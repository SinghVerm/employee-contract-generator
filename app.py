"""Employee document portal: signed contract, WHS Word checklist and ID uploads.

Run from the folder containing your EXISTING Contract_template.docx and private
.streamlit/secrets.toml:
    python -m streamlit run app.py

Keep your original Google Drive connection and drive_storage.py / registry.py.
The WHS employee-acknowledgement statements are read from WHS_template.docx.
"""
from __future__ import annotations

import csv
from datetime import date, datetime, timezone
from io import BytesIO
from pathlib import Path
from zoneinfo import ZoneInfo

import streamlit as st
from docx import Document
from docx.shared import Inches
from PIL import Image
from streamlit_drawable_canvas import st_canvas

from drive_storage import (
    ensure_folder,
    get_drive,
    safe_folder_label,
    upload_document,
    validate_upload,
)
from registry import search_employees, validate_employees

# -------- SETTINGS: keep these from the original contract generator --------
TEMPLATE_FILE = Path("Contract_template.docx")
WHS_TEMPLATE_FILE = Path("WHS_template.docx")
DEFAULT_POSITION = "Cleaning Services Employee"
DEFAULT_PAY_RATE = "$30.35 per hour"
PERTH = ZoneInfo("Australia/Perth")

COMPANIES = {
    "Cockburn Gateways": {
        "company_name": "Finolex Auto Pty Ltd",
        "company_address": "Cockburn Gateways",
        "manager_name": "Harpreet Singh",
        "manager_phone": "0414409284",
    },
    "Galleria": {
        "company_name": "Ks & K Pty Ltd",
        "company_address": "Star Car Wash - Galleria, Galleria Shopping Centre, "
                           "4 Collier Rd, Morley WA 6062",
        "manager_name": "Kuldeep Kuldeep",
        "manager_phone": "0497580120",
    },
    "Innaloo": {
        "company_name": "Star Car Wash",
        "company_address": "Westfield Innaloo (Level 1 car park near Kmart Service Centre), "
                           "Cnr Scarborough Beach Rd & Ellen Stirling Blvd, Innaloo WA 6021",
        "manager_name": "Raj Partap Singh Dhanju",
        "manager_phone": "0470382410",
    },
    "Midland": {
        "company_name": "SY Enterprises Group Pty Ltd",
        "company_address": "Star Carwash Midland, Midland Gate Shopping Centre, "
                           "274 Great Eastern Hwy, Midland WA 6056",
        "manager_name": "Jatinder Singh",
        "manager_phone": "0468378717",
    },
}

st.set_page_config(page_title="Employee Documents", layout="centered")


# -------- CSV AND DATE HELPERS --------
def local_today() -> date:
    return datetime.now(PERTH).date()


def display_date(value: date) -> str:
    return value.strftime("%d/%m/%Y")


def parse_date(value: str | None) -> date | None:
    text = (value or "").strip()
    if not text:
        return None
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            pass
    raise ValueError(f"Invalid date '{text}'. Use DD/MM/YYYY or YYYY-MM-DD.")


def load_people(path: Path) -> list[dict]:
    """Extend the existing registry with start_date and location from CSV.

    Existing name,documents files still work. Optional columns: start_date,
    location, position, pay_rate, offer_date. Old numeric ID column is ignored.
    """
    with path.open("r", encoding="utf-8-sig", newline="") as f:
        rows = list(csv.DictReader(f))
    if not rows:
        raise ValueError("employees.csv is empty")
    normalized = []
    for original in rows:
        row = {str(k or "").strip().casefold().replace(" ", "_"): (v or "").strip()
               for k, v in original.items()}
        normalized.append(row)
    basic = validate_employees(normalized)
    people = []
    for person, row in zip(basic, normalized):
        location = row.get("location", "")
        if location:
            location = next((k for k in COMPANIES if k.casefold() == location.casefold()), location)
        if location and location not in COMPANIES:
            raise ValueError(f"Invalid location for {person['name']}: {location}")
        person.update(
            start_date=parse_date(row.get("start_date")),
            offer_date=parse_date(row.get("offer_date")),
            location=location,
            position=row.get("position") or DEFAULT_POSITION,
            pay_rate=row.get("pay_rate") or DEFAULT_PAY_RATE,
        )
        people.append(person)
    return people


def load_settings_and_people():
    try:
        settings = dict(st.secrets)
        oauth = settings.get("google_oauth")
        if (not str(settings.get("root_folder_id", "")).strip() or not oauth
                or not all(oauth.get(k) for k in
                           ("client_id", "client_secret", "refresh_token"))):
            raise ValueError("Missing Google Drive secrets")
        if not Path("employees.csv").exists():
            raise ValueError("employees.csv not found")
        return settings, load_people(Path("employees.csv"))
    except Exception as error:
        st.error(f"App configuration problem: {error}")
        st.stop()


def read_whs_questions() -> list[dict]:
    """Read the 15 employee WHS acknowledgement statements from Word."""
    doc = Document(str(WHS_TEMPLATE_FILE))
    if len(doc.tables) < 2 or len(doc.tables[1].rows) < 2:
        raise ValueError("WHS_template.docx does not contain the expected hazard table")
    return [
        {"area": row.cells[0].text.strip(), "control": row.cells[1].text.strip()}
        for row in doc.tables[1].rows[1:]
    ]


# -------- WORD CONTRACT ENGINE (using your unchanged template) --------
def paragraphs_in_container(container):
    """Visit paragraphs in document, tables and nested table cells."""
    for paragraph in container.paragraphs:
        yield paragraph
    for table in container.tables:
        for row in table.rows:
            for cell in row.cells:
                yield from paragraphs_in_container(cell)


def all_paragraphs(doc):
    yield from paragraphs_in_container(doc)
    for section in doc.sections:
        yield from paragraphs_in_container(section.header)
        yield from paragraphs_in_container(section.footer)


def replace_in_paragraph(paragraph, substitutions: dict[str, str]):
    """Replace placeholders even if Word split the {{tag}} across runs.

    Keeps paragraph styling, and most run styling, unlike clearing every run.
    """
    for original, replacement in substitutions.items():
        while original in "".join(run.text for run in paragraph.runs):
            joined = "".join(run.text for run in paragraph.runs)
            start = joined.find(original)
            end = start + len(original)
            offset = 0
            first_idx = last_idx = None
            first_offset = last_offset = 0
            for i, run in enumerate(paragraph.runs):
                next_offset = offset + len(run.text)
                if first_idx is None and start < next_offset:
                    first_idx, first_offset = i, start - offset
                if first_idx is not None and end <= next_offset:
                    last_idx, last_offset = i, end - offset
                    break
                offset = next_offset
            if first_idx is None or last_idx is None:
                break
            runs = paragraph.runs
            if first_idx == last_idx:
                r = runs[first_idx]
                r.text = r.text[:first_offset] + replacement + r.text[last_offset:]
            else:
                runs[first_idx].text = runs[first_idx].text[:first_offset] + replacement
                for i in range(first_idx + 1, last_idx):
                    runs[i].text = ""
                runs[last_idx].text = runs[last_idx].text[last_offset:]


def insert_signature(doc, signature_png: bytes | None, signed_on: date | None):
    """Insert signature where {{signature}} appears; append if absent."""
    inserted = False
    for p in all_paragraphs(doc):
        if "{{signature}}" not in p.text:
            continue
        replace_in_paragraph(p, {"{{signature}}": ""})
        if signature_png:
            p.add_run().add_picture(BytesIO(signature_png), width=Inches(2.2))
            inserted = True
        else:
            p.add_run("________________________")
    if signature_png and not inserted:
        doc.add_paragraph("Employee signature:")
        doc.add_paragraph().add_run().add_picture(BytesIO(signature_png), width=Inches(2.2))
    if signature_png and signed_on:
        doc.add_paragraph("Employee declaration: I have read and accept this employment contract.")
        timestamp = datetime.now(PERTH).strftime("%d/%m/%Y at %I:%M %p AWST")
        doc.add_paragraph(f"Employee signature submitted on {timestamp}.")


def build_contract(employee: dict, company: dict, details: dict,
                   signed_png: bytes | None = None, signed_on: date | None = None) -> bytes:
    doc = Document(str(TEMPLATE_FILE))
    # {{date_today}} now uses a genuine known date. A fabricated random date
    # before employment would misstate when a signed contract was created.
    effective_date = employee.get("offer_date") or local_today()
    substitutions = {
        "{{position}}": employee["position"],
        "{{pay_rate}}": employee["pay_rate"],
        "{{employee_name}}": employee["name"],
        "{{employee_address}}": details["address"],
        "{{employee_phone}}": details["phone"],
        "{{phone}}": details["phone"],
        "{{date_of_birth}}": display_date(details["dob"]),
        "{{dob}}": display_date(details["dob"]),
        "{{start_date}}": display_date(details["start_date"]),
        "{{date_today}}": display_date(effective_date),
        "{{offer_date}}": display_date(effective_date),
        "{{signing_date}}": display_date(signed_on) if signed_on else "",
        "{{employee_dob}}": display_date(details["dob"]),
        "{{employee_date_of_birth}}": display_date(details["dob"]),
        "{{employee_mobile}}": details["phone"],
        "{{phone_number}}": details["phone"],
        "{{company_name}}": company["company_name"],
        "{{company_address}}": company["company_address"],
        "{{manager_name}}": company["manager_name"],
        "{{manager_phone}}": company["manager_phone"],
    }
    for p in all_paragraphs(doc):
        replace_in_paragraph(p, substitutions)
    insert_signature(doc, signed_png, signed_on)
    stream = BytesIO()
    doc.save(stream)
    return stream.getvalue()


def signature_to_png(canvas) -> bytes | None:
    if canvas.image_data is None:
        return None
    # A white, empty canvas is still image_data. Require a drawn stroke.
    shapes = (canvas.json_data or {}).get("objects", [])
    if not shapes:
        return None
    image = Image.fromarray(canvas.image_data.astype("uint8")).convert("RGB")
    out = BytesIO()
    image.save(out, format="PNG")
    return out.getvalue()


# -------- DRIVE SAVE HELPERS --------
def save_generated(drive, root_folder_id: str, employee: dict, category: str,
                   filename: str, content: bytes, mime: str) -> str:
    """Save a generated DOCX or PDF into the employee's existing Drive hierarchy."""
    from googleapiclient.http import MediaIoBaseUpload
    employee_id = ensure_folder(drive, root_folder_id, safe_folder_label(employee["name"]))
    category_id = ensure_folder(drive, employee_id, category)
    timestamp = datetime.now(timezone.utc).strftime("%Y-%m-%d_%H-%M-%S_%f_UTC")
    stem, suffix = filename.rsplit(".", 1)
    safe_filename = f"{safe_folder_label(stem)} - {timestamp}.{suffix.lower()}"
    media = MediaIoBaseUpload(BytesIO(content), mimetype=mime, resumable=True)
    item = drive.files().create(
        body={"name": safe_filename, "parents": [category_id]},
        media_body=media, fields="id", supportsAllDrives=True,
    ).execute()
    return item["id"]


# -------- WHS SIGNED WORD CHECKLIST ENGINE --------
def build_whs_document(employee: dict, site: str, checklist: list[dict],
                       answers: list[str], notes: list[str], assessment_date: date,
                       signed_on: date, actions: str, signature_png: bytes) -> bytes:
    """Fill the cleaned WHS Word template with genuine employee responses."""
    if not signature_png or len(checklist) != len(answers) or len(checklist) != len(notes):
        raise ValueError("Missing signature or WHS answers")
    if any(answer not in {"Yes", "No", "N/A"} for answer in answers):
        raise ValueError("Please answer every checklist item")
    if any(answer in {"No", "N/A"} and not note.strip() for answer, note in zip(answers, notes)):
        raise ValueError("Enter a reason or question for every No or N/A answer")

    doc = Document(str(WHS_TEMPLATE_FILE))
    if len(doc.tables[1].rows) != len(checklist) + 1:
        raise ValueError("WHS template and checklist question count do not match")
    replacements = {
        "{{site}}": site,
        "{{employee_name}}": employee["name"],
        "{{assessment_date}}": display_date(assessment_date),
        "{{date_signed}}": display_date(signed_on),
        "{{overall_assessment}}": (
            "Manager follow-up required" if any(a != "Yes" for a in answers)
            else "All statements acknowledged"
        ),
        "{{immediate_actions}}": actions.strip() or "None reported",
    }
    for i, (answer, note) in enumerate(zip(answers, notes), 1):
        replacements[f"{{{{status_{i:02d}}}}}"] = answer
        replacements[f"{{{{notes_{i:02d}}}}}"] = note.strip() or "—"
    for paragraph in all_paragraphs(doc):
        replace_in_paragraph(paragraph, replacements)
        if "{{employee_signature}}" in paragraph.text:
            replace_in_paragraph(paragraph, {"{{employee_signature}}": ""})
            paragraph.add_run().add_picture(BytesIO(signature_png), width=Inches(1.5))
    output = BytesIO()
    doc.save(output)
    return output.getvalue()


# -------- EMPLOYEE-FACING APP --------
settings, employees = load_settings_and_people()
st.title("Employee Documents")
st.caption("Select your name, complete the contract and WHS checklist, and attach your documents. Everything is saved together at the end.")

query = st.text_input("Find your name (enter at least 3 letters)", placeholder="Start typing your name")
if len(query.strip()) < 3:
    st.stop()

matches = search_employees(employees, query)
if not matches:
    st.info("No matching name found. Please check the spelling.")
    st.stop()

selected = st.selectbox(
    "Select your name", [p["name"] for p in matches],
    index=None, placeholder="Choose your name",
)
if not selected:
    st.stop()
employee = next(p for p in matches if p["name"] == selected)

# Cockburn Gateways is the only worksite for this portal.
location = "Cockburn Gateways"
company = COMPANIES[location]

# These variables are populated by the tabs but no data goes to Drive until
# the ONE final submission button below is pressed.
contract_ready = False
contract_details = None
contract_signature = None
whs_ready = False
whs_payload = None
whs_signature = None
attachments = []
missing_documents = []

contract_tab, whs_tab, documents_tab = st.tabs([
    "Employment Contract", "WHS Checklist", "Documents",
])

with contract_tab:
    st.subheader("Employment contract")
    if not TEMPLATE_FILE.exists():
        st.error("Contract_template.docx is missing from the app folder.")
    else:
        st.write(f"**Employee:** {employee['name']}")
        st.write(f"**Position:** {employee['position']}")
        if employee["start_date"]:
            start = employee["start_date"]
            st.write(f"**Employment start date:** {display_date(start)}")
        else:
            start = st.date_input(
                "Employment start date", value=None, format="DD/MM/YYYY",
                key=f"start_{selected}",
            )
        address = st.text_area("Residential address", key=f"address_{selected}")
        phone = st.text_input("Phone number", key=f"phone_{selected}")
        dob_text = st.text_input(
            "Date of birth (DD/MM/YYYY)",
            placeholder="e.g. 15/06/1990",
            max_chars=10,
            key=f"dob_{selected}",
        )
        dob = None
        if dob_text.strip():
            try:
                dob = datetime.strptime(dob_text.strip(), "%d/%m/%Y").date()
                if not date(1900, 1, 1) <= dob <= local_today():
                    raise ValueError
            except ValueError:
                st.error("Please enter a valid date in DD/MM/YYYY format.")
                dob = None
        details_complete = bool(start and address.strip() and phone.strip() and dob)
        if not details_complete:
            st.info("Enter your start date (if needed), address, phone and date of birth.")
        st.write("**Draw your signature**")
        contract_canvas = st_canvas(
            fill_color="rgba(255,255,255,0)", stroke_width=2,
            stroke_color="#000000", background_color="#FFFFFF",
            height=160, width=480, drawing_mode="freedraw",
            update_streamlit=True, return_image_data=True, key=f"signature_{selected}",
        )
        accepted_contract = st.checkbox(
            "I have read and accept the employment contract.",
            key=f"accepts_{selected}",
        )
        contract_signature = signature_to_png(contract_canvas)
        contract_ready = bool(details_complete and contract_signature and accepted_contract)
        if details_complete:
            contract_details = {
                "start_date": start, "address": address.strip(),
                "phone": phone.strip(), "dob": dob,
            }
        if contract_ready:
            st.success("Contract details and signature ready for final submission.")
        else:
            st.caption("Your contract will be generated and saved when you submit everything at the bottom of this page.")

with whs_tab:
    st.subheader("Workplace Health & Safety Induction and Acknowledgement")
    if not WHS_TEMPLATE_FILE.exists():
        st.error("WHS_template.docx is missing from the app folder.")
    else:
        try:
            questions = read_whs_questions()
        except (ValueError, OSError) as exc:
            st.error(f"WHS template problem: {exc}")
            questions = []
        if questions:
            site = "Star Car Wash - Cockburn Gateways"
            st.write(f"**Business / Site:** {site}")
            st.write(f"**Completed by:** {employee['name']}")
            assessment_date = st.date_input(
                "Date completed", value=local_today(), max_value=local_today(),
                format="DD/MM/YYYY", key=f"whs_assessment_{selected}",
            )
            st.caption(
                "Read each statement. Yes = I understand; No = I need instruction or "
                "clarification; N/A = this does not apply to my role. "
                "Explain every No or N/A in the notes. "
                "This acknowledgement does not replace practical WHS training."
            )
            answers, notes = [], []
            missing_answers, missing_notes = [], []
            for i, question in enumerate(questions, 1):
                with st.container(border=True):
                    st.markdown(f"**{i}. {question['area']}**")
                    st.write(question["control"])
                    answer = st.radio(
                        "Your response", ["Yes", "No", "N/A"], index=None,
                        horizontal=True, key=f"whs_answer_{selected}_{i}",
                    )
                    note = st.text_input(
                        "Notes / questions" + (
                            " (required for No or N/A)" if answer in {"No", "N/A"}
                            else " (optional)"
                        ), key=f"whs_note_{selected}_{i}",
                    )
                    answers.append(answer)
                    notes.append(note)
                    if answer is None:
                        missing_answers.append(i)
                    elif answer in {"No", "N/A"} and not note.strip():
                        missing_notes.append(i)
            actions = st.text_area(
                "Questions, training or other follow-up requested (if any)",
                key=f"whs_actions_{selected}",
            )
            st.write("**Acknowledgement outcome:** " + (
                "Manager follow-up required" if any(a not in {None, "Yes"} for a in answers)
                else "All statements acknowledged" if all(a == "Yes" for a in answers)
                else "Incomplete"
            ))
            st.write(f"**Employee:** {employee['name']}  |  **Date signed:** {display_date(local_today())}")
            st.write("**Draw your WHS signature**")
            whs_canvas = st_canvas(
                fill_color="rgba(255,255,255,0)", stroke_width=2,
                stroke_color="#000000", background_color="#FFFFFF",
                height=140, width=480, drawing_mode="freedraw",
                update_streamlit=True, return_image_data=True, key=f"whs_signature_{selected}",
            )
            acknowledged = st.checkbox(
                "I have answered honestly, recorded any training or clarification I need, "
                "and understand that this does not replace practical workplace training.",
                key=f"whs_acknowledge_{selected}",
            )
            whs_signature = signature_to_png(whs_canvas)
            if missing_answers:
                st.info("Please answer all 15 WHS statements.")
            if missing_notes:
                st.info("Please add a note for each No or N/A answer.")
            whs_ready = bool(
                not missing_answers and not missing_notes and whs_signature and acknowledged
            )
            whs_payload = {
                "site": site, "questions": questions, "answers": answers,
                "notes": notes, "assessment_date": assessment_date,
                "actions": actions,
            }
            if whs_ready:
                st.success("WHS checklist and signature ready for final submission.")
            else:
                st.caption("The completed WHS checklist will be saved with your other documents.")
            st.caption("Manager review/sign-off is left blank. Any No or N/A must be reviewed by management; a signature is not a substitute for safety training.")

with documents_tab:
    st.subheader("Upload documents")
    types = employee["documents"]
    passport_required = any(s.lower().startswith("passport") for s in types)
    visa_required = any(s.lower().startswith("visa") for s in types)
    other_required = [
        s for s in types if not s.lower().startswith(("passport", "visa"))
    ]
    if passport_required:
        st.write("**Passport** — attach at least one photo or PDF. Both pages may be in one file.")
        passports = st.file_uploader(
            "Passport (front, back or combined)", type=["pdf", "jpg", "jpeg", "png"],
            accept_multiple_files=True, key=f"passport_{selected}",
        )
        if not passports:
            missing_documents.append("Passport")
        attachments.extend(("Passport", file) for file in passports)
    if visa_required:
        st.write("**Visa evidence** — attach at least one file.")
        visas = st.file_uploader(
            "Visa / grant notice", type=["pdf", "jpg", "jpeg", "png"],
            accept_multiple_files=True, key=f"visa_{selected}",
        )
        if not visas:
            missing_documents.append("Visa")
        attachments.extend(("Visa Evidence", file) for file in visas)
    for name in other_required:
        files = st.file_uploader(
            name, type=["pdf", "jpg", "jpeg", "png"],
            accept_multiple_files=True, key=f"other_{selected}_{name}",
        )
        if not files:
            missing_documents.append(name)
        attachments.extend((name, file) for file in files)
    st.caption("PDF, JPG or PNG only; maximum 15 MB per file.")
    if missing_documents:
        st.info("Please attach at least one file for: " + ", ".join(missing_documents))
    else:
        st.success("Required documents attached and ready for final submission.")

# -------- SINGLE FINAL SUBMISSION --------
st.divider()
st.subheader("Final submission")

# In Streamlit, all tabs execute on each rerun, and widget values persist while
# the user switches between tabs. Do not upload anything earlier in the script.
ready = contract_ready and whs_ready and not missing_documents and bool(attachments)
needed = []
if not contract_ready:
    needed.append("signed employment contract")
if not whs_ready:
    needed.append("signed WHS checklist")
if missing_documents:
    needed.append("documents: " + ", ".join(missing_documents))
if needed:
    st.info("To enable final submission, complete: " + "; ".join(needed) + ".")

submission_key = f"submission_complete_{selected}"
submitted_already = st.session_state.get(submission_key, False)
if st.button(
    "Submit everything to Google Drive", type="primary", use_container_width=True,
    disabled=(not ready or submitted_already), key=f"submit_everything_{selected}",
):
    try:
        # Prepare/validate every file BEFORE making any Google Drive uploads.
        signed_on = local_today()
        contract_bytes = build_contract(
            employee, company, contract_details, contract_signature, signed_on,
        )
        whs_bytes = build_whs_document(
            employee, whs_payload["site"], whs_payload["questions"],
            whs_payload["answers"], whs_payload["notes"],
            whs_payload["assessment_date"], signed_on,
            whs_payload["actions"], whs_signature,
        )
        uploads = []
        for category, file in attachments:
            payload = file.getvalue()
            validate_upload(file.name, payload)
            uploads.append((category, file.name, payload))
    except (ValueError, OSError, KeyError) as exc:
        st.error(f"Please check your information and attachments: {exc}")
    except Exception as exc:
        st.error(f"Could not prepare your documents: {exc}")
    else:
        saved_count = 0
        try:
            with st.spinner("Saving contract, WHS checklist and documents to Google Drive..."):
                drive = get_drive(settings)
                root_id = str(settings["root_folder_id"])
                save_generated(
                    drive, root_id, employee, "Employment Contract",
                    "Signed Employment Contract.docx", contract_bytes,
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                )
                saved_count += 1
                save_generated(
                    drive, root_id, employee, "WHS Checklist",
                    "Signed WHS Checklist.docx", whs_bytes,
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                )
                saved_count += 1
                for category, name, data in uploads:
                    upload_document(drive, root_id, employee, category, name, data)
                    saved_count += 1
        except Exception:
            st.error(
                f"Submission stopped after saving {saved_count} file(s). "
                "Some files may already be on Google Drive. Check with your employer "
                "before retrying to avoid duplicate uploads."
            )
        else:
            st.session_state[submission_key] = True
            st.session_state[f"contract_download_{selected}"] = contract_bytes
            st.session_state[f"whs_download_{selected}"] = whs_bytes
            st.success("All documents saved to Google Drive successfully.")
            if any(a != "Yes" for a in whs_payload["answers"]):
                st.warning("Your WHS responses include No or N/A. Your manager should review the notes and provide any required training or clarification.")

if st.session_state.get(submission_key, False):
    st.success("Your submission is complete.")
    st.download_button(
        "Download signed contract",
        data=st.session_state[f"contract_download_{selected}"],
        file_name=f"{employee['name'].replace(' ', '_')}_Signed_Contract.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        key=f"download_contract_{selected}",
    )
    st.download_button(
        "Download signed WHS checklist",
        data=st.session_state[f"whs_download_{selected}"],
        file_name=f"{employee['name'].replace(' ', '_')}_Signed_WHS.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        key=f"download_whs_{selected}",
    )
