# Employee Document Portal — name-only version

A simple Streamlit interface. An employee searches their name (3+ characters), chooses a document type, and uploads a PDF/photo. The app creates one folder per employee in the Google Drive folder previously authorised by the employer. No employee ID, access code, or login is requested.

**IMPORTANT — NOT FOR LIVE ID DOCUMENT COLLECTION WITHOUT ADDITIONAL SECURITY.** Anyone who knows the Streamlit URL can select any employee, submit a file under their name, and use storage. The app deliberately does not reveal any employee's existing document list or whether files were received. For passport and visa documents, introduce genuine employee verification / individual invitation links, abuse controls, and suitable data-retention/access policies before public deployment. A checkbox is only an anti-mistake reminder, **not authentication**.

## Upgrading from the original project (Windows)

1. Close the running Streamlit application with `Ctrl+C` in Command Prompt.
2. Make a backup of your existing `employee_document_portal` folder.
3. Copy this version's `app.py`, `registry.py`, `drive_storage.py`, `employees.csv`, and `.streamlit/config.toml` into that existing folder, replacing the old copies. **Do not overwrite or publish `.streamlit/secrets.toml`**. Your previous Google Drive authorisation still works.
4. If `.streamlit/secrets.toml` does **not** exist yet, run `type google_secrets.toml > .streamlit\secrets.toml` in the project folder. No `employee_secrets.toml` is needed.
5. Run `python -m streamlit run app.py`.

Existing IDs and codes contained in `.streamlit/secrets.toml` are ignored as long as `employees.csv` exists. You can remove the old `[[employees]]` sections later and keep only the `root_folder_id` and `[google_oauth]` section. **Do not re-run Google authorisation.**

## Files and configuration

- `employees.csv`: 10 pre-entered employees, `name,documents` columns. Separate document types with `|`; edit to change requirements. Names must be unique (case insensitive). There is no ID field.
- `.streamlit/secrets.toml`: private Google OAuth credentials and root Drive folder ID, already generated previously.
- `connect_google_drive.py`: use **only** if you have not already authorised Google Drive.
- `drive_storage.py`: keeps Drive folder-per-name and type-per-document, saving timestamped copies.

Names are not confidential credentials, but they are employee personal data. The `.gitignore` excludes `employees.csv` and the credentials from Git. For Streamlit Community Cloud hosting, keep the roster in **app Secrets** as `[[employees]]` entries instead of committing a CSV into a public GitHub repository. If deployed without `employees.csv`, settings can include:

```toml
[[employees]]
name = "Example Employee"
documents = ["Passport - Photo Page", "Passport - Back Page", "Visa Evidence"]
```

## Local install from scratch

```powershell
python -m pip install -r requirements.txt
```

If never authorised, download a Google Desktop OAuth JSON file from your Google Cloud Console, save it as `client_secret.json`, run `python connect_google_drive.py`, and copy `google_secrets.toml` into `.streamlit/secrets.toml`. If already authorised, skip this entire step.

Then:

```powershell
python -m streamlit run app.py
```

Do not upload TFNs, bank details or any documents not requested. Uploads accept PDF/JPG/PNG up to 15 MB. An employee can submit another copy without replacing the previous one. Only the employer's Drive credentials are used to write files. Files are not downloadable through the app.

## Operational caveats

- The OAuth app in External/Testing may have refresh tokens expire after seven days; plan an appropriate longer-lived verified/published or Workspace-based access before production.
- No authenticated audit trail exists of which person actually uploaded a file.
- A public anonymous upload endpoint needs rate limiting, malware scanning, access protection, and retention/deletion procedures before handling identity documents.
- Use VEVO where appropriate to verify visa work rights; a file alone is not proof of current eligibility.
