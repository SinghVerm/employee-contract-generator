"""Validate employee names and required documents (no codes or employee IDs)."""
from __future__ import annotations

import csv
from pathlib import Path
from typing import Iterable

DEFAULT_DOCUMENTS = (
    "Passport - Photo Page", "Passport - Back Page", "Visa Evidence",
)


def validate_employees(raw: Iterable[dict]) -> list[dict]:
    employees: list[dict] = []
    seen: set[str] = set()
    for record in raw:
        name = str(record.get("name", "")).strip()
        if not name or len(name) > 120:
            raise ValueError("Each employee must have a name (up to 120 characters).")
        if name.casefold() in seen:
            raise ValueError(f"Duplicate employee name: {name}. Use unique names or add an ID internally.")
        seen.add(name.casefold())

        documents = record.get("documents", DEFAULT_DOCUMENTS)
        if isinstance(documents, str):
            documents = [part.strip() for part in documents.split("|") if part.strip()]
        else:
            documents = list(documents)
        if (not 1 <= len(documents) <= 15 or any(
            not isinstance(doc, str) or not doc.strip() or len(doc) > 100
            for doc in documents
        ) or len({doc.casefold() for doc in documents}) != len(documents)):
            raise ValueError(f"Invalid document requirements for {name}.")
        employees.append({"name": name, "documents": documents})
    if not employees:
        raise ValueError("No employees configured.")
    return employees


def load_employees(path: str | Path = "employees.csv") -> list[dict]:
    with open(path, newline="", encoding="utf-8-sig") as handle:
        reader = csv.DictReader(handle)
        if not reader.fieldnames or "name" not in reader.fieldnames:
            raise ValueError("employees.csv must contain a 'name' column.")
        return validate_employees(list(reader))


def search_employees(employees: list[dict], query: str) -> list[dict]:
    query = query.strip().casefold()
    if len(query) < 3:
        return []
    return [person for person in employees if query in person["name"].casefold()][:15]
