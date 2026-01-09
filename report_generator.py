from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Iterable, List, Sequence
from zipfile import ZipFile
import re
import unicodedata

import requests
from docxtpl import DocxTemplate, RichText

CRISTIN_API_URL = "https://api.cristin.no/v2/persons/{person_id}/results"
CRISTIN_PERSON_URL = "https://api.cristin.no/v2/persons/{person_id}"
DEFAULT_OUTPUT_DIR = Path("reports")
TEMPLATE_PATTERN = "Aarsrapport-plan_MAL.docx"
PER_PAGE_DEFAULT = 100
MAX_PAGES_DEFAULT = 25
REQUEST_TIMEOUT_SECONDS = 30
PLACEHOLDER_PATTERN = re.compile(r"{{\s*([a-zA-Z0-9_]+)\s*}}")

CATEGORY_KEYS = [
    "publisert_monografi_niva2",
    "publisert_monografi_niva1",
    "publisert_artikkel_niva2",
    "publisert_artikkel_niva1",
    "publisert_antologi_niva2",
    "publisert_antologi_niva1",
    "publisert_book_review",
    "publisert_annet",
]

MANUAL_FIELD_KEYS = [
    "forskningsarbeid_internasjonal_deltagelse",
    "forskningsarbeid_internasjonal_ledelse",
    "forskningsarbeid_nasjonal_deltagelse",
    "forskningsarbeid_nasjonal_ledelse",
    "forskningsarbeid_innvilget_soknad",
    "forskningsarbeid_utenlandsopphold",
    "forskningsarbeid_innovasjon",
    "forskningsarbeid_nasjonale_nettverk",
    "forskningsarbeid_internasjonale_nettverk",
    "formidling_faglig",
    "formidling_politisk",
    "formidling_kronikker",
    "formidling_popularvitenskapelig",
    "formidling_media",
    "veiledning_phd",
    "opponent_phd",
    "referee_vitenskapelige_artikler",
    "veiledning_masteroppgave",
    "sensur_masteroppgave",
    "professor_vurderinger",
]

REQUIRED_PLACEHOLDERS = [
    "report_year",
    "person_name",
    "institution_name",
    "institution_name_secondary",
    *CATEGORY_KEYS,
    *MANUAL_FIELD_KEYS,
]

NON_PUBLICATION_CATEGORY_CODES = {
    "ACADEMICLECTURE",
    "LECTURE",
    "POSTER",
    "OTHERPRES",
    "MEDIAINTERVIEW",
    "PROGRAMPARTICIP",
    "ARTICLEFEATURE",
    "READEROPINION",
}

FORMIDLING_CATEGORY_MAP = {
    "ACADEMICLECTURE": "formidling_faglig",
    "LECTURE": "formidling_faglig",
    "POSTER": "formidling_faglig",
    "OTHERPRES": "formidling_faglig",
    "MEDIAINTERVIEW": "formidling_media",
    "PROGRAMPARTICIP": "formidling_media",
    "ARTICLEFEATURE": "formidling_kronikker",
    "READEROPINION": "formidling_kronikker",
}


@dataclass(frozen=True)
class PublicationEntry:
    """Structured representation of a publication reference."""

    reference: str
    category_key: str
    year: str


def normalize_text(text: str) -> str:
    """Normalize text for matching (lowercase + remove diacritics)."""
    return "".join(
        char for char in unicodedata.normalize("NFKD", text) if ord(char) < 128
    ).lower().strip()


def fetch_publications(
    person_id: int | str,
    per_page: int = PER_PAGE_DEFAULT,
    max_pages: int = MAX_PAGES_DEFAULT,
    session: requests.Session | None = None,
) -> List[dict]:
    """
    Fetch all publication results for a CRISTIN person ID.

    Pagination is handled automatically. max_pages is a safety valve to prevent infinite loops.
    """
    person_id = str(person_id).strip()
    if not person_id.isdigit():
        raise ValueError("person_id must be numeric")

    base_url = CRISTIN_API_URL.format(person_id=person_id)
    page = 1
    results: List[dict] = []
    client = session or requests.Session()

    while page <= max_pages:
        response = client.get(
            base_url,
            params={"page": page, "per_page": per_page},
            timeout=REQUEST_TIMEOUT_SECONDS,
        )
        response.raise_for_status()
        page_results = response.json()

        if not page_results:
            break

        results.extend(page_results)
        page += 1

    return results


def fetch_person_details(
    person_id: int | str, session: requests.Session | None = None
) -> dict:
    """Fetch metadata for a CRISTIN person ID."""
    person_id = str(person_id).strip()
    if not person_id.isdigit():
        raise ValueError("person_id must be numeric")

    client = session or requests.Session()
    response = client.get(
        CRISTIN_PERSON_URL.format(person_id=person_id),
        timeout=REQUEST_TIMEOUT_SECONDS,
    )
    response.raise_for_status()
    return response.json()


def extract_person_name(person: dict) -> str:
    """Return a readable person name from a CRISTIN person payload."""
    for first_key, last_key in (
        ("first_name", "surname"),
        ("given_name", "family_name"),
        ("firstName", "lastName"),
    ):
        first = str(person.get(first_key, "")).strip()
        last = str(person.get(last_key, "")).strip()
        if first or last:
            return " ".join(part for part in (first, last) if part)

    for key in ("full_name", "display_name", "name"):
        value = str(person.get(key, "")).strip()
        if value:
            return value

    return "Ukjent navn"


def extract_affiliation_names(person: dict) -> tuple[str, str]:
    """Extract institution name(s) from a CRISTIN person payload."""
    affiliations = person.get("affiliations") or person.get("employments") or []
    names: List[str] = []
    for aff in affiliations:
        org = aff.get("organization") or aff.get("institution") or {}
        name_obj = org.get("name") or org.get("name_short") or org.get("title")
        if isinstance(name_obj, dict):
            for value in name_obj.values():
                if value:
                    names.append(str(value).strip())
        elif isinstance(name_obj, str) and name_obj.strip():
            names.append(name_obj.strip())

    deduped: List[str] = []
    for name in names:
        if name and name not in deduped:
            deduped.append(name)
        if len(deduped) >= 2:
            break

    primary = deduped[0] if deduped else ""
    secondary = deduped[1] if len(deduped) > 1 else ""
    return primary, secondary


def extract_category_code_name(item: dict) -> tuple[str | None, str | None]:
    """Return category code and display name if present."""
    category = item.get("category")
    if not isinstance(category, dict):
        return None, None

    code = category.get("code")
    name_obj = category.get("name")
    name = None
    if isinstance(name_obj, dict):
        name = next((value for value in name_obj.values() if value), None)
    elif isinstance(name_obj, str):
        name = name_obj

    return code, name


def sanitize_filename(value: str, fallback: str) -> str:
    """Sanitize a filename segment while keeping it readable."""
    cleaned = str(value or "").strip()
    if not cleaned:
        return fallback
    cleaned = re.sub(r'[<>:"/\\\\|?*]+', "", cleaned)
    cleaned = re.sub(r"\\s+", " ", cleaned).strip()
    cleaned = cleaned.rstrip(". ")
    return cleaned or fallback


def build_output_filename(person_name: str, person_id: int | str, report_year: int | str) -> str:
    """Build a user-friendly output filename."""
    safe_name = sanitize_filename(person_name, sanitize_filename(str(person_id), "report"))
    safe_year = sanitize_filename(str(report_year), "year")
    return f"Aarsrapport_{safe_year}_{safe_name}.docx"


def get_title(item: dict) -> str:
    """Return the first non-empty title from the multilingual title field."""
    title_obj = item.get("title", {})
    if isinstance(title_obj, dict):
        for value in title_obj.values():
            if value:
                return value
    if isinstance(title_obj, str) and title_obj.strip():
        return title_obj.strip()
    return "No title available"


def extract_doi(item: dict) -> str | None:
    """Extract the first DOI link if present."""
    for link in item.get("links", []):
        url_type = str(link.get("url_type", "")).upper()
        if url_type == "DOI":
            url = link.get("url")
            if url:
                return url
    return None


def extract_year(item: dict) -> str:
    """Extract publication year as a string."""
    year = item.get("year_published") or item.get("year") or ""
    return str(year) if year else ""


def format_authors(item: dict) -> str:
    """Format authors as 'Surname, First' separated by semicolons."""
    authors = []
    for author in item.get("contributors", {}).get("preview", []):
        surname = str(author.get("surname", "")).strip()
        first_name = str(author.get("first_name", "")).strip()
        if surname or first_name:
            authors.append(", ".join(part for part in (surname, first_name) if part)
            )
    return "; ".join(authors) or "Unknown author"


def format_reference(item: dict) -> str:
    """Build a simple APA-like reference string from a CRISTIN item."""
    authors = format_authors(item)
    year = extract_year(item) or "n.d."
    title = get_title(item)
    publication_info = (
        item.get("journal", {}).get("name")
        or item.get("event", {}).get("name")
        or item.get("channel", {}).get("title")
        or "No publication information available"
    )
    volume = item.get("volume") or "N/A"
    pages_obj = item.get("pages", {}) if isinstance(item.get("pages"), dict) else {}
    pages = f"{pages_obj.get('from', 'N/A')}-{pages_obj.get('to', 'N/A')}"
    doi = extract_doi(item)

    reference = f"{authors} ({year}). {title}. {publication_info}, Volume {volume}, pp. {pages}."
    if doi:
        reference = f"{reference} {doi}."
    return reference


def extract_level(item: dict) -> str | None:
    """Best-effort extraction of publication level (1 or 2)."""
    candidates = [
        item.get("level"),
        item.get("publication_level"),
        item.get("scientific_level"),
        item.get("publication_context", {}).get("level"),
        item.get("publicationContext", {}).get("level"),
        item.get("publication_channel", {}).get("level"),
        item.get("publicationChannel", {}).get("level"),
        item.get("channel", {}).get("level"),
        item.get("journal", {}).get("level"),
        item.get("channel", {}).get("nvi_level"),
        item.get("journal", {}).get("nvi_level"),
    ]

    for value in candidates:
        if value in (1, 2, "1", "2"):
            return str(value)
    return None


def classify_publication(item: dict) -> str | None:
    """Map CRISTIN category to a template placeholder key."""
    category_code, category_label = extract_category_code_name(item)
    if category_code in NON_PUBLICATION_CATEGORY_CODES:
        return None

    normalized = normalize_text(category_label or "")

    if category_code == "ARTICLE":
        level = extract_level(item) or "1"
        return f"publisert_artikkel_niva{level}"

    if category_code == "TEXTBOOK":
        level = extract_level(item) or "1"
        return f"publisert_monografi_niva{level}"

    if "book review" in normalized or "bokanmeldelse" in normalized:
        return "publisert_book_review"

    if "monograph" in normalized or "monografi" in normalized or "book" in normalized:
        level = extract_level(item) or "1"
        return f"publisert_monografi_niva{level}"

    if "anthology" in normalized or "antologi" in normalized or "edited" in normalized:
        level = extract_level(item) or "1"
        return f"publisert_antologi_niva{level}"

    if "article" in normalized or "artikkel" in normalized:
        level = extract_level(item) or "1"
        return f"publisert_artikkel_niva{level}"

    return "publisert_annet"


def build_entries(items: Sequence[dict], report_year: int | str) -> List[PublicationEntry]:
    """Filter and format publications for the selected year."""
    year_str = str(report_year)
    entries: List[PublicationEntry] = []
    for item in items:
        item_year = extract_year(item)
        if item_year != year_str:
            continue
        category_key = classify_publication(item)
        if not category_key:
            continue
        reference = format_reference(item)
        entries.append(
            PublicationEntry(reference=reference, category_key=category_key, year=item_year)
        )
    return entries


def group_references(entries: Iterable[PublicationEntry]) -> Dict[str, List[str]]:
    """Group references by template category key."""
    grouped: Dict[str, List[str]] = {key: [] for key in CATEGORY_KEYS}
    for entry in entries:
        grouped.setdefault(entry.category_key, []).append(entry.reference)
    for refs in grouped.values():
        refs.sort()
    return grouped


def build_template_context(
    entries: Sequence[PublicationEntry],
    report_year: int | str,
    person_name: str,
    institution_name: str,
    institution_name_secondary: str,
    manual_fields: Dict[str, str] | None = None,
) -> dict:
    """Prepare context for docxtpl template rendering."""
    grouped = group_references(entries)

    def join_refs(key: str) -> str | RichText:
        refs = grouped.get(key, [])
        if not refs:
            return ""

        rich_text = RichText()
        for idx, ref in enumerate(refs):
            rich_text.add(ref)
            if idx < len(refs) - 1:
                rich_text.add("\n\n")
        return rich_text

    context = {
        "report_year": str(report_year),
        "person_name": person_name,
        "institution_name": institution_name,
        "institution_name_secondary": institution_name_secondary,
    }

    for key in CATEGORY_KEYS:
        context[key] = join_refs(key)

    normalized_manual_fields = {key: "" for key in MANUAL_FIELD_KEYS}
    if manual_fields:
        for key in MANUAL_FIELD_KEYS:
            value = manual_fields.get(key, "")
            normalized_manual_fields[key] = str(value).strip()

    context.update(normalized_manual_fields)

    return context


def format_formidling_entry(item: dict) -> str:
    """Format a dissemination/activity entry for manual fields."""
    title = get_title(item)
    year = extract_year(item) or "n.d."
    authors = format_authors(item)
    event = item.get("event") or {}
    event_name = event.get("name")
    event_location = event.get("location")
    channel_title = (
        (item.get("channel") or {}).get("title")
        or (item.get("journal") or {}).get("name")
    )

    parts: List[str] = []
    if authors and authors != "Unknown author":
        parts.append(f"{authors} ({year}). {title}.")
    else:
        parts.append(f"{title} ({year}).")

    if event_name:
        venue = event_name
        if event_location:
            venue = f"{venue}, {event_location}"
        parts.append(f"{venue}.")
    elif channel_title:
        parts.append(f"{channel_title}.")

    return " ".join(parts).strip()


def build_auto_manual_fields(items: Sequence[dict], report_year: int | str) -> Dict[str, str]:
    """Populate manual fields from NVA categories where possible."""
    year_str = str(report_year)
    grouped: Dict[str, List[str]] = {key: [] for key in MANUAL_FIELD_KEYS}

    for item in items:
        item_year = extract_year(item)
        if item_year != year_str:
            continue

        category_code, category_label = extract_category_code_name(item)
        manual_key = None
        if category_code and category_code in FORMIDLING_CATEGORY_MAP:
            manual_key = FORMIDLING_CATEGORY_MAP[category_code]
        elif category_label:
            normalized = normalize_text(category_label)
            if "interview" in normalized:
                manual_key = "formidling_media"
            elif "poster" in normalized or "presentation" in normalized or "lecture" in normalized:
                manual_key = "formidling_faglig"

        if not manual_key:
            continue

        grouped.setdefault(manual_key, []).append(format_formidling_entry(item))

    return {key: "\n\n".join(values) if values else "" for key, values in grouped.items()}


def merge_manual_fields(
    auto_fields: Dict[str, str], manual_fields: Dict[str, str] | None
) -> Dict[str, str]:
    """Combine auto-populated and user-provided manual fields."""
    merged: Dict[str, str] = {key: "" for key in MANUAL_FIELD_KEYS}
    manual_fields = manual_fields or {}

    for key in MANUAL_FIELD_KEYS:
        auto_value = str(auto_fields.get(key, "")).strip()
        manual_value = str(manual_fields.get(key, "")).strip()
        if auto_value and manual_value:
            merged[key] = f"{auto_value}\n\n{manual_value}"
        else:
            merged[key] = manual_value or auto_value

    return merged


def resolve_template_path(base_dir: Path | None = None) -> Path:
    """Find the template matching the expected pattern."""
    search_dir = base_dir or Path(__file__).parent / "templates"
    candidates = list(search_dir.glob(TEMPLATE_PATTERN))
    if not candidates:
        raise FileNotFoundError(
            f"No template found in {search_dir}. Expected file named {TEMPLATE_PATTERN}."
        )
    return candidates[0]


def render_report(context: dict, template_path: Path, output_path: Path) -> Path:
    """Render the Word document using docxtpl with the provided context."""
    doc = DocxTemplate(template_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc.render(context)
    doc.save(output_path)
    return output_path


def extract_template_placeholders(template_path: Path) -> set[str]:
    """Extract simple Jinja placeholders from a docx template."""
    with ZipFile(template_path) as docx:
        try:
            xml = docx.read("word/document.xml").decode("utf-8")
        except KeyError:
            return set()
    return set(PLACEHOLDER_PATTERN.findall(xml))


def find_missing_placeholders(
    template_path: Path, required_placeholders: Sequence[str] = REQUIRED_PLACEHOLDERS
) -> List[str]:
    """Return required placeholders that do not appear in the template."""
    present = extract_template_placeholders(template_path)
    missing = sorted(set(required_placeholders) - present)
    return missing


def generate_report(
    person_id: int | str,
    report_year: int | str,
    output_dir: Path = DEFAULT_OUTPUT_DIR,
    template_path: Path | None = None,
    session: requests.Session | None = None,
    manual_fields: Dict[str, str] | None = None,
) -> Path:
    """High-level helper: fetch data, build context, and fill the template."""
    publications = fetch_publications(person_id, session=session)
    entries = build_entries(publications, report_year)
    auto_manual_fields = build_auto_manual_fields(publications, report_year)
    merged_manual_fields = merge_manual_fields(auto_manual_fields, manual_fields)

    person_name = "Ukjent navn"
    institution_name = ""
    institution_name_secondary = ""
    try:
        person = fetch_person_details(person_id, session=session)
    except requests.RequestException:
        person = {}
    else:
        person_name = extract_person_name(person)
        institution_name, institution_name_secondary = extract_affiliation_names(person)

    context = build_template_context(
        entries,
        report_year,
        person_name,
        institution_name,
        institution_name_secondary,
        manual_fields=merged_manual_fields,
    )
    template = template_path or resolve_template_path()
    output_filename = build_output_filename(person_name, person_id, report_year)
    output_path = output_dir / output_filename
    return render_report(context, template, output_path)


def save_publications_to_json(publications: Iterable[dict], output_path: Path) -> Path:
    """Persist raw publications to disk for debugging or offline work."""
    import json

    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open("w", encoding="utf-8") as fp:
        json.dump(list(publications), fp, ensure_ascii=False, indent=2)
    return output_path
