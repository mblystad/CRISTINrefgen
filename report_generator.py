from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Iterable, List, Sequence
import unicodedata

import requests
from docxtpl import DocxTemplate

CRISTIN_API_URL = "https://api.cristin.no/v2/persons/{person_id}/results"
CRISTIN_PERSON_URL = "https://api.cristin.no/v2/persons/{person_id}"
DEFAULT_OUTPUT_DIR = Path("reports")
TEMPLATE_PATTERN = "Aarsrapport-plan_MAL.docx"

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
    per_page: int = 100,
    max_pages: int = 25,
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
        response = client.get(base_url, params={"page": page, "per_page": per_page}, timeout=30)
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
    response = client.get(CRISTIN_PERSON_URL.format(person_id=person_id), timeout=30)
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
    for aff in affiliations:
        org = aff.get("organization") or aff.get("institution") or {}
        name_obj = org.get("name") or org.get("name_short") or org.get("title")
        if isinstance(name_obj, dict):
            for value in name_obj.values():
                if value:
                    return value, ""
        if isinstance(name_obj, str) and name_obj.strip():
            return name_obj.strip(), ""
    return "", ""


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
    ]

    for value in candidates:
        if value in (1, 2, "1", "2"):
            return str(value)
    return None


def classify_publication(item: dict) -> str:
    """Map CRISTIN category to a template placeholder key."""
    category_obj = item.get("category", {}).get("name", {})
    category_label = ""
    if isinstance(category_obj, dict):
        for value in category_obj.values():
            if value:
                category_label = value
                break
    elif isinstance(category_obj, str):
        category_label = category_obj

    normalized = normalize_text(category_label)

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
        reference = format_reference(item)
        category_key = classify_publication(item)
        entries.append(PublicationEntry(reference=reference, category_key=category_key, year=item_year))
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
) -> dict:
    """Prepare context for docxtpl template rendering."""
    grouped = group_references(entries)

    def join_refs(key: str) -> str:
        refs = grouped.get(key, [])
        return "\n".join(refs)

    context = {
        "report_year": str(report_year),
        "person_name": person_name,
        "institution_name": institution_name,
        "institution_name_secondary": institution_name_secondary,
    }

    for key in CATEGORY_KEYS:
        context[key] = join_refs(key)

    return context


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


def generate_report(
    person_id: int | str,
    report_year: int | str,
    output_dir: Path = DEFAULT_OUTPUT_DIR,
    template_path: Path | None = None,
    session: requests.Session | None = None,
) -> Path:
    """High-level helper: fetch data, build context, and fill the template."""
    publications = fetch_publications(person_id, session=session)
    entries = build_entries(publications, report_year)

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
    )
    template = template_path or resolve_template_path()
    output_path = output_dir / f"{person_id}_{report_year}_arsrapport.docx"
    return render_report(context, template, output_path)


def save_publications_to_json(publications: Iterable[dict], output_path: Path) -> Path:
    """Persist raw publications to disk for debugging or offline work."""
    import json

    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open("w", encoding="utf-8") as fp:
        json.dump(list(publications), fp, ensure_ascii=False, indent=2)
    return output_path
