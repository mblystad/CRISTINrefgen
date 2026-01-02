from __future__ import annotations

import json
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Iterable, List, Sequence, Tuple

import matplotlib.pyplot as plt
import requests
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Inches

CRISTIN_API_URL = "https://api.cristin.no/v2/persons/{person_id}/results"
DEFAULT_OUTPUT_DIR = Path("reports")
TEMPLATE_PATTERN = "*rsrapport-plan_MAL.docx"


@dataclass
class Reference:
    """Structured representation of a publication reference."""

    text: str
    category: str
    year: str


def fetch_publications(person_id: int | str, per_page: int = 100, max_pages: int = 25,
                       session: requests.Session | None = None) -> List[dict]:
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


def get_title(item: dict) -> str:
    """Return the first non-empty title from the multilingual title field."""
    title_obj = item.get("title", {})
    for value in title_obj.values():
        if value:
            return value
    return "No title available"


def extract_doi(item: dict) -> str:
    """Extract the first DOI link if present."""
    for link in item.get("links", []):
        if link.get("url_type", "").upper() == "DOI":
            return link.get("url", "No DOI available")
    return "No DOI available"


def build_reference(item: dict) -> Reference:
    """
    Build an APA-like reference string from a CRISTIN item.

    The format is intentionally simple; adjust as template requirements evolve.
    """
    authors = "; ".join(
        f"{author.get('surname', '').strip()}, {author.get('first_name', '').strip()}"
        for author in item.get("contributors", {}).get("preview", [])
        if author.get("surname") or author.get("first_name")
    ) or "Unknown author"

    year = str(item.get("year_published", "n.d."))
    title = get_title(item)
    publication_info = (
        item.get("journal", {}).get("name")
        or item.get("event", {}).get("name")
        or item.get("channel", {}).get("title")
        or "No publication information available"
    )
    volume = item.get("volume") or "N/A"
    pages_obj = item.get("pages", {})
    pages = f"{pages_obj.get('from', 'N/A')}-{pages_obj.get('to', 'N/A')}"
    doi = extract_doi(item)
    category = item.get("category", {}).get("name", {}).get("en") or "Unknown type"

    reference_text = (
        f"{authors} ({year}). {title}. {publication_info}, Volume {volume}, pp. {pages}. {doi}."
    )
    return Reference(text=reference_text, category=category, year=year)


def format_references(items: Sequence[dict]) -> Tuple[List[str], Dict[str, int]]:
    """Return formatted references sorted by year (desc) and category counts."""
    references: List[Reference] = [build_reference(item) for item in items]
    references.sort(key=lambda ref: ref.year, reverse=True)

    category_counts: Dict[str, int] = {}
    for ref in references:
        category_counts[ref.category] = category_counts.get(ref.category, 0) + 1

    return [ref.text for ref in references], category_counts


def generate_category_chart(category_counts: Dict[str, int], output_path: Path) -> Path:
    """Create a donut pie chart showing distribution of publication categories."""
    if not category_counts:
        raise ValueError("No categories to plot.")

    labels = [f"{category} ({count})" for category, count in category_counts.items()]
    counts = list(category_counts.values())

    plt.figure(figsize=(10, 7))
    wedges, texts, autotexts = plt.pie(
        counts,
        labels=labels,
        autopct="%1.1f%%",
        startangle=90,
        colors=plt.cm.Paired(range(len(counts))),
        textprops={"fontsize": 10},
        wedgeprops={"width": 0.35, "edgecolor": "w"},
        pctdistance=0.8,
        labeldistance=1.05,
    )

    for autotext in autotexts:
        autotext.set_color("white")
        autotext.set_fontsize(9)

    plt.title("Category Distribution", fontsize=14)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    plt.savefig(output_path, bbox_inches="tight")
    plt.close()
    return output_path


def build_template_context(references: Sequence[str], category_counts: Dict[str, int], chart_path: Path) -> dict:
    """Prepare context for docxtpl template rendering."""
    total_entries = sum(category_counts.values())
    return {
        "total_entries": total_entries,
        "category_counts": [{"name": k, "count": v} for k, v in category_counts.items()],
        "references": [{"text": ref} for ref in references],
        "category_pie_chart": chart_path.name,
    }


def resolve_template_path(base_dir: Path | None = None) -> Path:
    """Find the template matching the expected pattern."""
    search_dir = base_dir or Path(__file__).parent / "templates"
    candidates = list(search_dir.glob(TEMPLATE_PATTERN))
    if not candidates:
        raise FileNotFoundError(
            f"No template found in {search_dir}. Expected file matching {TEMPLATE_PATTERN}."
        )
    return candidates[0]


def render_report(context: dict, template_path: Path, output_path: Path, chart_path: Path) -> Path:
    """Render the Word document using docxtpl with the provided context and chart."""
    doc = DocxTemplate(template_path)

    # Ensure chart is embedded if present
    if chart_path.exists():
        context["category_pie_chart"] = InlineImage(doc, chart_path, width=Inches(4))
    else:
        context["category_pie_chart"] = None

    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc.render(context)
    doc.save(output_path)
    return output_path


def generate_report(person_id: int | str, output_dir: Path = DEFAULT_OUTPUT_DIR,
                    template_path: Path | None = None,
                    session: requests.Session | None = None) -> Path:
    """
    High-level helper: fetch data, format references, create chart, and fill the template.
    """
    publications = fetch_publications(person_id, session=session)
    references, category_counts = format_references(publications)

    chart_path = output_dir / "category_pie_chart.png"
    generate_category_chart(category_counts, chart_path)

    context = build_template_context(references, category_counts, chart_path)
    template = template_path or resolve_template_path()
    output_path = output_dir / f"{person_id}_arsrapport.docx"
    return render_report(context, template, output_path, chart_path)


def save_publications_to_json(publications: Iterable[dict], output_path: Path) -> Path:
    """Persist raw publications to disk for debugging or offline work."""
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open("w", encoding="utf-8") as fp:
        json.dump(list(publications), fp, ensure_ascii=False, indent=2)
    return output_path

