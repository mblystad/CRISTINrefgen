from pathlib import Path

import pytest

from report_generator import (
    build_reference,
    build_template_context,
    format_references,
    generate_category_chart,
)


def sample_record(category: str = "Journal article", year: int = 2024) -> dict:
    return {
        "contributors": {"preview": [{"surname": "Doe", "first_name": "Jane"}]},
        "year_published": year,
        "title": {"en": "Sample Publication"},
        "journal": {"name": "Journal of Testing"},
        "volume": "12",
        "pages": {"from": "10", "to": "20"},
        "links": [{"url_type": "DOI", "url": "https://doi.org/10.0000/example"}],
        "category": {"name": {"en": category}},
    }


def test_build_reference_returns_readable_text():
    ref = build_reference(sample_record(category="Report"))
    assert "Doe, Jane (2024)" in ref.text
    assert "Report" == ref.category
    assert ref.year == "2024"


def test_format_references_counts_categories_and_sorts():
    records = [sample_record(category="Article", year=2022), sample_record(category="Article", year=2023)]
    references, counts = format_references(records)
    assert counts == {"Article": 2}
    # Sorted newest first
    assert references[0].startswith("Doe, Jane (2023)")


def test_generate_category_chart_creates_file(tmp_path: Path):
    counts = {"Journal article": 2, "Book": 1}
    output = tmp_path / "chart.png"
    result = generate_category_chart(counts, output)
    assert result.exists()
    assert result == output


def test_build_template_context_tracks_totals(tmp_path: Path):
    chart = tmp_path / "chart.png"
    chart.touch()
    refs = ["ref1", "ref2"]
    counts = {"Journal article": 1, "Book": 1}
    context = build_template_context(refs, counts, chart)
    assert context["total_entries"] == 2
    assert len(context["category_counts"]) == 2
    assert context["category_counts"][0]["name"] in counts
