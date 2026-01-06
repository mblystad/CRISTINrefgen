from report_generator import (
    build_entries,
    build_template_context,
    classify_publication,
    extract_level,
    format_reference,
)


def sample_record(category: str = "Journal article", year: int = 2024, level: int = 2) -> dict:
    return {
        "contributors": {"preview": [{"surname": "Doe", "first_name": "Jane"}]},
        "year_published": year,
        "title": {"en": "Sample Publication"},
        "journal": {"name": "Journal of Testing", "level": str(level)},
        "volume": "12",
        "pages": {"from": "10", "to": "20"},
        "links": [{"url_type": "DOI", "url": "https://doi.org/10.0000/example"}],
        "category": {"name": {"en": category}},
    }


def test_format_reference_contains_expected_fields():
    record = sample_record()
    reference = format_reference(record)
    assert "Doe, Jane" in reference
    assert "(2024)." in reference
    assert "Sample Publication" in reference
    assert "Journal of Testing" in reference
    assert "doi.org" in reference


def test_extract_level_prefers_known_fields():
    record = sample_record(level=1)
    assert extract_level(record) == "1"


def test_classify_publication_uses_category_and_level():
    record = sample_record(category="Monograph", level=2)
    assert classify_publication(record) == "publisert_monografi_niva2"

    record = sample_record(category="Anthology", level=1)
    assert classify_publication(record) == "publisert_antologi_niva1"

    record = sample_record(category="Book review", level=1)
    assert classify_publication(record) == "publisert_book_review"


def test_build_entries_filters_by_year():
    records = [sample_record(year=2023), sample_record(year=2024)]
    entries = build_entries(records, report_year=2024)
    assert len(entries) == 1
    assert entries[0].year == "2024"


def test_build_template_context_populates_keys():
    records = [sample_record(category="Journal article", level=2)]
    entries = build_entries(records, report_year=2024)
    context = build_template_context(
        entries,
        report_year=2024,
        person_name="Jane Doe",
        institution_name="Test University",
        institution_name_secondary="",
    )

    assert context["report_year"] == "2024"
    assert context["person_name"] == "Jane Doe"
    assert "publisert_artikkel_niva2" in context
    assert "Doe, Jane" in context["publisert_artikkel_niva2"]
