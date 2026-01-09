# CRISTIN Annual Report Generator

This project generates a filled annual report (Aarsrapport) from a CRISTIN/NVA person ID and a Word template. The Streamlit UI lets a non-coder upload the institution template, select the report year, and download the completed report.

## Features

- Fetches publication data from the CRISTIN API for a given person ID.
- Filters publications by a user-selected year.
- Formats APA-style references and groups them by publication category.
- Fills a Word template (docxtpl) and returns the finished report.
- Streamlit UI for quick, no-code use, including optional manual activity fields.

## Requirements

- Python 3.8+

Install dependencies:

```bash
pip install -r requirements.txt
```

## Running the application

```bash
streamlit run app.py
```

Open the local Streamlit page, upload your template (or use the default template in `templates/`), enter the CRISTIN/NVA person ID, select the report year, and generate the report. The generated report is saved under `reports/` and can be downloaded directly from the UI.

## Template placeholders

The default template in `templates/Aarsrapport-plan_MAL.docx` contains Jinja placeholders that this app fills. Custom templates must include the same placeholders.

Required placeholders:

- `{{ report_year }}`
- `{{ person_name }}`
- `{{ institution_name }}`
- `{{ institution_name_secondary }}`
- `{{ publisert_monografi_niva2 }}`
- `{{ publisert_monografi_niva1 }}`
- `{{ publisert_artikkel_niva2 }}`
- `{{ publisert_artikkel_niva1 }}`
- `{{ publisert_antologi_niva2 }}`
- `{{ publisert_antologi_niva1 }}`
- `{{ publisert_book_review }}`
- `{{ publisert_annet }}`

Each publication placeholder is replaced with a double newline-separated list of APA-style references for the selected year.

Optional placeholders (manual entry in the UI):

- `{{ forskningsarbeid_internasjonal_deltagelse }}`
- `{{ forskningsarbeid_internasjonal_ledelse }}`
- `{{ forskningsarbeid_nasjonal_deltagelse }}`
- `{{ forskningsarbeid_nasjonal_ledelse }}`
- `{{ forskningsarbeid_innvilget_soknad }}`
- `{{ forskningsarbeid_utenlandsopphold }}`
- `{{ forskningsarbeid_innovasjon }}`
- `{{ forskningsarbeid_nasjonale_nettverk }}`
- `{{ forskningsarbeid_internasjonale_nettverk }}`
- `{{ formidling_faglig }}`
- `{{ formidling_politisk }}`
- `{{ formidling_kronikker }}`
- `{{ formidling_popularvitenskapelig }}`
- `{{ formidling_media }}`
- `{{ veiledning_phd }}`
- `{{ opponent_phd }}`
- `{{ referee_vitenskapelige_artikler }}`
- `{{ veiledning_masteroppgave }}`
- `{{ sensur_masteroppgave }}`
- `{{ professor_vurderinger }}`

Some formidling fields are also auto-populated from CRISTIN categories when available:

- `formidling_faglig`: lectures, academic lectures, posters, other presentations.
- `formidling_media`: interviews, programme participation.
- `formidling_kronikker`: feature articles, reader opinion pieces.

Manual input is appended after any auto-populated text.

## Project structure

- `app.py` - Streamlit UI.
- `report_generator.py` - CRISTIN API integration, categorization, and template rendering.
- `templates/Aarsrapport-plan_MAL.docx` - Default Word template with placeholders.
- `tests/` - Unit tests (pytest).
- `requirements.txt` - Python dependencies.

## Testing

Install test dependencies and run:

```bash
pytest -q
```

Tests cover:

- Categorizing publication data into report sections.
- Formatting APA references.
- Building the template context.

## Notes

- The CRISTIN API may have rate limits. If you see errors, retry after a short pause.
- The template is institution-standardized. Keep the structure intact and only replace content with placeholders.
