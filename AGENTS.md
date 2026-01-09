Project overview

This repository contains an application that generates an annual activity report
(Årsrapport) for a researcher based on their NVA ID. It fetches
publication data from the NVA/Cristin API, categorizes the publications
(monographs, articles, anthologies, book reviews, etc.), and compiles them
into a Word document using a template. It also summarises other activities
(research participation, dissemination, supervision) as specified in the
template. A minimal
Streamlit user interface allows users to input their NVA ID, trigger
report generation, and download the finished Word document.

Setup commands
# Create a Python virtual environment (recommended)
python3 -m venv .venv
source .venv/bin/activate

# Install dependencies
pip install -r requirements.txt

# Alternatively, install them manually:
# pip install python-docx docxtpl requests matplotlib streamlit pandas

Running the application

To start the Streamlit interface, run:

streamlit run app.py


This will launch a local web application in your browser. Enter a CRISTIN
person ID and click the "Generate Report" button. When complete, a link to
download the generated Word report will appear.

Project structure

app.py – Defines the Streamlit interface and glues together the report
generation functions.

report_generator.py – Contains the logic for interacting with the CRISTIN
API, formatting APA‑style references, and
filling the Word template via python-docx/docxtpl.

templates/Aarsrapport-plan_MAL.docx – The Word template for the annual
report. It includes numbered sections for research work, publishing,
dissemination, and supervision. You should replace the example text with
Jinja-style placeholders (e.g., {{published_articles_niva2}}) and ensure
report_generator.py fills them correctly.

requirements.txt – Lists Python dependencies.

tests/ – Unit tests (add with pytest).

Code style guidelines

Follow PEP 8
 for formatting.

Use descriptive variable and function names.

Add type hints and docstrings to all public functions.

Keep functions small and focused. Break complex tasks into smaller helpers.

Avoid hard‑coding values; use configuration constants where appropriate.

Testing instructions

Install testing dependencies (e.g., pytest) if not already present.

Write tests under the tests/ directory. Name test files test_*.py.

Use pytest to run all tests:

pytest -q


Provide unit tests for:

Parsing publication data into categories.

Formatting APA references.

Filling the Word template with sample data.

Streamlit component behaviour (can be tested via streamlit testing tools).

Additional notes

The application interacts with the CRISTIN API, which may have rate limits
or authentication requirements. Review the API documentation and store any
credentials securely (e.g., in environment variables).

When modifying the template, maintain section numbering and headings as in
the original. Keep the structure consistent with the sample provided in the
templates folder.

To embed charts in the Word document, save them as images (e.g., PNG) and
insert them using python-docx or docxtpl.