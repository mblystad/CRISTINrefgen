from __future__ import annotations

from datetime import date
from pathlib import Path
from uuid import uuid4

import streamlit as st

from report_generator import generate_report, resolve_template_path


st.set_page_config(
    page_title="CRISTIN Annual Report Generator",
    page_icon=":page_facing_up:",
    layout="centered",
)


def persist_uploaded_template(uploaded_file) -> Path:
    output_dir = Path("reports") / "uploaded_templates"
    output_dir.mkdir(parents=True, exist_ok=True)
    suffix = Path(uploaded_file.name).suffix or ".docx"
    destination = output_dir / f"template_{uuid4().hex}{suffix}"
    destination.write_bytes(uploaded_file.getbuffer())
    return destination


def main() -> None:
    st.title("CRISTIN Annual Report Generator")
    st.write(
        "Upload a Word template, enter a CRISTIN person ID, choose the report year, "
        "and download the filled report."
    )

    default_template = None
    try:
        default_template = resolve_template_path()
        st.caption(f"Default template: {default_template.name}")
    except FileNotFoundError:
        st.warning("No default template found in templates/. Please upload a template.")

    uploaded_template = st.file_uploader(
        "Template (.docx)",
        type=["docx"],
        help="Upload the institution template with placeholders.",
    )

    person_id = st.text_input("CRISTIN person ID", placeholder="123456")
    report_year = st.number_input(
        "Report year",
        min_value=1900,
        max_value=date.today().year + 1,
        value=date.today().year,
        step=1,
    )
    report_year = int(report_year)

    generate = st.button("Generate Report")

    if generate:
        if not person_id.strip():
            st.error("Please enter a CRISTIN person ID.")
            return

        if not person_id.strip().isdigit():
            st.error("CRISTIN person ID must be numeric.")
            return

        template_path = None
        if uploaded_template is not None:
            template_path = persist_uploaded_template(uploaded_template)
        else:
            template_path = default_template

        if template_path is None:
            st.error("Template missing. Upload a template or add one to the templates folder.")
            return

        with st.spinner("Generating report..."):
            try:
                output_path = generate_report(
                    person_id.strip(),
                    report_year=report_year,
                    template_path=template_path,
                )
            except Exception as exc:  # Broad on purpose to bubble errors to the UI
                st.error(f"Failed to generate report: {exc}")
                return

        st.success("Report generated successfully.")
        with output_path.open("rb") as fp:
            st.download_button(
                label="Download report",
                data=fp,
                file_name=output_path.name,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )

        st.info(f"Saved locally to {output_path}")


if __name__ == "__main__":
    main()
