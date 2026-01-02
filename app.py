from pathlib import Path

import streamlit as st

from report_generator import generate_report, resolve_template_path


st.set_page_config(page_title="CRISTIN Annual Report Generator", page_icon="📄", layout="centered")


def main() -> None:
    st.title("CRISTIN Annual Report Generator")
    st.write(
        "Enter a CRISTIN person ID to fetch publications, build the annual report, "
        "and download the finished Word document."
    )

    template_path = None
    try:
        template_path = resolve_template_path()
        st.success(f"Template found: {template_path}")
    except FileNotFoundError as exc:
        st.error(str(exc))

    person_id = st.text_input("CRISTIN person ID", placeholder="123456")
    generate = st.button("Generate Report")

    if generate:
        if not person_id.strip():
            st.error("Please enter a CRISTIN person ID.")
            return

        if template_path is None:
            st.error("Template missing. Add the template file to the templates folder and try again.")
            return

        with st.spinner("Generating report..."):
            try:
                output_path = generate_report(person_id.strip(), template_path=template_path)
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
