import pytest


def test_app_smoke():
    st_testing = pytest.importorskip("streamlit.testing.v1")
    app_test = st_testing.AppTest.from_file("app.py")
    app_test.run()

    assert app_test.title
    assert any(
        element.label == "NVA person ID" for element in app_test.text_input
    )
    assert any(
        element.label == "Generate Report" for element in app_test.button
    )
