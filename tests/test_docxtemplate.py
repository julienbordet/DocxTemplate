import docxtemplate.docxtemplate
from docxtemplate.docxtemplate import DocxTemplate
import pytest
import time
import os


def test_object_creation():
    dt = DocxTemplate("tests/test_data/Letter-template.docx",
                      "tests/test_data/Data-Example.xlsx",
                      "tests/test_target",
                      "")
    assert isinstance(dt, DocxTemplate)


def test_object_creation_exception():
    with pytest.raises(Exception):
        dt = DocxTemplate("invalid_file"
                          "tests/test_data/Data-Example.xlsx",
                          "tests/test_target",
                          "")

        dt = DocxTemplate("tests/test_data/Letter-template.docx",
                          "invalid_file"
                          "tests/test_target",
                          "")

        dt = DocxTemplate("tests/test_data/Letter-template.docx",
                          "tests/test_data/Data-Example.xlsx",
                          "invalid_file"
                          "")


def test_object_creation_invalid_xlsx():
    with pytest.raises(Exception):
        dt = DocxTemplate("tests/test_data/Letter-template.docx",
                          "tests/test_data/Fake-Data-Example.xlsx",
                          "tests/test_target",
                          "")


def test_object_creation_invalid_docx():
    with pytest.raises(Exception):
        dt = DocxTemplate("tests/test_data/Fake-Letter-template.docx",
                          "tests/test_data/Data-Example.xlsx",
                          "tests/test_target",
                          "")


def test_real_data_load_validation():
    expected_data = {
        'Moran': {'Nom': 'Moran', 'Prénom': 'Bob', 'Prime': '3000', 'Augmentation': '2'},
        'Laverdure': {'Nom': 'Laverdure', 'Prénom': 'Tanguy', 'Prime': '4000', 'Augmentation': '3'}
    }

    dt = DocxTemplate("tests/test_data/Letter-template.docx",
                      "tests/test_data/Data-Example.xlsx",
                      "tests/test_target",
                      "")
    dt.get_info()

    assert dt.data == expected_data


def test_generate_docx():
    epoch = str(int(round(time.time(), 0)))
    local_prefix = f"test_generate_docx-{epoch}"

    dt = DocxTemplate("tests/test_data/Letter-template.docx",
                      "tests/test_data/Data-Example.xlsx",
                      "tests/test_target",
                      local_prefix)
    dt.get_info()
    file_name = dt.generate_docx('Moran')

    assert os.path.isfile(file_name)
