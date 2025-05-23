import os
import pytest

from markitdown import MarkItDown, PagedDocumentConverterResult

from .test_module_misc import PDF_TEST_STRINGS

TEST_FILES_DIR = os.path.join(os.path.dirname(__file__), "test_files")


def test_pdf_return_pages():
    markitdown = MarkItDown()
    result = markitdown.convert(os.path.join(TEST_FILES_DIR, "test.pdf"), return_pages=True)
    assert isinstance(result, PagedDocumentConverterResult)
    assert len(result.pages) == 1
    assert result.pages[0].page_number == 1
    for expected in PDF_TEST_STRINGS:
        assert expected in result.pages[0].markdown
