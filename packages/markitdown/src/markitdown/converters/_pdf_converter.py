import sys
import io

from typing import BinaryIO, Any


from .._base_converter import (
    DocumentConverter,
    DocumentConverterResult,
    DocumentPage,
    PagedDocumentConverterResult,
)
from .._stream_info import StreamInfo
from .._exceptions import MissingDependencyException, MISSING_DEPENDENCY_MESSAGE


# Try loading optional (but in this case, required) dependencies
# Save reporting of any exceptions for later
_dependency_exc_info = None
try:
    import pdfminer
    import pdfminer.high_level
except ImportError:
    # Preserve the error and stack trace for later
    _dependency_exc_info = sys.exc_info()


ACCEPTED_MIME_TYPE_PREFIXES = [
    "application/pdf",
    "application/x-pdf",
]

ACCEPTED_FILE_EXTENSIONS = [".pdf"]


class PdfConverter(DocumentConverter):
    """
    Converts PDFs to Markdown. Most style information is ignored, so the results are essentially plain-text.
    """

    def accepts(
        self,
        file_stream: BinaryIO,
        stream_info: StreamInfo,
        **kwargs: Any,  # Options to pass to the converter
    ) -> bool:
        mimetype = (stream_info.mimetype or "").lower()
        extension = (stream_info.extension or "").lower()

        if extension in ACCEPTED_FILE_EXTENSIONS:
            return True

        for prefix in ACCEPTED_MIME_TYPE_PREFIXES:
            if mimetype.startswith(prefix):
                return True

        return False

    def convert(
        self,
        file_stream: BinaryIO,
        stream_info: StreamInfo,
        **kwargs: Any,  # Options to pass to the converter
    ) -> DocumentConverterResult:
        # Check the dependencies
        if _dependency_exc_info is not None:
            raise MissingDependencyException(
                MISSING_DEPENDENCY_MESSAGE.format(
                    converter=type(self).__name__,
                    extension=".pdf",
                    feature="pdf",
                )
            ) from _dependency_exc_info[
                1
            ].with_traceback(  # type: ignore[union-attr]
                _dependency_exc_info[2]
            )

        assert isinstance(file_stream, io.IOBase)  # for mypy
        return_pages = kwargs.get("return_pages", False)

        markdown_text = pdfminer.high_level.extract_text(file_stream)
        if not return_pages:
            return DocumentConverterResult(markdown=markdown_text)

        # Build page list with numbers
        pages: list[DocumentPage] = []
        file_stream.seek(0)
        page_count = sum(1 for _ in pdfminer.high_level.extract_pages(file_stream))
        for i in range(page_count):
            file_stream.seek(0)
            page_text = pdfminer.high_level.extract_text(file_stream, page_numbers=[i])
            pages.append(DocumentPage(page_number=i + 1, markdown=page_text))

        return PagedDocumentConverterResult(pages=pages)
