import sys

from typing import BinaryIO, Any

from ._html_converter import HtmlConverter
from ..converter_utils.docx.pre_process import pre_process_docx
from .._base_converter import DocumentConverterResult, PageInfo
from .._stream_info import StreamInfo
from .._exceptions import MissingDependencyException, MISSING_DEPENDENCY_MESSAGE

# Try loading optional (but in this case, required) dependencies
# Save reporting of any exceptions for later
_dependency_exc_info = None
try:
    import mammoth
    import docx
except ImportError:
    # Preserve the error and stack trace for later
    _dependency_exc_info = sys.exc_info()


ACCEPTED_MIME_TYPE_PREFIXES = [
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
]

ACCEPTED_FILE_EXTENSIONS = [".docx"]


class DocxConverter(HtmlConverter):
    """
    Converts DOCX files to Markdown. Style information (e.g.m headings) and tables are preserved where possible.
    """

    def __init__(self):
        super().__init__()
        self._html_converter = HtmlConverter()

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
        # Check: the dependencies
        if _dependency_exc_info is not None:
            raise MissingDependencyException(
                MISSING_DEPENDENCY_MESSAGE.format(
                    converter=type(self).__name__,
                    extension=".docx",
                    feature="docx",
                )
            ) from _dependency_exc_info[
                1
            ].with_traceback(  # type: ignore[union-attr]
                _dependency_exc_info[2]
            )

        # Check if page extraction is requested
        extract_pages = kwargs.get("extract_pages", False)

        style_map = kwargs.get("style_map", None)
        pre_process_stream = pre_process_docx(file_stream)

        if extract_pages:
            # Note: DOCX files don't have fixed pages like PDFs.
            # Page breaks depend on rendering settings (margins, font size, etc.)
            no_text_delimeter = "<No text>"
            number_of_chracters_delimeter = 20

            def find_paragraph_soft_page_breaks(doc_stream: BinaryIO) -> list[str]:
                """
                Extracts and returns the text fragments following soft page breaks in a DOCX document.
                """
                doc = docx.Document(doc_stream)
                soft_page_breaks = []

                for x in [p.rendered_page_breaks for p in doc.paragraphs]:
                    if x is not []:
                        for y in x:
                            if isinstance(y, docx.text.pagebreak.RenderedPageBreak):
                                following_text = (
                                    y.following_paragraph_fragment.text
                                    if y.following_paragraph_fragment
                                    else no_text_delimeter
                                )
                                soft_page_breaks.append(following_text[:number_of_chracters_delimeter])
                return soft_page_breaks

            def split_markdown_by_following_text(
                markdown_text: str, following_texts: list[str]
            ) -> list[str]:
                """
                Splits the given markdown text into segments based on specified following text delimiters.
                """
                pages = []
                current_position = 0
                page_number = 1

                for following_text in following_texts:
                    end_index = markdown_text.find(following_text, current_position)
                    if end_index == -1:
                        continue

                    content = markdown_text[current_position:end_index]
                    page_info = PageInfo(page_number=page_number, content=content)
                    pages.append(page_info)
                    current_position = end_index
                    page_number += 1

                if current_position < len(markdown_text):
                    remaining_content = markdown_text[current_position:]
                    remaining_page = PageInfo(page_number=page_number, content=remaining_content)
                    pages.append(remaining_page)

                return pages

            html_result = mammoth.convert_to_html(pre_process_stream, style_map=style_map)
            result = self._html_converter.convert_string(html_result.value, **kwargs)

            soft_page_breaks = find_paragraph_soft_page_breaks(pre_process_stream)

            pages = split_markdown_by_following_text(result.markdown, soft_page_breaks)

            return DocumentConverterResult(
                markdown=result.markdown,
                title=result.title,
                pages=pages,
            )

        else:

            # Convert to HTML
            html_result = mammoth.convert_to_html(pre_process_stream, style_map=style_map)
            # Convert HTML to markdown
            result = self._html_converter.convert_string(html_result.value, **kwargs)
            pages = None
            return DocumentConverterResult(
                markdown=result.markdown, title=result.title, pages=pages
            )
