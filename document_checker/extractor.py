from __future__ import annotations

from typing import List, Optional

from .docx_parser import DocxParser
from .models import Image, ListOfContent, ListOfFigures, ListOfTables, Paragraph, Table
from .organizer import BlockOrganizer
from .pdf_mapper import DocxMapper


class DocxExtractor:
    def __init__(
        self,
        docx_path: str,
        pdf_path: Optional[str] = None,
        work_dir: Optional[str] = None,
    ) -> None:
        self.docx_path = docx_path
        self.parser = DocxParser(docx_path)
        self.mapper = DocxMapper(docx_path, pdf_path=pdf_path, work_dir=work_dir)
        self.organizer = BlockOrganizer()

    def extract(
        self, map_pages: bool = True
    ) -> List[Paragraph | Table | Image | ListOfContent | ListOfTables | ListOfFigures]:
        blocks = self.parser.parse()
        paragraphs = [block for block in blocks if isinstance(block, Paragraph)]
        if map_pages:
            self.mapper.map_paragraphs(paragraphs)
        merged = self.organizer.organize(blocks)
        if map_pages:
            self.mapper.refine_table_pages(merged)
        return merged
