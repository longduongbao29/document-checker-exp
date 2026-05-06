from __future__ import annotations

from typing import Any, Dict, Iterable, List, Optional

from docx import Document
from docx.document import Document as DocxDocument
from docx.oxml.ns import qn
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.text.paragraph import Paragraph as DocxParagraph

from .models import Image, Paragraph, Table


def iter_block_items(parent: DocxDocument) -> Iterable[Paragraph | Table]:
    parent_elm = parent.element.body
    for child in parent_elm.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)


class DocxParser:
    def __init__(self, file_path: str) -> None:
        self.file_path = file_path
        self.doc: Optional[DocxDocument] = None

    def load_document(self) -> DocxDocument:
        self.doc = Document(self.file_path)
        return self.doc

    def parse(self) -> List[Paragraph | Table | Image]:
        doc = self.doc or self.load_document()
        blocks: List[Paragraph | Table | Image] = []
        order = 0
        for block in iter_block_items(doc):
            if isinstance(block, Paragraph):
                paragraph_style = self._paragraph_style(block)
                block.style_data = paragraph_style
                block.meta = {
                    "has_hyperlink": self._has_hyperlink(block),
                    "run_count": len(block.runs),
                }
                self._assign_identity(block, order)
                blocks.append(block)
                order += 1

                images = self._extract_images_from_paragraph(block, paragraph_style)
                for image in images:
                    self._assign_identity(image, order)
                    blocks.append(image)
                    order += 1
            elif isinstance(block, Table):
                self._fill_table_meta(block)
                self._assign_identity(block, order)
                blocks.append(block)
                order += 1
        return blocks
    def _assign_identity(self, block: Paragraph | Table | Image, order: int) -> None:
        block.order = order
        block.block_id = f"b{order:04d}"

    def _paragraph_style(self, paragraph: DocxParagraph) -> Dict[str, Any]:
        return {
            "paragraph": {
                "style_name": paragraph.style.name if paragraph.style else None,
                "alignment": getattr(paragraph.alignment, "name", None),
                "is_list": self._is_list_paragraph(paragraph),
            },
            "runs": [self._run_style(run) for run in paragraph.runs],
        }

    def _run_style(self, run: Any) -> Dict[str, Any]:
        font = run.font
        color = None
        if font.color is not None and font.color.rgb is not None:
            color = str(font.color.rgb)
        size = font.size.pt if font.size is not None else None
        underline = run.underline
        if underline is not None and not isinstance(underline, bool):
            underline = True

        return {
            "text": run.text,
            "style_name": run.style.name if run.style else None,
            "bold": run.bold,
            "italic": run.italic,
            "underline": underline,
            "font_name": font.name,
            "font_size": size,
            "color": color,
        }

    def _extract_images_from_paragraph(
        self, paragraph: Paragraph, paragraph_style: Dict[str, Any]
    ) -> List[Image]:
        if self.doc is None:
            return []
        images: List[Image] = []
        for run in paragraph.runs:
            blips = run._element.xpath(".//a:blip")
            for blip in blips:
                rel_id = blip.get(qn("r:embed"))
                if not rel_id:
                    continue
                image_part = self.doc.part.related_parts.get(rel_id)
                if image_part is None:
                    continue
                image = Image.from_image_part(image_part, rel_id)
                image.style_data = {"paragraph": paragraph_style.get("paragraph")}
                images.append(image)
        return images

    def _fill_table_meta(self, table: Table) -> None:
        rows = [[cell.text.strip() for cell in row.cells] for row in table.rows]
        table.text_content = " ".join(" ".join(row) for row in rows).strip()
        table.style_data = {
            "table": {
                "style_name": table.style.name if table.style else None,
            }
        }
        table.meta = {
            "rows": rows,
            "row_count": len(rows),
            "column_count": len(rows[0]) if rows else 0,
        }

    def _has_hyperlink(self, paragraph: DocxParagraph) -> bool:
        return bool(paragraph._p.xpath(".//w:hyperlink"))

    def _is_list_paragraph(self, paragraph: DocxParagraph) -> bool:
        return bool(paragraph._p.xpath("./w:pPr/w:numPr"))
