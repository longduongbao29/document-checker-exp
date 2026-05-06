from __future__ import annotations

from abc import ABC
from enum import Enum
from typing import Any, Dict, List, Optional, cast

from docx.image.exceptions import UnrecognizedImageError
from docx.image.image import Image as DocxImage
from docx.table import Table as DocxTable
from docx.text.paragraph import Paragraph as DocxParagraph


class BlockType(str, Enum):
    PARAGRAPH = "paragraph"
    TABLE = "table"
    IMAGE = "image"
    LIST_OF_CONTENT = "list_of_content"
    LIST_OF_TABLES = "list_of_tables"
    LIST_OF_FIGURES = "list_of_figures"


class BlockBase(ABC):
    __slots__ = ()

    def _init_block(self, block_type: BlockType) -> None:
        self.block_type = block_type
        self.block_id = ""
        self.order = 0
        self.pages = []
        self.style_data = {}
        self.meta = {}

    @property
    def type(self) -> BlockType:
        return self.block_type

    def get_text(self) -> str:
        return ""

    def to_markdown(self) -> str:
        return self.get_text()

    def to_markdown_with_style(self) -> str:
        style_dump = _style_to_string(self.style_data)
        base = self.to_markdown()
        if style_dump:
            return f"<!-- style: {style_dump} -->\n{base}"
        return base

    def to_dict(self) -> Dict[str, Any]:
        payload = {
            "id": self.block_id,
            "type": self.block_type.value,
            "order": self.order,
            "pages": list(self.pages),
            "text": self.get_text(),
            "style": self.style_data,
            "meta": self.meta,
        }
        title = getattr(self, "title", None)
        if title is not None:
            payload["title"] = title.text
        return payload


class Paragraph(DocxParagraph, BlockBase):
    __slots__ = ("block_type", "block_id", "order", "pages", "style_data", "meta")

    def __init__(self, element, parent) -> None:
        super().__init__(element, parent)
        self._init_block(BlockType.PARAGRAPH)

    def get_text(self) -> str:
        numbered = self.meta.get("numbered_text")
        if numbered:
            return numbered
        return self.text

    def style_name(self) -> str:
        style = getattr(self, "style", None)
        return getattr(style, "name", "") or ""

    def has_hyperlink(self) -> bool:
        return bool(self.meta.get("has_hyperlink"))


class Table(DocxTable, BlockBase):
    __slots__ = (
        "block_type",
        "block_id",
        "order",
        "pages",
        "style_data",
        "meta",
        "title",
        "text_content",
    )

    def __init__(self, element, parent) -> None:
        super().__init__(element, parent)
        self._init_block(BlockType.TABLE)
        self.title: Optional[Paragraph] = None
        self.text_content = ""

    def get_text(self) -> str:
        return self.text_content

    def to_markdown(self) -> str:
        rows = self.meta.get("rows", [])
        return _table_to_markdown(rows)


class Image(DocxImage, BlockBase):
    __slots__ = (
        "block_type",
        "block_id",
        "order",
        "pages",
        "style_data",
        "meta",
        "title",
    )

    def __init__(self, image_blob: bytes, filename: str, image_header) -> None:
        super().__init__(image_blob, filename, image_header)
        self._init_block(BlockType.IMAGE)
        self.title: Optional[Paragraph] = None

    @classmethod
    def from_image_part(cls, image_part, rel_id: str) -> "Image":
        import io

        try:
            instance = cast(
                Image,
                cls._from_stream(
                    io.BytesIO(image_part.blob),
                    image_part.blob,
                    filename=image_part.filename,
                ),
            )
        except UnrecognizedImageError:
            header = _UnknownImageHeader()
            filename = image_part.filename or "image.bin"
            instance = cls(image_part.blob, filename, header)
            instance.meta = {
                "image_name": filename,
                "relationship_id": rel_id,
                "unrecognized": True,
            }
            return instance
        instance.meta = {
            "image_name": image_part.filename,
            "relationship_id": rel_id,
        }
        return instance

    def to_markdown(self) -> str:
        name = self.meta.get("image_name") or "image"
        return f"![{name}]({name})"


class _UnknownImageHeader:
    content_type = "image/unknown"
    px_width = 0
    px_height = 0
    horz_dpi = 72
    vert_dpi = 72
    default_ext = "bin"


class ListBlock(BlockBase):
    __slots__ = (
        "block_type",
        "block_id",
        "order",
        "pages",
        "style_data",
        "meta",
        "title",
        "items",
    )

    def __init__(self, block_type: BlockType, title: Paragraph, items: List[Paragraph]):
        self._init_block(block_type)
        self.title = title
        self.items = [item for item in items if item.text.strip()]
        self.pages = _merge_pages([title] + self.items)
        self.style_data = {"heading": title.style_data}
        self.meta = {
            "item_count": len(self.items),
            "items": [
                {"text": item.text, "pages": list(item.pages)} for item in self.items
            ],
        }

    def get_text(self) -> str:
        return self.title.text

    def to_markdown(self) -> str:
        lines = [self.title.text]
        for item in self.items:
            text = item.text.strip()
            if text:
                lines.append(f"- {text}")
        return "\n".join(lines)


class ListOfContent(ListBlock):
    def __init__(self, title: Paragraph, items: List[Paragraph]) -> None:
        super().__init__(BlockType.LIST_OF_CONTENT, title, items)


class ListOfTables(ListBlock):
    def __init__(self, title: Paragraph, items: List[Paragraph]) -> None:
        super().__init__(BlockType.LIST_OF_TABLES, title, items)


class ListOfFigures(ListBlock):
    def __init__(self, title: Paragraph, items: List[Paragraph]) -> None:
        super().__init__(BlockType.LIST_OF_FIGURES, title, items)


def _style_to_string(style: Dict[str, Any]) -> str:
    if not style:
        return ""
    # Keep it compact and ASCII-safe.
    import json

    return json.dumps(style, ensure_ascii=True, separators=(",", ":"))


def _table_to_markdown(rows: List[List[str]]) -> str:
    if not rows:
        return ""
    header = rows[0]
    body = rows[1:] if len(rows) > 1 else []
    lines = []
    lines.append("| " + " | ".join(header) + " |")
    lines.append("|" + "|".join(["---"] * len(header)) + "|")
    for row in body:
        lines.append("| " + " | ".join(row) + " |")
    return "\n".join(lines)


def _merge_pages(paragraphs: List[Paragraph]) -> List[int]:
    pages: List[int] = []
    for paragraph in paragraphs:
        pages.extend(paragraph.pages)
    if not pages:
        return []
    start = min(pages)
    end = max(pages)
    return list(range(start, end + 1))
