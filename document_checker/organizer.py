from __future__ import annotations

import re

from typing import List, Optional, Sequence

from .models import (
    BlockType,
    Image,
    ListOfContent,
    ListOfFigures,
    ListOfTables,
    Paragraph,
    Table,
)


class BlockOrganizer:
    def __init__(
        self,
        max_list_page: int = 10,
        max_early_paragraph_order: int = 120,
        caption_search_window: int = 2,
    ) -> None:
        self.max_list_page = max_list_page
        self.max_early_paragraph_order = max_early_paragraph_order
        self.caption_search_window = caption_search_window

    def organize(
        self, blocks: Sequence[Paragraph | Table | Image]
    ) -> List[Paragraph | Table | Image | ListOfContent | ListOfTables | ListOfFigures]:
        merged = self._merge_list_blocks(list(blocks))
        merged = self._attach_captions(merged)
        self._assign_table_pages(merged)
        self._assign_orders(merged)
        return merged

    def _merge_list_blocks(
        self, blocks: List[Paragraph | Table | Image]
    ) -> List[Paragraph | Table | Image | ListOfContent | ListOfTables | ListOfFigures]:
        merged: List[Paragraph | Table | Image | ListOfContent | ListOfTables | ListOfFigures] = []
        index = 0
        while index < len(blocks):
            block = blocks[index]
            if isinstance(block, Paragraph) and self._is_early_paragraph(block):
                list_type = self._list_block_type(block)
                if list_type is not None:
                    result = self._collect_list_entries(blocks, index)
                    if result is not None:
                        list_block, last_index = result
                        merged.append(list_block)
                        index = last_index + 1
                        continue
            merged.append(block)
            index += 1
        return merged

    def _collect_list_entries(
        self, blocks: List[Paragraph | Table | Image], start_index: int
    ) -> Optional[tuple[ListOfContent | ListOfTables | ListOfFigures, int]]:
        heading = blocks[start_index]
        if not isinstance(heading, Paragraph):
            return None

        list_type = self._list_block_type(heading)
        if list_type is None:
            return None

        items: List[Paragraph] = []
        last_index = start_index
        for idx in range(start_index + 1, len(blocks)):
            current = blocks[idx]
            if not isinstance(current, Paragraph):
                break
            if self._is_list_entry(current):
                items.append(current)
                last_index = idx
                continue
            if current.text.strip() == "" and items:
                last_index = idx
                continue
            break

        if not items:
            return None

        if list_type == BlockType.LIST_OF_TABLES:
            return ListOfTables(heading, items), last_index
        if list_type == BlockType.LIST_OF_FIGURES:
            return ListOfFigures(heading, items), last_index
        return ListOfContent(heading, items), last_index

    def _attach_captions(
        self,
        blocks: List[Paragraph | Table | Image | ListOfContent | ListOfTables | ListOfFigures],
    ) -> List[Paragraph | Table | Image | ListOfContent | ListOfTables | ListOfFigures]:
        consumed: set[str] = set()
        for index, block in enumerate(blocks):
            if isinstance(block, (Table, Image)):
                caption = self._find_nearest_caption(blocks, index, consumed, block)
                if caption is not None:
                    block.title = caption
                    if caption.pages:
                        block.pages = list(caption.pages)
                    block.meta["title_text"] = caption.text
                    consumed.add(caption.block_id)
                else:
                    fallback = self._find_nearest_paragraph(blocks, index, consumed)
                    if fallback is not None and fallback.pages:
                        block.pages = list(fallback.pages)

        filtered: List[Paragraph | Table | Image | ListOfContent | ListOfTables | ListOfFigures] = []
        for block in blocks:
            if isinstance(block, Paragraph) and block.block_id in consumed:
                continue
            filtered.append(block)
        return filtered

    def _find_nearest_caption(
        self,
        blocks: Sequence[Paragraph | Table | Image | ListOfContent | ListOfTables | ListOfFigures],
        index: int,
        consumed: set[str],
        block: Table | Image,
    ) -> Optional[Paragraph]:
        best: Optional[Paragraph] = None
        best_distance = None
        best_score = 0
        for offset in range(1, self.caption_search_window + 1):
            candidates = [index - offset, index + offset]
            for candidate_index in candidates:
                if candidate_index < 0 or candidate_index >= len(blocks):
                    continue
                candidate = blocks[candidate_index]
                if not isinstance(candidate, Paragraph):
                    continue
                if candidate.block_id in consumed:
                    continue
                score = self._caption_score(block, candidate)
                if score <= 0:
                    continue
                if (
                    score > best_score
                    or best_distance is None
                    or (score == best_score and offset < best_distance)
                    or (
                        score == best_score
                        and offset == best_distance
                        and self._prefer_candidate(block, candidate_index, index)
                    )
                ):
                    best = candidate
                    best_distance = offset
                    best_score = score
        return best

    def _find_nearest_paragraph(
        self,
        blocks: Sequence[Paragraph | Table | Image | ListOfContent | ListOfTables | ListOfFigures],
        index: int,
        consumed: set[str],
    ) -> Optional[Paragraph]:
        for offset in range(1, self.caption_search_window + 1):
            candidates = [index - offset, index + offset]
            for candidate_index in candidates:
                if candidate_index < 0 or candidate_index >= len(blocks):
                    continue
                candidate = blocks[candidate_index]
                if not isinstance(candidate, Paragraph):
                    continue
                if candidate.block_id in consumed:
                    continue
                return candidate
        return None

    def _assign_orders(
        self,
        blocks: List[Paragraph | Table | Image | ListOfContent | ListOfTables | ListOfFigures],
    ) -> None:
        for order, block in enumerate(blocks):
            block.order = order
            block.block_id = f"b{order:04d}"

    def _assign_table_pages(
        self,
        blocks: List[Paragraph | Table | Image | ListOfContent | ListOfTables | ListOfFigures],
    ) -> None:
        for index, block in enumerate(blocks):
            if not isinstance(block, Table):
                continue

            prev_paragraph = self._find_neighbor_paragraph(blocks, index, -1)
            next_paragraph = self._find_neighbor_paragraph(blocks, index, 1)

            start_page = None
            end_page = None

            if prev_paragraph and prev_paragraph.pages:
                start_page = max(prev_paragraph.pages)
            if next_paragraph and next_paragraph.pages:
                end_page = min(next_paragraph.pages)

            if start_page is None and end_page is None:
                continue
            if start_page is None:
                start_page = end_page
            if end_page is None:
                end_page = start_page

            if self._is_real_table(block):
                start_page = start_page + 1
                if end_page < start_page:
                    end_page = start_page

            block.pages = list(range(start_page, end_page + 1))

    def _find_neighbor_paragraph(
        self,
        blocks: List[Paragraph | Table | Image | ListOfContent | ListOfTables | ListOfFigures],
        index: int,
        direction: int,
    ) -> Optional[Paragraph]:
        step = 1 if direction >= 0 else -1
        idx = index + step
        while 0 <= idx < len(blocks):
            candidate = blocks[idx]
            if isinstance(candidate, Paragraph):
                return candidate
            idx += step
        return None

    def _is_real_table(self, table: Table) -> bool:
        rows = table.meta.get("rows", [])
        if len(rows) < 2:
            return False
        header_text = " ".join(cell.strip() for cell in rows[0]).casefold()
        return "callout style" not in header_text

    def _is_early_paragraph(self, paragraph: Paragraph) -> bool:
        if paragraph.pages:
            return min(paragraph.pages) <= self.max_list_page
        return paragraph.order <= self.max_early_paragraph_order

    def _list_block_type(self, paragraph: Paragraph) -> Optional[BlockType]:
        text = paragraph.text.strip().casefold()
        style_name = paragraph.style_name().casefold()

        if self._contains_list_phrase(text, "table") or self._contains_list_phrase(
            style_name, "table"
        ):
            return BlockType.LIST_OF_TABLES
        if self._contains_list_phrase(text, "figure") or self._contains_list_phrase(
            style_name, "figure"
        ):
            return BlockType.LIST_OF_FIGURES
        if "table of contents" in text or "table of content" in text:
            return BlockType.LIST_OF_CONTENT
        if text in {"contents", "content"}:
            return BlockType.LIST_OF_CONTENT
        if "toc" in style_name and not paragraph.has_hyperlink() and text:
            return BlockType.LIST_OF_CONTENT
        return None

    def _contains_list_phrase(self, text: str, noun: str) -> bool:
        pattern = rf"\blist\s+of\s+{re.escape(noun)}s?\b"
        return bool(re.search(pattern, text))

    def _is_list_entry(self, paragraph: Paragraph) -> bool:
        style_name = paragraph.style_name().casefold()
        if paragraph.has_hyperlink():
            return True
        return "toc" in style_name

    def _caption_score(self, block: Table | Image, paragraph: Paragraph) -> int:
        style_name = paragraph.style_name().casefold()
        text = paragraph.text.strip()
        text_lower = text.casefold()

        if "note" in style_name:
            return 0

        if isinstance(block, Table):
            if "table title" in style_name or "table caption" in style_name:
                return 3
            if "caption" in style_name and text_lower.startswith("table "):
                return 2
            if text_lower.startswith("table "):
                return 1
            return 0

        if "caption" in style_name and "table" not in style_name:
            return 3
        if text_lower.startswith("figure "):
            return 2
        if text_lower.startswith("fig."):
            return 1
        return 0

    def _prefer_candidate(
        self, block: Table | Image, candidate_index: int, index: int
    ) -> bool:
        if isinstance(block, Table):
            return candidate_index < index
        return candidate_index > index
