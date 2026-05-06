from __future__ import annotations

import difflib
import os
import shutil
import subprocess
import tempfile
from dataclasses import dataclass
from typing import Iterable, List, Optional, Sequence

import pdfplumber
from rapidfuzz import fuzz

from .models import Image, ListOfContent, ListOfFigures, ListOfTables, Paragraph, Table
from .utils import normalize_text


@dataclass
class PdfPageText:
    page_number: int
    text: str
    normalized: str
    start: int
    end: int


@dataclass
class DocxSegment:
    blocks: List[Paragraph]
    text: str
    start: int = 0
    end: int = 0


class DocxMapper:
    def __init__(
        self,
        docx_path: str,
        pdf_path: Optional[str] = None,
        work_dir: Optional[str] = None,
        short_block_threshold: int = 30,
        segment_separator: str = " ",
        fuzzy_min_score: int = 60,
    ) -> None:
        self.docx_path = docx_path
        self.pdf_path = pdf_path
        self.work_dir = work_dir
        self.short_block_threshold = short_block_threshold
        self.segment_separator = segment_separator
        self.fuzzy_min_score = fuzzy_min_score

    def map_paragraphs(self, paragraphs: List[Paragraph]) -> List[Paragraph]:
        pdf_path = self._ensure_pdf()
        pages = self._extract_pdf_pages(pdf_path)
        if not pages:
            return paragraphs

        segments, docx_full_text = self._build_docx_segments(paragraphs)
        if not docx_full_text:
            self._fill_missing_paragraph_pages(paragraphs, pages)
            return paragraphs

        pdf_full_text = "".join(page.normalized for page in pages)
        match_blocks = self._align_text(docx_full_text, pdf_full_text)

        for segment in segments:
            if not segment.text:
                continue
            pages_for_segment = []
            if match_blocks:
                mapped = self._map_docx_range_to_pdf(
                    segment.start, segment.end, match_blocks, len(pdf_full_text)
                )
                if mapped is not None:
                    pages_for_segment = self._pages_for_range(*mapped, pages)
            if not pages_for_segment:
                pages_for_segment = self._fuzzy_pages_for_segment(segment, pages)
            for block in segment.blocks:
                block.pages = self._pages_for_block(block, pages_for_segment, pages)

        self._fill_missing_paragraph_pages(paragraphs, pages)

        return paragraphs

    def _ensure_pdf(self) -> str:
        if self.pdf_path and os.path.exists(self.pdf_path):
            return self.pdf_path
        output_dir = self.work_dir or tempfile.mkdtemp(prefix="docx_pdf_")
        return self._convert_docx_to_pdf(output_dir)

    def refine_table_pages(
        self,
        blocks: Sequence,
    ) -> None:
        self._assign_non_text_block_pages(blocks)

    def _convert_docx_to_pdf(self, output_dir: str) -> str:
        libreoffice = self._find_libreoffice()
        os.makedirs(output_dir, exist_ok=True)
        result = subprocess.run(
            [
                libreoffice,
                "--headless",
                "--convert-to",
                "pdf",
                "--outdir",
                output_dir,
                self.docx_path,
            ],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            check=False,
        )
        if result.returncode != 0:
            message = result.stderr.strip() or result.stdout.strip()
            raise RuntimeError(f"LibreOffice conversion failed: {message}")

        base_name = os.path.splitext(os.path.basename(self.docx_path))[0]
        pdf_path = os.path.join(output_dir, f"{base_name}.pdf")
        if not os.path.exists(pdf_path):
            raise FileNotFoundError("Converted PDF not found")
        return pdf_path

    def _find_libreoffice(self) -> str:
        for candidate in ("libreoffice", "soffice"):
            path = shutil.which(candidate)
            if path:
                return path
        raise RuntimeError("LibreOffice not found in PATH")

    def _extract_pdf_pages(self, pdf_path: str) -> List[PdfPageText]:
        pages: List[PdfPageText] = []
        cursor = 0
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text() or ""
                normalized = normalize_text(text)
                start = cursor
                end = cursor + len(normalized)
                pages.append(
                    PdfPageText(
                        page_number=page.page_number,
                        text=text,
                        normalized=normalized,
                        start=start,
                        end=end,
                    )
                )
                cursor = end
        return pages

    def _build_docx_segments(
        self, paragraphs: List[Paragraph]
    ) -> tuple[List[DocxSegment], str]:
        segments: List[DocxSegment] = []
        pending_empty: List[Paragraph] = []

        for paragraph in paragraphs:
            normalized = normalize_text(paragraph.text)
            if not normalized:
                pending_empty.append(paragraph)
                continue

            if (
                len(normalized) < self.short_block_threshold
                and segments
                and not self._is_heading_paragraph(paragraph)
            ):
                segments[-1].blocks.extend(pending_empty)
                pending_empty = []
                segments[-1].blocks.append(paragraph)
                if segments[-1].text:
                    segments[-1].text = (
                        f"{segments[-1].text}{self.segment_separator}{normalized}"
                    )
                else:
                    segments[-1].text = normalized
                continue

            blocks = pending_empty + [paragraph]
            pending_empty = []
            segments.append(DocxSegment(blocks=blocks, text=normalized))

        if pending_empty:
            if segments:
                segments[-1].blocks.extend(pending_empty)
            else:
                segments.append(DocxSegment(blocks=pending_empty, text=""))

        cursor = 0
        parts: List[str] = []
        for segment in segments:
            if not segment.text:
                continue
            if parts:
                cursor += len(self.segment_separator)
            segment.start = cursor
            segment.end = cursor + len(segment.text)
            parts.append(segment.text)
            cursor = segment.end

        full_text = self.segment_separator.join(parts)
        return segments, full_text

    def _is_heading_paragraph(self, paragraph: Paragraph) -> bool:
        style_name = paragraph.style_name().casefold()
        return "heading" in style_name or "title" in style_name

    def _align_text(
        self, docx_text: str, pdf_text: str
    ) -> Optional[List[difflib.Match]]:
        if not docx_text or not pdf_text:
            return None
        matcher = difflib.SequenceMatcher(
            None, docx_text, pdf_text, autojunk=False
        )
        matches = [match for match in matcher.get_matching_blocks() if match.size]
        if not matches:
            return None
        return matches

    def _map_docx_range_to_pdf(
        self,
        start: int,
        end: int,
        match_blocks: List[difflib.Match],
        pdf_len: int,
    ) -> Optional[tuple[int, int]]:
        if not match_blocks:
            return None
        if end <= start:
            end = start + 1

        mapped_start = self._map_docx_index(start, match_blocks)
        mapped_end = self._map_docx_index(max(start, end - 1), match_blocks) + 1

        mapped_start = max(0, min(mapped_start, pdf_len))
        mapped_end = max(0, min(mapped_end, pdf_len))

        if mapped_end < mapped_start:
            mapped_start, mapped_end = mapped_end, mapped_start
        return mapped_start, mapped_end

    def _map_docx_index(
        self, index: int, match_blocks: List[difflib.Match]
    ) -> int:
        for match in match_blocks:
            if match.a <= index < match.a + match.size:
                return match.b + (index - match.a)

        first = match_blocks[0]
        last = match_blocks[-1]
        if index < first.a:
            return first.b
        if index >= last.a + last.size:
            return last.b + last.size

        for prev, next_match in zip(match_blocks, match_blocks[1:]):
            prev_end = prev.a + prev.size
            if prev_end <= index < next_match.a:
                return prev.b + prev.size

        return last.b + last.size

    def _pages_for_range(
        self, start: int, end: int, pages: Sequence[PdfPageText]
    ) -> List[int]:
        if end <= start:
            end = start + 1
        matches = []
        for page in pages:
            if end <= page.start or start >= page.end:
                continue
            matches.append(page.page_number)
        return matches

    def _fuzzy_pages_for_segment(
        self, segment: DocxSegment, pages: Sequence[PdfPageText]
    ) -> List[int]:
        best_page = None
        best_score = 0
        for page in pages:
            score = fuzz.partial_ratio(segment.text, page.normalized)
            if score > best_score:
                best_score = score
                best_page = page.page_number
        if best_page is None or best_score < self.fuzzy_min_score:
            return []
        return [best_page]

    def _pages_for_block(
        self,
        block: Paragraph,
        pages_for_segment: List[int],
        pages: Sequence[PdfPageText],
    ) -> List[int]:
        if not pages_for_segment:
            return []
        if len(pages_for_segment) == 1:
            return list(pages_for_segment)

        normalized = normalize_text(block.text)
        candidate_pages = [
            page for page in pages if page.page_number in pages_for_segment
        ]
        best_page, best_score, second_score = self._best_two_pages_for_text(
            normalized, candidate_pages
        )
        if best_page is None:
            return list(pages_for_segment)
        if best_score >= self.fuzzy_min_score and best_score - second_score >= 10:
            return [best_page]
        if normalized and len(normalized) < self.short_block_threshold:
            if best_score >= self.fuzzy_min_score:
                return [best_page]

        return list(pages_for_segment)

    def _best_page_for_text(
        self, text: str, pages: Sequence[PdfPageText]
    ) -> Optional[int]:
        if not text:
            return None
        best_page = None
        best_score = 0
        for page in pages:
            score = fuzz.partial_ratio(text, page.normalized)
            if score > best_score:
                best_score = score
                best_page = page.page_number
        if best_page is None or best_score < self.fuzzy_min_score:
            return None
        return best_page

    def _best_two_pages_for_text(
        self, text: str, pages: Sequence[PdfPageText]
    ) -> tuple[Optional[int], int, int]:
        if not text:
            return None, 0, 0
        best_page = None
        best_score = 0
        second_score = 0
        for page in pages:
            score = fuzz.partial_ratio(text, page.normalized)
            if score > best_score:
                second_score = best_score
                best_score = score
                best_page = page.page_number
            elif score > second_score:
                second_score = score
        return best_page, best_score, second_score

    def _fill_missing_paragraph_pages(
        self, paragraphs: Iterable[Paragraph], pages: Sequence[PdfPageText]
    ) -> None:
        if not pages:
            return
        fallback = pages[0].page_number
        last_page: Optional[int] = None
        for paragraph in paragraphs:
            if paragraph.pages:
                last_page = paragraph.pages[-1]
                continue
            if last_page is not None:
                paragraph.pages = [last_page]
            else:
                paragraph.pages = [fallback]

    def _assign_non_text_block_pages(self, blocks: Sequence) -> None:
        text_blocks = []
        for index, block in enumerate(blocks):
            if isinstance(
                block,
                (
                    Paragraph,
                    ListOfContent,
                    ListOfTables,
                    ListOfFigures,
                ),
            ) and block.pages:
                text_blocks.append((index, block))

        if not text_blocks:
            return

        for index, block in enumerate(blocks):
            if not isinstance(block, (Table, Image)):
                continue
            anchor = self._nearest_text_block(index, text_blocks, block)
            if anchor is None:
                continue
            block.pages = list(anchor.pages)

    def _nearest_text_block(
        self,
        index: int,
        text_blocks: List[tuple[int, Paragraph | ListOfContent | ListOfTables | ListOfFigures]],
        block: Table | Image,
    ) -> Optional[Paragraph | ListOfContent | ListOfTables | ListOfFigures]:
        best_block = None
        best_distance = None
        prefer_previous = isinstance(block, Table)
        for text_index, text_block in text_blocks:
            distance = abs(text_index - index)
            if best_distance is None or distance < best_distance:
                best_distance = distance
                best_block = text_block
                continue
            if distance == best_distance and prefer_previous:
                if text_index < index and best_block is not None:
                    best_block = text_block
        return best_block
