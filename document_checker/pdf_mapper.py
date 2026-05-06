from __future__ import annotations

import os
import shutil
import subprocess
import tempfile
from dataclasses import dataclass
from typing import List, Optional

import re

import pdfplumber

from .models import Paragraph
from .utils import normalize_text, text_head, text_tail


@dataclass
class PdfPageText:
    page_number: int
    text: str
    normalized: str


class DocxMapper:
    def __init__(
        self,
        docx_path: str,
        pdf_path: Optional[str] = None,
        work_dir: Optional[str] = "temp",
        head_len: int = 300,
        tail_len: int = 300,
        min_head_len: int = 120,
        min_token_overlap: float = 0.7,
        min_tokens_for_overlap: int = 5,
        min_token_length: int = 3,
    ) -> None:
        self.docx_path = docx_path
        self.pdf_path = pdf_path
        self.work_dir = work_dir
        self.head_len = head_len
        self.tail_len = tail_len
        self.min_head_len = min_head_len
        self.min_token_overlap = min_token_overlap
        self.min_tokens_for_overlap = min_tokens_for_overlap
        self.min_token_length = min_token_length

    def map_paragraphs(self, paragraphs: List[Paragraph]) -> List[Paragraph]:
        pdf_path = self._ensure_pdf()
        pages = self._extract_pdf_pages(pdf_path)
        if not pages:
            return paragraphs

        current_page = 0
        for paragraph in paragraphs:
            if not paragraph.text.strip():
                paragraph.pages = [pages[current_page].page_number]
                continue

            normalized = self._normalize_paragraph_for_match(paragraph)
            head = text_head(normalized, self.head_len)
            tail = text_tail(normalized, self.tail_len)

            start = self._find_page_for_snippet(head, pages, current_page)
            if start is None and len(normalized) > self.min_head_len:
                head = text_head(normalized, self.min_head_len)
                start = self._find_page_for_snippet(head, pages, current_page)

            if start is None:
                paragraph.pages = [pages[current_page].page_number]
                continue

            end = self._find_page_for_snippet(tail, pages, start)
            if end is None:
                end = start

            paragraph.pages = [p.page_number for p in pages[start : end + 1]]
            current_page = end

        return paragraphs

    def _ensure_pdf(self) -> str:
        if self.pdf_path and os.path.exists(self.pdf_path):
            return self.pdf_path
        output_dir = self.work_dir or tempfile.mkdtemp(prefix="docx_pdf_")
        return self._convert_docx_to_pdf(output_dir)

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
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text() or ""
                normalized = normalize_text(text)
                pages.append(
                    PdfPageText(
                        page_number=page.page_number,
                        text=text,
                        normalized=normalized,
                    )
                )
        return pages

    def _find_page_for_snippet(
        self, snippet: str, pages: List[PdfPageText], start_index: int
    ) -> Optional[int]:
        if not snippet:
            return None
        for idx in range(start_index, len(pages)):
            page_text = pages[idx].normalized
            if snippet in page_text:
                return idx
            if self._token_overlap(snippet, page_text) >= self.min_token_overlap:
                return idx
        return None

    def _token_overlap(self, snippet: str, page_text: str) -> float:
        tokens = [
            token for token in snippet.split() if len(token) >= self.min_token_length
        ]
        if len(tokens) < self.min_tokens_for_overlap:
            return 0.0
        hits = sum(1 for token in tokens if token in page_text)
        return hits / len(tokens)

    def _normalize_paragraph_for_match(self, paragraph: Paragraph) -> str:
        text = paragraph.text
        if self._is_toc_entry(paragraph):
            text = self._strip_toc_page_number(text)
        return normalize_text(text)

    def _is_toc_entry(self, paragraph: Paragraph) -> bool:
        if paragraph.meta.get("has_hyperlink"):
            return True
        style_name = paragraph.style_name().casefold()
        return "toc" in style_name

    def _strip_toc_page_number(self, text: str) -> str:
        trimmed = text.strip()
        trimmed = re.sub(r"\s+[0-9]+\s*$", "", trimmed)
        trimmed = re.sub(r"\s+[ivxlcdm]+\s*$", "", trimmed, flags=re.IGNORECASE)
        return trimmed
