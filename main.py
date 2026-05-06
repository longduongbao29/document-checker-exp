from __future__ import annotations

import argparse
import json

from document_checker import DocxExtractor


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Extract docx blocks with style and page info."
    )
    parser.add_argument("docx_path", help="Path to the .docx file")
    parser.add_argument(
        "--no-page-map",
        action="store_true",
        help="Skip PDF page mapping step",
    )
    parser.add_argument(
        "--json",
        action="store_true",
        help="Print blocks as JSON",
    )
    parser.add_argument(
        "--out",
        help="Write JSON output to a file",
    )
    args = parser.parse_args()

    extractor = DocxExtractor(args.docx_path)
    blocks = extractor.extract(map_pages=not args.no_page_map)

    payload = [block.to_dict() for block in blocks]
    if args.out:
        with open(args.out, "w", encoding="utf-8") as handle:
            json.dump(payload, handle, indent=2, ensure_ascii=True)
        print(f"Wrote {len(payload)} blocks to {args.out}")
        return
    if args.json:
        print(json.dumps(payload, indent=2, ensure_ascii=True))
        return

    for block in blocks:
        preview = block.get_text().replace("\n", " ")
        if len(preview) > 80:
            preview = preview[:77] + "..."
        pages = ",".join(str(p) for p in block.pages) if block.pages else "-"
        print(f"{block.order:04d} {block.type.value} pages={pages} text={preview!r}")


if __name__ == "__main__":
    main()