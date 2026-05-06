from __future__ import annotations

import argparse
import json

from document_checker import DocxExtractor


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Write a trace log of extracted blocks to a text file."
    )
    parser.add_argument("docx_path", help="Path to the .docx file")
    parser.add_argument(
        "--no-page-map",
        action="store_true",
        help="Skip PDF page mapping step",
    )
    parser.add_argument(
        "--out",
        default="test.txt",
        help="Write trace output to a file (default: test.txt)",
    )
    parser.add_argument(
        "--json",
        action="store_true",
        help="Print blocks as JSON to stdout",
    )
    parser.add_argument(
        "--json-out",
        help="Write blocks as JSON to a file",
    )
    args = parser.parse_args()

    extractor = DocxExtractor(args.docx_path)
    blocks = extractor.extract(map_pages=not args.no_page_map)
    if args.json_out:
        payload = [block.to_dict() for block in blocks]
        with open(args.json_out, "w", encoding="utf-8") as handle:
            json.dump(payload, handle, indent=2, ensure_ascii=True)
        print(f"Wrote {len(payload)} blocks to {args.json_out}")
        return
    if args.json:
        payload = [block.to_dict() for block in blocks]
        print(json.dumps(payload, indent=2, ensure_ascii=True))
        return
    write_blocks_by_page(blocks, args.out)


def write_blocks_by_page(blocks, out_path: str) -> None:
    pages_map: dict[int, list] = {}
    unknown: list = []
    for block in blocks:
        if block.pages:
            for page in block.pages:
                pages_map.setdefault(page, []).append(block)
        else:
            unknown.append(block)

    with open(out_path, "w", encoding="utf-8") as handle:
        for page in sorted(pages_map):
            handle.write(f"===========Page {page} ================\n")
            for block in pages_map[page]:
                formatted = _format_block(block)
                if formatted:
                    handle.write(formatted)
            handle.write("\n")

        if unknown:
            handle.write("===========Page unknown ================\n")
            for block in unknown:
                formatted = _format_block(block)
                if formatted:
                    handle.write(formatted)

    print(f"Wrote {len(blocks)} blocks to {out_path}")


def _format_block(block) -> str:
    if block.type.value == "paragraph" and not block.get_text().strip():
        return ""

    lines = []
    if block.type.value != "paragraph":
        label = _display_block_label(block)
        lines.append(f"<{label}>")

    title = getattr(block, "title", None)
    if title is not None and getattr(title, "text", "").strip():
        if block.type.value not in {
            "list_of_content",
            "list_of_tables",
            "list_of_figures",
        }:
            lines.append(title.text.strip())

    if block.type.value in {
        "table",
        "list_of_content",
        "list_of_tables",
        "list_of_figures",
    }:
        content = block.to_markdown().strip()
    else:
        content = block.get_text().strip()

    if not content and block.type.value == "image":
        image_name = block.meta.get("image_name") if hasattr(block, "meta") else None
        if image_name:
            suffix = " (unrecognized)" if block.meta.get("unrecognized") else ""
            content = f"[image: {image_name}]{suffix}"

    if not content:
        return ""

    lines.append(content)
    lines.append("")
    return "\n".join(lines)


def _display_block_label(block) -> str:
    label_map = {
        "list_of_content": "Table of content",
        "list_of_tables": "List of tables",
        "list_of_figures": "List of figures",
    }
    return label_map.get(block.type.value, block.type.value.replace("_", " "))


if __name__ == "__main__":
    main()
