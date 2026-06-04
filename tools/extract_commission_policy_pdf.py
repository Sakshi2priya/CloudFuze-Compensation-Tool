"""
Extract plain text from the commission policy PDF to help transcribe numbers into
policy/commission_policy.json (tables that are images in the PDF must be typed manually).

Usage:
  python tools/extract_commission_policy_pdf.py path/to/policy.pdf
  python tools/extract_commission_policy_pdf.py path/to/policy.pdf -o policy/_extracted.txt
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path


def main() -> int:
    parser = argparse.ArgumentParser(description="Extract text from a commission policy PDF.")
    parser.add_argument("pdf_path", type=Path, help="Path to the PDF file")
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        default=None,
        help="Write extracted text to this file (UTF-8)",
    )
    args = parser.parse_args()
    pdf = args.pdf_path
    if not pdf.is_file():
        print(f"File not found: {pdf}", file=sys.stderr)
        return 1

    try:
        from pypdf import PdfReader
    except ImportError:
        print("Install dependency: pip install pypdf", file=sys.stderr)
        return 1

    reader = PdfReader(str(pdf))
    chunks: list[str] = []
    for i, page in enumerate(reader.pages):
        t = page.extract_text()
        chunks.append(t or "")
    text = "\n\n--- page break ---\n\n".join(chunks)

    if args.output:
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(text, encoding="utf-8")
        print(f"Wrote {args.output.resolve()}")
    else:
        print(text)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
