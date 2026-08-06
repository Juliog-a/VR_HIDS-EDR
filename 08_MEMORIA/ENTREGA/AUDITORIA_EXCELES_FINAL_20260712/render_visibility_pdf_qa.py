from __future__ import annotations

import argparse
import json
from pathlib import Path

import fitz


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("input_dir", type=Path)
    parser.add_argument("output_dir", type=Path)
    parser.add_argument("--scale", type=float, default=2.0)
    args = parser.parse_args()

    args.output_dir.mkdir(parents=True, exist_ok=True)
    manifest: list[dict[str, object]] = []
    pdfs = sorted(args.input_dir.glob("*.pdf"))
    for pdf_path in pdfs:
        with fitz.open(pdf_path) as document:
            for page_index, page in enumerate(document):
                pixmap = page.get_pixmap(
                    matrix=fitz.Matrix(args.scale, args.scale), alpha=False
                )
                output_path = (
                    args.output_dir / f"{pdf_path.stem}_p{page_index + 1:02d}.png"
                )
                pixmap.save(output_path)
                manifest.append(
                    {
                        "pdf": str(pdf_path),
                        "page": page_index + 1,
                        "png": str(output_path),
                        "width": pixmap.width,
                        "height": pixmap.height,
                        "bytes": output_path.stat().st_size,
                    }
                )

    manifest_path = args.output_dir / "RENDER_MANIFEST.json"
    manifest_path.write_text(
        json.dumps(manifest, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )
    print(
        json.dumps(
            {"pdfs": len(pdfs), "pages": len(manifest), "manifest": str(manifest_path)},
            ensure_ascii=False,
        )
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
