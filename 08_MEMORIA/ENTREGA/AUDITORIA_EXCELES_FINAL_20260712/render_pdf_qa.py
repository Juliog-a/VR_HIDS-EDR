from __future__ import annotations

import json
from pathlib import Path

import fitz


INPUT_DIR = Path(r"C:\Users\julio\Desktop\TFM\tmp\pdfs\excel_visual_qa")
OUTPUT_DIR = INPUT_DIR / "rendered"


def main() -> None:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    manifest: list[dict[str, object]] = []
    for pdf_path in sorted(INPUT_DIR.glob("*.pdf")):
        with fitz.open(pdf_path) as document:
            for page_index, page in enumerate(document):
                pixmap = page.get_pixmap(matrix=fitz.Matrix(2.0, 2.0), alpha=False)
                output_path = OUTPUT_DIR / f"{pdf_path.stem}_p{page_index + 1:02d}.png"
                pixmap.save(output_path)
                manifest.append(
                    {
                        "pdf": str(pdf_path),
                        "page": page_index + 1,
                        "png": str(output_path),
                        "width": pixmap.width,
                        "height": pixmap.height,
                        "size_bytes": output_path.stat().st_size,
                    }
                )
    manifest_path = OUTPUT_DIR / "RENDER_MANIFEST.json"
    manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({"pdfs": len(list(INPUT_DIR.glob('*.pdf'))), "pages": len(manifest), "manifest": str(manifest_path)}, ensure_ascii=False))


if __name__ == "__main__":
    main()
