from __future__ import annotations

import json
import math
import re
import sys
from pathlib import Path

import fitz
from PIL import Image, ImageDraw


def safe_name(text: str) -> str:
    return re.sub(r"[^A-Za-z0-9_-]+", "_", text).strip("_")[:80]


def main() -> int:
    if len(sys.argv) != 4:
        print("usage: render_pdf_review.py PDF OUT_DIR REPORT_JSON", file=sys.stderr)
        return 2

    pdf_path = Path(sys.argv[1]).resolve()
    out_dir = Path(sys.argv[2]).resolve()
    report_path = Path(sys.argv[3]).resolve()
    pages_dir = out_dir / "pages_100pct"
    contacts_dir = out_dir / "contact_sheets"
    focus_dir = out_dir / "focus_pages"
    for directory in (pages_dir, contacts_dir, focus_dir):
        directory.mkdir(parents=True, exist_ok=True)

    doc = fitz.open(pdf_path)
    zoom = 96 / 72
    matrix = fitz.Matrix(zoom, zoom)
    page_records: list[dict] = []
    all_text: list[str] = []
    page_paths: list[Path] = []

    for index, page in enumerate(doc):
        page_number = index + 1
        pix = page.get_pixmap(matrix=matrix, alpha=False, colorspace=fitz.csRGB)
        page_path = pages_dir / f"page_{page_number:03d}.png"
        pix.save(page_path)
        page_paths.append(page_path)

        text = page.get_text("text")
        all_text.append(text)
        words = page.get_text("words")
        blocks = page.get_text("blocks")
        rect = page.rect
        outside_text_blocks = 0
        for block in blocks:
            block_rect = fitz.Rect(block[:4])
            if (
                block_rect.x0 < rect.x0 - 1
                or block_rect.y0 < rect.y0 - 1
                or block_rect.x1 > rect.x1 + 1
                or block_rect.y1 > rect.y1 + 1
            ):
                outside_text_blocks += 1

        drawings_outside = 0
        for drawing in page.get_drawings():
            drawing_rect = drawing.get("rect")
            if drawing_rect and not rect.contains(drawing_rect):
                drawings_outside += 1

        images_outside = 0
        for image in page.get_image_info(xrefs=True):
            bbox = fitz.Rect(image.get("bbox", rect))
            if not rect.contains(bbox):
                images_outside += 1

        page_records.append(
            {
                "page": page_number,
                "width_pt": rect.width,
                "height_pt": rect.height,
                "words": len(words),
                "characters": len(text.strip()),
                "outside_text_blocks": outside_text_blocks,
                "outside_drawings": drawings_outside,
                "outside_images": images_outside,
                "image_count": len(page.get_image_info()),
            }
        )

    contact_paths: list[str] = []
    pages_per_contact = 12
    thumb_width = 520
    border = 12
    label_height = 32
    columns = 4
    rows = 3
    for batch_start in range(0, len(page_paths), pages_per_contact):
        batch = page_paths[batch_start : batch_start + pages_per_contact]
        thumbs: list[Image.Image] = []
        max_thumb_height = 0
        for page_path in batch:
            with Image.open(page_path) as image:
                thumb_height = round(image.height * thumb_width / image.width)
                thumb = image.resize((thumb_width, thumb_height), Image.Resampling.LANCZOS)
                thumbs.append(thumb.copy())
                max_thumb_height = max(max_thumb_height, thumb_height)

        canvas_width = columns * (thumb_width + 2 * border)
        canvas_height = rows * (max_thumb_height + label_height + 2 * border)
        canvas = Image.new("RGB", (canvas_width, canvas_height), "white")
        draw = ImageDraw.Draw(canvas)
        for offset, thumb in enumerate(thumbs):
            row = offset // columns
            col = offset % columns
            x = col * (thumb_width + 2 * border) + border
            y = row * (max_thumb_height + label_height + 2 * border) + border + label_height
            page_number = batch_start + offset + 1
            draw.text((x, y - label_height + 4), f"Página {page_number}", fill="black")
            canvas.paste(thumb, (x, y))
            draw.rectangle((x - 1, y - 1, x + thumb.width, y + thumb.height), outline="gray")
        end_page = batch_start + len(batch)
        contact_path = contacts_dir / f"contact_{batch_start + 1:03d}_{end_page:03d}.png"
        canvas.save(contact_path, quality=95)
        contact_paths.append(str(contact_path))

    full_text = "\n".join(all_text)
    searches = {
        "generic_sources": r"(?im)^\s*Fuente:\s*(?:elaboración|elaboración o captura)",
        "broken_references": r"(?i)Error\.\s*No se encuentra el origen de la referencia",
        "private_windows_paths": r"(?i)C:\\Users\\(?:julio|seguridad)\\",
        "private_linux_paths": r"(?i)/home/admin/",
        "private_ips": r"(?<!\d)(?:10\.\d{1,3}\.\d{1,3}\.\d{1,3}|172\.(?:1[6-9]|2\d|3[01])\.\d{1,3}\.\d{1,3}|192\.168\.\d{1,3}\.\d{1,3})(?!\d)",
        "template_residue": r"Equation Chapter|Código 5Control|Tabla 63:\s*Comparación final",
        "legacy_benchmark_scenarios": r"\b(?:BASE_SIN_VR|BASE_CON_VR|TEC_SIN_VR|TEC_CON_VR)\b",
        "legacy_wazuh_ruleset": r"tfm_wazuh_custom_rules_v1\.xml",
        "legacy_379_alerts": r"379\s+(?:alertas CLIENT_EVENT|alerts)",
    }
    search_counts = {key: len(re.findall(pattern, full_text)) for key, pattern in searches.items()}

    focus_terms = [
        "Tabla 52:",
        "Tabla 54:",
        "Tabla 56:",
        "Tabla 65:",
        "Tabla 68:",
        "benchmark",
        "Wazuh base",
        "Conclusiones",
        "Referencias bibliográficas",
        "Anexo M.",
        "Declaración sobre el uso de inteligencia artificial",
    ]
    focus_hits: dict[str, list[int]] = {}
    focus_pages: set[int] = {1, len(doc)}
    for term in focus_terms:
        hits = [i + 1 for i, text in enumerate(all_text) if term.casefold() in text.casefold()]
        focus_hits[term] = hits
        focus_pages.update(hits)

    for page_number in sorted(focus_pages):
        source = page_paths[page_number - 1]
        target = focus_dir / f"page_{page_number:03d}.png"
        if not target.exists():
            target.write_bytes(source.read_bytes())

    low_content_pages = [
        record for record in page_records if record["words"] < 25 and record["page"] not in {1, len(doc)}
    ]
    almost_empty_pages = [
        record for record in page_records if record["words"] < 8 and record["page"] not in {1, len(doc)}
    ]
    outside_pages = [
        record
        for record in page_records
        if record["outside_text_blocks"] or record["outside_drawings"] or record["outside_images"]
    ]

    report = {
        "pdf": str(pdf_path),
        "pages": len(doc),
        "render_dpi": 96,
        "page_png_count": len(page_paths),
        "contact_sheets": contact_paths,
        "focus_pages": sorted(focus_pages),
        "focus_hits": focus_hits,
        "search_counts": search_counts,
        "low_content_pages": low_content_pages,
        "almost_empty_pages": almost_empty_pages,
        "outside_page_findings": outside_pages,
        "page_records": page_records,
    }
    report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({k: report[k] for k in (
        "pages",
        "render_dpi",
        "page_png_count",
        "contact_sheets",
        "focus_pages",
        "search_counts",
        "low_content_pages",
        "almost_empty_pages",
        "outside_page_findings",
    )}, ensure_ascii=False, indent=2))
    doc.close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
