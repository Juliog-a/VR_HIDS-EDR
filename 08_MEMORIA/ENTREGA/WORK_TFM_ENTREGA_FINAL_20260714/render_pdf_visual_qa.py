import argparse
import json
from pathlib import Path

import fitz
from PIL import Image, ImageChops, ImageDraw, ImageFont


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("pdf")
    parser.add_argument("output_dir")
    args = parser.parse_args()

    pdf_path = Path(args.pdf).resolve()
    out_dir = Path(args.output_dir).resolve()
    pages_dir = out_dir / "pages"
    sheets_dir = out_dir / "contact_sheets"
    pages_dir.mkdir(parents=True, exist_ok=True)
    sheets_dir.mkdir(parents=True, exist_ok=True)

    doc = fitz.open(pdf_path)
    page_records = []
    rendered_paths = []
    suspicious_terms = (
        "Error! Reference source not found",
        "No se encuentra el origen de la referencia",
        "Bookmark not defined",
        "[[REF:",
        "[[SEQ_",
        "09_Benchmark_Plan",
        "v15.21",
    )

    for index, page in enumerate(doc):
        pix = page.get_pixmap(matrix=fitz.Matrix(2, 2), alpha=False)
        image_path = pages_dir / f"page_{index + 1:03d}.png"
        pix.save(image_path)
        rendered_paths.append(image_path)

        text = page.get_text("text")
        blocks = page.get_text("blocks")
        page_rect = page.rect
        outside = []
        for block in blocks:
            x0, y0, x1, y1 = block[:4]
            if x0 < -1 or y0 < -1 or x1 > page_rect.width + 1 or y1 > page_rect.height + 1:
                outside.append([round(x0, 2), round(y0, 2), round(x1, 2), round(y1, 2)])

        with Image.open(image_path) as im:
            rgb = im.convert("RGB")
            bg = Image.new("RGB", rgb.size, "white")
            diff = ImageChops.difference(rgb, bg)
            bbox = diff.getbbox()
            content_ratio = 0.0
            if bbox:
                content_ratio = ((bbox[2] - bbox[0]) * (bbox[3] - bbox[1])) / (rgb.width * rgb.height)

        page_records.append(
            {
                "page": index + 1,
                "width_points": round(page_rect.width, 2),
                "height_points": round(page_rect.height, 2),
                "text_chars": len(text.strip()),
                "content_bbox_ratio": round(content_ratio, 5),
                "outside_text_blocks": outside,
                "suspicious_terms": [term for term in suspicious_terms if term.lower() in text.lower()],
                "text_preview": " ".join(text.split())[:220],
            }
        )

    font = ImageFont.load_default()
    per_sheet = 20
    cols = 5
    rows = 4
    thumb_w, thumb_h = 238, 336
    label_h = 22
    for sheet_index in range((len(rendered_paths) + per_sheet - 1) // per_sheet):
        selected = rendered_paths[sheet_index * per_sheet : (sheet_index + 1) * per_sheet]
        sheet = Image.new("RGB", (cols * thumb_w, rows * (thumb_h + label_h)), "#d8dde6")
        draw = ImageDraw.Draw(sheet)
        for cell, image_path in enumerate(selected):
            with Image.open(image_path) as im:
                page_im = im.convert("RGB")
                page_im.thumbnail((thumb_w - 8, thumb_h - 8), Image.Resampling.LANCZOS)
                x = (cell % cols) * thumb_w + (thumb_w - page_im.width) // 2
                y = (cell // cols) * (thumb_h + label_h) + 4
                sheet.paste(page_im, (x, y))
                page_number = sheet_index * per_sheet + cell + 1
                label = f"Pagina {page_number}"
                draw.rectangle(
                    [cell % cols * thumb_w, (cell // cols + 1) * thumb_h + cell // cols * label_h,
                     (cell % cols + 1) * thumb_w, (cell // cols + 1) * (thumb_h + label_h)],
                    fill="#1f4e79",
                )
                draw.text(
                    (cell % cols * thumb_w + 8, (cell // cols + 1) * thumb_h + cell // cols * label_h + 5),
                    label,
                    fill="white",
                    font=font,
                )
        sheet.save(sheets_dir / f"contact_{sheet_index + 1:02d}.png")

    report = {
        "pdf": str(pdf_path),
        "pages": len(doc),
        "rendered_pages": len(rendered_paths),
        "contact_sheets": (len(rendered_paths) + per_sheet - 1) // per_sheet,
        "blank_or_nearly_blank_pages": [r["page"] for r in page_records if r["text_chars"] < 30],
        "pages_with_outside_text_blocks": [r["page"] for r in page_records if r["outside_text_blocks"]],
        "pages_with_suspicious_terms": [r["page"] for r in page_records if r["suspicious_terms"]],
        "page_records": page_records,
    }
    (out_dir / "visual_qa_report.json").write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({k: v for k, v in report.items() if k != "page_records"}, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
