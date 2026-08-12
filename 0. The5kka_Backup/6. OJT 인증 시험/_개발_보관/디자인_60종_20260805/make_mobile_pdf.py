from pathlib import Path

from PIL import Image
from reportlab.lib.utils import ImageReader
from reportlab.pdfgen import canvas


ROOT = Path(__file__).resolve().parent
OUTPUT = ROOT / "output" / "pdf" / "OJT_디자인_60종_핸드폰용.pdf"
IMAGES = [
    *(ROOT / "comparison" / f"screens_{i}.png" for i in range(1, 6)),
    *(ROOT / "comparison" / f"palette_{i}.png" for i in range(1, 3)),
    *(ROOT / "comparison" / f"buttons_{i}.png" for i in range(1, 3)),
]


def main() -> None:
    missing = [str(path) for path in IMAGES if not path.exists()]
    if missing:
        raise FileNotFoundError("Missing comparison images: " + ", ".join(missing))

    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    pdf = canvas.Canvas(str(OUTPUT))
    pdf.setTitle("OJT Exam Maker Mobile Design Catalog")
    for image_path in IMAGES:
        with Image.open(image_path) as image:
            width, height = image.size
        page_width = 900.0
        page_height = page_width * height / width
        pdf.setPageSize((page_width, page_height))
        pdf.drawImage(ImageReader(str(image_path)), 0, 0, width=page_width, height=page_height)
        pdf.showPage()
    pdf.save()
    print(OUTPUT)


if __name__ == "__main__":
    main()
