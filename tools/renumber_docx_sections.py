from pathlib import Path
import re

from docx import Document


ROOT = Path(__file__).resolve().parents[1]
SOURCE = ROOT / "readme_humdial_updated.docx"
OUTPUT = ROOT / "readme_humdial_numbered.docx"


def main():
    document = Document(SOURCE)
    counter = None
    updated = 0

    for paragraph in document.paragraphs:
        if paragraph.style.name == "Heading 1":
            counter = 0
            continue
        if counter is None or not paragraph.text.rstrip().endswith("Paper"):
            continue

        counter += 1
        prefix = re.sub(r"^\d+\.\s*", "", paragraph.runs[0].text)
        paragraph.runs[0].text = f"{counter}. {prefix}"
        updated += 1

    document.save(OUTPUT)
    print(f"Renumbered {updated} paper entries in {OUTPUT}")


if __name__ == "__main__":
    main()
