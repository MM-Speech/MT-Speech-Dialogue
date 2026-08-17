from pathlib import Path
import re

from docx import Document


ROOT = Path(__file__).resolve().parents[1]
DOCX = ROOT / "readme.docx"
README = ROOT / "README.md"

FALLBACK_LINKS = {
    "From Turn-Taking to Synchronous Dialogue: A Survey of Full-Duplex Spoken Language Models": "https://arxiv.org/abs/2509.14515",
    "A Survey of Full-Duplex Spoken Dialogue Systems: Architectural Hierarchy, Interaction Ontology, and Decision State Machine": "https://arxiv.org/abs/2606.19453",
    "EchoMind: An Interrelated Multi-level Benchmark for Evaluating Empathetic Speech Language Models": "https://arxiv.org/abs/2510.22758",
    "HumDial-EIBench: A Human-Recorded Multi-Turn Emotional Intelligence Benchmark for Audio Language Models": "https://arxiv.org/abs/2604.11594",
    "VoiceBench: Benchmarking LLM-Based Voice Assistants": "https://arxiv.org/abs/2410.17196",
}

MARKDOWN_PAPER = re.compile(
    r'^\d+\. \*\*"(?P<title>.+?)"\*\*\..*?\[\[Paper\]\((?P<url>https://[^)]+)\)\]$'
)
WORD_PAPER = re.compile(r'^"(?P<title>.+?)"\. (?P<rest>.+?)\. Paper$')


def existing_links():
    links = FALLBACK_LINKS.copy()
    for line in README.read_text(encoding="utf-8").splitlines():
        match = MARKDOWN_PAPER.match(line)
        if match:
            links[match.group("title")] = match.group("url")
    return links


def extract_sections():
    document = Document(DOCX)
    sections = []
    current = None
    for paragraph in document.paragraphs:
        if paragraph.style.name == "Heading 1":
            current = (paragraph.text.strip(), [])
            sections.append(current)
            continue
        if current is None:
            continue
        match = WORD_PAPER.match(paragraph.text.strip())
        if match:
            current[1].append(match.groupdict())
    return sections


def main():
    links = existing_links()
    sections = extract_sections()
    output = [
        "# MT-Speech-Dialogue",
        "",
        "> A curated paper list for **Multi-Turn Speech Dialogue**.",
        ">",
        "> Scope: papers directly addressing multi-turn, end-to-end, or full-duplex speech interaction. Broad domain surveys are listed separately under **General Surveys**.",
    ]
    total = 0
    for heading, papers in sections:
        output.extend(["", f"## {heading}", ""])
        for number, paper in enumerate(papers, start=1):
            title = paper["title"]
            url = links.get(title)
            if not url:
                raise ValueError(f"Missing stable link for: {title}")
            authors, venue = paper["rest"].rsplit(". ", 1)
            output.append(
                f'{number}. **"{title}"**. *{authors}* {venue}. [[Paper]({url})]'
            )
            total += 1
    README.write_text("\n".join(output) + "\n", encoding="utf-8")
    print(f"Updated {README} from {len(sections)} sections and {total} papers")


if __name__ == "__main__":
    main()
