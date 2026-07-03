from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
SOURCE = Path.home() / ".codex" / "attachments" / "1c842d35-f575-4b53-a784-5e335dc2a543" / "pasted-text.txt"
OUTPUT = ROOT / "frontend" / "public" / "legacy-vantis.html"


def repair_mojibake(text: str) -> str:
    try:
        return text.encode("latin1").decode("utf-8")
    except UnicodeEncodeError:
        return text.encode("latin1", errors="ignore").decode("utf-8", errors="ignore")


def main() -> None:
    html = SOURCE.read_text(encoding="utf-8", errors="replace")
    html = repair_mojibake(html)
    marker = "<script>\n// ── ESTADO ──"
    if "/vantis-mock.js" not in html:
        html = html.replace(marker, '<script src="/vantis-mock.js"></script>\n' + marker, 1)
    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    OUTPUT.write_text(html, encoding="utf-8")
    print(f"generated {OUTPUT} ({OUTPUT.stat().st_size} bytes)")


if __name__ == "__main__":
    main()
