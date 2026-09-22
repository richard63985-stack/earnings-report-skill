# -*- coding: utf-8 -*-
"""Fill the fixed research-report Word template from a markdown draft.

This script intentionally patches OOXML text in-place instead of serializing the
whole XML tree. The company template is sensitive to namespace rewrites.
"""

from __future__ import annotations

import argparse
import os
import html
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path


TEXT_RE = re.compile(r"(<w:t\b[^>]*>)(.*?)(</w:t>)", re.S)
P_RE = re.compile(r"<w:p\b[^>]*>.*?</w:p>", re.S)
SDT_START_RE = re.compile(r"<w:sdt(?=[\s>])[^>]*>", re.S)
TAG_RE = re.compile(r"<w:sdt(?=[\s>])[^>]*>|</w:sdt>", re.S)


def read_text(path: Path) -> str:
    return path.read_text(encoding="utf-8-sig")


def section(md: str, name: str) -> str:
    pat = re.compile(rf"^##\s+{re.escape(name)}\s*$([\s\S]*?)(?=^##\s+|\Z)", re.M)
    match = pat.search(md)
    return match.group(1).strip() if match else ""


def clean_blocks(text: str) -> list[str]:
    return [b.strip() for b in re.split(r"\n\s*\n", text.strip()) if b.strip()]


def parse_markdown(md_path: Path) -> dict[str, str | list[str]]:
    md = read_text(md_path)
    title_match = re.search(r"^#\s+(.+?)\s*$", md, re.M)
    if not title_match:
        raise ValueError("Markdown must contain one top-level '# 标题' line.")

    event = " ".join(clean_blocks(section(md, "事件")))
    risk = " ".join(clean_blocks(section(md, "风险提示")))
    advice_blocks = [
        block for block in clean_blocks(section(md, "投资建议")) if not block.startswith("注：")
    ]
    advice = " ".join(advice_blocks)

    summary_parts: list[str] = []
    current_heading: str | None = None
    current_body: list[str] = []
    for block in clean_blocks(section(md, "投资要点")):
        if block.startswith("▌"):
            if current_heading is not None:
                summary_parts.append(current_heading)
                summary_parts.append(" ".join(current_body).strip())
            current_heading = block.strip()
            current_body = []
        else:
            current_body.append(block.strip())

    if current_heading is not None:
        summary_parts.append(current_heading)
        summary_parts.append(" ".join(current_body).strip())

    investment_part_count = len(summary_parts)
    if advice:
        summary_parts.extend(["▌投资建议", advice])

    missing = []
    if not event:
        missing.append("## 事件")
    if investment_part_count != 6 or any(not part for part in summary_parts):
        missing.append("## 投资要点 (three nonempty headings and bodies)")
    if not risk:
        missing.append("## 风险提示")
    if missing:
        raise ValueError("Markdown is missing required content: " + ", ".join(missing))

    return {
        "title": title_match.group(1).strip(),
        "event": event,
        "summary_parts": summary_parts,
        "risk": risk,
    }


def xesc(text: str) -> str:
    return html.escape(text, quote=False)


def sdt_bounds_for_alias(xml: str, alias: str, occurrence: int = 0) -> tuple[int, int]:
    alias_pat = re.compile(
        r"<w:alias\b(?=[^>]*\bw:val=\"" + re.escape(alias) + r"\")[^>]*/?>"
    )
    matches = list(alias_pat.finditer(xml))
    if occurrence >= len(matches):
        raise ValueError(f"Content control alias not found: {alias} occurrence={occurrence}")

    alias_pos = matches[occurrence].start()
    starts = list(SDT_START_RE.finditer(xml, 0, alias_pos))
    if not starts:
        raise ValueError(f"Content control start not found for alias: {alias}")
    start = starts[-1].start()

    depth = 0
    for match in TAG_RE.finditer(xml, start):
        token = match.group(0)
        if token.startswith("</"):
            depth -= 1
            if depth == 0:
                return start, match.end()
        else:
            depth += 1

    raise ValueError(f"Content control end not found for alias: {alias}")


def replace_first_text(fragment: str, text: str) -> str:
    matches = list(TEXT_RE.finditer(fragment))
    if not matches:
        raise ValueError("No <w:t> node found in target fragment.")

    out: list[str] = []
    last = 0
    for idx, match in enumerate(matches):
        out.append(fragment[last : match.start()])
        out.append(match.group(1))
        out.append(xesc(text if idx == 0 else ""))
        out.append(match.group(3))
        last = match.end()
    out.append(fragment[last:])
    return "".join(out)


def replace_sdt_single(xml: str, alias: str, text: str, occurrence: int = 0) -> str:
    start, end = sdt_bounds_for_alias(xml, alias, occurrence)
    block = xml[start:end]
    return xml[:start] + replace_first_text(block, text) + xml[end:]


def replace_sdt_paragraphs(xml: str, alias: str, texts: list[str], occurrence: int = 0) -> str:
    start, end = sdt_bounds_for_alias(xml, alias, occurrence)
    block = xml[start:end]
    paragraphs = list(P_RE.finditer(block))
    if not paragraphs:
        raise ValueError(f"No paragraphs found in content control: {alias}")
    if len(texts) > len(paragraphs):
        raise ValueError(
            f"Too many summary paragraphs for template: {len(texts)} > {len(paragraphs)}"
        )

    out: list[str] = []
    last = 0
    for idx, paragraph in enumerate(paragraphs):
        out.append(block[last : paragraph.start()])
        text = texts[idx] if idx < len(texts) else ""
        out.append(replace_first_text(paragraph.group(0), text))
        last = paragraph.end()
    out.append(block[last:])
    return xml[:start] + "".join(out) + xml[end:]


def apply_layout_fix(xml: str) -> tuple[str, bool]:
    target = '<w:trHeight w:val="7680"/>'
    replacement = '<w:trHeight w:val="567"/>'
    if target in xml:
        return xml.replace(target, replacement, 1), True
    return xml, False


def write_docx(template: Path, output: Path, document_xml: str) -> None:
    output.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(template, "r") as zin, zipfile.ZipFile(output, "w") as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "word/document.xml":
                data = document_xml.encode("utf-8")
            zout.writestr(item, data)


def validate_docx(template: Path, output: Path, expected: dict[str, str | list[str]]) -> list[str]:
    issues: list[str] = []
    with zipfile.ZipFile(template, "r") as zt:
        template_xml = zt.read("word/document.xml").decode("utf-8")
    with zipfile.ZipFile(output, "r") as zo:
        bad = zo.testzip()
        output_xml = zo.read("word/document.xml").decode("utf-8")

    if bad is not None:
        issues.append(f"zip test failed at {bad}")
    if output_xml.count("<w:sdt") != template_xml.count("<w:sdt"):
        issues.append("content control count changed")
    if "<ns0:" in output_xml:
        issues.append("namespace rewrite detected: ns0")
    if "mc:Ignorable" not in output_xml:
        issues.append("missing mc:Ignorable")

    required_texts: list[str] = [
        str(expected["title"]),
        str(expected["event"]),
        *[str(x) for x in expected["summary_parts"]],
        str(expected["risk"]),
    ]
    for text in required_texts:
        if text and xesc(text) not in output_xml and text not in output_xml:
            issues.append("missing expected text: " + text[:80])

    for placeholder in ["报告标题报告标题", "文字待填充"]:
        if placeholder in output_xml:
            issues.append("template placeholder remains: " + placeholder)

    return issues


def word_page_count(path: Path) -> int | None:
    if sys.platform != "win32" or not shutil.which("pwsh"):
        return None
    ps = f"""
$doc = {ps_quote(str(path.resolve()))}
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {{
  $d = $word.Documents.Open($doc, $false, $true)
  $d.Repaginate()
  $pages = $d.ComputeStatistics(2)
  $d.Close($false)
  Write-Output $pages
}} finally {{
  $word.Quit()
}}
"""
    completed = subprocess.run(
        ["pwsh", "-NoProfile", "-Command", ps],
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        timeout=60,
    )
    if completed.returncode != 0:
        return None
    try:
        return int(completed.stdout.strip().splitlines()[-1])
    except Exception:
        return None


def export_pdf_with_word(path: Path, pdf_path: Path) -> bool:
    if sys.platform != "win32" or not shutil.which("pwsh"):
        return False
    pdf_path.parent.mkdir(parents=True, exist_ok=True)
    ps = f"""
$doc = {ps_quote(str(path.resolve()))}
$pdf = {ps_quote(str(pdf_path.resolve()))}
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {{
  $d = $word.Documents.Open($doc, $false, $true)
  $d.Repaginate()
  $d.ExportAsFixedFormat($pdf, 17)
  $d.Close($false)
  Write-Output "OK"
}} finally {{
  $word.Quit()
}}
"""
    completed = subprocess.run(
        ["pwsh", "-NoProfile", "-Command", ps],
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        timeout=120,
    )
    return completed.returncode == 0 and pdf_path.exists() and pdf_path.stat().st_size > 0


def ps_quote(text: str) -> str:
    return "'" + text.replace("'", "''") + "'"

def find_soffice(preferred: str = "") -> str | None:
    candidates = [preferred, os.environ.get("REPORT_SOFFICE", ""),
        shutil.which("soffice"), shutil.which("libreoffice"),
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
        str(Path(os.environ.get("PROGRAMFILES", "C:/Program Files")) / "LibreOffice/program/soffice.exe")]
    return next((str(Path(x).resolve()) for x in candidates if x and Path(x).is_file()), None)


def export_pdf_with_libreoffice(
    path: Path, pdf_path: Path, soffice_path: str = ""
) -> tuple[bool, str]:
    soffice = find_soffice(soffice_path)
    if not soffice:
        return False, "soffice not found"

    pdf_path.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix="lo_convert_", dir=str(pdf_path.parent)) as tmp:
        cmd = [
            soffice,
            "-env:UserInstallation=" + (Path(tmp)/"profile").resolve().as_uri(),
            "--headless",
            "--invisible",
            "--norestore",
            "--convert-to",
            "pdf",
            "--outdir",
            tmp,
            str(path.resolve()),
        ]
        completed = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=180,
        )
        produced = Path(tmp) / f"{path.stem}.pdf"
        if completed.returncode == 0 and produced.exists() and produced.stat().st_size > 0:
            shutil.copy2(produced, pdf_path)
            return True, (completed.stderr or completed.stdout or "").strip()
        message = "\n".join(
            [
                f"exit={completed.returncode}",
                "stdout=" + completed.stdout.strip(),
                "stderr=" + completed.stderr.strip(),
            ]
        )
        return False, message.strip()


def render_pdf_pages(pdf_path: Path, out_dir: Path) -> None:
    try:
        import pypdfium2 as pdfium
    except Exception as exc:
        raise RuntimeError("pypdfium2 is required for --render-pages") from exc

    out_dir.mkdir(parents=True, exist_ok=True)
    doc = pdfium.PdfDocument(str(pdf_path))
    try:
        for idx in range(len(doc)):
            page = doc[idx]
            try:
                image = page.render(scale=2.0).to_pil()
                image.save(out_dir / f"page-{idx + 1}.png")
            finally:
                page.close()
    finally:
        doc.close()


def fill(args: argparse.Namespace) -> int:
    template = Path(args.template)
    markdown = Path(args.markdown)
    output = Path(args.output)

    if output.exists() or output.resolve() == template.resolve():
        raise ValueError("Choose a new output path; never overwrite template or existing report")
    parsed = parse_markdown(markdown)
    with zipfile.ZipFile(template, "r") as zin:
        xml = zin.read("word/document.xml").decode("utf-8")

    xml = replace_sdt_single(xml, "报告日期", args.date)
    xml = replace_sdt_single(xml, "标题", str(parsed["title"]))
    original_xml = xml
    for alias, value in [("副标题", args.subtitle), ("投资评级", args.rating), ("评级变动", args.rating_change)]:
        if value is not None:
            xml = replace_sdt_single(xml, alias, value)
    xml = replace_sdt_single(xml, "事件公告", str(parsed["event"]))
    xml = replace_sdt_paragraphs(xml, "摘要", [str(x) for x in parsed["summary_parts"]])
    xml = replace_sdt_single(xml, "风险提示", str(parsed["risk"]))
    if args.report_no:
        xml = replace_sdt_single(xml, "报告编号", args.report_no, occurrence=0)
        xml = replace_sdt_single(xml, "报告编号", args.report_no, occurrence=1)

    layout_fixed = False
    if args.fix_legacy_row_height:
        xml, layout_fixed = apply_layout_fix(xml)
    protected = ["作者显示", "行业相对表现", "行业相对市场走势", "行业相关研究报告", "作者所属分组", "自我介绍"]
    protected += [alias for alias, value in [("副标题", args.subtitle), ("投资评级", args.rating), ("评级变动", args.rating_change), ("报告编号", args.report_no or None)] if value is None]
    for alias in protected:
        occurrences = len(re.findall(r'<w:alias\b[^>]*w:val="' + re.escape(alias) + '"', original_xml))
        for occurrence in range(occurrences):
            a, b = sdt_bounds_for_alias(original_xml, alias, occurrence)
            c, d = sdt_bounds_for_alias(xml, alias, occurrence)
            if original_xml[a:b] != xml[c:d]:
                raise ValueError("Protected template field changed: " + alias)
    write_docx(template, output, xml)
    with zipfile.ZipFile(template) as before, zipfile.ZipFile(output) as after:
        if before.namelist() != after.namelist() or any(before.read(n) != after.read(n) for n in before.namelist() if n != "word/document.xml"):
            raise ValueError("Unexpected change outside word/document.xml")

    issues = validate_docx(template, output, parsed)
    if issues:
        for issue in issues:
            print("VALIDATION_ISSUE:", issue, file=sys.stderr)
        return 2

    print(f"created={output}")
    print(f"layout_fixed={layout_fixed}")

    if args.word_check:
        pages = word_page_count(output)
        print(f"word_pages={pages if pages is not None else 'unavailable'}")

    if args.export_pdf:
        pdf_path = Path(args.export_pdf)
        ok = False
        engine_used = ""
        engine_message = ""

        if args.pdf_engine in {"word", "auto"}:
            ok = export_pdf_with_word(output, pdf_path)
            engine_used = "word" if ok else ""

        if not ok and args.pdf_engine in {"libreoffice", "auto"}:
            ok, engine_message = export_pdf_with_libreoffice(
                output, pdf_path, soffice_path=args.soffice
            )
            engine_used = "libreoffice" if ok else engine_used

        print(f"pdf_engine={engine_used if ok else 'failed'}")
        if engine_message:
            print("pdf_engine_message=" + engine_message.replace("\n", " | "))
        print(f"pdf={pdf_path if ok else 'failed'}")
        if not ok:
            return 3
        if ok and args.render_pages:
            render_pdf_pages(pdf_path, Path(args.render_pages))
            print(f"render_pages={args.render_pages}")

    return 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Fill 报告模板.docx from a reviewed markdown earnings-commentary draft."
    )
    parser.add_argument("--markdown", required=True, help="Reviewed markdown draft path.")
    parser.add_argument("--template", default="报告模板.docx", help="Word template path.")
    parser.add_argument("--output", required=True, help="Output docx path.")
    parser.add_argument("--date", required=True, help="Report date, e.g. 2026年06月23日.")
    parser.add_argument("--subtitle", default=None, help="Explicitly authorized subtitle; omit to preserve template.")
    parser.add_argument("--rating", default=None, help="User-approved rating; never infer a default recommendation.")
    parser.add_argument("--rating-change", default=None)
    parser.add_argument("--report-no", default="", help="Optional report number.")
    parser.add_argument("--fix-legacy-row-height", action="store_true", help="Apply the known 7680-row-height repair only when required for this template.")
    parser.add_argument("--word-check", action="store_true", help="Open with Word COM and print pages.")
    parser.add_argument("--export-pdf", default="", help="Optional PDF output path.")
    parser.add_argument(
        "--pdf-engine",
        choices=["word", "libreoffice", "auto"],
        default="auto",
        help="PDF export engine. Use 'auto' to try Word first, then LibreOffice.",
    )
    parser.add_argument("--soffice", default="", help="Optional explicit soffice executable path.")
    parser.add_argument("--render-pages", default="", help="Optional output directory for PDF page PNGs.")
    return parser


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()
    if args.render_pages and not args.export_pdf:
        parser.error("--render-pages requires --export-pdf")
    if (args.word_check or args.pdf_engine == "word") and (sys.platform != "win32" or not shutil.which("pwsh")):
        print("WARNING: Word COM unavailable in this environment.", file=sys.stderr)
    return fill(args)


if __name__ == "__main__":
    for stream in (sys.stdout, sys.stderr):
        stream.reconfigure(encoding="utf-8")
    raise SystemExit(main())
