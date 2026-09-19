# -*- coding: utf-8 -*-
"""Generate Monday-style visual evidence workpapers from a JSON config.

Usage:
    python create_visual_workpaper.py --config path/to/workpaper_config.json

The config is intentionally data-heavy: future companies should change the
sentence/source mapping JSON, not this script.
"""

from __future__ import annotations

import argparse
import os
import re
import hashlib
import json
import math
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from PIL import Image, ImageDraw, ImageFont
from playwright.sync_api import sync_playwright



def resolve_path(value: str | Path, base_dir: Path) -> Path:
    path = Path(value)
    if path.is_absolute():
        return path
    return (base_dir / path).resolve()


def has_path_part(path: Path, part: str) -> bool:
    return any(piece.lower() == part.lower() for piece in path.parts)


def validate_output_paths(output_xlsx: Path, process_dir: Path, cfg: dict) -> None:
    if cfg.get("allow_submit_materials_output"):
        return
    blocked = [path for path in (output_xlsx, process_dir) if has_path_part(path, "提交材料")]
    if blocked:
        paths = ", ".join(str(path) for path in blocked)
        raise ValueError(f"Output/process paths must not point into 提交材料 unless allow_submit_materials_output=true: {paths}")


def load_config(path: Path) -> dict:
    with path.open("r", encoding="utf-8-sig") as f:
        cfg = json.load(f)
    cfg["_config_dir"] = str(path.parent.resolve())
    return cfg


def font(size: int) -> ImageFont.ImageFont:
    candidates = [os.environ.get("REPORT_FONT", ""),
        "C:/Windows/Fonts/msyh.ttc", "C:/Windows/Fonts/simsun.ttc",
        "/System/Library/Fonts/PingFang.ttc", "/System/Library/Fonts/STHeiti Medium.ttc",
        "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc"]
    for candidate in candidates:
        if candidate and Path(candidate).is_file():
            return ImageFont.truetype(candidate, size)
    raise RuntimeError("Chinese font missing; set REPORT_FONT to a CJK .ttf/.ttc/.otf file")

BODY_FONT = font(28)
BOLD_FONT = font(30)
TITLE_FONT = font(34)


def text_width(draw: ImageDraw.ImageDraw, text: str, fnt: ImageFont.ImageFont) -> int:
    bbox = draw.textbbox((0, 0), text, font=fnt)
    return bbox[2] - bbox[0]


def text_height(draw: ImageDraw.ImageDraw, text: str, fnt: ImageFont.ImageFont) -> int:
    bbox = draw.textbbox((0, 0), text or "测", font=fnt)
    return bbox[3] - bbox[1]


def is_cjk(ch: str) -> bool:
    return "\u4e00" <= ch <= "\u9fff" or "\u3000" <= ch <= "\u303f" or "\uff00" <= ch <= "\uffef"


def iter_units(text: str) -> list[tuple[str, int, int]]:
    """Return drawable units with original start/end offsets."""
    units: list[tuple[str, int, int]] = []
    i = 0
    while i < len(text):
        ch = text[i]
        if ch == "\n":
            units.append(("\n", i, i + 1))
            i += 1
        elif ch.isspace():
            j = i + 1
            while j < len(text) and text[j].isspace() and text[j] != "\n":
                j += 1
            units.append((" ", i, j))
            i = j
        elif is_cjk(ch):
            units.append((ch, i, i + 1))
            i += 1
        else:
            j = i + 1
            while j < len(text) and (not text[j].isspace()) and (not is_cjk(text[j])):
                j += 1
            units.append((text[i:j], i, j))
            i = j
    return units


def compact_text_with_map(text: str) -> tuple[str, list[int]]:
    compact: list[str] = []
    mapping: list[int] = []
    for idx, ch in enumerate(text):
        if ch.isspace() or ch == "\u00a0":
            continue
        compact.append(ch.lower())
        mapping.append(idx)
    return "".join(compact), mapping


def highlight_ranges(
    text: str,
    highlights: list[str],
    strict: bool = False,
    context: str = "",
) -> list[tuple[int, int]]:
    ranges: list[tuple[int, int]] = []
    compact_body, body_map = compact_text_with_map(text)
    for phrase in highlights:
        if not phrase:
            continue
        compact_phrase, _ = compact_text_with_map(phrase)
        start = compact_body.find(compact_phrase)
        if start >= 0 and compact_phrase:
            end = start + len(compact_phrase) - 1
            ranges.append((body_map[start], body_map[end] + 1))
        elif strict:
            label = f" in {context}" if context else ""
            raise ValueError(f"Highlight not found{label}: {phrase}")
    return ranges


def intersects(start: int, end: int, ranges: list[tuple[int, int]]) -> bool:
    return any(start < r_end and end > r_start for r_start, r_end in ranges)


def wrap_units(
    draw: ImageDraw.ImageDraw,
    text: str,
    fnt: ImageFont.ImageFont,
    max_width: int,
) -> list[list[tuple[str, int, int]]]:
    lines: list[list[tuple[str, int, int]]] = []
    current: list[tuple[str, int, int]] = []
    width = 0
    for unit, start, end in iter_units(text):
        if unit == "\n":
            lines.append(current)
            current = []
            width = 0
            continue
        token_width = text_width(draw, unit, fnt)
        if current and width + token_width > max_width:
            lines.append(current)
            current = []
            width = 0
            if unit == " ":
                continue
        current.append((unit, start, end))
        width += token_width
    if current:
        lines.append(current)
    return lines


def make_text_image(
    sheet: str,
    heading: str,
    body: str,
    out: Path,
    highlights: list[str] | None = None,
    include_sheet_title: bool = True,
    width: int = 900,
) -> dict:
    margin_x = 30
    margin_y = 24
    max_text_width = width - margin_x * 2
    ranges = highlight_ranges(
        body,
        highlights or [],
        strict=bool(highlights),
        context=f"{sheet}/{heading}",
    )

    probe = Image.new("RGB", (width, 200), "white")
    draw = ImageDraw.Draw(probe)
    lines = wrap_units(draw, body, BODY_FONT, max_text_width)
    line_h = text_height(draw, "测", BODY_FONT) + 11
    title_h = text_height(draw, sheet, TITLE_FONT) + 24 if include_sheet_title else 0
    heading_h = text_height(draw, heading, BOLD_FONT) + 12
    height = margin_y * 2 + title_h + heading_h + max(1, len(lines)) * line_h + 12

    img = Image.new("RGB", (width, height), "white")
    draw = ImageDraw.Draw(img)
    y = margin_y
    purple = (104, 73, 167)

    if include_sheet_title:
        draw.text((margin_x, y), sheet, font=TITLE_FONT, fill=purple)
        y += text_height(draw, sheet, TITLE_FONT) + 8
        draw.line((margin_x, y, width - margin_x, y), fill=purple, width=3)
        y += 16

    draw.text((margin_x, y), heading, font=BOLD_FONT, fill=(20, 20, 20))
    y += heading_h

    for line in lines:
        x = margin_x
        for unit, start, end in line:
            unit_width = text_width(draw, unit, BODY_FONT)
            if intersects(start, end, ranges):
                draw.rectangle(
                    (x - 1, y - 2, x + unit_width + 1, y + line_h - 3),
                    fill=(216, 216, 216),
                )
            draw.text((x, y), unit, font=BODY_FONT, fill=(40, 40, 40))
            x += unit_width
        y += line_h

    out.parent.mkdir(parents=True, exist_ok=True)
    img.save(out)
    return {"path": str(out), "width": width, "height": height}


def extract_docx_paragraphs(docx_path: Path) -> list[str]:
    ns = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    with zipfile.ZipFile(docx_path) as zf:
        root = ET.fromstring(zf.read("word/document.xml"))
    paragraphs: list[str] = []
    for par in root.iter(ns + "p"):
        text = "".join(node.text or "" for node in par.iter(ns + "t")).strip()
        if text:
            paragraphs.append(text)
    return paragraphs


def select_body(paragraphs: list[str], section: dict) -> str:
    if section.get("body"):
        if section["body"] not in paragraphs:
            raise ValueError("Section body differs from final Word: " + section.get("sheet", ""))
        return section["body"]
    selector = section.get("body_selector", {})
    prefix = selector.get("prefix")
    contains = selector.get("contains")
    matches = [text for text in paragraphs if (prefix and text.startswith(prefix)) or (contains and contains in text)]
    if len(matches) == 1:
        return matches[0]
    if len(matches) > 1:
        raise ValueError("Body selector is ambiguous; provide the exact Word paragraph")
    raise ValueError(f"Cannot locate body for sheet {section.get('sheet')!r}")


def highlighted_html(source_path: Path, out_path: Path, terms: list[str]) -> Path:
    html = source_path.read_text(encoding="utf-8", errors="ignore")
    soup = BeautifulSoup(html, "html.parser")
    for node in soup(["script", "iframe", "object", "embed"]):
        node.decompose()
    for node in soup.find_all(True):
        for key in list(node.attrs):
            if key.lower().startswith("on"):
                del node.attrs[key]
    if soup.head is None:
        head = soup.new_tag("head")
        if soup.html:
            soup.html.insert(0, head)
        else:
            soup.insert(0, head)
    style = soup.new_tag("style")
    style.string = """
      mark.codex-highlight {
        background: rgba(255, 238, 88, 0.78) !important;
        outline: 4px solid rgba(239, 68, 68, 0.92) !important;
        outline-offset: 2px !important;
        color: inherit !important;
      }
      body { background: #ffffff !important; }
    """
    soup.head.append(style)
    script = soup.new_tag("script")
    script.string = f"""
    (() => {{
      const terms = {json.dumps(terms, ensure_ascii=False)};
      const compactWithMap = (s) => {{
        const compact = [];
        const map = [];
        const raw = String(s || '');
        for (let i = 0; i < raw.length; i++) {{
          const ch = raw[i];
          if (/\\s/.test(ch) || ch === '\\u00a0') continue;
          compact.push(ch.toLowerCase());
          map.push(i);
        }}
        return {{ compact: compact.join(''), map }};
      }};
      let markId = 0;
      const termHits = Object.fromEntries(terms.map((term) => [term, 0]));
      const walker = document.createTreeWalker(document.body || document.documentElement, NodeFilter.SHOW_TEXT);
      const nodes = [];
      while (walker.nextNode()) nodes.push(walker.currentNode);
      const findNext = (raw, from) => {{
        const body = compactWithMap(raw.slice(from));
        if (!body.compact) return null;
        let best = null;
        for (const term of terms) {{
          const wanted = compactWithMap(term).compact;
          if (!wanted) continue;
          const idx = body.compact.indexOf(wanted);
          if (idx < 0) continue;
          const start = from + body.map[idx];
          const end = from + body.map[idx + wanted.length - 1] + 1;
          if (!best || start < best.start || (start === best.start && end > best.end)) {{
            best = {{ term, start, end }};
          }}
        }}
        return best;
      }};
      for (const node of nodes) {{
        if (!node.parentNode) continue;
        const raw = node.nodeValue || '';
        if (!raw.trim()) continue;
        const frag = document.createDocumentFragment();
        let cursor = 0;
        let matched = false;
        while (cursor < raw.length) {{
          const next = findNext(raw, cursor);
          if (!next || next.end <= cursor) break;
          if (next.start > cursor) frag.appendChild(document.createTextNode(raw.slice(cursor, next.start)));
          const mark = document.createElement('mark');
          mark.className = 'codex-highlight';
          mark.id = 'codex-highlight-' + markId++;
          mark.textContent = raw.slice(next.start, next.end);
          frag.appendChild(mark);
          termHits[next.term] += 1;
          cursor = next.end;
          matched = true;
        }}
        if (!matched) continue;
        if (cursor < raw.length) frag.appendChild(document.createTextNode(raw.slice(cursor)));
        node.parentNode.replaceChild(frag, node);
      }}
      window.__codexHighlightCount = markId;
      window.__codexTermHits = termHits;
      window.__codexMissingTerms = terms.filter((term) => !termHits[term]);
    }})();
    """
    if soup.body:
        soup.body.append(script)
    else:
        soup.append(script)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text(str(soup), encoding="utf-8")
    return out_path


def screenshot_html(
    page,
    html_path: Path,
    out_path: Path,
    allow_unhighlighted: bool = False,
    allow_missing_terms: bool = False,
) -> dict:
    page.goto(html_path.as_uri(), wait_until="load")
    page.wait_for_timeout(500)
    count = int(page.evaluate("window.__codexHighlightCount || 0"))
    missing_terms = page.evaluate("window.__codexMissingTerms || []")
    term_hits = page.evaluate("window.__codexTermHits || {}")
    if missing_terms and not (allow_missing_terms or allow_unhighlighted):
        raise ValueError(f"HTML terms not found in {html_path.name}: {missing_terms}")
    if count <= 0:
        if not allow_unhighlighted:
            raise ValueError(f"No HTML highlights were found for {html_path.name}")
        page.screenshot(path=str(out_path), full_page=False)
        return {"highlight_count": count, "missing_terms": missing_terms, "term_hits": term_hits}

    clip = page.evaluate(
        """() => {
          const marks = Array.from(document.querySelectorAll('mark.codex-highlight'));
          const rects = marks.map((m) => m.getBoundingClientRect()).filter((r) => r.width > 0 && r.height > 0);
          if (!rects.length) return null;
          const doc = document.documentElement;
          const body = document.body || doc;
          const pageWidth = Math.max(doc.scrollWidth, body.scrollWidth, doc.clientWidth);
          const pageHeight = Math.max(doc.scrollHeight, body.scrollHeight, doc.clientHeight);
          let left = Math.min(...rects.map((r) => r.left + window.scrollX));
          let right = Math.max(...rects.map((r) => r.right + window.scrollX));
          let top = Math.min(...rects.map((r) => r.top + window.scrollY));
          let bottom = Math.max(...rects.map((r) => r.bottom + window.scrollY));
          const padX = 160, padY = 90;
          left = Math.max(0, left - padX);
          right = Math.min(pageWidth, right + padX);
          top = Math.max(0, top - padY);
          bottom = Math.min(pageHeight, bottom + padY);
          const minWidth = Math.min(1100, pageWidth);
          if (right - left < minWidth) {
            const center = (left + right) / 2;
            left = Math.max(0, center - minWidth / 2);
            right = Math.min(pageWidth, left + minWidth);
            left = Math.max(0, right - minWidth);
          }
          const minHeight = Math.min(240, pageHeight);
          if (bottom - top < minHeight) {
            const center = (top + bottom) / 2;
            top = Math.max(0, center - minHeight / 2);
            bottom = Math.min(pageHeight, top + minHeight);
            top = Math.max(0, bottom - minHeight);
          }
          return {x: Math.round(left), y: Math.round(top), width: Math.round(right-left), height: Math.round(bottom-top)};
        }"""
    )
    if not clip:
        if not allow_unhighlighted:
            raise ValueError(f"Cannot crop HTML screenshot because no highlight clip was found for {html_path.name}")
        page.screenshot(path=str(out_path), full_page=False)
        return {"highlight_count": count, "missing_terms": missing_terms, "term_hits": term_hits}

    full_path = out_path.with_name(f"{out_path.stem}_full.png")
    page.screenshot(path=str(full_path), full_page=True)
    with Image.open(full_path) as img:
        left = max(0, min(int(clip["x"]), img.width - 1))
        top = max(0, min(int(clip["y"]), img.height - 1))
        right = max(left + 1, min(left + int(clip["width"]), img.width))
        bottom = max(top + 1, min(top + int(clip["height"]), img.height))
        img.crop((left, top, right, bottom)).save(out_path)
    full_path.unlink(missing_ok=True)
    return {"highlight_count": count, "missing_terms": missing_terms, "term_hits": term_hits}


def crop_image(path: Path, out_path: Path, clip: list[int] | None = None) -> None:
    with Image.open(path) as img:
        if clip:
            x, y, w, h = [int(v) for v in clip]
            if x < 0 or y < 0 or w <= 0 or h <= 0 or x+w > img.width or y+h > img.height:
                raise ValueError("Crop rectangle must be inside the source image")
            left = max(0, min(x, img.width - 1))
            top = max(0, min(y, img.height - 1))
            right = max(left + 1, min(left + w, img.width))
            bottom = max(top + 1, min(top + h, img.height))
            img.crop((left, top, right, bottom)).save(out_path)
        else:
            img.save(out_path)


def validate_source_meta(evidence_id: str, meta: dict, source_path: Path) -> None:
    suffix = source_path.suffix.lower()
    allow_unhighlighted = bool(meta.get("allow_unhighlighted"))
    pre_highlighted = bool(meta.get("pre_highlighted"))

    if suffix in {".html", ".htm"} and not meta.get("terms") and not allow_unhighlighted:
        raise ValueError(f"{evidence_id}: HTML source requires non-empty terms unless allow_unhighlighted=true")
    if suffix == ".pdf" and not meta.get("clip") and not allow_unhighlighted:
        raise ValueError(f"{evidence_id}: PDF source requires clip coordinates unless allow_unhighlighted=true")
    if suffix in {".png", ".jpg", ".jpeg", ".webp"}:
        if not meta.get("clip") and not (pre_highlighted or allow_unhighlighted):
            raise ValueError(f"{evidence_id}: image source requires clip coordinates or pre_highlighted=true")


def render_pdf_source(source_path: Path, out_path: Path, page_index: int, clip: list[int] | None = None) -> None:
    try:
        import fitz  # type: ignore
    except ImportError as exc:
        raise RuntimeError("PDF sources require PyMuPDF/fitz. Install PyMuPDF or provide a pre-cropped image.") from exc

    with fitz.open(source_path) as doc:
        if page_index < 0 or page_index >= len(doc):
            raise ValueError(f"PDF page_index {page_index} out of range for {source_path} with {len(doc)} pages")
        page = doc[page_index]
        pix = page.get_pixmap(matrix=fitz.Matrix(2, 2), alpha=False)
    tmp = out_path.with_name(f"{out_path.stem}_rendered.png")
    pix.save(tmp)
    crop_image(tmp, out_path, clip)
    tmp.unlink(missing_ok=True)


def clean_source_label(meta: dict, source_path: Path) -> str:
    label = meta.get("label") or source_path.stem
    forbidden = ["网页位置", "：prepared", "（html", "(html"]
    lowered = str(label).lower()
    if any(term in lowered for term in forbidden):
        raise ValueError(f"Invalid source label; keep only the file name without type/location descriptions: {label}")
    return str(label)


def source_note(meta: dict, source_path: Path) -> str:
    label = clean_source_label(meta, source_path)
    page = str(meta.get("page", "")).strip()
    if page and page != "无页码":
        return f"资料来源：{label}，页码：{page}"
    return f"资料来源：{label}"


def make_source_assets(cfg: dict, dirs: dict[str, Path], base_dir: Path) -> dict[str, dict]:
    evidence_cfg: dict = cfg["evidence"]
    manifest: dict[str, dict] = {}
    browser_exe = os.environ.get("REPORT_BROWSER")
    html_items = []

    for evidence_id, meta in evidence_cfg.items():
        source = resolve_path(meta["source"], base_dir)
        validate_source_meta(evidence_id, meta, source)
        out = dirs["source"] / f"{evidence_id}.png"
        suffix = source.suffix.lower()
        if suffix in {".html", ".htm"}:
            html_items.append((evidence_id, meta, source, out))
        elif suffix == ".pdf":
            page_index = int(meta.get("page_index", 0))
            render_pdf_source(source, out, page_index, meta.get("clip"))
            manifest[evidence_id] = {
                "image": str(out),
                "note": source_note(meta, source),
                "calc": meta.get("calc", ""),
                "highlight_count": 1 if meta.get("clip") else 0,
                "allow_unhighlighted": bool(meta.get("allow_unhighlighted")),
            }
        elif suffix in {".png", ".jpg", ".jpeg", ".webp"}:
            crop_image(source, out, meta.get("clip"))
            manifest[evidence_id] = {
                "image": str(out),
                "note": source_note(meta, source),
                "calc": meta.get("calc", ""),
                "highlight_count": 1 if meta.get("clip") or meta.get("pre_highlighted") else 0,
                "allow_unhighlighted": bool(meta.get("allow_unhighlighted")),
            }
        else:
            raise ValueError(f"Unsupported source type for {source}")

    if html_items:
        viewport = cfg.get("source_viewport", {"width": 1400, "height": 1000})
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True, **({"executable_path": browser_exe} if browser_exe else {}))
            page = browser.new_page(
                viewport={"width": int(viewport.get("width", 1400)), "height": int(viewport.get("height", 1000))},
                device_scale_factor=1,
            )
            page.route("http://**/*", lambda route: route.abort())
            page.route("https://**/*", lambda route: route.abort())
            for evidence_id, meta, source, out in html_items:
                html = highlighted_html(source, dirs["html"] / f"{evidence_id}.html", meta.get("terms", []))
                shot = screenshot_html(
                    page,
                    html,
                    out,
                    allow_unhighlighted=bool(meta.get("allow_unhighlighted")),
                    allow_missing_terms=bool(meta.get("allow_missing_terms")),
                )
                manifest[evidence_id] = {
                    "image": str(out),
                    "note": source_note(meta, source),
                    "calc": meta.get("calc", ""),
                    "highlight_count": shot["highlight_count"],
                    "missing_terms": shot["missing_terms"],
                    "term_hits": shot["term_hits"],
                    "allow_unhighlighted": bool(meta.get("allow_unhighlighted")),
                }
            browser.close()

    for evidence_id, item in manifest.items():
        with Image.open(item["image"]) as img:
            item["width"] = img.width
            item["height"] = img.height
    return manifest


def validate_config(cfg: dict, base_dir: Path, output_xlsx: Path, process_dir: Path) -> None:
    validate_output_paths(output_xlsx, process_dir, cfg)
    if output_xlsx.suffix.lower() != ".xlsx":
        raise ValueError("output_xlsx must end in .xlsx")
    if output_xlsx.exists():
        raise ValueError("Output already exists; choose a new version")
    if not cfg.get("word_confirmed"):
        raise ValueError("User must confirm final Word before workpaper generation")
    word_path = resolve_path(cfg["word_path"], base_dir)
    if not word_path.exists():
        raise FileNotFoundError(f"Word file not found: {word_path}")

    if hashlib.sha256(word_path.read_bytes()).hexdigest() != cfg.get("word_sha256"):
        raise ValueError("Word hash changed or missing; obtain confirmation and refresh mapping")
    paragraphs = extract_docx_paragraphs(word_path)
    sections = cfg.get("sections", [])
    if not sections:
        raise ValueError("Config must include sections")
    evidence = cfg.get("evidence", {})
    if not evidence:
        raise ValueError("Config must include evidence")

    evidence_ids = set(evidence)
    if len({x.get("sheet") for x in sections}) != len(sections):
        raise ValueError("Duplicate sheet names")
    for name in evidence:
        if not re.fullmatch(r"[A-Za-z0-9_-]+", name):
            raise ValueError("Evidence IDs must use ASCII letters, digits, _ or -")
    for section in sections:
        sheet = section.get("sheet", "<missing sheet>")
        if not isinstance(sheet, str) or not sheet or len(sheet)>31 or re.search(r"[\\/*?:\[\]]",sheet):
            raise ValueError("Invalid Excel sheet name")
        body = select_body(paragraphs, section)
        overview = section.get("overview_evidence")
        if overview not in evidence_ids:
            raise ValueError(f"{sheet}: overview_evidence not found in evidence: {overview}")
        points = section.get("points") or []
        if not points:
            raise ValueError(f"{sheet}: must include at least one point")
        for idx, point in enumerate(points, start=1):
            if not point.get("evidence_ids"):
                raise ValueError("Every point requires evidence_ids")
            selected = point.get("highlights", []) if point.get("left_mode") == "full_body_highlight" else [point.get("text", "")]
            if not selected or any(not text or text not in body for text in selected):
                raise ValueError("Point text is not in the confirmed Word paragraph")
            mode = point.get("left_mode", "point_text")
            if mode == "full_body_highlight":
                if not (section.get("body") or section.get("body_selector")):
                    raise ValueError(
                        f"{sheet} point {idx}: left_mode=full_body_highlight requires section body or body_selector"
                    )
                if not point.get("highlights"):
                    raise ValueError(f"{sheet} point {idx}: left_mode=full_body_highlight requires highlights")
            else:
                point_body_for_image(section, point)
            for ev_id in point.get("evidence_ids", []):
                if ev_id not in evidence_ids:
                    raise ValueError(f"{sheet} point {idx}: evidence_id not found: {ev_id}")

    for evidence_id, meta in evidence.items():
        source = resolve_path(meta["source"], base_dir)
        if not source.exists():
            raise FileNotFoundError(f"{evidence_id}: source file not found: {source}")
        validate_source_meta(evidence_id, meta, source)
        source_note(meta, source)


def scaled_dims(path: Path, max_width: int, max_height: int | None = None) -> tuple[int, int]:
    with Image.open(path) as img:
        w, h = img.size
    scale = max_width / w
    if max_height is not None:
        scale = min(scale, max_height / h)
    return max(1, int(w * scale)), max(1, int(h * scale))


def add_image(ws, image_path: Path, cell: str, max_width: int, max_height: int | None = None) -> tuple[int, int]:
    width, height = scaled_dims(image_path, max_width, max_height)
    img = XLImage(str(image_path))
    img.width = width
    img.height = height
    ws.add_image(img, cell)
    return width, height


def rows_for_px(px: int, note_lines: int = 0) -> int:
    return max(8, math.ceil((px + 42 + note_lines * 24) / 24))


def style_note_cell(cell, fill: str = "FFFFFF") -> None:
    cell.alignment = Alignment(wrap_text=True, vertical="top")
    cell.font = Font(name="Microsoft YaHei", size=10, color="333333")
    cell.fill = PatternFill("solid", fgColor=fill)
    side = Side(style="thin", color="D9D9D9")
    cell.border = Border(left=side, right=side, top=side, bottom=side)


def add_separator(ws, row: int) -> None:
    ws.row_dimensions[row].height = 12
    fill = PatternFill("solid", fgColor="FFF2CC")
    side = Side(style="thin", color="F1C232")
    for col in ("A", "B", "C", "D"):
        cell = ws[f"{col}{row}"]
        cell.fill = fill
        cell.border = Border(left=side, right=side, top=side, bottom=side)


def configure_sheet(ws) -> None:
    ws.sheet_view.showGridLines = False
    for col, width in {"A": 94, "B": 4, "C": 128, "D": 28}.items():
        ws.column_dimensions[col].width = width
    ws.sheet_format.defaultRowHeight = 18
    ws.freeze_panes = "A1"
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.page_margins.left = 0.2
    ws.page_margins.right = 0.2
    ws.page_margins.top = 0.25
    ws.page_margins.bottom = 0.25


def point_body_for_image(section: dict, point: dict) -> str:
    mode = point.get("left_mode", "point_text")
    if mode == "full_body_highlight":
        return section["body"]
    if mode != "point_text":
        raise ValueError(f"Unsupported point left_mode: {mode}")
    if point.get("text"):
        return point["text"]
    highlights = point.get("highlights") or []
    if len(highlights) == 1:
        return highlights[0]
    raise ValueError(
        f"Point in sheet {section.get('sheet')} must provide text, "
        "or exactly one highlight, unless left_mode=full_body_highlight"
    )


def make_text_assets(sections: list[dict], dirs: dict[str, Path]) -> tuple[dict, dict]:
    para_manifest: dict[str, dict] = {}
    point_manifest: dict[str, list[dict]] = {}
    for section in sections:
        sheet = section["sheet"]
        para_manifest[sheet] = make_text_image(
            sheet,
            section.get("heading", sheet),
            section["body"],
            dirs["paragraph"] / f"{sheet}.png",
            include_sheet_title=True,
            width=900,
        )
        point_manifest[sheet] = []
        for idx, point in enumerate(section["points"], start=1):
            point_body = point_body_for_image(section, point)
            point_highlights = point.get("highlights", [])
            if point.get("left_mode", "point_text") == "point_text" and not point_highlights:
                point_highlights = [point_body]
            point_manifest[sheet].append(
                make_text_image(
                    sheet,
                    section.get("heading", sheet),
                    point_body,
                    dirs["point"] / f"{sheet}_{idx:02d}.png",
                    highlights=point_highlights,
                    include_sheet_title=(idx == 1),
                    width=900,
                )
            )
    return para_manifest, point_manifest


def create_workbook(
    output_xlsx: Path,
    sections: list[dict],
    para_manifest: dict,
    point_manifest: dict,
    source_manifest: dict,
) -> None:
    wb = Workbook()
    wb.remove(wb.active)

    for section in sections:
        ws = wb.create_sheet(section["sheet"])
        configure_sheet(ws)

        para_info = para_manifest[section["sheet"]]
        first_evidence = source_manifest[section["overview_evidence"]]
        _, para_h = add_image(ws, Path(para_info["path"]), "A1", max_width=660)
        note_rows = 1 + (1 if first_evidence.get("calc") else 0)
        ws["C1"] = first_evidence["note"]
        style_note_cell(ws["C1"])
        image_row = 2
        if first_evidence.get("calc"):
            ws["C2"] = first_evidence["calc"]
            style_note_cell(ws["C2"], fill="FFF2CC")
            image_row = 3
        _, ev_h = add_image(ws, Path(first_evidence["image"]), f"C{image_row}", max_width=1180, max_height=520)
        start_note_row = max(rows_for_px(para_h), rows_for_px(ev_h) + note_rows)
        current_row = start_note_row + 4

        for idx, point in enumerate(section["points"]):
            point_info = point_manifest[section["sheet"]][idx]
            _, left_h = add_image(ws, Path(point_info["path"]), f"A{current_row}", max_width=660)
            left_rows = rows_for_px(left_h)

            right_row = current_row
            for ev_id in point["evidence_ids"]:
                ev = source_manifest[ev_id]
                note_rows = 1
                ws[f"C{right_row}"] = ev["note"]
                style_note_cell(ws[f"C{right_row}"])
                if ev.get("calc"):
                    ws[f"C{right_row + 1}"] = ev["calc"]
                    style_note_cell(ws[f"C{right_row + 1}"], fill="FFF2CC")
                    note_rows += 1
                image_row = right_row + note_rows + 1
                _, height = add_image(ws, Path(ev["image"]), f"C{image_row}", max_width=1180, max_height=520)
                needed = note_rows + 1 + rows_for_px(height)
                right_row += needed + 2

            group_rows = max(left_rows, right_row - current_row)
            if idx < len(section["points"]) - 1:
                separator_row = current_row + group_rows
                add_separator(ws, separator_row)
                current_row = separator_row + 3
            else:
                current_row += group_rows + 1

        ws.print_area = f"A1:D{max(current_row, 40)}"

    output_xlsx.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_xlsx)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--config", required=True, help="Path to workpaper JSON config.")
    parser.add_argument("--validate-only", action="store_true", help="Validate config paths/references without generating Excel.")
    args = parser.parse_args()

    config_path = Path(args.config).resolve()
    cfg = load_config(config_path)
    base_dir = resolve_path(cfg.get("company_dir", "."), Path(cfg["_config_dir"]))
    output_xlsx = resolve_path(cfg["output_xlsx"], base_dir)
    process_dir = resolve_path(cfg.get("process_dir", "写作产出/_底稿过程文件"), base_dir)
    validate_config(cfg, base_dir, output_xlsx, process_dir)
    if args.validate_only:
        print(json.dumps({"validated": str(config_path)}, ensure_ascii=False, indent=2))
        return
    dirs = {
        "process": process_dir,
        "paragraph": process_dir / "paragraph_images",
        "point": process_dir / "point_images",
        "source": process_dir / "source_screenshots",
        "html": process_dir / "highlighted_html",
    }
    for path in dirs.values():
        path.mkdir(parents=True, exist_ok=True)

    word_path = resolve_path(cfg["word_path"], base_dir)
    paragraphs = extract_docx_paragraphs(word_path)
    sections = cfg["sections"]
    for section in sections:
        section["body"] = select_body(paragraphs, section)

    para_manifest, point_manifest = make_text_assets(sections, dirs)
    source_manifest = make_source_assets(cfg, dirs, base_dir)
    create_workbook(output_xlsx, sections, para_manifest, point_manifest, source_manifest)

    manifest = {
        "word": str(word_path),
        "output": str(output_xlsx),
        "sections": sections,
        "source_screenshots": source_manifest,
        "paragraph_images": para_manifest,
        "point_images": point_manifest,
    }
    (process_dir / "visual_workpaper_manifest.json").write_text(
        json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print(json.dumps({"output": str(output_xlsx)}, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
