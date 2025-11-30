# audit_handouts.py
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple, TypedDict, Union
import io
import re
import json
import hashlib
import time

import requests
from docxtpl import DocxTemplate
from docx import Document  # python-docx


class HandoutChunk(TypedDict):
    file: str
    unit: str
    heading: str
    text: str
    captions: List[str]


_CAPTION_RE = re.compile(
    r"(?im)^\s*(figure|fig\.|table|eq\.|equation)\s*\d+"
    r"(\s*[:.\-–—]\s*.*)?$"
)

def _sha256(b: bytes) -> str:
    return hashlib.sha256(b).hexdigest()

def _first_nonempty_line(text: str) -> str:
    for line in (text or "").splitlines():
        s = line.strip()
        if s:
            return s[:200]
    return ""

def _extract_captions(text: str) -> List[str]:
    if not text:
        return []
    caps = []
    for m in _CAPTION_RE.finditer(text):
        line = (m.group(0) or "").strip()
        if line and line not in caps:
            caps.append(line[:250])
    return caps

def _clean_text(s: str) -> str:
    if not s:
        return ""
    # normalize whitespace without destroying structure too much
    s = s.replace("\u00ad", "")  # soft hyphen
    s = re.sub(r"[ \t]+", " ", s)
    s = re.sub(r"\n{3,}", "\n\n", s)
    return s.strip()

def parse_pdf(path: Union[str, Path]) -> List[HandoutChunk]:
    path = Path(path)
    # Prefer pymupdf (fitz), else pypdf.
    try:
        import fitz  # pymupdf
        doc = fitz.open(str(path))
        out: List[HandoutChunk] = []
        for i in range(doc.page_count):
            page = doc.load_page(i)
            text = page.get_text("text") or ""
            text = _clean_text(text)
            out.append({
                "file": path.name,
                "unit": f"page {i+1}",
                "heading": _first_nonempty_line(text),
                "text": text,
                "captions": _extract_captions(text),
            })
        return out
    except Exception:
        pass

    try:
        from pypdf import PdfReader
        reader = PdfReader(str(path))
        out: List[HandoutChunk] = []
        for i, page in enumerate(reader.pages):
            text = page.extract_text() or ""
            text = _clean_text(text)
            out.append({
                "file": path.name,
                "unit": f"page {i+1}",
                "heading": _first_nonempty_line(text),
                "text": text,
                "captions": _extract_captions(text),
            })
        return out
    except Exception as e:
        raise RuntimeError(f"PDF parsing failed (install pymupdf or pypdf). Root error: {e}") from e

def parse_pptx(path: Union[str, Path]) -> List[HandoutChunk]:
    path = Path(path)
    try:
        from pptx import Presentation
    except Exception as e:
        raise RuntimeError("python-pptx is required for PPTX parsing. Please `pip install python-pptx`.") from e

    prs = Presentation(str(path))
    out: List[HandoutChunk] = []

    for idx, slide in enumerate(prs.slides, start=1):
        title = ""
        try:
            if slide.shapes.title and slide.shapes.title.text:
                title = slide.shapes.title.text.strip()
        except Exception:
            title = ""

        texts: List[str] = []
        for shape in slide.shapes:
            try:
                if getattr(shape, "has_text_frame", False) and shape.has_text_frame:
                    t = shape.text_frame.text
                    if t and t.strip():
                        texts.append(t.strip())
            except Exception:
                continue

        notes = ""
        try:
            if slide.has_notes_slide and slide.notes_slide and slide.notes_slide.notes_text_frame:
                notes = (slide.notes_slide.notes_text_frame.text or "").strip()
        except Exception:
            notes = ""

        full = "\n".join(texts)
        if notes:
            full = (full + "\n\n[NOTES]\n" + notes).strip()

        full = _clean_text(full)
        heading = title.strip() if title.strip() else _first_nonempty_line(full)

        out.append({
            "file": path.name,
            "unit": f"slide {idx}",
            "heading": heading,
            "text": full,
            "captions": _extract_captions(full),
        })

    return out

def build_corpus(paths: List[Union[str, Path]]) -> List[HandoutChunk]:
    """Parse + normalize into a single ordered corpus."""
    corpus: List[HandoutChunk] = []
    for p in paths:
        p = Path(p)
        ext = p.suffix.lower()
        if ext == ".pdf":
            corpus.extend(parse_pdf(p))
        elif ext == ".pptx":
            corpus.extend(parse_pptx(p))
        elif ext == ".ppt":
            # .ppt is not supported by python-pptx
            corpus.append({
                "file": p.name,
                "unit": "file",
                "heading": "unsupported .ppt",
                "text": "Unsupported format: .ppt (please export/save as .pptx).",
                "captions": [],
            })
        else:
            corpus.append({
                "file": p.name,
                "unit": "file",
                "heading": "unsupported",
                "text": f"Unsupported format: {ext}",
                "captions": [],
            })
    return corpus

def _compact_chunk(c: HandoutChunk, max_chars: int = 2600) -> HandoutChunk:
    t = c.get("text", "") or ""
    if len(t) <= max_chars:
        return c
    # keep start + end, preserve some structure
    head = t[: int(max_chars * 0.75)]
    tail = t[- int(max_chars * 0.2):]
    t2 = (head + "\n...\n" + tail).strip()
    out: HandoutChunk = {
        "file": c["file"],
        "unit": c["unit"],
        "heading": c.get("heading", "") or "",
        "text": t2,
        "captions": (c.get("captions") or [])[:15],
    }
    return out

def _flatten_weekly(draft_bundle: Dict[str, Any]) -> str:
    w = (draft_bundle or {}).get("weekly", {}) or {}
    theory = w.get("theory_rows", []) or []
    practical = w.get("practical_rows", []) or []
    def fmt(kind: str, rows: List[Dict[str, Any]]) -> List[str]:
        out = [f"[{kind}]"]
        for i, r in enumerate(rows, start=1):
            if not any((r.get(k) for k in ["week","topic","hours","methods","assessment","clos","gas"])):
                continue
            out.append(
                f"- Row {i}: week={r.get('week','')}, hours={r.get('hours','')}, "
                f"topic={r.get('topic','')}, CLOs={r.get('clos','')}, GAs={r.get('gas','')}, "
                f"methods={r.get('methods','')}, assessment={r.get('assessment','')}"
            )
        return out
    lines = fmt("Theory", theory) + [""] + fmt("Practical", practical)
    return "\n".join(lines).strip()

AUDIT_SYSTEM_PROMPT = """You are an academic course handout auditor for UTAS.
You must output ONLY valid JSON matching the exact contract provided by the user.
Use the CDP snapshot (goals/CLOs/topics/sources) as ground truth.
Use ONLY the uploaded handouts as evidence; be conservative.
If you cannot confidently support a judgement, set evidence="insufficient" and explain what is missing in remarks.
If previous handouts are provided, comment ONLY on observable differences; otherwise mark updated_vs_previous as not_applicable.
Never include markdown fences or extra commentary outside JSON."""

def prepare_llm_payload(
    cdp_bundle: Dict[str, Any],
    corpus: List[HandoutChunk],
    corpus_prev: Optional[List[HandoutChunk]] = None,
) -> Dict[str, Any]:
    course = (cdp_bundle or {}).get("course", {}) or {}
    cdp_snapshot = {
        "course_code": course.get("course_code", ""),
        "course_title": course.get("course_title", ""),
        "semester": course.get("semester", ""),
        "academic_year": course.get("academic_year", ""),
        "goals": (cdp_bundle or {}).get("goals_text", "") or "",
        "clos": (cdp_bundle or {}).get("clos_table", "") or "",
        "weekly_topics": _flatten_weekly(cdp_bundle),
        "sources": (cdp_bundle or {}).get("sources_text", "") or "",
    }

    current_compact = [_compact_chunk(c) for c in corpus if (c.get("text") or "").strip()]
    prev_compact = None
    if corpus_prev:
        prev_compact = [_compact_chunk(c) for c in corpus_prev if (c.get("text") or "").strip()]

    user_payload = {
        "cdp_snapshot": cdp_snapshot,
        "handouts_current": current_compact,
        "handouts_previous": prev_compact,
        "contract": {
            "pc_review": [
                {"criterion": "coverage_vs_outcomes", "rating": "below|meet|above", "evidence": "...", "remarks": "..."},
                {"criterion": "updated_vs_previous",  "rating": "below|meet|above|not_applicable", "evidence": "...", "remarks": "..."},
                {"criterion": "innovative_methods",   "rating": "below|meet|above", "evidence": "...", "remarks": "..."},
                {"criterion": "logical_sequence",     "rating": "below|meet|above", "evidence": "...", "remarks": "..."},
            ],
            "cc_review": [
                {"criterion": "template_utilized",      "yes_no": "yes|no", "evidence": "...", "remarks": "..."},
                {"criterion": "outcomes_per_chapter",   "yes_no": "yes|no", "evidence": "...", "remarks": "..."},
                {"criterion": "numbering_structure",    "yes_no": "yes|no", "evidence": "...", "remarks": "..."},
                {"criterion": "linguistic_clarity",     "yes_no": "yes|no", "evidence": "...", "remarks": "..."},
                {"criterion": "organization",           "yes_no": "yes|no", "evidence": "...", "remarks": "..."},
                {"criterion": "captions_present",       "yes_no": "yes|no", "evidence": "...", "remarks": "..."},
                {"criterion": "proper_citation",        "yes_no": "yes|no", "evidence": "...", "remarks": "..."},
                {"criterion": "relevant_references",    "yes_no": "yes|no", "evidence": "...", "remarks": "..."},
            ],
            "overall_summary": "...",
            "action_items": ["..."],
            "trace": [
                {"chunk": {"file": "...", "unit": "slide 7", "heading": "..."}, "highlights": ["..."], "linked_criteria": ["coverage_vs_outcomes"]}
            ],
        }
    }

    user_prompt = (
        "Evaluate the course handouts against the CDP snapshot.\n"
        "Return ONLY JSON that matches the contract inside this payload.\n\n"
        + json.dumps(user_payload, ensure_ascii=False)
    )

    return {
        "system": AUDIT_SYSTEM_PROMPT,
        "user": user_prompt,
        "cdp_snapshot": cdp_snapshot,
    }

def _extract_json(text: str) -> Dict[str, Any]:
    # Claude is usually clean, but guard anyway.
    s = (text or "").strip()
    if not s:
        raise ValueError("Empty model response.")
    if s[0] != "{":
        i = s.find("{")
        j = s.rfind("}")
        if i >= 0 and j > i:
            s = s[i:j+1]
    return json.loads(s)

def run_handout_audit(
    payload: Dict[str, Any],
    api_key: str,
    model: str = "anthropic/claude-3.5-sonnet",
    app_url: Optional[str] = None,
    app_title: str = "UTAS CDP Builder",
    timeout_s: int = 120,
) -> Dict[str, Any]:
    if not api_key:
        raise RuntimeError("Missing OpenRouter API key.")

    body = {
        "model": model,
        "messages": [
            {"role": "system", "content": payload["system"]},
            {"role": "user", "content": payload["user"]},
        ],
        "temperature": 0.1,
        "max_tokens": 2400,
    }

    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json",
    }
    if app_url:
        headers["HTTP-Referer"] = app_url
    if app_title:
        headers["X-Title"] = app_title

    r = requests.post(
        "https://openrouter.ai/api/v1/chat/completions",
        headers=headers,
        json=body,
        timeout=timeout_s,
    )
    if r.status_code >= 400:
        raise RuntimeError(f"OpenRouter error {r.status_code}: {r.text[:4000]}")

    data = r.json()
    txt = (
        data.get("choices", [{}])[0]
            .get("message", {})
            .get("content", "")
    )
    out = _extract_json(txt)

    # Minimal normalization/guardrails
    out.setdefault("pc_review", [])
    out.setdefault("cc_review", [])
    out.setdefault("overall_summary", "")
    out.setdefault("action_items", [])
    out.setdefault("trace", [])
    return out

def ensure_audit_summary_template(template_path: Union[str, Path]) -> Path:
    template_path = Path(template_path)
    template_path.parent.mkdir(parents=True, exist_ok=True)
    if template_path.exists():
        return template_path

    doc = Document()
    doc.add_heading("Course Handout Audit Summary", level=1)
    doc.add_paragraph("Course: {{ course_code }} — {{ course_title }}")
    doc.add_paragraph("Academic Year: {{ academic_year }}   |   Semester: {{ semester }}")
    doc.add_paragraph("Generated at: {{ generated_at }}")

    doc.add_heading("Program Coordinator Review", level=2)
    doc.add_paragraph("{% for r in pc_review %}")
    doc.add_paragraph("• {{ r.criterion }} — {{ r.rating }}")
    doc.add_paragraph("  Evidence: {{ r.evidence }}")
    doc.add_paragraph("  Remarks: {{ r.remarks }}")
    doc.add_paragraph("{% endfor %}")

    doc.add_heading("Curriculum Committee Review", level=2)
    doc.add_paragraph("{% for r in cc_review %}")
    doc.add_paragraph("• {{ r.criterion }} — {{ r.yes_no }}")
    doc.add_paragraph("  Evidence: {{ r.evidence }}")
    doc.add_paragraph("  Remarks: {{ r.remarks }}")
    doc.add_paragraph("{% endfor %}")

    doc.add_heading("Overall Summary", level=2)
    doc.add_paragraph("{{ overall_summary }}")

    doc.add_heading("Action Items", level=2)
    doc.add_paragraph("{% for a in action_items %}• {{ a }}{% endfor %}")

    doc.save(str(template_path))
    return template_path

def render_audit_summary_docx(
    template_path: Union[str, Path],
    audit_json: Dict[str, Any],
    cdp_snapshot: Dict[str, Any],
) -> bytes:
    template_path = ensure_audit_summary_template(template_path)
    tpl = DocxTemplate(str(template_path))

    ctx = {
        "course_code": cdp_snapshot.get("course_code", ""),
        "course_title": cdp_snapshot.get("course_title", ""),
        "academic_year": cdp_snapshot.get("academic_year", ""),
        "semester": cdp_snapshot.get("semester", ""),
        "generated_at": time.strftime("%Y-%m-%d %H:%M"),
        "pc_review": audit_json.get("pc_review", []) or [],
        "cc_review": audit_json.get("cc_review", []) or [],
        "overall_summary": audit_json.get("overall_summary", "") or "",
        "action_items": audit_json.get("action_items", []) or [],
    }
    tpl.render(ctx)

    buf = io.BytesIO()
    tpl.save(buf)
    return buf.getvalue()
