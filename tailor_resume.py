#!/usr/bin/env python3
"""
tailor_resume.py — tailor a .docx resume to a job description using the Claude API,
editing text in place so fonts/layout/design are preserved.

Usage:
    export ANTHROPIC_API_KEY=sk-ant-...
    python tailor_resume.py --resume my_resume.docx --jd "https://linkedin.com/jobs/..." --out tailored.docx
    python tailor_resume.py --resume my_resume.docx --jd jd.txt --out tailored.docx --review
    python tailor_resume.py --resume my_resume.docx --jd "paste the JD text directly here" --out tailored.docx

Flags:
    --resume PATH       Path to your master resume .docx (also acts as the layout template)
    --jd TEXT           A URL, a path to a .txt file, or raw pasted JD text
    --out PATH          Output .docx path (default: <resume>_tailored.docx)
    --model NAME        claude-sonnet-5 (default), claude-haiku-4-5-20251001, or claude-opus-5
    --review            Pause after keyword extraction so you can edit the list
    --dry-run           Show keywords + proposed edits, write nothing
    --max-change-pct N  Max allowed length change per paragraph before flagging (default 25)
    --instructions PATH Text file of persistent instructions for Claude (default: instructions.txt)
    --masterfile PATH   Markdown file of your fuller work history for Claude to draw on (default: masterfile.md)
    --apply-edits PATH  Skip the API and apply a saved edits JSON file to --resume, writing --out
"""

import argparse
import json
import os
import re
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

import requests

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
ET.register_namespace("w", W_NS)
NSMAP = {"w": W_NS}

API_URL = "https://api.anthropic.com/v1/messages"
ANTHROPIC_VERSION = "2023-06-01"

# Per-million-token USD rates, current as of Sep 2026. Update if pricing changes:
# https://platform.claude.com/docs/en/about-claude/pricing
RATES = {
    "claude-haiku-4-5-20251001": (1.0, 5.0),
    "claude-sonnet-5": (2.0, 10.0),
    "claude-opus-5": (5.0, 25.0),
}
DEFAULT_MODEL = "claude-sonnet-5"


# --------------------------------------------------------------------------
# JD loading
# --------------------------------------------------------------------------

def load_jd(jd_arg: str) -> str:
    if jd_arg.startswith("http://") or jd_arg.startswith("https://"):
        return fetch_jd_url(jd_arg)
    p = Path(jd_arg)
    if p.exists() and p.is_file():
        return p.read_text(encoding="utf-8", errors="ignore")
    return jd_arg  # treat as raw pasted text


def fetch_jd_url(url: str) -> str:
    try:
        from bs4 import BeautifulSoup
    except ImportError:
        sys.exit("Missing dependency: pip install beautifulsoup4")

    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"
    }
    try:
        resp = requests.get(url, headers=headers, timeout=15)
        resp.raise_for_status()
    except requests.RequestException as e:
        sys.exit(
            f"Could not fetch JD URL ({e}).\n"
            f"LinkedIn and some job boards block scrapers or require login.\n"
            f"Workaround: copy the JD text into a .txt file and pass that instead, "
            f"or paste the raw text directly as the --jd argument."
        )

    soup = BeautifulSoup(resp.text, "html.parser")
    for tag in soup(["script", "style", "nav", "header", "footer", "noscript"]):
        tag.decompose()
    text = re.sub(r"\n\s*\n+", "\n\n", soup.get_text("\n")).strip()

    if len(text) < 200:
        sys.exit(
            "Fetched page had very little text — it's likely JS-rendered or login-walled "
            "(common on LinkedIn). Paste the JD text into a .txt file and pass that instead."
        )
    return text


# --------------------------------------------------------------------------
# DOCX read/write — edits text inside existing runs, preserves everything else
# --------------------------------------------------------------------------

def qn(tag):
    return f"{{{W_NS}}}{tag}"


def extract_paragraphs(document_xml: bytes):
    """Return (tree, root, list of {id, text, elem}) for every w:p with visible text."""
    tree = ET.ElementTree(ET.fromstring(document_xml))
    root = tree.getroot()
    paragraphs = []
    for i, p in enumerate(root.iter(qn("p"))):
        runs = p.findall(f"./{qn('r')}")
        texts = []
        for r in runs:
            for t in r.findall(qn("t")):
                if t.text:
                    texts.append(t.text)
        full_text = "".join(texts).strip()
        if full_text:
            paragraphs.append({"id": i, "text": full_text, "elem": p})
    return tree, root, paragraphs


def apply_edits(paragraphs, edits: dict, max_change_pct: int):
    """edits: {paragraph_id(str or int): new_text}. Mutates XML elements in place."""
    flagged = []
    by_id = {p["id"]: p for p in paragraphs}

    for pid_raw, new_text in edits.items():
        pid = int(pid_raw)
        if pid not in by_id:
            continue
        para = by_id[pid]
        elem = para["elem"]
        old_text = para["text"]

        change_pct = abs(len(new_text) - len(old_text)) / max(len(old_text), 1) * 100
        if change_pct > max_change_pct:
            flagged.append((pid, old_text, new_text, change_pct))

        runs = elem.findall(f"./{qn('r')}")
        if not runs:
            continue

        # Put full new text in the first run's first <w:t>, preserving its rPr (formatting).
        # Clear text from any other <w:t> elements in this paragraph (collapses inline
        # sub-formatting to the first run's style — a known trade-off, see README).
        first_run = runs[0]
        first_t = first_run.find(qn("t"))
        if first_t is None:
            first_t = ET.SubElement(first_run, qn("t"))
        first_t.text = new_text
        first_t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")

        for r in runs[1:]:
            for t in r.findall(qn("t")):
                t.text = ""
        for extra_t in first_run.findall(qn("t"))[1:]:
            extra_t.text = ""

    return flagged


def edit_docx(resume_path: Path, edits: dict, out_path: Path, max_change_pct: int):
    with tempfile.TemporaryDirectory() as tmp:
        tmp = Path(tmp)
        with zipfile.ZipFile(resume_path, "r") as z:
            z.extractall(tmp)

        doc_xml_path = tmp / "word" / "document.xml"
        document_xml = doc_xml_path.read_bytes()
        tree, root, paragraphs = extract_paragraphs(document_xml)
        flagged = apply_edits(paragraphs, edits, max_change_pct)

        tree.write(doc_xml_path, xml_declaration=True, encoding="UTF-8", method="xml")

        if out_path.exists():
            out_path.unlink()
        with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for f in tmp.rglob("*"):
                if f.is_file():
                    zout.write(f, f.relative_to(tmp))

    return flagged


def get_editable_paragraphs(resume_path: Path):
    with zipfile.ZipFile(resume_path, "r") as z:
        document_xml = z.read("word/document.xml")
    _, _, paragraphs = extract_paragraphs(document_xml)
    return [{"id": p["id"], "text": p["text"]} for p in paragraphs]


# --------------------------------------------------------------------------
# Claude API calls — strict JSON-only, no commentary
# --------------------------------------------------------------------------

def call_claude(model: str, system: str, user: str, max_tokens: int = 2000):
    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        sys.exit("Set ANTHROPIC_API_KEY in your environment first.")

    try:
        resp = requests.post(
            API_URL,
            headers={
                "x-api-key": api_key,
                "anthropic-version": ANTHROPIC_VERSION,
                "content-type": "application/json",
            },
            json={
                "model": model,
                "max_tokens": max_tokens,
                # Cached so repeated runs (e.g. tailoring for several jobs in one
                # sitting) only pay full price for the system prompt once.
                "system": [{"type": "text", "text": system, "cache_control": {"type": "ephemeral"}}],
                "messages": [{"role": "user", "content": user}],
            },
            timeout=60,
        )
        resp.raise_for_status()
    except requests.RequestException as e:
        sys.exit(f"Claude API call failed: {e}")

    data = resp.json()
    text = "".join(b.get("text", "") for b in data.get("content", []) if b.get("type") == "text")
    usage = data.get("usage", {"input_tokens": 0, "output_tokens": 0})
    return text, usage, data.get("stop_reason")


def parse_json_strict(text: str, what: str, stop_reason: str = None, dump_path: Path = None):
    cleaned = re.sub(r"^```(json)?|```$", "", text.strip(), flags=re.MULTILINE).strip()
    try:
        return json.loads(cleaned)
    except json.JSONDecodeError:
        if dump_path:
            dump_path.write_text(text, encoding="utf-8")
        if stop_reason == "max_tokens":
            sys.exit(
                f"Claude's response for {what} was cut off before finishing (hit the output "
                f"token limit) — try a shorter job description, a shorter resume, or raise "
                f"max_tokens for that call in the script."
                + (f" Raw (truncated) output saved to {dump_path}." if dump_path else "")
            )
        sys.exit(
            f"Model did not return valid JSON for {what}."
            + (f" Raw output saved to {dump_path} — fix it by hand, then rerun with "
               f"--apply-edits {dump_path}." if dump_path else f" Raw output:\n{text}")
        )


def with_custom_instructions(system: str, custom_instructions: str) -> str:
    if not custom_instructions:
        return system
    return system + "\n\nAlso follow these persistent instructions from the user:\n" + custom_instructions


def extract_keywords(jd_text: str, model: str, custom_instructions: str = ""):
    system = (
        "You extract ATS-relevant keywords from a job description. "
        "Respond with ONLY a JSON object, no prose, no markdown fences, no preamble. "
        'Schema: {"must_have": ["..."], "nice_to_have": ["..."]}. '
        "Keep each list to at most 15 concise items (skills, tools, certifications, "
        "domain terms) — not full sentences."
    )
    system = with_custom_instructions(system, custom_instructions)
    text, usage, stop_reason = call_claude(model, system, jd_text, max_tokens=800)
    return parse_json_strict(text, "keyword extraction", stop_reason), usage


def rewrite_paragraphs(
    paragraphs,
    keywords: dict,
    jd_text: str,
    model: str,
    custom_instructions: str = "",
    masterfile_text: str = "",
    dump_path: Path = None,
):
    system = (
        "You tailor resume text to a job description without changing its factual claims, "
        "job titles, dates, employers, or overall length by more than roughly 15%. "
        "You are given a numbered list of resume paragraphs and a target keyword list. "
        "Rewrite ONLY the paragraphs that are resume content suitable for tailoring "
        "(summary lines, skills lists, bullet points describing responsibilities/achievements). "
        "Do NOT touch headers, contact info, dates, company names, section titles, or education "
        "credentials — omit those paragraph ids entirely from your output. "
        "Naturally weave in keywords from the list only where truthful and relevant; "
        "never fabricate skills or experience not implied by the original text. "
        "Respond with ONLY a JSON object mapping paragraph id (as a string) to the new text, "
        'e.g. {"4": "new text", "7": "new text"}. No prose, no markdown fences, no commentary.'
    )
    if masterfile_text:
        system += (
            "\n\nYou are also given the candidate's fuller personal work record below (their "
            "masterfile) — richer than the resume itself. Use it only to pull a more specific, "
            "truthful detail (a metric, tool, or outcome) into a paragraph you are already "
            "rewriting, when it's a better fit for that paragraph's original topic than what's "
            "there now. Never use it to add a fact unconnected to that paragraph's topic, invent "
            "an achievement, or change an employer, title, or date.\n\n"
            "=== MASTERFILE ===\n" + masterfile_text
        )
    system = with_custom_instructions(system, custom_instructions)
    user = json.dumps(
        {
            "job_description": jd_text[:6000],
            "target_keywords": keywords,
            "resume_paragraphs": paragraphs,
        }
    )
    text, usage, stop_reason = call_claude(model, system, user, max_tokens=4096)
    return parse_json_strict(text, "paragraph rewrite", stop_reason, dump_path), usage


# --------------------------------------------------------------------------
# Cost estimate
# --------------------------------------------------------------------------

def estimate_cost(model: str, usages: list):
    # Cache writes cost 1.25x the normal input rate, cache reads 0.1x — same ratios
    # Anthropic applies to every model, so no separate rate table is needed here.
    in_rate, out_rate = RATES.get(model, RATES[DEFAULT_MODEL])
    total_in = sum(u.get("input_tokens", 0) for u in usages)
    total_out = sum(u.get("output_tokens", 0) for u in usages)
    cache_write = sum(u.get("cache_creation_input_tokens", 0) for u in usages)
    cache_read = sum(u.get("cache_read_input_tokens", 0) for u in usages)
    cost = (
        (total_in / 1_000_000) * in_rate
        + (total_out / 1_000_000) * out_rate
        + (cache_write / 1_000_000) * in_rate * 1.25
        + (cache_read / 1_000_000) * in_rate * 0.1
    )
    return cost, total_in, total_out, cache_read


# --------------------------------------------------------------------------
# Main
# --------------------------------------------------------------------------

def main():
    ap = argparse.ArgumentParser(description="Tailor a .docx resume to a job description.")
    ap.add_argument("--resume", required=True, type=Path)
    ap.add_argument("--jd", help="Required unless --apply-edits is given")
    ap.add_argument("--out", type=Path)
    ap.add_argument(
        "--apply-edits",
        type=Path,
        help="Skip the API entirely and apply a JSON file of {paragraph_id: new_text} "
        "edits straight to --resume, writing --out. Use this to retry applying edits "
        "you already paid for — e.g. the *_edits.json saved after a normal run, or a "
        "*_rewrite_raw.txt you fixed by hand.",
    )
    ap.add_argument("--model", default=DEFAULT_MODEL, choices=list(RATES.keys()))
    ap.add_argument("--review", action="store_true", help="Edit keyword list before rewriting")
    ap.add_argument("--dry-run", action="store_true", help="Show plan, write nothing")
    ap.add_argument("--max-change-pct", type=int, default=25)
    ap.add_argument(
        "--instructions",
        type=Path,
        default=Path("instructions.txt"),
        help="Text file of persistent instructions for Claude (tone, style, things to "
        "always/never do). Reused as-is across job submissions, so it's cheap to keep "
        "loading it — see README. Default: instructions.txt (skipped if missing).",
    )
    ap.add_argument(
        "--masterfile",
        type=Path,
        default=Path("masterfile.md"),
        help="Markdown file of your fuller work history/experience, beyond what's in the "
        "resume itself — Claude can pull specific details from it into a bullet it's "
        "already rewriting. Reused as-is across job submissions — see README for how "
        "this stays cheap. Default: masterfile.md (skipped if missing).",
    )
    args = ap.parse_args()

    if not args.resume.exists():
        sys.exit(f"Resume not found: {args.resume}")
    out_path = args.out or args.resume.with_name(args.resume.stem + "_tailored.docx")

    if args.apply_edits:
        if not args.apply_edits.exists():
            sys.exit(f"Edits file not found: {args.apply_edits}")
        edits = json.loads(args.apply_edits.read_text(encoding="utf-8"))
        flagged = edit_docx(args.resume, edits, out_path, args.max_change_pct)
        print(f"✓ Saved {out_path} | {len(edits)} paragraph(s) updated from {args.apply_edits}")
        if flagged:
            print(f"⚠ {len(flagged)} paragraph(s) changed length by >{args.max_change_pct}% — review layout:")
            for pid, old, new, pct in flagged:
                print(f"  [{pid}] {pct:.0f}% change")
        return

    if not args.jd:
        sys.exit("--jd is required unless --apply-edits is given")

    custom_instructions = ""
    if args.instructions.exists():
        custom_instructions = args.instructions.read_text(encoding="utf-8").strip()

    masterfile_text = ""
    if args.masterfile.exists():
        masterfile_text = args.masterfile.read_text(encoding="utf-8").strip()

    jd_text = load_jd(args.jd)
    paragraphs = get_editable_paragraphs(args.resume)

    keywords, usage1 = extract_keywords(jd_text, args.model, custom_instructions)

    if args.review:
        print("\nMust-have:", ", ".join(keywords.get("must_have", [])))
        print("Nice-to-have:", ", ".join(keywords.get("nice_to_have", [])))
        edit = input("\nEdit? Enter comma-separated must-have list to replace it, or press Enter to keep: ").strip()
        if edit:
            keywords["must_have"] = [k.strip() for k in edit.split(",") if k.strip()]

    raw_dump_path = out_path.parent / (out_path.stem + "_rewrite_raw.txt")
    edits, usage2 = rewrite_paragraphs(
        paragraphs, keywords, jd_text, args.model, custom_instructions, masterfile_text, dump_path=raw_dump_path
    )

    # Saved regardless of --dry-run so a paid-for rewrite is never lost — reapply anytime
    # with --apply-edits, no API call needed.
    edits_path = out_path.parent / (out_path.stem + "_edits.json")
    edits_path.write_text(json.dumps(edits, indent=2), encoding="utf-8")

    cost, total_in, total_out, cache_read = estimate_cost(args.model, [usage1, usage2])
    cache_note = f" ({cache_read} cached)" if cache_read else ""

    if args.dry_run:
        print(f"\n[dry-run] {len(edits)} paragraph(s) would change:")
        for pid, new_text in edits.items():
            print(f"  [{pid}] {new_text[:100]}")
        print(f"\nEst. cost: ${cost:.4f} ({total_in} in{cache_note} / {total_out} out tokens)")
        print(f"Edits saved to {edits_path} (apply later with --apply-edits {edits_path})")
        return

    flagged = edit_docx(args.resume, edits, out_path, args.max_change_pct)

    print(f"✓ Saved {out_path} | {len(edits)} paragraph(s) updated | Est. cost: ${cost:.4f}{cache_note}")
    if flagged:
        print(f"⚠ {len(flagged)} paragraph(s) changed length by >{args.max_change_pct}% — review layout:")
        for pid, old, new, pct in flagged:
            print(f"  [{pid}] {pct:.0f}% change")


if __name__ == "__main__":
    main()
