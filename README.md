# resume_tailor

Local CLI that tailors a `.docx` resume to a job description using the Claude API,
editing text in place so your formatting/layout is never disturbed.

## Setup

```bash
pip install -r requirements.txt
export ANTHROPIC_API_KEY=sk-ant-...
```

## Usage

```bash
# JD from a URL (non-LinkedIn pages work best — see note below)
python tailor_resume.py --resume resume.docx --jd "https://example.com/job/123"

# JD pasted straight into the terminal
python tailor_resume.py --resume resume.docx --jd "We are looking for a backend engineer..."

# JD saved to a text file
python tailor_resume.py --resume resume.docx --jd jd.txt --out tailored.docx

# Review/edit the extracted keyword list before rewriting
python tailor_resume.py --resume resume.docx --jd jd.txt --review

# See what would change without writing anything (also prints cost estimate)
python tailor_resume.py --resume resume.docx --jd jd.txt --dry-run
```

Output defaults to `<resume>_tailored.docx` next to the input file.

## Persistent instructions

To give Claude standing behavior you want applied to every job you tailor for (tone,
things to always avoid, a fixed way to phrase your summary line, etc.), create an
`instructions.txt` file next to the script and write it in plain English — it's picked
up automatically. Use a different file with `--instructions path/to/file.txt`. If the
file doesn't exist, the tool just runs without extra instructions.

## Reducing API cost across runs

The system prompt (built-in instructions + your `instructions.txt`, if any) is sent
with [prompt caching](https://platform.claude.com/docs/en/build-with-claude/prompt-caching)
enabled. As long as that text doesn't change, the second and later calls you make
within the ~5 minute cache window (e.g. tailoring your resume for several job postings
back-to-back) are billed a small fraction of the normal input rate for that part of the
prompt instead of full price. The cost line printed after each run shows how many
tokens were served from cache.

## What it does

1. Loads the job description (URL fetch, file, or raw text).
2. Extracts every text paragraph from your `.docx` and sends it, with the JD, to Claude
   to pull out must-have / nice-to-have keywords.
3. Sends the paragraph list + keywords + JD to Claude, which returns *only* the
   paragraph IDs that should change and their new text — headers, dates, company
   names, and education are explicitly left alone.
4. Rewrites just those paragraphs' text directly in the underlying XML, reusing the
   existing run's formatting, then re-zips the `.docx`. No regeneration, no template
   re-rendering — your fonts, colors, spacing, and tables are untouched.
5. Prints one summary line with an estimated API cost, computed from actual token
   usage returned by the API.
6. Saves the edits Claude returned to `<out>_edits.json`, so the API call is never
   wasted even if you want to re-apply it later.

## Re-applying edits without calling the API

Every normal run saves the paragraph edits it got back from Claude to
`<out>_edits.json` before writing the `.docx`. If you ever want to apply that same
JSON to the resume again (e.g. after tweaking `--max-change-pct`, or recovering from
a failed run), skip the API entirely:

```bash
python tailor_resume.py --resume resume.docx --apply-edits tailored_edits.json --out tailored.docx
```

This reads the JSON and edits the Word document directly — you never need to hand-edit
the JSON yourself. If Claude's raw response ever fails to parse as JSON, it's saved to
`<out>_rewrite_raw.txt` instead so nothing is lost; that's the one case where you'd
need to fix the text by hand before it can be applied.

## Known limitations

- **LinkedIn / JS-rendered job pages** often block simple HTTP fetches or require
  login. If the fetch fails or returns almost no text, copy the JD into a `.txt`
  file (or paste it directly as the `--jd` argument) instead of using the URL.
- **Mixed formatting within one paragraph** (e.g., a single bolded word in the
  middle of a sentence) is flattened to that paragraph's first run's style when the
  paragraph is edited. If a bullet has important inline formatting, check it after
  running.
- **Length changes**: paragraphs whose new text differs from the original by more
  than `--max-change-pct` (default 25%) are flagged in the output so you can check
  for line-wrap/overflow issues — the tool does not auto-shrink font size.

## Pricing

Costs are computed from real API usage using per-million-token rates for the model
you pass with `--model` (default `claude-sonnet-5`). Rates are hardcoded in the
script as of Sep 2026 — check
[Anthropic's pricing page](https://platform.claude.com/docs/en/about-claude/pricing)
if it's been a while, and update the `RATES` dict at the top of `tailor_resume.py`.
