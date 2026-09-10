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
