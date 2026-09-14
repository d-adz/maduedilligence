# M&A Due Diligence — Contract Comparison Demo (RLB Steiermark)

Built for **Raiffeisen-Landesbank Steiermark** (Graz/Raaba, Austria) to
the 02 Aug 2026 brief from jbormann2: compare two versions of an
acquisition contract, detect clause-level changes, and flag material or
risky edits for legal review — with a before/after comparison, a short
explanation of what changed, and a confidence score.

Extended per the 10 Sep 2026 follow-up for Austria-specific due diligence:
flag language touching **insolvency**, **legal proceedings**, and
**IP/patent ownership**, using an M&A-relevant entity filter (organisations,
dates, monetary amounts, legal references, percentages) rather than a flat
keyword dump.

Adapted from the [Incom Leone contract-comparison demo](https://github.com/d-adz/leone-demo)
— same core engine (parsing, clause splitting, comparison), retargeted
from English/Slovenian to **English/German** for Austrian acquisitions.

## What's in this build

**Branding** — themed in Raiffeisen's yellow-and-black identity rather than
a generic colour scheme: an inline SVG rendering of the gable-cross
(Giebelkreuz) emblem appears in the sidebar and main header, and
`.streamlit/config.toml` sets yellow as the app-wide Streamlit theme
colour (buttons, radios, checkboxes, sliders). The gable-cross is a
simplified geometric rendition built from scratch, not a traced copy of
RLB Steiermark's actual registered logo file — swap in their real logo
asset if pixel-accuracy matters.

**Full bilingual UI toggle** — a `🌐 Language / Sprache` selector in the
sidebar switches every label, button, and message across all four pages
between English and Austrian German instantly (client-side string swap,
not a live translation call — no lag, no API dependency). This is
separate from the contract-language detection below, which works on the
uploaded documents themselves regardless of which UI language is active.

**AI-powered Chat and Image Q&A** — both pages call the **Gemini API**
(`gemini-2.5-flash`, via the `google-genai` SDK) rather than showing
placeholder text:
- **Chat** is grounded in whatever's been uploaded to the **Library**
  page — uploaded documents' extracted text is fed into Gemini as system
  context, so questions get answered against the actual due-diligence
  corpus rather than generically.
- **Image Q&A** sends the uploaded image directly to Gemini for a real
  visual answer.
- If no API key is configured, both pages **fall back cleanly to mock
  placeholder replies** rather than erroring out — the app never crashes
  from a missing key.

**API key setup** — checked in this order:
1. `st.secrets["GEMINI_API_KEY"]` — for a real deployment, set this in
   `.streamlit/secrets.toml` locally (must be excluded via `.gitignore` —
   **never commit this file**, since it holds a real credential) or in
   Streamlit Community Cloud's Secrets panel if deployed there.
2. A session-only password field in the sidebar, shown only when no
   secret is found — convenient for quick local testing without touching
   config files. Never written to disk.

Get a free key with no credit card at **aistudio.google.com**. If
Legal-Pythia already has a shared Google AI Studio project with a paid
tier (worth checking with jbormann2/Otmane before using your own), that
would give higher rate limits than a personal free-tier key.

## Run it

```
pip install -r requirements.txt
streamlit run app.py
```

**Python version:** requires a currently-supported Python (3.10+
recommended, 3.12/3.13 confirmed working). Python 3.8 is end-of-life and
Google's SDKs have been dropping support for it — if you hit strange
`pip install` results (very old/unexpected package versions installing),
check `python --version` first.

Then upload or paste the two sample contracts
(`sample_contract_english.txt`, `sample_contract_german.txt`) — they're
deliberately written with differences so the demo has something real to
flag:

- **Purchase price raised** from EUR 4.2m to EUR 4.85m (Clause changed)
- **Payment term extended** from 30 to 45 days (Clause changed)
- **Insolvency history disclosed** — a prior 2021 bankruptcy proceeding,
  now closed, surfaces in the revised version (Due-diligence flag: Insolvency)
- **Pending litigation newly disclosed** — a live dispute with a former
  supplier appears only in the revised version (Due-diligence flag: Legal
  proceedings, Risk increased)
- **IP ownership gap revealed** — a core patent turns out to be licensed
  from a third party, not owned outright (Due-diligence flag: IP/patent
  ownership, Risk increased)
- **Liability cap removed entirely** — capped at EUR 1m in the original,
  unlimited in the revised version (Risk increased)
- **Notice period shortened** from sixty to thirty days (Clause changed)

Sample contracts feature a fictional acquisition of a Graz-based target
(Grazer Antriebstechnik GmbH) by a fictional Styrian holding company, to
keep the demo grounded in RLB Steiermark's own region rather than a
generic placeholder deal.

## What's real vs mocked

- File parsing (TXT/DOCX/PDF), language detection, clause splitting, and
  word-overlap comparison are all functional, not faked.
- Language detection is a lightweight heuristic (German function words +
  umlauts/ß), not a full language-ID model — fine for a demo, worth
  swapping for `langdetect`/`fasttext` if this goes further.
- Clause matching is naive word-overlap after a German→English glossary
  translation pass, same tradeoff as the Leone demo — fine for a demo,
  worth replacing with embeddings for production use.
- **Due-diligence category flags** (insolvency / legal proceedings / IP)
  are keyword-based, matched against the contract text only. They do
  **not** query the actual registers — RIS (UGB/ABGB), the e-Justice EU
  insolvency register, or USP.gv.at — those are listed as next steps, not
  wired in.
- **Entity extraction** is regex-based (organisations via GmbH/AG/KG
  suffixes, dates, monetary amounts, percentages, § legal references),
  standing in for a real NER pipeline (e.g. spaCy `de_core_news_lg`). It
  intentionally requires context (a currency symbol, %, date format, or
  legal-citation prefix) before matching, so isolated numbers and generic
  phrases are excluded by construction rather than filtered after the
  fact — matching the "default M&A-relevant view + toggle to show
  everything" approach discussed with jbormann2.
- **Chat and Image Q&A are genuinely AI-powered** (Gemini), not mocked —
  but only when an API key is configured; otherwise both fall back to
  placeholder text automatically.

## Layout

Same as the Leone demo's brief: Contract A (left) | Contract B (middle) |
Differences & risk panel (right), with labels "Clause changed", "Clause
missing", "Risk increased", "Why this matters", "Business impact" — plus
a due-diligence category callout and an extracted-entities list per
flagged clause (with a toggle to show all extracted entities vs. the
M&A-relevant default view).

## Next steps

- Test live with the sample contracts to confirm the layout reads well
- Wire the due-diligence flags to actual register lookups (RIS OGD-RIS
  API, e-Justice EU) rather than keyword matching alone, if this needs to
  move beyond demo stage
- Swap the language detector and entity extractor for proper libraries
  (`langdetect`/`fasttext`, spaCy German NER) if time allows
- Confirm with jbormann2/Otmane whether to use Legal-Pythia's existing
  Google AI Studio project (paid tier, higher limits) instead of a
  personal free-tier Gemini key, if this gets shared beyond individual
  testing
- Consider adding RLB Steiermark's actual logo/brand assets if this needs
  to go beyond the simplified inline SVG gable-cross rendition