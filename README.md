# M&A Due Diligence — Contract Comparison Demo

Built to the 02 Aug 2026 brief from jbormann2: compare two versions of an
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

The whole page UI (not just the contract text) can be switched between
English and Austrian German via the language selector at the top of the
sidebar — this is an instant client-side toggle over a fixed set of UI
strings, not a live translation call, so there's no lag and no external
API dependency. Austrian German conventions are used where they differ
from standard German phrasing.

## Run it

```
pip install -r requirements.txt
streamlit run app.py
```

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

## Layout

Same as the Leone demo's brief: Contract A (left) | Contract B (middle) |
Differences & risk panel (right), with labels "Clause changed", "Clause
missing", "Risk increased", "Why this matters", "Business impact" — plus
a new due-diligence category callout and an extracted-entities list per
flagged clause.

## Next steps

- Test live with the sample contracts to confirm the layout reads well
- Wire the due-diligence flags to actual register lookups (RIS OGD-RIS
  API, e-Justice EU) rather than keyword matching alone, if this needs to
  move beyond demo stage
- Swap the language detector and entity extractor for proper libraries
  (`langdetect`/`fasttext`, spaCy German NER) if time allows
- Consider adding a firm logo/colour accent if there's a specific client
  or internal brand to match