"""
Legal-Pythia — M&A Due Diligence Contract Comparison Demo
Built for Raiffeisen-Landesbank Steiermark (RLB Steiermark, Graz/Raaba,
Austria) — one-page bilingual (English/German) contract comparison demo
for Austrian M&A due diligence.

Per brief (02 Aug 2026, jbormann2):
- Build a demo that compares two versions of an acquisition contract,
  detects clause-level changes, and flags material or risky edits for
  legal review.
- Output: clear before/after comparison, a short explanation of what
  changed, and a confidence score so a lawyer can quickly review flagged
  items.

Per follow-up (10 Sep 2026, jbormann2) — Austria-specific due diligence:
- Check whether the target company is insolvent
- Check whether there are legal proceedings against them
- Check whether they own all their patents/IP
Reference sources supplied: RIS (UGB Commercial Code & ABGB Civil Code),
OGD-RIS API, e-Justice EU insolvency/bankruptcy registers, USP.gv.at
insolvency info, Baker McKenzie Austria M&A guide.

Structure and demo conventions carried over from the Incom Leone
contract-comparison build (same team, same brief style): mocking/
prefilling AI-style results where a full backend isn't needed is fine —
prioritise a convincing, useful demo over a "real" production pipeline.
"""

from __future__ import annotations
import re
import streamlit as st
from docx import Document
from pypdf import PdfReader

st.set_page_config(page_title="Raiffeisen-Landesbank Steiermark — M&A Due Diligence", layout="wide")

BRAND_YELLOW = "#FFCC00"
BRAND_DARK = "#000000"

# Simplified inline SVG rendering of the Raiffeisen gable-cross (Giebelkreuz)
# emblem — the traditional crossed gable-end finials mark used across the
# Austrian Raiffeisen banking group. Rendered as inline SVG (not a hotlinked
# image) so it always displays with no external dependency or broken-link
# risk. This is a simplified geometric rendition, not a pixel copy of any
# specific registered logo file.
GIEBELKREUZ_SVG = """
<svg width="42" height="42" viewBox="0 0 100 100" xmlns="http://www.w3.org/2000/svg">
  <circle cx="50" cy="50" r="48" fill="{yellow}" stroke="{dark}" stroke-width="2"/>
  <g fill="{dark}">
    <path d="M50 18 L74 40 L66 40 L66 30 L58 30 L58 40 L50 34 L42 40 L34 40 L34 30 L26 30 L26 40 L18 40 Z"/>
    <path d="M50 82 L74 60 L66 60 L66 70 L58 70 L58 60 L50 66 L42 60 L34 60 L34 70 L26 70 L26 60 L18 60 Z"/>
  </g>
</svg>
""".format(yellow=BRAND_YELLOW, dark=BRAND_DARK)

if "library_files" not in st.session_state:
    st.session_state.library_files = []
if "chat_history" not in st.session_state:
    st.session_state.chat_history = []
if "ui_lang" not in st.session_state:
    st.session_state.ui_lang = "English"


# ---------------------------------------------------------------------------
# 0. UI TRANSLATIONS (English / Austrian German)
#
# A page-wide toggle rather than a one-shot "translate" action: switching
# is instant (no API call, no re-render lag) and the choice persists across
# reruns via session_state. Austrian German conventions used where they
# differ from standard German (e.g. "Jänner" not "Januar") since this is
# specifically for Austrian M&A due diligence, not a generic DE locale.
# ---------------------------------------------------------------------------

UI_STRINGS = {
    "nav_caption": {"English": "Navigation", "German": "Navigation"},
    "sidebar_title": {"English": "RLB Steiermark — M&A Due Diligence", "German": "RLB Steiermark — M&A Due-Diligence-Prüfung"},
    "page_title": {"English": "Raiffeisen-Landesbank Steiermark — M&A Due Diligence", "German": "Raiffeisen-Landesbank Steiermark — M&A Due-Diligence-Prüfung"},
    "page_subtitle": {
        "English": "Acquisition Contract Comparison — English & German (Austria)",
        "German": "Vergleich von Unternehmenskaufverträgen — Englisch & Deutsch (Österreich)",
    },
    "page_caption": {
        "English": "Compares two versions of an acquisition contract clause by clause, detects "
                    "language automatically, flags material or risky edits, and highlights "
                    "language touching insolvency, legal proceedings, or IP/patent ownership.",
        "German": "Vergleicht zwei Fassungen eines Unternehmenskaufvertrags klauselweise, "
                   "erkennt die Sprache automatisch, markiert wesentliche oder riskante "
                   "Änderungen und hebt Formulierungen zu Insolvenz, Rechtsstreitigkeiten "
                   "oder geistigem Eigentum/Patenten hervor.",
    },
    "contract_a_header": {"English": "Contract A (earlier version)", "German": "Vertrag A (frühere Fassung)"},
    "contract_b_header": {"English": "Contract B (revised version)", "German": "Vertrag B (überarbeitete Fassung)"},
    "input_label": {"English": "Input", "German": "Eingabe"},
    "upload_option": {"English": "Upload", "German": "Hochladen"},
    "paste_option": {"English": "Paste", "German": "Einfügen"},
    "upload_a": {"English": "Upload Contract A", "German": "Vertrag A hochladen"},
    "upload_b": {"English": "Upload Contract B", "German": "Vertrag B hochladen"},
    "paste_a": {"English": "Paste Contract A text", "German": "Text von Vertrag A einfügen"},
    "paste_b": {"English": "Paste Contract B text", "German": "Text von Vertrag B einfügen"},
    "diff_header": {"English": "Differences & Risk", "German": "Unterschiede & Risiko"},
    "provide_both": {
        "English": "Provide both Contract A and Contract B to compare.",
        "German": "Bitte sowohl Vertrag A als auch Vertrag B angeben, um sie zu vergleichen.",
    },
    "comparing_spinner": {"English": "Comparing…", "German": "Wird verglichen…"},
    "detected_as": {"English": "**Contract A detected as:** {a}  \n**Contract B detected as:** {b}",
                     "German": "**Vertrag A erkannt als:** {a}  \n**Vertrag B erkannt als:** {b}"},
    "metric_risk": {"English": "Risk increased", "German": "Erhöhtes Risiko"},
    "metric_changed": {"English": "Clauses changed", "German": "Geänderte Klauseln"},
    "metric_missing": {"English": "Clauses missing", "German": "Fehlende Klauseln"},
    "metric_dd": {"English": "Due-diligence flags", "German": "Due-Diligence-Hinweise"},
    "dd_categories_header": {"English": "**Due-diligence categories touched:**",
                              "German": "**Betroffene Due-Diligence-Kategorien:**"},
    "dd_caption": {
        "English": "These are keyword-based flags from the contract text only. "
                    "Confirm status against RIS (UGB/ABGB), the e-Justice EU "
                    "insolvency register, and USP.gv.at before relying on them.",
        "German": "Dies sind rein textbasierte Hinweise aus dem Vertragstext. "
                   "Der Status ist vor einer Verwendung im RIS (UGB/ABGB), im "
                   "e-Justice-EU-Insolvenzregister und auf USP.gv.at zu bestätigen.",
    },
    "flagged_summary": {"English": "**{n} flagged clause(s)** out of {total} compared",
                         "German": "**{n} markierte Klausel(n)** von {total} verglichenen"},
    "show_all_entities": {
        "English": "Show all extracted entities (default view shows M&A-relevant only)",
        "German": "Alle erkannten Entitäten anzeigen (Standardansicht zeigt nur M&A-relevante)",
    },
    "contract_a_label": {"English": "**Contract A:**", "German": "**Vertrag A:**"},
    "contract_b_label": {"English": "**Contract B:**", "German": "**Vertrag B:**"},
    "why_matters": {"English": "**Why this matters:**", "German": "**Warum das wichtig ist:**"},
    "business_impact": {"English": "**Business impact:**", "German": "**Auswirkung auf das Geschäft:**"},
    "dd_categories_inline": {"English": "**Due-diligence categories:**", "German": "**Due-Diligence-Kategorien:**"},
    "extracted_entities": {"English": "**Extracted entities:**", "German": "**Erkannte Entitäten:**"},
    "standard_clauses": {"English": "{n} standard clause(s) (click to view)",
                          "German": "{n} unauffällige Klausel(n) (zum Anzeigen klicken)"},
    "similarity_label": {"English": "similarity", "German": "Übereinstimmung"},
    "library_header": {"English": "Library", "German": "Bibliothek"},
    "library_caption": {"English": "Documents stored in the corpus — PDFs are ingested for retrieval.",
                         "German": "Im Bestand gespeicherte Dokumente — PDFs werden zur Suche indexiert."},
    "ingest_button": {"English": "⬆ Ingest", "German": "⬆ Aufnehmen"},
    "ingest_upload_label": {"English": "Ingest a document", "German": "Dokument aufnehmen"},
    "no_documents": {"English": "No documents yet — click Ingest to upload one.",
                      "German": "Noch keine Dokumente — auf „Aufnehmen“ klicken, um eines hochzuladen."},
    "col_name": {"English": "**NAME**", "German": "**NAME**"},
    "col_type": {"English": "**TYPE**", "German": "**TYP**"},
    "chat_header": {"English": "Chat with your documents", "German": "Chat mit Ihren Dokumenten"},
    "chat_caption": {"English": "Retrieval-augmented answers, grounded in your corpus.",
                      "German": "Antworten auf Basis Ihres Dokumentenbestands."},
    "ask_anything": {"English": "Ask anything", "German": "Fragen Sie alles"},
    "ask_anything_sub": {"English": "Answers come with the exact chunks they were grounded in.",
                          "German": "Antworten werden mit den genauen Textstellen belegt, auf denen sie beruhen."},
    "chat_input_placeholder": {"English": "Ask a question — type @ to add context…",
                                "German": "Stellen Sie eine Frage — @ eingeben, um Kontext hinzuzufügen…"},
    "chat_mock_reply": {
        "English": "(demo) Based on the documents in your library, here's a placeholder answer to: \"{prompt}\"",
        "German": "(Demo) Basierend auf den Dokumenten in Ihrer Bibliothek, hier eine Platzhalterantwort auf: „{prompt}“",
    },
    "image_qa_header": {"English": "Image Q&A", "German": "Bild-Frage & Antwort"},
    "image_qa_caption": {"English": "Ask a question about an uploaded image or scan.",
                          "German": "Stellen Sie eine Frage zu einem hochgeladenen Bild oder Scan."},
    "image_qa_upload_label": {"English": "Pick an image from the library or upload one",
                               "German": "Ein Bild aus der Bibliothek auswählen oder hochladen"},
    "image_qa_info": {"English": "Select an image first, then ask away.",
                       "German": "Bitte zuerst ein Bild auswählen und dann fragen."},
    "image_qa_question_label": {"English": "Ask about this image", "German": "Frage zu diesem Bild"},
    "image_qa_ask_button": {"English": "Ask", "German": "Fragen"},
    "image_qa_mock_answer": {
        "English": "**(demo) Answer:** This is a placeholder response to \"{q}\" — wire up a real vision model to replace this.",
        "German": "**(Demo) Antwort:** Dies ist eine Platzhalterantwort auf „{q}“ — hier ein echtes Bildmodell einbinden.",
    },
}

def t(key: str, **kwargs) -> str:
    lang = st.session_state.get("ui_lang", "English")
    template = UI_STRINGS.get(key, {}).get(lang, UI_STRINGS.get(key, {}).get("English", key))
    return template.format(**kwargs) if kwargs else template


# ---------------------------------------------------------------------------
# 1. LANGUAGE DETECTION (lightweight, no external API needed)
# ---------------------------------------------------------------------------

GERMAN_MARKERS = {
    "und", "ist", "für", "auf", "sich", "der", "die", "das", "dass", "sind",
    "oder", "wenn", "wie", "vertrag", "vertrags", "vertrages", "partei",
    "parteien", "vereinbarung", "artikel", "paragraph", "verpflichtung",
    "verpflichtungen", "zahlung", "frist", "rechte", "pflichten", "käufer",
    "verkäufer", "unternehmen", "gesellschaft", "haftung", "kündigung",
}

def detect_language(text: str) -> str:
    words = set(re.findall(r"[a-zäöüß]+", text.lower()))
    de_hits = len(words & GERMAN_MARKERS)
    has_umlauts = bool(re.search(r"[äöüß]", text.lower()))
    if de_hits >= 2 or has_umlauts:
        return "German"
    return "English"


# ---------------------------------------------------------------------------
# 2. FILE PARSING
# ---------------------------------------------------------------------------

def extract_text(uploaded_file) -> str:
    name = uploaded_file.name.lower()
    if name.endswith(".txt"):
        return uploaded_file.read().decode("utf-8", errors="ignore")
    elif name.endswith(".docx"):
        doc = Document(uploaded_file)
        return "\n".join(p.text for p in doc.paragraphs)
    elif name.endswith(".pdf"):
        reader = PdfReader(uploaded_file)
        return "\n".join(page.extract_text() or "" for page in reader.pages)
    else:
        st.warning(f"Unsupported file type: {uploaded_file.name}")
        return ""


# ---------------------------------------------------------------------------
# 3. CLAUSE SPLITTING
# ---------------------------------------------------------------------------

def split_into_clauses(text: str) -> list[dict]:
    pattern = r"(?=^\s*(?:\d+\.|\§\s*\d+|Article\s+\d+|Artikel\s+\d+|Paragraph\s+\d+)[^\n]*)"
    chunks = re.split(pattern, text, flags=re.MULTILINE)
    clauses = []
    for chunk in chunks:
        chunk = chunk.strip()
        if not chunk:
            continue
        first_line = chunk.split("\n", 1)[0][:80]
        num_match = re.match(r"\s*(\d+)", first_line)
        clause_num = int(num_match.group(1)) if num_match else None
        clauses.append({"heading": first_line, "text": chunk, "number": clause_num})
    return clauses


# ---------------------------------------------------------------------------
# 4. RELEVANT-ENTITY EXTRACTION (M&A-focused NER-style filter)
#
# Lightweight, regex-based stand-in for a real NER pipeline (e.g. spaCy
# de_core_news_lg). Flags the entity types that matter for M&A review —
# organisations, dates, monetary amounts, legal references, percentages —
# while suppressing isolated numbers and generic phrases. This is the
# "default M&A-relevant view" jbormann2 suggested; a full extracted-entity
# list is always available via the toggle so the threshold can be
# validated against real due-diligence files before tightening it.
# ---------------------------------------------------------------------------

ENTITY_PATTERNS = {
    "Legal reference": re.compile(
        r"§+\s*\d+[a-z]?(?:\s*Abs\.?\s*\d+)?\s*(?:UGB|ABGB|GmbHG|AktG)?", re.IGNORECASE
    ),
    "Monetary amount": re.compile(
        r"(?:EUR|€|USD|\$)\s?[\d.,]+(?:\s?(?:million|billion|Mio\.?|Mrd\.?))?"
        r"|[\d.,]+\s?(?:EUR|€|USD|\$)"
    ),
    "Percentage": re.compile(r"\d+(?:[.,]\d+)?\s?%"),
    "Date": re.compile(
        r"\b\d{1,2}[./]\d{1,2}[./]\d{2,4}\b"
        r"|\b\d{1,2}\s(?:January|February|March|April|May|June|July|August|"
        r"September|October|November|December|Januar|Februar|März|April|Mai|"
        r"Juni|Juli|August|September|Oktober|November|Dezember)\s\d{4}\b"
    ),
    "Organisation": re.compile(
        r"\b[A-ZÄÖÜ][\w&.\-]*(?:\s[A-ZÄÖÜ][\w&.\-]*)*\s"
        r"(?:GmbH|AG|KG|OG|GesmbH|Ltd\.?|Inc\.?|LLC|Group|Holding)\b"
    ),
}

NOISE_PHRASES = {
    "the parties", "this agreement", "the company", "die parteien",
    "dieser vertrag", "das unternehmen",
}

def extract_relevant_entities(text: str) -> list[dict]:
    """Extract M&A-relevant entities from a clause, suppressing noise.

    Isolated numbers with no unit/currency/legal-reference context are not
    matched by the patterns above at all (they require a currency symbol,
    %, legal-citation prefix, or date format), so they're excluded by
    construction rather than filtered after the fact.
    """
    found = []
    for entity_type, pattern in ENTITY_PATTERNS.items():
        for m in pattern.finditer(text):
            value = m.group(0).strip()
            if not value or value.lower() in NOISE_PHRASES:
                continue
            found.append({"type": entity_type, "value": value})
    return found


# ---------------------------------------------------------------------------
# 5. DUE DILIGENCE RISK CATEGORIES
#
# Per jbormann2's three specific checks: insolvency, legal proceedings
# against the target, and IP/patent ownership. Keyword-based flagging
# (same tradeoff as the generic risk-keyword approach) — flags language
# in the contract that touches these areas so a reviewer can jump
# straight to the relevant clauses rather than reading the whole
# document. This does NOT query external registers (RIS, e-Justice,
# USP.gv.at) — those are the real data sources for actually verifying
# status, and are listed as a next step rather than wired in for the demo.
# ---------------------------------------------------------------------------

DD_CATEGORIES = {
    "Insolvency": {
        "keywords": [
            "insolven", "bankrupt", "liquidation", "over-indebted",
            "receivership", "insolvenz", "zahlungsunfähig", "überschuldet",
            "konkurs", "sanierungsverfahren",
        ],
        "why": "Insolvency proceedings or financial distress at the target "
               "materially affects deal viability and warrants an RIS/e-Justice "
               "insolvency register check before signing.",
    },
    "Legal proceedings": {
        "keywords": [
            "litigation", "lawsuit", "claim", "dispute", "proceeding",
            "arbitration", "rechtsstreit", "klage", "verfahren", "streitigkeit",
            "gerichtsverfahren",
        ],
        "why": "Pending or threatened legal action against the target is a "
               "standard due-diligence red flag and should be cross-checked "
               "against court and register records.",
    },
    "IP / patent ownership": {
        "keywords": [
            "patent", "intellectual property", "trademark", "copyright",
            "licence", "license", "know-how", "marke", "urheberrecht",
            "lizenz", "gewerbliches schutzrecht",
        ],
        "why": "Confirming the target actually owns (rather than licenses) its "
               "patents and IP is essential — unclear IP title is a common "
               "source of post-acquisition disputes.",
    },
}

def flag_due_diligence_categories(text: str) -> list[str]:
    lowered = text.lower()
    hits = []
    for category, info in DD_CATEGORIES.items():
        if any(kw in lowered for kw in info["keywords"]):
            hits.append(category)
    return hits


# ---------------------------------------------------------------------------
# 6. CLAUSE COMPARISON
# ---------------------------------------------------------------------------

RISK_KEYWORDS = [
    "liability", "indemnif", "termination", "penalty", "confidential",
    "warranty", "haftung", "schadenersatz", "kündigung", "vertragsstrafe",
    "vertraulichkeit", "gewährleistung",
] + [kw for cat in DD_CATEGORIES.values() for kw in cat["keywords"]]

def is_risky(text: str) -> bool:
    lowered = text.lower()
    return any(kw in lowered for kw in RISK_KEYWORDS)


DE_TO_EN_GLOSSARY = {
    "vertrag": "agreement", "vertrags": "agreement", "vertrages": "agreement",
    "vertragspartei": "party", "vertragsparteien": "party",
    "partei": "party", "parteien": "party", "vereinbarung": "agreement",
    "käufer": "buyer", "käufers": "buyer", "verkäufer": "seller",
    "verkäufers": "seller", "zielgesellschaft": "target company",
    "unternehmen": "company", "gesellschaft": "company",
    "jede": "each", "andere": "other", "keine": "neither", "beide": "both",
    "zahlung": "payment", "zahlungen": "payment", "zahlen": "pay",
    "rechnung": "invoice", "rechnungen": "invoice", "ausstellen": "issue",
    "zinsen": "interest", "jährlich": "annum", "verzug": "late",
    "laufzeit": "term", "dauert": "lasts", "beginnt": "commence",
    "wirksam": "effect", "unterzeichnung": "signing",
    "automatisch": "automatic", "verlängerung": "renewal",
    "kündigen": "terminate", "kündigung": "termination",
    "frist": "notice period", "schriftlich": "written", "schriftliche": "written",
    "sofern": "unless", "beliebige": "any", "welche": "which",
    "vertraulichkeit": "confidentiality", "vertrauliche": "confidential",
    "nichtöffentliche": "non-public", "geschäftliche": "business",
    "technische": "technical", "preisbezogene": "pricing",
    "informationen": "information", "offenlegt": "discloses",
    "offenlegen": "disclose", "dritten": "third", "personen": "party",
    "zustimmung": "consent", "vorherige": "prior", "vorheriger": "prior",
    "behalten": "keep", "nicht": "not", "darf": "may", "ohne": "without",
    "haftung": "liability", "haftbar": "liable", "schaden": "damages",
    "mittelbare": "indirect", "folgeschäden": "consequential",
    "begrenzt": "limited", "gesamt": "total", "gewährleistung": "warranty",
    "garantiert": "warrants", "übereinstimmen": "conform",
    "vereinbarten": "agreed", "spezifikationen": "specifications",
    "standards": "standards", "sicherheit": "safety",
    "produkte": "products", "produkten": "products",
    "gegenstand": "scope", "bestellungen": "orders", "beigefügten": "attached",
    "unternehmens": "company", "teil": "part",
    "recht": "law", "geltende": "governing", "anwendbare": "applicable",
    "österreich": "austria", "österreichs": "austria", "regelt": "governs",
    "patent": "patent", "patente": "patents", "marke": "trademark",
    "lizenz": "licence", "know-how": "know-how",
    "insolvenz": "insolvency", "konkurs": "bankruptcy",
    "und": "and", "ist": "is", "für": "for", "auf": "on", "sich": "itself",
    "die": "the", "der": "the", "das": "the", "dass": "that", "sind": "are",
    "oder": "or", "wenn": "if", "wie": "as", "alle": "all",
    "diese": "this", "dieser": "this", "dieses": "this", "muss": "must",
    "werden": "will", "nicht": "not", "ihnen": "them", "von": "from",
    "durch": "by", "bei": "at", "in": "in", "mit": "with", "über": "about",
    "während": "during",
}

def translate_known_terms(text: str) -> str:
    words = re.findall(r"[a-zäöüßA-ZÄÖÜ]+|\S", text.lower())
    return " ".join(DE_TO_EN_GLOSSARY.get(w, w) for w in words)


def highlight_key_facts_html(text: str) -> str:
    def repl_number(m):
        return f"<mark style='background:#FFE8B3;padding:1px 3px;border-radius:3px;'>{m.group(0)}</mark>"

    def repl_term(m):
        return f"<mark style='background:#D6E8FF;padding:1px 3px;border-radius:3px;'>{m.group(0)}</mark>"

    highlighted = re.sub(r"\d+(?:[.,]\d+)?%?", repl_number, text)

    term_pattern = "|".join(sorted(RISK_KEYWORDS, key=len, reverse=True))
    if term_pattern:
        highlighted = re.sub(
            rf"\b({term_pattern})\w*", repl_term, highlighted, flags=re.IGNORECASE
        )

    return highlighted


def compare_clauses(clauses_a: list[dict], clauses_b: list[dict]) -> list[dict]:
    results = []
    by_number_b = {c["number"]: c for c in clauses_b if c["number"] is not None}

    for a in clauses_a:
        match = by_number_b.get(a["number"]) if a["number"] is not None else None
        best_match = match["text"] if match else None

        if best_match is not None:
            translated_a = translate_known_terms(a["text"])
            translated_b = translate_known_terms(best_match)
            words_a = set(translated_a.split())
            words_b = set(translated_b.split())
            best_score = (
                len(words_a & words_b) / len(words_a | words_b)
                if words_a and words_b else 0.0
            )
        else:
            best_score = 0.0

        risky = is_risky(a["text"]) or (best_match and is_risky(best_match))
        dd_flags = flag_due_diligence_categories(a["text"] + " " + (best_match or ""))
        entities = extract_relevant_entities(a["text"])

        if best_score < 0.15 or best_match is None:
            label = "Clause missing"
            risk = "Risk increased" if (risky or dd_flags) else "Low risk"
            why = "This clause appears in Contract A but has no clear counterpart in Contract B."
            business_impact = "A required protection or obligation may not be covered — worth flagging for legal review."
        elif best_score < 0.55:
            label = "Clause changed"
            risk = "Risk increased" if (risky or dd_flags) else "Worth a look"
            why = "The wording differs enough that the meaning may have shifted."
            business_impact = "Terms may no longer match the approved standard — confirm the change was intentional."
        else:
            label = "Standard"
            risk = "Risk increased" if dd_flags else "Low risk"
            why = "This clause closely matches the reference."
            business_impact = "No action needed." if not dd_flags else "Matches reference, but touches a due-diligence category below — worth a quick look."

        results.append({
            "heading": a["heading"],
            "text_a": a["text"],
            "text_b": best_match or "— no matching clause found —",
            "label": label,
            "risk": risk,
            "why_this_matters": why,
            "business_impact": business_impact,
            "similarity": round(best_score, 2),
            "dd_flags": dd_flags,
            "entities": entities,
        })

    return results


# ---------------------------------------------------------------------------
# 7. PAGES
# ---------------------------------------------------------------------------

with st.sidebar:
    lang_choice = st.selectbox(
        "🌐 Language / Sprache",
        ["English", "German"],
        index=["English", "German"].index(st.session_state.ui_lang),
        format_func=lambda x: "English" if x == "English" else "Deutsch (Österreich)",
        key="ui_lang_select",
    )
    st.session_state.ui_lang = lang_choice

    st.markdown(
        f"<div style='background:{BRAND_YELLOW};padding:8px 12px;border-radius:6px;"
        f"margin-bottom:8px;display:flex;align-items:center;gap:10px;'>"
        f"{GIEBELKREUZ_SVG}"
        f"<h2 style='color:{BRAND_DARK};margin:0;font-size:1.15em;'>{t('sidebar_title')}</h2></div>",
        unsafe_allow_html=True,
    )
    st.caption(t("nav_caption"))
    page = st.radio(
        "Navigation",
        ["Library", "Chat", "Compare", "Image Q&A"],
        index=2,
        label_visibility="collapsed",
    )

if page == "Compare":
    st.markdown(
        f"<div style='display:flex;align-items:center;gap:14px;border-left:6px solid {BRAND_YELLOW};padding-left:14px;'>"
        f"{GIEBELKREUZ_SVG}"
        f"<div>"
        f"<h1 style='color:{BRAND_DARK}; margin-bottom:0;'>{t('page_title')}</h1>"
        f"<p style='color:#666666; font-size:1.1em; margin-top:0;'>{t('page_subtitle')}</p>"
        f"</div></div>",
        unsafe_allow_html=True,
    )
    st.caption(t("page_caption"))

    col_a, col_b, col_diff = st.columns([1, 1, 1.3])

    with col_a:
        st.subheader(t("contract_a_header"))
        mode_a = st.radio(t("input_label"), [t("upload_option"), t("paste_option")], key="mode_a", horizontal=True)
        if mode_a == t("upload_option"):
            file_a = st.file_uploader(t("upload_a"), type=["txt", "docx", "pdf"], key="file_a")
            text_a_input = None
        else:
            file_a = None
            text_a_input = st.text_area(t("paste_a"), height=300, key="text_a")

    with col_b:
        st.subheader(t("contract_b_header"))
        mode_b = st.radio(t("input_label"), [t("upload_option"), t("paste_option")], key="mode_b", horizontal=True)
        if mode_b == t("upload_option"):
            file_b = st.file_uploader(t("upload_b"), type=["txt", "docx", "pdf"], key="file_b")
            text_b_input = None
        else:
            file_b = None
            text_b_input = st.text_area(t("paste_b"), height=300, key="text_b")

    ready_a = bool(file_a) or bool(text_a_input)
    ready_b = bool(file_b) or bool(text_b_input)

    with col_diff:
        st.subheader(t("diff_header"))
        if not (ready_a and ready_b):
            st.info(t("provide_both"))
        else:
            with st.spinner(t("comparing_spinner")):
                text_a = extract_text(file_a) if file_a else text_a_input
                text_b = extract_text(file_b) if file_b else text_b_input

                lang_a = detect_language(text_a)
                lang_b = detect_language(text_b)

                clauses_a = split_into_clauses(text_a)
                clauses_b = split_into_clauses(text_b)

                results = compare_clauses(clauses_a, clauses_b)

            st.markdown(t("detected_as", a=lang_a, b=lang_b))

            high_risk_count = sum(1 for r in results if r["risk"] == "Risk increased")
            changed_count = sum(1 for r in results if r["label"] == "Clause changed")
            missing_count = sum(1 for r in results if r["label"] == "Clause missing")
            dd_count = sum(1 for r in results if r["dd_flags"])

            m1, m2, m3, m4 = st.columns(4)
            m1.metric(t("metric_risk"), high_risk_count)
            m2.metric(t("metric_changed"), changed_count)
            m3.metric(t("metric_missing"), missing_count)
            m4.metric(t("metric_dd"), dd_count)

            all_dd_flags = sorted({f for r in results for f in r["dd_flags"]})
            if all_dd_flags:
                st.markdown(t("dd_categories_header"))
                for cat in all_dd_flags:
                    st.markdown(f"- 🔎 **{cat}** — {DD_CATEGORIES[cat]['why']}")
                st.caption(t("dd_caption"))

            flagged = [r for r in results if r["label"] != "Standard" or r["dd_flags"]]
            standard = [r for r in results if r["label"] == "Standard" and not r["dd_flags"]]

            RISK_ORDER = {"Low risk": 0, "Worth a look": 1, "Risk increased": 2}
            flagged.sort(key=lambda r: RISK_ORDER.get(r["risk"], 1), reverse=True)

            RISK_STYLE = {
                "Low risk": ("#2E7D32", "🟢"),
                "Worth a look": ("#E8A33D", "🟠"),
                "Risk increased": ("#C0392B", "🔴"),
            }

            st.markdown(t("flagged_summary", n=len(flagged), total=len(results)))

            show_all_entities = st.checkbox(t("show_all_entities"), value=False)

            for r in flagged:
                badge_color, risk_icon = RISK_STYLE.get(r["risk"], ("#E8A33D", "🟠"))
                dd_suffix = f" · {' / '.join(r['dd_flags'])}" if r["dd_flags"] else ""
                with st.expander(f"{risk_icon} {r['heading']} — {r['label']} ({r['risk']}){dd_suffix}"):
                    st.markdown(
                        f"<span style='background:{badge_color};color:white;"
                        f"padding:2px 8px;border-radius:4px;font-size:0.85em;'>"
                        f"{r['risk']}</span>",
                        unsafe_allow_html=True,
                    )
                    highlighted_a = highlight_key_facts_html(r["text_a"][:400] + ("…" if len(r["text_a"]) > 400 else ""))
                    highlighted_b = highlight_key_facts_html(r["text_b"][:400] + ("…" if len(r["text_b"]) > 400 else ""))
                    st.markdown(t("contract_a_label"))
                    st.markdown(
                        f"<div style='line-height:1.7;font-size:0.95em;'>{highlighted_a}</div>",
                        unsafe_allow_html=True,
                    )
                    st.markdown(t("contract_b_label"))
                    st.markdown(
                        f"<div style='line-height:1.7;font-size:0.95em;'>{highlighted_b}</div>",
                        unsafe_allow_html=True,
                    )
                    st.markdown(f"{t('why_matters')} {r['why_this_matters']}")
                    st.markdown(f"{t('business_impact')} {r['business_impact']}")

                    if r["dd_flags"]:
                        st.markdown(f"{t('dd_categories_inline')} " + ", ".join(r["dd_flags"]))

                    entities_to_show = r["entities"] if not show_all_entities else extract_relevant_entities(r["text_a"] + " " + r["text_b"])
                    if entities_to_show:
                        st.markdown(t("extracted_entities"))
                        for e in entities_to_show:
                            st.markdown(f"- `{e['type']}`: {e['value']}")

            if standard:
                with st.expander(t("standard_clauses", n=len(standard))):
                    for r in standard:
                        st.write(f"- {r['heading']} ({t('similarity_label')}: {r['similarity']})")

elif page == "Library":
    header_col, btn_col = st.columns([5, 1])
    with header_col:
        st.subheader(t("library_header"))
        st.caption(t("library_caption"))
    with btn_col:
        st.write("")
        ingest_clicked = st.button(t("ingest_button"), key="ingest_btn")

    if ingest_clicked:
        st.session_state.show_uploader = True

    if st.session_state.get("show_uploader"):
        uploaded = st.file_uploader(
            t("ingest_upload_label"), type=["txt", "docx", "pdf"], key="library_upload"
        )
        if uploaded is not None and uploaded.name not in [f["name"] for f in st.session_state.library_files]:
            st.session_state.library_files.append(
                {"name": uploaded.name, "type": uploaded.type or "application/octet-stream"}
            )
            st.session_state.show_uploader = False
            st.rerun()

    if not st.session_state.library_files:
        st.info(t("no_documents"))
    else:
        header_row = st.columns([0.4, 5, 2])
        header_row[1].markdown(t("col_name"))
        header_row[2].markdown(t("col_type"))
        for i, f in enumerate(st.session_state.library_files):
            row = st.columns([0.4, 5, 2])
            row[0].checkbox("Select", key=f"lib_check_{i}", label_visibility="collapsed")
            row[1].markdown(f"📄 {f['name']}")
            row[2].markdown(f"`{f['type']}`")

elif page == "Chat":
    st.subheader(t("chat_header"))
    st.caption(t("chat_caption"))

    if not st.session_state.chat_history:
        st.markdown(
            "<div style='text-align:center; padding:60px 0;'>"
            f"<div style='font-size:1.3em; font-weight:600;'>{t('ask_anything')}</div>"
            f"<div style='color:#999; margin-top:6px;'>{t('ask_anything_sub')}</div>"
            "</div>",
            unsafe_allow_html=True,
        )

    for msg in st.session_state.chat_history:
        with st.chat_message(msg["role"]):
            st.write(msg["content"])

    prompt = st.chat_input(t("chat_input_placeholder"))
    if prompt:
        st.session_state.chat_history.append({"role": "user", "content": prompt})
        mock_reply = t("chat_mock_reply", prompt=prompt)
        st.session_state.chat_history.append({"role": "assistant", "content": mock_reply})
        st.rerun()

elif page == "Image Q&A":
    st.subheader(t("image_qa_header"))
    st.caption(t("image_qa_caption"))

    img_col, qa_col = st.columns([1, 1.5])
    with img_col:
        image_file = st.file_uploader(
            t("image_qa_upload_label"),
            type=["png", "jpg", "jpeg"],
            key="image_qa_upload",
        )
        if image_file:
            st.image(image_file, use_container_width=True)

    with qa_col:
        if not image_file:
            st.info(t("image_qa_info"))
        else:
            img_question = st.text_input(t("image_qa_question_label"), key="image_qa_question")
            if st.button(t("image_qa_ask_button"), key="image_qa_ask") and img_question:
                st.markdown(t("image_qa_mock_answer", q=img_question))