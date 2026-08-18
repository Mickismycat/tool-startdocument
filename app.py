import json
import re
import tempfile
from io import BytesIO
from datetime import date
from pathlib import Path
from typing import Any, Dict, List

import streamlit as st
from docx import Document
from openai import OpenAI
from pptx import Presentation
from pptx.util import Pt, Emu
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pypdf import PdfReader
from pptx.dml.color import RGBColor

APP_TITLE = "Startdocument Generator"
APP_DIR = Path(__file__).resolve().parent
TEMPLATE_CANDIDATES = [
    APP_DIR / "templates" / "Startdocument_Cooble_template.pptx",
    APP_DIR / "Startdocument_Cooble_template.pptx",
    Path("templates/Startdocument_Cooble_template.pptx"),
    Path("Startdocument_Cooble_template.pptx"),
]

def get_template_path() -> Path:
    for candidate in TEMPLATE_CANDIDATES:
        if candidate.exists():
            return candidate
    checked = ", ".join(str(c) for c in TEMPLATE_CANDIDATES)
    raise RuntimeError(
        "PowerPoint-template niet gevonden. Controleer of 'Startdocument_Cooble_template.pptx' "
        "in de map 'templates' staat of in de hoofdmap van de repository. Gecheckte paden: " + checked
    )
DEFAULT_MODEL = "gpt-4.1"


# v2.4: doelgroepresearch strikt los van vacaturecontext; pullfactoren taxonomy-locked op extern webonderzoek.
def generate_with_openai_pipeline(vacature: str, intake: str, linkedin_size: str, extra: str, status=None) -> Dict[str, Any]:
    if status:
        status.write("Stap 1/7: feiten uit vacature en intake halen")
    facts = call_openai_json(build_fact_extraction_prompt(vacature, intake, extra), use_web=False)

    fallback = extract_basis_fallback(vacature, intake)
    for fk in ["klantnaam", "vacaturenaam", "salaris"]:
        if is_empty_or_placeholder(facts.get(fk, "")) and fallback.get(fk):
            facts[fk] = fallback[fk]
    facts["salaris"] = normalize_salary_display(facts.get("salaris", ""))

    extracted_no_go = extract_no_go_companies_from_intake(intake + "\n" + extra)
    if extracted_no_go:
        merged = []
        for item in clean_list(facts.get("no_go_bedrijven", [])) + extracted_no_go:
            c = clean_company_name(item)
            if c and c not in merged:
                merged.append(c)
        facts["no_go_bedrijven"] = merged

    if status:
        status.write("Stap 2/8: doelgroep en concurrenten online onderzoeken")
    market = call_openai_json(build_target_market_research_prompt(facts, linkedin_size), use_web=True)

    if status:
        status.write("Stap 3/8: belangrijkste arbeidsvoorwaarden online onderzoeken")
    conditions = call_openai_json(build_employment_conditions_research_prompt(facts), use_web=True)

    if status:
        status.write("Stap 4/8: pullfactoren online onderzoeken")
    pull = call_openai_json(build_pullfactors_research_prompt(facts), use_web=True)
    if pullfactors_are_invalid(pull.get("pullfactoren", []), facts.get("klantnaam", "")):
        pull = call_openai_json(build_pullfactors_research_prompt(facts, strict_retry=True), use_web=True)
    if pullfactors_are_invalid(pull.get("pullfactoren", []), facts.get("klantnaam", "")):
        raise RuntimeError("Online onderzoek leverde geen geldige pullfactoren op. De tool stopt bewust in plaats van vacature-inhoud of arbeidsmarktkrapte als pullfactor te gebruiken.")

    if status:
        status.write("Stap 5/8: leeftijd en man-vrouwverhouding online onderzoeken")
    demographics = call_openai_json(build_demographics_research_prompt(facts), use_web=True)
    research = merge_research_parts(market, conditions, pull, demographics)

    if status:
        status.write("Stap 6/8: startdocument-content schrijven")
    data = call_openai_json(build_writer_prompt(facts, research, vacature, intake, linkedin_size, extra), use_web=False)

    if status:
        status.write("Stap 7/8: presentatiekwaliteit aanscherpen")
    try:
        data = call_openai_json(build_presentation_prompt(data, facts, research), use_web=False)
    except Exception as presentation_error:
        data.setdefault("kwaliteitscontrole", {}).setdefault("waarschuwingen", []).append(str(presentation_error))

    if status:
        status.write("Stap 8/8: business rules toepassen")
    data = apply_business_rules(data, intake + "\n" + extra, linkedin_size, vacature, extra)
    # Researchvelden zijn leidend voor externe marktdata.
    data.setdefault("doelgroepanalyse", {})["pullfactoren"] = presentation_bullets(normalize_pullfactors(research.get("pullfactoren", [])), 3)
    data.setdefault("voorwaarden", {})["belangrijkste_arbeidsvoorwaarden"] = presentation_bullets(normalize_conditions(research.get("belangrijkste_arbeidsvoorwaarden", [])), 3)
    stable_demo = deterministic_demographics(facts, research)
    data.setdefault("doelgroepanalyse", {})["geslacht"] = stable_demo.get("geslacht", {"man": "", "vrouw": ""})
    data.setdefault("doelgroepanalyse", {})["leeftijdsverdeling"] = stable_demo.get("leeftijdsverdeling", normalize_age_distribution(research.get("leeftijdsverdeling", [])))
    data.setdefault("basisgegevens", {})["salaris"] = normalize_salary_display(data.get("basisgegevens", {}).get("salaris", ""))
    data.setdefault("kwaliteitscontrole", {})["pipeline"] = "v2.4: facts -> occupation-only web market -> web conditions -> taxonomy-locked web pullfactors -> web demographics -> writer -> template-only fill"
    return data


st.set_page_config(page_title=APP_TITLE, page_icon="📄", layout="wide")


def read_docx(file) -> str:
    doc = Document(file)
    paragraphs = [p.text for p in doc.paragraphs if p.text.strip()]
    table_text = []
    for table in doc.tables:
        for row in table.rows:
            cells = [cell.text.strip() for cell in row.cells if cell.text.strip()]
            if cells:
                table_text.append(" | ".join(cells))
    return "\n".join(paragraphs + table_text)


def read_pdf(file) -> str:
    reader = PdfReader(file)
    pages = []
    for page in reader.pages:
        text = page.extract_text() or ""
        if text.strip():
            pages.append(text)
    return "\n".join(pages)


def read_uploaded_file(file) -> str:
    if file is None:
        return ""
    name = file.name.lower()
    try:
        if name.endswith(".docx"):
            return read_docx(file)
        if name.endswith(".pdf"):
            return read_pdf(file)
        if name.endswith(".txt"):
            return file.read().decode("utf-8", errors="ignore")
    except Exception as exc:
        raise RuntimeError(f"Kon bestand '{file.name}' niet uitlezen: {exc}")
    return ""


def clean_list(items: List[str]) -> List[str]:
    return [str(x).strip(" •-\n\t") for x in (items or []) if str(x).strip(" •-\n\t")]


def split_one_topic_per_bullet(items: List[str]) -> List[str]:
    """Zorgt dat samengestelde bullets worden opgeknipt naar één onderwerp per bullet."""
    result: List[str] = []
    for raw in clean_list(items):
        text = raw.strip()
        # Splits alleen op duidelijke opsommingen, niet op woorden als "Learning & Development".
        parts = re.split(r"\s*(?:;|/|\+|,\s*(?=[A-ZÀ-ÖØ-Þ]))\s*", text)
        expanded: List[str] = []
        for part in parts:
            # Splits op " en " alleen als beide delen kort genoeg zijn om losse thema's te zijn.
            sub = re.split(r"\s+en\s+", part)
            if len(sub) == 2 and all(1 <= len(x.split()) <= 4 for x in sub):
                expanded.extend(sub)
            else:
                expanded.append(part)
        for part in expanded:
            part = part.strip(" •-\n\t")
            if part and part not in result:
                result.append(part)
    return result




def first_nonempty_line(text: str) -> str:
    for line in (text or "").splitlines():
        line = line.strip()
        if line and len(line) < 90:
            return line
    return ""


def extract_basis_fallback(vacature_text: str, intake_text: str) -> Dict[str, str]:
    """Deterministische fallback voor klantnaam, vacaturenaam en salaris uit input."""
    combined = f"{intake_text}\n{vacature_text}"
    fallback: Dict[str, str] = {}

    # Klantnaam uit intakevelden of herkenbare vacaturetekst.
    m = re.search(r"(?im)^\s*Klant\s*:\s*(.+)$", combined)
    if m:
        fallback["klantnaam"] = m.group(1).strip()
    else:
        m = re.search(r"(?i)\bWij zijn\s+([A-Z][A-Za-z0-9&+ .'-]{2,40})", vacature_text)
        if m:
            name = re.split(r"[\n\.;,]", m.group(1).strip())[0].strip()
            fallback["klantnaam"] = name

    # Vacaturenaam uit intakeveld; als leeg, eerste regel vacaturetekst.
    m = re.search(r"(?im)^\s*Vacature\s*:\s*(.+)$", combined)
    if m and m.group(1).strip():
        fallback["vacaturenaam"] = m.group(1).strip()
    else:
        title = first_nonempty_line(vacature_text)
        # Vermijd plaats/metadata als titel.
        if title and not re.search(r"\b(remote|hybrid|nl|ov|sp|locatie)\b", title, re.I):
            fallback["vacaturenaam"] = title

    # Salaris: euro-range, salarisschaal/schaal of expliciete salarisregel.
    salary_patterns = [
        r"(?i)(?:salaris(?:range)?|salarisindicatie)\s*[:\-]?\s*([^\n]{3,80})",
        r"(?i)\b(schaal\s*[0-9][0-9A-Za-z/\- ]{0,30})",
        r"(?i)(€\s?[\d\.,]+\s*(?:-|–|tot)\s*€?\s?[\d\.,]+)",
        r"(?i)([\d\.,]+\s*(?:-|–|tot)\s*[\d\.,]+\s*(?:euro|bruto|per maand)?)",
    ]
    for pat in salary_patterns:
        m = re.search(pat, combined)
        if m:
            val = m.group(1).strip(" .;,")
            if val and not re.search(r"weten we niet|onbekend", val, re.I):
                fallback["salaris"] = val
                break
            # Als er staat: salaris weten we niet, schaal 8/9/10, pak alsnog schaal.
            scale = re.search(r"(?i)\b(schaal\s*[0-9][0-9A-Za-z/\- ]{0,30})", m.group(0))
            if scale:
                fallback["salaris"] = scale.group(1).strip()
                break
    return fallback


def is_empty_or_placeholder(value: str) -> bool:
    value = str(value or "").strip()
    return value == "" or value.lower() in {"onbekend", "n.v.t.", "nvt", "in overleg", "-", "..."}


def is_placeholder_company(value: str) -> bool:
    return bool(re.search(r"(?i)^bedrijf\s+[a-z0-9]$|^concurrent\s+[a-z0-9]$|^organisatie\s+[a-z0-9]$", str(value or "").strip()))


def infer_competitors_offline(vacature_text: str, intake_text: str, current: List[str]) -> List[str]:
    """Fallback wanneer de AI placeholders geeft. Beperkt, maar beter dan Bedrijf A/B/C."""
    text = f"{vacature_text}\n{intake_text}".lower()
    companies = []
    # Bedrijven die expliciet in intake staan, mogen ook als concurrent/check-eerst zichtbaar worden.
    companies.extend(extract_no_go_companies_from_intake(intake_text))
    maps = [
        (["waterkwaliteit", "afvalwater", "waterwet", "waterschap", "watermanagement"], ["Witteveen+Bos", "Royal HaskoningDHV", "Sweco", "Antea Group", "Arcadis", "TAUW"]),
        (["luchtkwaliteit", "milieuconsultant", "emissie", "vergunning", "omgevingswet"], ["Royal HaskoningDHV", "Witteveen+Bos", "Sweco", "Antea Group", "Arcadis", "TAUW"]),
        (["business analist", "informatieanalist", "product owner"], ["Sogeti", "Capgemini", "Ordina", "CGI", "Conclusion", "Atos"]),
        (["lead engineer", "engineer", "warmtenet", "ondergrondse infra"], ["BAM", "Heijmans", "VolkerWessels", "Strukton", "Equans", "SPIE"]),
        (["accountmanager", "retentie", "upsell", "sales manager"], ["LeasePlan", "Alphabet", "Arval", "Athlon", "ALD Automotive"]),
    ]
    for keywords, names in maps:
        if any(k in text for k in keywords):
            companies.extend(names)
            break
    for item in current or []:
        if item and not is_placeholder_company(item):
            companies.append(item)
    result = []
    for c in companies:
        c = clean_company_name(c)
        if c and not is_placeholder_company(c) and c not in result:
            result.append(c)
    return result[:8]


def normalize_condition_label(text: str) -> str:
    """Normaliseer onderzoeksresultaten naar één generieke arbeidsvoorwaardencategorie.

    Concrete aantallen/bedragen uit werkgeversvacatures mogen nooit in de slide terechtkomen.
    """
    text = str(text or "").strip(" •-\n\t")
    low = text.lower()
    category_rules = [
        (("salaris", "loon", "beloning", "pay", "salary"), "Salaris"),
        (("pensioen",), "Pensioenregeling"),
        (("vakantie", "verlof", "adv", "vrije dagen"), "Vakantiedagen"),
        (("hybride", "thuiswerk", "remote", "flexibel werken"), "Hybride werken"),
        (("opleiding", "training", "ontwikkeling", "studiebudget", "leerbudget"), "Ontwikkelmogelijkheden"),
        (("mobiliteit", "leaseauto", "auto van de zaak", "reiskosten", "ov-vergoeding", "fietsregeling"), "Mobiliteit"),
        (("bonus", "variabele beloning"), "Bonusregeling"),
        (("eindejaarsuitkering", "13e maand", "dertiende maand"), "Eindejaarsuitkering"),
        (("werktijd", "werkweek", "uren", "rooster"), "Flexibele werktijden"),
        (("vitaliteit", "fitness", "gezondheid"), "Vitaliteitsregeling"),
    ]
    for needles, label in category_rules:
        if any(n in low for n in needles):
            return label
    # Geen cijfers/percentages/bedragen toelaten in resterende labels.
    text = re.sub(r"€?\s*\d[\d\.,%/-]*", "", text)
    text = re.sub(r"\s+", " ", text).strip(" ,;-:")
    # Houd labels kort; concrete zinnen zijn niet gewenst.
    if len(text.split()) > 4:
        text = " ".join(text.split()[:4])
    return text[:1].upper() + text[1:] if text else text


def normalize_conditions(items: List[str]) -> List[str]:
    out: List[str] = []
    for item in split_one_topic_per_bullet(items):
        item = normalize_condition_label(item)
        if item and item not in out:
            out.append(item)
    return out[:3]


def normalize_pullfactor_label(text: str) -> str:
    """Pullfactoren mogen in v2.4 alleen nog uit de vaste, neutrale onderzoekstaxonomie komen."""
    text = re.sub(r"\s+", " ", str(text or "").strip(" •-\n\t.,;:"))
    if not text:
        return ""
    # Kleine normalisatie voor spelling/variant, maar géén inhoudelijke herinterpretatie uit vacaturecontext.
    aliases = {
        "werk prive balans": "Werk-privébalans",
        "werk-privé balans": "Werk-privébalans",
        "work-life balance": "Werk-privébalans",
        "professionele groei": "Professionele ontwikkeling",
        "leren en ontwikkelen": "Professionele ontwikkeling",
        "doorgroei": "Doorgroeimogelijkheden",
        "vrijheid": "Autonomie",
        "waardering": "Waardering en erkenning",
        "erkenning": "Waardering en erkenning",
        "zekerheid": "Baanzekerheid",
        "leiderschap": "Goed leiderschap",
        "werksfeer": "Prettige werksfeer",
    }
    if text in ALLOWED_PULLFACTORS:
        return text
    mapped = aliases.get(text.lower())
    return mapped if mapped in ALLOWED_PULLFACTORS else ""


def normalize_pullfactors(items: List[str]) -> List[str]:
    out: List[str] = []
    for item in clean_list(items):
        label = normalize_pullfactor_label(item)
        if label and label not in out:
            out.append(label)
    # Geen fallback vanuit vacature of modelkennis: research moet drie geldige factoren opleveren.
    return out[:3]

def build_target_market_research_prompt(facts: Dict[str, Any], linkedin_size: str) -> str:
    doelgroep = public_occupation_query(facts)
    return f"""
Je bent recruitment researcher voor de Nederlandse arbeidsmarkt. Doe extern internetonderzoek naar deze doelgroep:
Doelgroep: {doelgroep}
Land: Nederland
LinkedIn-doelgroepgrootte indien handmatig aangeleverd: {linkedin_size}

Zoek online naar:
1. passende functietitels en senioriteit;
2. bedrijven waar deze doelgroep werkt;
3. doelgroepgrootte als LinkedIn-veld ontbreekt.

Harde regels:
- Gebruik geen vacaturetekst, intake, klantnaam of werkgevercontext als bron voor doelgroepinformatie.
- Concurrentenanalyse is altijd op bedrijfsniveau.
- Geef echte bedrijfsnamen. Nooit Bedrijf A/B/C, Concurrent 1 of Organisatie X.
- Doelgroepomschrijving moet specifiek zijn voor functiefamilie, domein, senioriteit en sector.
- Pullfactoren en arbeidsvoorwaarden laat je leeg; die worden in aparte researchmodules onderzocht.

Geef uitsluitend JSON terug:
{{
  "doelgroep_titel": "",
  "doelgroep_omschrijving": "",
  "verwachte_doelgroepgrootte": "",
  "belangrijkste_functietitels": [],
  "concurrenten_bedrijven": [],
  "zoekrichting": [],
  "pullfactoren": [],
  "belangrijkste_arbeidsvoorwaarden": [],
  "geslacht": {{"man": "", "vrouw": ""}},
  "leeftijdsverdeling": [],
  "research_bronnen": [],
  "research_toelichting": ""
}}
""".strip()


def build_employment_conditions_research_prompt(facts: Dict[str, Any]) -> str:
    doelgroep = public_occupation_query(facts)
    return f"""
Je bent arbeidsmarktonderzoeker. Doe ACTUEEL INTERNETONDERZOEK naar deze Nederlandse doelgroep:
Doelgroep: {doelgroep}
Land: Nederland

Onderzoek uitsluitend deze vraag:
Welke 3 arbeidsvoorwaarden vinden mensen in deze beroepsgroep in het algemeen belangrijk wanneer zij ergens in dienst zijn?

Regels:
- Gebruik verplicht web_search en externe arbeidsmarktbronnen zoals werknemersenquêtes, brancheonderzoeken, CNV/FNV/Randstad/Indeed/Nationale Vacaturebank/sectoronderzoek.
- Gebruik geen vacaturetekst, intake, werkgever of vacaturepagina als bron.
- Geef generieke categorieën, niet de concrete invulling van één werkgever.
- Voorbeelden: Salaris, Pensioenregeling, Vakantiedagen, Hybride werken, Flexibele werktijden, Mobiliteit, Ontwikkelmogelijkheden, Eindejaarsuitkering.
- Geen bedragen, percentages, aantallen dagen, uren of concrete werkgeversvoorwaarden.
- Eén onderwerp per item.

Geef uitsluitend JSON:
{{
  "belangrijkste_arbeidsvoorwaarden": ["", "", ""],
  "bronnen": [],
  "toelichting": ""
}}
""".strip()


def build_demographics_research_prompt(facts: Dict[str, Any]) -> str:
    doelgroep = public_occupation_query(facts)
    return f"""
Je bent arbeidsmarktonderzoeker. Doe ACTUEEL INTERNETONDERZOEK naar de demografie van deze Nederlandse beroepsdoelgroep:
Doelgroep: {doelgroep}
Land: Nederland

Onderzoek:
1. man-vrouwverhouding binnen deze beroepsgroep of dichtstbijzijnde officiële functiefamilie;
2. leeftijdsverdeling binnen dezelfde beroepsgroep/functiefamilie.

Bronhiërarchie:
1. CBS / StatLine of andere officiële Nederlandse statistiek;
2. UWV, ROA, SBB of officiële branche-/beroepsorganisaties;
3. gerenommeerde Nederlandse arbeidsmarkt- of sectoronderzoeken.

Regels:
- Gebruik geen vacaturetekst, intake of werkgevercontext.
- Baseer man-vrouw en leeftijd op dezelfde beroepsafbakening waar mogelijk.
- Rond af op hele procenten.
- Gebruik altijd deze categorieën: 15-24, 25-34, 35-49, 50+.

Geef uitsluitend JSON:
{{
  "geslacht": {{"man": "", "vrouw": ""}},
  "leeftijdsverdeling": ["15-24: %", "25-34: %", "35-49: %", "50+: %"],
  "bronnen": [],
  "afbakening": "",
  "toelichting": ""
}}
""".strip()



# -----------------------------------------------------------------------------
# v2.3: juiste Cooble-template is leidend. Alleen placeholders worden ingevuld.
# Geen lettergrootte, positie, marges, regelafstand of static layout aanpassen.
# -----------------------------------------------------------------------------

def _pct_int(value: Any):
    m = re.search(r"(\d{1,3})(?:[.,]\d+)?\s*%?", str(value or ""))
    if not m:
        return None
    return max(0, min(100, int(round(float(m.group(1))))))


def _round_to_5(value: int) -> int:
    return max(0, min(100, int(round(value / 5.0) * 5)))


def _normalize_gender_web(gender: Dict[str, Any]) -> Dict[str, str]:
    """Normaliseer webresearch naar stabiele afgeronde percentages; verzin geen benchmark."""
    man = _pct_int((gender or {}).get("man", ""))
    vrouw = _pct_int((gender or {}).get("vrouw", ""))
    if man is None and vrouw is None:
        return {"man": "", "vrouw": ""}
    if man is None:
        man = 100 - vrouw
    if vrouw is None:
        vrouw = 100 - man
    # Kleine bron-/afrondingsverschillen worden gestabiliseerd naar stappen van 5%.
    man = _round_to_5(man)
    vrouw = 100 - man
    return {"man": f"{man}%", "vrouw": f"{vrouw}%"}


def _normalize_age_web(items: List[str]) -> List[str]:
    wanted = ["15-24", "25-34", "35-49", "50+"]
    parsed: Dict[str, int] = {}
    for raw in clean_list(items):
        compact = str(raw).replace(" ", "")
        pct = _pct_int(raw)
        if pct is None:
            continue
        for label in wanted:
            if label in compact:
                parsed[label] = pct
                break
    if len(parsed) != 4:
        return []
    rounded = {k: _round_to_5(v) for k, v in parsed.items()}
    diff = 100 - sum(rounded.values())
    # Corrigeer alleen het afrondingsverschil op de grootste categorie.
    largest = max(rounded, key=rounded.get)
    rounded[largest] = max(0, rounded[largest] + diff)
    return [f"{k}: {rounded[k]}%" for k in wanted]


def build_demographics_research_prompt(facts: Dict[str, Any], strict_retry: bool = False) -> str:
    doelgroep = public_occupation_query(facts)
    retry = "" if not strict_retry else """
DIT IS EEN HERHAALPOGING OMDAT DE EERSTE DATA ONVOLLEDIG WAS.
Je MOET vier leeftijdspercentages en een man/vrouwverdeling teruggeven.
Gebruik desnoods één niveau bredere officiële beroepsklasse, maar blijf bij dezelfde afbakening voor beide analyses.
"""
    return f"""
Je bent een arbeidsmarktonderzoeker. Doe ACTUEEL WEBONDERZOEK naar uitsluitend de DEMOGRAFIE van deze Nederlandse beroepsdoelgroep:
Doelgroep: {doelgroep}
Land: Nederland

BELANGRIJK:
- Je krijgt bewust GEEN werkgever, vacaturetekst, intake, bedrijfsnaam of concrete vacaturecontext.
- Gebruik die informatie dus ook niet in je onderzoek.
- Onderzoek de beroepsgroep/functiefamilie als geheel.

Zoek twee dingen:
1. man-vrouwverhouding;
2. leeftijdsverdeling.

VASTE BRONVOLGORDE:
1. CBS / StatLine;
2. UWV, ROA, SBB of een officiële branche-/beroepsorganisatie;
3. pas daarna een gerenommeerde Nederlandse arbeidsmarktbron.

CONSISTENTIEREGELS:
- Gebruik voor man/vrouw en leeftijd dezelfde beroepsafbakening waar mogelijk.
- Als exacte functiedata niet bestaan, kies één aantoonbaar dichtstbijzijnde officiële beroepsklasse en gebruik die consequent.
- Gebruik de meest recente Nederlandse cijfers die je kunt vinden.
- Maak geen vrije AI-schatting.
- Man + vrouw = exact 100%.
- Leeftijd = exact deze vier categorieën en samen 100%: 15-24, 25-34, 35-49, 50+.
- Rond de gevonden cijfers af op hele procenten; de applicatie normaliseert daarna voor presentatiestabiliteit.
- Geef minimaal één concrete bron/domein op.
{retry}

Geef uitsluitend JSON:
{{
  "geslacht": {{"man": "", "vrouw": ""}},
  "leeftijdsverdeling": ["15-24: 0%", "25-34: 0%", "35-49: 0%", "50+: 0%"],
  "bronnen": [],
  "afbakening": "",
  "toelichting": ""
}}
""".strip()


def deterministic_demographics(facts: Dict[str, Any], research: Dict[str, Any]) -> Dict[str, Any]:
    """v2.3: uitsluitend webdata; geen zelfverzonnen vaste functietabellen."""
    gender = _normalize_gender_web((research or {}).get("geslacht", {}))
    age = _normalize_age_web((research or {}).get("leeftijdsverdeling", []))
    if not gender.get("man") or not gender.get("vrouw") or len(age) != 4:
        retry = call_openai_json(build_demographics_research_prompt(facts, strict_retry=True), use_web=True)
        gender = _normalize_gender_web(retry.get("geslacht", {}))
        age = _normalize_age_web(retry.get("leeftijdsverdeling", []))
    if not gender.get("man") or not gender.get("vrouw") or len(age) != 4:
        raise RuntimeError(
            "Demografisch webonderzoek leverde geen complete man/vrouw- en leeftijdsverdeling op. "
            "De tool stopt bewust in plaats van percentages te verzinnen."
        )
    return {
        "geslacht": gender,
        "leeftijdsverdeling": age,
        "afbakening": (research or {}).get("demografie_afbakening", "")
    }


def _strip_target_size(value: Any) -> str:
    text = str(value or "").strip()
    text = re.sub(r"^[±+\-\s]+", "", text)
    return text


def _list3(items: Any) -> List[str]:
    vals = clean_list(items if isinstance(items, list) else [])
    return (vals + ["", "", ""])[:3]


def _replace_tokens_in_text_frame(shape, replacements: Dict[str, str]) -> None:
    """Vervang tekst in bestaande runs. Geen clear(), geen nieuwe fonts, geen nieuwe paragrafen."""
    if not (hasattr(shape, "text_frame") and shape.has_text_frame):
        return
    for paragraph in shape.text_frame.paragraphs:
        for run in paragraph.runs:
            txt = run.text
            if not txt or "{{" not in txt:
                continue
            for token, value in replacements.items():
                if token in txt:
                    txt = txt.replace(token, str(value or ""))
            run.text = txt


def _render_template_shape_v23(slide, shape, data: Dict[str, Any], replacements: Dict[str, str]) -> None:
    if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
        for child in shape.shapes:
            _render_template_shape_v23(slide, child, data, replacements)
        return
    if not (hasattr(shape, "text_frame") and shape.has_text_frame):
        return
    original = full_text(shape)
    if "{{leeftijdsverdeling}}" in original:
        left, top, width, height = shape.left, shape.top, shape.width, shape.height
        # De placeholder is alleen een anker; hij wordt onzichtbaar gemaakt.
        for p in shape.text_frame.paragraphs:
            for r in p.runs:
                r.text = ""
        img = create_age_chart_image(get_nested_v12(data, "doelgroepanalyse.leeftijdsverdeling", []))
        slide.shapes.add_picture(img, left, top, width=width, height=height)
        return
    _replace_tokens_in_text_frame(shape, replacements)


def generate_pptx(data: Dict[str, Any]) -> bytes:
    """v2.3: vul uitsluitend placeholders in het goedgekeurde Cooble-template."""
    template_path = get_template_path()
    prs = Presentation(str(template_path))

    tasks = _list3(get_nested_v12(data, "functieprofiel.taken_verantwoordelijkheden", []))
    eisen = _list3(get_nested_v12(data, "kandidaatprofiel.eisen", []))
    voorkeuren = _list3(get_nested_v12(data, "kandidaatprofiel.voorkeuren", []))
    usps = _list3(get_nested_v12(data, "functieprofiel.usp_functie", []))
    pulls = _list3(normalize_pullfactors(get_nested_v12(data, "doelgroepanalyse.pullfactoren", [])))
    voorwaarden = _list3(normalize_conditions(get_nested_v12(data, "voorwaarden.belangrijkste_arbeidsvoorwaarden", [])))
    afspraken = (clean_list(data.get("afspraken", [])) + ["", "", ""])[:3]
    nogo = clean_list(get_nested_v12(data, "kandidaatprofiel.no_go_sourcing", []))

    replacements = {
        "{{vacaturenaam}}": get_nested_v12(data, "basisgegevens.vacaturenaam", ""),
        "{{datum}}": get_nested_v12(data, "basisgegevens.datum", "") or date.today().strftime("%d-%m-%Y"),
        "{{intake_samenvatting}}": presentation_summary(data.get("intake_samenvatting", "")),
        "{{taak_1}}": tasks[0], "{{taak_2}}": tasks[1], "{{taak_3}}": tasks[2],
        "{{eis_1}}": eisen[0], "{{eis_2}}": eisen[1], "{{eis_3}}": eisen[2],
        "{{voorkeur_1}}": voorkeuren[0], "{{voorkeur_2}}": voorkeuren[1], "{{voorkeur_3}}": voorkeuren[2],
        "{{usp_1}}": usps[0], "{{usp_2}}": usps[1], "{{usp_3}}": usps[2],
        "{{no_go_sourcing}}": ", ".join(nogo),
        "{{salaris}}": normalize_salary_display(get_nested_v12(data, "basisgegevens.salaris", "")),
        "{{locatie}}": get_nested_v12(data, "basisgegevens.locatie", ""),
        "{{uren}}": get_nested_v12(data, "basisgegevens.uren", ""),
        "{{doelgroepgrootte}}": _strip_target_size(get_nested_v12(data, "doelgroepanalyse.verwachte_doelgroepgrootte", "")),
        "{{pull_1}}": pulls[0], "{{pull_2}}": pulls[1], "{{pull_3}}": pulls[2],
        "{{voorwaarde_1}}": voorwaarden[0], "{{voorwaarde_2}}": voorwaarden[1], "{{voorwaarde_3}}": voorwaarden[2],
        "{{geslacht_man}}": get_nested_v12(data, "doelgroepanalyse.geslacht.man", ""),
        "{{geslacht_vrouw}}": get_nested_v12(data, "doelgroepanalyse.geslacht.vrouw", ""),
        "{{afspraak_1}}": afspraken[0], "{{afspraak_2}}": afspraken[1], "{{afspraak_3}}": afspraken[2],
    }

    for slide in prs.slides:
        for shape in list(slide.shapes):
            _render_template_shape_v23(slide, shape, data, replacements)

    # Safety check: geen zichtbare placeholdertokens in het eindbestand.
    leftovers = []
    def collect(shapes):
        for sh in shapes:
            if sh.shape_type == MSO_SHAPE_TYPE.GROUP:
                collect(sh.shapes)
            elif hasattr(sh, "text") and "{{" in sh.text:
                leftovers.append(sh.text)
    for slide in prs.slides:
        collect(slide.shapes)
    if leftovers:
        raise RuntimeError("Niet alle PowerPoint-placeholders konden worden ingevuld: " + " | ".join(leftovers[:5]))

    with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as tmp:
        prs.save(tmp.name)
        result = Path(tmp.name).read_bytes()
    return result

st.title("📄 Startdocument Generator")
st.caption("Upload de vacature en intake. Controleer de preview en download daarna een nette PowerPoint in de vaste Cooble-template.")

with st.sidebar:
    st.header("Status")
    mode = st.radio("Modus", ["AI-generatie", "Testmodus zonder API-key"], index=0)
    st.caption("Testmodus gebruikt voorbeelddata en controleert of de app en PowerPoint-export werken.")
    if st.secrets.get("OPENAI_API_KEY", ""):
        st.success("API-key gevonden")
    else:
        st.warning("Geen API-key gevonden")

st.subheader("1. Input")
col1, col2 = st.columns(2)
with col1:
    st.markdown("**Vacature**")
    vacature_file = st.file_uploader("Vacature uploaden", type=["docx", "pdf", "txt"], key="vac_file")
    vacature_paste = st.text_area("Of plak vacaturetekst", height=240)
with col2:
    st.markdown("**Intake**")
    intake_file = st.file_uploader("Intake uploaden", type=["docx", "pdf", "txt"], key="intake_file")
    intake_paste = st.text_area("Of plak intake-notities", height=240)

c1, c2 = st.columns([1, 2])
with c1:
    linkedin_size = st.text_input("Doelgroepgrootte gevonden op LinkedIn", placeholder="bijv. ± 500")
with c2:
    extra_notes = st.text_area("Extra opmerkingen", placeholder="bijv. salaris niet benoemen / extra compact schrijven / regio belangrijk", height=80)

if "data" not in st.session_state:
    st.session_state.data = None

if st.button("Genereer analyse", type="primary"):
    try:
        vacature_text = (read_uploaded_file(vacature_file) + "\n" + vacature_paste).strip()
        intake_text = (read_uploaded_file(intake_file) + "\n" + intake_paste).strip()
        if mode == "AI-generatie" and not vacature_text:
            st.error("Voeg minimaal een vacaturetekst toe.")
        elif mode == "AI-generatie" and not intake_text:
            st.error("Voeg minimaal intake-informatie toe.")
        else:
            with st.status("Analyse wordt gemaakt...", expanded=True) as status:
                st.write("Vacature en intake uitlezen")
                if mode == "Testmodus zonder API-key":
                    st.write("Testdata laden")
                    data = demo_data()
                else:
                    st.write("AI-pipeline starten")
                    data = generate_with_openai_pipeline(vacature_text, intake_text, linkedin_size, extra_notes, status=status)
                data = apply_business_rules(data, intake_text + "\n" + extra_notes, linkedin_size, vacature_text, extra_notes)
                st.session_state.data = data
                status.update(label="Analyse klaar", state="complete", expanded=False)
            st.success("Analyse klaar. Controleer en pas eventueel aan.")
    except Exception as e:
        st.exception(e)

if st.session_state.data:
    data = ensure_core_keys(st.session_state.data)
    st.subheader("2. Preview & aanpassen")
    tabs = st.tabs(["Basis", "Functie", "Doelgroep", "Sourcing", "Afspraken"])

    with tabs[0]:
        b = data.setdefault("basisgegevens", {})
        c1, c2, c3 = st.columns(3)
        b["klantnaam"] = c1.text_input("Klantnaam", b.get("klantnaam", ""))
        b["vacaturenaam"] = c2.text_input("Vacaturenaam", b.get("vacaturenaam", ""))
        b["datum"] = date.today().strftime("%d-%m-%Y")
        c3.text_input("Datum", b["datum"], disabled=True, help="Altijd de generatiedatum van vandaag.")
        c4, c5, c6 = st.columns(3)
        b["locatie"] = c4.text_input("Locatie", b.get("locatie", ""))
        b["uren"] = c5.text_input("Uren", b.get("uren", ""))
        b["salaris"] = c6.text_input("Salaris", b.get("salaris", ""))
        data["intake_samenvatting"] = st.text_area("Intake / vacaturesamenvatting", data.get("intake_samenvatting", ""), height=170)

    with tabs[1]:
        f = data.setdefault("functieprofiel", {})
        k = data.setdefault("kandidaatprofiel", {})
        f["taken_verantwoordelijkheden"] = editable_list("Taken & verantwoordelijkheden", f.get("taken_verantwoordelijkheden", []), "taken", 3, hard_max=True)
        k["eisen"] = editable_list("Eisen", k.get("eisen", []), "eisen", 3, hard_max=True)
        k["voorkeuren"] = editable_list("Voorkeuren", k.get("voorkeuren", []), "voorkeuren", 3, hard_max=True)
        f["usp_functie"] = editable_list("USP's van de functie", f.get("usp_functie", []), "usp", 3, hard_max=True)
        k["no_go_sourcing"] = editable_list("No-go sourcing", k.get("no_go_sourcing", []), "nogo", 5)

    with tabs[2]:
        d = data.setdefault("doelgroepanalyse", {})
        v = data.setdefault("voorwaarden", {})
        c1, c2 = st.columns(2)
        d["verwachte_doelgroepgrootte"] = c1.text_input("Doelgroepgrootte", d.get("verwachte_doelgroepgrootte", ""))
        d["regio"] = c2.text_input("Regio", d.get("regio", "Nederland"))
        d["pullfactoren"] = editable_list("Pullfactoren", d.get("pullfactoren", []), "pull", 3, hard_max=True)
        v["belangrijkste_arbeidsvoorwaarden"] = editable_list("Belangrijkste arbeidsvoorwaarden", v.get("belangrijkste_arbeidsvoorwaarden", []), "av", 3, hard_max=True)
        g = d.setdefault("geslacht", {})
        c3, c4 = st.columns(2)
        g["man"] = c3.text_input("Geslacht man", g.get("man", ""))
        g["vrouw"] = c4.text_input("Geslacht vrouw", g.get("vrouw", ""))
        d["leeftijdsverdeling"] = editable_list("Leeftijdsverdeling", d.get("leeftijdsverdeling", []), "age", 5)

    with tabs[3]:
        s = data.setdefault("sourcingplan", {})
        c = data.setdefault("concurrentenanalyse", {})
        s["doelgroep"] = st.text_area("Doelgroep", s.get("doelgroep", ""), height=100)
        s["strategie"] = st.text_area("Strategie", s.get("strategie", ""), height=120)
        s["zoekrichting"] = editable_list("Zoekrichting", s.get("zoekrichting", []), "zoek", 5)
        c["bedrijven"] = editable_list("Concurrenten / bedrijven", c.get("bedrijven", []), "conc", 8)
        c["toelichting"] = st.text_area("Toelichting concurrentenanalyse", c.get("toelichting", ""), height=90)

    with tabs[4]:
        data["afspraken"] = editable_list("Afspraken", data.get("afspraken", []), "afspraken", 5)

    st.session_state.data = data

    st.subheader("3. PowerPoint maken")
    try:
        data = apply_business_rules(data, "", linkedin_size)
        st.session_state.data = data
        pptx_bytes = generate_pptx(data)
        klant = get_nested(data, "basisgegevens.klantnaam", "klant") or "klant"
        vacature = get_nested(data, "basisgegevens.vacaturenaam", "vacature") or "vacature"
        filename = f"Startdocument_{klant}_{vacature}.pptx"
        filename = re.sub(r"[^A-Za-z0-9_\-\.]+", "_", filename)
        st.download_button(
            "Maak & download PowerPoint",
            pptx_bytes,
            filename,
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            type="primary",
        )
    except Exception as e:
        st.exception(e)
