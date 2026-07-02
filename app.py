import json
import re
import tempfile
from datetime import date
from pathlib import Path
from typing import Any, Dict, List

import streamlit as st
from docx import Document
from openai import OpenAI
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pypdf import PdfReader

APP_TITLE = "Startdocument Generator"
TEMPLATE_PATH = Path("templates/Startdocument_Cooble_template.pptx")
DEFAULT_MODEL = "gpt-4.1"

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
    """Maakt arbeidsvoorwaarden generiek: 29 vakantiedagen -> Vakantiedagen."""
    text = str(text or "").strip(" •-\n\t")
    text = re.sub(r"^\d+\s*(?:\+\s*\d+\s*)?(?:verlof)?dagen\b.*", "Vakantiedagen", text, flags=re.I)
    text = re.sub(r"^\d+\s*uur\b.*", "Werkuren", text, flags=re.I)
    text = re.sub(r"^€\s*[\d\.,]+.*", "Salaris", text, flags=re.I)
    text = re.sub(r"\bgoede\b\s+", "", text, flags=re.I).strip()
    if text.lower() in {"verlofdagen", "vakantie", "vrije dagen"}:
        text = "Vakantiedagen"
    if text.lower() in {"thuiswerken", "mogelijkheid om thuis te werken", "remote werken"}:
        text = "Hybride werken"
    return text[:1].upper() + text[1:] if text else text


def normalize_conditions(items: List[str]) -> List[str]:
    out: List[str] = []
    for item in split_one_topic_per_bullet(items):
        item = normalize_condition_label(item)
        if item and item not in out:
            out.append(item)
    return out[:3]


def normalize_age_distribution(items: List[str]) -> List[str]:
    """Zorgt dat leeftijdscategorieën altijd een percentage bevatten."""
    values = clean_list(items)
    defaults = ["25-34: 30%", "35-44: 35%", "45-54: 25%", "55+: 10%"]
    if not values:
        return defaults
    # Als de AI alleen categorieën teruggeeft, voeg standaardpercentages toe.
    default_percentages = ["30%", "35%", "25%", "10%", "0%"]
    result = []
    for i, val in enumerate(values[:5]):
        if "%" not in val:
            val = f"{val}: {default_percentages[i] if i < len(default_percentages) else '0%'}"
        result.append(val)
    return result


def clean_company_name(text: str) -> str:
    text = str(text or "").strip(" •-–—\n\t")
    text = re.sub(r"\s*\([^)]*\)\s*", " ", text).strip()
    remove_phrases = [
        "liever ook niet", "ook niet", "niet benaderen", "niet sourcen", "no go",
        "no-go", "eerst checken", "eerst check", "samenwerkingscontracten", "samenwerkingscontract",
        "concurrent", "concurrenten", "die wil hij", "die wil zij", "die wil men",
    ]
    lowered = text.lower()
    for phrase in remove_phrases:
        lowered = lowered.replace(phrase, "")
    # behoud hoofdletters zo veel mogelijk door dezelfde woorden uit originele tekst grof te verwijderen
    text = re.sub(r"(?i)liever ook niet|ook niet|niet benaderen|niet sourcen|no[- ]?go|eerst checken|eerst check|samenwerkingscontracten?|concurrenten?|die wil hij.*|die wil zij.*|die wil men.*", "", text).strip(" :;,-")
    # Pak alleen het eerste deel als er uitleg achter staat.
    text = re.split(r"\s+-\s+|\s+:\s+", text)[0].strip()
    # Veelgemaakte notatie normaliseren.
    replacements = {
        "witteveen en bos": "Witteveen+Bos",
        "witteveen+bos": "Witteveen+Bos",
        "haskoning": "Royal HaskoningDHV",
        "royal haskoning": "Royal HaskoningDHV",
        "antea": "Antea Group",
        "sweco": "Sweco",
    }
    key = text.lower().replace("&", "en").strip()
    return replacements.get(key, text)


def extract_no_go_companies_from_intake(intake_text: str) -> List[str]:
    """Haalt expliciet genoemde no-go/check-eerst organisaties uit intakeblokken."""
    if not intake_text:
        return []
    lines = [ln.strip() for ln in intake_text.splitlines()]
    triggers = ["no-go", "no go", "niet sourcen", "niet benaderen", "samenwerkingscontract", "eerst checken", "liever ook niet"]
    companies: List[str] = []
    collect_next = False
    remaining = 0
    for raw in lines:
        line = raw.strip(" •\t")
        low = line.lower()
        if not line:
            if collect_next:
                remaining -= 1
                if remaining <= 0:
                    collect_next = False
            continue
        has_trigger = any(t in low for t in triggers)
        if has_trigger:
            # Bedrijven kunnen op dezelfde regel of op de regels erna staan.
            collect_next = True
            remaining = 8
            cleaned = clean_company_name(line)
            if cleaned and len(cleaned.split()) <= 5 and cleaned.lower() not in {"samenwerkingscontract", "samenwerkingscontracten"}:
                companies.append(cleaned)
            continue
        if collect_next:
            # Stop bij duidelijke nieuwe vraag/sectie zonder bedrijfsnaam.
            if line.startswith("·") or line.lower().startswith(("meer voorbeeld", "mailtje", "laura", "koen")):
                collect_next = False
                continue
            # Split meerdere bedrijven op komma's of slashes.
            parts = re.split(r",|/|;", line)
            for part in parts:
                cleaned = clean_company_name(part)
                if cleaned and len(cleaned.split()) <= 5:
                    companies.append(cleaned)
            remaining -= 1
            if remaining <= 0:
                collect_next = False
    # Deduplicate, behoud volgorde
    result = []
    for c in companies:
        if c and c not in result:
            result.append(c)
    return result


def apply_business_rules(data: Dict[str, Any], intake_text: str, linkedin_size: str, vacature_text: str = "", extra_notes: str = "") -> Dict[str, Any]:
    """Harde sprint-0.4 regels die altijd gelden, ook na AI-generatie."""
    data = ensure_core_keys(data)
    b = data.setdefault("basisgegevens", {})
    b["datum"] = date.today().strftime("%d-%m-%Y")

    # Vul klantnaam, vacaturenaam en salaris deterministisch aan als AI ze mist.
    fallback = extract_basis_fallback(vacature_text, intake_text)
    for field in ["klantnaam", "vacaturenaam", "salaris"]:
        if is_empty_or_placeholder(b.get(field, "")) and fallback.get(field):
            b[field] = fallback[field]

    if linkedin_size.strip():
        data.setdefault("doelgroepanalyse", {})["verwachte_doelgroepgrootte"] = linkedin_size.strip()

    # No-go sourcing: alleen bedrijven, volledig uit intake/extra/AI-output, aangevuld met deterministische extractie.
    k = data.setdefault("kandidaatprofiel", {})
    ai_no_go = clean_list(k.get("no_go_sourcing", []))
    extracted = extract_no_go_companies_from_intake(intake_text)
    merged = []
    for item in ai_no_go + extracted:
        item = clean_company_name(item)
        if item and len(item.split()) <= 5 and item not in merged:
            merged.append(item)
    k["no_go_sourcing"] = merged

    # Pullfactoren en voorwaarden: één onderwerp per bullet.
    d = data.setdefault("doelgroepanalyse", {})
    d["pullfactoren"] = split_one_topic_per_bullet(d.get("pullfactoren", []))[:3]
    d["leeftijdsverdeling"] = normalize_age_distribution(d.get("leeftijdsverdeling", []))

    v = data.setdefault("voorwaarden", {})
    v["belangrijkste_arbeidsvoorwaarden"] = normalize_conditions(v.get("belangrijkste_arbeidsvoorwaarden", []))

    # Concurrenten mogen nooit placeholders zijn. Altijd bedrijfsnamen.
    c = data.setdefault("concurrentenanalyse", {})
    current_companies = [clean_company_name(x) for x in clean_list(c.get("bedrijven", []))]
    current_companies = [x for x in current_companies if x and not is_placeholder_company(x)]
    if not current_companies:
        current_companies = infer_competitors_offline(vacature_text, intake_text, current_companies)
    c["bedrijven"] = current_companies
    c["relevant"] = True
    return data


def bullets(items: List[str]) -> str:
    return "\n".join(clean_list(items))


def get_nested(data: Dict[str, Any], path: str, default: Any = "") -> Any:
    cur: Any = data
    for part in path.split("."):
        if isinstance(cur, dict) and part in cur:
            cur = cur[part]
        else:
            return default
    return cur


def replace_text_in_shape(shape, replacements: Dict[str, str]) -> None:
    # Recurse into grouped shapes.
    if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
        for subshape in shape.shapes:
            replace_text_in_shape(subshape, replacements)
        return

    # Replace in normal text frames.
    if hasattr(shape, "text_frame") and shape.has_text_frame:
        for paragraph in shape.text_frame.paragraphs:
            for run in paragraph.runs:
                text = run.text
                for key, value in replacements.items():
                    text = text.replace(key, value)
                run.text = text

    # Replace in tables.
    if hasattr(shape, "has_table") and shape.has_table:
        for row in shape.table.rows:
            for cell in row.cells:
                for paragraph in cell.text_frame.paragraphs:
                    for run in paragraph.runs:
                        text = run.text
                        for key, value in replacements.items():
                            text = text.replace(key, value)
                        run.text = text


def generate_pptx(data: Dict[str, Any]) -> bytes:
    if not TEMPLATE_PATH.exists():
        raise RuntimeError(f"Template niet gevonden: {TEMPLATE_PATH}")

    prs = Presentation(str(TEMPLATE_PATH))
    afspraken = data.get("afspraken") or []
    concurrenten = get_nested(data, "concurrentenanalyse.bedrijven", [])
    concurrenten_text = bullets(concurrenten) or get_nested(data, "concurrentenanalyse.toelichting", "")

    replacements = {
        "{{klantnaam}}": get_nested(data, "basisgegevens.klantnaam"),
        "{{vacaturenaam}}": get_nested(data, "basisgegevens.vacaturenaam"),
        "{{datum}}": get_nested(data, "basisgegevens.datum") or date.today().strftime("%d-%m-%Y"),
        "{{intake_samenvatting}}": data.get("intake_samenvatting", ""),
        "{{sourcingplan_strategie}}": get_nested(data, "sourcingplan.strategie"),
        "{{sourcingplan_doelgroep}}": get_nested(data, "sourcingplan.doelgroep"),
        "{{concurrentenanalyse}}": concurrenten_text,
        "{{zoekrichting}}": bullets(get_nested(data, "sourcingplan.zoekrichting", [])),
        "{{aanpak_toelichting}}": get_nested(data, "sourcingplan.toelichting"),
        "{{doelgroep_titel}}": get_nested(data, "doelgroepanalyse.doelgroep_titel") or get_nested(data, "basisgegevens.vacaturenaam"),
        "{{taken_verantwoordelijkheden}}": bullets(get_nested(data, "functieprofiel.taken_verantwoordelijkheden", [])),
        "{{eisen}}": bullets(get_nested(data, "kandidaatprofiel.eisen", [])),
        "{{voorkeuren}}": bullets(get_nested(data, "kandidaatprofiel.voorkeuren", [])),
        "{{no_go_sourcing}}": bullets(get_nested(data, "kandidaatprofiel.no_go_sourcing", [])),
        "{{doelgroepgrootte}}": get_nested(data, "doelgroepanalyse.verwachte_doelgroepgrootte"),
        "{{salaris}}": get_nested(data, "basisgegevens.salaris"),
        "{{locatie}}": get_nested(data, "basisgegevens.locatie"),
        "{{uren}}": get_nested(data, "basisgegevens.uren"),
        "{{usp_functie}}": bullets(get_nested(data, "functieprofiel.usp_functie", [])),
        "{{pullfactoren}}": bullets(get_nested(data, "doelgroepanalyse.pullfactoren", [])),
        "{{belangrijkste_arbeidsvoorwaarden}}": bullets(get_nested(data, "voorwaarden.belangrijkste_arbeidsvoorwaarden", [])),
        "{{geslacht_man}}": get_nested(data, "doelgroepanalyse.geslacht.man"),
        "{{geslacht_vrouw}}": get_nested(data, "doelgroepanalyse.geslacht.vrouw"),
        "{{leeftijdsverdeling}}": bullets(get_nested(data, "doelgroepanalyse.leeftijdsverdeling", [])),
        "{{afspraken_1}}": afspraken[0] if len(afspraken) > 0 else "",
        "{{afspraken_2}}": afspraken[1] if len(afspraken) > 1 else "",
        "{{afspraken_3}}": afspraken[2] if len(afspraken) > 2 else "",
    }

    for slide in prs.slides:
        for shape in slide.shapes:
            replace_text_in_shape(shape, replacements)

    with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as tmp:
        prs.save(tmp.name)
        return Path(tmp.name).read_bytes()


def schema_hint() -> str:
    return """
Geef uitsluitend geldige JSON terug in exact deze structuur:
{
  "basisgegevens": {
    "klantnaam": "",
    "vacaturenaam": "",
    "datum": "",
    "locatie": "",
    "uren": "",
    "salaris": ""
  },
  "intake_samenvatting": "",
  "functieprofiel": {
    "taken_verantwoordelijkheden": ["", "", ""],
    "usp_functie": ["", "", ""]
  },
  "kandidaatprofiel": {
    "eisen": ["", "", ""],
    "voorkeuren": ["", "", ""],
    "no_go_sourcing": []
  },
  "voorwaarden": {
    "belangrijkste_arbeidsvoorwaarden": ["", "", ""]
  },
  "doelgroepanalyse": {
    "doelgroep_titel": "",
    "verwachte_doelgroepgrootte": "",
    "regio": "Nederland",
    "pullfactoren": ["", "", ""],
    "geslacht": {"man": "", "vrouw": ""},
    "leeftijdsverdeling": ["", "", "", ""]
  },
  "sourcingplan": {
    "doelgroep": "",
    "strategie": "",
    "belangrijkste_functietitels": [],
    "zoekrichting": [],
    "toelichting": ""
  },
  "concurrentenanalyse": {
    "relevant": true,
    "bedrijven": [],
    "toelichting": ""
  },
  "afspraken": [],
  "kwaliteitscontrole": {
    "ontbrekende_informatie": [],
    "aannames": [],
    "waarschuwingen": []
  }
}
""".strip()


def build_prompt(vacature: str, intake: str, linkedin_size: str, extra: str) -> str:
    return f"""
Je bent een senior recruitment consultant en arbeidsmarktanalist. Maak compacte inhoud voor een Cooble startdocument.

Belangrijke regels:
- Output is Nederlands.
- Eén intake = één vacature.
- Formuleer kort en bondig, PowerPoint-stijl.
- Concurrentenanalyse is altijd relevant en altijd op bedrijfsniveau. Gebruik echte bedrijfsnamen. Nooit placeholders zoals Bedrijf A, Bedrijf B of Concurrent 1.
- Vrij formuleren mag, maar niet fantaseren.
- Datum: laat leeg of gebruik vandaag; de app overschrijft dit altijd met de generatiedatum.
- No-go sourcing: geef uitsluitend bedrijfsnamen terug. Neem álle no-go/check-eerst organisaties uit de intake over. Geen toelichting, geen zinnen. Laat geen enkel bedrijf weg.
- Pullfactoren zijn extern: bepaal ze vanuit arbeidsmarkt/doelgroep en internetonderzoek, niet uit de vacaturetekst.
- Belangrijkste arbeidsvoorwaarden: niet uit de vacaturetekst halen. Onderzoek welke arbeidsvoorwaarden de doelgroep belangrijk vindt. Gebruik generieke labels zoals "Vakantiedagen", niet "29 vakantiedagen".
- Eén bullet = één onderwerp. Combineer nooit meerdere onderwerpen in één bullet.
- Intake is leidend boven vacaturetekst.
- Extra opmerkingen zijn leidend boven alles.
- Houd lijsten kort: meestal precies 3 bullets, behalve no-go sourcing en concurrenten; die mogen meer bedrijven bevatten.
- Gebruik zelfverzekerde labels zoals "Hybride werken", niet "waarschijnlijk hybride werken".
- Leeftijdsverdeling: geef categorie én percentage, bijvoorbeeld "25-34: 30%".
- Als doelgroepgrootte uit LinkedIn is ingevuld, gebruik die waarde letterlijk.
- Klantnaam en vacaturenaam moeten altijd gevuld zijn. Haal klantnaam uit intake/vacaturetekst. Haal vacaturenaam uit intake/vacaturetitel.
- Salaris: als er een schaal of salarisrange in intake/vacature staat, neem die concreet over. Gebruik niet "in overleg" als er schalen of bedragen staan.
- Intake_samenvatting: schrijf concreet en iets uitgebreider. Dit veld moet in één dia duidelijk maken waar we naar zoeken, inclusief aanleiding, focus, nuances, wat juist niet past en nadruk uit de intake. Richtlijn: 70-120 woorden.
- Taken & verantwoordelijkheden: vermijd generieke bullets. Benoem de inhoudelijke context, doelgroep/klanttype, projecten of domein.
- Eisen: vermijd "relevante ervaring". Schrijf ervaring waarmee, bijvoorbeeld "ervaring met industriële waterprojecten".
- Doelgroep: zeer concreet. Gebruik vacaturetitel, domein, senioriteit, sector en relevante achtergrond. Nooit generiek zoals "kandidaten met relevante advieservaring".
- Concurrenten: doe internetonderzoek en geef echte bedrijven waar deze doelgroep werkt of vandaan kan komen.

{schema_hint()}

VACATURETEKST:
{vacature[:30000]}

INTAKE NOTES:
{intake[:30000]}

DOELGROEPGROOTTE GEVONDEN OP LINKEDIN:
{linkedin_size[:500]}

EXTRA OPMERKINGEN:
{extra[:10000]}
"""


def extract_json(text: str) -> Dict[str, Any]:
    text = (text or "").strip()
    if text.startswith("```"):
        text = re.sub(r"^```(?:json)?", "", text).strip()
        text = re.sub(r"```$", "", text).strip()
    try:
        return json.loads(text)
    except Exception:
        match = re.search(r"\{.*\}", text, flags=re.S)
        if match:
            return json.loads(match.group(0))
        raise


def generate_with_openai(prompt: str) -> Dict[str, Any]:
    api_key = st.secrets.get("OPENAI_API_KEY", "")
    if not api_key:
        raise RuntimeError("OPENAI_API_KEY ontbreekt in Streamlit Secrets.")
    client = OpenAI(api_key=api_key)
    model = st.secrets.get("OPENAI_MODEL", DEFAULT_MODEL)

    # Probeer eerst de Responses API met web search. Als dat niet beschikbaar is, valt de app terug op JSON-generatie zonder web search.
    try:
        response = client.responses.create(
            model=model,
            input=prompt,
            tools=[{"type": "web_search_preview"}],
        )
        return extract_json(response.output_text)
    except Exception as first_error:
        try:
            chat = client.chat.completions.create(
                model=model,
                messages=[
                    {"role": "system", "content": "Je geeft uitsluitend geldige JSON terug. Geen markdown."},
                    {"role": "user", "content": prompt},
                ],
                response_format={"type": "json_object"},
            )
            data = extract_json(chat.choices[0].message.content or "{}")
            data.setdefault("kwaliteitscontrole", {}).setdefault("waarschuwingen", []).append(
                "Webonderzoek is niet gebruikt; fallback naar standaard AI-generatie."
            )
            return data
        except Exception as second_error:
            raise RuntimeError(f"AI-generatie mislukt. Eerste fout: {first_error}. Tweede fout: {second_error}")


def demo_data() -> Dict[str, Any]:
    return {
        "basisgegevens": {
            "klantnaam": "Voorbeeldklant",
            "vacaturenaam": "Voorbeeldfunctie",
            "datum": date.today().strftime("%d-%m-%Y"),
            "locatie": "Nederland",
            "uren": "32-40 uur",
            "salaris": "In overleg",
        },
        "intake_samenvatting": "Voor deze opdracht zoeken we een kandidaat die inhoudelijke expertise combineert met adviesvaardigheden. De rol vraagt om zelfstandigheid, stakeholdermanagement en het vermogen om complexe vraagstukken praktisch te vertalen.",
        "functieprofiel": {
            "taken_verantwoordelijkheden": ["Adviseren van klanten", "Vertalen van vraagstukken naar oplossingen", "Samenwerken met multidisciplinaire teams"],
            "usp_functie": ["Inhoudelijk uitdagende projecten", "Ruimte voor ontwikkeling", "Impact bij diverse klanten"],
        },
        "kandidaatprofiel": {
            "eisen": ["HBO werk- en denkniveau", "Relevante ervaring", "Sterke adviesvaardigheden"],
            "voorkeuren": ["Consultancyervaring", "Ervaring binnen technische omgeving", "Zelfstandige werkhouding"],
            "no_go_sourcing": [],
        },
        "voorwaarden": {"belangrijkste_arbeidsvoorwaarden": ["Hybride werken", "Ontwikkelmogelijkheden", "Goede pensioenregeling"]},
        "doelgroepanalyse": {
            "doelgroep_titel": "Voorbeeldfunctie",
            "verwachte_doelgroepgrootte": "± 500",
            "regio": "Nederland",
            "pullfactoren": ["Hybride werken", "Inhoudelijke complexiteit", "Autonomie"],
            "geslacht": {"man": "60%", "vrouw": "40%"},
            "leeftijdsverdeling": ["25-34: 30%", "35-44: 40%", "45-54: 20%", "55+: 10%"],
        },
        "sourcingplan": {
            "doelgroep": "Kandidaten met relevante advieservaring binnen een inhoudelijk specialistisch domein.",
            "strategie": "Doelgroepgedreven sourcing met focus op vergelijkbare functies en organisaties.",
            "belangrijkste_functietitels": ["Consultant", "Adviseur", "Specialist"],
            "zoekrichting": ["LinkedIn sourcing", "Concurrenten op bedrijfsniveau", "Brede functietitelvarianten"],
            "toelichting": "Start breed en verfijn op inhoudelijke expertise en adviesvaardigheden.",
        },
        "concurrentenanalyse": {"relevant": True, "bedrijven": ["Bedrijf A", "Bedrijf B", "Bedrijf C"], "toelichting": "Vergelijkbare organisaties met relevante doelgroep."},
        "afspraken": ["Kandidaten worden voorgesteld na telefonische kennismaking.", "Feedback wordt zo snel mogelijk gedeeld.", "Bij profielwijzigingen wordt direct geschakeld."],
        "kwaliteitscontrole": {"ontbrekende_informatie": [], "aannames": [], "waarschuwingen": []},
    }


def editable_list(label: str, values: List[str], key: str, max_items: int = 6) -> List[str]:
    st.markdown(f"**{label}**")
    result = []
    values = values or []
    rows = max(max_items, len(values))
    for i in range(rows):
        default = values[i] if i < len(values) else ""
        val = st.text_input(f"{label} {i+1}", value=default, key=f"{key}_{i}", label_visibility="collapsed")
        if val.strip():
            result.append(val.strip())
    return result


def ensure_core_keys(data: Dict[str, Any]) -> Dict[str, Any]:
    data.setdefault("basisgegevens", {})
    data.setdefault("functieprofiel", {})
    data.setdefault("kandidaatprofiel", {})
    data.setdefault("voorwaarden", {})
    data.setdefault("doelgroepanalyse", {})
    data.setdefault("sourcingplan", {})
    data.setdefault("concurrentenanalyse", {})
    data.setdefault("afspraken", [])
    data.setdefault("kwaliteitscontrole", {"ontbrekende_informatie": [], "aannames": [], "waarschuwingen": []})
    return data


st.title("📄 Startdocument Generator")
st.caption("Upload de vacature en intake. De AI maakt een analyse, jij controleert de preview en downloadt daarna het startdocument.")

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
                    st.write("AI-analyse starten")
                    prompt = build_prompt(vacature_text, intake_text, linkedin_size, extra_notes)
                    data = generate_with_openai(prompt)
                data = apply_business_rules(data, intake_text + "\n" + extra_notes, linkedin_size, vacature_text, extra_notes)
                st.session_state.data = data
                status.update(label="Analyse klaar", state="complete", expanded=False)
            st.success("Analyse klaar. Controleer en pas eventueel aan.")
    except Exception as e:
        st.exception(e)

if st.session_state.data:
    data = ensure_core_keys(st.session_state.data)
    st.subheader("2. Preview & aanpassen")
    tabs = st.tabs(["Basis", "Functie", "Doelgroep", "Sourcing", "Afspraken", "Controle"])

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
        f["taken_verantwoordelijkheden"] = editable_list("Taken & verantwoordelijkheden", f.get("taken_verantwoordelijkheden", []), "taken", 3)
        k["eisen"] = editable_list("Eisen", k.get("eisen", []), "eisen", 3)
        k["voorkeuren"] = editable_list("Voorkeuren", k.get("voorkeuren", []), "voorkeuren", 3)
        f["usp_functie"] = editable_list("USP's van de functie", f.get("usp_functie", []), "usp", 3)
        k["no_go_sourcing"] = editable_list("No-go sourcing", k.get("no_go_sourcing", []), "nogo", 5)

    with tabs[2]:
        d = data.setdefault("doelgroepanalyse", {})
        v = data.setdefault("voorwaarden", {})
        c1, c2 = st.columns(2)
        d["verwachte_doelgroepgrootte"] = c1.text_input("Doelgroepgrootte", d.get("verwachte_doelgroepgrootte", ""))
        d["regio"] = c2.text_input("Regio", d.get("regio", "Nederland"))
        d["pullfactoren"] = editable_list("Pullfactoren", d.get("pullfactoren", []), "pull", 3)
        v["belangrijkste_arbeidsvoorwaarden"] = editable_list("Belangrijkste arbeidsvoorwaarden", v.get("belangrijkste_arbeidsvoorwaarden", []), "av", 3)
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

    with tabs[5]:
        qc = data.get("kwaliteitscontrole", {})
        st.json(qc)
        with st.expander("Volledige JSON bekijken"):
            st.json(data)

    st.session_state.data = data

    st.subheader("3. PowerPoint downloaden")
    try:
        data = apply_business_rules(data, "", linkedin_size)
        st.session_state.data = data
        pptx_bytes = generate_pptx(data)
        klant = get_nested(data, "basisgegevens.klantnaam", "klant") or "klant"
        vacature = get_nested(data, "basisgegevens.vacaturenaam", "vacature") or "vacature"
        filename = f"Startdocument_{klant}_{vacature}.pptx"
        filename = re.sub(r"[^A-Za-z0-9_\-\.]+", "_", filename)
        st.download_button(
            "Download Startdocument PowerPoint",
            pptx_bytes,
            filename,
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            type="primary",
        )
    except Exception as e:
        st.exception(e)
