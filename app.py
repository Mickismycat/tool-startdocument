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
from pypdf import PdfReader

APP_TITLE = "Startdocument Generator"
TEMPLATE_PATH = Path("templates/Startdocument_Cooble_template.pptx")
DEFAULT_MODEL = "gpt-4.1-mini"

st.set_page_config(page_title=APP_TITLE, page_icon="📄", layout="wide")


def read_docx(file) -> str:
    doc = Document(file)
    parts = [p.text for p in doc.paragraphs if p.text.strip()]
    return "\n".join(parts)


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
    if name.endswith(".docx"):
        return read_docx(file)
    if name.endswith(".pdf"):
        return read_pdf(file)
    if name.endswith(".txt"):
        return file.read().decode("utf-8", errors="ignore")
    return ""


def bullets(items: List[str]) -> str:
    clean = [str(x).strip(" •-\n\t") for x in (items or []) if str(x).strip()]
    return "\n".join(clean)


def get_nested(data: Dict[str, Any], path: str, default: Any = "") -> Any:
    cur = data
    for part in path.split("."):
        if isinstance(cur, dict) and part in cur:
            cur = cur[part]
        else:
            return default
    return cur


def safe_join(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, list):
        return bullets(value)
    if isinstance(value, dict):
        return json.dumps(value, ensure_ascii=False)
    return str(value)


def replace_text_in_shape(shape, replacements: Dict[str, str]):
    if not hasattr(shape, "text_frame"):
        return
    for paragraph in shape.text_frame.paragraphs:
        for run in paragraph.runs:
            text = run.text
            for key, value in replacements.items():
                text = text.replace(key, value)
            run.text = text


def generate_pptx(data: Dict[str, Any]) -> bytes:
    prs = Presentation(str(TEMPLATE_PATH))
    afspraken = data.get("afspraken") or []
    concurrenten = get_nested(data, "concurrentenanalyse.bedrijven", [])
    concurrenten_text = bullets(concurrenten)
    if not concurrenten_text:
        concurrenten_text = get_nested(data, "concurrentenanalyse.toelichting", "")

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
        "{{salaris}}": get_nested(data, "basisgegevens.salaris"),
        "{{locatie}}": get_nested(data, "basisgegevens.locatie"),
        "{{uren}}": get_nested(data, "basisgegevens.uren"),
        "{{usp_functie}}": bullets(get_nested(data, "functieprofiel.usp_functie", [])),
        "{{pullfactoren}}": bullets(get_nested(data, "doelgroepanalyse.pullfactoren", [])),
        "{{belangrijkste_arbeidsvoorwaarden}}": bullets(get_nested(data, "voorwaarden.belangrijkste_arbeidsvoorwaarden", [])),
        "{{geslacht_man}}": get_nested(data, "doelgroepanalyse.geslacht.man"),
        "{{geslacht_vrouw}}": get_nested(data, "doelgroepanalyse.geslacht.vrouw"),
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
    "leeftijdsverdeling": []
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


def build_prompt(vacature: str, intake: str, linkedin_notes: str, extra: str) -> str:
    return f"""
Je bent een senior recruitment consultant en arbeidsmarktanalist. Maak compacte inhoud voor een Cooble startdocument.

Regels:
- Output is Nederlands.
- Formuleer kort en bondig, PowerPoint-stijl.
- Eén intake = één vacature.
- Concurrentenanalyse is altijd relevant en altijd op bedrijfsniveau.
- Vrij formuleren mag, maar niet fantaseren.
- No-go organisaties alleen uit intake of extra opmerkingen halen, nooit zelf bedenken.
- Pullfactoren zijn extern: bepaal ze vanuit de arbeidsmarkt/doelgroep, niet uit de vacaturetekst.
- Arbeidsvoorwaarden: bepaal welke voorwaarden voor deze doelgroep belangrijk zijn; combineer arbeidsmarktanalyse met wat de vacature biedt.
- Intake is leidend boven vacaturetekst.
- Extra opmerkingen zijn leidend boven alles.
- Houd lijsten kort: meestal 3 bullets.
- Gebruik zelfverzekerde labels zoals "Hybride werken", niet "waarschijnlijk hybride werken".

{schema_hint()}

VACATURETEKST:
{vacature[:30000]}

INTAKE NOTES:
{intake[:30000]}

LINKEDIN / DOELGROEP-NOTITIES:
{linkedin_notes[:10000]}

EXTRA OPMERKINGEN:
{extra[:10000]}
"""


def extract_json(text: str) -> Dict[str, Any]:
    text = text.strip()
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

    # First try Responses API with web search. If not available for the account/model, fallback to standard JSON generation.
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
            return extract_json(chat.choices[0].message.content or "{}")
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
    for i in range(max(max_items, len(values))):
        default = values[i] if i < len(values) else ""
        val = st.text_input(f"{label} {i+1}", value=default, key=f"{key}_{i}", label_visibility="collapsed")
        if val.strip():
            result.append(val.strip())
    return result


st.title("📄 Startdocument Generator")
st.caption("Upload of plak vacature- en intake-informatie. De tool maakt een ingevuld Cooble startdocument.")

with st.sidebar:
    st.header("Instellingen")
    mode = st.radio("Modus", ["AI-generatie", "Testmodus zonder API-key"], index=0)
    st.caption("Gebruik testmodus om te controleren of de app en PowerPoint-export werken.")

st.subheader("1. Input")
col1, col2 = st.columns(2)
with col1:
    vacature_file = st.file_uploader("Vacature uploaden", type=["docx", "pdf", "txt"], key="vac_file")
    vacature_paste = st.text_area("Of plak vacaturetekst", height=240)
with col2:
    intake_file = st.file_uploader("Intake uploaden", type=["docx", "pdf", "txt"], key="intake_file")
    intake_paste = st.text_area("Of plak intake-notities", height=240)

linkedin_notes = st.text_area("LinkedIn / doelgroep-notities (optioneel)", height=110)
extra_notes = st.text_area("Extra opmerkingen (optioneel)", height=90)

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
            with st.spinner("Analyse wordt gemaakt..."):
                if mode == "Testmodus zonder API-key":
                    data = demo_data()
                else:
                    prompt = build_prompt(vacature_text, intake_text, linkedin_notes, extra_notes)
                    data = generate_with_openai(prompt)
                st.session_state.data = data
            st.success("Analyse klaar. Controleer en pas eventueel aan.")
    except Exception as e:
        st.exception(e)

if st.session_state.data:
    data = st.session_state.data
    st.subheader("2. Preview & aanpassen")
    tabs = st.tabs(["Basis", "Functie", "Doelgroep", "Sourcing", "Afspraken", "Controle"])

    with tabs[0]:
        b = data.setdefault("basisgegevens", {})
        c1, c2, c3 = st.columns(3)
        b["klantnaam"] = c1.text_input("Klantnaam", b.get("klantnaam", ""))
        b["vacaturenaam"] = c2.text_input("Vacaturenaam", b.get("vacaturenaam", ""))
        b["datum"] = c3.text_input("Datum", b.get("datum", date.today().strftime("%d-%m-%Y")))
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

    st.subheader("3. PowerPoint downloaden")
    try:
        pptx_bytes = generate_pptx(data)
        filename = f"Startdocument_{get_nested(data, 'basisgegevens.klantnaam', 'klant')}_{get_nested(data, 'basisgegevens.vacaturenaam', 'vacature')}.pptx"
        filename = re.sub(r"[^A-Za-z0-9_\-\.]+", "_", filename)
        st.download_button("Download Startdocument PowerPoint", pptx_bytes, filename, mime="application/vnd.openxmlformats-officedocument.presentationml.presentation", type="primary")
    except Exception as e:
        st.exception(e)
