# Startdocument Generator v2.4.8

Hotfix op v2.4.7 voor de fout `Web Search cannot be used with JSON mode`.

Wijzigingen:
- Webresearch en JSON-output zijn nu technisch gescheiden in twee stappen.
- Stap 1 gebruikt OpenAI Responses API + `web_search` verplicht voor live internetonderzoek.
- Stap 2 structureert uitsluitend de gevonden onderzoeksnotities naar geldig JSON zonder webtool.
- Er is geen fallback naar vacaturetekst voor doelgroep-, pullfactor-, arbeidsvoorwaarden- of demografieonderzoek.
- Bij mislukte webresearch stopt de tool bewust met een duidelijke foutmelding.
- Extra syntax- en functiedefinitiecontrole uitgevoerd.

Upload de volledige inhoud naar GitHub en laat Streamlit opnieuw deployen.
