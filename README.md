# Startdocument Generator v2.4.7

Hotfix op v2.4.6 voor de fout `Expecting ',' delimiter` tijdens online arbeidsvoorwaardenonderzoek.

Wijzigingen:
- Webresearch gebruikt nu het actuele OpenAI Responses API `web_search` tooltype.
- JSON wordt op API-niveau afgedwongen met `text.format = json_object`.
- Websearch blijft verplicht; er is geen fallback naar vacaturetekst of algemene modelkennis.
- Automatische herstelpoging bij een onvolledige/technisch mislukte response.
- Alle functionaliteit van v2.4.6 blijft behouden.

Upload de volledige inhoud naar GitHub en laat Streamlit opnieuw deployen.
