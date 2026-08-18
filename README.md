# Startdocument Generator v1.8

Wijzigingen in v1.8:
- Salaris toont alleen een getal/range; bij salarisschalen blijft `Schaal` staan.
- Pullfactoren worden verplicht via live web search onderzocht.
- Pullfactoren worden genormaliseerd naar nette, korte labels of natuurlijke korte zinnen.
- Losse fragmenten zoals `certificering` worden vermeden/gecorrigeerd.
- Bestaande Cooble-template en PowerPoint-logica blijven ongewijzigd.

# Startdocument Generator v1.7

Belangrijkste wijziging: online arbeidsmarktonderzoek is nu **verplicht** voor doelgroep, arbeidsvoorwaarden en pullfactoren.

- Webresearch gebruikt de actuele Responses API `web_search` tool.
- `tool_choice="required"` dwingt af dat er daadwerkelijk wordt gezocht; de app valt niet meer stilletjes terug op vacaturetekst of algemene AI-kennis.
- Arbeidsvoorwaarden worden onderzocht op basis van de **doelgroep**, niet op basis van wat de opdrachtgever aanbiedt.
- De arbeidsvoorwaarden-slide bevat alleen generieke categorieën zoals `Salaris`, `Pensioenregeling` of `Vakantiedagen`; nooit bedragen, percentages of concrete aantallen dagen.
- Pullfactoren blijven een apart onderzoek naar wat de doelgroep in beweging brengt en wat zij in een vacature wil terugzien.
- Als live webresearch technisch mislukt, stopt de analyse met een duidelijke foutmelding in plaats van onjuiste vacaturevoorwaarden te tonen.
- PowerPoint en Cooble-template zijn verder ongewijzigd ten opzichte van v1.6.

Upload alle bestanden en mappen naar dezelfde GitHub-repository. Streamlit redeployt daarna automatisch.
