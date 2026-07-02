# Startdocument Generator v0.5

Wijzigingen in v0.5:
- Klantnaam en vacaturenaam worden altijd aangevuld uit vacature/intake als AI ze mist.
- Salaris pakt nu ook schalen of ranges uit intake/vacature en zet niet zomaar 'in overleg'.
- Intake-samenvatting is concreter en uitgebreider volgens startdocument-doel.
- Taken, eisen en doelgroep zijn aangescherpt: minder generiek, meer functie/domeinspecifiek.
- Concurrentenanalyse mag geen placeholders meer bevatten zoals Bedrijf A/B/C.
- Fallback voor concurrenten toegevoegd als webonderzoek geen namen oplevert.
- No-go sourcing blijft: alleen bedrijven en zo volledig mogelijk uit intake.
- Datum blijft altijd de generatiedatum.

Deploy:
1. Upload/vervang alle bestanden in je GitHub repo.
2. Streamlit redeployt automatisch.
3. Controleer Secrets: OPENAI_API_KEY = "..."
