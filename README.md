# Startdocument Generator v2.4.1

Bugfix release voor v2.4. Herstelt de ontbrekende OpenAI-helperfuncties waardoor `call_openai_json` niet gedefinieerd was. Deze versie behoudt de strikte externe doelgroepresearch en de template-first PowerPoint-export.

# Startdocument Generator v2.3

Deze versie gebruikt het opnieuw aangeleverde, correcte **Cooble-template** als vaste layout.

## Belangrijkste wijzigingen
- Het template bevat vaste placeholders op de exacte bestaande tekstposities.
- De PowerPoint-code verandert **geen** fonts, lettergroottes, posities, marges of regelafstand.
- Alleen placeholdertekst wordt vervangen; de leeftijdsverdeling gebruikt één vast beeldanker.
- Man/vrouw en leeftijd komen uitsluitend uit verplicht webonderzoek; er zijn geen vaste, verzonnen demografische tabellen meer.
- Demografie wordt afgerond in stappen van 5% voor meer stabiliteit tussen runs.
- Als demografisch webonderzoek onvolledig is, volgt één automatische herhaalpoging; daarna stopt de tool in plaats van lege of verzonnen percentages te tonen.

Upload alle bestanden en mappen naar GitHub, inclusief `templates/Startdocument_Cooble_template.pptx`.
