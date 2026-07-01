# Startdocument Generator v0.4

Online Streamlit-app voor het genereren van een Cooble startdocument.

## Nieuw in v0.4

- Datum wordt altijd automatisch gezet op de generatiedatum van vandaag.
- No-go sourcing toont alleen bedrijfsnamen.
- No-go sourcing wordt aangevuld vanuit intake-notities, zodat losse bedrijfsnamen in een no-go/check-eerst blok niet wegvallen.
- Belangrijkste arbeidsvoorwaarden worden generiek gemaakt, bijvoorbeeld `Vakantiedagen` in plaats van `29 vakantiedagen`.
- Eén bullet = één onderwerp. Samengestelde bullets worden waar mogelijk opgesplitst.
- Leeftijdsverdeling bevat categorieën én percentages.
- Slide-template heeft nu ook een placeholder voor leeftijdsverdeling.

## Deployen

1. Pak de zip uit.
2. Upload/vervang alle bestanden in je GitHub-repository.
3. Commit changes.
4. Streamlit redeployt automatisch.
5. Controleer in Streamlit Secrets dat dit aanwezig is:

```toml
OPENAI_API_KEY = "jouw-nieuwe-api-key"
```

## Bestanden

- `app.py` — hoofdapplicatie
- `requirements.txt` — Python packages
- `templates/Startdocument_Cooble_template.pptx` — vaste PowerPoint-template

## Testen

Start eerst met **Testmodus zonder API-key** om te controleren of preview en PowerPoint-export werken.
Daarna kun je **AI-generatie** gebruiken met een echte vacature en intake.
