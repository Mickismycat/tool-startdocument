# Startdocument Generator v1.0

Online Streamlit-app voor het genereren van een Cooble startdocument.

## Nieuw in v1.0

Deze sprint richt zich op het eindproduct: de PowerPoint.

- PowerPoint-export verbeterd met een nettere presentation engine.
- Lange tekstvakken worden automatisch iets kleiner gezet zodat tekst minder snel buiten het vak valt.
- Ongebruikte placeholders worden vóór export leeggemaakt.
- Bullets worden nu als echte scanbare bullets geëxporteerd.
- `doelgroep_regio` wordt correct gevuld in de doelgroepanalyse.
- Dubbele tekst `Kandidaat` boven voorkeuren is uit de template gehaald.
- v0.7/v0.8 outputregels blijven behouden: intake als lopende tekst, maximaal 3 bullets en één onderwerp per bullet.

## Deploy

1. Pak de zip uit.
2. Upload/vervang alle bestanden in je GitHub-repository.
3. Streamlit Cloud redeployt automatisch.
4. Controleer bij Streamlit Secrets:

```toml
OPENAI_API_KEY = "jouw-api-key"
```

Optioneel:

```toml
OPENAI_MODEL = "gpt-4.1"
```

## Bestanden

- `app.py` – Streamlit-app + AI-pipeline + PowerPoint-export
- `requirements.txt` – Python dependencies
- `templates/Startdocument_Cooble_template.pptx` – vaste PowerPoint-template
