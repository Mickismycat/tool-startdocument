# Startdocument Generator v0.6

Online Streamlit-app voor het genereren van een Cooble startdocument.

## Nieuw in v0.6

- AI-pipeline in meerdere stappen: feitenextractie → arbeidsmarktresearch → writer → kwaliteitscontrole.
- Research-stap voor doelgroep, arbeidsvoorwaarden, pullfactoren en concurrenten.
- Automatische controle op generieke teksten zoals “relevante ervaring”.
- Concurrentenanalyse blijft verplicht en op bedrijfsniveau.
- No-go sourcing blijft alleen bedrijfsnamen uit intake/extra opmerkingen.
- Datum blijft altijd de generatiedatum.

## Deploy

Upload alle bestanden naar GitHub. Streamlit Cloud redeployt automatisch.

## Secrets

Zet in Streamlit Secrets:

```toml
OPENAI_API_KEY = "jouw-api-key"
```

Optioneel:

```toml
OPENAI_MODEL = "gpt-4.1"
```
