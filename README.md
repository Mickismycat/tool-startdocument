# Startdocument Generator v0.3

Online Streamlit-app voor het genereren van een Cooble startdocument op basis van:

1. Vacaturetekst
2. Intake-notities
3. Doelgroepgrootte gevonden op LinkedIn (optioneel)
4. Extra opmerkingen (optioneel)

## Update in GitHub

Upload/vervang deze bestanden in je bestaande repository:

- `app.py`
- `requirements.txt`
- `templates/Startdocument_Cooble_template.pptx`

## Streamlit Secrets

Plaats je OpenAI API-key in Streamlit Cloud onder **App → Settings → Secrets**:

```toml
OPENAI_API_KEY = "sk-proj-..."
```

Optioneel:

```toml
OPENAI_MODEL = "gpt-4.1-mini"
```

## Gebruik

1. Upload/plak vacaturetekst.
2. Upload/plak intake-notities.
3. Vul eventueel doelgroepgrootte gevonden op LinkedIn in, bijvoorbeeld `± 500`.
4. Vul eventueel extra opmerkingen in.
5. Klik op **Genereer analyse**.
6. Controleer en bewerk de preview.
7. Download het PowerPoint-startdocument.

## Ontwerpregels

- Eén intake = één vacature.
- Concurrentenanalyse is altijd relevant en op bedrijfsniveau.
- Pullfactoren zijn extern en worden niet uit de vacature afgeleid.
- Arbeidsvoorwaarden worden gekozen vanuit wat de doelgroep belangrijk vindt.
- No-go sourcing wordt niet verzonnen.
- Vrij formuleren mag, feiten niet verzinnen.
