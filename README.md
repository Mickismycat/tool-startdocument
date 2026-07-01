# Startdocument Generator

Online Streamlit-app voor het genereren van een Cooble startdocument PowerPoint op basis van vacaturetekst, intake-notities en optionele doelgroepnotities.

## Bestanden

```text
app.py
requirements.txt
templates/Startdocument_Cooble_template.pptx
.streamlit/secrets.toml.example
```

## Deploy naar Streamlit Cloud

1. Maak een **private GitHub repository** aan, bijvoorbeeld `startdocument-tool`.
2. Upload alle bestanden uit deze map naar de repository.
3. Ga naar Streamlit Cloud.
4. Klik op **Create app**.
5. Kies je GitHub repository.
6. Main file path: `app.py`.
7. Klik op **Deploy**.

## OpenAI API-key toevoegen

Zet je API-key nooit in GitHub.

Ga in Streamlit Cloud naar:

```text
App → Settings → Secrets
```

Plak daar:

```toml
OPENAI_API_KEY = "sk-proj-..."
```

Optioneel kun je het model overschrijven:

```toml
OPENAI_MODEL = "gpt-4.1-mini"
```

## Gebruik

1. Open de Streamlit-link.
2. Upload of plak de vacaturetekst.
3. Upload of plak de intake-notities.
4. Vul optioneel LinkedIn/doelgroep-notities en extra opmerkingen in.
5. Klik op **Genereer analyse**.
6. Controleer en pas de preview aan.
7. Download de PowerPoint.

## Testmodus

De app heeft een testmodus zonder API-key. Daarmee kun je controleren of Streamlit en PowerPoint-export werken voordat je de AI-koppeling gebruikt.

## Belangrijke ontwerpregels

- Eén intake = één vacature.
- Concurrentenanalyse is altijd relevant en op bedrijfsniveau.
- No-go organisaties worden nooit zelf bedacht.
- Pullfactoren zijn extern en doelgroepgericht.
- Formulering mag vrij zijn, maar de AI mag niet fantaseren.
