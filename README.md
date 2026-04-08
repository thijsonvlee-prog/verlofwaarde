# Waarde vakantiedag

Streamlit-webapp die per medewerker één **gewogen-gemiddelde aanvulling
tijdens verlof p/u** uitrekent op basis van het beginsaldo per opbouwjaar en
het aanvullingstarief dat in dát opbouwjaar gold. Dit ene tarief geldt
vervolgens voor élk verlofuur dat in het lopende kalenderjaar wordt
opgenomen — geen administratie meer per opgenomen uur.

## Hoe het werkt

1. Upload het **verlofsaldi-export** (.xlsx) uit de salarisadministratie.
2. Upload het **aanvulling-per-uur-export** (.xlsx).
3. Controleer per **basis-verloftype** (jaartal weggelaten) of het
   aanvullingsrechtig is. Sensible defaults zijn al ingevuld.
4. Klik **Bereken** → tabel met per medewerker totaal uren, totaal
   aanvulling en gewogen p/u, plus een download-knop voor een .xlsx-rapport.

### Rekenregel

Voor iedere verlofsaldi-regel:

- `opbouwjaar` = 4-cijferig jaartal uit `Type verlof`. Geen jaartal → rij
  wordt volledig overgeslagen (Negatief verlof e.d.).
- `basis_type` = `Type verlof` met het jaartal weggehaald.
- `aanvulling p/u` = lookup in de aanvullingstabel op `(Mdw., opbouwjaar)`.
- Niet-aanvullingsrechtige basis-types krijgen `€0` in de teller maar tellen
  wel mee in de noemer (verlagen het gemiddelde).
- `gewicht = beginsaldo × effectieve aanvulling p/u`.

Per medewerker: `gewogen p/u = Σ gewicht / Σ beginsaldo`.

## Lokaal draaien

```bash
pip install -r requirements.txt
streamlit run app.py
```

De app start op http://localhost:8501. Zonder `.streamlit/secrets.toml` is er
geen wachtwoord. Voor productie:

```bash
cp .streamlit/secrets.toml.example .streamlit/secrets.toml
# en stel APP_PASSWORD in
```

## Tests

```bash
pytest
```

De testsuite verifieert onder andere het Joris-voorbeeld uit de
specificatie: gewogen p/u ≈ **€6,52**.

## Deployment

### Streamlit Community Cloud (aanbevolen)

1. Push naar GitHub.
2. Op [share.streamlit.io](https://share.streamlit.io) een nieuwe app
   maken vanuit deze repo, branch `main` of `claude/define-vacation-day-value-9FQ3Z`,
   entrypoint `app.py`.
3. Bij **Advanced settings → Secrets** invoeren:
   ```toml
   APP_PASSWORD = "een-sterk-wachtwoord"
   ```
4. Deploy. Iedere push naar de branch triggert een rebuild.

### Docker / Render.com (alternatief)

```bash
docker build -t verlofwaarde .
docker run -p 8501:8501 -e APP_PASSWORD=geheim verlofwaarde
```

`render.yaml` ligt klaar voor [Render.com](https://render.com): koppel de
repo, zet `APP_PASSWORD` als geheime env-var, klaar.

## Architectuur

```
app.py                       # Streamlit UI + flow
verlofwaarde/
├── parsing.py               # xlsx → genormaliseerde DataFrames
├── calculator.py            # Gewogen-gemiddelde-berekening
├── instellingen.py          # Defaults voor aanvullingsrechtig per basis-type
└── export.py                # Resultaat → xlsx-bytes voor download
tests/
├── test_parsing.py
├── test_calculator.py
└── fixtures/                # Joris-voorbeeld als .xlsx
```

## Privacy

Salarisdata is gevoelig. De ingebouwde password-gate (via Streamlit secrets)
is een minimale beveiliging — voor productie op gevoelige data is echte SSO
of een deployment achter reverse proxy met IP-whitelisting aan te bevelen.
Alle uploads worden alleen in-memory verwerkt; er wordt niets persistent
opgeslagen.
