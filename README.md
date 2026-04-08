# Waarde vakantiedag — gewogen aanvulling p/u

Excel-template die per medewerker één **gewogen gemiddelde aanvulling p/u**
berekent op basis van het openstaande verlofsaldo per opbouwjaar. Hiermee
krijgt elk in het lopende kalenderjaar opgenomen verlofuur dezelfde
aanvulling, zonder dat de salarisadministratie per uur hoeft te bepalen uit
welk opbouwjaar het komt.

## Achtergrond

Een chauffeur die een dag verlof opneemt heeft recht op dezelfde beloning als
op een gewone werkdag: 8 × het actuele uurloon plus een aanvulling voor het
gemiste gemiddelde aan overuren en toeslagen (de "aanvulling tijdens verlof").

Die aanvulling is per opbouwjaar anders, want zij is gebaseerd op de
gemiddelde overuren van dat jaar en op het uurloon van dát jaar. Als een
medewerker verlof opneemt dat in verschillende jaren is opgebouwd, zou je
strikt genomen per uur moeten bepalen uit welk potje het komt — onwerkbaar in
de salarisadministratie.

De oplossing: **éénmaal per jaar per medewerker** een gewogen gemiddelde
aanvulling p/u uitrekenen op basis van het verlofsaldo dat er aan het begin
van het jaar staat. Daarna krijgt elk in dat jaar opgenomen verlofuur
diezelfde aanvulling — ongeacht uit welk opbouwjaar het komt.

### Rekenregel

Voor iedere verlofregel met een opbouwjaar in de naam:

```
gewicht = Beginsaldo × aanvulling p/u (uit opbouwjaar)
```

Niet-aanvullingsrechtige verlofsoorten (instelbaar per basis-type) krijgen
gewicht 0, maar hun uren tellen wél in de noemer — zo verlagen ze het
gewogen gemiddelde, conform het uitgangspunt "over ieder verlofuur wordt
aanvulling uitbetaald, maar voor uren waarover eigenlijk geen aanvulling
hoort te gaan rekenen we €0 mee".

Per medewerker (alle DV's samen):

```
gewogen aanvulling p/u = SOM(gewicht) / SOM(Beginsaldo, meetellend)
```

Types zonder jaartal in de naam (bijv. "Negatief verlof") tellen niet mee,
niet in de teller en niet in de noemer.

## Gebruik

1. Installeer de build-dependency en genereer de template:

   ```
   pip install -r requirements.txt
   python scripts/build_template.py
   ```

   Resultaat: `dist/waarde_vakantiedag_template.xlsx`. Het bestand staat
   ook ingecheckt in de repo, dus stap 1 is alleen nodig als je het wilt
   herbouwen.

2. Open de xlsx in **Excel 365** (moderne formules als `LET`, `FILTER`,
   `XLOOKUP`, `SEQUENCE` zijn vereist).

3. Werkblad `Verlofsaldi` — plak hier het verlofsaldo-export. De kolommen
   moeten exact zijn:

   ```
   Mdw. | Naam | DV | Type verlof | Bkjr. | Actueel saldo deze periode |
   Geboekt / nog op te nemen | Vorige periode | Perioderecht |
   Extra perioderecht | Beginsaldo | Corr. | Opgenomen | Saldo
   ```

4. Werkblad `Aanvulling per uur` — plak hier het aanvullings-export:

   ```
   Mdw. | Naam | DV | LC | Omschrijving | Totaalbedrag | Bkjr.
   ```

5. Werkblad `Instellingen` — controleer per **basis-verloftype** (de naam
   zonder jaartal) of de verlofsoort aanvullingsrechtig is (`Ja`/`Nee`).
   Kolom D/E toont automatisch alle basistypes die in de data voorkomen,
   met status `OK` of `ONTBREEKT`. Ontbrekende types moet je toevoegen.

6. Werkblad `Resultaat` — per medewerker (alle DV's samengevoegd) de
   totale uren, totale aanvulling, en het gewogen gemiddelde p/u.

## Voorbeeld — Joris Alofs (medewerker 10004)

Met de meegeleverde fixtures:

| Verloftype                              | Beginsaldo | Opbouwjaar | Aanvulling p/u |    Gewicht |
|-----------------------------------------|-----------:|:----------:|---------------:|-----------:|
| Negatief verlof                         |       0,00 |     —      |              — |          — |
| Bovenwettelijke rechten 2025            |       5,68 |    2025    |          €4,54 |      25,79 |
| Wettelijke rechten 2026                 |     160,00 |    2026    |          €6,57 |   1.051,20 |
| Bovenwettelijke rechten 2026            |      28,00 |    2026    |          €6,57 |     183,96 |
| Bovenwettelijke rechten 2026 (aanv.)    |      56,00 |    2026    |          €6,57 |     367,92 |

Totalen voor Mdw. 10004:

- Totaal uren: **249,68**
- Totaal aanvulling: **€1.628,87**
- Gewogen aanvulling p/u: **€6,5238**

## Projectstructuur

```
.
├── README.md                               Deze file
├── requirements.txt                        openpyxl (build-dep)
├── scripts/
│   └── build_template.py                   Genereert de xlsx
├── fixtures/
│   ├── voorbeeld_verlofsaldi.csv           Demo-data (verlofsaldi)
│   └── voorbeeld_aanvulling.csv            Demo-data (aanvulling p/u)
└── dist/
    └── waarde_vakantiedag_template.xlsx    Gegenereerde template
```

De logica zit uitsluitend in `scripts/build_template.py`: dát bestand maakt
de werkbladen, vult de hulpkolommen met formules en schrijft de xlsx. Wil je
de berekening aanpassen, wijzig dan de formules daar en draai het script
opnieuw.
