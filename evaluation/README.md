# Evaluation Framework

Evaluacija pouzdanosti Free Lector nalaza putem ručnog obeležavanja uzoraka.

## Cilj

Kvantifikovati **preciznost** (precision) za svaku kategoriju pravila,
koristeći ručno obeležene uzorke iz stvarnih dokumenata.

**Zašto samo preciznost:** Free Lector ne može izmeriti recall jer ne postoji
iscrpan spisak svih grešaka u dokumentu. Za recall bi bio potreban "zlatni
standard" - kompletno lektorisan dokument. To zahteva značajno vreme
profesionalnog lektora. Zbog toga ne izveštavamo recall, F1 niti ukupnu tačnost (accuracy).

**Zašto je confidence heuristički:** Vrednost `predicted_confidence` u pravilima
je ručno procenjena heuristička vrednost, ne empirijski kalibrisana verovatnoća.
Brier score pomaže da se proveri koliko su te procene u praksi tačne.

## Schema: labels.csv

| Kolona | Obavezna | Opis |
|--------|----------|------|
| document_id | Da | Stabilni ID dokumenta (iz export JSON-a) |
| version_id | Da | Version ID dokumenta - obezbeđuje da se obeležavanje odnosi na tačnu verziju |
| rule_id | Da | ID pravila koje je proizvelo nalaz (npr. `duplicate_words_body`) |
| finding_id | Da | ID nalaza (npr. f-17) |
| location_key | Da | Lokacija nalaza (paragraphId) |
| predicted_confidence | Da | Confidence vrednost koju je pravilo dodelilo (0.0-1.0) |
| label | Da | 1 = ispravno prijavljen nalaz (true positive), 0 = lažni pozitiv (false positive) |
| rule_version | Da | Verzija pravila (npr. r4-82tests) - omogućava praćenje promena pravila kroz vreme |
| reviewer | Da | Identifikator osobe koja je obeležila |
| note | Ne | Opcioni komentar (može biti prazan) |

### Definicija labela

- **1 (True Positive):** Nalaz je korektan - postoji realna greška koju je alat prijavio.
- **0 (False Positive):** Nalaz je pogrešan - alat je prijavio grešku koja ne postoji.

### Svrha ključnih kolona

- **version_id** - Vezuje obeležene podatke za tačnu verziju dokumenta. Ako se dokument promeni, stari
  labeli se ne mešaju sa novim.
- **rule_version** - Pravila se menjaju kroz vreme. Ovo omogućava poređenje preciznosti
  pre i posle izmene pravila.
- **rule_id** - Identifikator pravila. Evaluacija se grupiše po rule_id, ne po category.

## Kako dodati labele

1. Pokrenuti audit na test dokumentu i izvesti JSON.
2. Nasumično uzorkovati nalaze za obeležavanje (koristiti `sample-size.js`).
3. Za svaki nalaz dodati red u `labels.csv`:
   - Koristiti `document_id` i `version_id` iz JSON eksporta
   - `rule_id` iz nalaza
   - `finding_id` iz nalaza
   - `location_key` - paragraphId nalaza
   - `predicted_confidence` - confidence vrednost nalaza
   - `label` - 1 ili 0
   - `rule_version` - verzija pravila (navesti poslednju poznatu)
   - `reviewer` - vaš identifikator
   - `note` - opcioni komentar

## Pokretanje

```bash
# Evaluacija preciznosti
node evaluation/evaluate-confidence.js

# Izračun potrebne veličine uzorka
node evaluation/sample-size.js --p 0.5 --margin 0.05 --confidence 0.95 --findings-per-document 5

# Sa design effect-om za korelisane nalaze
node evaluation/sample-size.js --p 0.5 --margin 0.05 --confidence 0.95 --design-effect 1.5
```

## Interpretacija rezultata

### Wilson interval

Wilson 95% confidence interval daje raspon u kojem se očekuje prava preciznost pravila.
Širina intervala zavisi od veličine uzorka - manji uzorak → širi interval.

### Brier score

Brier score meri koliko su `predicted_confidence` vrednosti blizu stvarnih ishoda.
Savršena kalibracija = 0. Nasumično pogađanje sa p=0.5 daje ~0.25.

### Reliability bins

Reliability diagram deli confidence vrednosti u bin-ove i poredi srednju predviđenu
sa stvarnom stopom ispravnosti. Idealno: mean_predicted ≈ observed_correct.

## Korelacija nalaza

Nalazi iz istog dokumenta **nisu statistički nezavisni**. Ako jedno pravilo ima
problem sa specifičnim tipom teksta, svi njegovi nalazi u tom dokumentu biće
pogođeni. Wilson intervali po pojedinačnim nalazima mogu **potceniti** stvarnu neizvesnost.

Zato `sample-size.js` podržava `--design-effect` faktor. Kada je designEffect=1,
skript ispisuje upozorenje: "Ovo je optimistična donja procena koja pretpostavlja
nezavisne nalaze."

## Zašto McNemar test nije implementiran

McNemar test služi za poređenje dve verzije sistema na istom uzorku. Trenutno:

1. Identifikatori lokacija (location_key) nisu garantovano stabilni kroz verzije parsera/pravila
2. Prvo je potreban zlatni standard - skup kandidatskih lokacija koji ostaje stabilan
3. Tek nakon što postoji stabilno uparivanje, McNemar se može primeniti

Dok ti uslovi ne budu ispunjeni, implementacija bi davala nepouzdane rezultate.

## Zašto compare-versions.js ne postoji

Fajl `compare-versions.js` nije implementiran jer zahteva:
1. Obeležene podatke za najmanje dve verzije pravila
2. Isti set dokumenata proverenih sa obe verzije
3. McNemar ili ekvivalentan test

Kada ti uslovi budu ispunjeni, skript će biti dodat.

## Fajlovi

- `labels.csv` - ručno obeleženi uzorci
- `evaluate-confidence.js` - izračunava precision, Wilson CI, Brier score po rule_id
- `sample-size.js` - izračunava potrebnu veličinu uzorka (Cochran formula + design effect)
- `report/` - generisani izveštaji (gitignore-ovano)
- `README.md` - ovaj fajl
