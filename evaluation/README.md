# Evaluation Framework

Evaluacija pouzdanosti Free Lector nalaza putem ručnog obeležavanja uzoraka.

## Cilj

Kvantifikovati preciznost (precision) i odziv (recall) za svaku kategoriju nalaza,
koristeći ručno obeležene uzorke iz stvarnih dokumenata.

## Tok rada

1. Pokrenuti audit na test dokumentu.
2. Nasumično uzorkovati nalaze za obeležavanje (koristiti `sample-size.js` za izračunavanje potrebne veličine uzorka).
3. Ručno popuniti `labels.csv` sa kolonama: finding_id, category, human_label, confidence_override, comment.
4. Pokrenuti `evaluate-confidence.js` za generisanje izveštaja.

## Fajlovi

- `labels.csv` — ručno obeleženi uzorci
- `evaluate-confidence.js` — izračunava metriku (precision, recall, F1) po kategoriji
- `sample-size.js` — izračunava potrebnu veličinu uzorka za željenu marginu greške
- `report/` — generisani izveštaji (gitignore-ovano)

## Pokretanje

```bash
npm run evaluate:confidence
npm run evaluate:sample-size
```

## Format labels.csv

| Kolona | Opis |
|--------|------|
| finding_id | ID nalaza (npr. F-0001) |
| category | Kategorija nalaza |
| human_label | TP (true positive), FP (false positive), FN (false negative) |
| confidence_override | Opciono: korigovana pouzdanost (0.0-1.0) |
| comment | Opcioni komentar |
