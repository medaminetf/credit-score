# Notation AFRICAPITAL — Émetteurs de dette privée

Application Streamlit qui lit `Liste_Emetteurs.xlsx`, calcule les ratios et la
note (A/B/C/D) selon la méthodologie AFRICAPITAL, et affiche le résultat avec
l'interface de la maquette.

## Installation

```bash
pip install -r requirements.txt
```

## Lancement

```bash
streamlit run app.py
```

Place `Liste_Emetteurs.xlsx` dans le même dossier que `app.py` (chargement
automatique), ou charge-le via la barre latérale.

## Structure attendue du fichier Excel

- **Feuil1** : univers (Emetteur / Type / Secteur)
- **financier** : données des sociétés financières (en-têtes en ligne 2)
- **corporate** : données des corporates (en-têtes en ligne 2)

Il suffit de remplir les colonnes de données brutes — les ratios, les notes
par indicateur et la note finale sont calculés par l'application. Les
émetteurs sans données complètes apparaissent en « N/A » et sont exclus du
tableau de l'univers noté.

## ⚠️ Note importante

La formule Excel du « scoring 2 » (Total exigible / PNB) dans la feuille
`financier` est **inversée** par rapport au PDF de méthodologie
(`>40 → A … ≤20 → D` au lieu de `≤20 → A … >40 → D`). L'application suit le
PDF : un levier faible donne une meilleure note. Avec cette correction :
BMCI = **A** (B/A/A) et CIH = **B** (C/B/A).
