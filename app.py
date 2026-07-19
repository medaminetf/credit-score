# -*- coding: utf-8 -*-
"""
Notation AFRICAPITAL des émetteurs de dette privée
---------------------------------------------------
Application Streamlit : lit le fichier Liste_Emetteurs.xlsx (feuilles
"Feuil1", "financier", "corporate"), calcule les 3 ratios par famille
d'émetteur, convertit chaque ratio en note (A/B/C/D), calcule la note
globale (moyenne des points A=4, B=3, C=2, D=1) et affiche le résultat.

Lancement :  streamlit run app.py
"""

import io
from pathlib import Path

import pandas as pd
import streamlit as st

# ----------------------------------------------------------------------------
# Configuration générale
# ----------------------------------------------------------------------------

st.set_page_config(
    page_title="Notation Émetteurs — Dette Privée",
    page_icon="📊",
    layout="wide",
)

NAVY = "#1F2A44"
GOLD = "#C9A227"

RATING_META = {
    "A": {"color": "#16A34A", "bg": "#F0FDF4", "border": "#86EFAC", "label": "Qualité de crédit forte"},
    "B": {"color": "#2563EB", "bg": "#EFF6FF", "border": "#93C5FD", "label": "Qualité de crédit satisfaisante"},
    "C": {"color": "#D97706", "bg": "#FFFBEB", "border": "#FCD34D", "label": "Sous surveillance"},
    "D": {"color": "#DC2626", "bg": "#FEF2F2", "border": "#FCA5A5", "label": "Qualité de crédit dégradée"},
    "N/A": {"color": "#6B7280", "bg": "#F9FAFB", "border": "#D1D5DB", "label": "Données non disponibles"},
}

POINTS = {"A": 4, "B": 3, "C": 2, "D": 1}

st.markdown(
    f"""
    <style>
      .block-container {{ padding-top: 2rem; }}
      .acm-title {{
          text-align: center; font-size: 2.3rem; font-weight: 800;
          color: {NAVY}; margin-bottom: .2rem;
      }}
      .acm-issuer {{ font-size: 1.6rem; font-weight: 700; color: #111827; }}
      .acm-badge {{
          display: inline-block; padding: 2px 10px; border-radius: 8px;
          background: #FEE2E2; color: #B91C1C; font-size: .85rem; font-weight: 600;
          margin-right: 8px;
      }}
      .acm-sector {{ color: #4B5563; font-size: .9rem; }}
      .rating-card {{
          border-radius: 12px; padding: 14px 30px; text-align: center;
          border: 2px solid; min-width: 190px;
      }}
      .rating-letter {{ font-size: 2.2rem; font-weight: 800; line-height: 1.1; }}
      .rating-label {{ font-size: .75rem; color: #374151; }}
      .kpi-label {{ color: #374151; font-size: .9rem; margin-bottom: .1rem; }}
      .kpi-value {{ font-size: 2rem; font-weight: 700; color: #111827; }}
      .kpi-note {{ font-size: .8rem; font-weight: 700; }}
      hr {{ margin: 1.2rem 0; }}
    </style>
    """,
    unsafe_allow_html=True,
)

# ----------------------------------------------------------------------------
# Moteur de notation (conforme au PDF « Méthodologie Notation AFRICAPITAL »)
# ----------------------------------------------------------------------------

def note_solvabilite(x):
    """Fonds propres / total actif (proxy du ratio réglementaire)."""
    if x is None: return "N/A"
    return "A" if x >= 0.10 else "B" if x >= 0.08 else "C" if x >= 0.06 else "D"

def note_exigible_pnb(x):
    """Total exigible / PNB — un levier faible est meilleur (PDF : ≤20x = A)."""
    if x is None: return "N/A"
    return "A" if x <= 20 else "B" if x <= 30 else "C" if x <= 40 else "D"

def note_marge_intermediation(x):
    """Produits d'intérêts / Charges d'intérêts."""
    if x is None: return "N/A"
    return "A" if x >= 2.5 else "B" if x >= 2 else "C" if x >= 1.5 else "D"

def note_gearing(x):
    """Dette nette / Fonds propres."""
    if x is None: return "N/A"
    return "A" if x <= 0.5 else "B" if x <= 0.7 else "C" if x <= 0.9 else "D"

def note_dn_ebitda(x):
    """Dette nette / EBITDA (années de désendettement)."""
    if x is None: return "N/A"
    return "A" if x <= 1 else "B" if x <= 2 else "C" if x <= 4 else "D"

def note_couverture(x):
    """EBITDA / Frais financiers."""
    if x is None: return "N/A"
    return "A" if x >= 15 else "B" if x >= 10 else "C" if x >= 5 else "D"

def note_globale(notes):
    """Moyenne des points (A=4 … D=1) reconvertie en lettre."""
    pts = [POINTS[n] for n in notes if n in POINTS]
    if len(pts) < 3:
        return None, "N/A"
    m = sum(pts) / len(pts)
    lettre = "A" if m >= 3.5 else "B" if m >= 2.5 else "C" if m >= 1.5 else "D"
    return round(m, 2), lettre

def _num(v):
    """Convertit une cellule en float, None si vide/non numérique."""
    if v is None or (isinstance(v, str) and not v.strip()):
        return None
    try:
        f = float(v)
        return None if f != f else f  # exclut les NaN
    except (TypeError, ValueError):
        return None

def _div(a, b):
    if a is None or b in (None, 0):
        return None
    return a / b

# ----------------------------------------------------------------------------
# Lecture robuste du fichier Excel — s'adapte à différentes mises en forme
# (nom des feuilles, ligne d'en-tête, ordre/orthographe des colonnes)
# ----------------------------------------------------------------------------

import re
import unicodedata


def _norm(s):
    """Normalise un libellé : sans accents, minuscules, espaces compactés."""
    if s is None:
        return ""
    s = unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode("ascii")
    s = s.lower().strip()
    s = re.sub(r"[’']", "'", s)
    s = re.sub(r"\s+", " ", s)
    return s


# Alias possibles pour chaque champ requis (on matche par normalisation)
ALIASES = {
    "societes":              ["societes", "societe", "emetteur", "emetteurs"],
    "capitaux_propres":      ["capitaux propres", "capitaux propre"],
    "total_actif":           ["total actif", "total actifs"],
    "dettes_subordonnees":   ["dettes subordonnees", "dette subordonnee"],
    "etablissements_credit": ["etablissements de credits", "etablissement de credit", "etablissements de credit"],
    "dettes_clientele":      ["dettes envers la clientele", "dette envers la clientele"],
    "autres_dettes_titre":   ["autre dettes representees par un titre", "autres dettes representees par un titre"],
    "pnb":                   ["produit net bancaire", "pnb"],
    "produits_interet":      ["produits d'interet", "produit d'interet"],
    "charges_interet":       ["charges d'interets", "charge d'interets", "charges d'interet"],
    "dette_nette":           ["dette nette"],
    "resultat_exploitation": ["resultat d'exploitation", "resultat dexploitation"],
    "ebitda":                ["ebitda"],
    "frais_financiers":      ["charges d'interet (frais financiers)", "frais financiers",
                              "charges d'interets (frais financiers)"],
    "type":                  ["type"],
    "secteur":               ["secteur"],
}


def _map_columns(columns):
    """Associe chaque clé canonique (ALIASES) à son nom de colonne réel, si présent."""
    norm_map = {c: _norm(c) for c in columns}
    mapping = {}
    for key, aliases in ALIASES.items():
        aliases_norm = {_norm(a) for a in aliases}
        for orig, norm in norm_map.items():
            if norm in aliases_norm:
                mapping[key] = orig
                break
    return mapping


def _detect_header_row(ws, max_scan=5):
    """Trouve la ligne (1-indexée) contenant l'en-tête 'Societes'/'Emetteur'."""
    societes_aliases = {_norm(a) for a in ALIASES["societes"]}
    for r in range(1, max_scan + 1):
        for cell in ws[r]:
            if _norm(cell.value) in societes_aliases:
                return r
    return 1  # repli : première ligne


def _find_sheet(sheetnames, keywords):
    """Trouve la première feuille dont le nom normalisé contient un des mots-clés."""
    for name in sheetnames:
        n = _norm(name)
        if any(k in n for k in keywords):
            return name
    return None


@st.cache_data(show_spinner=False)
def charger_donnees(contenu: bytes):
    import openpyxl

    wb = openpyxl.load_workbook(io.BytesIO(contenu), read_only=True, data_only=True)
    sheetnames = wb.sheetnames

    sheet_fin = _find_sheet(sheetnames, ["financ"])
    sheet_corp = _find_sheet(sheetnames, ["corporate", "corpo"])
    sheet_univers = _find_sheet(sheetnames, ["feuil1", "univers", "liste"])

    if sheet_fin is None or sheet_corp is None:
        raise ValueError(
            "Impossible de repérer les feuilles 'financier' et 'corporate' dans ce fichier. "
            f"Feuilles trouvées : {sheetnames}"
        )

    xls = pd.ExcelFile(io.BytesIO(contenu))

    # --- Univers (optionnel) ----------------------------------------------
    univers = None
    if sheet_univers is not None:
        header_u = _detect_header_row(wb[sheet_univers])
        raw_u = pd.read_excel(xls, sheet_univers, header=header_u - 1)
        raw_u.columns = [str(c).strip() for c in raw_u.columns]
        cmap_u = _map_columns(raw_u.columns)
        if "societes" in cmap_u:
            univers = pd.DataFrame({
                "Emetteur": raw_u[cmap_u["societes"]].astype(str).str.strip(),
                "Type": raw_u[cmap_u["type"]] if "type" in cmap_u else None,
                "Secteur": raw_u[cmap_u["secteur"]] if "secteur" in cmap_u else None,
            }).dropna(subset=["Emetteur"])
            univers = univers[univers["Emetteur"].str.len() > 0]

    # --- Sociétés financières ----------------------------------------------
    header_f = _detect_header_row(wb[sheet_fin])
    fin_raw = pd.read_excel(xls, sheet_fin, header=header_f - 1)
    fin_raw.columns = [str(c).strip() for c in fin_raw.columns]
    fin_raw = fin_raw.loc[:, ~fin_raw.columns.str.startswith("Unnamed")]
    cmap_f = _map_columns(fin_raw.columns)

    def g(row, key):
        col = cmap_f.get(key)
        return _num(row.get(col)) if col else None

    fin_rows = []
    for _, r in fin_raw.iterrows():
        nom = r.get(cmap_f.get("societes"))
        if not isinstance(nom, str) or not nom.strip():
            continue
        cp   = g(r, "capitaux_propres")
        ta   = g(r, "total_actif")
        dsub = g(r, "dettes_subordonnees")
        ec   = g(r, "etablissements_credit")
        dcl  = g(r, "dettes_clientele")
        adt  = g(r, "autres_dettes_titre")
        pnb  = g(r, "pnb")
        pi   = g(r, "produits_interet")
        ci   = g(r, "charges_interet")

        exigible = None
        parts = [dsub, ec, dcl, adt]
        if any(p is not None for p in parts):
            exigible = sum(p for p in parts if p is not None)

        r1 = _div(cp, ta)          # ratio de solvabilité (proxy)
        r2 = _div(exigible, pnb)   # total exigible / PNB
        r3 = _div(pi, ci)          # produits / charges d'intérêts

        n1, n2, n3 = note_solvabilite(r1), note_exigible_pnb(r2), note_marge_intermediation(r3)
        moy, lettre = note_globale([n1, n2, n3])

        fin_rows.append({
            "Emetteur": nom.strip(), "Famille": "Société financière",
            "Ratio 1": r1, "Ratio 2": r2, "Ratio 3": r3,
            "Note 1": n1, "Note 2": n2, "Note 3": n3,
            "Score": moy, "Note finale": lettre,
        })
    df_fin = pd.DataFrame(fin_rows)

    # --- Corporates ----------------------------------------------------------
    header_c = _detect_header_row(wb[sheet_corp])
    corp_raw = pd.read_excel(xls, sheet_corp, header=header_c - 1)
    corp_raw.columns = [str(c).strip() for c in corp_raw.columns]
    corp_raw = corp_raw.loc[:, ~corp_raw.columns.str.startswith("Unnamed")]
    cmap_c = _map_columns(corp_raw.columns)

    def gc(row, key):
        col = cmap_c.get(key)
        return _num(row.get(col)) if col else None

    corp_rows = []
    for _, r in corp_raw.iterrows():
        nom = r.get(cmap_c.get("societes"))
        if not isinstance(nom, str) or not nom.strip():
            continue
        cp     = gc(r, "capitaux_propres")
        dn     = gc(r, "dette_nette")
        rex    = gc(r, "resultat_exploitation")
        ebitda = gc(r, "ebitda")
        ff     = gc(r, "frais_financiers")

        # EBITDA de repli : résultat d'exploitation si l'EBITDA manque
        if ebitda is None:
            ebitda = rex

        r1 = _div(dn, cp)        # gearing
        r2 = _div(dn, ebitda)    # dette nette / EBITDA
        r3 = _div(ebitda, ff)    # EBITDA / frais financiers

        n1, n2, n3 = note_gearing(r1), note_dn_ebitda(r2), note_couverture(r3)
        moy, lettre = note_globale([n1, n2, n3])

        corp_rows.append({
            "Emetteur": nom.strip(), "Famille": "Corporate",
            "Ratio 1": r1, "Ratio 2": r2, "Ratio 3": r3,
            "Note 1": n1, "Note 2": n2, "Note 3": n3,
            "Score": moy, "Note finale": lettre,
        })
    df_corp = pd.DataFrame(corp_rows)

    # --- Univers de repli : si aucune feuille Type/Secteur n'existe --------
    if univers is None:
        combo = pd.concat([df_fin[["Emetteur", "Famille"]], df_corp[["Emetteur", "Famille"]]], ignore_index=True)
        univers = pd.DataFrame({
            "Emetteur": combo["Emetteur"],
            "Type": combo["Famille"],
            "Secteur": "—",
        })

    return univers, df_fin, df_corp


# Libellés des indicateurs par famille
LIBELLES = {
    "Société financière": [
        ("Ratio de solvabilité", "pct"),
        ("Total exigible / PNB", "x"),
        ("Produits d'intérêts / Charges d'intérêts", "x"),
    ],
    "Corporate": [
        ("Gearing (Dette nette / Fonds propres)", "x"),
        ("Dette nette / EBITDA", "x"),
        ("EBITDA / Frais financiers", "x"),
    ],
}

def fmt(v, kind):
    if v is None or pd.isna(v):
        return "—"
    if kind == "pct":
        return f"{v * 100:.1f}%".replace(".", ",")
    return f"{v:.1f}x".replace(".", ",")

# ----------------------------------------------------------------------------
# Interface
# ----------------------------------------------------------------------------

with st.sidebar:
    st.header("Sélection")
    fichier = st.file_uploader("Fichier Excel des émetteurs", type=["xlsx"])

NOMS_PAR_DEFAUT = ["Data-Emetteurs.xlsx", "Liste_Emetteurs.xlsx"]
defaut = next((Path(__file__).parent / n for n in NOMS_PAR_DEFAUT if (Path(__file__).parent / n).exists()), None)

if fichier is not None:
    contenu = fichier.getvalue()
elif defaut is not None:
    contenu = defaut.read_bytes()
else:
    st.info("Charge un fichier Excel des émetteurs dans la barre latérale pour commencer.")
    st.stop()

try:
    univers, df_fin, df_corp = charger_donnees(contenu)
except ValueError as e:
    st.error(str(e))
    st.stop()
resultats = pd.concat([df_fin, df_corp], ignore_index=True)

# Jointure avec l'univers pour récupérer Type et Secteur
univers_idx = univers.set_index(univers["Emetteur"].str.upper().str.strip())

def info_univers(nom):
    key = nom.upper().strip()
    # correspondance exacte puis partielle (ex. "CIH" vs "CIH Bank")
    if key in univers_idx.index:
        row = univers_idx.loc[key]
    else:
        match = univers_idx[univers_idx.index.str.contains(key, regex=False)]
        if match.empty:
            return "—", "—"
        row = match.iloc[0]
    if isinstance(row, pd.DataFrame):
        row = row.iloc[0]
    return row.get("Type", "—"), row.get("Secteur", "—")

resultats[["Type", "Secteur"]] = resultats["Emetteur"].apply(
    lambda n: pd.Series(info_univers(n))
)

with st.sidebar:
    types_dispo = sorted(univers["Type"].dropna().unique().tolist())
    types_sel = st.multiselect("Type d'émetteur", types_dispo, default=types_dispo)

    dispo = resultats[resultats["Type"].isin(types_sel)] if types_sel else resultats
    notes_ok = dispo[dispo["Note finale"] != "N/A"]
    choix = notes_ok["Emetteur"].tolist() or dispo["Emetteur"].tolist()
    emetteur = st.selectbox("Émetteur", choix) if choix else None

st.markdown('<div class="acm-title">Notation AFRICAPITAL des émetteurs de dette privée</div>', unsafe_allow_html=True)
st.markdown("---")

if emetteur is None:
    st.warning("Aucun émetteur disponible pour cette sélection.")
    st.stop()

ligne = resultats[resultats["Emetteur"] == emetteur].iloc[0]
note = ligne["Note finale"] if pd.notna(ligne["Note finale"]) else "N/A"
meta = RATING_META[note]

# --- En-tête émetteur + carte de notation -----------------------------------
c1, c2 = st.columns([3, 1])
with c1:
    st.markdown(f'<div class="acm-issuer">{emetteur}</div>', unsafe_allow_html=True)
    st.markdown(
        f'<span class="acm-badge">{ligne["Type"]}</span>'
        f'<span class="acm-sector">Secteur : {ligne["Secteur"]}</span>',
        unsafe_allow_html=True,
    )
with c2:
    st.markdown(
        f"""
        <div class="rating-card" style="background:{meta['bg']}; border-color:{meta['border']};">
          <div class="rating-letter" style="color:{meta['color']};">{note}</div>
          <div class="rating-label">{meta['label']}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

if ligne["Score"] is not None and pd.notna(ligne["Score"]):
    st.caption(f"Score moyen : {str(ligne['Score']).replace('.', ',')} / 4")

# --- Indicateurs clés --------------------------------------------------------
st.markdown("### Indicateurs clés")
libelles = LIBELLES[ligne["Famille"]]
fmt1 = "pct" if ligne["Famille"] == "Société financière" else libelles[0][1]

cols = st.columns(3)
for col, (lib, kind), rk, nk in zip(
    cols, libelles, ["Ratio 1", "Ratio 2", "Ratio 3"], ["Note 1", "Note 2", "Note 3"]
):
    n = ligne[nk]
    couleur = RATING_META.get(n, RATING_META["N/A"])["color"]
    with col:
        st.markdown(
            f"""
            <div class="kpi-label">{lib}</div>
            <div class="kpi-value">{fmt(ligne[rk], kind)}</div>
            <div class="kpi-note" style="color:{couleur};">Note : {n}</div>
            """,
            unsafe_allow_html=True,
        )

st.markdown("---")

# --- Univers des émetteurs notés --------------------------------------------
st.markdown("### Univers des émetteurs notés")

tableau = resultats[resultats["Note finale"] != "N/A"][
    ["Emetteur", "Type", "Secteur", "Note 1", "Note 2", "Note 3", "Score", "Note finale"]
].rename(columns={"Emetteur": "Émetteur", "Note finale": "Note"})

def couleur_note(v):
    meta = RATING_META.get(v)
    if meta:
        return f"color:{meta['color']}; font-weight:700;"
    return ""

st.dataframe(
    tableau.style.map(couleur_note, subset=["Note 1", "Note 2", "Note 3", "Note"]),
    use_container_width=True,
    hide_index=True,
)

# --- Export Excel ------------------------------------------------------------
buffer = io.BytesIO()
export = resultats[resultats["Note finale"] != "N/A"].copy()
with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
    export.to_excel(writer, sheet_name="Notation", index=False)
    wb, ws = writer.book, writer.sheets["Notation"]
    header_fmt = wb.add_format({"bold": True, "bg_color": NAVY, "font_color": "white", "border": 1})
    for i, col in enumerate(export.columns):
        ws.write(0, i, col, header_fmt)
        ws.set_column(i, i, max(14, len(str(col)) + 2))

st.download_button(
    "📥 Exporter la notation (Excel)",
    data=buffer.getvalue(),
    file_name="Notation_AFRICAPITAL.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

st.caption(
    "Modèle d'aide à la décision — ne se substitue ni à l'analyse fondamentale "
    "ni aux éléments qualitatifs (gouvernance, actionnariat, position concurrentielle)."
)
