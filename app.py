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
# Lecture du fichier Excel
# ----------------------------------------------------------------------------

@st.cache_data(show_spinner=False)
def charger_donnees(contenu: bytes):
    xls = pd.ExcelFile(io.BytesIO(contenu))

    # --- Univers (Feuil1) -------------------------------------------------
    univers = pd.read_excel(xls, "Feuil1")
    univers.columns = [c.strip() for c in univers.columns]
    univers = univers.dropna(subset=["Emetteur"])

    # --- Sociétés financières --------------------------------------------
    fin_raw = pd.read_excel(xls, "financier", header=1)
    fin_raw = fin_raw.loc[:, ~fin_raw.columns.astype(str).str.startswith("Unnamed")]
    fin_raw.columns = [str(c).strip() for c in fin_raw.columns]

    fin_rows = []
    for _, r in fin_raw.iterrows():
        nom = r.get("Societes")
        if not isinstance(nom, str) or not nom.strip():
            continue
        cp   = _num(r.get("Capitaux propre"))
        ta   = _num(r.get("Total actif"))
        dsub = _num(r.get("Dettes subordonnees"))
        ec   = _num(r.get("Etablissements de credits"))
        dcl  = _num(r.get("dettes envers la clientele"))
        adt  = _num(r.get("autre dettes representees par un titre"))
        pnb  = _num(r.get("Produit net bancaire"))
        pi   = _num(r.get("Produits d'interet"))
        ci   = _num(r.get("charges d'interets"))

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

    # --- Corporates -------------------------------------------------------
    corp_raw = pd.read_excel(xls, "corporate", header=1)
    corp_raw = corp_raw.loc[:, ~corp_raw.columns.astype(str).str.startswith("Unnamed")]
    corp_raw.columns = [str(c).strip() for c in corp_raw.columns]

    corp_rows = []
    for _, r in corp_raw.iterrows():
        nom = r.get("Societes")
        if not isinstance(nom, str) or not nom.strip():
            continue
        cp     = _num(r.get("Capitaux propres"))
        dn     = _num(r.get("Dette nette"))
        rex    = _num(r.get("Resultat d'exploitation"))
        ebitda = _num(r.get("EBITDA"))
        ff     = _num(r.get("Charges d'interet (Frais financiers)"))

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
    fichier = st.file_uploader("Fichier Excel (Liste_Emetteurs.xlsx)", type=["xlsx"])

defaut = Path(__file__).parent / "Liste_Emetteurs.xlsx"
if fichier is not None:
    contenu = fichier.getvalue()
elif defaut.exists():
    contenu = defaut.read_bytes()
else:
    st.info("Charge le fichier Liste_Emetteurs.xlsx dans la barre latérale pour commencer.")
    st.stop()

univers, df_fin, df_corp = charger_donnees(contenu)
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
