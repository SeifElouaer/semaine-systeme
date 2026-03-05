"""
=============================================================================
INTERFACE STREAMLIT — Ballet des Transporteurs GWD
=============================================================================
Lance l'interface avec :
    streamlit run app_ballet.py

Ce fichier importe modele.py sans le modifier.
Il doit être placé dans le même dossier que :
  • modele.py
  • Données GWD 2026 VF.xlsx
=============================================================================
"""

import sys
import os
import importlib

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px

# ---------------------------------------------------------------------------
# Import du module solver (sans le modifier)
# ---------------------------------------------------------------------------
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
solver = importlib.import_module("modele")

# ---------------------------------------------------------------------------
# Config page
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="Ballet des Transporteurs — GWD",
    page_icon="🚛",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ---------------------------------------------------------------------------
# CSS
# ---------------------------------------------------------------------------
st.markdown("""
<style>
    .main-title {
        font-size: 2rem; font-weight: 700; color: #1a5276;
        border-bottom: 3px solid #2e86c1; padding-bottom: 0.4rem;
        margin-bottom: 1.2rem;
    }
    .section-title {
        font-size: 1.15rem; font-weight: 600; color: #1a5276;
        margin-top: 1.5rem; margin-bottom: 0.5rem;
    }
    .kpi-box {
        background: #eaf4fc; border-left: 5px solid #2e86c1;
        border-radius: 6px; padding: 0.7rem 1rem;
        margin-bottom: 0.5rem;
    }
    .kpi-label { font-size: 0.78rem; color: #666; text-transform: uppercase; }
    .kpi-value { font-size: 1.6rem; font-weight: 700; color: #1a5276; }
    .selected-badge {
        display:inline-block; background:#27ae60; color:white;
        border-radius:4px; padding:2px 8px; font-size:0.8rem; margin:2px;
    }
    .excluded-badge {
        display:inline-block; background:#e74c3c; color:white;
        border-radius:4px; padding:2px 8px; font-size:0.8rem; margin:2px;
    }
</style>
""", unsafe_allow_html=True)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def min_to_hhmm(m: float, base_h: int = 18) -> str:
    total = int(round(m)) + base_h * 60
    return f"{(total // 60) % 24:02d}h{total % 60:02d}"

def min_to_abs(m: float, base_h: int = 18) -> float:
    """Minutes depuis base_h → minutes absolues depuis minuit (pour Plotly)."""
    return (m + base_h * 60) % (24 * 60)

def abs_to_hhmm(m_abs: float) -> str:
    h = int(m_abs) // 60 % 24
    mn = int(m_abs) % 60
    return f"{h:02d}h{mn:02d}"

COLORS_TRANS = px.colors.qualitative.Set2


# ---------------------------------------------------------------------------
# Chargement des données Excel
# ---------------------------------------------------------------------------
@st.cache_data(show_spinner="Chargement du fichier Excel…")
def charger_donnees(uploaded_file_bytes, filename):
    """Charge les données depuis l'Excel en utilisant le module solver."""
    import tempfile, openpyxl, unicodedata

    def normalise(s):
        return unicodedata.normalize("NFD", s).encode("ascii", "ignore").decode().lower()

    # Écriture temporaire si fichier uploadé, sinon cherche en local
    if uploaded_file_bytes is not None:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
            tmp.write(uploaded_file_bytes)
            tmp_path = tmp.name
    else:
        script_dir = os.path.dirname(os.path.abspath(__file__))
        candidates = [
            os.path.join(script_dir, solver.FICHIER_EXCEL),
            os.path.join(os.getcwd(), solver.FICHIER_EXCEL),
        ]
        tmp_path = next((p for p in candidates if os.path.isfile(p)), None)
        if tmp_path is None:
            return None, None, f"Fichier '{solver.FICHIER_EXCEL}' introuvable."

    try:
        wb = openpyxl.load_workbook(tmp_path, data_only=True)
        feuille = next(
            (name for name in wb.sheetnames if "transporteur" in normalise(name)), None
        )
        if feuille is None:
            return None, None, f"Feuille 'Données Transporteurs' introuvable. Feuilles : {wb.sheetnames}"

        ws   = wb[feuille]
        rows = [list(r) for r in ws.iter_rows(values_only=True)]

        transporteurs = {}
        camions       = {}
        camion_vers_trans = {}

        idx_trans = idx_camion = None
        for i, row in enumerate(rows):
            if row[0] is None: continue
            cell = str(row[0]).strip().lower()
            if cell == "transporteur" and idx_trans is None:   idx_trans  = i
            elif cell == "camion"      and idx_camion is None: idx_camion = i

        for row in rows[idx_trans + 1 : idx_camion]:
            if row[0] is None: continue
            nom, score, pen = str(row[0]).strip(), row[1], row[2]
            cstr = str(row[3]).strip() if row[3] else ""
            if not nom or not isinstance(score, (int, float)): continue
            transporteurs[nom] = {"score": int(score), "penalite": int(pen)}
            for c in [c.strip() for c in cstr.split(",") if c.strip()]:
                camion_vers_trans[c] = nom

        for row in rows[idx_camion + 1:]:
            if row[0] is None: continue
            cid = str(row[0]).strip()
            if not cid: continue
            arr, dur, lim = row[1], row[2], row[3]
            if arr is None or dur is None or lim is None: continue
            trans = camion_vers_trans.get(cid)
            if trans is None: continue
            camions[cid] = {
                "transporteur": trans,
                "S":  transporteurs[trans]["score"],
                "P":  transporteurs[trans]["penalite"],
                "Ta": solver.heure_to_min(arr),
                "Td": int(dur),
                "Tl": solver.heure_to_min(lim),
            }

        return transporteurs, camions, None
    except Exception as e:
        return None, None, str(e)
    finally:
        if uploaded_file_bytes is not None and os.path.exists(tmp_path):
            os.unlink(tmp_path)


# ---------------------------------------------------------------------------
# Résolution
# ---------------------------------------------------------------------------
@st.cache_data(show_spinner="Résolution en cours (CBC)…")
def resoudre(_transporteurs, _camions, K, Q, M=1440):
    """Lance le solveur et retourne les résultats sous forme de dicts."""
    try:
        import pulp
    except ImportError:
        return None, "PuLP non installé. Exécutez : pip install pulp"

    # Injecter les données dans le module solver
    solver.TRANSPORTEURS.clear()
    solver.TRANSPORTEURS.update(_transporteurs)
    solver.CAMIONS.clear()
    solver.CAMIONS.update(_camions)
    solver.N = len(_camions)
    solver.K = K
    solver.Q = Q
    solver.M = M

    try:
        prob, u, x, d, r, y, camions, quais = solver.construire_modele()
    except Exception as e:
        return None, str(e)

    prob.solve(pulp.PULP_CBC_CMD(msg=0, timeLimit=300))

    if prob.status != 1:
        return None, f"Pas de solution optimale (statut : {pulp.LpStatus[prob.status]})"

    # Extraction des résultats
    Z          = pulp.value(prob.objective)
    sel        = [i for i in camions if pulp.value(u[i]) > 0.5]
    excl       = [i for i in camions if pulp.value(u[i]) < 0.5]
    score_eco  = sum(_camions[i]["S"] * pulp.value(u[i]) for i in camions)
    penalites  = sum(_camions[i]["P"] * pulp.value(r[i]) for i in camions)
    att_tot    = sum(pulp.value(d[i]) - _camions[i]["Ta"] * pulp.value(u[i]) for i in camions)

    rows_result = []
    for i in sel:
        Ta_i    = _camions[i]["Ta"]
        di      = pulp.value(d[i])
        ri      = pulp.value(r[i])
        quai_i  = next(j for j in quais if pulp.value(x[i, j]) > 0.5)
        attente = max(0.0, di - Ta_i)
        rows_result.append({
            "Camion":       i,
            "Transporteur": _camions[i]["transporteur"],
            "Éco":          _camions[i]["S"],
            "Quai":         quai_i,
            "Ta_min":       Ta_i,
            "d_min":        di,
            "Fin_min":      di + _camions[i]["Td"],
            "Tl_min":       _camions[i]["Tl"],
            "Td_min":       _camions[i]["Td"],
            "Attente_min":  attente,
            "Retard_min":   ri,
            "Pénalité_TND": _camions[i]["P"] * ri,
            "Arrivée":      min_to_hhmm(Ta_i),
            "Début":        min_to_hhmm(di),
            "Fin":          min_to_hhmm(di + _camions[i]["Td"]),
            "Limite":       min_to_hhmm(_camions[i]["Tl"]),
        })

    return {
        "Z": Z, "sel": sel, "excl": excl,
        "score_eco": score_eco, "penalites": penalites, "att_tot": att_tot,
        "rows": rows_result,
        "n_vars": len(prob.variables()),
        "n_cons": len(prob.constraints),
    }, None


# ---------------------------------------------------------------------------
# Graphiques
# ---------------------------------------------------------------------------

def gantt_chart(rows, quais):
    """Diagramme de Gantt par quai."""
    trans_list = list({r["Transporteur"] for r in rows})
    color_map  = {t: COLORS_TRANS[i % len(COLORS_TRANS)] for i, t in enumerate(trans_list)}

    fig = go.Figure()
    base_h = 18

    for row in rows:
        debut_abs = (row["d_min"] + base_h * 60) % (24 * 60)
        fin_abs   = (row["Fin_min"] + base_h * 60) % (24 * 60)
        quai_y    = f"Quai {row['Quai']}"
        retard    = row["Retard_min"]
        attente   = row["Attente_min"]
        color     = color_map[row["Transporteur"]]

        # Barre d'attente (gris clair)
        if attente > 0:
            arr_abs = (row["Ta_min"] + base_h * 60) % (24 * 60)
            fig.add_trace(go.Bar(
                x=[attente], y=[quai_y], base=arr_abs,
                orientation="h",
                marker_color="rgba(180,180,180,0.4)",
                marker_line_width=0,
                name="Attente",
                showlegend=False,
                hovertemplate=(
                    f"<b>{row['Camion']}</b> — Attente<br>"
                    f"Arrivée : {row['Arrivée']}<br>"
                    f"Début   : {row['Début']}<br>"
                    f"Attente : {int(attente)}m<extra></extra>"
                ),
            ))

        # Barre de chargement
        fig.add_trace(go.Bar(
            x=[row["Td_min"]], y=[quai_y], base=debut_abs,
            orientation="h",
            marker_color=color,
            marker_line_color="white", marker_line_width=1.5,
            name=row["Transporteur"],
            legendgroup=row["Transporteur"],
            showlegend=row["Transporteur"] not in [t.name for t in fig.data if hasattr(t, "legendgroup") and t.legendgroup == row["Transporteur"] and t.showlegend],
            hovertemplate=(
                f"<b>{row['Camion']}</b> ({row['Transporteur']})<br>"
                f"Début   : {row['Début']}<br>"
                f"Fin     : {row['Fin']}<br>"
                f"Durée   : {int(row['Td_min'])}m<br>"
                f"Retard  : {int(retard)}m<br>"
                f"Pénalité: {int(row['Pénalité_TND'])} ₮<extra></extra>"
            ),
        ))

        # Label camion au centre de la barre
        centre = debut_abs + row["Td_min"] / 2
        fig.add_annotation(
            x=centre, y=quai_y,
            text=f"<b>{row['Camion']}</b>",
            showarrow=False, font=dict(size=11, color="white"),
            xanchor="center", yanchor="middle",
        )

        # Trait vertical rouge pour la limite
        lim_abs = (row["Tl_min"] + base_h * 60) % (24 * 60)
        if retard > 0:
            fig.add_shape(type="line",
                x0=lim_abs, x1=lim_abs,
                y0=quai_y, y1=quai_y,
                yref="y", xref="x",
                line=dict(color="red", width=2, dash="dot"),
            )

    # Ticks toutes les heures
    tick_vals = list(range(0, 1440, 60))
    tick_text = [f"{h:02d}h00" for h in range(24)]

    fig.update_layout(
        barmode="overlay",
        title="📅 Diagramme de Gantt — Ordonnancement sur les quais",
        xaxis=dict(
            title="Heure",
            tickvals=tick_vals, ticktext=tick_text,
            range=[
                min((r["d_min"] + base_h * 60 - 30) % 1440 for r in rows),
                max((r["Fin_min"] + base_h * 60 + 60) % 1440 for r in rows),
            ],
            gridcolor="#eee",
        ),
        yaxis=dict(title="Quai", categoryorder="array",
                   categoryarray=[f"Quai {j}" for j in sorted(quais)]),
        height=300 + len(quais) * 80,
        plot_bgcolor="white",
        legend=dict(title="Transporteur", orientation="h",
                    y=-0.25, x=0.5, xanchor="center"),
        margin=dict(l=80, r=30, t=60, b=120),
    )
    return fig


def bar_scores(rows):
    """Barres : score éco par camion sélectionné, coloré par transporteur."""
    df = pd.DataFrame(rows).sort_values("Éco", ascending=True)
    fig = px.bar(
        df, x="Éco", y="Camion", orientation="h",
        color="Transporteur", text="Éco",
        color_discrete_sequence=COLORS_TRANS,
        title="🌿 Score écologique par camion sélectionné",
    )
    fig.update_traces(textposition="outside")
    fig.update_layout(
        height=380, plot_bgcolor="white",
        xaxis=dict(range=[0, 110], gridcolor="#eee"),
        legend=dict(orientation="h", y=-0.2, x=0.5, xanchor="center"),
        margin=dict(l=60, r=30, t=50, b=80),
    )
    return fig


def bar_penalites(rows):
    """Barres : pénalités par camion."""
    df = pd.DataFrame(rows)
    df["Retard_label"] = df["Retard_min"].apply(lambda v: f"{int(v)}m")
    fig = px.bar(
        df, x="Camion", y="Pénalité_TND",
        color="Transporteur", text="Retard_label",
        color_discrete_sequence=COLORS_TRANS,
        title="💰 Pénalités de retard par camion (TND)",
    )
    fig.update_traces(textposition="outside")
    fig.update_layout(
        height=360, plot_bgcolor="white",
        yaxis=dict(gridcolor="#eee"),
        legend=dict(orientation="h", y=-0.2, x=0.5, xanchor="center"),
        margin=dict(l=60, r=30, t=50, b=80),
    )
    return fig


def pie_attente(rows):
    """Camembert : répartition du temps d'attente par camion."""
    df = pd.DataFrame(rows)
    df_att = df[df["Attente_min"] > 0]
    if df_att.empty:
        return None
    fig = px.pie(
        df_att, values="Attente_min", names="Camion",
        title="⏳ Répartition du temps d'attente",
        color_discrete_sequence=COLORS_TRANS,
        hole=0.4,
    )
    fig.update_traces(textinfo="label+percent")
    fig.update_layout(height=360, margin=dict(l=30, r=30, t=50, b=30))
    return fig


def scatter_arrivee_debut(rows):
    """Scatter : heure d'arrivée vs heure de début de chargement."""
    df = pd.DataFrame(rows)
    df["Arrivée_abs"] = df["Ta_min"].apply(lambda m: (m + 18*60) % 1440)
    df["Début_abs"]   = df["d_min"].apply(lambda m: (m + 18*60) % 1440)

    fig = px.scatter(
        df, x="Arrivée_abs", y="Début_abs",
        color="Transporteur", size="Td_min",
        text="Camion",
        color_discrete_sequence=COLORS_TRANS,
        title="🕐 Arrivée vs Début de chargement",
        labels={"Arrivée_abs": "Heure d'arrivée", "Début_abs": "Heure de début"},
    )

    # Ligne y=x (pas d'attente)
    mn = min(df["Arrivée_abs"].min(), df["Début_abs"].min()) - 30
    mx = max(df["Arrivée_abs"].max(), df["Début_abs"].max()) + 30
    fig.add_shape(type="line", x0=mn, y0=mn, x1=mx, y1=mx,
                  line=dict(color="grey", dash="dash", width=1))

    ticks  = list(range(0, 1440, 120))
    labels = [f"{h:02d}h00" for h in range(0, 24, 2)]
    fig.update_layout(
        height=380, plot_bgcolor="white",
        xaxis=dict(tickvals=ticks, ticktext=labels, gridcolor="#eee"),
        yaxis=dict(tickvals=ticks, ticktext=labels, gridcolor="#eee"),
        legend=dict(orientation="h", y=-0.25, x=0.5, xanchor="center"),
        margin=dict(l=60, r=30, t=50, b=100),
    )
    fig.update_traces(textposition="top center", textfont_size=10)
    return fig


# ---------------------------------------------------------------------------
# Interface principale
# ---------------------------------------------------------------------------

def main():
    # --- Sidebar ---
    with st.sidebar:
        st.image("https://img.icons8.com/fluency/96/truck.png", width=64)
        st.markdown("## ⚙️ Configuration")
        st.markdown("---")

        uploaded = st.file_uploader(
            "📂 Fichier Excel (.xlsx)",
            type=["xlsx"],
            help=f"Par défaut : '{solver.FICHIER_EXCEL}' (même dossier)"
        )

        st.markdown("**Paramètres du modèle**")
        Q_val = st.number_input("Nombre de quais (Q)", min_value=1, max_value=10, value=2)
        K_auto = st.checkbox("K automatique (N − 3)", value=True)

        uploaded_bytes = uploaded.read() if uploaded else None

        # Chargement préalable pour connaître N
        transporteurs, camions, err_load = charger_donnees(uploaded_bytes, getattr(uploaded, "name", ""))
        if transporteurs and camions:
            N_val = len(camions)
            if K_auto:
                K_val = max(1, N_val - 3)
                st.info(f"K = {N_val} − 3 = **{K_val}** camions sélectionnés")
            else:
                K_val = st.number_input("Nombre de camions à sélectionner (K)",
                                        min_value=1, max_value=N_val, value=max(1, N_val-3))
        else:
            K_val, N_val = 9, 12

        st.markdown("---")
        run_btn = st.button("▶ Lancer l'optimisation", use_container_width=True, type="primary")

    # --- Header ---
    st.markdown('<div class="main-title">🚛 Ballet des Transporteurs — Green Wood Design</div>',
                unsafe_allow_html=True)
    st.caption("Optimisation PLNM · Sélection · Affectation · Ordonnancement sur quais de chargement")

    # --- Gestion des erreurs de chargement ---
    if err_load:
        st.error(f"❌ Erreur de chargement : {err_load}")
        st.stop()

    if transporteurs is None:
        st.warning("👈 Veuillez fournir un fichier Excel ou placer "
                   f"'{solver.FICHIER_EXCEL}' dans le même dossier que ce script.")
        st.stop()

    # --- Aperçu des données ---
    with st.expander("📊 Aperçu des données Excel", expanded=False):
        col1, col2 = st.columns(2)
        with col1:
            st.markdown("**Transporteurs**")
            df_trans = pd.DataFrame([
                {"Transporteur": k, "Score Éco": v["score"], "Pénalité (TND/min)": v["penalite"]}
                for k, v in transporteurs.items()
            ])
            st.dataframe(df_trans, use_container_width=True, hide_index=True)
        with col2:
            st.markdown("**Camions**")
            df_cam = pd.DataFrame([
                {
                    "Camion": k, "Transporteur": v["transporteur"],
                    "Arrivée": min_to_hhmm(v["Ta"]),
                    "Durée (min)": v["Td"],
                    "Limite": min_to_hhmm(v["Tl"]),
                }
                for k, v in camions.items()
            ])
            st.dataframe(df_cam, use_container_width=True, hide_index=True)

    # --- Lancement de l'optimisation ---
    if run_btn or "result" in st.session_state:

        if run_btn:
            with st.spinner("Résolution en cours…"):
                result, err_solve = resoudre(transporteurs, camions, K_val, Q_val)
            if err_solve:
                st.error(f"❌ {err_solve}")
                st.stop()
            st.session_state["result"] = result
        else:
            result = st.session_state["result"]

        rows  = result["rows"]
        quais = sorted({r["Quai"] for r in rows})
        df    = pd.DataFrame(rows)

        # --- KPIs ---
        st.markdown("---")
        st.markdown('<div class="section-title">📈 Indicateurs clés</div>', unsafe_allow_html=True)
        c1, c2, c3, c4, c5 = st.columns(5)
        def kpi(col, label, value, suffix=""):
            col.markdown(
                f'<div class="kpi-box"><div class="kpi-label">{label}</div>'
                f'<div class="kpi-value">{value}{suffix}</div></div>',
                unsafe_allow_html=True
            )
        kpi(c1, "Valeur Z*",         f"{result['Z']:.0f}")
        kpi(c2, "Score éco total",   f"{result['score_eco']:.0f}")
        kpi(c3, "Pénalités retard",  f"{result['penalites']:.0f}", " TND")
        kpi(c4, "Temps attente tot", f"{result['att_tot']:.0f}",   " min")
        kpi(c5, "Camions retenus",   f"{len(result['sel'])}/{len(camions)}")

        # --- Camions sélectionnés / exclus ---
        st.markdown('<div class="section-title">✅ Sélection des camions</div>', unsafe_allow_html=True)
        sel_html  = " ".join(f'<span class="selected-badge">{c}</span>' for c in result["sel"])
        excl_html = " ".join(f'<span class="excluded-badge">{c}</span>'  for c in result["excl"])
        st.markdown(f"**Sélectionnés :** {sel_html}", unsafe_allow_html=True)
        st.markdown(f"**Exclus :** {excl_html}", unsafe_allow_html=True)

        # --- Gantt ---
        st.markdown("---")
        st.plotly_chart(gantt_chart(rows, quais), use_container_width=True)

        # --- Tableaux détaillés par quai ---
        st.markdown('<div class="section-title">📋 Détail par quai</div>', unsafe_allow_html=True)
        tabs = st.tabs([f"Quai {j}" for j in quais])
        for tab, j in zip(tabs, quais):
            with tab:
                df_q = df[df["Quai"] == j].sort_values("d_min")[[
                    "Camion","Transporteur","Éco",
                    "Arrivée","Début","Attente_min","Fin","Limite","Retard_min","Pénalité_TND"
                ]].rename(columns={
                    "Attente_min":   "Attente (m)",
                    "Retard_min":    "Retard (m)",
                    "Pénalité_TND":  "Pénalité (₮)",
                })
                df_q["Attente (m)"]  = df_q["Attente (m)"].apply(lambda v: f"{int(v)}m")
                df_q["Retard (m)"]   = df_q["Retard (m)"].apply(lambda v: f"{int(v)}m")
                df_q["Pénalité (₮)"] = df_q["Pénalité (₮)"].apply(lambda v: f"{int(v)} ₮")
                st.dataframe(df_q, use_container_width=True, hide_index=True)

        # --- Graphiques analytiques ---
        st.markdown("---")
        st.markdown('<div class="section-title">📊 Analyses</div>', unsafe_allow_html=True)

        col_a, col_b = st.columns(2)
        with col_a:
            st.plotly_chart(bar_scores(rows), use_container_width=True)
        with col_b:
            st.plotly_chart(bar_penalites(rows), use_container_width=True)

        col_c, col_d = st.columns(2)
        with col_c:
            pie = pie_attente(rows)
            if pie:
                st.plotly_chart(pie, use_container_width=True)
            else:
                st.info("✅ Aucun temps d'attente — tous les camions sont chargés dès leur arrivée.")
        with col_d:
            st.plotly_chart(scatter_arrivee_debut(rows), use_container_width=True)

        # --- Décomposition Z* ---
        st.markdown("---")
        st.markdown('<div class="section-title">🧮 Décomposition de Z*</div>', unsafe_allow_html=True)
        decomp_data = {
            "Composante": ["+ Score éco total", "− Pénalités retard", "− Temps attente total", "= Z*"],
            "Valeur":     [
                f"{result['score_eco']:.0f}",
                f"−{result['penalites']:.0f} TND",
                f"−{result['att_tot']:.0f} min",
                f"{result['Z']:.2f}",
            ],
        }
        st.dataframe(
            pd.DataFrame(decomp_data),
            use_container_width=False, hide_index=True,
        )

        # Info modèle
        st.caption(
            f"Modèle PLNM · {result['n_vars']} variables · "
            f"{result['n_cons']} contraintes · Solveur CBC"
        )
    else:
        st.info("👈 Configurez les paramètres puis cliquez sur **▶ Lancer l'optimisation**.")


if __name__ == "__main__":
    main()