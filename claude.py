"""
=============================================================================
BALLET DES TRANSPORTEURS SUR LES QUAIS DE CHARGEMENT
Modélisation en Programme Linéaire en Nombres Mixtes (PLNM) avec PuLP
Cas Pédagogique — Green Wood Design (GWD) — ENIT 2026
=============================================================================

Données source : Données_GWD_2026_VF.xlsx — Feuille : Données Transporteurs

Problème :
  • Sélectionner K=9 camions parmi N=12
  • Affecter chaque camion sélectionné à l'un des Q=2 quais
  • Déterminer l'ordre de passage optimal sur chaque quai
  
Objectif (max Z) :
  Z = Σ Sᵢ·uᵢ  −  Σ Pᵢ·rᵢ  −  Σ (dᵢ − Taᵢ)
      ↑ score éco   ↑ pénalités   ↑ attente
=============================================================================

Usage :
    pip install pulp openpyxl
    python ballet_transporteurs_pulp.py
    
    # Pour charger les données depuis Excel :
    python ballet_transporteurs_pulp.py --xlsx Données_GWD_2026_VF.xlsx
=============================================================================
"""

import sys
import argparse
import pulp

# ---------------------------------------------------------------------------
# 1. DONNÉES (embarquées — identiques à la feuille Excel)
# ---------------------------------------------------------------------------

# Transporteurs : {nom: (score_eco, penalite)}
TRANSPORTEURS = {
    "GreenWay":     (95, 8),
    "RapidCargo":   (40, 5),
    "Proxilog":     (75, 6),
    "StarFret":     (60, 7),
    "UltraExpress": (30, 4),
}

# Camions : {id: (transporteur, heure_arrivee_min, duree_min, heure_limite_min)}
# Toutes les heures sont converties en minutes depuis 18h00 (base de référence)
# pour éviter les problèmes de minuit (ex: 20h00 → 120, 01h00 → 420, etc.)
def heure_to_min(h_str: str, base_h: int = 18) -> int:
    """Convertit 'HHhMM' en minutes depuis base_h heures."""
    h_str = h_str.strip().lower().replace("h", ":")
    if ":" in h_str:
        parts = h_str.split(":")
        h, m = int(parts[0]), int(parts[1]) if len(parts) > 1 and parts[1] else 0
    else:
        h, m = int(h_str), 0
    minutes = h * 60 + m
    # Ajustement si l'heure est "le lendemain" (avant base_h)
    if h < base_h:
        minutes += 24 * 60
    return minutes - base_h * 60


# Données brutes issues du fichier Excel
_CAMIONS_RAW = {
    #  id    transporteur     arrivée   durée  limite
    "T1":  ("GreenWay",     "20h00",   180,   "02h00"),
    "T2":  ("RapidCargo",   "21h00",   240,   "03h00"),
    "T3":  ("Proxilog",     "22h00",   120,   "01h00"),
    "T4":  ("StarFret",     "23h00",   180,   "02h00"),
    "T5":  ("UltraExpress", "00h00",   300,   "06h00"),
    "T6":  ("GreenWay",     "01h00",   180,   "05h00"),
    "T7":  ("RapidCargo",   "02h00",   240,   "07h00"),
    "T8":  ("Proxilog",     "03h00",   120,   "05h00"),
    "T9":  ("StarFret",     "04h00",   180,   "08h00"),
    "T10": ("UltraExpress", "22h00",   240,   "03h00"),
    "T11": ("GreenWay",     "23h00",   180,   "04h00"),
    "T12": ("Proxilog",     "00h00",   120,   "02h00"),
}

# Construction du dict camions avec heures converties en minutes
CAMIONS = {}
for cid, (trans, arr, dur, lim) in _CAMIONS_RAW.items():
    score, penalite = TRANSPORTEURS[trans]
    CAMIONS[cid] = {
        "transporteur": trans,
        "S": score,           # Score éco (hérité du transporteur)
        "P": penalite,        # Pénalité TND/min (héritée du transporteur)
        "Ta": heure_to_min(arr),   # Heure d'arrivée (min depuis 18h00)
        "Td": dur,            # Durée de chargement (min)
        "Tl": heure_to_min(lim),   # Heure limite de départ (min depuis 18h00)
    }

# Paramètres globaux
N = 12      # Nombre total de camions
K = 9       # Nombre de camions à sélectionner
Q = 2       # Nombre de quais
M = 1440    # Grande constante (Big-M) = 24h en minutes


# ---------------------------------------------------------------------------
# 2. CHARGEMENT OPTIONNEL DEPUIS EXCEL
# ---------------------------------------------------------------------------

def charger_depuis_excel(chemin: str) -> None:
    """Recharge CAMIONS et TRANSPORTEURS depuis le fichier Excel."""
    try:
        import openpyxl
    except ImportError:
        print("[AVERTISSEMENT] openpyxl non disponible — utilisation des données embarquées.")
        return

    wb = openpyxl.load_workbook(chemin, data_only=True)
    ws = wb["Données Transporteurs"]
    rows = [row for row in ws.iter_rows(values_only=True) if any(v is not None for v in row)]

    # --- Lecture des transporteurs (lignes 1-5 = indices 0-4) ---
    TRANSPORTEURS.clear()
    for row in rows[1:6]:
        nom, score, pen, _ = row[0], row[1], row[2], row[3]
        if nom and isinstance(score, (int, float)):
            TRANSPORTEURS[nom] = (int(score), int(pen))

    # --- Lecture des camions (après l'entête "Camion") ---
    start = next(i for i, r in enumerate(rows) if r[0] == "Camion") + 1
    CAMIONS.clear()
    for row in rows[start:]:
        if row[0] is None or not str(row[0]).startswith("T"):
            break
        cid, arr, dur, lim = str(row[0]), str(row[1]), int(row[2]), str(row[3])
        # Retrouver le transporteur de ce camion
        trans = next(
            t for t, (_, __) in TRANSPORTEURS.items()
            if cid in [c.strip() for c in
                       [r[3] for r in rows[1:6] if r[0] == t][0].split(",")]
        )
        score, penalite = TRANSPORTEURS[trans]
        CAMIONS[cid] = {
            "transporteur": trans,
            "S": score,
            "P": penalite,
            "Ta": heure_to_min(arr),
            "Td": dur,
            "Tl": heure_to_min(lim),
        }

    print(f"[INFO] Données chargées depuis {chemin} : {len(CAMIONS)} camions, {len(TRANSPORTEURS)} transporteurs.")


# ---------------------------------------------------------------------------
# 3. CONSTRUCTION DU MODÈLE PULP
# ---------------------------------------------------------------------------

def construire_modele():
    """Construit et retourne le modèle PuLP."""
    try:
        import pulp
    except ImportError:
        raise ImportError("PuLP non installé. Exécutez : pip install pulp")

    camions = list(CAMIONS.keys())   # ['T1', ..., 'T12']
    quais   = list(range(1, Q + 1))  # [1, 2]

    prob = pulp.LpProblem("Ballet_Transporteurs_GWD", pulp.LpMaximize)

    # ------------------------------------------------------------------
    # 3.1  Variables de décision
    # ------------------------------------------------------------------

    # u[i] ∈ {0,1} : sélection du camion i
    u = pulp.LpVariable.dicts("u", camions, cat="Binary")

    # x[i][j] ∈ {0,1} : affectation camion i → quai j
    x = pulp.LpVariable.dicts("x",
                               [(i, j) for i in camions for j in quais],
                               cat="Binary")

    # d[i] ≥ 0 : heure de début de chargement (en min depuis 18h00)
    d = pulp.LpVariable.dicts("d", camions, lowBound=0, cat="Continuous")

    # r[i] ≥ 0 : retard du camion i (en minutes)
    r = pulp.LpVariable.dicts("r", camions, lowBound=0, cat="Continuous")

    # y[i][k][j] ∈ {0,1} : i passe avant k sur quai j
    paires = [(i, k, j)
              for idx_i, i in enumerate(camions)
              for k in camions[idx_i + 1:]   # i < k pour éviter les doublons
              for j in quais]
    y = pulp.LpVariable.dicts("y", paires, cat="Binary")

    # ------------------------------------------------------------------
    # 3.2  Fonction objectif
    # ------------------------------------------------------------------
    # max Z = Σ Sᵢ·uᵢ  −  Σ Pᵢ·rᵢ  −  Σ (dᵢ − Taᵢ·uᵢ)
    # Note : on utilise Taᵢ·uᵢ pour que l'attente soit 0 quand uᵢ=0

    prob += (
        pulp.lpSum(CAMIONS[i]["S"] * u[i] for i in camions)
        - pulp.lpSum(CAMIONS[i]["P"] * r[i] for i in camions)
        - pulp.lpSum(d[i] - CAMIONS[i]["Ta"] * u[i] for i in camions)
    ), "Objectif_Z"

    # ------------------------------------------------------------------
    # 3.3  Contraintes
    # ------------------------------------------------------------------

    # C1 — Sélection exacte de K=9 camions
    prob += pulp.lpSum(u[i] for i in camions) == K, "C1_selection_exacte"

    # C2 — Affectation liée à la sélection : Σⱼ xᵢⱼ = uᵢ
    for i in camions:
        prob += (
            pulp.lpSum(x[i, j] for j in quais) == u[i],
            f"C2_affectation_{i}"
        )

    # C3 — Début de chargement après arrivée
    for i in camions:
        Ta_i = CAMIONS[i]["Ta"]
        # dᵢ ≥ Taᵢ · uᵢ
        prob += d[i] >= Ta_i * u[i], f"C3a_apres_arrivee_{i}"
        # dᵢ ≤ M · uᵢ  (si non sélectionné → dᵢ = 0)
        prob += d[i] <= M * u[i], f"C3b_zero_si_non_selectionne_{i}"

    # C4 — Calcul du retard
    for i in camions:
        Td_i = CAMIONS[i]["Td"]
        Tl_i = CAMIONS[i]["Tl"]
        # rᵢ ≥ dᵢ + Tdᵢ − Tlᵢ
        prob += r[i] >= d[i] + Td_i - Tl_i, f"C4a_retard_{i}"
        # rᵢ ≤ M · uᵢ  (pas de pénalité fantôme si non sélectionné)
        prob += r[i] <= M * u[i], f"C4b_retard_zero_si_exclu_{i}"

    # C5 & C6 — Non-chevauchement sur chaque quai (Big-M)
    for (i, k, j) in paires:
        Td_i = CAMIONS[i]["Td"]
        Td_k = CAMIONS[k]["Td"]

        # C5 : si yᵢₖⱼ=1 (i avant k) et tous deux sur quai j → i finit avant k commence
        prob += (
            d[i] + Td_i
            <= d[k]
            + M * (1 - y[i, k, j])
            + M * (1 - x[i, j])
            + M * (1 - x[k, j]),
            f"C5_i_avant_k_{i}_{k}_q{j}"
        )

        # C6 : si yᵢₖⱼ=0 (k avant i) et tous deux sur quai j → k finit avant i commence
        prob += (
            d[k] + Td_k
            <= d[i]
            + M * y[i, k, j]
            + M * (1 - x[i, j])
            + M * (1 - x[k, j]),
            f"C6_k_avant_i_{i}_{k}_q{j}"
        )

    return prob, u, x, d, r, y, camions, quais


# ---------------------------------------------------------------------------
# 4. RÉSOLUTION ET AFFICHAGE DES RÉSULTATS
# ---------------------------------------------------------------------------

def afficher_resultats(prob, u, x, d, r, y, camions, quais):
    """Affiche les résultats après résolution."""
    import pulp

    statut = pulp.LpStatus[prob.status]
    print("\n" + "=" * 65)
    print("  RÉSULTATS — Ballet des Transporteurs GWD")
    print("=" * 65)
    print(f"  Statut du solveur : {statut}")

    if prob.status != 1:
        print("  [ERREUR] Pas de solution optimale trouvée.")
        return

    Z = pulp.value(prob.objective)
    print(f"  Valeur objective Z* = {Z:.2f}")
    print()

    # --- Camions sélectionnés ---
    selectionnes = [i for i in camions if pulp.value(u[i]) > 0.5]
    exclus = [i for i in camions if pulp.value(u[i]) < 0.5]

    print(f"  ✔  Camions sélectionnés ({len(selectionnes)}) : {', '.join(selectionnes)}")
    print(f"  ✘  Camions exclus       ({len(exclus)})  : {', '.join(exclus)}")

    # --- Détails par quai ---
    print()
    for j in quais:
        camions_quai = [i for i in selectionnes if pulp.value(x[i, j]) > 0.5]
        # Tri par heure de début de chargement
        camions_quai.sort(key=lambda i: pulp.value(d[i]))

        print(f"  ┌─ QUAI {j} ({len(camions_quai)} camions) " + "─" * 40)
        print(f"  │  {'Camion':<6}  {'Transporteur':<14}  {'Éco':>4}  "
              f"{'Début':>8}  {'Fin':>8}  {'Limite':>8}  {'Retard':>7}  {'Pénalité':>9}")
        print(f"  │  {'─'*6}  {'─'*14}  {'─'*4}  {'─'*8}  {'─'*8}  {'─'*8}  {'─'*7}  {'─'*9}")

        for i in camions_quai:
            di    = pulp.value(d[i])
            ri    = pulp.value(r[i])
            fin   = di + CAMIONS[i]["Td"]
            Tl_i  = CAMIONS[i]["Tl"]
            Pi    = CAMIONS[i]["P"]

            def min_to_hhmm(m):
                """Convertit des minutes (depuis 18h00) en HH:MM."""
                total = int(round(m)) + 18 * 60
                return f"{(total // 60) % 24:02d}h{total % 60:02d}"

            print(f"  │  {i:<6}  {CAMIONS[i]['transporteur']:<14}  "
                  f"{CAMIONS[i]['S']:>4}  "
                  f"{min_to_hhmm(CAMIONS[i]['Ta']):>8}→{min_to_hhmm(di):>5}  "
                  f"{min_to_hhmm(fin):>8}  "
                  f"{min_to_hhmm(Tl_i):>8}  "
                  f"{ri:>6.0f}m  "
                  f"{Pi * ri:>8.0f} ₮")
        print(f"  └" + "─" * 55)
        print()

    # --- Décomposition de Z ---
    score_eco   = sum(CAMIONS[i]["S"] * pulp.value(u[i]) for i in camions)
    penalites   = sum(CAMIONS[i]["P"] * pulp.value(r[i]) for i in camions)
    attente_tot = sum(pulp.value(d[i]) - CAMIONS[i]["Ta"] * pulp.value(u[i]) for i in camions)

    print(f"  Décomposition de Z* :")
    print(f"    + Score éco total       : {score_eco:.0f}")
    print(f"    − Pénalités retard      : {penalites:.0f} TND")
    print(f"    − Temps attente total   : {attente_tot:.0f} min")
    print(f"    ═══════════════════════════")
    print(f"      Z*                    = {Z:.2f}")
    print()


# ---------------------------------------------------------------------------
# 5. POINT D'ENTRÉE
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description="Ballet des Transporteurs GWD — Modèle PLNM PuLP"
    )
    parser.add_argument(
        "--xlsx", type=str, default=None,
        help="Chemin vers le fichier Excel Données_GWD_2026_VF.xlsx"
    )
    parser.add_argument(
        "--solver", type=str, default="CBC",
        help="Solveur PuLP à utiliser : CBC (défaut), GLPK, CPLEX, GUROBI…"
    )
    parser.add_argument(
        "--timelimit", type=int, default=300,
        help="Limite de temps en secondes (défaut : 300)"
    )
    args = parser.parse_args()

    # --- Chargement des données ---
    if args.xlsx:
        charger_depuis_excel(args.xlsx)
    else:
        print("[INFO] Utilisation des données embarquées (identiques à l'Excel).")
        print("[INFO] Pour charger depuis Excel : python script.py --xlsx Données_GWD_2026_VF.xlsx")

    print("\n[INFO] Affichage des données du problème :")
    print(f"       N={N} camions, K={K} à sélectionner, Q={Q} quais, M={M}")
    print()

    # --- Construction ---
    print("[INFO] Construction du modèle PuLP...")
    try:
        prob, u, x, d, r, y, camions, quais = construire_modele()
    except ImportError:
        print("\n[ERREUR] PuLP n'est pas installé.")
        print("  Installez-le avec : pip install pulp")
        print()
        n_bin = N + N * Q + N * (N - 1) // 2 * Q
        print("  Structure du modèle (sans PuLP) :")
        print(f"    Variables binaires  : u({N}) + x({N}×{Q}) + y({N*(N-1)//2}×{Q}) = {n_bin}")
        print(f"    Variables continues : d({N}) + r({N}) = {2*N}")
        print(f"    Familles de contraintes : C1 à C8")
        sys.exit(1)

    # --- Résolution ---
    print(f"[INFO] Résolution avec le solveur {args.solver} (limite : {args.timelimit}s)...")
    solver_map = {
        "CBC":    pulp.PULP_CBC_CMD(msg=1, timeLimit=args.timelimit),
        "GLPK":   pulp.GLPK_CMD(msg=1),
        "CPLEX":  pulp.CPLEX_CMD(msg=1),
        "GUROBI": pulp.GUROBI_CMD(msg=1),
    }
    solver = solver_map.get(args.solver.upper(), pulp.PULP_CBC_CMD(msg=1, timeLimit=args.timelimit))

    prob.solve(solver)

    # --- Résultats ---
    afficher_resultats(prob, u, x, d, r, y, camions, quais)


if __name__ == "__main__":
    main()