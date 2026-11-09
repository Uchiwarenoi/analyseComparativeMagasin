import argparse
import random
import time
import statistics
from collections import Counter, defaultdict
from datetime import date, timedelta
from pathlib import Path

from app.common import (
    ML, BB, CI, SOURCES, DATE_FICHIERS, AN_DEBUT, AN_FIN, BF_DEBUT, BF_FIN,
    PANIER, Journal, Chrono, fr, pct, md_table,
)
from app.chargement import (
    charger_catalogue, charger_marche_lagune, charger_baobab, charger_comptoir,
)
from app.rapprochement import rapprocher
from app.historique import (
    Historique, controler_avec_historique, analyse_historique,
    indice_hebdomadaire, panier_historique,
)
from app.comparaison import comparatif, synthese, calcul_panier
from app.exportation import ecrire_csv
from app.interface import menu

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--data", default=str(Path(__file__).resolve().parent / "donnees"))
    ap.add_argument("--out", default=str(Path(__file__).resolve().parent / "resultats"))
    ap.add_argument("--sans-menu", action="store_true", help="fait l'analyse et écrit les fichiers, sans ouvrir le menu")
    a = ap.parse_args()
    D, O = Path(a.data), Path(a.out)
    O.mkdir(parents=True, exist_ok=True)
    j, chrono = Journal(), Chrono()
    t_total = time.perf_counter()

    with chrono("Chargement des 3 enseignes (CSV + JSON + Excel)"):
        catalogue = charger_catalogue(D / "catalogue_reference.csv")
        offres = (charger_marche_lagune(D / "marche_lagune.csv", j) + charger_baobab(D / "baobab_shop.json", j)
                  + charger_comptoir(D / "comptoir_ivoire.xlsx", j))
    n_brut = len(offres)
    with chrono("Chargement historique + construction de l'index"):
        hist = Historique.depuis_csv(D / "historique_prix.csv", catalogue, j)
    with chrono("Rapprochement avec le catalogue"):
        offres = rapprocher(offres, catalogue, j)
    with chrono("Contrôle des prix avec l'historique"):
        controler_avec_historique(offres, hist, j)

    par_pid = defaultdict(dict)                               # produit_id -> {enseigne: Offre}  : accès O(1)
    for o in offres:
        par_pid[o.produit_id][o.source] = o

    ecrire_csv(O / "rapprochement.csv", ["enseigne", "reference", "libelle_source", "produit_id", "nom_catalogue",
                                         "score", "prix_affiche", "frais", "remise_pct", "prix_paye", "disponible", "stock"],
               [(o.source, o.ref, o.libelle, o.produit_id, catalogue[o.produit_id]["nom"], round(o.score, 2),
                 o.prix_affiche if o.prix_affiche is not None else "", o.frais, o.remise,
                 o.prix_paye if o.prix_paye is not None else "", "oui" if o.en_stock else "non", o.stock_info) for o in offres])

    with chrono("Comparatif + synthèse + panier"):
        comp = comparatif(par_pid, catalogue)
        victoires, seuls, ecarts, top_fcfa, top_pct = synthese(comp)
        lignes_panier, totaux, manques, mixte = calcul_panier(par_pid, catalogue)
    ecrire_csv(O / "comparatif.csv",
               ["produit_id", "nom", "categorie"] + [f"prix_{s}" for s in SOURCES] +
               ["nb_enseignes", "enseigne_moins_chere", "prix_min", "ecart_fcfa", "ecart_pct"],
               [[c["produit_id"], c["nom"], c["categorie"]] + [c["prix"].get(s, "") for s in SOURCES] +
                [c["nb"], " / ".join(c["gagnants"]), c["mn"], c["ecart"], round(c["ecart_pct"], 1)] for c in comp])
    ecrire_csv(O / "panier.csv", ["produit_id", "nom", "quantite"] + [f"prix_unitaire_{s}" for s in SOURCES] + ["moins_cher"],
               [[p, n, q] + [pr.get(s, "indisponible") for s in SOURCES] + [" / ".join(g)] for p, n, q, pr, g in lignes_panier])

    with chrono("Analyse de l'historique (Black Friday, semaines, panier)"):
        hres = analyse_historique(hist, catalogue)
        moy_an = {(r["source"], r["pid"]): r["an"] for r in hres}
        hebdo = {s: indice_hebdomadaire(hist, s, moy_an) for s in SOURCES}
        pan_h = {s: panier_historique(hist, s) for s in SOURCES}
    ecrire_csv(O / "black_friday.csv",
               ["enseigne", "produit_id", "nom", "prix_moyen_BF_24-30nov2025", "prix_moyen_annee", "ecart_BF_vs_annee_pct",
                "prix_le_plus_bas", "date_prix_le_plus_bas", "BF_vs_plus_bas_pct"],
               [[r["source"], r["pid"], r["nom"], round(r["bf"]), round(r["an"]), round(r["ecart"], 1), r["mini"],
                 r["date_mini"].isoformat(), round(r["bf_vs_mini"], 1)] for r in hres])

    # --- Benchmark complexité : recherche naïve (balayage) vs index dict + bisect -----------------
    random.seed(42)
    cles = list(hist.dates)
    Q = 2000
    requetes = [(*random.choice(cles), AN_DEBUT + timedelta(days=random.randrange(14, 365))) for _ in range(Q)]
    t0 = time.perf_counter()
    r_naif = [hist.prix_a_naif(s, p, d) for s, p, d in requetes]
    t_naif = time.perf_counter() - t0
    t0 = time.perf_counter()
    r_idx = [hist.prix_a(s, p, d) for s, p, d in requetes]
    t_idx = time.perf_counter() - t0
    assert r_naif == r_idx, "les deux méthodes doivent donner les mêmes prix"
    chrono.durees["TOTAL (hors benchmark)"] = time.perf_counter() - t_total - t_naif - t_idx

    # --- Rapport -----------------------------------------------------------------------------------
    nom = lambda pid: f"{pid} {catalogue[pid]['nom']}"
    R = []
    R.append("# Analyse comparative — Marché Lagune / Baobab Shop / Comptoir Ivoire\n")
    R.append("_Données fictives. Prix en FCFA, **prix réellement payés** : Marché Lagune = prix affiché (livraison incluse) ; "
             "Baobab Shop = prix + frais de livraison ; Comptoir Ivoire = prix catalogue après remise, retrait en boutique "
             "(transport du client non compté). Un article en rupture ou sans prix n'est jamais retenu._\n")
    R.append("## 1. Chargement, nettoyage, rapprochement\n")
    niv = Counter(l["niveau"] for l in j.lignes)
    R.append(f"{n_brut} lignes lues dans les 3 enseignes -> {len(offres)} articles rapprochés d'un produit du catalogue. "
             f"Journal complet : `journal_nettoyage.csv` ({niv['ERREUR']} erreurs, {niv['ATTENTION']} avertissements, {niv['INFO']} infos). "
             "Points principaux :\n")
    R.append(md_table(["Fichier", "Élément", "Problème", "Traitement"],
                      [(l["fichier"], l["element"], l["probleme"], l["traitement"]) for l in j.lignes if l["niveau"] == "ERREUR"]))
    R.append("\nAutres problèmes traités (avertissements) : " + "; ".join(
        f"{l['element']} — {l['probleme']}" for l in j.lignes if l["niveau"] == "ATTENTION" and l["fichier"] != "comptoir_ivoire.xlsx") + ".\n")
    R.append("**Rapprochement** : les libellés sont normalisés (accents, casse, unités `128GB`=`128 Go`, `1HP`=`1 CV`, "
             "`24\"`=`24 pouces`, synonymes `Frigo`/`Fridge`/`Réfrigérateur`, etc.), découpés en mots, puis comparés à chaque "
             "produit du catalogue par une similarité pondérée (un mot rare, comme `A25`, pèse plus qu'un mot courant). "
             "Un article est accepté si son score ≥ 0,30 et s'il bat le 2e candidat d'au moins 0,10 ; les fautes (`Samsun`) sont "
             "corrigées par proximité orthographique. Les SKU Baobab servent de contrôle croisé (" + str(sum(1 for l in j.lignes if "SKU suggère" in l["probleme"])) + " désaccord). Détail : `rapprochement.csv`.\n")
    R.append("## 2. Comparatif\n")
    R.append(f"{len(comp)} produits sur 30 disponibles dans au moins deux enseignes (`comparatif.csv`). "
             "Écart = prix le plus élevé − prix le plus bas (en FCFA et en % du prix le plus bas).\n")
    R.append(md_table(["Produit"] + SOURCES + ["Moins cher", "Écart FCFA", "Écart %"],
                      [[nom(c["produit_id"])] + [fr(c["prix"][s]) if s in c["prix"] else "—" for s in SOURCES] +
                       [" / ".join(c["gagnants"]), fr(c["ecart"]), pct(c["ecart_pct"])] for c in comp]))
    R.append("\n## 3. Synthèse\n")
    R.append("**Combien de fois chaque enseigne est-elle la moins chère ?** (un ex æquo compte pour chaque enseigne à égalité)\n")
    R.append(md_table(["Enseigne", "Moins chère (total)", "dont seule", "dont ex æquo", "Sur n produits comparables"],
                      [[s, len(victoires[s]), seuls[s], len(victoires[s]) - seuls[s], ecarts[s][0]] for s in SOURCES]))
    R.append("\n**Écart moyen par rapport au prix le plus bas** (0 quand l'enseigne est la moins chère ; moyenne sur les produits où elle a un prix retenu) :\n")
    R.append(md_table(["Enseigne", "Écart moyen %", "Écart moyen FCFA", "n"],
                      [[s, pct(ecarts[s][2]), fr(ecarts[s][1]), ecarts[s][0]] for s in SOURCES]))
    R.append("\n**Les 3 produits où il y a le plus à gagner à comparer** (classés par écart en FCFA) :\n")
    R.append(md_table(["Produit", "Moins cher", "Plus cher", "Écart FCFA", "Écart %"],
                      [[nom(c["produit_id"]), f"{' / '.join(c['gagnants'])} ({fr(c['mn'])})",
                        f"{max(c['prix'], key=c['prix'].get)} ({fr(c['mx'])})", fr(c["ecart"]), pct(c["ecart_pct"])] for c in top_fcfa]))
    R.append("\nClassés par écart en % : " + ", ".join(f"{nom(c['produit_id'])} ({pct(c['ecart_pct'])})" for c in top_pct) + ".\n")
    R.append("## 4. Panier de rentrée\n")
    R.append("1 Tecno Spark 20, 1 Lenovo IdeaPad 3, 2 clés USB SanDisk, 1 disque dur Seagate, 1 micro-ondes Samsung "
             "(frais de livraison Baobab comptés par article, donc × 2 pour les clés USB).\n")
    R.append(md_table(["Article", "Qté"] + SOURCES + ["Moins cher"],
                      [[nom(p), q] + [fr(pr[s]) if s in pr else "indisponible" for s in SOURCES] + [" / ".join(g)]
                       for p, _, q, pr, g in [(l[0], l[1], l[2], l[3], l[4]) for l in lignes_panier]] +
                      [["**Total panier**", ""] + [f"**{fr(totaux[s])}**" if not manques[s] else "incomplet" for s in SOURCES] + [""]]))
    complets = {s: t for s, t in totaux.items() if not manques[s]}
    meilleur = min(complets, key=complets.get)
    R.append(f"\n- Panier complet dans une seule enseigne : le moins cher est **{meilleur}** à **{fr(complets[meilleur])} FCFA**"
             + "".join(f" ; {s} : {fr(t)} (+{fr(t - complets[meilleur])})" for s, t in complets.items() if s != meilleur) + ".")
    R.append(f"- En achetant chaque article là où il est le moins cher : **{fr(mixte)} FCFA**, soit "
             f"{fr(complets[meilleur] - mixte)} FCFA ({pct((complets[meilleur] - mixte) / complets[meilleur] * 100)}) de moins que le meilleur panier unique. "
             "Contrepartie : plusieurs enseignes, donc plusieurs déplacements/livraisons.\n")
    R.append("## 5. Historique : quand fallait-il acheter ?\n")
    R.append("Prix **affichés** de l'historique (Baobab sans livraison, Comptoir sans remise) ; la comparaison se fait enseigne par enseigne "
             "contre sa propre moyenne annuelle (01/11/2025–31/10/2026), donc ces biais ne la faussent pas. Comptoir Ivoire change de prix le lundi : "
             "le prix en vigueur est reporté jusqu'au changement suivant, ce qui donne bien une moyenne pondérée par jour. Détail par produit : `black_friday.csv`.\n")
    lignes_bf = []
    for s in SOURCES:
        rs = [r for r in hres if r["source"] == s]
        moins = sum(1 for r in rs if r["ecart"] < 0)
        lignes_bf.append([s, len(rs), pct(statistics.mean(r["ecart"] for r in rs)), f"{moins} / {len(rs)}",
                          pct(statistics.mean(r["bf_vs_mini"] for r in rs))])
    R.append(md_table(["Enseigne", "Produits", "Semaine BF vs moyenne annuelle (moyenne des écarts)", "Produits moins chers en BF",
                       "BF vs plus bas prix de l'année"], lignes_bf))
    R.append("\n**Semaines les moins chères** (indice 100 % = moyenne annuelle ; semaines lundi-dimanche) :\n")
    lignes_sem = []
    for s in SOURCES:
        h = hebdo[s]
        tri = sorted(h, key=h.get)
        bf_lundi = BF_DEBUT
        lignes_sem.append([s] + [f"{w:%d/%m/%Y} ({pct(h[w] * 100)})" for w in tri[:3]] +
                          [f"{pct(h[bf_lundi] * 100)} (rang {tri.index(bf_lundi) + 1}/{len(tri)})"])
    R.append(md_table(["Enseigne", "1re", "2e", "3e", "Semaine du 24/11/2025"], lignes_sem))
    R.append("\n**Mois du prix le plus bas de l'année** (nombre de produits dont le minimum tombe ce mois-là) :\n")
    lignes_m = []
    for s in SOURCES:
        c = Counter(f"{r['date_mini']:%Y-%m}" for r in hres if r["source"] == s)
        lignes_m.append([s, ", ".join(f"{m} ({n})" for m, n in c.most_common(3))])
    R.append(md_table(["Enseigne", "Top 3 des mois"], lignes_m))
    R.append("\n**Le panier de rentrée dans l'historique** (prix affichés) :\n")
    lignes_p = []
    for s in SOURCES:
        pj = pan_h[s]
        bf = statistics.mean(v for d, v in pj.items() if BF_DEBUT <= d <= BF_FIN)
        moy = statistics.mean(pj.values())
        dmin = min(pj, key=pj.get)
        lignes_p.append([s, fr(moy), fr(bf), pct((bf / moy - 1) * 100), f"{fr(pj[dmin])} le {dmin:%d/%m/%Y}", fr(pj[AN_FIN])])
    R.append(md_table(["Enseigne", "Moyenne annuelle", "Semaine BF", "BF vs année", "Plus bas de l'année", "Au 31/10/2026"], lignes_p))
    R.append("\n## 6. Complexité et structures de données\n")
    R.append(f"Historique : {len(hist.lignes)} lignes, {len(hist.dates)} séries (enseigne, produit) de ≤ 365 points.\n")
    R.append(md_table(["Étape", "Temps (s)"], [[k, f"{v:.4f}"] for k, v in chrono.durees.items()]))
    R.append(f"\n**Benchmark** : {Q} recherches « prix de tel produit chez telle enseigne à telle date » (résultats identiques, vérifié) :\n")
    R.append(md_table(["Méthode", "Complexité par recherche", f"Temps pour {Q} recherches"],
                      [["Balayage de la liste des 20 876 lignes", "O(n)", f"{t_naif:.3f} s"],
                       ["Index dict + liste triée + bisect", "O(log m), m ≤ 365", f"{t_idx:.5f} s"]]))
    R.append(f"\nL'index est environ **{t_naif / max(t_idx, 1e-9):,.0f} fois plus rapide**. Choix retenus :".replace(",", " "))
    R.append("""
- **`dict` clé `(enseigne, produit_id)` -> séries** : accès direct en O(1) à la bonne série au lieu de filtrer 20 876 lignes ; construit en une passe O(n).
- **Séries triées par date + `bisect`** : trouver le prix à une date coûte O(log m) (m ≤ 365) ; trier une fois coûte O(n log m) au chargement et sert ensuite à toutes les requêtes. Ça gère aussi Comptoir Ivoire (1 prix/semaine) : `bisect_right − 1` renvoie le dernier prix publié, sans stocker de lignes par jour.
- **`dict` produit_id -> {enseigne: offre}** pour le comparatif : accès O(1), comparatif en O(P × S) = 30 × 3.
- **Rapprochement** : 76 articles × 30 produits × quelques mots = quelques milliers de comparaisons, donc un simple balayage suffit ; avec un catalogue de 100 000 produits il faudrait un index inversé (mot -> produits) pour ne comparer que les candidats.
- **Ensembles (`set`) de mots** pour l'intersection/union des libellés : O(taille du plus petit).
- Bibliothèque standard uniquement (`openpyxl` facultatif) pour garder les coûts visibles ; l'historique tient en mémoire (≈ 20 k tuples), la mémoire est en O(n).
""")
    (O / "rapport.md").write_text("\n".join(R), encoding="utf-8")
    j.ecrire(O / "journal_nettoyage.csv")
    if not a.sans_menu:
        menu("\n".join(R), O)


if __name__ == "__main__":
    main()
