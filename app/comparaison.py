import statistics

from app.common import SOURCES, PANIER

def comparatif(par_pid, catalogue):
    lignes = []
    for pid, cat in catalogue.items():
        prix = {s: o.prix_paye for s, o in par_pid.get(pid, {}).items() if o.exploitable}
        if len(prix) < 2:
            continue
        mn, mx = min(prix.values()), max(prix.values())
        lignes.append(dict(produit_id=pid, nom=cat["nom"], categorie=cat["categorie"], prix=prix, nb=len(prix),
                           gagnants=[s for s, p in prix.items() if p == mn], mn=mn, mx=mx,
                           ecart=mx - mn, ecart_pct=(mx - mn) / mn * 100))
    return lignes


def synthese(comp):
    victoires = {s: [c["produit_id"] for c in comp if s in c["gagnants"]] for s in SOURCES}
    seuls = {s: sum(1 for c in comp if c["gagnants"] == [s]) for s in SOURCES}
    ecarts = {}
    for s in SOURCES:
        v = [(c["prix"][s] - c["mn"], (c["prix"][s] - c["mn"]) / c["mn"] * 100) for c in comp if s in c["prix"]]
        ecarts[s] = (len(v), statistics.mean(x for x, _ in v) if v else None, statistics.mean(y for _, y in v) if v else None)
    top_fcfa = sorted(comp, key=lambda c: -c["ecart"])[:3]
    top_pct = sorted(comp, key=lambda c: -c["ecart_pct"])[:3]
    return victoires, seuls, ecarts, top_fcfa, top_pct


def calcul_panier(par_pid, catalogue):
    lignes, totaux, manques = [], {s: 0 for s in SOURCES}, {s: [] for s in SOURCES}
    mixte = 0
    for pid, q in PANIER.items():
        prix = {s: o.prix_paye for s, o in par_pid.get(pid, {}).items() if o.exploitable}
        for s in SOURCES:
            if s in prix:
                totaux[s] += q * prix[s]
            else:
                manques[s].append(pid)
        mn = min(prix.values())
        mixte += q * mn
        lignes.append((pid, catalogue[pid]["nom"], q, prix, [s for s, p in prix.items() if p == mn]))
    return lignes, totaux, manques, mixte

