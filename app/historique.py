import csv
import bisect
import statistics
from collections import defaultdict
from datetime import date, timedelta

from app.common import SOURCES, AN_DEBUT, AN_FIN, BF_DEBUT, BF_FIN, DATE_FICHIERS, PANIER, fr

class Historique:
    def __init__(self, lignes):
        self.lignes = lignes                                  # liste brute (pour la comparaison naïve)
        tmp = defaultdict(list)
        for d, s, p, prix in lignes:
            tmp[(s, p)].append((d, prix))
        self.dates, self.prix = {}, {}
        for k, serie in tmp.items():
            serie.sort()                                      # O(m log m) par série
            self.dates[k] = [d for d, _ in serie]
            self.prix[k] = [p for _, p in serie]

    @classmethod
    def depuis_csv(cls, chemin, catalogue, j):
        lignes, vus = [], set()
        with open(chemin, encoding="utf-8-sig", newline="") as f:
            for n, r in enumerate(csv.DictReader(f), start=2):
                try:
                    d, prix = date.fromisoformat(r["date"]), int(r["prix_fcfa"])
                except ValueError:
                    j.add("historique_prix.csv", f"ligne {n}", "ERREUR", f"valeur illisible {r}", "ligne ignorée")
                    continue
                if prix <= 0 or r["source"] not in SOURCES or r["produit_id"] not in catalogue:
                    j.add("historique_prix.csv", f"ligne {n}", "ERREUR", f"ligne incohérente {r}", "ligne ignorée")
                    continue
                if (d, r["source"], r["produit_id"]) in vus:
                    j.add("historique_prix.csv", f"ligne {n}", "ERREUR", "doublon (date, source, produit)", "ignoré")
                    continue
                vus.add((d, r["source"], r["produit_id"]))
                lignes.append((d, r["source"], r["produit_id"], prix))
        j.add("historique_prix.csv", "fichier", "INFO",
              f"{len(lignes)} lignes valides, aucun trou ni doublon ; Comptoir Ivoire = 1 prix/semaine (lundi), prix catalogue sans remise ; "
              "Baobab = amount sans livraison", "prix en vigueur reporté jusqu'au changement suivant (Comptoir)")
        return cls(lignes)

    def prix_a(self, source, pid, d):
        """Prix en vigueur à la date d (dernier prix publié à cette date ou avant) : O(log m)."""
        k = (source, pid)
        if k not in self.dates:
            return None
        i = bisect.bisect_right(self.dates[k], d) - 1
        return self.prix[k][i] if i >= 0 else None

    def moyenne(self, source, pid, d0, d1):
        vals, d = [], d0
        while d <= d1:
            p = self.prix_a(source, pid, d)
            if p is not None:
                vals.append(p)
            d += timedelta(days=1)
        return sum(vals) / len(vals) if vals else None

    def prix_a_naif(self, source, pid, d):
        """Même résultat par balayage complet de la liste : O(n). Sert uniquement au benchmark."""
        meilleur = None
        for dd, s, p, prix in self.lignes:
            if s == source and p == pid and dd <= d and (meilleur is None or dd > meilleur[0]):
                meilleur = (dd, prix)
        return meilleur[1] if meilleur else None


def controler_avec_historique(offres, hist, j):
    """Compare le prix affiché au prix du 31/10/2026 de l'historique : repère les prix aberrants (zéro en trop)."""
    for o in offres:
        if o.prix_affiche is None:
            continue
        ref = hist.prix_a(o.source, o.produit_id, DATE_FICHIERS)
        if ref is None:
            continue
        r = o.prix_affiche / ref
        if r >= 3 or r <= 1 / 3:
            cand = [c for c in (o.prix_affiche / 10, o.prix_affiche * 10) if abs(c / ref - 1) <= 0.10]
            if cand:
                j.add("contrôle historique", f"{o.source} {o.ref} ({o.produit_id})", "ERREUR",
                      f"prix aberrant {fr(o.prix_affiche)} FCFA (historique au 31/10 : {fr(ref)})",
                      f"corrigé en {fr(cand[0])} (facteur 10, confirmé par l'historique)")
                o.prix_affiche = int(cand[0])
            else:
                j.add("contrôle historique", f"{o.source} {o.ref} ({o.produit_id})", "ERREUR",
                      f"prix aberrant {fr(o.prix_affiche)} (historique : {fr(ref)})", "prix exclu")
                o.prix_affiche = None
        elif o.prix_affiche != ref:
            j.add("contrôle historique", f"{o.source} {o.ref} ({o.produit_id})", "INFO",
                  f"prix affiché {fr(o.prix_affiche)} ≠ historique du 31/10 {fr(ref)}", "prix du fichier conservé")


def analyse_historique(hist, catalogue):
    res = []
    for (s, pid) in sorted(hist.dates, key=lambda k: (SOURCES.index(k[0]), k[1])):
        an = hist.moyenne(s, pid, AN_DEBUT, AN_FIN)
        bf = hist.moyenne(s, pid, BF_DEBUT, BF_FIN)
        prix, dates = hist.prix[(s, pid)], hist.dates[(s, pid)]
        i = prix.index(min(prix))
        res.append(dict(source=s, pid=pid, nom=catalogue[pid]["nom"], an=an, bf=bf, ecart=(bf / an - 1) * 100,
                        mini=prix[i], date_mini=dates[i], bf_vs_mini=(bf / prix[i] - 1) * 100))
    return res


def indice_hebdomadaire(hist, source, moy_an):
    """Indice moyen (prix / moyenne annuelle du produit) par semaine lundi-dimanche complète."""
    pids = [p for (s, p) in hist.dates if s == source]
    par_sem, d = defaultdict(list), AN_DEBUT
    while d <= AN_FIN:
        v = [hist.prix_a(source, p, d) / moy_an[(source, p)] for p in pids if hist.prix_a(source, p, d) is not None]
        if v:
            par_sem[d - timedelta(days=d.weekday())].append(statistics.mean(v))
        d += timedelta(days=1)
    return {w: statistics.mean(v) for w, v in par_sem.items() if len(v) == 7}


def panier_historique(hist, source):
    jours, d = {}, AN_DEBUT
    while d <= AN_FIN:
        ps = {p: hist.prix_a(source, p, d) for p in PANIER}
        if all(v is not None for v in ps.values()):
            jours[d] = sum(PANIER[p] * v for p, v in ps.items())
        d += timedelta(days=1)
    return jours

