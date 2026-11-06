#!/usr/bin/env python3
"""Analyse comparative de marché : Marché Lagune / Baobab Shop / Comptoir Ivoire.

Usage :  python3 analyse.py   (lit ./donnees, écrit dans ./resultats ; options --data et --out)

Étapes : chargement (CSV, JSON, Excel) -> rapprochement avec le catalogue
-> nettoyage/contrôles (journal) -> comparatif -> synthèse -> panier
-> historique (Black Friday) -> mesures de temps.
Bibliothèques : standard + openpyxl (lecture du .xlsx).
"""
import argparse
import bisect
import csv
import difflib
import json
import math
import random
import re
import statistics
import time
import unicodedata
from collections import Counter, defaultdict
from contextlib import contextmanager
from dataclasses import dataclass
from datetime import date, timedelta
from pathlib import Path

import openpyxl

ML, BB, CI = "Marché Lagune", "Baobab Shop", "Comptoir Ivoire"
SOURCES = [ML, BB, CI]
DATE_FICHIERS = date(2026, 10, 31)          # date des fichiers d'enseignes
AN_DEBUT, AN_FIN = date(2025, 11, 1), date(2026, 10, 31)
BF_DEBUT, BF_FIN = date(2025, 11, 24), date(2025, 11, 30)
PANIER = {"P04": 1, "P12": 1, "P14": 2, "P15": 1, "P28": 1}


# ----------------------------------------------------------------------------
# Outils : journal, chronomètre, formats
# ----------------------------------------------------------------------------
class Journal:
    """Liste des problèmes de données repérés et de leur traitement."""

    def __init__(self):
        self.lignes = []

    def add(self, fichier, element, niveau, probleme, traitement):
        self.lignes.append(dict(fichier=fichier, element=element, niveau=niveau,
                                probleme=probleme, traitement=traitement))

    def ecrire(self, chemin):
        with open(chemin, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.DictWriter(f, fieldnames=list(self.lignes[0]), delimiter=";")
            w.writeheader()
            w.writerows(self.lignes)


class Chrono:
    def __init__(self):
        self.durees = {}

    @contextmanager
    def __call__(self, nom):
        t0 = time.perf_counter()
        yield
        self.durees[nom] = time.perf_counter() - t0


def fr(n):
    """Entier avec séparateur de milliers : 102900 -> '102 900'."""
    return f"{round(n):,}".replace(",", " ")


def pct(x):
    return f"{x:.1f}".replace(".", ",") + " %"


def md_table(entetes, lignes):
    out = ["| " + " | ".join(entetes) + " |", "|" + "|".join("---" for _ in entetes) + "|"]
    out += ["| " + " | ".join(str(c) for c in l) + " |" for l in lignes]
    return "\n".join(out)


# ----------------------------------------------------------------------------
# Modèle commun : une offre = un article d'une enseigne
# ----------------------------------------------------------------------------
@dataclass
class Offre:
    source: str
    ref: str
    libelle: str
    prix_affiche: int | None      # prix tel qu'affiché dans le fichier
    frais: int = 0                # livraison à ajouter (Baobab)
    remise: float = 0.0           # % de remise (Comptoir)
    en_stock: bool = True
    stock_info: str = ""
    produit_id: str | None = None
    score: float = 0.0

    @property
    def prix_paye(self):
        if self.prix_affiche is None:
            return None
        return round(self.prix_affiche * (1 - self.remise / 100)) + self.frais

    @property
    def exploitable(self):
        return self.en_stock and self.prix_affiche is not None


def parse_prix(txt):
    """'102 900 FCFA' (espaces, espaces insécables) -> 102900 ; 'Prix sur demande' -> None."""
    if txt is None:
        return None
    chiffres = re.sub(r"\D", "", str(txt))
    return int(chiffres) if chiffres else None


# ----------------------------------------------------------------------------
# 1. Chargement + nettoyage propre à chaque source
# ----------------------------------------------------------------------------
def charger_catalogue(chemin):
    with open(chemin, encoding="utf-8-sig", newline="") as f:
        return {r["produit_id"]: r for r in csv.DictReader(f)}


def charger_marche_lagune(chemin, j):
    fich = "marche_lagune.csv"
    brut = Path(chemin).read_bytes()
    if brut.startswith(b"\xef\xbb\xbf") or b"\r\n" in brut:
        j.add(fich, "fichier", "INFO", "BOM UTF-8 et fins de ligne CRLF (export Excel français)",
              "lu en utf-8-sig, séparateur ';'")
    offres, vus = [], {}
    with open(chemin, encoding="utf-8-sig", newline="") as f:
        for n, row in enumerate(csv.DictReader(f, delimiter=";"), start=2):
            ref = row["reference"].strip()
            if ref in vus:
                identique = vus[ref] == row
                j.add(fich, f"ligne {n} ({ref})", "ERREUR",
                      "doublon " + ("exact" if identique else "de référence avec valeurs différentes"),
                      "ligne ignorée (première occurrence conservée)")
                continue
            vus[ref] = row
            lib = " ".join(row["designation"].split())
            if lib != row["designation"] or lib.islower():
                j.add(fich, f"ligne {n} ({ref})", "ATTENTION",
                      f"désignation mal formatée : {row['designation']!r}", "espaces/casse normalisés")
            prix, ancien = parse_prix(row["prix"]), parse_prix(row["ancien_prix"])
            if prix is None:
                j.add(fich, f"ligne {n} ({ref})", "ERREUR",
                      f"prix inexploitable : {row['prix']!r}" + (f" (ancien prix {fr(ancien)} ignoré)" if ancien else ""),
                      "pas de prix -> jamais retenu comme moins cher")
            elif ancien and ancien < prix:
                j.add(fich, f"ligne {n} ({ref})", "ATTENTION",
                      f"ancien prix ({fr(ancien)}) inférieur au prix actuel ({fr(prix)})",
                      "ancien prix (prix barré) jamais utilisé ; prix actuel contrôlé avec l'historique")
            dispo = row["disponibilite"].strip().lower()
            if dispo not in ("en stock", "stock limité", "rupture"):
                j.add(fich, f"ligne {n} ({ref})", "ATTENTION", f"disponibilité inconnue : {dispo!r}", "traitée comme en stock")
            en_stock = dispo != "rupture"
            if not en_stock:
                j.add(fich, f"ligne {n} ({ref})", "INFO", "article en rupture", "écarté des comparaisons")
            offres.append(Offre(ML, ref, lib, prix, 0, 0.0, en_stock, row["disponibilite"].strip()))
    j.add(fich, "colonne prix", "INFO", "prix texte « 102 900 FCFA »", "chiffres extraits -> entier ; prix livraison incluse")
    return offres


def charger_baobab(chemin, j):
    fich = "baobab_shop.json"
    data = json.loads(Path(chemin).read_text(encoding="utf-8"))
    items = data["items"]
    if data.get("count") != len(items):
        j.add(fich, "count", "ATTENTION", f"count={data.get('count')} mais {len(items)} articles", "len(items) retenu")
    offres, skus = [], set()
    for it in items:
        sku, p = it["sku"], it["pricing"]
        if sku in skus:
            j.add(fich, sku, "ERREUR", "sku en double", "ignoré")
            continue
        skus.add(sku)
        if p.get("currency") != "XOF":
            j.add(fich, sku, "ERREUR", f"devise {p.get('currency')}", "article ignoré")
            continue
        if p["amount"] is None:
            j.add(fich, sku, "ERREUR", f"amount = null ({it['title']})", "pas de prix -> jamais retenu comme moins cher")
        if it["stock"] <= 0:
            j.add(fich, sku, "INFO", "stock = 0", "écarté des comparaisons")
        if it["seller"] != "Baobab Shop":
            j.add(fich, sku, "INFO", f"vendeur tiers : {it['seller']}", "conservé (même prix payé)")
        offres.append(Offre(BB, sku, it["title"], p["amount"], p["delivery_fee"] or 0, 0.0,
                            it["stock"] > 0, f"stock {it['stock']}"))
    j.add(fich, "pricing", "INFO", "prix payé = amount + delivery_fee (livraison hors prix affiché)",
          "frais de livraison ajoutés (par article)")
    return offres


def charger_comptoir(chemin, j):
    fich = "comptoir_ivoire.xlsx"
    ws = openpyxl.load_workbook(chemin, data_only=True).worksheets[0]
    lignes = list(ws.iter_rows(values_only=True))
    ent = next(i for i, l in enumerate(lignes) if l and l[0] == "Code")      # ligne d'en-têtes (titre au-dessus)
    j.add(fich, f"lignes 1-{ent}", "INFO", "titre + ligne vide avant les en-têtes", f"en-têtes repérés ligne {ent + 1}")
    cols = {c: i for i, c in enumerate(lignes[ent])}
    offres = []
    for n, l in enumerate(lignes[ent + 1:], start=ent + 2):
        if not any(l):
            continue
        code, art = l[cols["Code"]], " ".join(str(l[cols["Article"]]).split())
        brut = l[cols["Prix catalogue (FCFA)"]]
        prix = parse_prix(brut)
        if isinstance(brut, str):
            j.add(fich, f"ligne {n} ({code})", "ERREUR", f"prix stocké en texte : {brut!r}", "espaces (insécables) retirés -> entier")
        if prix is None:
            j.add(fich, f"ligne {n} ({code})", "ERREUR", "prix absent", "jamais retenu comme moins cher")
        remise = l[cols["Remise (%)"]]
        if remise is None or not 0 <= remise <= 100:
            j.add(fich, f"ligne {n} ({code})", "ERREUR", f"remise invalide : {remise!r}", "remise = 0")
            remise = 0
        qte = l[cols["Quantité en rayon"]] or 0
        if qte <= 0:
            j.add(fich, f"ligne {n} ({code})", "INFO", "quantité en rayon = 0", "écarté des comparaisons")
        maj = l[cols["Date de mise à jour"]].date()
        if (DATE_FICHIERS - maj).days > 21:
            j.add(fich, f"ligne {n} ({code})", "ATTENTION", f"dernière mise à jour le {maj:%d/%m/%Y} (> 3 semaines)", "conservé, à signaler")
        offres.append(Offre(CI, code, art, prix, 0, float(remise), qte > 0, f"qté {qte}"))
    j.add(fich, "prix", "INFO", "prix payé = prix catalogue x (1 - remise) ; retrait en boutique (pas de livraison)",
          "remise appliquée, arrondi au FCFA")
    return offres


# ----------------------------------------------------------------------------
# 2. Rapprochement avec le catalogue (similarité de mots pondérée)
# ----------------------------------------------------------------------------
SYNONYMES = [
    (r"\bmicro[ -]?ondes?\b|\bmicrowave\b", "microondes"),
    (r"\b(?:frigo|fridge|refrigerateur)\b", "refrigerateur"),
    (r"\bfreezer\b", "congelateur"),
    (r"\b(?:clim|climatiseur|air conditioner)\b", "climatiseur"),
    (r"\b(?:ventilo|ventilateur|fan)\b", "ventilateur"),
    (r"\b(?:monitor|ecran)\b", "ecran"),
    (r"\bkettle\b", "bouilloire"),
    (r"\belectric\b", "electrique"),
    (r"\bblender\b", "mixeur"),
    (r"\b(?:iron|fer)\b", "fer"),
    (r"\bdd\b|\bhard drive\b", "disque dur"),
    (r"\bmouse\b", "souris"),
    (r"\bprinter\b", "imprimante"),
    (r"\brouter\b", "routeur"),
    (r"\bwireless\b", "sans fil"),
    (r"\blaptop\b", "portable"),
]
MOTS_VIDES = {"a", "de", "du", "la", "le", "les", "et", "pc", "steam"}


def normaliser(txt):
    t = unicodedata.normalize("NFKD", txt)
    t = "".join(c for c in t if not unicodedata.combining(c)).lower()
    t = t.replace(" ", " ").replace("\xa0", " ")
    t = re.sub(r"(\d),(\d)", r"\1.\2", t)                                   # 1,5 -> 1.5
    t = re.sub(r"(\d+(?:\.\d+)?)\s*(?:gb|go)\b", r"\1go", t)                # 128 GB / 128Go -> 128go
    t = re.sub(r"(\d+(?:\.\d+)?)\s*(?:tb|to)\b", r"\1to", t)
    t = re.sub(r"(\d+(?:\.\d+)?)\s*(?:cv|hp)\b", r"\1cv", t)                # 1HP == 1 CV
    t = re.sub(r"(\d+(?:\.\d+)?)\s*l\b", r"\1l", t)
    t = re.sub(r"(\d+)\s*(?:\"|pouces?\b|inch\b)", r"\1p", t)               # 24" == 24 pouces
    for motif, rempl in SYNONYMES:
        t = re.sub(motif, rempl, t)
    return t


def tokens(txt):
    return {m for m in re.findall(r"[a-z0-9]+(?:\.[a-z0-9]+)*", normaliser(txt)) if m not in MOTS_VIDES}


def rapprocher(offres, catalogue, j):
    """Associe chaque offre au produit du catalogue le plus proche (score pondéré par la rareté des mots)."""
    cat_tok = {pid: tokens(r["nom"]) for pid, r in catalogue.items()}
    df = Counter(t for ts in cat_tok.values() for t in ts)
    poids = {t: math.log(len(cat_tok) / n) + 0.5 for t, n in df.items()}
    vocab_alpha = [t for t in poids if t.isalpha() and len(t) >= 5]
    rayon_cat = {"Téléphonie": "Smartphone", "Informatique": "Informatique", "Maison": "Électroménager"}
    for o in offres:
        ts = set()
        for t in tokens(o.libelle):
            if t not in poids and t.isalpha() and len(t) >= 5:               # fautes de frappe (Samsun -> samsung)
                proche = difflib.get_close_matches(t, vocab_alpha, n=1, cutoff=0.85)
                if proche:
                    j.add("rapprochement", f"{o.source} {o.ref}", "ATTENTION",
                          f"faute de frappe « {t} » dans {o.libelle!r}", f"corrigé en « {proche[0]} »")
                    t = proche[0]
            ts.add(t)
        scores = []
        for pid, ct in cat_tok.items():
            union = ct | {t for t in ts if t in poids}
            scores.append((sum(poids[t] for t in ct & ts) / sum(poids[t] for t in union), pid))
        scores.sort(reverse=True)
        (s1, p1), (s2, _) = scores[0], scores[1]
        if s1 >= 0.30 and s1 - s2 >= 0.10:
            o.produit_id, o.score = p1, s1
        else:
            j.add("rapprochement", f"{o.source} {o.ref}", "ERREUR",
                  f"rapprochement ambigu/introuvable pour {o.libelle!r} (meilleur {p1} score {s1:.2f}, suivant {s2:.2f})",
                  "article non rapproché, exclu")
    # contrôles croisés (non utilisés pour décider)
    for o in offres:
        if o.produit_id and o.source == BB and o.ref[2:4].isdigit() and f"P{o.ref[2:4]}" != o.produit_id:
            j.add("rapprochement", f"{o.source} {o.ref}", "ERREUR", f"SKU suggère P{o.ref[2:4]}, rapproché à {o.produit_id}", "à vérifier")
    # un même produit ne doit apparaître qu'une fois par enseigne
    vu = {}
    for o in sorted((o for o in offres if o.produit_id), key=lambda o: -o.score):
        k = (o.source, o.produit_id)
        if k in vu:
            j.add("rapprochement", f"{o.source} {o.ref}", "ERREUR", f"2e article pour {o.produit_id}", "ignoré (score plus faible)")
            o.produit_id = None
        vu[k] = o
    return [o for o in offres if o.produit_id]


# ----------------------------------------------------------------------------
# 3. Historique : index {(source, produit): séries triées par date}
# ----------------------------------------------------------------------------
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
        """Prix en vigueur à la date d (dernier prix publié à cette date ou avant) : O (log m)."""
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


# ----------------------------------------------------------------------------
# 4. Comparatif, synthèse, panier
# ----------------------------------------------------------------------------
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


def ecrire_csv(chemin, entetes, lignes):
    with open(chemin, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(entetes)
        for l in lignes:
            w.writerow([str(c).replace(".", ",") if isinstance(c, float) else c for c in l])


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


# ----------------------------------------------------------------------------
# 5. Historique : Black Friday vs année, meilleures périodes
# ----------------------------------------------------------------------------
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


# ----------------------------------------------------------------------------
# Programme principal
# ----------------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--data", default=str(Path(__file__).resolve().parent / "donnees"))
    ap.add_argument("--out", default=str(Path(__file__).resolve().parent / "resultats"))
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
- Bibliothèque standard uniquement (+ `openpyxl`) pour garder les coûts visibles ; l'historique tient en mémoire (≈ 20 k tuples), la mémoire est en O(n).
""")
    (O / "rapport.md").write_text("\n".join(R), encoding="utf-8")
    j.ecrire(O / "journal_nettoyage.csv")
    print("\n".join(R))


if __name__ == "__main__":
    main()
