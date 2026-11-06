"""Normalisation des libellés et rapprochement avec le catalogue."""
import re
import unicodedata
import difflib
import math
from collections import Counter

from app.common import BB

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

