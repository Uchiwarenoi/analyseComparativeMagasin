import csv
import re
import time
from contextlib import contextmanager
from dataclasses import dataclass

from datetime import date

ML, BB, CI = "Marché Lagune", "Baobab Shop", "Comptoir Ivoire"
SOURCES = [ML, BB, CI]
DATE_FICHIERS = date(2026, 10, 31)
AN_DEBUT, AN_FIN = date(2025, 11, 1), date(2026, 10, 31)
BF_DEBUT, BF_FIN = date(2025, 11, 24), date(2025, 11, 30)
PANIER = {"P04": 1, "P12": 1, "P14": 2, "P15": 1, "P28": 1}

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

