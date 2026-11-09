"""Chargement des fichiers CSV, JSON et Excel des enseignes."""
import csv
import json
import re
from datetime import timedelta
from pathlib import Path

try:
    import openpyxl
except ImportError:
    openpyxl = None

from app.common import DATE_FICHIERS, ML, BB, CI, Offre, parse_prix, fr

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


def lire_xlsx_standard(chemin):
    """Lit la 1re feuille d'un .xlsx sans openpyxl (un .xlsx est un zip de fichiers XML). Retourne une liste de lignes."""
    import zipfile
    import xml.etree.ElementTree as ET
    from datetime import datetime
    ns = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
    with zipfile.ZipFile(chemin) as z:
        noms = z.namelist()
        textes = []
        if "xl/sharedStrings.xml" in noms:
            for si in ET.fromstring(z.read("xl/sharedStrings.xml")).iter(ns + "si"):
                textes.append("".join(t.text or "" for t in si.iter(ns + "t")))
        est_date = []                                   # style de cellule -> format date ?
        if "xl/styles.xml" in noms:
            st = ET.fromstring(z.read("xl/styles.xml"))
            perso = {int(n.get("numFmtId")): n.get("formatCode", "").lower() for n in st.iter(ns + "numFmt")}
            for xf in st.find(ns + "cellXfs").findall(ns + "xf"):
                i = int(xf.get("numFmtId", 0))
                code = perso.get(i, "")
                est_date.append(14 <= i <= 22 or 45 <= i <= 47 or any(c in code for c in ("yy", "dd", "mm/")))
        feuille = sorted(n for n in noms if n.startswith("xl/worksheets/sheet") and n.endswith(".xml"))[0]
        lignes = []
        for row in ET.fromstring(z.read(feuille)).iter(ns + "row"):
            cellules = {}
            for c in row.findall(ns + "c"):
                col = 0
                for ch in re.match(r"[A-Z]+", c.get("r")).group():
                    col = col * 26 + ord(ch) - 64
                t, v = c.get("t"), c.find(ns + "v")
                if t == "inlineStr":
                    val = "".join(x.text or "" for x in c.iter(ns + "t"))
                elif v is None or v.text is None:
                    val = None
                elif t == "s":
                    val = textes[int(v.text)]
                elif t in ("str", "e"):
                    val = v.text
                elif t == "b":
                    val = v.text == "1"
                else:
                    val = float(v.text)
                    val = int(val) if val == int(val) else val
                    s_ = int(c.get("s", 0))
                    if s_ < len(est_date) and est_date[s_]:
                        val = datetime(1899, 12, 30) + timedelta(days=val)
                cellules[col] = val
            while len(lignes) < int(row.get("r")) - 1:      # lignes vides absentes du XML (ex. ligne 2)
                lignes.append(())
            largeur = max(cellules, default=0)
            lignes.append(tuple(cellules.get(i) for i in range(1, largeur + 1)))
    return lignes


def charger_comptoir(chemin, j):
    fich = "comptoir_ivoire.xlsx"
    if openpyxl is not None:
        lignes = list(openpyxl.load_workbook(chemin, data_only=True).worksheets[0].iter_rows(values_only=True))
    else:
        lignes = lire_xlsx_standard(chemin)
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

