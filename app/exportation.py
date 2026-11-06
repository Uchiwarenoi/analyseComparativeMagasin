import csv

def ecrire_csv(chemin, entetes, lignes):
    with open(chemin, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(entetes)
        for l in lignes:
            w.writerow([str(c).replace(".", ",") if isinstance(c, float) else c for c in l])

