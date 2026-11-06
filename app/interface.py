"""Menu de consultation des résultats dans la console."""
import re

def menu(rapport, dossier):
    sections = {s.split(" ", 2)[1].rstrip("."): s for s in re.split(r"\n(?=## )", rapport) if s.startswith("## ")}
    entrees = [("Chargement, nettoyage et rapprochement", "1"), ("Comparatif des prix", "2"), ("Synthèse", "3"),
               ("Panier de rentrée", "4"), ("Historique : Black Friday et meilleures périodes", "5"),
               ("Complexité et temps d'exécution", "6")]
    print(f"\nAnalyse terminée. Fichiers écrits dans : {dossier}")
    while True:
        print("\n===== MENU =====")
        for i, (nom, _) in enumerate(entrees, 1):
            print(f" {i}. {nom}")
        print(" 0. Quitter")
        try:
            choix = input("Votre choix : ").strip()
        except EOFError:
            return
        if choix == "0":
            return
        if choix.isdigit() and 1 <= int(choix) <= len(entrees):
            print("\n" + sections[entrees[int(choix) - 1][1]])
        else:
            print("Choix non valide.")

