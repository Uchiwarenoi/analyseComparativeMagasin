# Analyse comparative de marché

**Projet Python — L2 Informatique**

Ce programme réalise une analyse comparative des prix de trois enseignes fictives : **Marché Lagune**, **Baobab Shop** et **Comptoir Ivoire**. Il charge leurs données depuis trois formats différents, rapproche les articles d'un catalogue commun, compare les prix réellement payés, étudie un panier de rentrée et analyse un historique de prix.

> **Important :** toutes les données de l'exercice sont fictives. Les enseignes et les prix ne représentent pas des commerces ou des tarifs réels.

---

## 1. Objectifs du projet

Le programme répond aux huit travaux demandés dans le sujet :

1. **Chargement des données** : lire les fichiers CSV, JSON et Excel dans une structure commune.
2. **Nettoyage** : repérer les anomalies, appliquer un traitement adapté et enregistrer les problèmes dans un journal.
3. **Rapprochement** : associer les articles des enseignes aux 30 références du catalogue (`P01` à `P30`).
4. **Comparatif** : comparer les prix payés pour les produits présents dans au moins deux enseignes, identifier le ou les prix les plus bas et calculer les écarts en FCFA et en pourcentage.
5. **Synthèse** : compter les victoires de chaque enseigne, calculer les écarts moyens et identifier les trois produits pour lesquels comparer les prix est le plus intéressant.
6. **Panier de rentrée** : calculer le coût du panier dans chaque enseigne et le coût d'un panier composé en choisissant, pour chaque article, l'offre exploitable la moins chère.
7. **Historique** : comparer les prix de la semaine du Black Friday (du 24 au 30 novembre 2025) à la moyenne annuelle, puis étudier les semaines et périodes les plus intéressantes.
8. **Complexité et performances** : mesurer les temps d'exécution et comparer une recherche naïve à une recherche utilisant un index.

Un produit en rupture de stock ou sans prix exploitable ne doit jamais être retenu comme offre la moins chère.

---

## 2. Présentation des données

Les fichiers d'entrée doivent être placés dans le dossier `donnees/`.

| Fichier | Format | Description |
|---|---|---|
| `catalogue_reference.csv` | CSV | Catalogue commun de 30 produits : identifiant, nom, catégorie et marque. |
| `marche_lagune.csv` | CSV | Offres de Marché Lagune. Le séparateur est le point-virgule (`;`). |
| `baobab_shop.json` | JSON | Articles retournés par l'API fictive de Baobab Shop. |
| `comptoir_ivoire.xlsx` | Excel | Tarifs de Comptoir Ivoire. Les en-têtes se trouvent à la troisième ligne de la feuille. |
| `historique_prix.csv` | CSV | Historique quotidien des prix du 1er novembre 2025 au 31 octobre 2026, soit plus de 20 000 lignes. |

### Règles de calcul des prix payés

Les prix ne sont pas calculés de la même manière selon l'enseigne :

- **Marché Lagune** : le prix indiqué inclut déjà la livraison.
- **Baobab Shop** : prix payé = prix de l'article + frais de livraison.
- **Comptoir Ivoire** : prix payé = prix catalogue diminué de la remise. Le retrait se fait en boutique ; le transport du client n'est pas inclus.

Dans le fichier `historique_prix.csv`, les prix sont des prix affichés : l'historique de Baobab Shop ne comprend pas la livraison et celui de Comptoir Ivoire ne comprend pas la remise. L'analyse historique compare chaque enseigne à sa propre moyenne annuelle.

---

## 3. Structure du projet

L'application est organisée en plusieurs fichiers afin de séparer les responsabilités.

```text
Analyse comparative Magasin/
│
├── main.py
├── requirements.txt
│
├── app/
│   ├── __init__.py
│   ├── common.py
│   ├── chargement.py
│   ├── rapprochement.py
│   ├── comparaison.py
│   ├── historique.py
│   ├── exportation.py
│   └── interface.py
│
├── donnees/
│   ├── catalogue_reference.csv
│   ├── marche_lagune.csv
│   ├── baobab_shop.json
│   ├── comptoir_ivoire.xlsx
│   └── historique_prix.csv
│
└── resultats/                 # créé automatiquement à l'exécution
```

### Rôle des fichiers Python

| Fichier | Rôle |
|---|---|
| `main.py` | Point d'entrée : organise les étapes, lance les traitements et génère le rapport. |
| `app/common.py` | Définit les constantes, la classe `Offre`, le journal des anomalies, le chronomètre et les fonctions communes. |
| `app/chargement.py` | Charge le catalogue et les trois fichiers des enseignes ; traite les premières anomalies de lecture. |
| `app/rapprochement.py` | Normalise les libellés et associe les articles aux identifiants du catalogue. |
| `app/comparaison.py` | Produit le comparatif, les statistiques et les calculs du panier. |
| `app/historique.py` | Charge et indexe l'historique, contrôle les prix et réalise les analyses temporelles. |
| `app/exportation.py` | Écrit les tableaux de résultats au format CSV. |
| `app/interface.py` | Affiche le menu de consultation du rapport dans la console. |

---

## 4. Installation

### Prérequis

- Python **3.8 ou une version supérieure**.
- `pip`, le gestionnaire de paquets Python.
- Un terminal (PowerShell, CMD ou terminal intégré de PyCharm).
- Les fichiers de données fournis avec le sujet.

Vérifier la version de Python :

```bash
python --version
```

### Ouvrir le projet

Extraire l'archive du projet, puis ouvrir le dossier `Analyse comparative Magasin` dans PyCharm ou dans un éditeur Python.

Vérifier que le dossier `donnees/` se trouve au même niveau que `main.py`.

### Installer la dépendance Excel

Le projet utilise la bibliothèque `openpyxl` pour lire le fichier Excel `.xlsx`.

```bash
python -m pip install --upgrade pip
python -m pip install openpyxl
```

Si un fichier `requirements.txt` est fourni et contient la dépendance `openpyxl`, il est également possible d'utiliser :

```bash
python -m pip install -r requirements.txt
```

> `openpyxl` est nécessaire pour lire le classeur Excel. Les autres traitements utilisent principalement la bibliothèque standard de Python.

### (Facultatif) Créer un environnement virtuel

Sous Windows :

```bash
python -m venv .venv
.venv\Scripts\activate
python -m pip install openpyxl
```

Pour quitter l'environnement virtuel :

```bash
deactivate
```

---

## 5. Exécution du programme

Depuis le dossier qui contient `main.py`, lancer :

```bash
python main.py
```

Le programme effectue les traitements, crée le dossier `resultats/`, écrit les fichiers de sortie, puis affiche un menu de consultation dans la console.

### Menu de consultation

Le menu permet de consulter les différentes parties du rapport :

| Choix | Contenu |
|---:|---|
| 1 | Chargement, nettoyage et rapprochement |
| 2 | Comparatif des prix |
| 3 | Synthèse statistique |
| 4 | Panier de rentrée |
| 5 | Historique : Black Friday et meilleures périodes |
| 6 | Complexité et temps d'exécution |
| 0 | Quitter le programme |

### Exécuter sans afficher le menu

Pour lancer l'analyse et générer les fichiers sans ouvrir le menu :

```bash
python main.py --sans-menu
```

### Choisir un autre dossier de données

L'option `--data` permet d'indiquer le chemin du dossier contenant les fichiers d'entrée :

```bash
python main.py --data "C:\chemin\vers\donnees"
```

### Choisir un autre dossier de résultats

L'option `--out` permet de choisir où enregistrer les fichiers générés :

```bash
python main.py --out "C:\chemin\vers\resultats"
```

Les options peuvent être combinées :

```bash
python main.py --data "C:\chemin\vers\donnees" --out "C:\chemin\vers\resultats" --sans-menu
```

Les chemins doivent désigner des dossiers existants pour les données et accessibles en écriture pour les résultats. Le dossier de sortie est créé automatiquement s'il n'existe pas.

---

## 6. Résultats générés

Après une exécution réussie, le dossier `resultats/` contient notamment :

| Fichier | Contenu |
|---|---|
| `rapprochement.csv` | Détail des articles rapprochés : enseigne, libellé d'origine, identifiant du catalogue, score, prix, disponibilité et stock. |
| `comparatif.csv` | Comparaison des prix payés par produit, nombre d'enseignes, prix minimum et écarts. |
| `panier.csv` | Prix unitaires des articles du panier dans chaque enseigne et enseigne(s) la/les moins chère(s). |
| `black_friday.csv` | Comparaison détaillée, par produit et enseigne, entre la semaine du Black Friday et la moyenne annuelle. |
| `journal_nettoyage.csv` | Anomalies, avertissements et informations détectés pendant le traitement, avec leur traitement. |
| `rapport.md` | Rapport complet : résultats, tableaux, synthèses, panier, historique et mesures de performance. |

Le fichier `rapport.md` peut être ouvert dans PyCharm, Visual Studio Code ou un éditeur Markdown.

### Comprendre les écarts de prix

Pour un produit :

- **Écart en FCFA** = prix le plus élevé − prix le plus bas.
- **Écart en %** = (écart en FCFA / prix le plus bas) × 100.

En cas d'égalité au prix minimum, plusieurs enseignes peuvent être indiquées comme les moins chères.

### Comprendre le panier

Le panier demandé dans le sujet est composé de :

| Référence | Article | Quantité |
|---|---|---:|
| P04 | Tecno Spark 20 | 1 |
| P12 | Lenovo IdeaPad 3 | 1 |
| P14 | Clé USB SanDisk | 2 |
| P15 | Disque dur Seagate | 1 |
| P28 | Micro-ondes Samsung | 1 |

Le programme distingue les paniers complets des paniers incomplets. Une enseigne qui ne propose pas tous les articles ne doit pas être présentée comme ayant le panier complet le moins cher. Le panier mixte choisit, pour chaque article, l'offre exploitable la moins chère ; il peut toutefois impliquer plusieurs achats ou déplacements.

### Comprendre l'analyse historique

La période annuelle étudiée va du **1er novembre 2025 au 31 octobre 2026**. La semaine du Black Friday est définie du **24 au 30 novembre 2025**. Les prix historiques sont des prix affichés et sont analysés séparément des prix réellement payés dans le comparatif courant.

---
