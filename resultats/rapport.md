# Analyse comparative — Marché Lagune / Baobab Shop / Comptoir Ivoire

_Données fictives. Prix en FCFA, **prix réellement payés** : Marché Lagune = prix affiché (livraison incluse) ; Baobab Shop = prix + frais de livraison ; Comptoir Ivoire = prix catalogue après remise, retrait en boutique (transport du client non compté). Un article en rupture ou sans prix n'est jamais retenu._

## 1. Chargement, nettoyage, rapprochement

76 lignes lues dans les 3 enseignes -> 76 articles rapprochés d'un produit du catalogue. Journal complet : `journal_nettoyage.csv` (7 erreurs, 8 avertissements, 25 infos). Points principaux :

| Fichier | Élément | Problème | Traitement |
|---|---|---|---|
| marche_lagune.csv | ligne 14 (ML-1000) | doublon exact | ligne ignorée (première occurrence conservée) |
| marche_lagune.csv | ligne 20 (ML-1119) | prix inexploitable : 'Prix sur demande' (ancien prix 110 400 ignoré) | pas de prix -> jamais retenu comme moins cher |
| marche_lagune.csv | ligne 27 (ML-1168) | prix inexploitable : '' | pas de prix -> jamais retenu comme moins cher |
| marche_lagune.csv | ligne 31 (ML-1147) | doublon exact | ligne ignorée (première occurrence conservée) |
| baobab_shop.json | BB27999 | amount = null (Philips Steam Iron) | pas de prix -> jamais retenu comme moins cher |
| comptoir_ivoire.xlsx | ligne 20 (CI216) | prix stocké en texte : '279\u202f300' | espaces (insécables) retirés -> entier |
| contrôle historique | Marché Lagune ML-1021 (P04) | prix aberrant 779 000 FCFA (historique au 31/10 : 77 900) | corrigé en 77 900 (facteur 10, confirmé par l'historique) |

Autres problèmes traités (avertissements) : ligne 5 (ML-1021) — ancien prix (88 700) inférieur au prix actuel (779 000); ligne 18 (ML-1105) — désignation mal formatée : '  souris sans fil logitech m185 '; Comptoir Ivoire CI200 — faute de frappe « samsun » dans 'Samsun Galaxy A15 128Go'.

**Rapprochement** : les libellés sont normalisés (accents, casse, unités `128GB`=`128 Go`, `1HP`=`1 CV`, `24"`=`24 pouces`, synonymes `Frigo`/`Fridge`/`Réfrigérateur`, etc.), découpés en mots, puis comparés à chaque produit du catalogue par une similarité pondérée (un mot rare, comme `A25`, pèse plus qu'un mot courant). Un article est accepté si son score ≥ 0,30 et s'il bat le 2e candidat d'au moins 0,10 ; les fautes (`Samsun`) sont corrigées par proximité orthographique. Les SKU Baobab servent de contrôle croisé (0 désaccord). Détail : `rapprochement.csv`.

## 2. Comparatif

23 produits sur 30 disponibles dans au moins deux enseignes (`comparatif.csv`). Écart = prix le plus élevé − prix le plus bas (en FCFA et en % du prix le plus bas).

| Produit | Marché Lagune | Baobab Shop | Comptoir Ivoire | Moins cher | Écart FCFA | Écart % |
|---|---|---|---|---|---|---|
| P01 Samsung Galaxy A15 128 Go | — | 104 000 | 92 990 | Comptoir Ivoire | 11 010 | 11,8 % |
| P02 Samsung Galaxy A25 128 Go | 151 500 | 142 000 | 150 300 | Baobab Shop | 9 500 | 6,7 % |
| P04 Tecno Spark 20 128 Go | 77 900 | 75 000 | 76 300 | Baobab Shop | 2 900 | 3,9 % |
| P05 Tecno Camon 30 256 Go | 136 900 | 135 400 | — | Baobab Shop | 1 500 | 1,1 % |
| P06 Infinix Hot 40i 128 Go | 69 200 | 69 500 | 67 165 | Comptoir Ivoire | 2 335 | 3,5 % |
| P07 Infinix Note 40 256 Go | 132 400 | 132 700 | — | Marché Lagune | 300 | 0,2 % |
| P09 Apple iPhone 13 128 Go | 433 700 | 417 200 | 443 700 | Baobab Shop | 26 500 | 6,4 % |
| P10 Xiaomi Redmi Note 13 128 Go | 136 900 | 134 100 | — | Baobab Shop | 2 800 | 2,1 % |
| P11 HP 250 G9 Core i5 8 Go 512 Go | 390 800 | 354 200 | 407 000 | Baobab Shop | 52 800 | 14,9 % |
| P12 Lenovo IdeaPad 3 Core i3 8 Go 256 Go | 309 400 | 298 200 | 287 375 | Comptoir Ivoire | 22 025 | 7,7 % |
| P14 Clé USB SanDisk 64 Go | 6 100 | 7 300 | 6 200 | Marché Lagune | 1 200 | 19,7 % |
| P15 Disque dur externe Seagate 1 To | 37 600 | 38 300 | 39 500 | Marché Lagune | 1 900 | 5,1 % |
| P16 Souris sans fil Logitech M185 | 9 200 | — | 9 200 | Marché Lagune / Comptoir Ivoire | 0 | 0,0 % |
| P17 Imprimante HP DeskJet 2720e | 44 000 | 45 700 | — | Marché Lagune | 1 700 | 3,9 % |
| P20 Réfrigérateur Hisense 205 L | 186 100 | 183 000 | 168 385 | Comptoir Ivoire | 17 715 | 10,5 % |
| P21 Congélateur Nasco 200 L | 176 000 | — | 151 725 | Comptoir Ivoire | 24 275 | 16,0 % |
| P22 TV Samsung 43 pouces Crystal UHD | 251 100 | 244 400 | — | Baobab Shop | 6 700 | 2,7 % |
| P23 TV Hisense 32 pouces | — | 96 000 | 88 740 | Comptoir Ivoire | 7 260 | 8,2 % |
| P24 Climatiseur LG 1 CV | 256 600 | 253 300 | 279 300 | Baobab Shop | 26 000 | 10,3 % |
| P26 Mixeur Moulinex 1,5 L | — | 30 900 | 29 200 | Comptoir Ivoire | 1 700 | 5,8 % |
| P27 Fer à repasser Philips | 15 700 | — | 15 700 | Marché Lagune / Comptoir Ivoire | 0 | 0,0 % |
| P28 Micro-ondes Samsung 23 L | 67 200 | 65 400 | 68 500 | Baobab Shop | 3 100 | 4,7 % |
| P29 Bouilloire électrique 1,7 L | — | 11 600 | 9 200 | Comptoir Ivoire | 2 400 | 26,1 % |

## 3. Synthèse

**Combien de fois chaque enseigne est-elle la moins chère ?** (un ex æquo compte pour chaque enseigne à égalité)

| Enseigne | Moins chère (total) | dont seule | dont ex æquo | Sur n produits comparables |
|---|---|---|---|---|
| Marché Lagune | 6 | 4 | 2 | 19 |
| Baobab Shop | 9 | 9 | 0 | 20 |
| Comptoir Ivoire | 10 | 8 | 2 | 18 |

**Écart moyen par rapport au prix le plus bas** (0 quand l'enseigne est la moins chère ; moyenne sur les produits où elle a un prix retenu) :

| Enseigne | Écart moyen % | Écart moyen FCFA | n |
|---|---|---|---|
| Marché Lagune | 3,8 % | 7 771 | 19 |
| Baobab Shop | 4,7 % | 2 702 | 20 |
| Comptoir Ivoire | 2,8 % | 6 667 | 18 |

**Les 3 produits où il y a le plus à gagner à comparer** (classés par écart en FCFA) :

| Produit | Moins cher | Plus cher | Écart FCFA | Écart % |
|---|---|---|---|---|
| P11 HP 250 G9 Core i5 8 Go 512 Go | Baobab Shop (354 200) | Comptoir Ivoire (407 000) | 52 800 | 14,9 % |
| P09 Apple iPhone 13 128 Go | Baobab Shop (417 200) | Comptoir Ivoire (443 700) | 26 500 | 6,4 % |
| P24 Climatiseur LG 1 CV | Baobab Shop (253 300) | Comptoir Ivoire (279 300) | 26 000 | 10,3 % |

Classés par écart en % : P29 Bouilloire électrique 1,7 L (26,1 %), P14 Clé USB SanDisk 64 Go (19,7 %), P21 Congélateur Nasco 200 L (16,0 %).

## 4. Panier de rentrée

1 Tecno Spark 20, 1 Lenovo IdeaPad 3, 2 clés USB SanDisk, 1 disque dur Seagate, 1 micro-ondes Samsung (frais de livraison Baobab comptés par article, donc × 2 pour les clés USB).

| Article | Qté | Marché Lagune | Baobab Shop | Comptoir Ivoire | Moins cher |
|---|---|---|---|---|---|
| P04 Tecno Spark 20 128 Go | 1 | 77 900 | 75 000 | 76 300 | Baobab Shop |
| P12 Lenovo IdeaPad 3 Core i3 8 Go 256 Go | 1 | 309 400 | 298 200 | 287 375 | Comptoir Ivoire |
| P14 Clé USB SanDisk 64 Go | 2 | 6 100 | 7 300 | 6 200 | Marché Lagune |
| P15 Disque dur externe Seagate 1 To | 1 | 37 600 | 38 300 | 39 500 | Marché Lagune |
| P28 Micro-ondes Samsung 23 L | 1 | 67 200 | 65 400 | 68 500 | Baobab Shop |
| **Total panier** |  | **504 300** | **491 500** | **484 075** |  |

- Panier complet dans une seule enseigne : le moins cher est **Comptoir Ivoire** à **484 075 FCFA** ; Marché Lagune : 504 300 (+20 225) ; Baobab Shop : 491 500 (+7 425).
- En achetant chaque article là où il est le moins cher : **477 575 FCFA**, soit 6 500 FCFA (1,3 %) de moins que le meilleur panier unique. Contrepartie : plusieurs enseignes, donc plusieurs déplacements/livraisons.

## 5. Historique : quand fallait-il acheter ?

Prix **affichés** de l'historique (Baobab sans livraison, Comptoir sans remise) ; la comparaison se fait enseigne par enseigne contre sa propre moyenne annuelle (01/11/2025–31/10/2026), donc ces biais ne la faussent pas. Comptoir Ivoire change de prix le lundi : le prix en vigueur est reporté jusqu'au changement suivant, ce qui donne bien une moyenne pondérée par jour. Détail par produit : `black_friday.csv`.

| Enseigne | Produits | Semaine BF vs moyenne annuelle (moyenne des écarts) | Produits moins chers en BF | BF vs plus bas prix de l'année |
|---|---|---|---|---|
| Marché Lagune | 28 | -11,3 % | 28 / 28 | 1,5 % |
| Baobab Shop | 26 | -11,7 % | 26 / 26 | 1,2 % |
| Comptoir Ivoire | 22 | 3,7 % | 0 / 22 | 10,5 % |

**Semaines les moins chères** (indice 100 % = moyenne annuelle ; semaines lundi-dimanche) :

| Enseigne | 1re | 2e | 3e | Semaine du 24/11/2025 |
|---|---|---|---|---|
| Marché Lagune | 24/11/2025 (88,7 %) | 14/09/2026 (94,6 %) | 07/09/2026 (94,6 %) | 88,7 % (rang 1/51) |
| Baobab Shop | 24/11/2025 (88,3 %) | 14/09/2026 (94,8 %) | 07/09/2026 (95,0 %) | 88,3 % (rang 1/51) |
| Comptoir Ivoire | 14/09/2026 (94,7 %) | 07/09/2026 (94,9 %) | 19/10/2026 (96,0 %) | 103,7 % (rang 48/51) |

**Mois du prix le plus bas de l'année** (nombre de produits dont le minimum tombe ce mois-là) :

| Enseigne | Top 3 des mois |
|---|---|
| Marché Lagune | 2025-11 (23), 2026-09 (5) |
| Baobab Shop | 2025-11 (23), 2026-09 (3) |
| Comptoir Ivoire | 2026-10 (16), 2026-09 (6) |

**Le panier de rentrée dans l'historique** (prix affichés) :

| Enseigne | Moyenne annuelle | Semaine BF | BF vs année | Plus bas de l'année | Au 31/10/2026 |
|---|---|---|---|---|---|
| Marché Lagune | 530 746 | 475 757 | -10,4 % | 472 100 le 25/11/2025 | 504 300 |
| Baobab Shop | 502 671 | 446 329 | -11,2 % | 441 400 le 30/11/2025 | 483 000 |
| Comptoir Ivoire | 522 376 | 544 200 | 4,2 % | 476 700 le 14/09/2026 | 499 200 |

## 6. Complexité et structures de données

Historique : 20876 lignes, 76 séries (enseigne, produit) de ≤ 365 points.

| Étape | Temps (s) |
|---|---|
| Chargement des 3 enseignes (CSV + JSON + Excel) | 0.0299 |
| Chargement historique + construction de l'index | 0.0604 |
| Rapprochement avec le catalogue | 0.0112 |
| Contrôle des prix avec l'historique | 0.0001 |
| Comparatif + synthèse + panier | 0.0004 |
| Analyse de l'historique (Black Friday, semaines, panier) | 0.0858 |
| TOTAL (hors benchmark) | 0.1989 |

**Benchmark** : 2000 recherches « prix de tel produit chez telle enseigne à telle date » (résultats identiques, vérifié) :

| Méthode | Complexité par recherche | Temps pour 2000 recherches |
|---|---|---|
| Balayage de la liste des 20 876 lignes | O(n) | 1.687 s |
| Index dict + liste triée + bisect | O(log m), m ≤ 365 | 0.00099 s |

L'index est environ **1 711 fois plus rapide**. Choix retenus :

- **`dict` clé `(enseigne, produit_id)` -> séries** : accès direct en O(1) à la bonne série au lieu de filtrer 20 876 lignes ; construit en une passe O(n).
- **Séries triées par date + `bisect`** : trouver le prix à une date coûte O(log m) (m ≤ 365) ; trier une fois coûte O(n log m) au chargement et sert ensuite à toutes les requêtes. Ça gère aussi Comptoir Ivoire (1 prix/semaine) : `bisect_right − 1` renvoie le dernier prix publié, sans stocker de lignes par jour.
- **`dict` produit_id -> {enseigne: offre}** pour le comparatif : accès O(1), comparatif en O(P × S) = 30 × 3.
- **Rapprochement** : 76 articles × 30 produits × quelques mots = quelques milliers de comparaisons, donc un simple balayage suffit ; avec un catalogue de 100 000 produits il faudrait un index inversé (mot -> produits) pour ne comparer que les candidats.
- **Ensembles (`set`) de mots** pour l'intersection/union des libellés : O(taille du plus petit).
- Bibliothèque standard uniquement (`openpyxl` facultatif) pour garder les coûts visibles ; l'historique tient en mémoire (≈ 20 k tuples), la mémoire est en O(n).
