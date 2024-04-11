# Extraction de données IMARIS - Morphologie Microgliale

Extraction Data IMARIS

## Description

Ce script Python est conçu pour automatiser l'extraction, le traitement, et l'analyse des données issues de fichiers Excel générés par le logiciel IMARIS. Il permet de traiter des ensembles de données spécifiques à la recherche en neuroscience, en se concentrant sur l'analyse de variables telles que l'aire du filament, la profondeur et le niveau des branches complètes, la longueur et le volume des filaments, et d'autres mesures critiques.

## Fonctionnalités

- Extraction des données depuis des fichiers Excel format .xls
- Classification des données extraites en fonction de critères prédéfinis (ex : groupe de souris)
- Calcul des moyennes et des sommes pour des variables spécifiques
- Réalisation de tests statistiques (ANOVA) pour évaluer la significativité des différences entre les groupes
- Génération d'un fichier Excel consolidé avec les résultats d'analyse

## Prérequis

- Python 3.8 ou supérieur
- pandas
- numpy
- scipy

## Installation

1. Assurez-vous que Python 3.8+ est installé sur votre système.
2. Installez les dépendances requises en exécutant :

```bash
pip install pandas numpy scipy
```

## Structure des Dossiers

- `fichier_xlsx/` : Répertoire contenant les fichiers .xls à traiter.
- `resultats_extraction.xlsx` : Fichier Excel généré contenant les résultats de l'analyse.
- `log_extraction.txt` : Fichier log détaillant les opérations réalisées par le script.

## Utilisation

1. Placez vos fichiers .xls dans le répertoire `fichier_xlsx/` dans le même repertoire que le script `Extraction_data_IMARIS.py`. Les fichiers doivent être nommés de manière cohérente pour faciliter le traitement automatisé. Par exemple, `souris1_groupe1_microglie1_vertionFichier.xls`, `souris1_groupe1_microglie2_vertionFichier.xls`, etc. Les fichiers doivent être au format .xls et placés dans le répertoire `fichier_xlsx
2. Exécutez le script en ligne de commande :

```bash
python Extraction_data_IMARIS.py
```

3. Consultez le fichier `resultats_extraction.xlsx` pour les résultats et `log_extraction.txt` pour les détails des opérations effectuées.

## Développement

Ce script a été développé avec les considérations suivantes :

- **Modularité** : Chaque fonction a un objectif clair et peut être réutilisée ou adaptée séparément.
- **Extensibilité** : Facilement adaptable pour inclure de nouvelles variables ou méthodes d'analyse.
- **Robustesse** : Gestion des erreurs pour éviter les interruptions lors du traitement de multiples fichiers.

## Auteur

FRAYSSE Yoann -

## Contact

yoann.fraysse@inserm.fr
