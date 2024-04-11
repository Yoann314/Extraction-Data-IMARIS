#!/usr/bin/python
# -*- coding: utf-8 -*-
 
import os
import re
import pandas as pd
import numpy as np
from scipy.stats import ttest_ind
from scipy.stats import f_oneway



"""
Vous trouverez ici toutes les variable extraite par le script. 
Si les nom de variable ne sont pas les mêmes que dans les fichiers Excel, vous pouvez les changer ici ou en rajouter si necessaire.
Si vous souaitez en rajouter, il faut aussi rajouter la variable dans la fonction 'verification_variable' pour que le script puisse la prendre en compte.
"""



VARIABLE_1 = "Filament Area (sum)"
VARIABLE_2 = "Filament Full Branch Depth"
VARIABLE_3 = "Filament Full Branch Level"
VARIABLE_4 = "Filament Length (sum)"
VARIABLE_5 = "Filament Volume (sum)"
VARIABLE_6 = "Segment Length"
VARIABLE_7 = "Filament Length (sum)"
VARIABLE_8 = "Area"
VARIABLE_9 = "Intensity Mean"
VARIABLE_10 = "Volume"

# Chemin vers le répertoire contenant les fichiers .xls
repertoire = 'fichier_xlsx'

# Liste et trie les fichiers .xls par ordre alphabétique
fichiers_xls = sorted([f for f in os.listdir(repertoire) if f.endswith('.xls')])

# Chemin vers le fichier d'inventaire
log = 'log_extraction.txt'

# Chemin du fichier Excel de sortie
chemin_fichier_excel = 'resultats_extraction.xlsx'



"""
Ces dictionnaires ont la structure suivante :
{
    'Variable1': {
        'Souris1_microglia1': {'mean': 0.123, 'sum': 0.456},
        'Souris2_microglia2': {'mean': 0.789, 'sum': 0.012},
        ...
    },
    'Variable2': {
        'Souris1_microglia1': {'mean': 0.345, 'sum': 0.678},
        'Souris2_microglia2': {'mean': 0.901, 'sum': 0.234},
        ...
    },
    ...
}
"""

variables_dict_one = {}
variables_dict_tree = {}



def extraire_informations(nom_fichier):
    """
    Extrait les informations spécifiques du nom du fichier.

    Args:
        nom_fichier (str): Le chemin du fichier à analyser.

    Returns:
        tuple: Un tuple contenant les parties extraites du nom du fichier.
    """

    pattern = re.compile(r"_(\d{3})_(NI_AD|NI|Tg_AD|Tg|AD).*?microglia(\d+(?:_\d+)?)(?:\.(\d+))?\.xls")
    match = pattern.search(nom_fichier)
    if match:
        return match.groups()
    return None, None, None, None

def significance_level(p_val):
    """
    Détermine le niveau de significativité en fonction de la valeur p

    Args:
        p_val (float): La valeur p à évaluer.

    Returns:
        str: Le niveau de significativité correspondant.
    """

    if p_val < 0.001:
        return '***'
    elif p_val < 0.01:
        return '**'
    elif p_val < 0.05:
        return '*'
    else:
        return 'ns'
    
def determine_group(x):
    """
    Détermine le groupe en fonction de la sous-chaîne trouvée dans le nom du fichier.

    Args:
        x (str): Le nom du fichier à analyser.
    
    Returns:
        str: Le groupe correspondant à la sous-chaîne trouvée.
    """

    # Dictionnaire de correspondance sous-chaîne vers groupe
    group_map = {
        'Tg_AD': 'Tg_AD',
        'Tg': 'Tg',
        'NI_AD': 'NI_AD',
        'NI': 'NI'
    }
    # Parcourt le dictionnaire pour trouver un match
    for key, group in group_map.items():
        if key in x:
            return group
    # Si aucun match n'est trouvé, retourne une valeur par défaut ou lève une exception
    return "Groupe Inconnu"  # Ou utilisez raise ValueError("Groupe inconnu")

        
def extraire_numeros(chaine):
    """
    Extrait le premier nombre trouvé dans une chaîne de caractères.

    Args:
        chaine (str): La chaîne de caractères à analyser.

    Returns:
        str: Le premier nombre trouvé ou la chaîne originale si aucun nombre n'est trouvé.
    """

    # Utilisation de re.search pour trouver la première occurrence de tout nombre
    match = re.search(r'\d+', chaine)
    if match:
        return match.group()  # Retourne la première séquence de chiffres trouvée
    return chaine  # Retourne la chaîne originale si aucun chiffre n'est trouvé

        
def verification_variable(variable):
    """
    Vérifie si la variable est une varaible d'interet. Les variables d'intérêt sont les suivantes :

    Args:
        variable (str): La variable à vérifier.
    
    Returns:
        bool: True si la variable est d'intérêt, False sinon.
    
    
    version == "1"
        Means : 
            Filament Area (sum)
            Filament Full Branch Depth
            Filament Full Branch Level
            Filament Length (sum)
            Filament Volume (sum)
            Segment Length
        Sum : 
            Filament Length (sum)
    version == "3"
        Means :
            Area
            Intensity Mean
            Volume
    """
    
    if variable == VARIABLE_1 or variable == VARIABLE_2 or variable == VARIABLE_3 or variable == VARIABLE_4 or variable == VARIABLE_5 or variable == VARIABLE_6 or variable == VARIABLE_7 or variable == VARIABLE_8 or variable == VARIABLE_9 or variable == VARIABLE_10:
        return True
    return False

def traiter_fichier_version(repertoire, nom_fichier, version):
    """
    Traite un fichier IMARIS en fonction de sa version.

    Args:
        repertoire (str): Le chemin du répertoire contenant le fichier.
        nom_fichier (str): Le nom du fichier à traiter.
        version (str): La version du fichier IMARIS.

    Returns:
        None
    """

    souris, groupe, microglia, _ = extraire_informations(nom_fichier)
    try:
        fichier.write(f"Lecture du fichier : {nom_fichier}, Groupe : {groupe}, Souris : {souris}, Microglie : microglia{microglia}, Version : {version}\n")
        df = pd.read_excel(os.path.join(repertoire, nom_fichier))
        identifiant_unique = groupe + '_' + souris + '_' + microglia

        if df.columns[0] == "Average":
            df.columns = df.iloc[0]
            df = df.drop(0).reset_index(drop=True)
            
            # Déterminer les indices une seule fois ici
            indice_mean = df.columns.get_loc("Mean") if "Mean" in df.columns else None
            indice_sum = df.columns.get_loc("Sum") if "Sum" in df.columns and version == "1" else None

            # Filtrez le DataFrame pour ne conserver que les lignes correspondant aux variables d'intérêt
            df_interet = df[df.iloc[:, 0].apply(verification_variable)]

            # Pour chaque variable d'intérêt, mise à jour des dictionnaires
            for _, row in df_interet.iterrows():
                variable = row.iloc[0]
                if version == "1":
                    variables_dict_one.setdefault(variable, {})[identifiant_unique] = {'mean': row.iloc[indice_mean], 'sum': row.iloc[indice_sum]}
                elif version == "3":
                    variables_dict_tree.setdefault(variable, {})[identifiant_unique] = {'mean': row.iloc[indice_mean]}

        else:
            fichier.write(f"\nLa première ligne ne contient pas 'Average', le fichier na pas ete correctement enregistrer dans IMARIS. Nom du fichier : {nom_fichier}, Groupe : {groupe}, Souris : {souris}, Microglie : microglia{microglia}, Version : {version}\n\n")
            print(f"\nLa première ligne ne contient pas 'Average', le fichier na pas ete correctement enregistrer dans IMARIS. Nom du fichier : {nom_fichier}, Groupe : {groupe}, Souris : {souris}, Microglie : microglia{microglia}, Version : {version}\n\n")

    except Exception as e:
        fichier.write(f"\nUne erreur est survenue lors de la lecture du fichier : {nom_fichier}, Groupe : {groupe}, Souris : {souris}, Microglie : microglia{microglia}, Version : {version}\n\n")
        print(f"\nUne erreur est survenue lors de la lecture du fichier : {nom_fichier}, Groupe : {groupe}, Souris : {souris}, Microglie : microglia{microglia}, Version : {version}\n\n")



######### Exploration des fichiers IMARIS #########



# Ouvre le fichier d'inventaire en mode écriture
with open(log, 'w') as fichier:
    fichier.write("Exploration des fichiers IMARIS\n\n")
    fichier.write("Chemin vers le répertoire contenant les fichiers à traiter : {}\n\n".format(repertoire))
    fichier.write("Chemin du fichier Excel de sortie : {}\n\n".format(chemin_fichier_excel))

    for nom_fichier in fichiers_xls:
        _, _, _, version = extraire_informations(nom_fichier)
        if version in ["1", "3"]:  # Assurez-vous que 'version' est une chaîne
            traiter_fichier_version(repertoire, nom_fichier, version)
        else:
            fichier.write(f"Version non reconnue pour le fichier {nom_fichier}\n")



######### Traitement des fichiers Excel #########



# Préparer deux listes pour stocker les données de mean et sum séparément
donnees_mean_one = []
donnees_sum_one = []
donnees_mean_tree = []

if not variables_dict_one:
    print("Aucune data n'as pu être extraite des fichiers de la version 1")
else:
    for variable, fichiers in variables_dict_one.items():
        for fichier, stats in fichiers.items():
            donnees_mean_one.append({'Variable': variable, 'Fichier': fichier, 'Mean': stats['mean']})
            donnees_sum_one.append({'Variable': variable, 'Fichier': fichier, 'Sum': stats['sum']})

if not variables_dict_tree:
    print("Aucune data n'as pu être extraite des fichiers de la version 3")
else:
    for variable, fichiers in variables_dict_tree.items():
        for fichier, stats in fichiers.items():
            donnees_mean_tree.append({'Variable': variable, 'Fichier': fichier, 'Mean': stats['mean']})


# Convertir les listes en DataFrames
df_mean_one = pd.DataFrame(donnees_mean_one)
df_sum_one = pd.DataFrame(donnees_sum_one)
df_mean_tree = pd.DataFrame(donnees_mean_tree)

# Restructurer les DataFrames pour avoir une ligne par variable et une colonne par fichier
pivot_mean_one = df_mean_one.pivot(index='Variable', columns='Fichier', values='Mean')
pivot_sum_one = df_sum_one.pivot(index='Variable', columns='Fichier', values='Sum')
pivot_mean_tree = df_mean_tree.pivot(index='Variable', columns='Fichier', values='Mean')

# Créer un writer pandas Excel
with pd.ExcelWriter(chemin_fichier_excel, engine='openpyxl') as writer:
    pivot_mean_one.T.to_excel(writer, sheet_name='Means')
    pivot_sum_one.T.to_excel(writer, sheet_name='Sums')
    pivot_mean_tree.T.to_excel(writer, sheet_name='Means Hull')


######### Calcul des moyennes et tests t #########
    

# Lire le fichier Excel
df = pd.read_excel(chemin_fichier_excel, sheet_name=None)  # Lit toutes les feuilles

# Pour chaque feuille dans le fichier Excel
for sheet_name, sheet_df in df.items():
    # Identifier les groupes
    ###sheet_df['Groupe'] = sheet_df['Fichier'].apply(lambda x: 'NI_AD' if 'NI_AD' in x else 'NI')
    sheet_df['Groupe'] = sheet_df['Fichier'].apply(determine_group)

    # Sélectionner uniquement les colonnes numériques pour le calcul des moyennes, exclure 'Fichier' et 'Groupe'
    columns_numeriques = sheet_df.select_dtypes(include=[np.number]).columns

    # Calculer les moyennes pour chaque groupe en se limitant aux colonnes numériques
    mean_ni = sheet_df.loc[sheet_df['Groupe'] == 'NI', columns_numeriques].mean()
    mean_ni_ad = sheet_df.loc[sheet_df['Groupe'] == 'NI_AD', columns_numeriques].mean()
    mean_tg = sheet_df.loc[sheet_df['Groupe'] == 'Tg', columns_numeriques].mean()
    mean_tg_ad = sheet_df.loc[sheet_df['Groupe'] == 'Tg_AD', columns_numeriques].mean()

    # Préparer les données pour les tests t et les niveaux de significativité pour chaque colonne numérique
    anova_results = {}
    for column in columns_numeriques:
        # Extraction des données pour chaque groupe
        group_ni = sheet_df.loc[sheet_df['Groupe'] == 'NI', column].dropna()
        group_ni_ad = sheet_df.loc[sheet_df['Groupe'] == 'NI_AD', column].dropna()
        group_tg = sheet_df.loc[sheet_df['Groupe'] == 'Tg', column].dropna()
        group_tg_ad = sheet_df.loc[sheet_df['Groupe'] == 'Tg_AD', column].dropna()
        
        # Application de l'ANOVA
        f_stat, p_val = f_oneway(group_ni, group_ni_ad, group_tg, group_tg_ad)
        anova_results[column] = {'F_stat': f_stat, 'p_val': p_val, 'significance': significance_level(p_val)}

    # Conversion des résultats ANOVA en DataFrame pour facilité d'ajout
    anova_df = pd.DataFrame(anova_results).T
    
    # Ajouter les lignes de moyennes, tests t, et significativité au DataFrame original
    additional_rows = pd.DataFrame([mean_ni, mean_ni_ad, mean_tg, mean_tg_ad, anova_df['p_val'], anova_df['significance']],
                                   index=['Mean_NI', 'Mean_NI_AD', 'Mean_TG', 'Mean_TG_AD', 'F_stat', 'Significance'])
    sheet_df = pd.concat([sheet_df, additional_rows])
    
    # Sauvegarder le DataFrame modifié dans le dictionnaire df sous la même clé (nom de feuille)
    df[sheet_name] = sheet_df
    
    # Ajouter les noms des lignes pour les moyennes, tests t, et significativité
    sheet_df.iloc[len(sheet_df) - 6, 0] = "Mean NI"
    sheet_df.iloc[len(sheet_df) - 5, 0] = "Mean_NI_AD"
    sheet_df.iloc[len(sheet_df) - 4, 0] = "Mean_TG"
    sheet_df.iloc[len(sheet_df) - 3, 0] = "Mean_TG_AD"
    sheet_df.iloc[len(sheet_df) - 2, 0] = "Anova p-value"
    sheet_df.iloc[len(sheet_df) - 1, 0] = "Significance"

    sheet_df['Fichier'] = sheet_df['Fichier'].apply(extraire_numeros)

with pd.ExcelWriter(chemin_fichier_excel, engine='openpyxl') as writer:
    for sheet_name, sheet_df in df.items():
        sheet_df.to_excel(writer, sheet_name=sheet_name, index=False)

print("############################################\n\nLe script a terminé avec succès ! \nRetouvez le détail des opérations dans le fichier log_extraction.txt et les résultats dans le fichier resultats_extraction.xlsx\n\n############################################")