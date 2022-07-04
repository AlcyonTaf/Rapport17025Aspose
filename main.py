# -*- coding: utf-8 -*-
import pandas as pd
import aspose.words as aw

# Soit on iter sur tous le documents dans l'ordre, soit on utilise les index.
# Index tableau :
# 0 = tableau informations demandes = Pas de modif
# 1 = Tableau informations item = Pas de modif
# 2 = Tableau informations essais = Mettre status des essais et ou num page de l'essai
# 3 = Tableau de révision = on ne fait rien
# 4 = 1er essai : modifier format des tableau imbriqué =Voir details ci-dessous
# 4 + n-1 = Tableau des autres essais
# Dernier tableau : Signature = on ne fait rien

# Détails tableau imbriqué dans essais :
# ligne 4 : Tableau maintenance equipement : one ne fait rien
# ligne 7 : Tableau des caractéristiques demandés : Mise ne forme a faire
# Ligne 9 : Tableau conditions de l'essais : Mise en forme a faire
# Ligne 11 : Tableau des résultats : Mise en forme a faire
# On va faire la même mise en forme pour chacun de ces tableaux

'''
#############
# Variables #
#############
'''
nbr_tableau_fixe = 5  # PAs util, Nbr de tableau qui ne bouge pas, c'est-à-dire qui ne sont pas des tableaux d'essais
list_row_nested_tab_essais = [4, 7, 9, 11]  # Emplacement des tableaux imbriqué dans les tableaux essais

'''
 Fonction
'''


def extract_df_from_nested_table(table_to_check_in, row_index):
    child = table_to_check_in.rows[row_index].cells[0].tables
    data = []
    if child.count == 1:
        print("Table nested!")
        # Maintenant, il faut récupérer cette table et faire une boucle sur chaque cellule
        # print(child.count)
        table_to_df = child[0]
        # print(table_to_df.rows.count)
        # print(table_to_df.rows[0].cells.count)
        df = [['' for i in range(table_to_df.rows[0].cells.count)] for j in range(table_to_df.rows.count)]

        for y, row in enumerate(table_to_df.rows):
            for j, cell in enumerate(row.as_row().cells):
                # print(cell.to_string(aw.SaveFormat.TEXT))
                df[y][j] = cell.to_string(aw.SaveFormat.TEXT).replace("\r", "")

        data = pd.DataFrame(df)
        # print(data)
        return data, True

    elif child.count > 1:
        print("Il y a plus d'une table ce n'est pas normale")
        raise ValueError("Il y a plus d'une table ce n'est pas normale")
    else:
        print("pas de table")
        data = pd.DataFrame()
        return data, False


def get_items():
    """
    Function pour récupérer le tableau d'information sur les items
    Va permettre de mettre la ref client avec la ref lims pour les essais
    Le tableau d'items est toujours en position deux
    :return: df avec le contenu du tableau
    """
    table_items = doc.get_child(aw.NodeType.TABLE, 1, True).as_table()
    df_items = [['' for i in range(table_items.rows[0].cells.count)] for j in range(table_items.rows.count)]
    for y, row in enumerate(table_items.rows):
        for j, cell in enumerate(row.as_row().cells):
            # print(cell.to_string(aw.SaveFormat.TEXT))
            df_items[y][j] = cell.to_string(aw.SaveFormat.TEXT).replace("\r", "")

    df_items = pd.DataFrame(df_items)
    return df_items


def apply_formatting_to_table(table_to_check_in, row_index):
    """
    il va falloir supprimer  la table imbriqué puis rajouter les lignes et cellule en fonction du dataframe

    :param table_to_check_in:
    :param row_index:
    :return:
    """
    print('test')


if __name__ == '__main__':
    doc = aw.Document(".\TestWord\SansMacro.docx")

    # toutes les tables directement dans le documents, pas les imbriqué ni dans header
    all_tables = doc.select_nodes("//Body/Table")

    # On récupère les positions des tables d'essais:
    # print("nbr de tableau d'essai : " + str(len(tb) - nbr_tableau_fixe))
    list_index_tab = range(4, all_tables.count - 1)
    # on boucle sur chaque tableau d'essai
    for i in list_index_tab:
        table_essais = doc.get_child(aw.NodeType.TABLE, i, True).as_table()

        # On iter sur chaque table imbriquée du tableau d'essai, pour cela on utilise l'index de la ligne
        for y in list_row_nested_tab_essais:
            try:
                result = extract_df_from_nested_table(table_essais, y)

            except ValueError as err:
                print("Erreur lors de la lecture des tableaux imbriqué dans les tables d'essai : {}0".format(err))
            else:
                # On vérifie si df n'est pas vide
                if result[1]:
                    print("ok donnée on continue")
                    print(result[0])
                    # Ici on va faire la mise en forme
                else:
                    print("rien")
