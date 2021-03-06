import aspose.words as aw
import pandas as pd

doc = aw.Document(".\TestWord\SansMacro.docx")

table_essais = doc.get_child(aw.NodeType.TABLE, 4, True).as_table()

table_equipement = doc.get_child(aw.NodeType.TABLE, 5, True).as_table()


# alltables = doc.get_child_nodes(aw.NodeType.TABLE, True)
# alltables = doc.select_nodes("//Body/Table")
#
# print(alltables.count)
#
# for a, tab in enumerate(alltables):
#     print("table N° : " + str(a))
#     for row in tab.as_table().rows:
#         for cell in row.as_row().cells:
#             print(cell.to_string(aw.SaveFormat.TEXT))


# On essaie de récupérer uniquement les tables inbriqué dans la table essais
# test = table_essais.get_child_nodes(aw.NodeType.TABLE, True)
# print(test.count)
# for table in test:
#     print(table.as_table())


# table_equipement.parent_node

# Boucle sur les cells pour voir si elle contient des tables
# for row in table_essais.rows:
#     for cell in row.as_row().cells:
#         print(cell.to_string(aw.SaveFormat.TEXT))
#         #print(cell.as_cell().tables)
#         childtables = cell.as_cell().tables
#         if childtables.count > 0:
#             print("Table nested!")

# TOdo : Faire fonction avec en entré numéro de ligne, qui va vérifier si un tableau existe dans la cellule de cette
#  ligne. Si oui on extrait les valeurs dans un df
def extract_df_from_nested_table(table_to_check_in, row_index):
    child = table_to_check_in.rows[row_index].cells[0].tables
    if child.count == 1:
        print("Table nested!")
        # Maintenant, il faut récupérer cette table et faire une boucle sur chaque cellule
        # print(child.count)
        table_to_df = child[0]
        # print(table_to_df.rows.count)
        # print(table_to_df.rows[0].cells.count)
        df_nested_table = [['' for i in range(table_to_df.rows[0].cells.count)] for j in range(table_to_df.rows.count)]

        for y, row in enumerate(table_to_df.rows):
            for j, cell in enumerate(row.as_row().cells):
                # print(cell.to_string(aw.SaveFormat.TEXT))
                df_nested_table[y][j] = cell.to_string(aw.SaveFormat.TEXT).replace("\r", "")

        data = pd.DataFrame(df_nested_table)
        print(data)
        return data

    elif child.count > 1:
        print("Il y a plus d'une table ce n'est pas normale")
    else:
        print("pas de table")


# Function pour récupérer le tableau d'information sur les items
# Va permettre de mettre la ref client avec la ref lims pour les essais
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
            #print(cell.to_string(aw.SaveFormat.TEXT))
            df_items[y][j] = cell.to_string(aw.SaveFormat.TEXT).replace("\r", "")

    df_items = pd.DataFrame(df_items)
    return df_items



df = extract_df_from_nested_table(table_essais, 11)

# Tous les tableaux ne vont pas etre traiter de la même maniere
# cas des tableaux de chimie : On repere "Eléments à tester" en position 0,1
# On fera une function pour manipuler et mettre en forme le dataframe avant de modifier les tableaux word
print(df.iat[0, 1])
if df.iat[0, 1] == 'Eléments à tester':
    # On supprime la colonne Elements a tester
    df = df.drop(columns=1)
    # print(df.iloc[1:])
    # print(df.iloc[1:] =='')
    # print(df.loc[:,(df.iloc[1:] != '').any()])
    # On supprime les colonnes ou pour toutes les lignes, il n'y a aucune valeur
    df = df.loc[:, (df.iloc[1:] != '').any()].reset_index(drop=True, )
    print(df)
else:
    print("pas chimie")



