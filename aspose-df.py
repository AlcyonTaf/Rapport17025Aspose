import aspose.words as aw
import pandas as pd

doc = aw.Document(".\TestWord\SansMacro.docx")

table_essais = doc.get_child(aw.NodeType.TABLE, 4, True).as_table()

table_equipement = doc.get_child(aw.NodeType.TABLE, 5, True).as_table()

#alltables = doc.get_child_nodes(aw.NodeType.TABLE, True)
alltables = doc.select_nodes("//Body/Table")

print(alltables.count)

for a, tab in enumerate(alltables):
    print("table N° : " + str(a))
    for row in tab.as_table().rows:
        for cell in row.as_row().cells:
            print(cell.to_string(aw.SaveFormat.TEXT))


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
        #print(child.count)
        table_to_df = child[0]
        #print(table_to_df.rows.count)
        #print(table_to_df.rows[0].cells.count)
        df = [['' for i in range(table_to_df.rows[0].cells.count)] for j in range(table_to_df.rows.count)]

        for y, row in enumerate(table_to_df.rows):
            for j, cell in enumerate(row.as_row().cells):
                print(cell.to_string(aw.SaveFormat.TEXT))
                df[y][j] = cell.to_string(aw.SaveFormat.TEXT).replace("\r", "")

        data = pd.DataFrame(df)
        print(data)

    elif child.count > 1:
        print("Il y a plus d'une table ce n'est pas normale")
    else:
        print("pas de table")


extract_df_from_nested_table(table_essais, 1)
