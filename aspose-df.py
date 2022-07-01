import aspose.words as aw

doc = aw.Document(".\TestWord\SansMacro.docx")

table_essais = doc.get_child(aw.NodeType.TABLE, 4, True).as_table()

table_equipement = doc.get_child(aw.NodeType.TABLE, 5, True).as_table()


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
    if child.count > 0:
        print("Table nested!")
    else:
        print("pas de table")


extract_df_from_nested_table(table_essais,8)