# Import Aspose.Words for Python via .NET module
import aspose.words as aw



# Create and save a simple document
doc = aw.Document(".\TestWord\SansMacro.docx")

table = doc.get_child(aw.NodeType.TABLE, 4, True).as_table()
# allTables = doc.get_child_nodes(aw.NodeType.TABLE, True)
doc.get_child(aw.NodeType.TABLE, 5, True).remove()

x= 7

#testcell = dict()
row = table.rows[4]
cell = table.rows[4].cells[0]
cell.remove_all_children()
new_row = row.clone(True).as_row()

#cell_clone = cell.clone(False).as_cell()

cell.cell_format.width /= x
print(cell.cell_format.width)
for i in range(0, x):
    newcell = cell.clone(False).as_cell()
    cell.parent_node.append_child(newcell)
    #testcell[i] = cell.clone(False).as_cell()
    #testcell[i].cell_format.width /= 4


table.insert_after(new_row, row)

#newcell = cell.clone(False).as_cell()


#newcell.cell_format.width /= 4


# for i in range(0,4):
#     if i == 0:
#         cell.parent_node.insert_after(testcell[i], cell)
#     else:
#         cell.parent_node.insert_after(testcell[i], testcell[i-1])

table.allow_auto_fit = True

# row = table.rows[4]
# cell = row.first_cell
#
# row2 = aw.tables.Row(doc)
#
# cell2 = aw.tables.Cell(doc)
# cell2.append_child(aw.Paragraph(doc))
# cell2.first_paragraph.append_child(aw.Run(doc, "test"))
# cell3 = aw.tables.Cell(doc)
# cell3.append_child(aw.Paragraph(doc))
# cell3.first_paragraph.append_child(aw.Run(doc, "test"))
# cell4 = aw.tables.Cell(doc)
# cell4.append_child(aw.Paragraph(doc))
# cell4.first_paragraph.append_child(aw.Run(doc, "test"))
#
#
# row2.append_child(cell2)
# row2.append_child(cell3)
# row2.append_child(cell4)
#
# table.rows.insert(4, row2)




#cell.cell_format.width = newwidth
# cell2.cell_format.width = newwidth
# cell3.cell_format.width = newwidth
#
# row.insert_after(cell2, cell)
# row.insert_after(cell3, cell2)


#
# for row in table.rows:
#     for cell in row.as_row().cells:
#         print(cell.to_string(aw.SaveFormat.TEXT))


doc.save('test2.docx')