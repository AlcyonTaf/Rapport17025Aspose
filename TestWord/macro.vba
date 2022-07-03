Option Base 1

Sub test()
Dim x As Integer, y As Integer, I As Integer
Dim tbl As Table
Dim tTab() As Variant
Dim cCells As Cells
Dim cCell As Cell
x = 0
y = 0
For Each tbl In ActiveDocument.Tables
    With tbl
        If .Title = "test" Then
         Debug.Print "Nombre de colonne : " & .Columns.Count
         Debug.Print "Nombre de ligne : " & .Rows.Count
         ReDim tTab(.Rows.Count, .Columns.Count)
         Debug.Print "ubound :" & UBound(tTab, 1)

            'on remplie le tab avec les valeurs pour sa on parcour le tableau
            .Select
            Set cCells = Selection.Cells
            For Each cCell In cCells
                Debug.Print cCell.ColumnIndex
                Debug.Print cCell.RowIndex
                Debug.Print cCell.Range.Text
                tTab(cCell.RowIndex, cCell.ColumnIndex) = cCell.Range.Text


            Next cCell


        End If


    End With

 'Boucle sur tableau pour test
 For x = 1 To UBound(tTab, 1)
    For y = 1 To UBound(tTab, 2)
        Debug.Print "X =" & x & "-Y =" & y & "-Valeur: "; tTab(x, y)
    Next y
Next x


Next
End Sub

Sub testimbrique()
Dim x As Integer, y As Integer, I As Integer, Z As Integer, w As Integer, v As Integer
Dim tbl As Table, tblimbrique As Table
Dim addtbl As Table
Dim rRangetbl As Range
Dim iNbrRow As Integer, iNbrColumn As Integer
Dim iNbrColumnMax As Integer, iNbrRowParItem As Integer, iFirstRowValue As Integer
Dim iAddRow As Integer, iAddColum As Integer
Dim iNbrItem As Integer, iColVal As Integer, iColEntete As Integer
Dim iNbrtblImbrique As Integer
Dim iIndexTabRow As Integer, iIndexTabColum As Integer
Dim ICount As Integer
Dim iNumPageDebut As Integer, iNumPageFin As Integer, iNbrTableauInfo As Integer, iNumTableauListeEssai As Integer, iNumTableauListeEssaiColPage As Integer
Dim tTabRefInterneAutre() As Variant
Dim cCells As Cells
Dim cCell As Cell
Dim bTrouver As Boolean, bContreEssai As Boolean, bEquipement As Boolean

'#####CONFIGURATION#####
'Nombre de colonne de valeur de résultat max sur 1 ligne, en plus de la référence de l'item
iNbrColumnMax = 9
'Nombre de ligne d'entete dans le tableau
iFirstRowValue = 2
'Nbr de tableau de la 1er page, normalement ne bouge pas
iNbrTableauInfo = 4
'Numero du tableau avec la liste des essais
iNumTableauListeEssai = 3
'Numero de la colonne ou mettre les nums de page
iNumTableauListeEssaiColPage = 5

'####Traitement Ref Interne - Ref Autre ####
'Création d'un tableau vba avec la référence interne et la référence Autre
'Ceci va permettre de rajouter la ref autre dans le tableau de résultats
    'On sélectionne le tableau grace au signets
    'Selection.GoTo What:=wdGoToBookmark, Name:="TableauItem"
    ActiveDocument.Tables(2).Select

    With Selection
        Set cCells = Selection.Cells
                    'On dimensionne la table en fonction du tableau
                    iNbrRow = .Rows.Count
                    iNbrColumn = .Columns.Count
                    ReDim tTabRefInterneAutre(iNbrRow, iNbrColumn)
                    'Debug.Print "ubound :" & UBound(tTabRefInterneAutre, 1)
                    'On sauvegarde le tableau dans la table
                    For Each cCell In cCells
                        'Debug.Print cCell.ColumnIndex
                        'Debug.Print cCell.RowIndex
                        'Debug.Print cCell.Range.Text
                        Set rRangetbl = cCell.Range
                        rRangetbl.MoveEnd Unit:=wdCharacter, Count:=-1
                        tTabRefInterneAutre(cCell.RowIndex, cCell.ColumnIndex) = rRangetbl.Text
                    Next cCell
    End With



'#### Traitement des tableaux imbriqués ####

For Each tbl In ActiveDocument.Tables
    'declaration tableau
    Dim tTab() As Variant
    'Dim cCells As Cells
    'Dim cCell As Cell

    With tbl


    'Pour compter le nbr de tableau
    intt = intt + 1
    'Debug.Print "Tableau N°" & intt


    'Debug.Print "Nombre de colonne : " & .Columns.Count
    'Debug.Print "Nombre de ligne : " & .Rows.Count
    'Debug.Print "Nombre de ligne : " & .Tables.Count
    iNbrtblImbrique = .Tables.Count

    'Vérification si contre essai Statut : Essai refusé contre essai demandé
    If .Rows(1).Cells(1).Range.Text Like "*Essai refusé contre essai demandé*" Then
        'Debug.Print "Contre Essai"
        bContreEssai = True
        .Rows(1).Cells(1).Range.Font.ColorIndex = wdRed
    Else
        'Debug.Print "Pas contre essai"
        bContreEssai = False
    End If

    'test si tableau imbrique
    'Debug.Print .Tables.Count
    If .Tables.Count >= 1 And intt > 2 Then
        'On boucle par incrément d'une variable v car sinon probleme du au nombre de tableau imbrique qui diminue
        For v = 1 To iNbrtblImbrique
             'Debug.Print "Tableau imbriqué!"

            'Vérification tableau equipement
            If InStr(.Tables(1).Cell(1, 1).Range.Text, "Nom") Then
                bEquipement = True
            Else
                bEquipement = False
            End If

            'Vérification : Tableau de condition de l'essai ou de résultats
            If InStr(.Tables(1).Cell(1, 1).Range.Text, "Items de l'essai") Then

                'Vérification si tableau de condition vide : si,oui on supprime le tableau
                'TODO : Voir si on ne pourrait pas plutot vérifier le contenue de la cellule au dessus qui doit contenir "Conditions de(s) l'essai(s)"
                If InStr(.Tables(1).Cell(1, 2).Range.Text, "NA") Then
                    .Tables(1).Select
                    .Tables(1).Delete
                    'On ecrit N/A
                    Selection.TypeText Text:="N/A"
                    'On passe au tableau suivant
                    GoTo NextVIteration
                End If


                'Supression de la colonne avec la liste des éléments a tester
                'Debug.Print "Test colonne 2 : " & .Tables(1).Cell(1, 2).Range.Text
                If InStr(.Tables(1).Cell(1, 2).Range.Text, "Eléments à tester") Then .Tables(1).Columns(2).Delete
                '.Tables(1).Columns(2).Delete

                'Suppression des colonnes vides pour les tableaux de résultats
                'Les tableaux commence au numéro 3 et finisent a ActiveDocument.Tables.Count -1
                'Actuellement on vérifie juste que la 2eme ligne ne contient pas de valeur
                'Risque d'erreur si un essai n'est pas fait pour tout les items!
                 For I = .Tables(1).Columns.Count To 2 Step -1
                    If Len(.Tables(1).Cell(2, I).Range.Text) = 2 Then .Tables(1).Columns(I).Delete
                 Next

                End If

             'On sélectionne le tableau imbriqué
             .Tables(1).Select
             Set cCells = Selection.Cells
             'On dimensionne la table en fonction du tableau
             iNbrRow = .Tables(1).Rows.Count
             iNbrColumn = .Tables(1).Columns.Count
             ReDim tTab(iNbrRow, iNbrColumn)
             'Debug.Print "ubound :" & UBound(tTab, 1)
             'On sauvegarde le tableau dans la table
             For Each cCell In cCells
                 'Debug.Print cCell.ColumnIndex
                 'Debug.Print cCell.RowIndex
                 'Debug.Print cCell.Range.Text
                 Set rRangetbl = cCell.Range
                 rRangetbl.MoveEnd Unit:=wdCharacter, Count:=-1
                 tTab(cCell.RowIndex, cCell.ColumnIndex) = rRangetbl.Text
             Next cCell

             'On supprime la table imbriqué une fois sauvegarder dans le tableau
             .Tables(1).Delete
             'On recupére la position du curseur a ce moment
             'Debug.Print "Position colonne curseur :" & Selection.Information(wdEndOfRangeColumnNumber)
             iIndexTabColum = Selection.Information(wdEndOfRangeColumnNumber)
             iIndexTabRow = Selection.Information(wdEndOfRangeRowNumber)
             'Debug.Print "Position ligne curseur :" & Selection.Information(wdEndOfRangeRowNumber)

             'Calcul pour voir comment modifier le tableau
             'Debug.Print "Mod" & (iNbrColumn - 1) Mod iNbrColumnMax
             'Debug.Print "division entier" & (iNbrColumn - 1) \ iNbrColumnMax
             If ((iNbrColumn - 1) Mod iNbrColumnMax) = 0 Then
                 'Debug.Print "Nbr de ligne par item:" & (iNbrColumn - 1) \ iNbrColumnMax
                 iNbrRowParItem = ((iNbrColumn - 1) \ iNbrColumnMax) * 2
             Else
                 'Debug.Print "Nbr de ligne par item:" & ((iNbrColumn - 1) \ iNbrColumnMax) + 1
                 iNbrRowParItem = (((iNbrColumn - 1) \ iNbrColumnMax) + 1) * 2
             End If

             'Calcul du nombre d'item
             iNbrItem = iNbrRow - 1
             'Debug.Print "Nbr d'item:" & iNbrItem

             'On modifie le tableau :
             'Nbr de colum = iNbrColumnMax +1
             'Nbr de ligne  =
             iAddRow = ((iNbrRow - 1) * iNbrRowParItem)

             'On ajoute uniquement le nombre de colonne nécessaire
             If iNbrColumn < iNbrColumnMax Then
                iAddColum = iNbrColumn
             Else
                iAddColum = iNbrColumnMax + 1
             End If

             'apparment probleme pour split le tableau en plus de 13 lignes et 22 colonne
             'Il va falloir créer le tableau autrement dans ce cas la.
             'Autre probleme la limite semble changer parfois, je ne sais pas pourquoi
             'On va regler à 10 la limite
             If iAddRow >= 5 Then
                Selection.Cells.Split NumRows:=5, NumColumns:=iAddColum, MergeBeforeSplit:=False
                Selection.InsertRowsBelow (iAddRow - 5)
            Else
                Selection.Cells.Split NumRows:=iAddRow, NumColumns:=iAddColum, MergeBeforeSplit:=False
            End If


             'Prévoir une boucle sur chaque item pour fussionner les lignes correspondantes
             'Essayer de merge les cellules avec le nom des items avant l'ajout des valeurs
             'On centre également le texte
             For w = 1 To iNbrItem
                 If w = 1 Then
                 ActiveDocument.Range(Start:=.Cell(iIndexTabRow, 1).Range.Start, _
                     End:=.Cell((iIndexTabRow - 1 + (iNbrRowParItem * w)), 1).Range.End).Select
                     Selection.Cells.merge
                     Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
                     Selection.Cells.VerticalAlignment = wdCellAlignVerticalCenter
                     Selection.Font.Bold = True
                 Else
                 ActiveDocument.Range(Start:=.Cell(iIndexTabRow + (iNbrRowParItem * (w - 1)), 1).Range.Start, _
                     End:=.Cell((iIndexTabRow - 1 + (iNbrRowParItem * w)), 1).Range.End).Select
                     Selection.Cells.merge
                     Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
                     Selection.Cells.VerticalAlignment = wdCellAlignVerticalCenter
                     Selection.Font.Bold = True
                 End If
             Next w

             'Maintenant que le tableau est modifier on remplie avec les valeurs
             'On sélectionne la tableau imbriqué
             ActiveDocument.Range(Start:=.Cell(iIndexTabRow, 1).Range.Start, _
                 End:=.Cell(iAddRow + (iIndexTabRow - 1), iAddColum).Range.End).Select
             Set cCells = Selection.Cells

             'Voir pour faire la mise en forme des tableaux ici
             Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
             With cCells
                With .Borders(wdBorderLeft)
                    .LineStyle = wdLineStyleSingle
                    .LineWidth = wdLineWidth050pt
                    .Color = -603946753
                End With
                With .Borders(wdBorderRight)
                    .LineStyle = wdLineStyleSingle
                    .LineWidth = wdLineWidth050pt
                    .Color = -603946753
                End With
                With .Borders(wdBorderTop)
                    .LineStyle = wdLineStyleSingle
                    .LineWidth = wdLineWidth050pt
                    .Color = -603946753
                End With
                With .Borders(wdBorderBottom)
                    .LineStyle = wdLineStyleSingle
                    .LineWidth = wdLineWidth150pt
                    .Color = -603946753
                End With
                With .Borders(wdBorderHorizontal)
                    .LineStyle = wdLineStyleSingle
                    .LineWidth = wdLineWidth050pt
                    .Color = -603946753
                End With
                With .Borders(wdBorderVertical)
                    .LineStyle = wdLineStyleSingle
                    .LineWidth = wdLineWidth050pt
                    .Color = -603946753
                End With
                .VerticalAlignment = wdCellAlignVerticalCenter
                .Borders(wdBorderDiagonalDown).LineStyle = wdLineStyleNone
                .Borders(wdBorderDiagonalUp).LineStyle = wdLineStyleNone
                .Borders.Shadow = False
            End With

             'Mise en forme global des tableau Caractéristiques, conditions et résultats
             'Bordure bas de tableau
             'With Selection.Borders(wdBorderBottom)
             '   .LineStyle = Options.DefaultBorderLineStyle
             '   .LineWidth = Options.DefaultBorderLineWidth
             '   .Color = Options.DefaultBorderColor
             'End With


             Z = 2
             iColVal = 2
             iColEntete = 2
             For Each cCell In cCells
                 'Debug.Print "colonne:" & cCell.ColumnIndex
                 'Debug.Print "Ligne:" & (cCell.RowIndex - iIndexTabRow + 1)
                 'Copie de la ref de l'Item
                 If cCell.ColumnIndex = 1 Then
                     'Debug.Print "Cas column 1"
                     'Debug.Print "calcul Z :" & (cCell.RowIndex - iIndexTabRow + 1) Mod iNbrRowParItem
                     'On incrémente Z de 1 a chaque fois qu'on a passer toutes les lignes d'un item
                     'Attention possition des lignes : on enleve les premiers
                     'Debug.Print " Z :" & Z
                     'Ici il faut concatener les 2 reférence item (lims et client)

                     'comparaison de ref item pour ensuite sortir Ref autre
                     'En y = 1 ref lims et y = 2 ref aute
                     bTrouver = False
                     For x = 2 To UBound(tTabRefInterneAutre, 1)
                             'Debug.Print "Comparaison : " & tTab(Z, 1) & " et " & tTabRefInterneAutre(x, 1)
                         If tTabRefInterneAutre(x, 1) = tTab(Z, 1) Then
                             'Debug.Print "OK sa match"
                             'Debug.Print tTabRefInterneAutre(x, 2)
                             cCell.Range.Text = tTab(Z, 1) & " - " & tTabRefInterneAutre(x, 2)
                             bTrouver = True
                             Exit For
                         End If
                     Next x
                     'Si on a rien trouver on met la valeur sans conctatener
                     If bTrouver = False Then cCell.Range.Text = tTab(Z, 1)

                     'Mise en forme fond 1/2
                     If IsPair(Z) Then cCell.Range.Shading.BackgroundPatternColor = -603923969

                     'Mise en forme si contre essai
                     If bContreEssai And Not bEquipement Then cCell.Range.Font.StrikeThrough = wdToggle

                     Z = Z + 1
                     'remise a zero compteur colonne car on passe a l'item suivant
                     If iColVal = UBound(tTab, 2) + 1 Then iColVal = 2
                     If iColEntete = UBound(tTab, 2) + 1 Then iColEntete = 2
                     'If (cCell.RowIndex - iFirstRowValue) Mod iNbrRowParItem = 0 Then Z = Z + 1
                 Else
                     'Maintenant il faut s'occuper des valeurs des résultats
                     'Si ligne impair = entete
                     'Si ligne pair = valeur
                     'Debug.Print "Cas colum <>1"
                     If IsPair(cCell.RowIndex - iIndexTabRow + 1) Then
                         If iColVal <= UBound(tTab, 2) Then
                             'Debug.Print "Cas valeur "
                             'Debug.Print "Z-1:" & Z - 1
                             cCell.Range.Text = tTab(Z - 1, iColVal)
                             iColVal = iColVal + 1
                             'If iColVal = UBound(tTab, 2) Then iColVal = 2
                             'Mise en forme fond 1/2
                                If IsPair(Z) = False Then cCell.Range.Shading.BackgroundPatternColor = -603923969
                             'Mise en forme si contre essai
                                If bContreEssai And Not bEquipement Then cCell.Range.Font.StrikeThrough = wdToggle
                         End If
                     Else
                         If iColEntete <= UBound(tTab, 2) Then
                             'Debug.Print "Cas entete "
                             cCell.Range.Text = tTab(1, iColEntete)
                             cCell.Range.Font.Bold = True
                             iColEntete = iColEntete + 1
                             'If iColEntete = UBound(tTab, 2) Then iColEntete = 2
                             'Mise en forme fond 1/2
                                If IsPair(Z) = False Then cCell.Range.Shading.BackgroundPatternColor = -603923969
                             'Mise en forme si contre essai
                                If bContreEssai And Not bEquipement Then cCell.Range.Font.StrikeThrough = wdToggle
                         End If
                     End If
                 End If





             Next cCell
NextVIteration:
         Next v

         'Essai split derniere cellule pour mise en forme signature
         'TODO : NumColumns doit etre égale au nombre de signature!
         .Cell(.Rows.Count, 1).Select
            ICount = (Len(Selection.Text) - Len(Replace(Selection.Text, ";", ""))) / Len(";")
            'Une fois compter on supprimer le ;
            For x = 1 To ICount
            Selection.Find.ClearFormatting
            Selection.Find.Replacement.ClearFormatting
                With Selection.Find
                    .Text = ";"
                    .Replacement.Text = ""
                    .Forward = True
                    .Wrap = wdFindStop
                    .Format = False
                    .MatchCase = False
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
            Selection.Find.Execute Replace:=wdReplaceOne
            Next x

         'Debug.Print "Nombre de signature : " & ICount


         Selection.Cells.Split NumRows:=1, NumColumns:=ICount, MergeBeforeSplit:=False

         'Calcul des numéro de page des essais a inscrire dans le tableau avec la liste des essais.
         'On sélectionne la 1er cellule pour connaitre le Num de la page de debut
         .Cell(1, 1).Select
         iNumPageDebut = Selection.Information(wdActiveEndAdjustedPageNumber)
         'Debug.Print "Page debut : " & iNumPageDebut
         'On sélectionne le tableau en entier pour connaitre le Num de la page de fin
         .Select
         iNumPageFin = Selection.Information(wdActiveEndAdjustedPageNumber)
         'Debug.Print "Page fin : " & iNumPageFin
         'On regarde si le tableau est sur plusieur page pour adapter le texte
         If iNumPageFin = iNumPageDebut Then
            'Cas ou 1 page :
            ActiveDocument.Tables(iNumTableauListeEssai).Cell(1 + (intt - iNbrTableauInfo), iNumTableauListeEssaiColPage).Range.Text = iNumPageDebut
         Else
            'cas ou plusieur page
            ActiveDocument.Tables(iNumTableauListeEssai).Cell(1 + (intt - iNbrTableauInfo), iNumTableauListeEssaiColPage).Range.Text = iNumPageDebut & " à " & iNumPageFin
         End If
         'On met en forme dans le cas d'un contre essai
         If bContreEssai And Not bEquipement Then
         'On barre et on met en rouge la 1er col qui est le nom de l'essai
         ActiveDocument.Tables(iNumTableauListeEssai).Cell(1 + (intt - iNbrTableauInfo), 1).Range.Font.StrikeThrough = wdToggle
         ActiveDocument.Tables(iNumTableauListeEssai).Cell(1 + (intt - iNbrTableauInfo), 1).Range.Font.ColorIndex = wdRed
         End If

     Else
     Debug.Print "probleme nbr tableau imbriqué"
     .Select
     End If

    End With


Erase tTab

Next
End Sub

Function IsPair(nombre As Integer) As String
IsPair = IIf((nombre And 1), False, True)
End Function


Sub testmiseenforme()
Dim x As Integer, y As Integer, I As Integer
Dim tbl As Table
Dim tTab() As Variant
Dim cCells As Cells
Dim cCell As Cell
x = 0
y = 0
Dim iNbrRowParItem As Integer

iNbrRowParItem = 4
Set tbl = ActiveDocument.Tables(3)

With tbl
    Debug.Print "Nombre de colonne : " & .Columns.Count
    Debug.Print "Nombre de ligne : " & .Rows.Count
    Debug.Print "Nombr de tableau imbriqué : " & .Tables.Count
    Debug.Print .Tables(1).Columns.Count

    Debug.Print Selection.Information(wdEndOfRangeColumnNumber)
    Debug.Print Selection.Information(wdEndOfRangeRowNumber)

    '.Select
    'Debug.Print Selection.Information(wdActiveEndPageNumber)


End With

End Sub




Sub merge()
Dim x As Integer, y As Integer, I As Integer
Dim tbl As Table
Dim tTab() As Variant
Dim cCells As Cells
Dim cCell As Cell
x = 0
y = 0
For Each tbl In ActiveDocument.Tables
    With tbl
        If .Title = "merge" Then
         Debug.Print "Nombre de colonne : " & .Columns.Count
         Debug.Print "Nombre de ligne : " & .Rows.Count
         ReDim tTab(.Rows.Count, .Columns.Count)
         Debug.Print "ubound :" & UBound(tTab, 1)

            'on remplie le tab avec les valeurs pour sa on parcour le tableau
            .Select
            Set cCells = Selection.Cells
            For Each cCell In cCells
                Debug.Print "colonne:" & cCell.ColumnIndex
                Debug.Print "ligne:" & cCell.RowIndex
                Debug.Print cCell.Range.Text
            Next cCell


        End If


    End With


Next
End Sub


Sub creationtab()
Set addtbl = ActiveDocument.Tables.Add(Range:=Selection.Range, NumRows:=3, NumColumns:=3)
                'Debug.Print Selection.Information(wdStartOfRangeRowNumber)
                With addtbl
                    .Borders.Enable = True
                    .Borders(wdBorderBottom).LineWidth = wdLineWidth050pt
                    .Borders(wdBorderLeft).LineWidth = wdLineWidth050pt
                    .Borders(wdBorderRight).LineWidth = wdLineWidth050pt
                    .Borders(wdBorderTop).LineWidth = wdLineWidth050pt
                End With

End Sub


Sub testtabitem()
Dim tTabRefInterneAutre() As Variant
Dim rRangetbl As Range 'existe déja

'#A rajouter avec les autres
Dim cCells As Cells
Dim cCell As Cell
'####Traitement Ref Interne - Ref Autre ####
'Création d'un tableau vba avec la référence interne et la référence Autre
'Ceci va permettre de rajouter la ref autre dans le tableau de résultats
    'On sélectionne le tableau grace au signets
    Selection.GoTo What:=wdGoToBookmark, Name:="TableauItem"

    With Selection
        Set cCells = Selection.Cells
                    'On dimensionne la table en fonction du tableau
                    iNbrRow = .Rows.Count
                    iNbrColumn = .Columns.Count
                    ReDim tTabRefInterneAutre(iNbrRow, iNbrColumn)
                    Debug.Print "ubound :" & UBound(tTabRefInterneAutre, 1)
                    'On sauvegarde le tableau dans la table
                    For Each cCell In cCells
                        'Debug.Print cCell.ColumnIndex
                        'Debug.Print cCell.RowIndex
                        'Debug.Print cCell.Range.Text
                        Set rRangetbl = cCell.Range
                        rRangetbl.MoveEnd Unit:=wdCharacter, Count:=-1
                        tTabRefInterneAutre(cCell.RowIndex, cCell.ColumnIndex) = rRangetbl.Text
                    Next cCell
    End With

    'comparaison de ref item pour ensuite sortir Ref autre
    'En y = 1 ref lims et y = 2 ref aute
    For x = 2 To UBound(tTabRefInterneAutre, 1)
        If tTabRefInterneAutre(x, 1) = "I19000038" Then
            Debug.Print "OK sa match"
            Debug.Print tTabRefInterneAutre(x, 2)
        End If
    Next x

 For x = 2 To UBound(tTabRefInterneAutre, 1)
    For y = 1 To 2
        Debug.Print "X =" & x & "-Y =" & y & "-Valeur: "; tTabRefInterneAutre(x, y)
    Next y
Next x


End Sub





