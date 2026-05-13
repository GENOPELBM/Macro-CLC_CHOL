Attribute VB_Name = "Module3"
Sub PANEL_HFE()

    Dim genes As Variant
    
    ' Liste des gènes à chercher dans la feuille "Mergevariant"
    genes = Array("SLC40A1", "BMP6", "LOC108783645, HFE", "HFE", "FTL", "HFE2", "HAMP", "TFR2")

    Call FiltrerFeuilles(genes, additionalWords, "PANEL_HFE")
    
End Sub

Sub PANEL_CHOL()

    Dim genes As Variant
    
    ' Liste des gènes à chercher dans la feuille "Mergevariant"
    genes = Array("LDLRAP1", "PCSK9", "APOB", "LDLR", "APOE")

    Call FiltrerFeuilles(genes, additionalWords, "PANEL_CHOL")
    
End Sub

Sub PANEL_SCU()

    Dim genes As Variant
    
    ' Liste des gènes à chercher dans la feuille "Mergevariant"
    genes = Array("ATP7B")

    Call FiltrerFeuilles(genes, additionalWords, "PANEL_SCU")
    
End Sub

Function FiltrerFeuilles(listeGenes As Variant, additionalWords As Variant, nomPanel As String) As Variant

    Dim dicCriteria As Object
    Dim dicCriteria1 As Object
    Dim ws As Worksheet, ws2 As Worksheet, wsAll As Worksheet
    Dim lastRow As Long, lastRow2 As Long
    Dim data As Variant, dataCNV As Variant
    Dim combinedWords As Variant
    Dim i As Long, k As Long
    Dim ii As Long, kk As Long
    Dim rng As Range, rng2 As Range

    Call InitializeAdditionalWords

    '========================
    ' PANEL NAME
    '========================
    Set wsAll = ThisWorkbook.Sheets("Allin1")
    wsAll.Range("A1").Value = nomPanel

    '========================
    ' RESET TOTAL (IMPORTANT)
    '========================
    Set ws = ThisWorkbook.Sheets("Mergevariant")
    If ws.FilterMode Then ws.ShowAllData
    If ws.AutoFilterMode Then ws.AutoFilterMode = False

    Set ws2 = ThisWorkbook.Sheets("MergeCNV")
    If ws2.FilterMode Then ws2.ShowAllData
    If ws2.AutoFilterMode Then ws2.AutoFilterMode = False

    '========================
    ' M E R G E V A R I A N T
    '========================
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    Set rng = ws.Range("A3:AA" & lastRow)

    Set dicCriteria = CreateObject("Scripting.Dictionary")
    dicCriteria.CompareMode = 1

    data = ws.Range("E2:E" & lastRow).Value

    For i = 1 To UBound(data, 1)
        For k = LBound(listeGenes) To UBound(listeGenes)
            If InStr(1, data(i, 1), listeGenes(k), vbTextCompare) > 0 Then
                dicCriteria(data(i, 1)) = True
                Exit For
            End If
        Next k
    Next i

    If dicCriteria.Count > 0 Then
        rng.AutoFilter Field:=5, _
            Criteria1:=dicCriteria.keys, _
            Operator:=xlFilterValues
    Else
        ' Ne pas afficher de gène si pas dans la liste
        rng.Rows.Hidden = False
        rng.AutoFilter Field:=5, Criteria1:="§§§NO_MATCH§§§"
    End If

    '========================
    ' M E R G E C N V
    '========================
    lastRow2 = ws2.Cells(ws2.Rows.Count, "A").End(xlUp).Row
    Set rng2 = ws2.Range("A3:AA" & lastRow2)

    combinedWords = CombineArrays(listeGenes, additionalWords)

    Set dicCriteria1 = CreateObject("Scripting.Dictionary")
    dicCriteria1.CompareMode = 1

    dataCNV = ws2.Range("E3:E" & lastRow2).Value

    For ii = 1 To UBound(dataCNV, 1)
        For kk = LBound(combinedWords) To UBound(combinedWords)
            If InStr(1, dataCNV(ii, 1), combinedWords(kk), vbTextCompare) > 0 Then
                dicCriteria1(dataCNV(ii, 1)) = True
                Exit For
            End If
        Next kk
    Next ii

    If dicCriteria1.Count > 0 Then

        rng2.AutoFilter Field:=5, _
            Criteria1:=dicCriteria1.keys, _
            Operator:=xlFilterValues

        rng2.AutoFilter Field:=13, _
            Criteria1:=">1.4", _
            Operator:=xlOr, _
            Criteria2:="<-1.4"

    Else
        rng2.Rows.Hidden = False
        rng2.AutoFilter Field:=5, Criteria1:="§§§NO_MATCH§§§"
    End If

End Function

& lastRow).PasteSpecial Paste:=xlPasteValues
    
    ' Supprime colonne db snp en trop
    Columns("U").Delete
    
' Mise en couleur en fonction de sa classification dans notre base
    Range("T4:T" & lastRow).Select
    Selection.FormatConditions.Add Type:=xlTextString, String:="Benign", _
        TextOperator:=xlContains
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
  