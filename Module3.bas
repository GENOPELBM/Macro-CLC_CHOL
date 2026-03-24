Attribute VB_Name = "Module3"
Sub PANEL_HFE()

    Dim genes As Variant
    
    ' Liste des gènes à chercher dans la feuille "Mergevariant"
    genes = Array("SLC40A1", "BMP6", "LOC108783645, HFE", "HFE", "FTL", "HFE2", "HAMP", "TFR2")

    Call FiltrerFeuilles(genes, additionalWords)
    
End Sub

Sub PANEL_CHOL()

    Dim genes As Variant
    
    ' Liste des gènes à chercher dans la feuille "Mergevariant"
    genes = Array("LDLRAP1", "PCSK9", "APOB", "LDLR", "APOE")

    Call FiltrerFeuilles(genes, additionalWords)
    
End Sub

Sub PANEL_SCU()

    Dim genes As Variant
    
    ' Liste des gènes à chercher dans la feuille "Mergevariant"
    genes = Array("ATP7B")

    Call FiltrerFeuilles(genes, additionalWords)
    
End Sub


Function FiltrerFeuilles(listeGenes As Variant, additionalWords As Variant) As Variant

    Dim dicCriteria As Object
    Dim dicCriteria1 As Object
    Dim ColumToFilter As Variant
    Dim ColumToFilter2 As Variant
    Dim i As Integer
    Dim k As Integer
    Dim ii As Integer
    Dim kk As Integer
    Dim ws As Worksheet
    Dim ws2 As Worksheet
    Dim lastRow As Long
    Dim lastRow2 As Long
    Dim combinedWords As Variant
    Call InitializeAdditionalWords
    
    
    ' Filtrer sur la feuille "Mergevariant"
    Set ws = ThisWorkbook.Sheets("Mergevariant")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' Créer le dictionnaire
    Set dicCriteria = CreateObject("Scripting.Dictionary")
    dicCriteria.CompareMode = 1

    With ws.Range(ws.Cells(2, 5), ws.Cells(lastRow, 5))
        ColumToFilter = .Cells.Value
        For i = 1 To UBound(ColumToFilter, 1)
            For k = LBound(listeGenes) To UBound(listeGenes)
                If InStr(1, ColumToFilter(i, 1), listeGenes(k), vbTextCompare) > 0 Then
                    If Not dicCriteria.Exists(ColumToFilter(i, 1)) Then
                        dicCriteria.Add Key:=ColumToFilter(i, 1), Item:=ColumToFilter(i, 1)
                    End If
                End If
            Next k
        Next i

        ' Filtrer en utilisant les clefs du dictionnaire
        If CBool(dicCriteria.Count) Then
            ws.Range("$A$3:$AA$" & lastRow).AutoFilter Field:=5, Criteria1:=dicCriteria.keys, Operator:=xlFilterValues
        End If
    End With
    
    ' Créer dictionnaire
    Set dicCriteria1 = CreateObject("Scripting.Dictionary")
    dicCriteria1.CompareMode = 1
    
    ' APPEL DU ADD-IN ADD-IN-MERGE-PANELS
    ' BIEN RENSEIGNER LE BON additionalWords
     combinedWords = CombineArrays(listeGenes, additionalWords)
        
    
    ' Filtrer sur la feuille "MergeCNV"
    Sheets("MergeCNV").Select
    Rows("3:3").Select
    Selection.AutoFilter
    Set ws2 = ThisWorkbook.Sheets("MergeCNV")
    lastRow2 = ws2.Cells(ws2.Rows.Count, "A").End(xlUp).Row


    With Range(Cells(3, 5), Cells(lastRow2, 5))
        ColumToFilter2 = .Cells.Value
        For ii = 1 To UBound(ColumToFilter2, 1)
            For kk = LBound(combinedWords) To UBound(combinedWords)
                If InStr(1, ColumToFilter2(ii, 1), combinedWords(kk), vbTextCompare) > 0 Then
                    If Not dicCriteria1.Exists(ColumToFilter2(ii, 1)) Then
                        dicCriteria1.Add Key:=ColumToFilter2(ii, 1), Item:=ColumToFilter2(ii, 1)
                    End If
                End If
            Next kk
        Next ii
    
        ' Filtrer en utilisant les clefs du dictionnaire
        If CBool(dicCriteria1.Count) Then
            With ActiveSheet.Range("$A$3:$AA$" & lastRow2)
                .AutoFilter Field:=5, Criteria1:=dicCriteria1.keys, Operator:=xlFilterValues
                .AutoFilter Field:=13, Criteria1:=">1.4", Operator:=xlOr, Criteria2:="<-1.4"
            End With
        End If
    End With
    Sheets("Feuil1").Select
End Function

