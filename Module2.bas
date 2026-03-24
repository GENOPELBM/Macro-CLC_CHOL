Attribute VB_Name = "Module2"
Sub Miseenpagevariant()

    
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim delimiter As String
    Dim rng As Range
    Dim i As Long
    Dim col As Variant
    Dim cell As Range
    
    
    Application.ScreenUpdating = False

' Sélectionner une feuille
    Sheets("Mergevariant").Select
    
    Set ws = ThisWorkbook.Sheets("Mergevariant")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

' Réduction de l'affichage du nombre de chiffres après la virgule
    Range("J4:L" & lastRow).Select
    Selection.NumberFormat = "0.0"
    Range("N4:N" & lastRow).Select
    Selection.NumberFormat = "0.00"

' Mettre les colonne en ordre similaire à celui des couvertures et CNV
    Range("A3:X" & lastRow).Select
    Selection.Copy
    Range("AA3").Select
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ' Copier coller nom echantillon
    Range("AS3:AS" & lastRow).Select
    Selection.Copy
    Range("A3").Select
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("A4").FormulaR1C1 = "=MID(RC[45], FIND(""F-"", RC[45]) + 2, FIND(""_S"", RC[45]) - FIND(""F-"", RC[45]) - 2)"
    Range("A4:A" & lastRow).FillDown
    With Range("A4:A" & lastRow)
    .Copy
    .PasteSpecial Paste:=xlPasteValues
    End With
    ' Copier colonne Type sur colonne B
    Range("AC3:AC" & lastRow).Copy
    Range("B3:B" & lastRow).PasteSpecial Paste:=xlPasteValues
    ' Copier colonne Chromosome sur colonne C
    Range("AA3:AA" & lastRow).Copy
    Range("C3:C" & lastRow).PasteSpecial Paste:=xlPasteValues
    'Copier colonne Region sur D
    Range("AB3:AB" & lastRow).Copy
    Range("D3:D" & lastRow).PasteSpecial Paste:=xlPasteValues
    'Copier colonne gene sur E
    Range("AO3:AO" & lastRow).Copy
    Range("E3:E" & lastRow).PasteSpecial Paste:=xlPasteValues
    'Copier colonne Homopolymer sur F
    Range("AM3:AM" & lastRow).Copy
    Range("F3:F" & lastRow).PasteSpecial Paste:=xlPasteValues
    'Copier colonne reference sur G
    Range("AD3:AD" & lastRow).Copy
    Range("G3:G" & lastRow).PasteSpecial Paste:=xlPasteValues
    'Copier colonne allele alternatif sur H
    Range("AE3:AE" & lastRow).Copy
    Range("H3:H" & lastRow).PasteSpecial Paste:=xlPasteValues
    'Copier colonne Zigosity sur I
    Range("AG3:AG" & lastRow).Copy
    Range("I3:I" & lastRow).PasteSpecial Paste:=xlPasteValues
    'Copier colonne Count sur J
    Range("AH3:AH" & lastRow).Copy
    Range("J3:J" & lastRow).PasteSpecial Paste:=xlPasteValues
    'Copier colonne Coverage sur K
    Range("AI3:AI" & lastRow).Copy
    Range("K3:K" & lastRow).PasteSpecial Paste:=xlPasteValues
    'Copier colonne Frequency sur L
    Range("AJ3:AJ" & lastRow).Copy
    Range("L3:L" & lastRow).PasteSpecial Paste:=xlPasteValues
    'Copier colonne forward/reverse sur M
    Range("AK3:AK" & lastRow).Copy
    Range("M3:M" & lastRow).PasteSpecial Paste:=xlPasteValues
    'Copier colonne Avg Quality sur N
    Range("AL3:AL" & lastRow).Copy
    Range("N3:N" & lastRow).PasteSpecial Paste:=xlPasteValues
    'Copier colonne Prop singleton sur O
    Range("AN3:AN" & lastRow).Copy
    Range("O3:O" & lastRow).PasteSpecial Paste:=xlPasteValues
    'Copier colonne Exon Number sur P
    Range("AP3:AP" & lastRow).Copy
    Range("P3:P" & lastRow).PasteSpecial Paste:=xlPasteValues
    'Copier colonne Coding region sur Q
    Range("AQ3:AQ" & lastRow).Copy
    Range("Q3:Q" & lastRow).PasteSpecial Paste:=xlPasteValues
    'Copier colonne AA Change sur R
    Range("AR3:AR" & lastRow).Copy
    Range("R3:R" & lastRow).PasteSpecial Paste:=xlPasteValues
    'Copier colonne dbsnp sur S
    Range("AU3:AU" & lastRow).Copy
    Range("S3:S" & lastRow).PasteSpecial Paste:=xlPasteValues
    'Copier colonne dbxref sur T
    Range("AS3:AS" & lastRow).Copy
    Range("T3:T" & lastRow).PasteSpecial Paste:=xlPasteValues
    
    ' Supprime colonne db snp en trop
    Columns("U").Delete
    
' Mise en couleur en fonction de sa classification dans notre base
    Range("T4:T" & lastRow).Select
    Selection.FormatConditions.Add Type:=xlTextString, String:="Benign", _
        TextOperator:=xlContains
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 5287936
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlTextString, String:= _
        "Presumed Benign", TextOperator:=xlContains
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 5296274
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlTextString, String:= _
        "Unknown Significance", TextOperator:=xlContains
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlTextString, String:= _
        "Presumed Pathogenic", TextOperator:=xlContains
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 36607
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlTextString, String:="Pathogenic", _
        TextOperator:=xlContains
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 255
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlTextString, String:="Artefact", _
        TextOperator:=xlContains
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.349986266670736
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
  
    
' Format colonnes
    Columns("A:A").ColumnWidth = 11.5
    Columns("B:B").ColumnWidth = 4.5
    Columns("C:C").ColumnWidth = 4
    Columns("D:D").ColumnWidth = 21
    Columns("E:E").ColumnWidth = 16
    Columns("F:F").ColumnWidth = 3
    Columns("G:G").ColumnWidth = 6.2
    Columns("H:H").ColumnWidth = 6.2
    Columns("I:I").ColumnWidth = 8.2
    Columns("J:J").ColumnWidth = 6.5
    Columns("K:K").ColumnWidth = 12.5
    Columns("L:L").ColumnWidth = 7
    Columns("M:M").ColumnWidth = 7
    Columns("N:N").ColumnWidth = 7
    Columns("O:O").ColumnWidth = 7
    Columns("P:P").ColumnWidth = 7
    Columns("Q:Q").ColumnWidth = 25
    Columns("R:R").ColumnWidth = 25
    Columns("S:S").ColumnWidth = 15
    Columns("T:T").ColumnWidth = 17
    Columns("U:U").ColumnWidth = 11
    
' A mettre à la fin pour chaque feuille pour séparer les feuille sur la feuille combinée
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "X"
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "X"
    Range("A3").Select
    ActiveCell.FormulaR1C1 = "Sample"
    Range("A1:A2").Select
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    

'nom du fichier 2 premières lignes
    Range("B1").Select
    Selection.ClearContents
    ActiveCell.FormulaR1C1 = _
        "=MID(CELL(""nomfichier""),36,FIND(""]"",CELL(""nomfichier"")))"
    Selection.Copy
    Range("B2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("B1").Select
    Selection.ClearContents
    Range("K1").Select
    ActiveCell.FormulaR1C1 = "CHOLV4"
    Range("M1").Select
    ActiveCell.FormulaR1C1 = "Fait par : "
    
'mise en couleur de l'entète de la table en gris
    Range("A3:T3").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark2
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With

'mise en page ==> retour à la ligne
    Range("A3:Z" & lastRow).Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlTop
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
'mise en page impression
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .PrintTitleRows = ""
        .PrintTitleColumns = ""
    End With
    Application.PrintCommunication = True
    ActiveSheet.PageSetup.PrintArea = ""
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .LeftHeader = ""
        .CenterHeader = ""
        .RightHeader = ""
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = ""
        .LeftMargin = Application.InchesToPoints(0.25)
        .RightMargin = Application.InchesToPoints(0.25)
        .TopMargin = Application.InchesToPoints(0.75)
        .BottomMargin = Application.InchesToPoints(0.75)
        .HeaderMargin = Application.InchesToPoints(0.3)
        .FooterMargin = Application.InchesToPoints(0.3)
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        .PrintQuality = 600
        .CenterHorizontally = False
        .CenterVertically = False
        .Orientation = xlLandscape
        .Draft = False
        .PaperSize = xlPaperA4
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 0
        .PrintErrors = xlPrintErrorsDisplayed
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .ScaleWithDocHeaderFooter = True
        .AlignMarginsHeaderFooter = True
        .EvenPage.LeftHeader.Text = ""
        .EvenPage.CenterHeader.Text = ""
        .EvenPage.RightHeader.Text = ""
        .EvenPage.LeftFooter.Text = ""
        .EvenPage.CenterFooter.Text = ""
        .EvenPage.RightFooter.Text = ""
        .FirstPage.LeftHeader.Text = ""
        .FirstPage.CenterHeader.Text = ""
        .FirstPage.RightHeader.Text = ""
        .FirstPage.LeftFooter.Text = ""
        .FirstPage.CenterFooter.Text = ""
        .FirstPage.RightFooter.Text = ""
    End With
    Application.PrintCommunication = True

    
' Supprimer tableau original
    Range("Z3:AU" & lastRow).Delete
    
' Convertir virgule  en point virgule
    Range("S4:S" & lastRow).Select
    Selection.Replace What:=",", Replacement:=";", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:=".", Replacement:=",", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.TextToColumns Destination:=Range("W4"), DataType:=xlDelimited, _
        TextQualifier:=xlNone, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=True, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(Array(1, 1), Array(2, 1)), TrailingMinusNumbers:=True
        
        
   ' Étape 4 : Arrondir les colonnes W, X, Y, Z à 2 décimales
    For Each col In Array("W", "X", "Y", "Z")
        ' Créer la plage dynamique pour chaque colonne
        Set rng = ws.Range(col & "4:" & col & lastRow)
        
        ' Vérifier et arrondir les cellules numériques
        For Each cell In rng
            If IsNumeric(cell.Value) Then
                cell.Value = Application.WorksheetFunction.Round(cell.Value, 2)
            End If
        Next cell
    Next col
    
    ' Supprimer les cellules vides ou contenant seulement des virgules
    For Each col In Array("W", "X", "Y", "Z")
        Set rng = ws.Range(col & "4:" & col & lastRow)
        For Each cell In rng
            If cell.Value = "," Or IsEmpty(cell.Value) Then
                cell.ClearContents
            End If
        Next cell
    Next col

' Supprimer colonnes
    Range("W3:Z" & lastRow).Delete
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

'sélectionner une feuille
    Sheets("MergeCNV").Select
   
    
    Set ws = ThisWorkbook.Sheets("MergeCNV")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
'sélectionner le numéro des échantillons
    Range("Z4").Select
    ActiveCell.FormulaR1C1 = _
        "=MID(RC[-25],FIND(""ck"",RC[-25])+2,FIND(""_S"",RC[-25])-FIND(""ck"",RC[-25])-2)"
    Range("Z4").AutoFill Destination:=Range("Z4:Z" & lastRow)
       
'insérer une colonne et copier le nom des échantillons
    Columns(1).Insert
    Range("AA4:AA" & lastRow).Select
    Selection.Copy
    Range("A4").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

'macro CNV à adapter
    Rows("3:3").Select
    Selection.AutoFilter
    ActiveSheet.Range("$A$3:$T$" & lastRow).AutoFilter Field:=13, Criteria1:=">1.4", _
        Operator:=xlOr, Criteria2:="<-1.4"
        
    
    Range("M4:M" & lastRow).Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLessEqual _
        , Formula1:="=-1,41"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 6264043
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreaterEqual, _
        Formula1:="=1,41"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
' Permutation colonnes Comments et Regional P value

    Sheets("MergeCNV").Range("S1:S" & lastRow).Cut Destination:=Range("AB1")
    Sheets("MergeCNV").Range("V1:V" & lastRow).Cut Destination:=Range("S1")
    Sheets("MergeCNV").Range("AB1:AB" & lastRow).Cut Destination:=Range("V1")
    
    Sheets("MergeCNV").Range("Q1:Q" & lastRow).Cut Destination:=Range("AC1")
    Sheets("MergeCNV").Range("T1:T" & lastRow).Cut Destination:=Range("Q1")
    Sheets("MergeCNV").Range("AC1:AC" & lastRow).Cut Destination:=Range("T1")

'nom du fichier 2 premières lignes
    Range("B1").Select
    Selection.ClearContents
    ActiveCell.FormulaR1C1 = _
        "=MID(CELL(""nomfichier""),36,FIND(""]"",CELL(""nomfichier"")))"
    Selection.Copy
    Range("B2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("B1").Select
    Selection.ClearContents
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "Macro CNV"

'mise en couleur de l'entète de la table en gris
    Range("A3:V3").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark2
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
'fin macro CNV à adapter

'mettre en forme les nombres à virgule
    Range("H4:I" & lastRow).Select
    Selection.NumberFormat = "0"
    Range("K4:O" & lastRow).Select
    Selection.NumberFormat = "0.00"
    Range("R4:S" & lastRow).Select
    Selection.NumberFormat = "0.00"

'à mettre à la fin pour séparer les feuille sur la feuille combinée
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "X"
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "X"
    Range("A3").Select
    ActiveCell.FormulaR1C1 = "Sample"
    Range("A1:A2").Select
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    Range("B3").Select
    ActiveCell.FormulaR1C1 = "Type"
    Range("B4:B" & lastRow).Select
    Selection.ClearContents
    Range("B4:B" & lastRow).Formula = "CNV"
    Range("AA4:AA" & lastRow).Select
    Selection.ClearContents

'Format colonnes
    Columns("A:A").ColumnWidth = 11.5
    Columns("B:B").ColumnWidth = 4.5
    Columns("C:C").ColumnWidth = 4
    Columns("D:D").ColumnWidth = 21
    Columns("E:E").ColumnWidth = 16
    Columns("F:F").ColumnWidth = 3
    Columns("G:G").ColumnWidth = 6.2
    Columns("H:H").ColumnWidth = 6.2
    Columns("I:I").ColumnWidth = 8.2
    Columns("J:J").ColumnWidth = 6.5
    Columns("K:K").ColumnWidth = 10
    Columns("L:L").ColumnWidth = 7
    Columns("M:M").ColumnWidth = 7
    Columns("N:N").ColumnWidth = 7
    Columns("O:O").ColumnWidth = 7
    Columns("P:P").ColumnWidth = 7
    Columns("Q:Q").ColumnWidth = 18
    Columns("R:R").ColumnWidth = 18
    Columns("S:S").ColumnWidth = 7
    Columns("T:T").ColumnWidth = 11
    Columns("U:U").ColumnWidth = 11

'mise en page ==> retour à la ligne
    Range("A3:Z" & lastRow).Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlTop
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With

'mise en page impression
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .PrintTitleRows = ""
        .PrintTitleColumns = ""
    End With
    Application.PrintCommunication = True
    ActiveSheet.PageSetup.PrintArea = ""
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .LeftHeader = ""
        .CenterHeader = ""
        .RightHeader = ""
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = ""
        .LeftMargin = Application.InchesToPoints(0.25)
        .RightMargin = Application.InchesToPoints(0.25)
        .TopMargin = Application.InchesToPoints(0.75)
        .BottomMargin = Application.InchesToPoints(0.75)
        .HeaderMargin = Application.InchesToPoints(0.3)
        .FooterMargin = Application.InchesToPoints(0.3)
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        .PrintQuality = 600
        .CenterHorizontally = False
        .CenterVertically = False
        .Orientation = xlLandscape
        .Draft = False
        .PaperSize = xlPaperA4
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 0
        .PrintErrors = xlPrintErrorsDisplayed
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .ScaleWithDocHeaderFooter = True
        .AlignMarginsHeaderFooter = True
        .EvenPage.LeftHeader.Text = ""
        .EvenPage.CenterHeader.Text = ""
        .EvenPage.RightHeader.Text = ""
        .EvenPage.LeftFooter.Text = ""
        .EvenPage.CenterFooter.Text = ""
        .EvenPage.RightFooter.Text = ""
        .FirstPage.LeftHeader.Text = ""
        .FirstPage.CenterHeader.Text = ""
        .FirstPage.RightHeader.Text = ""
        .FirstPage.LeftFooter.Text = ""
        .FirstPage.CenterFooter.Text = ""
        .FirstPage.RightFooter.Text = ""
    End With
    Application.PrintCommunication = True

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
        
'sélectionner une feuille
    Sheets("MergeCOUV").Select
    
    Set ws = ThisWorkbook.Sheets("MergeCOUV")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
'sélectionner le numéro des échantillons
    Range("Z4").Select
    ActiveCell.FormulaR1C1 = _
        "=MID(RC[-25],FIND("" -"",RC[-25])+2,FIND(""_S"",RC[-25])-FIND("" -"",RC[-25])-2)"
    Range("Z4").AutoFill Destination:=Range("Z4:Z" & lastRow)
    
'insérer une colonne et copier le nom des échantillons
    Columns(1).Insert
    Range("AA4:AA" & lastRow).Select
    Selection.Copy
    Range("A4").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

    
    
       
 'macro COUV à adapter
    Sheets("MergeCOUV").Select
    Rows("3:3").Select
    Selection.AutoFilter
    ' Coller la formule en S4
    ws.Range("T4").Activate
    ActiveCell.FormulaR1C1 = "=IF(OR(RC[-12]<100, ISNUMBER(SEARCH(""x"", RC[-17])), ISNUMBER(SEARCH(""y"", RC[-17]))), ""oui"", """")"
    ' Étendre la formule jusqu'à la dernière ligne
    ws.Range("T4:T" & lastRow).FillDown
    ActiveSheet.Range("$A$3:$T$" & lastRow).AutoFilter Field:=20, Criteria1:="oui"
    ' Masquer la nouvelle colonne (caractères en blanc)
    ws.Columns("T:T").Select
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    
    
'nom du fichier 2 premières lignes
    Range("B1").Select
    Selection.ClearContents
    ActiveCell.FormulaR1C1 = _
        "=MID(CELL(""nomfichier""),36,FIND(""]"",CELL(""nomfichier"")))"
    Selection.Copy
    Range("B2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("B1").Select
    Selection.ClearContents
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "Macro COUV"
    
'mise en couleur de l'entète de la table en gris
    Range("A3:S3").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark2
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
'fin macro COUV à adapter

'mettre en forme les nombres à virgule
    Range("K4:K5" & lastRow).Select
    Selection.NumberFormat = "0.0"
    Range("N4:N" & lastRow).Select
    Selection.NumberFormat = "0.0"
    Range("Q4:Q" & lastRow).Select
    Selection.NumberFormat = "0.0"

'à mettre à la fin pour séparer les feuille sur la feuille combinée
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "X"
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "X"
    Range("A3").Select
    ActiveCell.FormulaR1C1 = "Sample"
    Range("A1:A2").Select
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    Range("B3").Select
    ActiveCell.FormulaR1C1 = "Type"
    Range("B4:B" & lastRow).Select
    Selection.ClearContents
    Range("B4:B" & lastRow).Formula = "couv"
    Range("AA4:AA" & lastRow).Select
    Selection.ClearContents

'Format colonnes
    Columns("A:A").ColumnWidth = 11.5
    Columns("B:B").ColumnWidth = 4.5
    Columns("C:C").ColumnWidth = 4
    Columns("D:D").ColumnWidth = 21
    Columns("E:E").ColumnWidth = 23
    Columns("F:F").ColumnWidth = 3
    Columns("G:G").ColumnWidth = 6.2
    Columns("H:H").ColumnWidth = 6.2
    Columns("I:I").ColumnWidth = 8.2
    Columns("J:J").ColumnWidth = 6.5
    Columns("K:K").ColumnWidth = 10
    Columns("L:L").ColumnWidth = 7
    Columns("M:M").ColumnWidth = 7
    Columns("N:N").ColumnWidth = 7
    Columns("O:O").ColumnWidth = 7
    Columns("P:P").ColumnWidth = 7
    Columns("Q:Q").ColumnWidth = 11
    Columns("R:R").ColumnWidth = 7
    Columns("S:S").ColumnWidth = 7
    Columns("T:T").ColumnWidth = 11
    Columns("U:U").ColumnWidth = 11

'mise en page ==> retour à la ligne
    Range("A3:Z" & lastRow).Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlTop
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
'mise en page impression
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .PrintTitleRows = ""
        .PrintTitleColumns = ""
    End With
    Application.PrintCommunication = True
    ActiveSheet.PageSetup.PrintArea = ""
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .LeftHeader = ""
        .CenterHeader = ""
        .RightHeader = ""
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = ""
        .LeftMargin = Application.InchesToPoints(0.25)
        .RightMargin = Application.InchesToPoints(0.25)
        .TopMargin = Application.InchesToPoints(0.75)
        .BottomMargin = Application.InchesToPoints(0.75)
        .HeaderMargin = Application.InchesToPoints(0.3)
        .FooterMargin = Application.InchesToPoints(0.3)
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        .PrintQuality = 600
        .CenterHorizontally = False
        .CenterVertically = False
        .Orientation = xlLandscape
        .Draft = False
        .PaperSize = xlPaperA4
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 0
        .PrintErrors = xlPrintErrorsDisplayed
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .ScaleWithDocHeaderFooter = True
        .AlignMarginsHeaderFooter = True
        .EvenPage.LeftHeader.Text = ""
        .EvenPage.CenterHeader.Text = ""
        .EvenPage.RightHeader.Text = ""
        .EvenPage.LeftFooter.Text = ""
        .EvenPage.CenterFooter.Text = ""
        .EvenPage.RightFooter.Text = ""
        .FirstPage.LeftHeader.Text = ""
        .FirstPage.CenterHeader.Text = ""
        .FirstPage.RightHeader.Text = ""
        .FirstPage.LeftFooter.Text = ""
        .FirstPage.CenterFooter.Text = ""
        .FirstPage.RightFooter.Text = ""
    End With
    Application.PrintCommunication = True
    
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

    ' Vérifier la position des entêtes dans chaque feuille et colorier les cases en vert si pas d'erreur
    If ThisWorkbook.Sheets("MergeCNV").Range("M3").Value = "Fold-change (adjusted)" Then
        With ThisWorkbook.Sheets("Feuil1").Range("L16:L20").Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent6
            .TintAndShade = 0.399975585192419
            .PatternTintAndShade = 0
        End With
    End If
    
    If ThisWorkbook.Sheets("MergeCOUV").Range("H3").Value = "Percentage with coverage above 30" Then
        With ThisWorkbook.Sheets("Feuil1").Range("L10:L15").Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent6
            .TintAndShade = 0.399975585192419
            .PatternTintAndShade = 0
        End With
    End If
    
    If ThisWorkbook.Sheets("Mergevariant").Range("E3").Value = "Homo_sapiens_refseq_GRCh37.p13_Genes" Or ThisWorkbook.Sheets("Mergevariant").Range("E3").Value = "gene (Homo_sapiens_refseq_GRCh37.p13_Genes)" Then
        With ThisWorkbook.Sheets("Feuil1").Range("L7:L9").Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent6
            .TintAndShade = 0.399975585192419
            .PatternTintAndShade = 0
        End With
    End If
    
    Sheets("Feuil1").Select
    Application.ScreenUpdating = True
    

End Sub





