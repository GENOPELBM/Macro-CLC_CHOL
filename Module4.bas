Attribute VB_Name = "Module4"
Sub Combinaisonfeuilles()

    Dim wsDest As Worksheet
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim destLastRow As Long
    
'copier coller entre feuilles


    Application.ScreenUpdating = False
    
    
    ' Set the destination worksheet
    Set wsDest = ThisWorkbook.Sheets("Allin1") ' Change "DestinationSheet" to the name of your destination sheet
    
    ' Loop through each sheet except the destination sheet and "Feuil1"
    For Each ws In ThisWorkbook.Sheets
        If ws.Name <> wsDest.Name And ws.Name <> "Feuil1" Then
             ' Find the last row in the destination sheet
            destLastRow = wsDest.Cells(wsDest.Rows.Count, "A").End(xlUp).Row + 1
            
            ' Find the last row in the current sheet
            lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
            
            ' Copy Data
            ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, 26)).Copy Destination:=wsDest.Cells(destLastRow, 1)
            
    
        End If
    Next ws



    Sheets("Allin1").Select
'Format colonnes
    Columns("A:A").ColumnWidth = 11.5
    Columns("B:B").ColumnWidth = 4.5
    Columns("C:C").ColumnWidth = 4
    Columns("D:D").ColumnWidth = 21
    Columns("E:E").ColumnWidth = 23
    Columns("F:F").ColumnWidth = 6.2
    Columns("G:G").ColumnWidth = 6.2
    Columns("H:H").ColumnWidth = 6.2
    Columns("I:I").ColumnWidth = 8.2
    Columns("J:J").ColumnWidth = 6.5
    Columns("K:K").ColumnWidth = 8
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
 
'filtre par patient et colonne non vides
    ActiveSheet.Range("$A$3:$AA$" & lastRow).AutoFilter Field:=1, Criteria1:= _
        Array("X", "Sample"), Operator:=xlFilterValues

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
     
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
    
    Application.ScreenUpdating = True

End Sub




