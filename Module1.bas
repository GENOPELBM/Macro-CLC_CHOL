Attribute VB_Name = "Module1"
Sub importfichiers()

    Dim actWBB As Workbook
    Dim actDOS As String
    Dim COUVnom As String
    Dim COUVchem As String
    Dim CNVnom As String
    Dim CNVchem As String
    Dim VARIANTnom As String
    Dim VARIANTchem As String

    Application.ScreenUpdating = False
    
    
    Set actWBB = ActiveWorkbook
    actDOS = ThisWorkbook.Path
    
    
    COUVnom = Dir(actDOS & "\Merge_COUV30XCHOL*.csv")
    COUVchem = actDOS & "\" & COUVnom
    Workbooks.Open Filename:=COUVchem, Format:=4, delimiter:=";", Local:=True
    ActiveSheet.Copy After:=actWBB.Sheets(1)
    ActiveSheet.Name = "MergeCOUV"
    Rows("1:1").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Workbooks(COUVnom).Close SaveChanges:=False
    
    CNVnom = Dir(actDOS & "\Merge_CNVCHOL*.csv")
    CNVchem = actDOS & "\" & CNVnom
    'Workbooks.Open (CNVchem)
    Workbooks.Open Filename:=CNVchem, Format:=4, delimiter:=";", Local:=True
    ActiveSheet.Copy After:=actWBB.Sheets(2)
    ActiveSheet.Name = "MergeCNV"
    Rows("1:1").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Workbooks(CNVnom).Close SaveChanges:=False
    
    
    VARIANTnom = Dir(actDOS & "\Merge_VariantCHOL*.xlsx")
    VARIANTchem = actDOS & "\" & VARIANTnom
    Workbooks.Open (VARIANTchem)
    ActiveSheet.Copy After:=actWBB.Sheets(3)
    ActiveSheet.Name = "Mergevariant"
    Rows("1:1").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Workbooks(VARIANTnom).Close SaveChanges:=False
    
    Sheets("Feuil1").Select
    
    Application.ScreenUpdating = True
    

End Sub



