VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub OrganizeList_Click()

Application.ScreenUpdating = False

    Dim sheet As Worksheet
    
    Set sheet = ActiveWorkbook.Worksheets("Sheet1")

    Dim AllRows As Range
    Dim rrow As Range
    Dim nextRole As Range
    
    Dim lastCol As Long
    Dim pasteRow As Long
    Dim counter As Long
    
    
    Set AllRows = sheet.Range("A2:A3000")  ' make smaller for test purposes
    
    For Each rrow In AllRows
    
    counter = 2                             ' plus two because first col is empty, therfore only 1 used to start
    pasteRow = rrow.row - 1
    
    'If rrow.Value2 <> "" And sheet.Cells(rrow.row, "B").Value2 <> "" Then
       
    
        Do While rrow.Value2 = ""
        
            lastCol = sheet.Cells(rrow.row - 1, rrow.Columns.Count).End(xlToLeft).Column
            lastCol = lastCol + counter
            rrow.Offset(0, 1).Copy
            Set nextRole = sheet.Cells(pasteRow, lastCol)
            nextRole.PasteSpecial (xlValues)
            rrow.EntireRow.Delete Shift:=xlUp
            Set rrow = sheet.Cells(pasteRow + 1, "A")
            counter = counter + 1
            
        Loop
        
    'End If
        
    Next rrow
        
    
Application.ScreenUpdating = True

End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)



End Sub
