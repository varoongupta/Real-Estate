VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True


Private Sub clearSTLchart_Click()

Dim answer As Integer
answer = MsgBox("Are you sure you want to clear the chart?", vbYesNo + vbQuestion, "ClearChart")

If answer = vbYes Then
Worksheets("STL Applicable").Range("B3:Z9999").ClearContents
End If



End Sub

Private Sub CommandButton2_Click()

Worksheets("Applicable STL Codes (JDE)").Range("H2:H50").Copy
Worksheets("STL Applicable").Range("B2").PasteSpecial Transpose:=True

Worksheets("Lots in Community (JDE)").Range("C2:C500").Copy
Worksheets("STL Applicable").Range("A3").PasteSpecial Paste:=xlPasteValues

End Sub

Private Sub stlUpdate_Click()
    
    Dim startTime As Double
    startTime = Timer
    
    Application.ScreenUpdating = False
    
    Dim peA As Worksheet
    Dim stA As Worksheet
    Dim lpf As Worksheet
    Dim lpfN As Worksheet
    Dim lic As Worksheet
    Dim AstlC As Worksheet
    
    
    Set peA = ActiveWorkbook.Worksheets("Plan Elev Restrictions")   ' Plan Eval not Allowed
    Set stA = ActiveWorkbook.Worksheets("STL Applicable") ' STL allowed
    Set lpf = ActiveWorkbook.Worksheets("Lot Plan Fit")
    Set lpfN = ActiveWorkbook.Worksheets("Lot Plan Fit Exclusions")
    Set lic = ActiveWorkbook.Worksheets("Lots in Community (JDE)")
    Set AstlC = ActiveWorkbook.Worksheets("Applicable STL Codes (JDE)")
    

    Dim rowRange As Range
    Dim colRange As Range
    
    Dim LastCol As Long
    Dim LastRow As Long
    
    LastRow = stA.Cells(stA.Rows.Count, "A").End(xlUp).Row
    LastCol = stA.Cells(2, stA.Columns.Count).End(xlToLeft).Column
    
    Dim rrow As Range
    Dim cell As Range
    
    Dim lot As Integer
    Dim stlC As Integer
    Dim elev As Integer
    Dim comm As Double
    
    Dim lpfNextLot As Range
    Dim lpfNextComm As Range
    Dim lpfNextSTL As Range
    Dim lpfNextElev As Range
    
    Dim lcMatch As Double
    
    ' Set rowRange = stA.Range("B4:B40")
    Set rowRange = stA.Range(stA.Cells(3, 2), stA.Cells(LastRow, 2))
    
    
    'Loop through each row
    
    For Each rrow In rowRange
        
        ' Set colRange = stA.Range("B" & rrow.Row & ":" & LastCol & rrow.Row)
        
        Set colRange = stA.Range(stA.Cells(rrow.Row, 2), stA.Cells(rrow.Row, LastCol))
        ' Set colRange = pea.Range(Cells(rrow.Row, 2), Cells(rrow.Row, LastCol))
        
        
        'Loop through all cells in row up to last col
        For Each cell In colRange


            'Do something to each cell
            
            If stA.Cells(rrow.Row, cell.Column).Value = "" Then
             ' If pea.Range(rrow.Row, cell.Column) = "1" Or "y' Then
             
             
             ' Sets offset to reach header for corresponding value
             lot = -(cell.Column - 1)
             stlC = -(cell.Row - 2)

             
             ' Prints Lot
             cell.Offset(, lot).Copy
             Set lpfNextLot = lpfN.Cells(lpfN.Rows.Count, "C").End(xlUp).Offset(1, 0)
             lpfNextLot.PasteSpecial (xlValues)
             
             ' Prints Community
             lcMatch = Application.WorksheetFunction.Match(cell.Offset(stlC, 0).Value, AstlC.Range("$H$2:$H$1000"), 0)
             Application.WorksheetFunction.Index(AstlC.Range("$A$2:$A$1000"), lcMatch, 0).Copy
             Set lpfNextComm = lpfNextLot.Offset(0, -2) ' Pastes to column two cells left of lot inserted
             lpfNextComm.PasteSpecial (xlValues)
             
             ' Prints STL
             cell.Offset(stlC, 0).Copy
             Set lpfNextSTL = lpfNextLot.Offset(0, 4) ' Pastses to column four cells right of lot inserted
             lpfNextSTL.PasteSpecial (xlValues)
            
             lpfNextLot.Offset(0, 1) = "+"
             lpfNextLot.Offset(0, 2) = "+"
             lpfNextLot.Offset(0, 3) = "+"
             
            End If
        Next cell
    Next rrow
    
    Application.ScreenUpdating = True
    
    MsgBox "Total time was: " & (Timer - startTime)

End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)

End Sub

Sub PrintApplicableList()

    If stA.Cells(rrow.Row, cell.Column).Value = "1" Then
        ' If pea.Range(rrow.Row, cell.Column) = "1" Or "y' Then
             
             
        ' Sets offset to reach header for corresponding value
        lot = -(cell.Column - 1)
        stlC = -(cell.Row - 2)
        elev = -(cell.Row - 3)
             
        ' Prints Lot
        cell.Offset(, lot).Copy
        Set lpfNextLot = lpf.Cells(lpf.Rows.Count, "C").End(xlUp).Offset(1, 0)
        lpfNextLot.PasteSpecial (xlValues)
             
        ' Prints Community
        lcMatch = Application.WorksheetFunction.Match(cell.Offset(stlC, 0).Value, AstlC.Range("$H$2:$H$1000"), 0)
        Application.WorksheetFunction.Index(AstlC.Range("$A$2:$A$1000"), lcMatch, 0).Copy
        Set lpfNextComm = lpfNextLot.Offset(0, -2) ' Pastes to column two cells left of lot inserted
        lpfNextComm.PasteSpecial (xlValues)
             
        ' Prints STL
        cell.Offset(stlC, 0).Copy
        Set lpfNextSTL = lpfNextLot.Offset(0, 4) ' Pastses to column four cells right of lot inserted
        lpfNextSTL.PasteSpecial (xlValues)
             
        lpfNextLot.Offset(0, 1) = "+"
        lpfNextLot.Offset(0, 2) = "+"
        lpfNextLot.Offset(0, 3) = "+"
             
            
       End If

End Sub
