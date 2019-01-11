VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True



Private Sub clearPlanElevChart_Click()

Dim answer As Integer
answer = MsgBox("Are you sure you want to clear the chart?", vbYesNo + vbQuestion, "ClearChart")

If answer = vbYes Then
Worksheets("Plan Elev Restrictions").Range("B4:Z9999").ClearContents
End If

End Sub

Private Sub CommandButton1_Click()

Worksheets("Plan Elev in Community (JDE)").Range("A2:A200").Copy
Worksheets("Plan Elev Restrictions").Range("B2").PasteSpecial Transpose:=True

Worksheets("Plan Elev in Community (JDE)").Range("B2:B200").Copy
Worksheets("Plan Elev Restrictions").Range("B3").PasteSpecial Transpose:=True

Worksheets("Lots in Community (JDE)").Range("C2:C500").Copy
Worksheets("Plan Elev Restrictions").Range("A4").PasteSpecial Paste:=xlPasteValues

End Sub




Private Sub PlanElevUpdate_Click()

    Application.ScreenUpdating = False
    
    Dim peA As Worksheet
    Dim lpf As Worksheet
    Dim lpfN As Worksheet
    Dim lic As Worksheet
    
    
    Set peA = ActiveWorkbook.Worksheets("Plan Elev Restrictions")   ' Plan Eval Allowed
    Set lpf = ActiveWorkbook.Worksheets("Lot Plan Fit")
    Set lpfN = ActiveWorkbook.Worksheets("Lot Plan Fit Exclusions")
    Set lic = ActiveWorkbook.Worksheets("Lots in Community (JDE)")
    
    Dim rowRange As Range
    Dim colRange As Range
    
    Dim LastCol As Long
    Dim LastRow As Long
    
    LastRow = peA.Cells(peA.Rows.Count, "A").End(xlUp).Row
    LastCol = peA.Cells(2, peA.Columns.Count).End(xlToLeft).Column

    Dim rrow As Range
    Dim cell As Range
    
    Dim lot As Integer
    Dim plan As Integer
    Dim elev As Integer
    Dim comm As Double
    
    Dim lpfNextLot As Range
    Dim lpfNextComm As Range
    Dim lpfNextPlan As Range
    Dim lpfNextElev As Range
    
    Dim lcMatch As Double
    
    
    ' Set rowRange = peA.Range("B4:B40")
    Set rowRange = peA.Range(peA.Cells(4, 2), peA.Cells(LastRow, 2))
    
    'Loop through each row
    For Each rrow In rowRange
        
        ' Set colRange = peA.Range("B" & rrow.Row & ":" & LastCol & rrow.Row)
        
        Set colRange = peA.Range(peA.Cells(rrow.Row, 2), peA.Cells(rrow.Row, LastCol))
        ' Set colRange = pea.Range(Cells(rrow.Row, 2), Cells(rrow.Row, LastCol))
        
        
        'Loop through all cells in row up to last col
        For Each cell In colRange


            'Do something to each cell
            If peA.Cells(rrow.Row, cell.Column).Value = "1" Then
             ' If pea.Range(rrow.Row, cell.Column) = "1" Or "y' Then
             
             
             ' Sets offset to reach header for corresponding value
             lot = -(cell.Column - 1)
             plan = -(cell.Row - 2)
             elev = -(cell.Row - 3)
             
             ' Prints Lot
             cell.Offset(, lot).Copy
             Set lpfNextLot = lpfN.Cells(lpfN.Rows.Count, "C").End(xlUp).Offset(1, 0)
             lpfNextLot.PasteSpecial (xlValues)
             
             ' Prints Community
             lcMatch = Application.WorksheetFunction.Match(cell.Offset(0, lot).Value, lic.Range("$C$2:$C$1000"), 0)
             Application.WorksheetFunction.Index(lic.Range("$A$2:$A$1000"), lcMatch, 0).Copy
             Set lpfNextComm = lpfNextLot.Offset(0, -2) ' Pastes to column two cells left of lot inserted
             lpfNextComm.PasteSpecial (xlValues)
             
             ' Prints Plan
             cell.Offset(plan, 0).Copy
             Set lpfNextPlan = lpfNextLot.Offset(0, 1)
             lpfNextPlan.PasteSpecial (xlValues) ' Pastses to column one cell right of lot inserted
             ' lpfNextPlan.PasteSpecial (xlValues)
             
             ' Prints Elev
             cell.Offset(elev, 0).Copy
             Set lpfNextElev = lpfNextLot.Offset(0, 2)
             lpfNextElev.PasteSpecial (xlValues) ' Pastses to column two cells right of lot inserted
             ' lpfNextElev.PasteSpecial (xlValues)
             
             lpfNextLot.Offset(0, 3) = "+"
             lpfNextLot.Offset(0, 4) = "+"
             
            End If

             
             
        Next cell
    Next rrow
    
    Application.ScreenUpdating = True

End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)

End Sub

Sub PrintApplicableList()

        If peA.Cells(rrow.Row, cell.Column).Value = "" Then
             ' If pea.Cells(rrow.Row, cell.Column).Comment.Text = "1" Or "y' Then
             
             
             ' Sets offset to reach header for corresponding value
             lot = -(cell.Column - 1)
             plan = -(cell.Row - 2)
             elev = -(cell.Row - 3)
             
             ' Prints Lot
             cell.Offset(, lot).Copy
             Set lpfNextLot = lpf.Cells(lpf.Rows.Count, "C").End(xlUp).Offset(1, 0)
             lpfNextLot.PasteSpecial (xlValues)
             
             ' Prints Community
             lcMatch = Application.WorksheetFunction.Match(cell.Offset(0, lot).Value, lic.Range("$D$2:$D$1000"), 0)
             Application.WorksheetFunction.Index(lic.Range("$A$2:$A$1000"), lcMatch, 0).Copy
             Set lpfNextComm = lpfNextLot.Offset(0, -2) ' Pastes to column two cells left of lot inserted
             lpfNextComm.PasteSpecial (xlValues)
             
             ' Prints Plan
             cell.Offset(plan, 0).Copy
             Set lpfNextPlan = lpfNextLot.Offset(0, 1) ' Pastses to column one cell right of lot inserted
             lpfNextPlan.PasteSpecial (xlValues)
             
             ' Prints Elev
             cell.Offset(elev, 0).Copy
             Set lpfNextElev = lpfNextLot.Offset(0, 2) ' Pastses to column two cells right of lot inserted
             lpfNextElev.PasteSpecial (xlValues)
             
             lpfNextLot.Offset(0, 3) = "+"
             lpfNextLot.Offset(0, 4) = "+"
             

            
            End If

End Sub
