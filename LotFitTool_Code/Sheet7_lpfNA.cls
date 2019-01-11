VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CommandButton1_Click()

Dim answer As Integer
answer = MsgBox("Are you sure you want to clear the chart?", vbYesNo + vbQuestion, "ClearChart")

If answer = vbYes Then
ActiveWorkbook.Worksheets("Lot Plan Fit").Range("A3:H6000").Clear
ActiveWorkbook.Worksheets("Lot Plan Fit Exclusions").Range("A3:H9999").Clear
End If

End Sub
