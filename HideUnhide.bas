Attribute VB_Name = "HideUnhide"
Option Explicit
Sub count_Formula()
Dim cell As Range
Dim countF As Long

For Each cell In ActiveSheet.UsedRange
    If cell.HasFormula Then
        countF = countF + 1
    End If

Next cell
Range("B6").Value = countF
End Sub

Sub unhide_all()
Dim sh As Worksheet
For Each sh In ThisWorkbook.Worksheets
sh.Visible = True

Next sh


End Sub
