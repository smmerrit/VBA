Attribute VB_Name = "Module1"

Public Sub FinalReportLoop()

    Dim i As Integer
    
    For i = 1 To Worksheets.Count - 1
        Worksheets(i).Select
    
        
        Range("A1").Select
        
        If ActiveCell.Value <> "" Then
        
            AddHeaders
            FormatHeaders
            AutomateSum
            
            Range("A2").Select
            Selection.CurrentRegion.Select
            Selection.Copy
            
            Worksheets("Yearly Report").Select
            
            Range("A30000").Select
            Selection.End(xlUp).Select
            
            ActiveCell.Offset(3, 0).Select
            
            ActiveSheet.Paste
            
        End If
        
    Next i
    
    Columns("C:F").EntireColumn.AutoFit

End Sub


Public Sub AutomateSum()
    Dim lastCell As String
    
    Range("F2").Select
    
    Selection.End(xlDown).Select
    lastCell = ActiveCell.Address(False, False)
    
    ActiveCell.Offset(1, 0).Select
    
    ActiveCell.Value = "=SUM(F2:" + lastCell + ")"
    ActiveCell.Font.Bold = True
End Sub

Sub AddHeaders()
'
' AddHeaders Macro
'

'
    Rows("1:1").Select
    Selection.Insert Shift:=xlDown
    ActiveWindow.SmallScroll Down:=-3
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Division"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "Category"
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "Jan"
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "Feb"
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "Mar"
    Range("F1").Select
    ActiveCell.FormulaR1C1 = "Total Expense"
    Range("A2").Select
End Sub

Sub FormatHeaders()
'
' FormatData Macro
'

'
    Range("A1:F1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = -0.249977111117893
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    Selection.Font.Bold = True
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Range("C2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Style = "Currency"
    Range("A2").Select
End Sub


