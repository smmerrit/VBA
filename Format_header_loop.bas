Attribute VB_Name = "Module1"
Public Sub cleanUpData()
    Dim i As Integer
    i = 1
    Do While i <= Worksheets.Count
        Worksheets(i).Select
        AddHeaders
        FormatHeader
        
        i = i + 1
        
    
    Loop
End Sub




Sub AddHeaders()
Attribute AddHeaders.VB_Description = "This marco places headers on the worksheets"
Attribute AddHeaders.VB_ProcData.VB_Invoke_Func = "y\n14"
'
' AddHeaders Macro
' This marco places headers on the worksheets
'
' Keyboard Shortcut: Ctrl+y
'
    Rows("1:1").Select
    Selection.Insert Shift:=xlDown
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Region"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "Category"
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "Jan"
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "Feb"
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "Mar"
    Range("F1").Select
    ActiveCell.FormulaR1C1 = "Total"
    Range("A2").Select
End Sub
Sub FormatHeader()
Attribute FormatHeader.VB_Description = "Format the headers for this worksheet"
Attribute FormatHeader.VB_ProcData.VB_Invoke_Func = "u\n14"
'
' FormatHeader Macro
' Format the headers for this worksheet
'
' Keyboard Shortcut: Ctrl+u
'
    Rows("1:1").Select
    Selection.Font.Bold = True
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Region"
    Range("A1:F1").Select
    Range("F1").Activate
    Application.Width = 598.5
    Application.Height = 456
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark2
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("F1").Select
End Sub
