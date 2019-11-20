Attribute VB_Name = "Module1"
Public Sub UserSortInput()
    Dim userInput As String
    Dim promptmsg As String
    Dim tryagain As Integer
    
    promptmsg = "Enter a numeric value to sort..." & vbCrLf & _
        "1 --- Sort by dvision" & vbCrLf & _
        "2 --- Sort by catagory" & vbCrLf & _
        "3 --- Sort by Total"
        
    userInput = InputBox(promptmsg)
    
    If userInput = "1" Then
        DivisionSort
    ElseIf userInput = "2" Then
        CategorySort
    ElseIf userInput = "3" Then
        TotalSort
    Else
        tryagain = MsgBox("Invalid value try again", vbYesNo)
        If tryagain = 6 Then
            UserSortInput
        
        End If
        
    
    
    End If
    
        
        
    
    
End Sub




Sub DivisionSort()
Attribute DivisionSort.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Sort List by Division Ascending
'

'
    Selection.Sort Key1:=Range("A4"), Order1:=xlAscending, Header:=xlGuess, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
        DataOption1:=xlSortNormal

End Sub

Sub CategorySort()
'
' Sort List by Category Ascending
'

'
    Selection.Sort Key1:=Range("B4"), Order1:=xlAscending, Header:=xlGuess, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
        DataOption1:=xlSortNormal

End Sub

Sub TotalSort()
'
' Sort List by Total Sales Ascending
'

'
    Selection.Sort Key1:=Range("F4"), Order1:=xlAscending, Header:=xlGuess, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
        DataOption1:=xlSortNormal

End Sub

























