Attribute VB_Name = "Module1"

Public Sub importTextfile()
    Dim TextFile As Workbook
    Dim OpenFiles() As Variant
    Dim i As Integer
    OpenFiles = Application.GetOpenFilename(Title:="Select File(s) to import", MultiSelect:=True)
    Application.ScreenUpdating = False
    For i = 1 To Application.CountA(OpenFiles)
        Set TextFile = Workbooks.Open(OpenFiles(i))
        TextFile.Sheets(1).Range("A1").CurrentRegion.Copy
        Workbooks(1).Activate
        Workbooks(1).Worksheets.Add
        ActiveSheet.Paste
        ActiveSheet.Name = TextFile.Name
        Application.CutCopyMode = False
        TextFile.Close
    Next i
    Application.ScreenUpdating = True
    
    
End Sub
