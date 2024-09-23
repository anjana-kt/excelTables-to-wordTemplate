Sub WordTransferData()
    Dim DocApp As Object
    Dim DocFile As Object
    Dim DocName As String
    Dim WordRange As Object
    Dim tableIndexes As Variant
    
    DocName = "C:\path\Demo_Template.docm"
    ' Add the excel table index as start:end  
    tableIndexes = Array("C15:G21", "O14:S44", "J109:L126")
    
    On Error Resume Next
    Set DocApp = GetObject(, "Word.Application")
    If Err.Number = 429 Then
        Err.Clear
        Set DocApp = CreateObject("Word.Application")
    End If
    
    DocApp.Visible = True
    
    If Dir(DocName) = "" Then
        MsgBox "File " & DocName & vbCrLf & "not found " & vbCrLf & DocName, vbExclamation, "Document doesn't exist."
        Exit Sub
    End If
    
    DocApp.Activate
    Set DocFile = DocApp.Documents(DocName)
    
    If DocFile Is Nothing Then Set DocFile = DocApp.Documents.Open(DocName)
    
    DocFile.Activate
    ' DocFile.Range.Paste
    
    For i = 0 To UBound(tableIndexes)
        ' Copy table from excel
        Worksheets("DR Results").Range(tableIndexes(i)).Copy
        
        ' Find and delete marker
        Set WordRange = DocFile.Content
        With WordRange.Find
            .Text = "<<Table" & (i + 1) & ">>"
            .Replacement.Text = ""
            .Execute Replace:=wdReplaceOne
        End With
        
        ' Paste the the table onto marker location
        If WordRange.Find.Found Then
            WordRange.Paste
        End If
    
    Next i
    
    For i = 1 To DocFile.Tables.Count
        Set tbl = DocFile.Tables(i)
        ' Shade the background of the entire table to violet
        tbl.Shading.BackgroundPatternColor = RGB(220, 228, 244)
        
        If i > 12 Then
            Exit For ' No more table has to be colored
        End If
    Next
    
    DocFile.Save
    ' DocApp.Quit
    Set DocFile = Nothing
    Set DocApp = Nothing
    Application.CutCopyMode = True

End Sub
