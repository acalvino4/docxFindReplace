Attribute VB_Name = "helpers"
Public Function createFolder(folderName As String)
    Dim folder As Object
    Set folder = CreateObject("Scripting.FileSystemObject")
    If Not folder.FolderExists(ThisDocument.Path & folderName) Then
        folder.createFolder (ThisDocument.Path & folderName)
    End If
End Function
Public Function findreplace(strFind As String, strReplace As String)
    With Selection.find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = strFind
        .Replacement.Text = strReplace
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        
        .Execute replace:=wdReplaceAll
    End With
End Function

