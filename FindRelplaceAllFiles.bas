Attribute VB_Name = "FindRelplaceAllFiles"
Sub findReplaceAllFiles()
    Application.ScreenUpdating = False
    Dim doc As Document
    Dim file As String
    Const readpath As String = "\Files\"
    Const savepath As String = "\FilesWithSubs"
    Call helpers.createFolder(savepath)
    Dim newfilename As String
    
    ' Read lookup table
    Dim lookuptable As Variant
    Dim csv As Workbook
    Set csv = Workbooks.Open(filename:=ThisDocument.Path & "\lookuptable.csv")
    lookuptable = csv.Sheets(1).UsedRange.Value
    csv.Close
    
    file = Dir(PathName:=ThisDocument.Path & readpath)
    Dim counter As Integer
    counter = 1
    While file <> vbNullString
        ' Set new doc and initialize new filename
        Set doc = Documents.Open(filename:=ThisDocument.Path & readpath & file)
        newfilename = "\untitled" & counter
        counter = counter + 1
        
        ' Perform the find & replace for all elements in lookup table
        Dim i As Long
        For i = 1 To UBound(lookuptable)
            Dim find As String
            find = lookuptable(i, 1)
            Dim replace As String
            replace = lookuptable(i, 2)
            
            ' Set filename
            If find & ".docx" = file Then
                newfilename = replace
            End If
            
            doc.Windows(1).View.SeekView = wdSeekPrimaryHeader
            Call helpers.findreplace(find, replace)
            doc.Windows(1).View.SeekView = wdSeekMainDocument
            Call helpers.findreplace(find, replace)
            doc.Windows(1).View.SeekView = wdSeekPrimaryFooter
            Call helpers.findreplace(find, replace)
        Next i
        
        ' save doc and move to next file
        doc.SaveAs2 (ThisDocument.Path & savepath & "\" & newfilename)
        doc.Close
        file = Dir
    Wend
    Application.ScreenUpdating = True
End Sub
