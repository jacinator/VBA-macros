Attribute VB_Name = "Module1"
Sub DateFileName()
    Dim fromFileName As String
    Dim docxFilePath As String
    Dim docxFileRoot As String
    Dim docxFileName As String
    
    'The source filename for this macro can be found on the currently
    'active document. We get it from there and store it in a variable
    'so we can manipulate it.
    fromFileName = ActiveDocument.Name
    
    'We need to reference the file path (the current file's folder) a
    'few times over the rest of this macro. Here we retrieve it and
    'store it. Be default it doesn't end with a "\", so we add one on.
    docxFilePath = ActiveDocument.Path & Application.PathSeparator
    
    'This macro's purpose is to add a date format prefix to the file
    'name. We're adding that prefix here.
    docxFileRoot = Format(Now(), "yyyy-MM-dd-") & fromFileName
    
    'Here we are combining the path with the file name. This is done
    'here just so that we aren't doing concatentation in the "SaveAs"
    'method.
    docxFileName = docxFilePath & docxFileRoot
    
    ActiveDocument.SaveAs2 FileName:=docxFileName
    
    If Dir(docxFileName) > "" Then
        Kill docxFilePath & fromFileName
    End If
    
    ActiveDocument.Close
End Sub
