Attribute VB_Name = "Module2"
Sub SaveFileXLSX()
    Dim fromFileName As String
    Dim xlsxFilePath As String
    Dim xlsxFileRoot As String
    Dim xlsxFileName As String
    
    'The source filename for this macro can be found on the currently
    'active workbook. We get it from there and store it in a variable
    'so we can manipulate it.
    fromFileName = ActiveWorkbook.Name
    
    'We need to reference the file path (the current file's folder) a
    'few times over the rest of this macro. Here we retrieve it and
    'store it. Be default it doesn't end with a "\", so we add one on.
    xlsxFilePath = ActiveWorkbook.Path & Application.PathSeparator
    
    'We only need the full file name with the file type extension once,
    'to delete the file. Here we strip the extesion from the filename
    'and add the current date as a prefix.
    xlsxFileRoot = Format(Now(), "yyyy-MM-dd-") & Left(fromFileName, InStr(ActiveWorkbook.Name, ".") - 1)
    
    'If the file root that we just compiled ends with a "-C", then we
    'don't need to keep those two characters. This statement strips
    'those unneeded characters from the root.
    If Right(xlsxFileRoot, 2) = "-C" Then
        xlsxFileRoot = Left(xlsxFileRoot, Len(xlsxFileRoot) - 2)
    End If
    
    'Here we are combining all the parts for the file name that we have
    'prepared so far with the ".xlsx" file type extension. This is done
    'here just so that we aren't doing concatentation in the "SaveAs"
    'method.
    xlsxFileName = xlsxFilePath & xlsxFileRoot & ".xlsx"
    
    'Save the current workbook as an ".xlsx" file using the new file
    'name that we created.
    ActiveWorkbook.SaveAs Filename:=xlsxFileName, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    
    'This statement checks that the new  ".xlsx" file has actually been
    'created. If it has, the source file is then deleted.
    If Dir(xlsxFileName) > "" Then
        Kill xlsxFilePath & fromFileName
    End If
    
    'Finally, we want to close this workbook.
    ActiveWorkbook.Close
End Sub 'eof SaveFileXLSX
