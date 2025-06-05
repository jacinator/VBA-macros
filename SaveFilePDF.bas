Sub Save_As_PDF()
'
' Save_As_PDF Macro
'
'
    Dim strText As String
    'Dim strPath As String
    'strPath = "\\storage2\data\Departments\Fulfillment\Database\Year End Receipts\_2016 Year End Receipts - TEST\2016 PDF Receipts"
    
    
    
    strText = ActiveDocument.SelectContentControlsByTitle("FileName")(1).Range.Text
    
    Dim strFilename As String
    strFilename = ActiveDocument.Path & "\" & "PDF Receipts" & "\" & strText & ".pdf"
    
    Dim FileInQuestion As String
    FileInQuestion = Dir(strFilename)
    
    Dim MyDataObj As New DataObject
    MyDataObj.SetText strFilename
    
    
    If FileInQuestion = "" Then
    
    ActiveDocument.ExportAsFixedFormat OutputFileName:=strFilename, _
            ExportFormat:=wdExportFormatPDF, OpenAfterExport:=True
    
    MyDataObj.PutInClipboard
    
    'ActiveDocument.SaveAs2 FileName:=strFilename, _
        'Fileformat:=wdFormatPDF
        
   Else
         MsgBox "File Already Exists. Please manually save this file. Overwrite the previous file if you are correcting an incorrect file from earlier, or add a suffix to the end of the filename to distinguish this one from the previous file."
         
                          
         With Dialogs(wdDialogFileSaveAs)
         .Name = strFilename
         .Format = wdFormatPDF
         .Show
         End With
         
         'Application.GetSaveAsFilename
      End If
    
End Sub