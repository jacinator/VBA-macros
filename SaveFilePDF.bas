Attribute VB_Name = "NewMacros"
Option Explicit

Sub Save_As_PDF()

    Dim strText As String
    Dim strFilename As String
    Dim cleanFilename As String
    Dim yearPart As String, fullNamePart As String, lastName As String, acctPart As String
    Dim parts() As String
    Dim fExists As Boolean
    Dim i As Integer

    '--- 1. Get and parse the content control text ---------------------
    strText = ActiveDocument.SelectContentControlsByTitle("FileName")(1).Range.Text
    parts = Split(strText, "-")

    If UBound(parts) = 2 Then
        yearPart = Left(parts(0), 4)            ' e.g., "2024" from "2024Receipt"
        fullNamePart = parts(1)                 ' e.g., "JohnSmith"
        acctPart = parts(2)                     ' e.g., "123456"

        ' Extract LastName from FirstLast (JohnSmith ? Smith)
        For i = 2 To Len(fullNamePart)
            If Mid(fullNamePart, i, 1) Like "[A-Z]" Then
                lastName = Mid(fullNamePart, i)
                Exit For
            End If
        Next i

        If lastName = "" Then lastName = fullNamePart ' fallback if no capital found

        cleanFilename = yearPart & " Receipt - " & acctPart & " - " & lastName & ".pdf"
    Else
        MsgBox "Unexpected filename format: " & strText, vbCritical
        Exit Sub
    End If

    '--- 2. Build full file path ---------------------------------------
    strFilename = ActiveDocument.Path & "\PDF Receipts\" & cleanFilename
    fExists = (Dir(strFilename) <> "")

    '--- 3. Export PDF or show Save As dialog --------------------------
    If Not fExists Then
        ActiveDocument.ExportAsFixedFormat _
            OutputFileName:=strFilename, _
            ExportFormat:=wdExportFormatPDF, _
            OpenAfterExport:=True
    Else
        MsgBox "File already exists � save manually (overwrite or rename).", vbExclamation
        With Dialogs(wdDialogFileSaveAs)
            .Name = strFilename
            .Format = wdFormatPDF
            .Show
        End With
    End If

    '--- 4. Copy file path to clipboard using Forms DataObject ---------
    Dim Clip As New DataObject
    Clip.SetText strFilename
    Clip.PutInClipboard

End Sub
