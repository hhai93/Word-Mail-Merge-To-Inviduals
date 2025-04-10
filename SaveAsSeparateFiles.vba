Sub SaveAsSeparateFiles()
    Dim i As Long
    Dim doc As Document
    Dim newDoc As Document
    Dim dataField As String
    Dim fileFormat As Long
    
    ' Replace "Name" with the column name you want to use for the file name
    dataField = "Name"
    
    ' Set file format: wdFormatPDF for PDF, wdFormatDocumentDefault for .docx, etc.
    fileFormat = wdFormatPDF  ' Change this to wdFormatDocumentDefault for .docx, wdFormatRTF for .rtf, etc.
    
    Set doc = ActiveDocument
    With doc.MailMerge
        .Destination = wdSendToNewDocument
        .Execute
    End With
    
    For i = 1 To ActiveDocument.Sections.Count - 1
        Set newDoc = Documents.Add
        ActiveDocument.Sections(i).Range.Copy
        newDoc.Content.Paste
        newDoc.SaveAs2 FileName:="C:\YourFolderPath\" & _
            doc.MailMerge.DataSource.DataFields(dataField).Value & GetFileExtension(fileFormat), _
            FileFormat:=fileFormat
        newDoc.Close
        doc.MailMerge.DataSource.ActiveRecord = wdNextRecord
    Next i
    
    ActiveDocument.Close SaveChanges:=False
End Sub

' Helper function to get file extension based on format
Function GetFileExtension(fileFormat As Long) As String
    Select Case fileFormat
        Case wdFormatPDF
            GetFileExtension = ".pdf"
        Case wdFormatDocumentDefault
            GetFileExtension = ".docx"
        Case wdFormatRTF
            GetFileExtension = ".rtf"
        Case wdFormatHTML
            GetFileExtension = ".html"
        Case Else
            GetFileExtension = ".docx" ' Default to .docx if unknown
    End Select
End Function
