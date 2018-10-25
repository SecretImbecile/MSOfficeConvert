Attribute VB_Name = "ConvertDoc"
Sub TranslateDocIntoDocx()

    ' Variable Definitions
    Dim strFolder As String
    Dim strFile As String
    Dim intFileCount As Integer
    Dim objWordApplication As New Word.Application
    Dim objWordDocument As Word.Document

    ' Get the current directory
    strFolder = ActiveDocument.Path & "\"
    Debug.Print (strFolder)
    
    ' Load the first file using wildcard
    intFileCount = 0
    strFile = Dir(strFolder & "*.doc", vbNormal)
  
    While strFile <> ""
        
        ' Extra check because the wildcard doesn't work properly
        If Right(strFile, 4) = ".doc" Then
            
            With objWordApplication
                ' Open each file
                Set objWordDocument = .Documents.Open(FileName:=strFolder & strFile, AddToRecentFiles:=False, ReadOnly:=True, Visible:=False)
    
                ' Save each file in modern format
                With objWordDocument
                    .SaveAs2 FileName:=strFolder & Replace(strFile, ".doc", ".docx"), FileFormat:=16
                    .Close
                End With
            End With
            
            ' Show progess
            intFileCount = intFileCount + 1
            Debug.Print ("Processed " & intFileCount & " files")
        
        End If
        
        ' Load the next file
        strFile = Dir()
    
    Wend
    
    Set objWordDocument = Nothing
    Set objWordApplication = Nothing
End Sub
