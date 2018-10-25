Attribute VB_Name = "ConvertPpt"
Sub TranslatePptIntoPptx()

    ' Variable Definitions
    Dim strFolder As String
    Dim strFile As String
    Dim intFileCount As Integer
    Dim objPPApplication As New PowerPoint.Application
    Dim objPPDocument As PowerPoint.Presentation

    ' Get the current directory
    strFolder = ActivePresentation.Path & "\"
    Debug.Print (strFolder)
    
    ' Load the first file using wildcard
    intFileCount = 0
    strFile = Dir(strFolder & "*.ppt", vbNormal)
  
    While strFile <> ""
        
        ' Extra check because the wildcard doesn't work properly
        If Right(strFile, 4) = ".ppt" Then
            
            With objPPApplication
                ' Open each file
                Set objPPDocument = Presentations.Open(FileName:=strFolder & strFile, ReadOnly:=True, WithWindow:=False)
    
                ' Save each file in modern format
                With objPPDocument
                    .SaveAs FileName:=strFolder & Replace(strFile, ".ppt", ".pptx")
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
    
    Set objPPDocument = Nothing
    Set objPPApplication = Nothing
End Sub
