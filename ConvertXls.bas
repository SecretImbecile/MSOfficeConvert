Attribute VB_Name = "ConvertXls"
Sub TranslateXlsIntoXlsx()

    ' Variable Definitions
    Dim strFolder As String
    Dim strFile As String
    Dim intFileCount As Integer
    Dim objExcelApplication As New Excel.Application
    Dim objExcelDocument As Excel.Workbook

    ' Get the current directory
    strFolder = ActiveWorkbook.Path & "\"
    Debug.Print (strFolder)
    
    ' Load the first file using wildcard
    intFileCount = 0
    strFile = Dir(strFolder & "*.xls", vbNormal)
    
    ' Disable screen updates until completed
    Application.ScreenUpdating = False
  
    While strFile <> ""
        
        ' Extra check because the wildcard doesn't work properly
        If Right(strFile, 4) = ".xls" Then
            
            With objExcelApplication
                ' Open each file
                Set objExcelDocument = Workbooks.Open(FileName:=strFolder & strFile, ReadOnly:=True)
    
                ' Save each file in modern format
                With objExcelDocument
                    .SaveAs FileName:=strFolder & Replace(strFile, ".xls", ".xlsx"), FileFormat:=51
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
    
    Application.ScreenUpdating = True
    
    Set objExcelDocument = Nothing
    Set objExcelApplication = Nothing
End Sub

