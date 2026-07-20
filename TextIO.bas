Attribute VB_Name = "TextIO"
Public iOutputFile As Integer ' file number to use for output file
Public iInputFile  As Integer
Public sOutputFile As String  ' filename (full path)

'===============================================================================================
' note: these routines work equally well for .txt or .csv, simply determined by filename passed
'===============================================================================================

Public Sub InitializeOutputFile(sFilename As String)
    iOutputFile = FreeFile()
    Open sFilename For Output As #iOutputFile
    Close #iOutputFile
End Sub

Public Sub WriteToFile(sFilename As String, sContent As String)
    Open sFilename For Append As #iOutputFile
    Print #iOutputFile, sContent
    Close #iOutputFile
End Sub

Public Sub WriteCSVRow(sFilename As String, sArray As Variant, sDelimiter)
    ' arg sArray is defined as a variant to allow initializing it via VBA's Array command,
    ' but it's expected to be string array data
    WriteToFile sFilename, Join(sArray, sDelimiter)
End Sub

Public Sub WriteCSVRowByRef(sFilename As String, sArray() As String, sDelimiter As String, iRowNum As Integer)
    ' delimit a specified "row" of a 2D array and write the concatenation into an output file
    ' dimension #2 is assumed to be a "row", so looping through sArray(0 To UBound, iRowNum)
        
    Dim sSliceArr() As String
    Dim i           As Integer
    
    ReDim sSliceArr(UBound(sArray, 1))
    For i = 0 To UBound(sArray, 1)
        sSliceArr(i) = sArray(i, iRowNum)
    Next i
    WriteToFile sFilename, Join(sSliceArr, sDelimiter)
    
End Sub

Public Sub WriteCSV2DArray(sFilename As String, sArray() As String, sDelimiter As String)
    
    Dim i       As Integer
    
    For i = 0 To UBound(sArray, 2)
        WriteCSVRowByRef sFilename, sArray, sDelimiter, i
    Next i
    
End Sub

Public Function ReadFile(sFilename As String) As String()

    Dim sRet()  As String
    Dim sLine   As String
    
    iInputFile = FreeFile()
    Open sFilename For Input As #iInputFile
    ReDim sRet(0)
    Line Input #iInputFile, sLine
    sRet(0) = sLine
    Do While Not EOF(iInputFile)
        Line Input #iInputFile, sLine
        ReDim Preserve sRet(UBound(sRet) + 1)
        sRet(UBound(sRet)) = sLine
    Loop
    Close #iInputFile
    ReadFile = sRet

End Function

