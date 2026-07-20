Attribute VB_Name = "ExtractConfigDetails1"
Option Explicit

Dim sOutputFilename As String

Sub main()

    Dim swApp       As SldWorks.SldWorks
    Dim sInputDir   As String
    Dim vStyle      As Variant ' recursive or not
    Dim bRecursive  As Boolean
    Dim oFSO        As New Scripting.FileSystemObject
    Dim oWscript    As Object
    Dim sDesktopDir As String

    Set swApp = Application.SldWorks
    
    Set oWscript = CreateObject("WScript.Shell")
    sDesktopDir = oWscript.SpecialFolders("Desktop")
    Set oWscript = Nothing
    sOutputFilename = sDesktopDir & "\ExtractConfigDetails_output.csv"
    If Dir(sOutputFilename) <> "" Then
        MsgBox "Output file 'ExtractConfigDetails_output.csv' already exists on the desktop. Please delete and retry."
        Exit Sub
    End If
    'Debug.Print "sOutputFilename = " & sOutputFilename
    InitializeOutputFile sOutputFilename
    'WriteCSVRow sOutputFilename, Array("Folder", "Filename", "Search term (if it exists in Enovia)"), ","
    WriteCSVRow sOutputFilename, Array("Filename", "Raw configuration name", "Instance object title"), ","
    
    InitializeDocumentManager
    sInputDir = BrowseButton("Browse to folder containing Solidworks files", "C:\3DEXPERIENCE")
    'sInputDir = "C:\ECM_SW_extracts\sample" ' testing
    'Debug.Print "sInputDir = " & sInputDir
    If sInputDir = "" Then Exit Sub
    
    'vStyle = MsgBox("Scan just this folder? ('No' means all subdirectories, recursively)", vbYesNoCancel)
    vStyle = vbNo
    Debug.Print Now & " - starting"
    If vStyle = vbCancel Then Exit Sub
    bRecursive = IIf(vStyle = vbYes, False, True)
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    ' look in folder (including recursive if user said so, and read files inside)
    Traverse oFSO.GetFolder(sInputDir), bRecursive
    
    Debug.Print Now & " - finished"
    
    MsgBox "Done! See file 'ExtractConfigDetails_output.csv' on your desktop."
    
End Sub

Sub Traverse(oFolder As Object, bRecursive As Boolean)
    
    Dim oSubFolder As Object
    Dim oFile As Object
        
    For Each oFile In oFolder.Files
        EvaluateFile oFile.Path
    Next
    
    If bRecursive Then
        For Each oSubFolder In oFolder.SubFolders
            Traverse oSubFolder, bRecursive
        Next
    End If
    
End Sub

Sub EvaluateFile(sFilename As String)

    Dim dmDoc           As SwDMDocument29
    Dim dmDocOpenErr    As SwDmDocumentOpenError
    Dim swDMSearchOpt   As SwDMSearchOption
    Dim vCompRefs       As Variant
    Dim i               As Integer
    'Dim sDirPath        As String
    Dim sFilenameShort  As String
    'Dim sFilenameSearch As String
    Dim dmConfMgr       As SwDMConfigurationMgr2
    Dim vConfigNames    As Variant ' string array
    Dim sInstanceName   As String
    
    'Debug.Print "evaluating " & sFilename
    
    'sDirPath = Left(sFilename, InStrRev(sFilename, "\") - 1)
    sFilenameShort = Mid(sFilename, InStrRev(sFilename, "\") + 1)
    'sFilenameSearch = Left(sFilenameShort, InStrRev(sFilenameShort, ".")) & "*"
    If Dir(sFilename) = "" Then Exit Sub
    
    'If Right(UCase(sFilename), 7) <> ".SLDPRT" And Right(UCase(sFilename), 7) <> ".SLDASM" And Right(UCase(sFilename), 7) <> ".SLDDRW" Then Exit Sub
    If Right(UCase(sFilename), 7) <> ".SLDPRT" And Right(UCase(sFilename), 7) <> ".SLDASM" Then Exit Sub
    
    Set dmDoc = dmApp.GetDocument(sFilename, swDmDocumentUnknown, False, dmDocOpenErr)
    'If dmDocOpenErr <> swDmDocumentOpenErrorFileReadOnly Then
    '    WriteCSVRow sOutputFilename, Array(sDirPath, sFilenameShort, sFilenameSearch), ","
    'End If
    
    vConfigNames = Null
    sInstanceName = ""
    If Not dmDoc Is Nothing Then
        Set dmConfMgr = dmDoc.ConfigurationManager
        vConfigNames = dmConfMgr.GetConfigurationNames2(Empty)
    End If
    
    If IsArray(vConfigNames) Then
        If Not IsEmpty(vConfigNames) Then
            For i = 0 To UBound(vConfigNames)
                sInstanceName = vConfigNames(i) & "(" & Left(sFilenameShort, InStrRev(sFilenameShort, ".") - 1) & ")"
                WriteCSVRow sOutputFilename, Array(sFilenameShort, vConfigNames(i), sInstanceName), ","
            Next i
        End If
    End If
    
    If Not dmDoc Is Nothing Then dmDoc.CloseDoc

End Sub
