Attribute VB_Name = "BrowseInterface"
Function BrowseButton(sTitle As String, sStartDir As String) As String

    '=================================================================================================
    ' Generate a browse dialog that can be assigned to a button.
    ' Arguments:
    '   sTitle: string title of the dialog
    '   sStartDir: default location to start browsing (17 for a folder, 4000 for a file)
    ' Return:
    '   user-chosen file or folder
    '=================================================================================================

    Dim objDialog As Object ' dialog window object
    Dim objBrowse As Object ' Browse window based on the dialog object
    
    If Right(sStartDir, 1) = "\" Then sStartDir = Left(sStartDir, Len(sStartDir) - 1)
    
    ' set browser object using browse dialog
    Set objDialog = CreateObject("Shell.Application")
    If sStartDir = "17" Then
        Set objBrowse = objDialog.BrowseForFolder(0, sTitle, &H40, 17)
    ElseIf sStartDir = "4000" Then
        Set objBrowse = objDialog.BrowseForFolder(0, sTitle, &H4000, 17)
    Else
        Set objBrowse = objDialog.BrowseForFolder(0, sTitle, &H40, CStr(sStartDir))
    End If
    
    ' set return value: append final slash if necessary, return nothing if Cancel
    If objBrowse Is Nothing Then
        BrowseButton = ""
    Else
        BrowseButton = objBrowse.ParentFolder.ParseName(objBrowse.title).Path
    End If
    
    Set objDialog = Nothing
    Set objBrowse = Nothing

End Function
