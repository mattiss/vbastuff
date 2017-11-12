Attribute VB_Name = "FilesToolbox"

Sub CreateNewZip(sPath)
'http://www.rondebruin.nl/win/s7/win001.htm
'Create empty Zip File
'Changed by keepITcool Dec-12-2005
    If Len(Dir(sPath)) > 0 Then Kill sPath
    Open sPath For Output As #1
    Print #1, Chr$(80) & Chr$(75) & Chr$(5) & Chr$(6) & String(18, 0)
    Close #1
End Sub

Sub ZipFilesInFolder(sFolderName As String, sZipFileName As String)
'http://www.rondebruin.nl/win/s7/win001.htm
    Dim oApp As Object
    'Create empty Zip File
    CreateNewZip sZipFileName

    Set oApp = CreateObject("Shell.Application")
    'Copy the files to the compressed folder
    oApp.Namespace((sZipFileName)).CopyHere oApp.Namespace((sFolderName)).Items

    'Keep script waiting until Compressing is done
    On Error Resume Next
    Do Until oApp.Namespace((sZipFileName)).Items.Count = _
       oApp.Namespace((sFolderName)).Items.Count
        Application.Wait (Now + TimeValue("0:00:01"))
    Loop
    On Error GoTo 0
End Sub
Sub UnzipFile(sZipFileName As String, sFolderName As String)
    Dim oApp As Object

    'Extract the files into the newly created folder
    Set oApp = CreateObject("Shell.Application")
    oApp.Namespace((sFolderName)).CopyHere oApp.Namespace((sZipFileName)).Items
End Sub

Function GetFolderFromFilePath(sFilePath As String) As String
    'given a full path and file, strip the filename off the end and return the path
    Set FSO = CreateObject("Scripting.FilesystemObject")
    GetFolderFromFilePath = FSO.GetParentFolderName(sFilePath)
End Function
Function GetFileNameFromFilePath(sFilePath As String) As String
    'given a full path and file, strip the filename off the end and return the path
    Set FSO = CreateObject("Scripting.FilesystemObject")
    GetFileNameFromFilePath = FSO.GetFileName(sFilePath)
    GetFileNameFromFilePath = Split(GetFileNameFromFilePath, ".")(0)
End Function
Function JoinPath(sFolderName As String, sFileName As String) As String
    Set FSO = CreateObject("Scripting.FileSystemObject")
    JoinPath = FSO.BuildPath(sFolderName, sFileName)
End Function
Sub TestJoinPath()
    Debug.Print (JoinPath("c:\temp", "test.txt"))
    Debug.Print (JoinPath("c:\temp\", "test.txt"))
End Sub
Sub TestGetFolderFromFilePath()
    Debug.Print (GetFolderFromFilePath("C:\Windows\System32\drivers\etc\"))
    Debug.Print (GetFolderFromFilePath("C:\Windows\System32\drivers\etc\"))
    Debug.Print (GetFolderFromFilePath("C:\Windows\System32\drivers\etc\hosts"))
End Sub
Sub TestGetFileNameFromFilePath()
    Debug.Print (GetFileNameFromFilePath("C:\Windows\System32\drivers\etc\"))
    Debug.Print (GetFileNameFromFilePath("C:\Windows\System32\drivers\etc\hosts"))
    Debug.Print (GetFileNameFromFilePath("C:\Windows\notepad.exe"))
End Sub
Sub MakeDir(sFolderName As String)
    MkDir sFolderName
End Sub
Sub RemoveDir(sFolderName As String)
    Set FSO = CreateObject("Scripting.FilesystemObject")
    FSO.deletefolder sFolderName
End Sub
