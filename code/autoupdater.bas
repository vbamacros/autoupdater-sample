Attribute VB_Name = "Autoupdater"
Option Explicit
Const LOCAL_VERSION_TAG = "v1.0.0"
Const RELEASES_URL$ = "https://api.github.com/repos/vbamacros/autoupdater-sample/releases/"

Private Sub Autoupdate()
    
    ' Remember to include this in the "ThisWorkbook" module:
    ' Private Sub Workbook_Open()
    ' Application.Run ("Autoupdater.Autoupdate")
    ' End Sub
    
    On Error GoTo Quit
    
    ' Dim all objects first, so they can be set to Nothing on Quit
    Dim remoteInfo
    
    ' Check for updates
    Set remoteInfo = fetchLatestReleaseInfo()
    
    Dim isUpToDate As Boolean
    isUpToDate = (LOCAL_VERSION_TAG = remoteInfo("obj.tag_name"))
    If isUpToDate Then GoTo Quit
    
    ' Download
    Dim tempDirPath As String: tempDirPath = NewTempDir()
    If tempDirPath = "" Then GoTo Quit 'Something went wrong

    Dim downloadUrl$, downloadName$, pathToNewFile$
    downloadUrl = remoteInfo("obj.assets(0).browser_download_url")
    downloadName = remoteInfo("obj.assets(0).name")
    pathToNewFile = DownloadFile(downloadUrl, tempDirPath & "\" & downloadName)
    
    If pathToNewFile = "" Then GoTo Quit
    
    ' Get consent
    If ThisWorkbook.Saved = False Then GoTo Quit
    
    Dim proceed
    proceed = MsgBox( _
            "An update is available for this macro." & vbNewLine & _
            "Do you want to install it?", vbYesNo, "Autoupdater" _
        )
    If proceed <> vbYes Then GoTo Quit
    
    ' Install
    Dim scriptPath$: scriptPath = MakeBatScript(tempDirPath)
    shell "CMD.exe /C " & Chr(34) & scriptPath & Chr(34), vbHide
    
    ThisWorkbook.Close False
    
    Exit Sub
    
Quit:
    DeleteAllMyTempDirs
    Set remoteInfo = Nothing
    On Error GoTo 0
End Sub

Function MakeBatScript(tempDir)
    Dim batFilePath$: batFilePath = tempDir & "\replace-and-restart.bat"
    Dim updated$: updated = tempDir & "\" & ThisWorkbook.Name
    Dim here$: here = ThisWorkbook.FullName ' First: old file; after replacement: new file.
    
    Dim file&: file = FreeFile
    
    Open batFilePath For Output As #file
    Print #file, "@ECHO off"
    Print #file, "TIMEOUT 5"
    ' Replace old with new (/Y = yes, overwrite)
    Print #file, Replace("COPY /Y ~" & updated & "~ ~" & here & "~", "~", Chr(34))
    ' Re-open this (updated) workbook
    Print #file, Replace("START /I ~excel.exe~ ~" & here, "~", Chr(34))
    ' Delete temporary folder and files (/S = all subfolders/files; /Q = Quiet)
    Print #file, Replace("RD /S /Q ~" & tempDir & "~", "~", Chr(34))
    Close #file
    
    MakeBatScript = batFilePath
End Function

Function DownloadFile(url, fullDestinationPath) As String
    On Error GoTo Finally
    
    Dim newFilePath$: newFilePath = ""
    
    Dim XML: Set XML = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    XML.Open "GET", url, False
    XML.send
    
    If XML.Status <> 200 Then GoTo Finally
    
    ' Succesful request, save file
    With CreateObject("ADODB.Stream")
        .Open: .Type = 1: .Write XML.responseBody
        .SaveToFile fullDestinationPath, 2 'adSaveCreateOverWrite
    End With
    
    newFilePath = fullDestinationPath
    
Finally:
    On Error GoTo 0
    Set XML = Nothing
    DownloadFile = newFilePath ' May return empty string
End Function

Function fetchLatestReleaseInfo()
    Dim url$: url = RELEASES_URL & "latest"
    
    With CreateObject("MSXML2.ServerXMLHTTP.6.0")
        .Open "GET", url, False
        .send
    Dim json$: json = .responseText
    End With
    
    Set fetchLatestReleaseInfo = ParseJSON(json)
End Function

Private Sub DeleteAllMyTempDirs()
    Dim parentPath$: parentPath = Environ("TEMP") & "\"
    Dim pattern$: pattern = "vbaAutoupdater_*"
    
    If Dir(parentPath, vbDirectory) = "" Then Exit Sub ' Doesn't exist
    
    Dim nextSubDir$: nextSubDir = Dir(parentPath & pattern, vbDirectory)
    With CreateObject("Scripting.FileSystemObject")
        Do While nextSubDir <> ""
            .DeleteFolder parentPath & nextSubDir
            nextSubDir = Dir ' Dir is iterator ; -returns "" when no more left
        Loop
    End With
End Sub

Function NewTempDir()
    Dim sysTempFolder$: sysTempFolder = Environ("TEMP")
    Dim uniqueName$: uniqueName = "vbaAutoupdater_" & Format(Now, "yyyyMMdd_HHmmss")
    Dim newDirPath$: newDirPath = sysTempFolder & "\" & uniqueName
    
    On Error Resume Next
        VBA.FileSystem.MkDir newDirPath
        NewTempDir = IIf(Err.Number <> 0, "", newDirPath)
    On Error GoTo 0
End Function
