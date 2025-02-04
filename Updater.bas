Attribute VB_Name = "Updater"
Public Sub Update()
    Dim httpRequest As Object
    Dim linkStr As String
    Dim installUrl As String
    Dim currentVersion As Double

    ' Get current version from your Info module or constant
    On Error Resume Next
    currentVersion = Info.INFO_VERSION
    If Err.Number <> 0 Then
        currentVersion = 0.0
    End If
    On Error GoTo 0

    ' Configuration: URL to your update file
    installUrl = "https://raw.githubusercontent.com/xnmakers/ItemSelector/refs/heads/main/INSTALL"
    
    ' Create HTTP request to get the update file
    Set httpRequest = CreateObject("WinHttp.WinHttpRequest.5.1")
    httpRequest.Open "GET", installUrl, False
    httpRequest.SetRequestHeader "User-Agent", "Mozilla/5.0"
    httpRequest.Send

    If httpRequest.Status <> 200 Then
        MsgBox "Failed to check updates. Status: " & httpRequest.Status, vbExclamation
        Exit Sub
    End If

    linkStr = httpRequest.responseText
    Dim linkList() As String
    linkList = Split(linkStr, vbLf)

    ' Download the updated modules and store them in a dictionary
    Dim fileContentDict As Object
    Set fileContentDict = CreateObject("Scripting.Dictionary")
    
    Dim i As Long
    For i = LBound(linkList) To UBound(linkList)
        If Trim(linkList(i)) = "" Then Exit For
        
        Dim fileUrl As String
        fileUrl = Trim(linkList(i))
        
        Set httpRequest = CreateObject("WinHttp.WinHttpRequest.5.1")
        httpRequest.Open "GET", fileUrl, False
        httpRequest.Send
        
        If httpRequest.Status <> 200 Then
            MsgBox "Failed to download: " & fileUrl, vbExclamation
            Exit Sub
        End If
        
        Dim content As String
        content = httpRequest.responseText
        
        ' Extract filename from URL (everything after the last "/")
        Dim fileName As String
        fileName = Mid(fileUrl, InStrRev(fileUrl, "/") + 1)
        fileContentDict(fileName) = content
    Next i


    Dim newVersion As Double
    newVersion = 0 ' Initialize newVersion to a default value

    ' Extract version number from Info.bas if it exists in the downloaded files
    If fileContentDict.Exists("Info.bas") Then
        Dim infoContent As String
        infoContent = fileContentDict("Info.bas")
        
        Dim versionPattern As String
        versionPattern = "Public Const INFO_VERSION As Double = "
        
        Dim versionPos As Long
        versionPos = InStr(infoContent, versionPattern)
        
        If versionPos > 0 Then
            Dim versionStart As Long
            versionStart = versionPos + Len(versionPattern)
            
            Dim versionEnd As Long
            versionEnd = InStr(versionStart, infoContent, vbCrLf)
            
            If versionEnd > 0 Then
                Dim extractedVersion As String
                extractedVersion = Mid(infoContent, versionStart, versionEnd - versionStart)
                
                ' Attempt to convert the extracted version to a Double
                On Error Resume Next
                newVersion = CDbl(Trim(extractedVersion))
                On Error GoTo 0
            End If
        End If
    End If
    
    ' Check if the new version is valid and greater than the current version
    If newVersion > 0 And newVersion <= currentVersion Then
        MsgBox "Already up to date.", vbInformation
        Exit Sub
    End If

    ' Create a temporary folder in the Temp directory
    Dim tempFolder As String
    tempFolder = Environ("Temp") & "\VBAUpdate\"
    If Dir(tempFolder, vbDirectory) = "" Then
        MkDir tempFolder
    End If

    ' Write each downloaded module to a temporary file
    Dim key As Variant
    For Each key In fileContentDict.Keys
        Dim fullPath As String
        fullPath = tempFolder & key
        
        Dim fileNum As Integer
        fileNum = FreeFile
        Open fullPath For Output As #fileNum
        Print #fileNum, fileContentDict(key)
        Close #fileNum
    Next key

    ' Delete old modules (except the Updater module itself)
    Dim vbaProject As Object
    Set vbaProject = ThisWorkbook.VBProject
    Dim filesToDelete() As String

    On Error Resume Next
    filesToDelete = Split(Info.INFO_FILES_LIST, vbCrLf)
    If Err.Number <> 0 Then
        filesToDelete = Split("Core.bas" & vbCrLf & _
                              "Info.bas" & vbCrLf & _
                              "ClassItemSelector.frm" & vbCrLf & _
                              "Assigner.cls" & vbCrLf & _
                              "ClassNode.cls" & vbCrLf & _
                              "DataLibrary.cls", vbCrLf)
    End If
    On Error GoTo 0    
    
    Dim fileToDelete As Variant
    For Each fileToDelete In filesToDelete
        On Error Resume Next
        Dim componentName As String
        componentName = Left(Trim(fileToDelete), InStrRev(Trim(fileToDelete), ".") - 1)
        ' Skip deletion of the updater module to avoid interrupting the running code.

        If vbaProject.VBComponents.Count = 0 Or vbaProject.VBComponents(componentName) Is Nothing Then
            GoTo SkipDeletion
        End If

        If componentName <> "Updater" Then
            vbaProject.VBComponents.Remove vbaProject.VBComponents(componentName)
        End If
        On Error GoTo 0
        
    SkipDeletion:
    Next fileToDelete

    ' Schedule the import of the new modules to run shortly after this sub ends.
    ' (A short delay lets VBA complete the deletions.)
    Application.OnTime Now + TimeValue("00:00:01"), "ImportModules"
    
    MsgBox "Old modules deleted. New modules will be imported shortly.", vbInformation
End Sub

Public Sub ImportModules()
    Dim tempFolder As String
    tempFolder = Environ("Temp") & "\VBAUpdate\"
    
    Dim fileName As String
    ' Loop through all files in the temporary folder
    fileName = Dir(tempFolder & "*.*")
    Do While fileName <> ""
        Dim fileExtension As String
        ' Get the extension (in lower case) from the file name
        fileExtension = LCase$(Mid(fileName, InStrRev(fileName, ".") + 1))
        
        ' Check if the file is one of the recognized VBA component types.
        If fileExtension = "bas" Or fileExtension = "cls" Or fileExtension = "frm" Then
            ThisWorkbook.VBProject.VBComponents.Import tempFolder & fileName
        End If
        
        ' Remove the temporary file after processing
        Kill tempFolder & fileName
        
        ' Get the next file in the folder
        fileName = Dir()
    Loop
    
    ' Attempt to remove the temporary folder
    On Error Resume Next
    RmDir tempFolder
    On Error GoTo 0

    MsgBox "Update successful! Please restart Excel to use the new version.", vbInformation
End Sub