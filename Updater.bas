Public Sub Update()
    Dim httpRequest As Object
    Dim linkStr As String
    Dim installUrl As String
    Dim currentVersion As Double
    Dim newVersion As Double
    
    Dim vbaProject As Object
    Set vbaProject = ThisWorkbook.VBProject

    currentVersion = 0#
    newVersion = 0# ' Initialize newVersion to a default value
    ' Configuration: URL to your update file
    installUrl = "https://raw.githubusercontent.com/xnmakers/ItemSelector/refs/heads/main/INSTALL"

    Dim vbComponent As Object
    For Each vbComponent In vbaProject.VBComponents
        If vbComponent.Name = "Info" Then
            Dim codeContent As String
            Dim j As Long
            For j = 1 To vbComponent.codeModule.CountOfLines
                codeContent = codeContent & vbComponent.codeModule.lines(j, 1) & vbCrLf
            Next j
            currentVersion = GetVersionNumber(codeContent)
            Exit For
        End If
    Next vbComponent
    
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
    
    Dim k As Long
    For k = LBound(linkList) To UBound(linkList)
        If Trim(linkList(k)) = "" Then Exit For
        
        Dim fileUrl As String
        fileUrl = Trim(linkList(k))
        
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
    Next k

    Dim filesToDelete() As String
    ' Extract version number from Info.bas if it exists in the downloaded files
    If fileContentDict.Exists("Info.bas") Then
        Dim infoContent As String
        infoContent = fileContentDict("Info.bas")
        newVersion = GetVersionNumber(infoContent)
        filesToDelete = GetFileNames(infoContent)
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
    
    Dim fileToDelete As Variant
    For Each fileToDelete In filesToDelete

        Dim componentName As String
        Dim currentvbComp As String
        componentName = Left(Trim(fileToDelete), InStrRev(Trim(fileToDelete), ".") - 1)

        For Each vbComponent In vbaProject.VBComponents
            currentvbComp = vbComponent.Name
            If componentName <> "Updater" And currentvbComp = componentName Then
                vbaProject.VBComponents.Remove vbaProject.VBComponents(componentName)
            End If
        Next vbComponent
        
        On Error GoTo 0
        
    Next fileToDelete

    ' Schedule the import of the new modules to run shortly after this sub ends.
    ' (A short delay lets VBA complete the deletions.)
    Application.OnTime Now + TimeValue("00:00:01"), "ImportModules"
    
    MsgBox "Old modules deleted. New modules will be imported shortly.", vbInformation
End Sub

Private Sub ImportModules()
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

' Default value is 0
Private Function GetVersionNumber(codeModuleContent As String) As Double
    Dim currentVersion As Double
    currentVersion = 0#

    Dim lines() As String
    Dim line As Variant
    Dim versionString As String
    Dim verPos As Long
    
    ' Split the input string into lines
    lines = Split(codeModuleContent, vbCrLf)
    
    ' Loop through each line to find the version number
    For Each line In lines
        verPos = InStr(line, "Public Const INFO_VERSION As Double =")
        
        If verPos > 0 Then
            versionString = Trim(Mid(line, verPos + Len("Public Const INFO_VERSION As Double =")))
            currentVersion = CDbl(versionString)
            Exit For
        End If
    Next line

    GetVersionNumber = currentVersion
End Function

Private Function GetFileNames(codeModule As String) As String()
    Dim lines() As String
    Dim i As Long
    Dim fileNames() As String
    
    ' Default value for fileNames
    fileNames = Split("Core.bas" & vbCrLf & _
                      "Info.bas" & vbCrLf & _
                      "ClassItemSelector.frm" & vbCrLf & _
                      "Assigner.cls" & vbCrLf & _
                      "ClassNode.cls" & vbCrLf & _
                      "DataLibrary.cls", vbCrLf)
    
    ' Split the input string into lines
    lines = Split(codeModule, vbCrLf)
    
    ' Loop through the lines to find the one after 'Files to be Updated
    For i = LBound(lines) To UBound(lines)
        If InStr(lines(i), "'Files to be Updated") > 0 Then
            ' The file names are on the next line
            fileNames = Split(Mid(Trim(lines(i + 1)), 2), ";")
            Exit For
        End If
    Next i
    
    ' Trim each file name
    For i = LBound(fileNames) To UBound(fileNames)
        fileNames(i) = Trim(fileNames(i))
    Next i
    
    GetFileNames = fileNames
End Function

