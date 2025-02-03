Attribute VB_Name = "Installer"

'XNMAKERS CODE STARTS HERE
Public Sub Install()
    Dim httpRequest As Object
    Dim linkStr As String
    Dim installUrl As String
    
    ' Configuration
    installUrl = "https://raw.githubusercontent.com/xnmakers/ItemSelector/refs/heads/main/INSTALL"
    
    
    ' Create HTTP request with authentication header
    Set httpRequest = CreateObject("WinHttp.WinHttpRequest.5.1")
    httpRequest.Open "GET", installUrl, False
    httpRequest.SetRequestHeader "User-Agent", "Mozilla/5.0"
    httpRequest.SetRequestHeader "Accept", "*/*"    
    httpRequest.Send
    
    ' Check for errors
    If httpRequest.Status <> 200 Then
        MsgBox "Failed to download. Status: " & httpRequest.Status & vbCrLf & _
               "You may need to check the URL or your internet connection.", vbExclamation
        Exit Sub
    End If
    
    
    linkStr = httpRequest.responseText

    Dim linkList() As String
    Dim i As Integer    
    Dim fileContentDict As Object
    Set fileContentDict = CreateObject("Scripting.Dictionary")

    ' Split the response text by lines
    linkList = Split(linkStr, vbLf)
    
    ' Loop through each line (file link)
    For i = LBound(linkList) To UBound(linkList)
        If Trim(linkList(i)) <> "" Then
            Dim fileRequest As Object
            
            ' Create a new HTTP request for each file link
            Set fileRequest = CreateObject("WinHttp.WinHttpRequest.5.1")
            fileRequest.Open "GET", Trim(linkList(i)), False
            fileRequest.SetRequestHeader "User-Agent", "Mozilla/5.0"
            fileRequest.SetRequestHeader "Accept", "*/*"
            fileRequest.Send
            ' Extract the file name from the URL
            Dim fName As String

            fName = Mid(linkList(i), InStrRev(linkList(i), "/") + 1)
            If fName = "" Then
                MsgBox "Invalid file name detected. Ending program.", vbExclamation
                End
            End If

            ' Check for errors
            If fileRequest.Status <> 200 Then
                MsgBox "Failed to install the program. Status: " & fileRequest.Status & vbCrLf & _
                       "Failed to download file: " & linkList(i), vbExclamation
                End
            End If            
            
            ' Extract the content starting from "'XNMAKERS CODE STARTS HERE"
            Dim content As String
            content = fileRequest.responseText
            Dim startPos As Long
            startPos = InStr(content, "'XNMAKERS CODE STARTS HERE")
            
            If startPos > 0 Then
                content = Mid(content, startPos)
            Else
                MsgBox "The marker 'XNMAKERS CODE STARTS HERE' was not found in the file: " & fName, vbExclamation
                End
            End If
            
            ' Add the file content to the dictionary
            fileContentDict(fName) = content
        End If
    Next i
    Dim vbaProject As Object
    Dim vbComponent As Object
    Dim fileName As Variant
    
    ' Get the VBA project
    Set vbaProject = ThisWorkbook.VBProject
    
    Dim vbaMacroDict As Object
    Set vbaMacroDict = CreateObject("Scripting.Dictionary")
    
    ' Iterate through all VBA components in the project
    For Each vbComponent In vbaProject.VBComponents
        Dim componentName As String
        Dim componentCode As String
        componentName = vbComponent.Name
        
        Select Case vbComponent.Type
            Case 1 ' vbext_ct_StdModule
                componentName = componentName & ".bas"
            Case 2 ' vbext_ct_ClassModule
                componentName = componentName & ".cls"
            Case 3 ' vbext_ct_MSForm
                componentName = componentName & ".frm"
        End Select
        
        ' Add the component name with extension to the dictionary
        vbaMacroDict(componentName) = vbComponent.Name
    Next vbComponent

    ' Iterate through the dictionary and create modules
    For Each fileName In fileContentDict.Keys
        Dim fileExtension As String
        fileExtension = Mid(fileName, InStrRev(fileName, ".") + 1)        
        
        ' Check if the component already exists in the dictionary
        If vbaMacroDict.Exists(fileName) Then
            MsgBox "The file " & fileName & " already exists in the project. Installation Process is Canceled.", vbInformation
            End
        End If

        Select Case LCase(fileExtension)
            Case "bas"
                ' Add a new standard module
                Set vbComponent = vbaProject.VBComponents.Add(1) ' 1 = vbext_ct_StdModule
            Case "cls"
                ' Add a new class module
                Set vbComponent = vbaProject.VBComponents.Add(2) ' 2 = vbext_ct_ClassModule
            Case "frm"
                ' Add a new user form
                Set vbComponent = vbaProject.VBComponents.Add(3) ' 3 = vbext_ct_MSForm
            Case Else
                MsgBox "Unknown file extension: " & fileExtension, vbExclamation
                End
        End Select        
        
        Dim macroComponentName As String
        macroComponentName = Left(fileName, InStrRev(fileName, ".") - 1)

        ' Assign the name to the component
        vbComponent.Name = macroComponentName
    NextFile:
        ' Add the content to the module
        vbComponent.CodeModule.AddFromString fileContentDict(fileName)
    Next fileName

    ' Add your code here to further process the fileContentList collection
    MsgBox "All files are downloaded and processed successfully."
End Sub