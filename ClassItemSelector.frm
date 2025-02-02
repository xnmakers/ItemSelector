VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ClassItemSelector 
   Caption         =   "UserForm1"
   ClientHeight    =   3040
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   4580
   OleObjectBlob   =   "ClassItemSelector.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ClassItemSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'XNMAKERS CODE STARTS HERE
Option Explicit

' ----------------------------------------------------------------
' ----------------------------------------------------------------
' ---------API Declarations updated for 64-bit compatibility -----
' ----------------------------------------------------------------
' ----------------------------------------------------------------
Private Declare PtrSafe Function OleTranslateColor Lib "oleaut32.dll" ( _
    ByVal lOleColor As Long, _
    ByVal hPalette As Long, _
    ByRef lRGBColor As Long) As Long

Private Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongA" ( _
    ByVal hwnd As LongPtr, ByVal nIndex As Long) As LongPtr

Private Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongA" ( _
    ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr

Private Declare PtrSafe Function FindWindowA Lib "user32" ( _
    ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr

Private Declare PtrSafe Function DrawMenuBar Lib "user32" ( _
    ByVal hwnd As LongPtr) As Long
' ----------------------------------------------------------------
' ----------------------------------------------------------------
' ---------API Declarations updated for 64-bit compatibility -----
' ----------------------------------------------------------------
' ----------------------------------------------------------------


' ----------------------------------------------------------------
' ----------------------------------------------------------------
' -------------------Form members and variables ------------------
' ----------------------------------------------------------------
' ----------------------------------------------------------------
Private Const GWL_STYLE As Long = -16
Private Const WS_THICKFRAME As Long = &H40000
Private Const WS_MINIMIZEBOX As Long = &H20000
Private Const WS_MAXIMIZEBOX As Long = &H10000

Private Const TEXTBOXHEIGHT As Long = 20
Private Const BUTTONHEIGHT As Long = 20
Private Const TITLEHEIGHT As Long = 13
Private Const MARGIN As Long = 10
Private Const GAP As Long = 5
Private Const FONTSIZE As Long = 10
Private Const FORM_BORDER_WIDTH As Long = 20 ' Approximate border width
Private Const FORM_TITLEBAR_HEIGHT As Long = 30 ' Approximate title bar height
Private Const MININSIDEWIDTH As Long = 150
Private Const MININSIDEHEIGHT As Long = 200
Private Const DESCAREAHEIGHT As Long = 30

Private WithEvents Tb_ClassKey As MSForms.TextBox
Attribute Tb_ClassKey.VB_VarHelpID = -1
Private WithEvents Tb_ClassValue As MSForms.TextBox
Attribute Tb_ClassValue.VB_VarHelpID = -1
Private WithEvents Tv_ClassTree As MSComctlLib.TreeView
Attribute Tv_ClassTree.VB_VarHelpID = -1
Private WithEvents Btn_Insert As MSForms.CommandButton
Attribute Btn_Insert.VB_VarHelpID = -1
Private WithEvents Btn_Close As MSForms.CommandButton
Attribute Btn_Close.VB_VarHelpID = -1
Private WithEvents Btn_Help As MSForms.CommandButton
Attribute Btn_Help.VB_VarHelpID = -1
Private WithEvents Btn_Update As MSForms.CommandButton
Attribute Btn_Update.VB_VarHelpID = -1
Private WithEvents Lb_SearchKey As MSForms.Label
Attribute Lb_SearchKey.VB_VarHelpID = -1
Private WithEvents Lb_SearchValue As MSForms.Label
Attribute Lb_SearchValue.VB_VarHelpID = -1
Private WithEvents Lb_Title As MSForms.Label
Attribute Lb_Title.VB_VarHelpID = -1
Private WithEvents Lb_Desc As MSForms.Label
Attribute Lb_Desc.VB_VarHelpID = -1

Private pKeyFilterStr As String
Private pValueFilterStr As String
Private pCurrentDataLib As DataLibrary
' ----------------------------------------------------------------
' ----------------------------------------------------------------
' -------------------Form members and variables ------------------
' ----------------------------------------------------------------
' ----------------------------------------------------------------


' ----------------------------------------------------------------
' ----------------------------------------------------------------
' ------------User Form and Control Initialization----------------
' ----------------------------------------------------------------
' ----------------------------------------------------------------
' --------------Initialize UserForm
Private Sub UserForm_Initialize()
    On Error GoTo ErrorHandler

    Me.Caption = Info.INFO_TITLE
    ' Enable resizing of the UserForm
    Dim hwnd As LongPtr
    hwnd = FindWindowA("ThunderDFrame", Me.Caption)

    If hwnd <> 0 Then
        Dim currentStyle As LongPtr
        currentStyle = GetWindowLongPtr(hwnd, GWL_STYLE)
        SetWindowLongPtr hwnd, GWL_STYLE, currentStyle Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX
        DrawMenuBar hwnd
    End If

    ' Set initial dimensions of the UserForm (client area)
    Me.Width = 350 + FORM_BORDER_WIDTH
    Me.Height = 300 + FORM_TITLEBAR_HEIGHT

    ' Add controls (TextBoxes, TreeView, Buttons, Labels)
    InitializeControls
    
    LoadAndBuildTreeView
    
    Exit Sub
ErrorHandler:
    Core.IsUserFormLoaded = False
    MsgBox "Error on UserForm_Initialize(): " & Err.Description, vbCritical
End Sub

' --------------Initialize Controls
Private Sub InitializeControls()

    ' Add label for the first TextBox inside the TextBox
    Set Lb_Title = Me.Controls.Add("Forms.Label.1", "Lb_Title", True)
    Lb_Title.Font.Size = FONTSIZE * 1.2 ' Adjust font size
    Lb_Title.Font.Name = "Calibri" ' Optional: Change font style
    Lb_Title.BackStyle = fmBackStyleTransparent
    Lb_Title.Font.Bold = True
    
    ' Add the first TextBox
    Set Tb_ClassKey = Me.Controls.Add("Forms.TextBox.1", "Tb_ClassKey", True)
    Tb_ClassKey.TextAlign = fmTextAlignLeft
    Tb_ClassKey.Font.Size = FONTSIZE ' Adjust font size
    Tb_ClassKey.Font.Name = "Calibri" ' Optional: Change font style
    
    ' Add label for the first TextBox inside the TextBox
    Set Lb_SearchKey = Me.Controls.Add("Forms.Label.1", "Lb_SearchKey", True)
    Lb_SearchKey.Font.Size = FONTSIZE ' Adjust font size
    Lb_SearchKey.Font.Name = "Calibri" ' Optional: Change font style
    Lb_SearchKey.BackStyle = fmBackStyleTransparent
    Lb_SearchKey.ForeColor = RGB(128, 128, 128) ' 50% Gray
    Lb_SearchKey.Enabled = False

    ' Add the second TextBox
    Set Tb_ClassValue = Me.Controls.Add("Forms.TextBox.1", "Tb_ClassValue", True)
    Tb_ClassValue.TextAlign = fmTextAlignLeft
    Tb_ClassValue.Font.Size = FONTSIZE ' Adjust font size
    Tb_ClassValue.Font.Name = "Calibri" ' Optional: Change font style

    ' Add label for the second TextBox inside the TextBox
    Set Lb_SearchValue = Me.Controls.Add("Forms.Label.1", "Lb_SearchValue", True)
    Lb_SearchValue.Font.Size = FONTSIZE ' Adjust font size
    Lb_SearchValue.Font.Name = "Calibri" ' Optional: Change font style
    Lb_SearchValue.BackStyle = fmBackStyleTransparent
    Lb_SearchValue.ForeColor = RGB(128, 128, 128) ' 50% Gray
    Lb_SearchValue.Enabled = False

    ' Add label for the description panel
    Set Lb_Desc = Me.Controls.Add("Forms.Label.1", "Lb_Desc", True)
    Lb_Desc.Font.Size = FONTSIZE * 0.9 ' Adjust font size
    Lb_Desc.Font.Name = "Calibri" ' Optional: Change font style
    Lb_Desc.Caption = "Description: "
    Lb_Desc.BackStyle = fmBackStyleTransparent
    Lb_Desc.Enabled = True

    ' Add the TreeView control
    Set Tv_ClassTree = Me.Controls.Add("MSComctlLib.TreeCtrl.2", "Tv_ClassTree", True)
    Tv_ClassTree.LabelEdit = False

    ' Add the "Insert" button
    Set Btn_Insert = Me.Controls.Add("Forms.CommandButton.1", "Btn_Insert", True)
    Btn_Insert.Caption = "Insert"
    Btn_Insert.BackColor = RGB(245, 245, 245) ' 50% Gray

    ' Add the "Close" button
    Set Btn_Close = Me.Controls.Add("Forms.CommandButton.1", "Btn_Close", True)
    Btn_Close.Caption = "Close"
    Btn_Close.BackColor = RGB(245, 245, 245) ' 50% Gray
    
    ' Add the "Help" button
    Set Btn_Help = Me.Controls.Add("Forms.CommandButton.1", "Btn_Help", True)
    Btn_Help.Caption = "?"
    Btn_Help.Font.Bold = True
    Btn_Help.ForeColor = RGB(255, 255, 255)
    Btn_Help.BackColor = RGB(180, 180, 180) ' 50% Gray

    ' Add the "Update" button
    Set Btn_Update = Me.Controls.Add("Forms.CommandButton.1", "Btn_Update", True)
    Btn_Update.Caption = "U"
    Btn_Update.Font.Bold = True
    Btn_Update.ForeColor = RGB(255, 255, 255)
    Btn_Update.BackColor = RGB(144, 238, 144) ' Light Green
    
    SetControlSize
End Sub

' --------------Set all the control size, position, and visibility based on the form size
Private Sub SetControlSize()

    Dim minInsideW As Long
    Dim minInsideH As Long
    
    If Me.InsideWidth < MININSIDEWIDTH Then
        minInsideW = MININSIDEWIDTH
    Else
        minInsideW = Me.InsideWidth
    End If
    If Me.InsideHeight < MININSIDEHEIGHT Then
        minInsideH = MININSIDEHEIGHT
    Else
        minInsideH = Me.InsideHeight
    End If
    
    If Not Lb_Title Is Nothing Then
        Lb_Title.Top = MARGIN
        Lb_Title.Left = MARGIN
        Lb_Title.Height = TITLEHEIGHT
        Lb_Title.Width = minInsideW - MARGIN * 2
    End If
    
    If Not Tb_ClassKey Is Nothing Then
        Tb_ClassKey.Top = MARGIN + Lb_Title.Height + GAP
        Tb_ClassKey.Left = MARGIN
        Tb_ClassKey.Width = (minInsideW - MARGIN * 2 - GAP) / 2
        Tb_ClassKey.Height = TEXTBOXHEIGHT
    End If

    If Not Tb_ClassValue Is Nothing Then
        Tb_ClassValue.Top = MARGIN + Lb_Title.Height + GAP
        Tb_ClassValue.Left = Tb_ClassKey.Left + Tb_ClassKey.Width + GAP
        Tb_ClassValue.Width = (minInsideW - MARGIN * 2 - GAP) / 2
        Tb_ClassValue.Height = TEXTBOXHEIGHT
    End If
    
    If Not Lb_SearchKey Is Nothing Then
        Lb_SearchKey.Top = Tb_ClassKey.Top + 3
        Lb_SearchKey.Left = Tb_ClassKey.Left + 9
        Lb_SearchKey.Width = Tb_ClassKey.Width - 12
        Lb_SearchKey.Height = Tb_ClassKey.Height - 4
    End If
    
    If Not Lb_SearchValue Is Nothing Then
        Lb_SearchValue.Top = Tb_ClassValue.Top + 3
        Lb_SearchValue.Left = Tb_ClassValue.Left + 9
        Lb_SearchValue.Width = Tb_ClassValue.Width - 12
        Lb_SearchValue.Height = Tb_ClassValue.Height - 4
    End If

    If Not Tv_ClassTree Is Nothing Then
        Tv_ClassTree.Top = Tb_ClassKey.Top + Tb_ClassKey.Height + GAP
        Tv_ClassTree.Left = MARGIN
        Tv_ClassTree.Width = minInsideW - MARGIN * 2
        Tv_ClassTree.Height = minInsideH - MARGIN * 2 - GAP * 4 - TEXTBOXHEIGHT - BUTTONHEIGHT - TITLEHEIGHT - DESCAREAHEIGHT
    End If

    If Not Lb_Desc Is Nothing Then
        Lb_Desc.Top = Tv_ClassTree.Top + Tv_ClassTree.Height + GAP
        Lb_Desc.Left = MARGIN
        Lb_Desc.Width = minInsideW - MARGIN * 2
        Lb_Desc.Height = DESCAREAHEIGHT
    End If

    If Not Btn_Insert Is Nothing Then
        Btn_Insert.Height = BUTTONHEIGHT
        Btn_Insert.Width = (minInsideW - MARGIN * 2 - GAP * 3 - BUTTONHEIGHT * 2) * 3 / 4
        Btn_Insert.Top = minInsideH - Btn_Insert.Height - MARGIN
        Btn_Insert.Left = MARGIN
    End If

    If Not Btn_Close Is Nothing Then
        Btn_Close.Height = BUTTONHEIGHT
        Btn_Close.Width = (minInsideW - MARGIN * 2 - GAP * 3 - BUTTONHEIGHT * 2) * 1 / 4
        Btn_Close.Top = minInsideH - Btn_Close.Height - MARGIN
        Btn_Close.Left = Btn_Insert.Left + Btn_Insert.Width + GAP
    End If
    
    If Not Btn_Help Is Nothing Then
        Btn_Help.Height = BUTTONHEIGHT
        Btn_Help.Width = BUTTONHEIGHT
        Btn_Help.Top = minInsideH - Btn_Help.Height - MARGIN
        Btn_Help.Left = Btn_Close.Left + Btn_Close.Width + GAP
    End If

    If Not Btn_Update Is Nothing Then
        Btn_Update.Height = BUTTONHEIGHT
        Btn_Update.Width = BUTTONHEIGHT
        Btn_Update.Top = minInsideH - Btn_Help.Height - MARGIN
        Btn_Update.Left = Btn_Help.Left + Btn_Help.Width + GAP
    End If
End Sub

' --------------Resize User Form
Private Sub UserForm_Resize()
    ' Dynamically adjust the layout when resizing
    If Me.Controls.Count > 0 Then
        SetControlSize
    End If
End Sub

' --------------Terminate UserForm, Fired when Form is Unloaded
Private Sub UserForm_Terminate()
    Core.IsUserFormLoaded = False
End Sub
' ----------------------------------------------------------------
' ----------------------------------------------------------------
' ------------User Form and Control Initialization----------------
' ----------------------------------------------------------------
' ----------------------------------------------------------------


' ----------------------------------------------------------------
' ----------------------------------------------------------------
' --------------------Load and Build Tree View--------------------
' ----------------------------------------------------------------
' ----------------------------------------------------------------
Private Sub LoadAndBuildTreeView()
    On Error GoTo ErrorHandler
    
    Set Core.CurrentForm = Me
    Dim key As Variant
    
    pKeyFilterStr = Trim(Tb_ClassKey.Text)
    pValueFilterStr = Trim(Tb_ClassValue.Text)
    
    Core.LoadDataLibraries
    ' Load all the tree data first
    For Each key In Core.DataLibDict.Keys
        Core.DataLibDict(key).LoadNodeList
    Next key
    
    Set pCurrentDataLib = Core.GetDataLibrary(ActiveCell(1, 1))
    Me.BackColor = Core.CurrentColumnColor
    Lb_Desc.ForeColor = CalcForeColor(Me.BackColor)
    Lb_Title.ForeColor = CalcForeColor(Me.BackColor)

    If pCurrentDataLib Is Nothing Then
        If Core.DataLibDict.Count = 0 Then
            MsgBox "No data library has loaded.", vbExclamation
            End
        End If
        Set pCurrentDataLib = Core.DataLibDict.Items()(0)
        Me.BackColor = &H8000000B
    End If
    pCurrentDataLib.BuildTreeView Tv_ClassTree, pKeyFilterStr, pValueFilterStr
    
    ' Initialize the title of the form
    If Not Lb_Title Is Nothing Then
        Lb_Title.Caption = pCurrentDataLib.LibraryName & " CLASS LIBRARY: " & pCurrentDataLib.NodeCount & " Nodes"
        Lb_SearchKey.Caption = pCurrentDataLib.KeyHdg
        Lb_SearchValue.Caption = pCurrentDataLib.ValueHdg
    End If
    
    Exit Sub
ErrorHandler:
    Core.IsUserFormLoaded = False
    MsgBox "Error on ClassItemSelector.LoadAndBuildTreeView(): " & Err.Description, vbCritical
    End
End Sub
' ----------------------------------------------------------------
' ----------------------------------------------------------------
' --------------------Load and Build Tree View--------------------
' ----------------------------------------------------------------
' ----------------------------------------------------------------




' ----------------------------------------------------------------
' ----------------------------------------------------------------
' --------------------Control Event Handlers ---------------------
' ----------------------------------------------------------------
' ----------------------------------------------------------------
' --------------TextBox Class Name Changed
Private Sub Tb_ClassKey_Change()
    Lb_SearchKey.Visible = (Tb_ClassKey.Text = "")
    pKeyFilterStr = Trim(Tb_ClassKey.Text)

    If pCurrentDataLib.NodeCount < 700 Then
        pCurrentDataLib.BuildTreeView Tv_ClassTree, pKeyFilterStr, pValueFilterStr
    Else
        Core.StartDebounceTimer
    End If
End Sub

' --------------TextBox Class Code Changed
Private Sub Tb_ClassValue_Change()
    Lb_SearchValue.Visible = (Tb_ClassValue.Text = "")
    pValueFilterStr = Trim(Tb_ClassValue.Text)

    If pCurrentDataLib.NodeCount < 700 Then
        pCurrentDataLib.BuildTreeView Tv_ClassTree, pKeyFilterStr, pValueFilterStr
    Else
        Core.StartDebounceTimer
    End If
End Sub
' --------------Insert Button is clicked
Private Sub Btn_Insert_Click()
    ' Insert the selected TreeView item's text into the active cell
    If Not Tv_ClassTree.SelectedItem Is Nothing Then
        
        ' Insert into the active cell
        'On Error Resume Next
        Core.AssignFromNode ActiveCell.Cells(1, 1), Tv_ClassTree.SelectedItem.Tag
        'On Error GoTo 0
    Else
        MsgBox "Please select an item from the Tree View.", vbExclamation
    End If
End Sub

' --------------Close Button is clicked
Private Sub Btn_Close_Click()
    ' Close the form without action
    Unload Me
    Core.IsUserFormLoaded = False
End Sub

' --------------Help Button is clicked
Private Sub Btn_Help_Click()
    MsgBox Info.INFO_DESCRIPTION & vbCrLf & Info.INFO_KNOWNISSUE & vbCrLf & Info.INFO_CONTACT, vbInformation, "Item Selector Help"
End Sub

' --------------Update Button is clicked
Private Sub Btn_Update_Click()
    MsgBox "Update function is not available yet", vbInformation
End Sub

' --------------TreeView Node Selection Changed
Private Sub Tv_ClassTree_NodeClick(ByVal node As MSComctlLib.node)
    If Not node Is Nothing Then
        Lb_Desc.Caption = "Description: " & node.Tag.ClassDesc
    End If
End Sub

' --------------On Selection Changed
Public Sub OnClassSelectionChanged(currentCell As Range)
        
    If Core.IsUserFormLoaded Then
        Dim dataLib As DataLibrary
        Set dataLib = Core.GetDataLibrary(currentCell(1, 1))
        
        Me.BackColor = Core.CurrentColumnColor
        Lb_Desc.Caption = "Description: "
        Lb_Desc.ForeColor = CalcForeColor(Me.BackColor)
        Lb_Title.ForeColor = CalcForeColor(Me.BackColor)

        If Not dataLib Is Nothing Then
            If pCurrentDataLib.LibraryName <> dataLib.LibraryName Then
                Set pCurrentDataLib = dataLib
                pCurrentDataLib.BuildTreeView Tv_ClassTree, pKeyFilterStr, pValueFilterStr

                ' Change title name as the class library changes
                If Not Lb_Title Is Nothing Then
                    Lb_Title.Caption = pCurrentDataLib.LibraryName + " CLASS LIBRARY: " & pCurrentDataLib.NodeCount & " Nodes"
                    Lb_SearchKey.Caption = pCurrentDataLib.KeyHdg
                    Lb_SearchValue.Caption = pCurrentDataLib.ValueHdg
                End If
            End If
        End If
    End If
End Sub

' ----------------------------------------------------------------
' ----------------------------------------------------------------
' --------------------Control Event Handlers ---------------------
' ----------------------------------------------------------------
' ----------------------------------------------------------------


' ----------------------------------------------------------------
' ----------------------------------------------------------------
' --------------------Control Event Handlers ---------------------
' ----------------------------------------------------------------
' ----------------------------------------------------------------
' --------------Calculate ForeColor based on BackColor
Public Function CalcForeColor(BackColor As Long) As Long
    Dim lRGBColor As Long
    OleTranslateColor BackColor, 0, lRGBColor
    If lRGBColor < &H800000 Then
        CalcForeColor = &HFFFFFF
    Else
        CalcForeColor = &H0
    End If
End Function

' --------------Perform Debounce
Public Sub PerformDebounce()
    pCurrentDataLib.BuildTreeView Tv_ClassTree, pKeyFilterStr, pValueFilterStr
End Sub

' ----------------------------------------------------------------
' ----------------------------------------------------------------
' --------------------Control Event Handlers ---------------------
' ----------------------------------------------------------------
' ----------------------------------------------------------------





