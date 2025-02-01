Attribute VB_Name = "Core"
'XNMAKERS CODE STARTS HERE
Option Private Module
Option Explicit
' ----------------------------------------------------------------
' ----------------------------------------------------------------
' -------------------------- Authos's Note -----------------------
' ----------------------------------------------------------------
' ----------------------------------------------------------------
Private Const INFO_VERSION_MAJOR As String = "1"
Private Const INFO_VERSION_MINOR As String = "1"
Private Const INFO_AUTHOR As String = "Xingyun Jin"
Private Const INFO_EMAIL As String = "xingyun.jin@avangrid.com"
Private Const INFO_DATE As String = "01/28/2025"

Public Const INFO_TITLE = "Item Selector - v" & INFO_VERSION_MAJOR & "." & INFO_VERSION_MINOR
Public Const INFO_DESCRIPTION = "The Macro UI is to generate Tree Structure using the specified data." & _
                                "The Macro UI supports auto-fill of the items to the pairing cells, based on the data library that is mapped from Worksheet, [ClassDataMapping] ."
Public Const INFO_KNOWNISSUE = vbCrLf & "Known Issues:" & vbCrLf & _
                                "1. The Macro clears the Undo and Redo (ctrl+z, ctrl+y) stack by design" & vbCrLf & _
                                "2. Auto-fill of associating cell functions one at a time. Mass Auto-fill is not supported at the moment" & vbCrLf & _
                                "3. The Macro is not tested on Mac OS" & vbCrLf & _
                                "4. The Macro is not tested on Excel 2010 or earlier versions" & vbCrLf & _
                                "5. The Macro is not tested on Excel Online" & vbCrLf & _
                                "6. The Macro is not tested on Excel Mobile"

Public Const INFO_CONTACT = vbCrLf & "Please contact the Author, " & INFO_AUTHOR & " for more information " & vbCrLf & _
                            INFO_EMAIL

Private Const DATAENTRY_WORKSHEET As String = "ClassDataMapping"
Private Const DATAENTRY_TableFrom As String = "tbl_ClassDataMapping_From"
Private Const DATAENTRY_TableTo As String = "tbl_ClassDataMapping_To"

Public Enum ColumnProperty
    KeyColumn
    ValueColumn
    NotAssigned
End Enum
' ----------------------------------------------------------------
' ----------------------------------------------------------------
' -------------------------- Authos's Note -----------------------
' ----------------------------------------------------------------
' ----------------------------------------------------------------



' ----------------------------------------------------------------
' ----------------------------------------------------------------
' -------------------------- Timer : API Declarations ------------
' ----------------------------------------------------------------
' ----------------------------------------------------------------
Public Declare PtrSafe Function SetTimer Lib "user32" ( _
    ByVal hwnd As LongPtr, _
    ByVal nIDEvent As LongPtr, _
    ByVal uElapse As Long, _
    ByVal lpTimerFunc As LongPtr) As LongPtr

Public Declare PtrSafe Function KillTimer Lib "user32" ( _
    ByVal hwnd As LongPtr, _
    ByVal nIDEvent As LongPtr) As Long

Public TimerId As LongPtr
Public CurrentForm As ClassItemSelector
Private Const DEBOUCETIMER As Long = 300
' ----------------------------------------------------------------
' ----------------------------------------------------------------
' -------------------------- Timer : API Declarations ------------
' ----------------------------------------------------------------
' ----------------------------------------------------------------



' ----------------------------------------------------------------
' ----------------------------------------------------------------
' -------------------------- Members and Properties -------------
' ----------------------------------------------------------------
' ----------------------------------------------------------------
' Private Variables
Private pDataLibDict As Object
Private pAssignerList As Collection
Private pIsUserFormLoaded As Boolean
Private pCurrentColumnColor As Long
Private pCurrentAssigner As Assigner
' Public Properties
Public Property Get IsUserFormLoaded() As Boolean
    IsUserFormLoaded = pIsUserFormLoaded
End Property
Public Property Let IsUserFormLoaded(value As Boolean)
    pIsUserFormLoaded = value
End Property
Public Property Get DataLibDict() As Object
    Set DataLibDict = pDataLibDict
End Property
Public Property Get CurrentColumnColor() As Long
    CurrentColumnColor = pCurrentColumnColor
End Property
Public Property Get CurrentAssigner() As Assigner
    Set CurrentAssigner = pCurrentAssigner
End Property
' ----------------------------------------------------------------
' ----------------------------------------------------------------
' -------------------------- Members and Properties --------------
' ----------------------------------------------------------------
' ----------------------------------------------------------------



' ----------------------------------------------------------------
' ----------------------------------------------------------------
' -------------------------- Data Library and Data Assign --------
' ----------------------------------------------------------------
' ----------------------------------------------------------------

' --------------Load Data Libraries
Public Sub LoadDataLibraries()

    ' Check if the UserForm is already loaded, if user form is loaded, that means the data is already loaded
    If pIsUserFormLoaded Then
        Exit Sub
    End If

    Dim rngFrom As Range
    Dim wkst As Worksheet
    Dim rngTo As Range
    Dim row As Range
    
    Set pDataLibDict = CreateObject("Scripting.Dictionary")
    Set pAssignerList = New Collection

    On Error Resume Next
    Set wkst = ThisWorkbook.Sheets(DATAENTRY_WORKSHEET)    

    If wkst Is Nothing Then
        MsgBox "Worksheet '" & DATAENTRY_WORKSHEET & "' is not found! Please provide 'ClassDataMapping' worksheet.", vbExclamation, "Class Data Mapping Warning"
        End
    End If

    Set rngFrom = wkst.ListObjects(DATAENTRY_TableFrom).DataBodyRange
    If rngFrom Is Nothing Then
        MsgBox "Table '" & DATAENTRY_TableFrom & "' is not found! Please provide 'tbl_ClassDataMapping_From' table and define data libraries.", vbExclamation, "Class Data Mapping Warning"
        End
    End If

    Set rngTo = wkst.ListObjects(DATAENTRY_TableTo).DataBodyRange
        If rngTo Is Nothing Then
        MsgBox "Table '" & DATAENTRY_TableTo & "' is not found! Please provide 'tbl_ClassDataMapping_To' table and define data input columns.", vbExclamation, "Class Data Mapping Warning"
        End
    End If

    On Error GoTo ErrorHandler

    For Each row In rngFrom.Rows
        If Not pDataLibDict.Exists(row.Cells(1, 1).value) Then
            Dim data As DataLibrary
            Set data = New DataLibrary

            If IsError(row.Cells(1, 1).Value) Or IsError(row.Cells(1, 2).Value) Or _
               IsError(row.Cells(1, 3).Value) Or IsError(row.Cells(1, 4).Value) Or _
               IsError(row.Cells(1, 5).Value) Or IsError(row.Cells(1, 6).Value) Then
                MsgBox "Error in data at row " & row.Row & " at table: '" & DATAENTRY_TableFrom & "'. Please check the data.", vbExclamation, "Data Error in Defining Data Library"
                End
            End If
            
            data.Initialize row.Cells(1, 1).Value, row.Cells(1, 2).Value, _
                            row.Cells(1, 3).Value, row.Cells(1, 4).Value, _
                            row.Cells(1, 5).Value, row.Cells(1, 6).Value
            pDataLibDict.Add row.Cells(1, 1).Value, data
        End If
    Next row
    
    For Each row In rngTo.Rows
        Dim ass As Assigner
        Set ass = New Assigner
        If IsError(row.Cells(1, 1).Value) Or IsError(row.Cells(1, 2).Value) Or _
           IsError(row.Cells(1, 3).Value) Or IsError(row.Cells(1, 4).Value) Then
            MsgBox "Error in data at row " & row.Row & " at table: '" & DATAENTRY_TableTo & "'. Please check the data.", vbExclamation, "Data Error in Defining Assigner"
            End
        End If

        ass.Initialize row.Cells(1, 2).Value, row.Cells(1, 3).Value, _
                row.Cells(1, 4).Value, pDataLibDict(row.Cells(1, 1).Value)
        pAssignerList.Add ass
    Next row
    
    Exit Sub
ErrorHandler:
    If Not Application.EnableEvents Then
        Application.EnableEvents = True
    End If
    MsgBox "Error on Core.LoadDataLibraries(): " & Err.Description, vbCritical, "Core.LoadDataLibraries()"
    End
End Sub

' --------------Assign Value By Range
Public Sub AssignFromRange(cell As Range)
    Dim ass As Assigner
    
    For Each ass In pAssignerList

        If ass.TestCellInRange(cell) Then
            ass.AssignFromMatchingCell
            Exit Sub
        End If
        
    Next ass
End Sub

' --------------Assign Value By Node
Public Sub AssignFromNode(cell As Range, node As ClassNode)
    Dim ass As Assigner
    
    For Each ass In pAssignerList

        If ass.TestCellInRange(cell) Then
            ass.AssignFromNode node
            Exit Sub
        End If

    Next ass

    MsgBox "Unregister cell is selected. Please select valid cell to insert the data. ", vbExclamation
    End
End Sub


' --------------Get DataLibrary
Public Function GetDataLibrary(cell As Range) As DataLibrary
    Dim ass As Assigner
    
    Set GetDataLibrary = Nothing
    
    For Each ass In pAssignerList
        If ass.TestCellInRange(cell) Then
            Set pCurrentAssigner = ass
            pCurrentColumnColor = ass.Color
            Set GetDataLibrary = ass.dataLib
            Exit Function
        End If
    Next ass
    
End Function

' ----------------------------------------------------------------
' ----------------------------------------------------------------
' -------------------------- Data Library and Data Assign --------
' ----------------------------------------------------------------
' ----------------------------------------------------------------


' ----------------------------------------------------------------
' ----------------------------------------------------------------
' --------------------------Create Button-------------------------
' ----------------------------------------------------------------
' ----------------------------------------------------------------

Public Sub CreateItemSelectorButton(ByVal targetSheet As Worksheet)
    
    Dim shp           As Shape
    Dim buttonExists  As Boolean
    
    ' 1. Check if a button with the specific name already exists on the sheet
    For Each shp In targetSheet.Shapes
        If shp.Name = "button_E963474" Then
            buttonExists = True
            Exit For
        End If
    Next shp
    
    ' 2. If the button already exists, skip and exit the procedure
    If buttonExists Then
        Exit Sub
    End If
    
    ' 3. If no button is found:
    '    - Insert a new Row at Row 1
    '    - Set Row 1 height to 30
    targetSheet.Rows(1).Insert Shift:=xlDown
    targetSheet.Rows(1).RowHeight = 30
    
    ' 4. Create the button in the newly inserted Row 1, Column A
    Dim btnLeft    As Single
    Dim btnTop     As Single
    Dim newButton  As Shape
    
    ' Determine the top & left position for the button
    btnLeft = targetSheet.Cells(1, 1).Left + 2
    btnTop = targetSheet.Cells(1, 1).Top + 2
    
    ' Add a Forms button with the desired width (100) and height (30)
    Set newButton = targetSheet.Shapes.AddShape( _
                                Type:=msoShapeRoundedRectangle, _
                                Left:=btnLeft, _
                                Top:=btnTop, _
                                Width:=600, _
                                Height:=26)
    
    ' 5. Configure the newly created button
    With newButton
        .Name = "button_E963474" ' Unique name
        .OnAction = "OnButtonClicked"                  ' Assign macro to run when clicked
        
        .Adjustments.Item(1) = 0.5
        ' Apply green background and remove border
        .Fill.ForeColor.RGB = RGB(156, 205, 88)         ' Green fill
        .Line.Visible = msoTrue
        .Line.Weight = 0.2
        .Line.ForeColor.RGB = RGB(121, 159, 68)
    
        ' Set text
        .TextFrame2.TextRange.Text = "Click here to open Item Selector"    ' Button label
        .TextFrame2.TextRange.Font.Size = 18            ' Font size
        .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255) ' White text
        
        ' Center the text horizontally and vertically
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
    End With
    
End Sub

Public Sub OnButtonClicked()
    Dim targetShape As Shape

    ' Reference the clicked shape
    Set targetShape = ActiveSheet.Shapes(Application.Caller)

    ' Apply click effect: Darken the green fill temporarily
    With targetShape.Fill.ForeColor
        .RGB = RGB(126, 175, 68) ' Slightly darker green
    End With

    ' Pause to simulate a "press" effect
    Pause 0.1

    ' Restore the original fill color
    With targetShape.Fill.ForeColor
        .RGB = RGB(156, 205, 88) ' Original green
    End With

    ' Call the actual macro (OpenItemSelector)
    ClassItemSelector.Show vbModeless
    Core.IsUserFormLoaded = True
End Sub

' Helper function to pause execution for a specified duration
Private Sub Pause(seconds As Double)
    Dim startTime As Double
    startTime = timer
    Do While timer < startTime + seconds
        DoEvents
    Loop
End Sub
' ----------------------------------------------------------------
' ----------------------------------------------------------------
' --------------------------Create Button-------------------------
' ----------------------------------------------------------------
' ----------------------------------------------------------------



' ----------------------------------------------------------------
' ----------------------------------------------------------------
' --------------------------    Timer    -------------------------
' ----------------------------------------------------------------
' ----------------------------------------------------------------
' ---------------- Timer Callback Procedure
Private Sub TimerCallback(ByVal hwnd As LongPtr, ByVal uMsg As Long, _
                         ByVal idEvent As LongPtr, ByVal dwTime As Long)
    On Error Resume Next
    ' Call PerformFilter on the current form
    If Not CurrentForm Is Nothing Then
        CurrentForm.PerformDebounce
    End If

    ' Kill the timer to prevent repeated calls
    KillTimer 0, idEvent

    ' Reset TimerId
    TimerId = 0
    On Error GoTo 0
End Sub

' ---------------- Procedure to Start or Reset the Timer
Public Sub StartDebounceTimer()
    ' If a timer already exists, kill it
    If TimerId <> 0 Then
        KillTimer 0, TimerId
        TimerId = 0
    End If

    ' Set a new timer to trigger after 300 milliseconds (0.3 seconds)
    TimerId = SetTimer(0, 0, DEBOUCETIMER, AddressOf TimerCallback)
End Sub

' ----------------------------------------------------------------
' ----------------------------------------------------------------
' --------------------------    Timer    -------------------------
' ----------------------------------------------------------------
' ----------------------------------------------------------------

' ----------------------------------------------------------------
' Get Table object from the worksheet. The worksheet shall contain only 1 table.
' ----------------------------------------------------------------
Public Function GetTableObject(ws As Worksheet) As ListObject
    On Error GoTo ErrorHandler

    Dim tbl As ListObject
    
    Select Case ws.ListObjects.Count
        Case 0
            MsgBox "No tables found in worksheet!", vbExclamation, "Table Read Error"
            End
        Case 1
            Set GetTableObject = ws.ListObjects(1)
        Case Else
            MsgBox ws.Name & " Worksheet has more than one table." & vbCrLf & _
                   "Data Library worksheet shall only contain 1 table.", _
                   vbExclamation, "Table Read Error"
            End
    End Select
    
    Exit Function
    
ErrorHandler:
    MsgBox "Error on Core.GetTableObject(): " & Err.Description, vbCritical
    ' Return nothing if error occurs
    Set GetTableObject = Nothing
    End
End Function
