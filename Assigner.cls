VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Assigner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ----------------------------------------------------------------
' Private variables to hold property values
' ----------------------------------------------------------------
Private pDataLib As DataLibrary
Private pCellProperty As ColumnProperty

Private pKeyCell As Range
Private pValueCell As Range

Private pKeyHdrStr As String
Private pValueHdrStr As String

Private pWorksheetObj As Worksheet
Private pTableObj As ListObject
Private pKeyRng As Range
Private pValueRng As Range

Private pColor As Long

' ----------------------------------------------------------------
' Read-Only Properties
' ----------------------------------------------------------------
Public Property Get dataLib() As DataLibrary
    Set dataLib = pDataLib
End Property
Public Property Get Color() As Long
    Color = pColor
End Property
Public Property Get KeyHdrStr() As String
    KeyHdrStr = pKeyHdrStr
End Property
Public Property Get ValueHdrStr() As String
    ValueHdrStr = pValueHdrStr
End Property

' ----------------------------------------------------------------
' Initialize the Assigner
' ----------------------------------------------------------------
Public Sub Initialize(ws As String, keyStr As String, valueStr As String, lib As DataLibrary)

    Dim headerCell As Range

    pKeyHdrStr = keyStr
    pValueHdrStr = valueStr

    Set pDataLib = lib

    On Error Resume Next
    Set pWorksheetObj = ThisWorkbook.Worksheets(ws)
    If pWorksheetObj Is Nothing Then
        MsgBox "Worksheet '" & ws & "' not found! Please double-check the worksheet name in the 'ClassDataMapping' worksheet.", vbExclamation, "Worksheet Not Found"
        End
    End If

    Set pTableObj = Core.GetTableObject(pWorksheetObj)
    If pTableObj Is Nothing Then
        GoTo ErrorHandler
    End If
    
    Set pKeyRng = pTableObj.ListColumns(keyStr).DataBodyRange
    If pKeyRng Is Nothing Then
        MsgBox "Header name '" & keyStr & "' is not found in the worksheet '" & ws & "'! Please double check whether the name of the key header is correctly defined in the 'ClassDataMapping' worksheet.", vbExclamation, "Key header not found"
        End
    End If

    Set pValueRng = pTableObj.ListColumns(valueStr).DataBodyRange
    If pValueRng Is Nothing Then
        MsgBox "Header name '" & valueStr & "' is not found in the worksheet '" & ws & "'! Please double check whether the name of the value header is correctly defined in the 'ClassDataMapping' worksheet.", vbExclamation, "Value header not found"
        End
    End If

    Set pKeyCell = Nothing
    Set pValueCell = Nothing
    pCellProperty = ColumnProperty.NotAssigned
    
    ' Find the header cell for the specified heading
    Set headerCell = pTableObj.HeaderRowRange.Find(What:=keyStr, LookIn:=xlValues, LookAt:=xlWhole)

    ' Check if the header was found
    If Not headerCell Is Nothing Then
        ' Get the background color of the header
        pColor = headerCell.Interior.Color
    Else
        pColor = RGB(255, 255, 255) ' Default color as white
    End If
    
    Exit Sub
ErrorHandler:
    MsgBox "Error on Assigner.Initialize(): " & Err.Description, vbCritical
    End
End Sub

' ----------------------------------------------------------------
' --------------Assign Matching Cell From Range
' ----------------------------------------------------------------
Public Function AssignFromMatchingCell() As Boolean
    On Error GoTo ErrorHandler

    If pValueCell Is Nothing Or pKeyCell Is Nothing Then
        GoTo ErrorHandler
    End If
    
    Dim i As Long

    If pCellProperty = ColumnProperty.KeyColumn Then
        For i = 1 To pDataLib.KeyRng.Count
            If pKeyCell.Cells(1, 1).value = Trim(pDataLib.KeyRng.Cells(i, 1).value) Then
                pValueCell.value = pDataLib.ValueRng.Cells(i, 1).value
                AssignFromMatchingCell = True
                Exit Function
            End If
        Next i
    ElseIf pCellProperty = ColumnProperty.ValueColumn Then
        For i = 1 To pDataLib.ValueRng.Count
            If pValueCell.Cells(1, 1).value = Trim(pDataLib.ValueRng.Cells(i, 1).value) Then
                pKeyCell.value = pDataLib.KeyRng.Cells(i, 1).value
                AssignFromMatchingCell = True
                Exit Function
            End If
        Next i
    End If
    
    AssignFromMatchingCell = False
    Exit Function

ErrorHandler:
    AssignFromMatchingCell = False
    If Not Application.EnableEvents Then
        Application.EnableEvents = True
    End If
    MsgBox "Error on Assigner.AssignFromMatchingCell(): " & Err.Description, vbCritical
    End
End Function

' ----------------------------------------------------------------
' --------------Assign Matching Cell From Node
' ----------------------------------------------------------------
Public Function AssignFromNode(node As ClassNode) As Boolean
    On Error GoTo ErrorHandler

    Application.EnableEvents = False
    If pValueCell Is Nothing Or pKeyCell Is Nothing Then
        GoTo ErrorHandler
    End If
    
    pKeyCell.value = node.ClassKey
    pValueCell.value = node.ClassValue
    AssignFromNode = True

    Application.EnableEvents = True
    Exit Function
    
ErrorHandler:
    AssignFromNode = False
    If Not Application.EnableEvents Then
        Application.EnableEvents = True
    End If
    MsgBox "Error on Assigner.AssignFromNode(): " & Err.Description, vbCritical
    End
End Function

' ----------------------------------------------------------------
' --------------Test Cell Location
' ----------------------------------------------------------------
Public Function TestCellInRange(cell As Range) As Boolean
    On Error GoTo ErrorHandler
    
    Dim thisCell As Range
    Dim rowIndex As Long
    
    Set thisCell = cell.Cells(1, 1)

    If thisCell.Worksheet.Name <> pWorksheetObj.Name Then
        TestCellInRange = False
        Exit Function
    End If
    
    If Not Intersect(thisCell, pKeyRng) Is Nothing Then
    
        rowIndex = thisCell.row - pKeyRng.row + 1
        Set pKeyCell = thisCell
        Set pValueCell = pValueRng.Cells(rowIndex, 1)
        pCellProperty = ColumnProperty.KeyColumn
        TestCellInRange = True
        Exit Function
        
    End If
    If Not Intersect(cell, pValueRng) Is Nothing Then
        
        rowIndex = thisCell.row - pValueRng.row + 1
        Set pValueCell = thisCell
        Set pKeyCell = pKeyRng.Cells(rowIndex, 1)
        pCellProperty = ColumnProperty.ValueColumn
        TestCellInRange = True
        Exit Function
        
    End If
    
    TestCellInRange = False
    
    Exit Function
ErrorHandler:
    TestCellInRange = False
    If Not Application.EnableEvents Then
        Application.EnableEvents = True
    End If
    MsgBox "Error on Assigner.AssignFromNode(): " & Err.Description, vbCritical
    End
End Function



