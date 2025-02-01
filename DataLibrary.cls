VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataLibrary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'XNMAKERS CODE STARTS HERE
Option Explicit

' ----------------------------------------------------------------
' Private variables to hold property values
' ----------------------------------------------------------------
Private pLibraryName As String
Private pIsNodeListLoaded As Boolean

' Worksheet, Table, Headings Address
Private pWorksheetStr As String
Private pKeyHdg As String
Private pValueHdg As String
Private pParentHdg As String
Private pDescHdg As String

' Worksheet, Table, Ranges Objects
Private pWorksheetObj As Worksheet
Private pTableObj As ListObject
Private pKeyRng As Range
Private pValueRng As Range
Private pParentRng As Range
Private pDescRng As Range

' Collection to hold all ClassNodes
Private pNodeList As Collection
Private pRootNodeList As Collection

' ----------------------------------------------------------------
' Read-only Properties
' ----------------------------------------------------------------
Public Property Get LibraryName() As String
    LibraryName = pLibraryName
End Property
Public Property Get KeyHdg() As String
    KeyHdg = pKeyHdg
End Property
Public Property Get ValueHdg() As String
    ValueHdg = pValueHdg
End Property
Public Property Get WorksheetObj() As Worksheet
    Set WorksheetObj = pWorksheetObj
End Property
Public Property Get TableObj() As ListObject
    Set TableObj = pTableObj
End Property
Public Property Get KeyRng() As Range
    Set KeyRng = pKeyRng
End Property
Public Property Get ValueRng() As Range
    Set ValueRng = pValueRng
End Property
Public Property Get ParentRng() As Range
    Set ParentRng = pParentRng
End Property
Public Property Get DescRng() As Range
    Set DescRng = pDescRng
End Property
Public Property Get NodeList() As Collection
    Set NodeList = pNodeList
End Property
Public Property Get NodeCount() As Long
    NodeCount = pNodeList.Count
End Property

' ----------------------------------------------------------------
' Initialize the ChildrenClasses collection when a new ClassNode is created
' ----------------------------------------------------------------
Public Sub Initialize(libName As String, ws As String, keyStr As String, valueStr As String, parentStr As String, descStr As String)
    
    pLibraryName = libName
    pWorksheetStr = ws
    
    pKeyHdg = keyStr
    pValueHdg = valueStr
    pParentHdg = parentStr
    pDescHdg = descStr
    
    On Error Resume Next
    Set pWorksheetObj = ThisWorkbook.Worksheets(pWorksheetStr)
    If pWorksheetObj Is Nothing Then
        MsgBox "Worksheet '" & pWorksheetStr & "' not found! Please double-check if your Library: '" & pLibraryName & "' has the correct worksheet name defined within the 'ClassDataMapping' worksheet.", vbExclamation, "Worksheet name not found"
        End
    End If

    Set pTableObj = Core.GetTableObject(pWorksheetObj)
    If pTableObj Is Nothing Then
        GoTo ErrorHandler
    End If

    Set pKeyRng = pTableObj.ListColumns(pKeyHdg).DataBodyRange
    If pKeyRng Is Nothing Then
        MsgBox "Column Heading '" & pKeyHdg & "' is not found! Please double-check if your Library: '" & pLibraryName & "' has the correct heading name defined within the 'ClassDataMapping' worksheet.", vbExclamation, "Heading name not found"
        End
    End If

    Set pValueRng = pTableObj.ListColumns(pValueHdg).DataBodyRange
    If pValueRng Is Nothing Then
        MsgBox "Colume Heading '" & pValueHdg & "' is not found! Please double-check if your Library: '" & pLibraryName & "' has the correct heading name defined within the 'ClassDataMapping' worksheet.", vbExclamation, "Heading name not found"
        End
    End If

    Set pParentRng = pTableObj.ListColumns(pParentHdg).DataBodyRange
    If pParentRng Is Nothing Then
        MsgBox "Colume Heading '" & pParentRng & "' is not found! Please double-check if your Library: '" & pLibraryName & "' has the correct heading name defined within the 'ClassDataMapping' worksheet.", vbExclamation, "Heading name not found"
        End
    End If

    Set pDescRng = pTableObj.ListColumns(pDescHdg).DataBodyRange
    If pDescRng Is Nothing Then
        MsgBox "Colume Heading '" & pDescRng & "' is not found! Please double-check if your Library: '" & pLibraryName & "' has the correct heading name defined within the 'ClassDataMapping' worksheet.", vbExclamation, "Heading name not found"
        End
    End If

    pIsNodeListLoaded = False
    Exit Sub
ErrorHandler:
    MsgBox "Error on DataLibrary.Initialize(): " & Err.Description, vbCritical
    End
End Sub

' ----------------------------------------------------------------
' Load Data (Only load this when in the Form)
' ----------------------------------------------------------------
Public Sub LoadNodeList()
    On Error GoTo ErrorHandler
    
    Dim i As Long
    Dim currentNode As ClassNode
    Dim errorKey As String

    Set pNodeList = New Collection
    Set pRootNodeList = New Collection
    
    ' Create a dictionary to map key to ClassNode for quick lookup
    Dim uniqueNodeDict As Object
    Set uniqueNodeDict = CreateObject("Scripting.Dictionary")
    
    ' First pass: Create ClassNode objects for each row
    For i = 1 To pKeyRng.Count
        
        Set currentNode = New ClassNode
        currentNode.ClassKey = Trim(pKeyRng.Cells(i, 1).value)
        currentNode.ClassValue = Trim(pValueRng.Cells(i, 1).value)
        currentNode.ClassDesc = Trim(pDescRng.Cells(i, 1).value)
        currentNode.ParentKey = Trim(pParentRng.Cells(i, 1).value)
        
        errorKey = currentNode.ClassKey
        
        pNodeList.Add currentNode
        
        If Not uniqueNodeDict.Exists(currentNode.ClassKey) Then
            uniqueNodeDict.Add currentNode.ClassKey, currentNode
        End If

    Next i

    ' Second pass: Build parent-sub relationships
    For Each currentNode In pNodeList
        If uniqueNodeDict.Exists(currentNode.ParentKey) Then
            uniqueNodeDict(currentNode.ParentKey).AddSubNode currentNode
        Else
            If currentNode.ParentKey = "" Then
                pRootNodeList.Add currentNode
            Else
                Dim parentNode As ClassNode
                Set parentNode = New ClassNode
                parentNode.ClassKey = currentNode.ParentKey
                parentNode.ClassValue = ""
                parentNode.ClassDesc = "Root Class"
                parentNode.AddSubNode currentNode

                uniqueNodeDict.Add parentNode.ClassKey, parentNode
                pRootNodeList.Add parentNode
            End If
        End If

    Next currentNode
    
    pIsNodeListLoaded = True
    Exit Sub
ErrorHandler:
    MsgBox "Error on DataLibrary.LoadData(): " & Err.Description + "_" + errorKey, vbCritical
    End
End Sub

' ----------------------------------------------------------------
' Build Tree View (Only build this when in the Form)
' ----------------------------------------------------------------
Public Sub BuildTreeView(tvw As MSComctlLib.TreeView, keyFilteredStr As String, valueFilteredStr As String)
    On Error GoTo ErrorHandler
    
    Dim currentNode As ClassNode
    Dim treeNode As MSComctlLib.node
    Dim errorCode As String
    
    ' Clear existing nodes
    tvw.nodes.Clear
    
    ' Iterate through all nodes and check whether they are filtered visible
    For Each currentNode In pNodeList
        currentNode.FilterNodes keyFilteredStr, valueFilteredStr
    Next currentNode
    
    ' Iterate through root nodes
    For Each currentNode In pRootNodeList

        Dim treeNodeKey As String, treeNodeValue As String
        treeNodeKey = "key_" & currentNode.ClassKey & "_" & currentNode.ClassValue
        treeNodeValue = currentNode.ClassKey & "   -   " & currentNode.ClassValue

        errorCode = treeNodeKey

        ' Add root node to TreeView
        Set treeNode = tvw.nodes.Add(, , treeNodeKey, treeNodeValue)
        Set treeNode.Tag = currentNode
        
        ' Highlight if the currentNode is highlighted
        If currentNode.IsHighlighted Then
            treeNode.BackColor = RGB(255, 255, 0) ' Yellow highlight
        End If
        
        treeNode.expanded = True
        
        ' Recursively add child nodes
        AddSubNodes currentNode, treeNode, tvw.nodes, errorCode
        
    Next currentNode
    
    
    Exit Sub
ErrorHandler:
    MsgBox "Error on BuildTreeView(): " & Err.Description + "_" + errorCode, vbCritical
    End
End Sub

' ----------------------------------------------------------------
' Add Child Nodes (Only load this when in the Form)
' ----------------------------------------------------------------
Private Sub AddSubNodes(parentNode As ClassNode, parentTreeNode As MSComctlLib.node, tvwNodes As MSComctlLib.nodes, ByRef errorCode As String)

    Dim currentNode As ClassNode
    Dim treeNode As MSComctlLib.node
    
    For Each currentNode In parentNode.SubNodes
        ' Only add nodes that are filtered to be visible
        If currentNode.IsFiltered Or currentNode.HasFiltedSubNode Then
            
            Dim treeNodeKey As String, treeNodeValue As String
            treeNodeKey = "key_" & currentNode.ClassKey & "_" & currentNode.ClassValue
            treeNodeValue = currentNode.ClassKey & "   -   " & currentNode.ClassValue

            errorCode = treeNodeKey
            
            ' Add current node under the parent node
            Set treeNode = tvwNodes.Add(parentTreeNode, tvwChild, treeNodeKey, treeNodeValue)
            Set treeNode.Tag = currentNode
            
            ' Highlight if the currentTreeNode is highlighted
            If currentNode.IsHighlighted Then
                treeNode.BackColor = RGB(255, 255, 0) ' Yellow highlight
            End If
            
            ' Expand the tree if it has filtered sub
            If currentNode.HasHighlightedSubNode Then
                treeNode.expanded = True
            End If
            ' Recursively add grandchildren
            AddSubNodes currentNode, treeNode, tvwNodes, errorCode
        End If
    Next currentNode
End Sub

