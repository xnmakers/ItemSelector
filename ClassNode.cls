VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'XNMAKERS CODE STARTS HERE
Option Explicit

' -----------------------------
' Private Members
' -----------------------------
Private pParentKey As String
Private pClassKey As String
Private pClassValue As String
Private pClassDesc As String

Private pParentNode As ClassNode
Private pSubNodes As Collection

Private pIsHighlighted As Boolean
Private pIsFiltered As Boolean
' -----------------------------
' Public Properties
' -----------------------------
' Property for ParentKey
Public Property Get ParentKey() As String
    ParentKey = pParentKey
End Property
Public Property Let ParentKey(value As String)
    pParentKey = value
End Property

' Property for ClassKey
Public Property Get ClassKey() As String
    ClassKey = pClassKey
End Property
Public Property Let ClassKey(value As String)
    pClassKey = value
End Property

' Property for ClassValue
Public Property Get ClassValue() As String
    ClassValue = pClassValue
End Property
Public Property Let ClassValue(value As String)
    pClassValue = value
End Property

' Property for ClassDescription
Public Property Get ClassDesc() As String
    ClassDesc = pClassDesc
End Property
Public Property Let ClassDesc(value As String)
    pClassDesc = value
End Property

' Property for ParentNode (as ClassNode)
Public Property Get parentNode() As ClassNode
    Set parentNode = pParentNode
End Property
Public Property Let parentNode(value As ClassNode)
    Set pParentNode = value
End Property

' Read-only Property for SubNodes (as Collection of ClassNode)
Public Property Get SubNodes() As Collection
    Set SubNodes = pSubNodes
End Property

' Read-only Property for IsHighlighted
Public Property Get IsHighlighted() As Boolean
    IsHighlighted = pIsHighlighted
End Property

' Read-only Property for IsFiltered
Public Property Get IsFiltered() As Boolean
    IsFiltered = pIsFiltered
End Property

' Read-only Property for IsRootParent
Public Property Get IsRoot() As Boolean
    If pParentNode Is Nothing Then
        IsRoot = True
    Else
        IsRoot = False
    End If
End Property

' Read-only Property for HasHighlightedSubNode
Public Property Get HasHighlightedSubNode() As Boolean
    Dim node As ClassNode
    HasHighlightedSubNode = False
    
    For Each node In pSubNodes
        If node.IsHighlighted Or node.HasHighlightedSubNode Then
            HasHighlightedSubNode = True
            Exit Property
        End If
    Next node
End Property

' Read-only Property for HasFiltedSubNodes
Public Property Get HasFiltedSubNode() As Boolean
    Dim node As ClassNode
    HasFiltedSubNode = False

    For Each node In pSubNodes
        If node.IsFiltered Or node.HasFiltedSubNode Then
            HasFiltedSubNode = True
            Exit Property
        End If
    Next node
End Property


' -----------------------------
' Methods
' -----------------------------
' Initialize the ChildrenClasses collection when a new ClassNode is created
Private Sub Class_Initialize()
    Set pSubNodes = New Collection
    Set pParentNode = Nothing
    pIsHighlighted = False
    pIsFiltered = True
End Sub

' Method to check if Node matches filter and update properties
Public Sub FilterNodes(keyFilteredString As String, valueFilteredString As String)
    
    If keyFilteredString = "" And valueFilteredString = "" Then
        pIsHighlighted = False
        pIsFiltered = True
    ElseIf InStr(1, pClassKey, keyFilteredString, vbTextCompare) > 0 And _
                InStr(1, pClassValue, valueFilteredString, vbTextCompare) > 0 Then
        pIsHighlighted = True
        pIsFiltered = True
    Else
        pIsHighlighted = False
        pIsFiltered = False
    End If
End Sub

' Method to add a child ClassNode
Public Sub AddSubNode(node As ClassNode)
    pSubNodes.Add node
    node.parentNode = Me
End Sub

