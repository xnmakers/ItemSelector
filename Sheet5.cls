Option Explicit

Private Sub Worksheet_SelectionChange(ByVal target As Range)

    If Core.IsUserFormLoaded Then
        ClassItemSelector.OnClassSelectionChanged target.Cells(1, 1)
    End If
    Core.CreateItemSelectorButton target.Worksheet
    
End Sub

Private Sub Worksheet_Change(ByVal current As Range)
    On Error GoTo ErrorHandler ' Ensure events are re-enabled in case of an error
    
    ' Disable events to prevent recursion
    Application.EnableEvents = False
    
    Core.LoadDataLibraries
    Core.AssignFromRange current.Cells(1, 1)
    
ErrorHandler:
    ' Re-enable events
    If Not Application.EnableEvents Then
        Application.EnableEvents = True
    End If
End Sub
