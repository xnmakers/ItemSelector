Attribute VB_Name = "Info"
'XNMAKERS CODE STARTS HERE
Option Private Module
Option Explicit
Public Const INFO_VERSION As Double = 3.1
Public Const INFO_DATE As String = "02/10/2025"
Public Const INFO_FILES_LIST As String = "Core.bas" & vbCrLf & _
                                         "Info.bas" & vbCrLf & _
                                         "ClassItemSelector.frm" & vbCrLf & _
                                         "Assigner.cls" & vbCrLf & _
                                         "ClassNode.cls" & vbCrLf & _
                                         "DataLibrary.cls"

Public Const INFO_AUTHOR As String = "Xingyun Jin"
Public Const INFO_EMAIL As String = "xingyun.jin@avangrid.com"
Public Const INFO_TITLE = "Item Selector - v" & INFO_VERSION & " - " & INFO_DATE
Public Const INFO_DESCRIPTION = "The Macro UI is to generate Tree Structure using the specified data." & _
                                "The Macro UI supports auto-fill of the items to the pairing cells, based on the data library that is mapped from Worksheet, [ClassDataMapping] ."
Public Const INFO_KNOWNISSUE = vbCrLf & "Known Issues:" & vbCrLf & _
                                "1. The Macro clears the Undo and Redo (ctrl+z, ctrl+y) stack by design" & vbCrLf & _
                                "2. Auto-fill of associating cell functions one at a time. Mass Auto-fill is not supported at the moment" & vbCrLf & _
                                "3. The Macro is not tested on Mac OS" & vbCrLf & _
                                "4. The Macro is not tested on Excel 2010 or earlier versions" & vbCrLf & _
                                "5. The Macro is not tested on Excel Online" & vbCrLf & _
                                "6. The Macro is not tested on Excel Mobile"

Public Const INFO_CONTACT = vbCrLf & "Please contact the Author, " & INFO_AUTHOR & " for more information " & vbCrLf & INFO_EMAIL

