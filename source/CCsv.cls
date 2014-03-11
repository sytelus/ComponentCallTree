VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCsv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements CPrint
Private mvarLvw As ListView
Private mvarOrientation As Integer
Private mvarTitle As String
Private mvarFileName As String
Private mvarLastErrorMsg As String
Private handleError As ErrorHandler



Private Sub Class_Initialize()
Set handleError = New ErrorHandler
End Sub

Private Sub Class_Terminate()
Set handleError = Nothing
End Sub

Private Property Let CPrint_FileName(ByVal RHS As String)
mvarFileName = RHS
End Property

Private Property Let CPrint_FontName(ByVal RHS As String)

End Property

Private Property Let CPrint_FontSize(ByVal RHS As Integer)

End Property

Private Property Get CPrint_LastErrorMsg() As String
CPrint_LastErrorMsg = mvarLastErrorMsg
End Property

Private Property Let CPrint_lvwToPrint(RHS As Variant)

End Property

Private Property Set CPrint_lvwToPrint(ByVal RHS As MSComctlLib.IListView)
Set mvarLvw = RHS
End Property

Private Property Get CPrint_lvwToPrint() As MSComctlLib.IListView

End Property

Private Property Let CPrint_Orientation(ByVal RHS As Integer)
mvarOrientation = RHS

End Property

Private Function CPrint_printlvw() As Boolean
Dim outputLine As String
Dim fileNum As Integer
Dim cnt As Integer
Dim lstItem As ListItem
Dim subItem As ListSubItem

On Error GoTo errHandle


If mvarLvw Is Nothing Then
  CPrint_printlvw = False
  mvarLastErrorMsg = "List View Not Defined"
  Exit Function
End If
fileNum = FreeFile
outputLine = ""


Open mvarFileName For Output As fileNum

For Each lstItem In mvarLvw.ListItems
  outputLine = ""
  outputLine = outputLine & Chr$(34) & lstItem.Text & Chr$(34) & ","
  For Each subItem In lstItem.ListSubItems
    outputLine = outputLine & Chr$(34) & subItem.Text & Chr$(34) & ","
  Next subItem
  Print #fileNum, Left(outputLine, Len(outputLine) - 1)
Next lstItem

Close #fileNum
CPrint_printlvw = True
Exit Function

errHandle:
Call handleError.handleError(Err, "CCsv::PrintLvw")
mvarLastErrorMsg = Err.Description
CPrint_printlvw = False



End Function

Private Property Let CPrint_Title(ByVal RHS As String)
mvarTitle = RHS
End Property
