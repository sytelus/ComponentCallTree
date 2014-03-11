VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CHtml"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements CPrint
Private mvarOrientation As Integer
Private mvarTitle As String
Private mvarLvw As ListView
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

Private Property Get CPrint_lvwToPrint() As MSComctlLib.IListView

End Property

Private Property Set CPrint_lvwToPrint(ByVal RHS As MSComctlLib.IListView)
Set mvarLvw = RHS
End Property

Private Property Let CPrint_Orientation(ByVal RHS As Integer)
'This property is not needed in HTML
mvarOrientation = RHS
End Property


Private Function CPrint_printlvw() As Boolean

Dim outputLine As String
Dim fileNum As Integer
Dim cnt As Integer
Dim quotes As String
Dim lstItem As ListItem
Dim x As Integer
Dim subItem As ListSubItem


On Error GoTo errHandle
If mvarLvw Is Nothing Then
  CPrint_printlvw = False
  mvarLastErrorMsg = "List View Undefined"
  Exit Function
End If

  


fileNum = FreeFile
outputLine = ""

If mvarTitle = "" Then mvarTitle = mvarLvw.Name & " Output"

Open mvarFileName For Output As fileNum
'First Create the Header Information and the Column Headers
outputLine = ""
Print #fileNum, "<!-Created by PrintLvw.DLL Version " & App.Major & "." & App.Minor & "." & App.Revision & " ->"
Print #fileNum, "<!-Created on " & Date & " at " & Time & "->"
Print #fileNum, "<html>"
Print #fileNum, "<center>"
Print #fileNum, "<Title>" & mvarTitle
Print #fileNum, "</title>"
Print #fileNum, "<body bgcolor = " & Chr$(39) & "FFFFFF" & Chr$(39) & "text = " & Chr$(39) & "000000" & Chr$(39) & ">"
Print #fileNum, "<h1>"
Print #fileNum, mvarTitle
Print #fileNum, "</h1>"
Print #fileNum, "<Table border>"
Print #fileNum, "<tr>"
For x = 1 To mvarLvw.ColumnHeaders.Count
  Print #fileNum, "<td><center>" & mvarLvw.ColumnHeaders(x).Text & "</center></td>"
Next x
Print #fileNum, "</tr>"



cnt = 1
For Each lstItem In mvarLvw.ListItems
  Print #fileNum, "<tr>"
  Print #fileNum, "<td>" & lstItem.Text & "</td>"
  For Each subItem In lstItem.ListSubItems
    Print #fileNum, "<td>" & subItem.Text & "</td>"
  Next subItem
  Print #fileNum, "</tr>"
Next lstItem
  

Print #fileNum, "</table>"
Print #fileNum, "</Center>"
Print #fileNum, "</body>"
Print #fileNum, "</html>"
Close #fileNum

CPrint_printlvw = True
Exit Function

errHandle:
Call handleError.handleError(Err, "CHTML::PrintLvw")
mvarLastErrorMsg = Err.Description
CPrint_printlvw = False


End Function

Private Property Let CPrint_Title(ByVal RHS As String)
mvarTitle = RHS
End Property
