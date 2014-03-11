VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CLvwPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Enum iOrientation
    iPortrait = 1
    iLandscape = 2
End Enum

Public Enum printType
  eCSV = 1
  eHtml = 2
  ePrinter = 3
End Enum

Public lvwToPrint As ListView
Private mvarLastErrorMsg As String
Private handleError As ErrorHandler



Public Property Get LastErrorMsg() As String

LastErrorMsg = mvarLastErrorMsg

End Property
Public Sub about()

Load frmAbout
frmAbout.Show vbModal


End Sub

Public Function PrintListView(ByVal Orientation As iOrientation, _
                        Optional fonttoprint As String = "Arial", Optional defaultTitle As String, _
                        Optional FontSize As Integer = 12) As Integer

On Error GoTo errHandle

PrintListView = 0

If lvwToPrint Is Nothing Then
  PrintListView = 1
  mvarLastErrorMsg = "List View Undefined"
  Exit Function
End If

  
  
Dim printJob As CPrint
Dim rc As Boolean


Set printJob = New CPrinter

printJob.Orientation = Orientation
printJob.FontName = fonttoprint
printJob.FontSize = FontSize
printJob.title = defaultTitle
Set printJob.lvwToPrint = lvwToPrint
rc = printJob.PrintLvw
If rc = False Then
  mvarLastErrorMsg = printJob.LastErrorMsg
  PrintListView = 1
End If

Call SafePrinterEndDoc

Exit Function

errHandle:
Call handleError.handleError(Err, "CLvwPrint::PrintLvw")



End Function

Private Sub SafePrinterEndDoc()
    On Error Resume Next
    Call Printer.EndDoc
End Sub

Public Function ExportListView(ByVal exportType As printType, Optional fileName As String, Optional ByVal Orientation As iOrientation = iPortrait, _
                              Optional fonttoprint As String = "Arial", Optional defaultTitle As String, _
                              Optional FontSize As Integer = 12) As Integer


On Error GoTo errHandle

ExportListView = 0

If lvwToPrint Is Nothing Then
  ExportListView = 1
  mvarLastErrorMsg = "List View Not Defined"
  Exit Function
End If


Dim printJob As CPrint
Dim rc As Boolean


Select Case exportType
  Case eCSV
    Set printJob = New CCsv
  Case eHtml
    Set printJob = New CHtml
  Case ePrinter
    Set printJob = New CPrinter
  Case Else
    mvarLastErrorMsg = "Incorrect type of print job was specified."
    ExportListView = 1
    Exit Function
End Select

Set printJob.lvwToPrint = lvwToPrint
printJob.Orientation = 1
printJob.title = defaultTitle
printJob.fileName = fileName
printJob.FontName = fonttoprint
printJob.FontSize = FontSize
rc = printJob.PrintLvw
If rc = False Then
  mvarLastErrorMsg = printJob.LastErrorMsg
  ExportListView = 1
End If


Exit Function

errHandle:
Call handleError.handleError(Err, "CLvwPrint::ExportListView")




End Function





Private Sub Class_Initialize()

Set handleError = New ErrorHandler

End Sub

Private Sub Class_Terminate()

Set lvwToPrint = Nothing
Set handleError = Nothing


End Sub
