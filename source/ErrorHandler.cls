VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ErrorHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Private printErrorFile As String
Private fileNum As Integer

Public Sub handleError(ByVal ErrorObject As ErrObject, Optional location As String, Optional fileName As String)

printErrorFile = App.Path & "\lvwprint.err"
fileNum = FreeFile
Open printErrorFile For Append Shared As #fileNum
Print #fileNum, "An Error was raised on " & Date & " " & Time
Print #fileNum, "=====" & "Error Number = " & ErrorObject.Number
Print #fileNum, "=====" & "Error Desc = " & ErrorObject.Description 'Write the output to the outfile
Print #fileNum, "=====" & "Error Source = " & Err.Source
Close #fileNum



End Sub

