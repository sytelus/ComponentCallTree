VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarOrientation As Integer 'local copy
Private mvarTitle As String 'local copy
Public lvwToPrint As ListView
Private mvarFontName As String
Private mvarFontSize As Integer
Private mvarFileName As String
Private mvarLastErrorMsg As String
Private handleError As ErrorHandler

Public Property Get LastErrorMsg() As String

LastErrorMsg = mvarLastErrorMsg

End Property



Public Property Let fileName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Title = 5
    mvarFileName = vData
End Property


Public Property Let FontSize(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Title = 5
    mvarFontSize = vData
End Property



Public Property Let title(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Title = 5
    mvarTitle = vData
End Property

Public Property Let FontName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Title = 5
    mvarFontName = vData
End Property




Public Property Let Orientation(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Orientation = 5
    mvarOrientation = vData
End Property






Public Function PrintLvw() As Boolean
End Function


