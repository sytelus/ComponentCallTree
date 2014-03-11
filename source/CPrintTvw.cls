VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CPrintTvw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Option Explicit
Private mvarPrintConnectorLines As Boolean 'local copy
Public tvwToPrint As TreeView
Private mvarFontSize As Integer 'local copy
Private mvarFontBold As Boolean 'local copy
Private mvarTitle As String 'local copy
Private mvarSecondTitle As String 'local copy
Private mvarPrintFooter As Boolean 'local copy
Public Enum printOptions
    ePrintAll = 1  'The entire tree view
    ePrintFromSelected = 2  'Only the selected item and down
End Enum
Public Enum ConnectLines
  econnectlines = 1
  eNoConnectLinesWithIndent = 2
  eNoConnectLinesNoIndent = 3
End Enum

Public Event PrepareNode(ByVal Node As MSComctlLib.Node)


'To fire this event, use RaiseEvent with the following syntax:
'RaiseEvent ProcessingTreeView[(arg1, arg2, ... , argn)]
Public Event ProcessingTreeView()
'To fire this event, use RaiseEvent with the following syntax:
'RaiseEvent SendingToPrinter[(arg1, arg2, ... , argn)]
Public Event SendingToPrinter()
'local variable(s) to hold property value(s)
Private mvarTitleFontSize As Integer 'local copy
Private mvarSecondTitleFontSize As Integer 'local copy
Private mvarTitleFontBold As Boolean 'local copy
Private mvarSecondTitleFontBold As Boolean 'local copy
Private Const CONNECTOR_WIDTH = 100
Private Const DROP_WIDTH = 100
Private Const INDENT = 200
Public Event PrintJobComplete()
'local variable(s) to hold property value(s)
Private mvarConnectorStyle As ConnectLines 'local copy
'local variable(s) to hold property value(s)
Private mvarPrintStyle As printOptions 'local copy
Public Property Let PrintStyle(ByVal vData As printOptions)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PrintStyle = 5
    mvarPrintStyle = vData
End Property


Public Property Get PrintStyle() As printOptions
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PrintStyle
    PrintStyle = mvarPrintStyle
End Property



Public Property Let ConnectorStyle(ByVal vData As ConnectLines)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ConnectorStyle = 5
    mvarConnectorStyle = vData
End Property


Public Property Get ConnectorStyle() As ConnectLines
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ConnectorStyle
    ConnectorStyle = mvarConnectorStyle
End Property





Public Property Let SecondTitleFontBold(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SecondTitleFontBold = 5
    mvarSecondTitleFontBold = vData
End Property


Public Property Get SecondTitleFontBold() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SecondTitleFontBold
    SecondTitleFontBold = mvarSecondTitleFontBold
End Property



Public Property Let TitleFontBold(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TitleFontBold = 5
    mvarTitleFontBold = vData
End Property


Public Property Get TitleFontBold() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TitleFontBold
    TitleFontBold = mvarTitleFontBold
End Property



Public Property Let SecondTitleFontSize(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SecondTitleFontSize = 5
    mvarSecondTitleFontSize = vData
End Property


Public Property Get SecondTitleFontSize() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SecondTitleFontSize
    SecondTitleFontSize = mvarSecondTitleFontSize
End Property



Public Property Let TitleFontSize(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TitleFontSize = 5
    mvarTitleFontSize = vData
End Property


Public Property Get TitleFontSize() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TitleFontSize
    TitleFontSize = mvarTitleFontSize
End Property






Public Property Let PrintFooter(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PrintFooter = 5
    mvarPrintFooter = vData
End Property


Public Property Get PrintFooter() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PrintFooter
    PrintFooter = mvarPrintFooter
End Property



Public Property Let SecondTitle(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SecondTitle = 5
    mvarSecondTitle = vData
End Property


Public Property Get SecondTitle() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SecondTitle
    SecondTitle = mvarSecondTitle
End Property



Public Property Let title(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Title = 5
    mvarTitle = vData
End Property


Public Property Get title() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Title
    title = mvarTitle
End Property



Public Property Let FontBold(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FontBold = 5
    mvarFontBold = vData
End Property


Public Property Get FontBold() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FontBold
    FontBold = mvarFontBold
End Property



Public Property Let FontSize(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FontSize = 5
    mvarFontSize = vData
End Property


Public Property Get FontSize() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FontSize
    FontSize = mvarFontSize
End Property




Private Sub PrintSiblingLine(x As Integer, Y As Integer)

Select Case mvarConnectorStyle
  Case 1
    Printer.Line (x, Y)-(x, Printer.CurrentY)
End Select



End Sub
Private Sub PrintDropLine(lengthOfItem As Integer)

Select Case mvarConnectorStyle
  Case 1  'Connector Lines
    Printer.Line (Printer.currentX + lengthOfItem, Printer.CurrentY)-(Printer.currentX + lengthOfItem, Printer.CurrentY + DROP_WIDTH)
  Case 2 ' Connector with Indent
    Printer.currentX = Printer.currentX + lengthOfItem
    Printer.CurrentY = Printer.CurrentY + DROP_WIDTH
End Select



End Sub

Private Sub PrintConnector()

Select Case mvarConnectorStyle
  Case 1
    Printer.Line -(Printer.currentX + CONNECTOR_WIDTH, Printer.CurrentY)
  Case 2
    Printer.currentX = Printer.currentX + CONNECTOR_WIDTH
End Select



End Sub


Private Sub RecurseTree(ByRef tvwParent As Node, iLevel As Integer, currentX As Integer, curSiblingX As Integer, curSiblingY As Integer)


Dim newNode As Node
Dim widthOfNode As Single
Dim memberNum As Integer  'The number of the node at the selected level
Dim holdcurx As Integer
Dim holdCurXSibling As Integer
Dim holdCurYSibling As Integer
Dim lastPage As Integer

RaiseEvent PrepareNode(tvwParent)

lastPage = Printer.Page  'Figure out what page the output is on when we enter the routine

If Not tvwParent.Child Is Nothing Then  'We want to print the connector lines
  
  'Call PrintDropLine(currentX + Printer.TextWidth(tvwParent.Text))
  
    Call PrintDropLine(currentX + INDENT)
    Call PrintConnector
    holdcurx = Printer.currentX
    holdCurYSibling = Printer.CurrentY - DROP_WIDTH
    holdCurXSibling = Printer.currentX - CONNECTOR_WIDTH
    Printer.Print (tvwParent.Child)
  
  
    If lastPage <> Printer.Page Then  'Moved to the next page of output
      Call PrintFooterText
      Call RecurseTree(tvwParent.Child, iLevel + 1, holdcurx, holdCurXSibling, 0)
    Else
      Call RecurseTree(tvwParent.Child, iLevel + 1, holdcurx, holdCurXSibling, holdCurYSibling)
    End If
  
'  Set newNode = frmMain.tvwTest.Nodes.Add(tvwParent.Index, tvwChild, , tvwParent.Child.Text)



  
End If



If Not tvwParent.Next Is Nothing Then
  If iLevel = 0 And mvarPrintStyle = ePrintFromSelected Then
   Exit Sub
  End If
  If iLevel <> 0 Then
    Call PrintDropLine(curSiblingX)
    Select Case lastPage
     Case Is = Printer.Page
       Call PrintSiblingLine(curSiblingX, curSiblingY)
     Case Else
       Call PrintSiblingLine(curSiblingX, 0)
    End Select
    Call PrintConnector
     
    
    'Call PrintDropLine(holdCurXSibling)
    
    'Call PrintSiblingLine(holdCurXSibling, holdCurYSibling)
    
  End If
  holdcurx = Printer.currentX
  
  Printer.Print ("  " & tvwParent.Next.Text)
  
  If lastPage <> Printer.Page Then
    Call PrintFooterText
    Call RecurseTree(tvwParent.Next, iLevel, holdcurx, curSiblingX, 0)
  Else
    Call RecurseTree(tvwParent.Next, iLevel, holdcurx, curSiblingX, curSiblingY)
  End If
End If

End Sub


Public Function PrintTvw()

Dim selNode As Node
Dim newNode As Node

Select Case mvarPrintStyle
  Case 1 'Print As Shown, Print All
    Set selNode = tvwToPrint.Nodes(1)
  Case 2  'Print From Selected
    If tvwToPrint.SelectedItem Is Nothing Then
        Err.Raise 1000, , "Nothing selected in Tree"
    Else
        Set selNode = tvwToPrint.SelectedItem
    End If

End Select




Call PrintHeader
Call PrintFooterText


Printer.FontSize = Me.FontSize
Printer.FontBold = Me.FontBold
Printer.DrawWidth = 7
Printer.Print ("    " & selNode.Text)
RaiseEvent ProcessingTreeView
  
Call RecurseTree(selNode, 0, 0, 0, 0)

PrintEnd:
RaiseEvent SendingToPrinter
Printer.EndDoc
RaiseEvent PrintJobComplete

End Function



Private Sub PrintHeader()


If Me.title <> "" Then
  Printer.FontSize = Me.TitleFontSize
  Printer.FontBold = Me.TitleFontBold
  Printer.currentX = (Printer.Width - Printer.TextWidth(Me.title)) / 2
  Printer.Print (Me.title)
End If


If Me.SecondTitle <> "" Then
  Printer.FontSize = Me.SecondTitleFontSize
  Printer.FontBold = Me.SecondTitleFontBold
  Printer.currentX = (Printer.Width - Printer.TextWidth(Me.SecondTitle)) / 2
  Printer.Print Me.SecondTitle
End If

If Me.title <> "" Or Me.SecondTitle <> "" Then
  Printer.Print ""
  Printer.Print ""
End If



End Sub


Private Sub PrintFooterText()

If Me.PrintFooter = False Then Exit Sub

Dim holdY As Integer
Dim holdFontSize As Integer

'Now print the Footer on this page

holdY = Printer.CurrentY
holdFontSize = Printer.FontSize


Printer.FontSize = 6
'Debug.Print "Header - " & Printer.TextHeight("Footer")
Printer.CurrentY = Printer.Height - (Printer.TextHeight("Footer") * 6)
Printer.Print Date
Printer.CurrentY = Printer.Height - (Printer.TextHeight("Footer") * 6)
Printer.currentX = Printer.Width - Printer.TextWidth(Time) * 2.5
Printer.Print Time
Printer.CurrentY = Printer.Height - (Printer.TextHeight("Footer") * 6)
Printer.currentX = (Printer.Width - Printer.TextWidth("Page -" & Printer.Page & "-")) / 2
Printer.Print "-Page " & Printer.Page & "-"
Printer.CurrentY = holdY
Printer.FontSize = holdFontSize


End Sub
