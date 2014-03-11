VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CPrinter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements CPrint
Private mvarTitle As String
Private mvarOrientation As Integer
Private mvarLvw As ListView
Private mvarFont As String
Private mvarFontSize As Integer
Private totalWidth As Integer
Private holdItemWidth() As Integer
Private mvarLastErrorMsg As String
Private handleError As ErrorHandler
Private mvarFooterLocation As Integer




Private Sub Class_Initialize()
Set handleError = New ErrorHandler
End Sub

Private Sub Class_Terminate()
Set handleError = Nothing
End Sub

Private Property Let CPrint_FileName(ByVal RHS As String)

End Property

Private Property Let CPrint_FontName(ByVal RHS As String)
mvarFont = RHS
End Property

Private Property Let CPrint_FontSize(ByVal RHS As Integer)
mvarFontSize = RHS
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

Dim lstItem As ListItem
Dim x As Integer
Dim colHeader As ColumnHeader
Dim cntHeader As Integer


Dim footerHeight As Integer
Dim holdPage As Integer
Dim holdLineNumber As Integer
Dim widthOfOutput As Integer
Dim outputFits As Boolean

Dim freeSpace As Integer
Dim columnAllocation As Integer
Dim resp As Integer
Dim output As String
Dim holdXPos As Integer

On Error GoTo errHandle

If mvarLvw Is Nothing Then
  CPrint_printlvw = False
  mvarLastErrorMsg = "List View Not Defined"
  Exit Function
End If



Printer.Font = mvarFont
Printer.Orientation = mvarOrientation

If mvarTitle = "" Then mvarTitle = mvarLvw.Name & " Output"


outputFits = FitOutputToPage(mvarFontSize)

If Not outputFits And Printer.Orientation = 1 Then
  resp = MsgBox("The output that you have defined is currently wider than the page.  Would you like to rotate the output?", vbQuestion + vbYesNo, "Output to Wide")
  If resp = vbYes Then
    Printer.Orientation = 2
    If Not FitOutputToPage(mvarFontSize) Then
      MsgBox "Unable to fit output to the page even after rotating.  Please Adjust", vbCritical, "Output Overflow"
      CPrint_printlvw = False
      Exit Function
    End If
  End If
End If

If outputFits Then
  'Determine the amount of free space
  freeSpace = (Printer.Width - 600) - totalWidth  'Subtract 600 to ensure that all output fits on the page
  'Allocate the freeSpace each column
  columnAllocation = freeSpace \ mvarLvw.ColumnHeaders.Count
  For x = 1 To UBound(holdItemWidth)
    holdItemWidth(x) = holdItemWidth(x) + columnAllocation
  Next x
End If


Call PrintTitle(mvarTitle)
Call PrintHeader(1)

output = ""
holdXPos = 0
holdPage = Printer.Page


'Debug.Print "Paper Height = " & Printer.Height

For Each lstItem In mvarLvw.ListItems
  holdXPos = 0
  holdLineNumber = Printer.CurrentY
  
  'Debug.Print "Line Number = " & holdLineNumber
  'Debug.Print "Line Height = " & Printer.TextHeight("This is the text")
  
  If (Printer.CurrentY + Printer.TextHeight("SomeText")) > mvarFooterLocation Then 'Print a blank line
    Printer.Print output
  End If
  
  Printer.Print output  'Print a blank line to see if we are moving to the next page
  
  If holdPage <> Printer.Page Then  'We have moved to a new page
    Call PrintHeader(Printer.Page)
    holdLineNumber = Printer.CurrentY
    holdPage = Printer.Page
  End If
  

  Printer.CurrentY = holdLineNumber
  Printer.currentX = holdXPos
  Printer.Print lstItem.Text
  
  
  
  For x = 1 To mvarLvw.ColumnHeaders.Count - 1
    Printer.CurrentY = holdLineNumber
    Printer.currentX = holdXPos + holdItemWidth(x)
    holdXPos = Printer.currentX
    Printer.Print lstItem.SubItems(x)
  Next x
  'holdLineNumber = Printer.CurrentY
  
  'Printer.Print output
  
  
Next lstItem

'Printer.PSet

'footerHeight = Printer.Height
'Printer.NewPage
'Printer a footer on the page
'Printer.CurrentX = 0
'Printer.CurrentY = footerHeight
'Printer.FontSize = 6
'Printer.FontBold = True
'Printer.Print Date
'Printer.CurrentX = (Printer.Width - Printer.TextWidth(mytelFile.fileName)) / 2
'Printer.CurrentY = footerHeight
'Printer.Print mytelFile.fileName
'Printer.CurrentX = Printer.Width - Printer.TextWidth(Time)
'Printer.CurrentY = footerHeight
'Printer.Print Time

'printer.EndDoc

CPrint_printlvw = True
Exit Function

errHandle:
Call handleError.handleError(Err, "CPrinter::PrintLvw")
mvarLastErrorMsg = Err.Description
CPrint_printlvw = False


End Function

Private Property Let CPrint_Title(ByVal RHS As String)
mvarTitle = RHS

End Property

Private Function DetermineWidth(FontSize As Integer) As Integer



If mvarLvw Is Nothing Then
  Call Err.Raise("9888888", "PrintListView", "List View Not Defined.  Set the lvwToPrint Variable in the CLvwPrint Class")
End If

Dim holdWidth As Integer
Dim lstItem As ListItem
Dim x As Integer




holdWidth = 0
ReDim holdItemWidth(1 To mvarLvw.ColumnHeaders.Count) As Integer

Printer.FontSize = FontSize  'Set the font size to get the proper width calculation
Printer.FontBold = True   'Set the font to bold since the column headers will be bold


For x = 1 To UBound(holdItemWidth)
  holdItemWidth(x) = 0
Next x

For Each lstItem In mvarLvw.ListItems
  
  If holdItemWidth(1) < Printer.TextWidth(lstItem.Text) Then
    holdItemWidth(1) = Printer.TextWidth(lstItem.Text)
  End If
  For x = 2 To UBound(holdItemWidth)
     If holdItemWidth(x) < Printer.TextWidth(lstItem.SubItems(x - 1)) Then
       holdItemWidth(x) = Printer.TextWidth(lstItem.SubItems(x - 1))
     End If
  Next x
  
Next lstItem

'Make sure that if there are no items to print in a column we set the width equal to the title of the column
'or if the width is not greater than or equal to the width of the column header


For x = 1 To UBound(holdItemWidth)
  If holdItemWidth(x) = 0 Or holdItemWidth(x) < Printer.TextWidth(mvarLvw.ColumnHeaders(x).Text) Then
    holdItemWidth(x) = Printer.TextWidth(mvarLvw.ColumnHeaders(x).Text)
  End If
Next x


'Determine if the total amount of space is wider than the page.  If so then rotate the output

DetermineWidth = 0
For x = 1 To UBound(holdItemWidth)
  DetermineWidth = DetermineWidth + holdItemWidth(x)
Next x

totalWidth = DetermineWidth
Printer.FontBold = False  'Turn off the bold



End Function

Private Function FitOutputToPage(fonttoprint As Integer) As Boolean

Dim widthOfOutput As Integer
Dim x As Integer


FitOutputToPage = False


For x = fonttoprint To 5 Step -1  'Move down all of the fonts until it fits on a page
  widthOfOutput = DetermineWidth(x)
  If widthOfOutput < Printer.Width Then
    FitOutputToPage = True
    Printer.FontSize = x
    Exit For
  End If
  
Next x




End Function

Private Sub PrintHeader(Optional pageNumber As Integer)


Dim headerLineNumber As Integer
Dim holdXPos As Integer
Dim colHeader As ColumnHeader
Dim cntHeader As Integer
Dim output As String
Dim holdY As Integer
Dim holdFontSize As Integer



Printer.FontBold = True
Printer.FontUnderline = True

headerLineNumber = Printer.CurrentY
For Each colHeader In mvarLvw.ColumnHeaders
  Printer.CurrentY = headerLineNumber
  Select Case cntHeader
    Case 0
      Printer.Print colHeader.Text
    Case Else
      Printer.currentX = holdXPos + holdItemWidth(cntHeader)
      holdXPos = Printer.currentX
      Printer.Print colHeader.Text
      
  End Select
  cntHeader = cntHeader + 1
Next colHeader

Printer.FontBold = False
Printer.FontUnderline = False

output = ""
Printer.Print output

'Print the footer on this page.  Be sure to print it at a location that will not force to the next page
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
Printer.currentX = (Printer.Width - Printer.TextWidth("Page -" & pageNumber & "-")) / 2
Printer.Print "-Page " & pageNumber & "-"
Printer.CurrentY = holdY
Printer.FontSize = holdFontSize
mvarFooterLocation = Printer.Height - (Printer.TextHeight("Footer") * 6)


End Sub

Private Sub PrintTitle(titleToPrint As String)

Dim holdFontSize As Integer
Dim startLoc As Single
Dim holdY As Integer
Dim titleHeight As Integer


holdFontSize = Printer.FontSize

holdY = Printer.CurrentY

Printer.FontSize = 16
Printer.FontBold = True
startLoc = (Printer.Width - Printer.TextWidth(titleToPrint)) / 2
titleHeight = Printer.TextHeight(titleToPrint)

Printer.DrawWidth = 20
Printer.Line (startLoc, Printer.CurrentY + titleHeight)-(startLoc + Printer.TextWidth(titleToPrint), Printer.CurrentY + titleHeight)
'Printer.Line (startLoc + Printer.TextWidth(titleToPrint), holdY)-(startLoc + Printer.TextWidth(titleToPrint), holdY + titleHeight)

Printer.currentX = (Printer.Width - Printer.TextWidth(titleToPrint)) / 2
Printer.CurrentY = holdY
Printer.Print titleToPrint
Printer.FontUnderline = False
Printer.Print " "

Printer.FontBold = False
Printer.FontSize = holdFontSize

End Sub

