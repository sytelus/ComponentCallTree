VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPrint 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Print"
   ClientHeight    =   4605
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4050
   Icon            =   "Print.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4605
   ScaleWidth      =   4050
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraTab 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   3105
      Index           =   3
      Left            =   180
      TabIndex        =   6
      Top             =   1080
      Width           =   4335
      Begin VB.CheckBox chkPrintList 
         Caption         =   "&Print List of All Childs"
         Height          =   225
         Left            =   120
         TabIndex        =   40
         Top             =   210
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.TextBox txtListFontSize 
         Height          =   285
         Left            =   900
         TabIndex        =   38
         Text            =   "8"
         Top             =   2610
         Width           =   420
      End
      Begin VB.Frame Frame8 
         Height          =   30
         Left            =   0
         TabIndex        =   36
         Top             =   2460
         Width           =   6015
      End
      Begin VB.Frame Frame1 
         Height          =   30
         Left            =   0
         TabIndex        =   34
         Top             =   600
         Width           =   6015
         Begin VB.Frame Frame7 
            Height          =   30
            Left            =   0
            TabIndex        =   35
            Top             =   -60
            Width           =   6015
         End
      End
      Begin VB.OptionButton optLandscap 
         Caption         =   "Print In &Landscap mode"
         Height          =   225
         Left            =   120
         TabIndex        =   33
         Top             =   1710
         Value           =   -1  'True
         Width           =   1995
      End
      Begin VB.OptionButton optPortrait 
         Caption         =   "Print in Port&rait mode"
         Height          =   225
         Left            =   120
         TabIndex        =   32
         Top             =   2070
         Width           =   1815
      End
      Begin VB.TextBox txtListTitle 
         Height          =   285
         Left            =   990
         TabIndex        =   30
         Text            =   "List Of All Child Calls"
         Top             =   780
         Width           =   2220
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Font Si&ze:"
         Height          =   195
         Left            =   120
         TabIndex        =   39
         Top             =   2640
         Width           =   705
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "&Orientation:"
         Height          =   195
         Left            =   150
         TabIndex        =   37
         Top             =   1350
         Width           =   810
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "List &Title:"
         Height          =   195
         Left            =   120
         TabIndex        =   31
         Top             =   810
         Width           =   630
      End
   End
   Begin VB.Frame fraTab 
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      Height          =   2895
      Index           =   2
      Left            =   180
      TabIndex        =   5
      Top             =   750
      Width           =   4335
      Begin VB.OptionButton optPrintWholeTree 
         Caption         =   "&Whole Tree"
         Height          =   225
         Left            =   2550
         TabIndex        =   29
         Top             =   1320
         Width           =   1155
      End
      Begin VB.OptionButton optPrintSelectedNode 
         Caption         =   "Selected &Node"
         Height          =   225
         Left            =   120
         TabIndex        =   28
         Top             =   1320
         Value           =   -1  'True
         Width           =   1425
      End
      Begin VB.Frame Frame6 
         Height          =   30
         Left            =   -1050
         TabIndex        =   27
         Top             =   1650
         Width           =   6735
      End
      Begin VB.Frame Frame5 
         Height          =   30
         Left            =   -810
         TabIndex        =   26
         Top             =   600
         Width           =   5955
      End
      Begin VB.CheckBox chkNodeFontBold 
         Caption         =   "&Bold Font"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   1800
         Width           =   1695
      End
      Begin VB.TextBox txtNodeFontSize 
         Height          =   285
         Left            =   3300
         TabIndex        =   23
         Text            =   "8"
         Top             =   1830
         Width           =   420
      End
      Begin VB.CheckBox chkPrintTree 
         Caption         =   "Print &Tree Structure"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   210
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.ComboBox cmbPrintStyle 
         Height          =   315
         ItemData        =   "Print.frx":0442
         Left            =   1350
         List            =   "Print.frx":044F
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   810
         Width           =   2385
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Font Si&ze:"
         Height          =   195
         Left            =   2520
         TabIndex        =   25
         Top             =   1860
         Width           =   705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Connector S&tyle:"
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   810
         Width           =   1170
      End
   End
   Begin VB.Frame fraTab 
      BackColor       =   &H008080FF&
      BorderStyle     =   0  'None
      Height          =   3195
      Index           =   1
      Left            =   90
      TabIndex        =   4
      Top             =   540
      Width           =   3885
      Begin VB.Frame Frame4 
         Height          =   30
         Left            =   -570
         TabIndex        =   19
         Top             =   2400
         Width           =   6435
      End
      Begin VB.CheckBox chkPrintFooter 
         Caption         =   "Print &Footer"
         Height          =   285
         Left            =   60
         TabIndex        =   18
         Top             =   2640
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.Frame Frame3 
         Height          =   30
         Left            =   -780
         TabIndex        =   17
         Top             =   1140
         Width           =   6015
      End
      Begin VB.TextBox txtPageTitle 
         Height          =   285
         Left            =   960
         TabIndex        =   12
         Text            =   "VB Project Reference Analysis"
         Top             =   210
         Width           =   2220
      End
      Begin VB.CheckBox chkTitle1FontBold 
         Caption         =   "Bold Font For Title"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Value           =   1  'Checked
         Width           =   1605
      End
      Begin VB.TextBox txtTitle1FontSize 
         Height          =   285
         Left            =   3330
         TabIndex        =   10
         Text            =   "14"
         Top             =   720
         Width           =   420
      End
      Begin VB.TextBox txtPageTitle2 
         Height          =   285
         Left            =   990
         TabIndex        =   9
         Text            =   "Reference Call Tree"
         Top             =   1350
         Width           =   2220
      End
      Begin VB.CheckBox chkTitle2FontBold 
         Caption         =   "Bold Font For Title2"
         Height          =   255
         Left            =   60
         TabIndex        =   8
         Top             =   1920
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.TextBox txtTitle2FontSize 
         Height          =   285
         Left            =   3270
         TabIndex        =   7
         Text            =   "12"
         Top             =   1950
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Page &Title:"
         Height          =   195
         Left            =   90
         TabIndex        =   16
         Top             =   240
         Width           =   765
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Title Font Size:"
         Height          =   195
         Left            =   2220
         TabIndex        =   15
         Top             =   750
         Width           =   1050
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Page Title&2:"
         Height          =   195
         Left            =   60
         TabIndex        =   14
         Top             =   1380
         Width           =   855
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Title2 Font Size:"
         Height          =   195
         Left            =   2070
         TabIndex        =   13
         Top             =   1980
         Width           =   1140
      End
   End
   Begin MSComctlLib.TabStrip tbsPrintOptions 
      Height          =   3795
      Left            =   30
      TabIndex        =   3
      Top             =   120
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   6694
      MultiRow        =   -1  'True
      HotTracking     =   -1  'True
      Separators      =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&General"
            Object.ToolTipText     =   "General Print Options"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Tree"
            Object.ToolTipText     =   "Tree Print Options"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&List"
            Object.ToolTipText     =   "All Child List Print Options"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      Height          =   30
      Left            =   -60
      TabIndex        =   2
      Top             =   3960
      Width           =   4785
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   405
      Left            =   2790
      TabIndex        =   1
      Top             =   4110
      Width           =   1245
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   405
      Left            =   1410
      TabIndex        =   0
      Top             =   4110
      Width           =   1245
   End
End
Attribute VB_Name = "frmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mbIsOKPressed As Boolean
Dim WithEvents moTreeViewPrint As CPrintTvw
Attribute moTreeViewPrint.VB_VarHelpID = -1

Public Function DisplayForm(Optional ByVal voTreeView As TreeView = Nothing, Optional ByVal voListView As ListView = Nothing) As Boolean
    Dim bFormValidated As Boolean
    Dim sMessage As String
    
    bFormValidated = False
    
    Do While Not bFormValidated
    
        mbIsOKPressed = False
        
        Me.Show vbModal
        
        If mbIsOKPressed Then
            bFormValidated = ValidateForm(sMessage)
            If bFormValidated Then
                Call PrintTreeAndList(voTreeView, voListView)
            Else
                MsgBox sMessage
            End If
        Else
            bFormValidated = True
        End If
    Loop
        
    DisplayForm = mbIsOKPressed
    
    Unload Me
End Function

Private Function ValidateForm(ByRef rsMessage As String) As Boolean
    Dim bReturn As Boolean
    rsMessage = vbNullString
    bReturn = True
    
    If Not IsNumber(txtListFontSize) Then
        bReturn = False
        rsMessage = rsMessage & "Font Size for List field must be a number" & vbCrLf
    End If
    
    If Not IsNumber(txtNodeFontSize) Then
        bReturn = False
        rsMessage = rsMessage & "Font Size for a tree node must be a number" & vbCrLf
    End If
    
    If Not IsNumber(txtTitle1FontSize) Then
        bReturn = False
        rsMessage = rsMessage & "Font Size for page title must be a number" & vbCrLf
    End If
    
    If Not IsNumber(txtTitle2FontSize) Then
        bReturn = False
        rsMessage = rsMessage & "Font Size for page title2 must be a number" & vbCrLf
    End If
    
    If Not bReturn Then
        rsMessage = "Following is not correct: " & vbCrLf & rsMessage
    End If
    
    ValidateForm = bReturn
    
End Function

Private Function IsTextBoxEmpty(ByVal voTextBox As TextBox) As Boolean
    If Trim$(voTextBox.Text) = vbNullString Then
        IsTextBoxEmpty = True
    Else
        IsTextBoxEmpty = False
    End If
End Function


Private Sub PrintTreeAndList(Optional ByVal voTreeView As TreeView = Nothing, Optional ByVal voListView As ListView = Nothing)
    
    Screen.MousePointer = vbHourglass
    
    If Not (voTreeView Is Nothing) Then
        Call PrintTree(voTreeView)
    End If
    
    If Not (voListView Is Nothing) Then
        Call PrintList(voListView)
    End If
    
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub PrintTree(ByVal voTreeView As TreeView)
    If chkPrintTree.Value = vbChecked Then
        Set moTreeViewPrint = New CPrintTvw
        With moTreeViewPrint
          .FontBold = CheckToBool(chkNodeFontBold.Value)
          .FontSize = txtNodeFontSize.Text
          .PrintFooter = CheckToBool(chkPrintFooter.Value)
          .PrintStyle = IIf(optPrintWholeTree.Value, ePrintAll, ePrintFromSelected)
          .SecondTitle = txtPageTitle2.Text
          .SecondTitleFontBold = CheckToBool(chkTitle2FontBold.Value)
          .SecondTitleFontSize = txtTitle2FontSize.Text
          .title = txtPageTitle.Text
          .TitleFontBold = CheckToBool(chkTitle1FontBold.Value)
          .TitleFontSize = txtTitle2FontSize.Text
          Select Case cmbPrintStyle.ItemData(cmbPrintStyle.ListIndex)
            Case 1
                .ConnectorStyle = econnectlines
            Case 2
                .ConnectorStyle = eNoConnectLinesWithIndent
            Case 3
                .ConnectorStyle = eNoConnectLinesNoIndent
            End Select
        End With
        
        Set moTreeViewPrint.tvwToPrint = voTreeView
        Call moTreeViewPrint.PrintTvw
        
        Set moTreeViewPrint = Nothing
    End If
End Sub

Private Sub PrintList(ByVal voListView As ListView)
    If chkPrintList.Value = vbChecked Then
        Dim oListViewPrint As CLvwPrint
        Set oListViewPrint = New CLvwPrint
            With oListViewPrint
                Set .lvwToPrint = voListView
                Call .PrintListView(IIf(optLandscap.Value, iLandscape, iPortrait), , txtListTitle.Text, txtListFontSize.Text)
            End With
        Set oListViewPrint = Nothing
    End If
End Sub

Private Sub cmdCancel_Click()
    mbIsOKPressed = False
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    mbIsOKPressed = True
    Me.Hide
End Sub

Private Sub Form_Load()
    mbIsOKPressed = False
    
    Dim ofraTab As Frame
    For Each ofraTab In fraTab
        ofraTab.Visible = False
        ofraTab.BackColor = vbButtonFace
        ofraTab.Left = tbsPrintOptions.ClientLeft
        ofraTab.Top = tbsPrintOptions.ClientTop
        ofraTab.Width = tbsPrintOptions.ClientWidth
        ofraTab.Height = tbsPrintOptions.ClientHeight
    Next ofraTab
    Call tbsPrintOptions_Click

    cmbPrintStyle.ListIndex = 0
End Sub

Private Sub moTreeViewPrint_PrepareNode(ByVal Node As MSComctlLib.Node)
    Call frmMain.FillTreeNode(Node)
End Sub

Private Sub tbsPrintOptions_Click()
    Dim ofraTab As Frame
    For Each ofraTab In fraTab
        If ofraTab.Index = tbsPrintOptions.SelectedItem.Index Then
            ofraTab.Visible = True
        Else
            ofraTab.Visible = False
        End If
    Next ofraTab
End Sub

