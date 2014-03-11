VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmIconExplanation 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Icon Explanation"
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6315
   Icon            =   "IconExplanation.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   6315
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   135
      Left            =   0
      TabIndex        =   3
      Top             =   3840
      Width           =   6345
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   405
      Left            =   5040
      TabIndex        =   2
      Top             =   4170
      Width           =   1245
   End
   Begin MSComctlLib.ImageList imlIconsForExplanation 
      Left            =   2070
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "IconExplanation.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "IconExplanation.frx":0894
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "IconExplanation.frx":0CE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "IconExplanation.frx":1138
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "IconExplanation.frx":158A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "IconExplanation.frx":19DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "IconExplanation.frx":212E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lsvIconHelp 
      Height          =   3675
      Left            =   90
      TabIndex        =   0
      Top             =   150
      Width           =   6195
      _ExtentX        =   10927
      _ExtentY        =   6482
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "imlIconsForExplanation"
      SmallIcons      =   "imlIconsForExplanation"
      ForeColor       =   -2147483640
      BackColor       =   -2147483633
      Appearance      =   0
      Enabled         =   0   'False
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   17639
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "&icon Help:"
      Height          =   225
      Left            =   840
      TabIndex        =   1
      Top             =   60
      Visible         =   0   'False
      Width           =   1515
   End
End
Attribute VB_Name = "frmIconExplanation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    lsvIconHelp.ColumnHeaders(1).Width = lsvIconHelp.Width
    
    lsvIconHelp.ListItems.Clear
    lsvIconHelp.ListItems.Add(, , "    Unexpanded Component Node", 1, 1).Selected = False
    Call lsvIconHelp.ListItems.Add(, , "    Expanded Component Node", 2, 2)
    Call lsvIconHelp.ListItems.Add(, , "    Call From - To View", 3, 3)
    Call lsvIconHelp.ListItems.Add(, , "    Call To - From View", 4, 4)
    Call lsvIconHelp.ListItems.Add(, , "    Child component is referenced with different versions/descriptions", 5, 5)
    Call lsvIconHelp.ListItems.Add(, , "    Component is ActiveX Control (OCX)", 6, 6)
    Call lsvIconHelp.ListItems.Add(, , "    Component is standerd Win32 EXE", 7, 7)
End Sub
