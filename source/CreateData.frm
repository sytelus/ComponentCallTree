VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmCreateData 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Create Call Tree Data"
   ClientHeight    =   2280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4200
   Icon            =   "CreateData.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   4200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkIncludeSubFolders 
      Caption         =   "&Include Subfolders"
      Height          =   315
      Left            =   2520
      TabIndex        =   10
      Top             =   1200
      Value           =   1  'Checked
      Width           =   1635
   End
   Begin MSComDlg.CommonDialog cdlFileOpen 
      Left            =   180
      Top             =   1800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   ".mdb"
      DialogTitle     =   "Locate Database File"
      Filter          =   "Access Database (*.mdb)|*.mdb|All Files|*.*"
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1890
      TabIndex        =   9
      Top             =   1770
      Width           =   1065
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3090
      TabIndex        =   8
      Top             =   1770
      Width           =   1065
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Left            =   -300
      TabIndex        =   7
      Top             =   1500
      Width           =   4875
   End
   Begin VB.CommandButton cmdDatabaseFileBrose 
      Height          =   315
      Left            =   3750
      Picture         =   "CreateData.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   660
      Width           =   375
   End
   Begin VB.CheckBox chkDeleteOldData 
      Caption         =   "De&lete old data"
      Height          =   315
      Left            =   90
      TabIndex        =   5
      Top             =   1200
      Value           =   1  'Checked
      Width           =   1545
   End
   Begin VB.TextBox txtDatabaseFile 
      Height          =   315
      Left            =   1590
      TabIndex        =   4
      Top             =   660
      Width           =   2115
   End
   Begin VB.TextBox txtVBPLocation 
      Height          =   315
      Left            =   1590
      TabIndex        =   2
      Top             =   210
      Width           =   2115
   End
   Begin VB.CommandButton cmdVBPLocationBrowse 
      Height          =   315
      Left            =   3750
      Picture         =   "CreateData.frx":0544
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   210
      Width           =   375
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "&Database Location:"
      Height          =   195
      Left            =   90
      TabIndex        =   3
      Top             =   720
      Width           =   1395
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "&VBP files location:"
      Height          =   195
      Left            =   90
      TabIndex        =   0
      Top             =   240
      Width           =   1275
   End
End
Attribute VB_Name = "frmCreateData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mbIsOKPressed As Boolean

Private Sub cmdCancel_Click()
    mbIsOKPressed = False
    Me.Hide
End Sub

Private Sub cmdDatabaseFileBrose_Click()
    On Error GoTo ErrorTrap
    cdlFileOpen.Flags = cdlOFNFileMustExist Or cdlOFNPathMustExist
    cdlFileOpen.fileName = txtDatabaseFile.Text
    If txtDatabaseFile.Text = vbNullString Then
        cdlFileOpen.InitDir = App.Path
    End If
    cdlFileOpen.ShowOpen
    txtDatabaseFile.Text = cdlFileOpen.fileName
Exit Sub
ErrorTrap:
    If Err.Number <> cdlCancel Then
        ShowError
    End If
End Sub

Private Sub cmdOK_Click()
    mbIsOKPressed = True
    Me.Hide
End Sub

Public Function DisplayForm(ByRef rsVBPLocation As String, ByRef rsDatabaseFile As String, ByRef rbDeleteOldData As Boolean, ByRef rbIncludeSubFolders As Boolean) As Boolean
    Dim bFormValidated As Boolean
    Dim sMessage As String
    
    'Data xfer
    txtVBPLocation.Text = rsVBPLocation
    txtDatabaseFile.Text = rsDatabaseFile
    chkDeleteOldData.Value = BoolToCheck(rbDeleteOldData)
    chkIncludeSubFolders.Value = BoolToCheck(rbIncludeSubFolders)
    
    bFormValidated = False
    
    Do While Not bFormValidated
    
        mbIsOKPressed = False
        
        Me.Show vbModal
        
        If mbIsOKPressed Then
            bFormValidated = ValidateForm(sMessage)
            If bFormValidated Then
                rsVBPLocation = Trim(txtVBPLocation.Text)   'Trim is needed because Dir$ doesn't work otherwise
                rsDatabaseFile = Trim(txtDatabaseFile.Text)
                rbDeleteOldData = CheckToBool(chkDeleteOldData.Value)
                rbIncludeSubFolders = CheckToBool(chkIncludeSubFolders.Value)
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
    
    If IsTextBoxEmpty(txtDatabaseFile) Then
        bReturn = False
        rsMessage = rsMessage & "Database file name can't be blank" & vbCrLf
    End If
    
    If IsTextBoxEmpty(txtVBPLocation) Then
        bReturn = False
        rsMessage = rsMessage & "VBP files location can't be blank" & vbCrLf
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

Private Sub cmdVBPLocationBrowse_Click()
    txtVBPLocation.Text = BrowseForFolder(txtVBPLocation.Text, Me.hwnd, "Select Folder Where VBP Files Are Located")
End Sub

Private Sub ShowError()
    Me.MousePointer = vbDefault
    Screen.MousePointer = vbDefault
    MsgBox "Error " & Err.Number & " : " & Err.Description
End Sub

