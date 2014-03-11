VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmLoadData 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Load Call Tree Data"
   ClientHeight    =   1335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4200
   Icon            =   "LoadData.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1335
   ScaleWidth      =   4200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1890
      TabIndex        =   5
      Top             =   840
      Width           =   1065
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3090
      TabIndex        =   4
      Top             =   840
      Width           =   1065
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Left            =   -300
      TabIndex        =   3
      Top             =   570
      Width           =   4875
   End
   Begin VB.CommandButton cmdDatabaseFileBrose 
      Height          =   315
      Left            =   3750
      Picture         =   "LoadData.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   150
      Width           =   375
   End
   Begin VB.TextBox txtDatabaseFile 
      Height          =   315
      Left            =   1590
      TabIndex        =   1
      Top             =   150
      Width           =   2115
   End
   Begin MSComDlg.CommonDialog cdlFileOpen 
      Left            =   150
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   ".mdb"
      DialogTitle     =   "Locate Database File"
      Filter          =   "Access Database (*.mdb)|*.mdb|All Files|*.*"
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "&Database Location:"
      Height          =   195
      Left            =   90
      TabIndex        =   0
      Top             =   210
      Width           =   1395
   End
End
Attribute VB_Name = "frmLoadData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mbIsOkPressed As Boolean

Private Sub cmdCancel_Click()
    mbIsOkPressed = False
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    mbIsOkPressed = True
    Me.Hide
End Sub

Public Function DisplayForm(ByRef rsDatabaseFile As String) As Boolean
    Dim bFormValidated As Boolean
    Dim sMessage As String
    
    'Data xfer
    txtDatabaseFile.Text = rsDatabaseFile
    
    bFormValidated = False
    
    Do While Not bFormValidated
    
        mbIsOkPressed = False
        
        Me.Show vbModal
        
        If mbIsOkPressed Then
            bFormValidated = ValidateForm(sMessage)
            If bFormValidated Then
                rsDatabaseFile = txtDatabaseFile.Text
            Else
                MsgBox sMessage
            End If
        Else
            bFormValidated = True
        End If
    Loop
        
    DisplayForm = mbIsOkPressed
    
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
    
    ValidateForm = bReturn
    
End Function

Private Function IsTextBoxEmpty(ByVal voTextBox As TextBox) As Boolean
    If Trim$(voTextBox.Text) = vbNullString Then
        IsTextBoxEmpty = True
    Else
        IsTextBoxEmpty = False
    End If
End Function

Private Sub cmdDatabaseFileBrose_Click()
    On Error GoTo ErrorTrap
    cdlFileOpen.Flags = cdlOFNFileMustExist Or cdlOFNPathMustExist
    cdlFileOpen.FileName = txtDatabaseFile.Text
    If txtDatabaseFile.Text = vbNullString Then
        cdlFileOpen.InitDir = App.Path
    End If
    cdlFileOpen.ShowOpen
    txtDatabaseFile.Text = cdlFileOpen.FileName
Exit Sub
ErrorTrap:
    If Err.Number <> cdlCancel Then
        ShowError
    End If
End Sub

Private Sub ShowError()
    Me.MousePointer = vbDefault
    Screen.MousePointer = vbDefault
    MsgBox "Error " & Err.Number & " : " & Err.Description
End Sub

