VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Reference Explorer"
   ClientHeight    =   5625
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9780
   Icon            =   "main.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5625
   ScaleWidth      =   9780
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cdlPrinterSetup 
      Left            =   4650
      Top             =   2580
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picSplitter 
      BorderStyle     =   0  'None
      Height          =   4545
      Left            =   3750
      MousePointer    =   9  'Size W E
      ScaleHeight     =   4545
      ScaleWidth      =   45
      TabIndex        =   6
      Top             =   720
      Width           =   45
   End
   Begin MSComctlLib.Toolbar tlbMain 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   9780
      _ExtentX        =   17251
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "imgList"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "NEW"
            Object.ToolTipText     =   "Create Call Tree Data"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "OPEN"
            Object.ToolTipText     =   "Open Call Tree Data"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "PRINT"
            Object.ToolTipText     =   "Print Tree And List"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "FROM-TO"
            Object.ToolTipText     =   "View From-To Call Tree"
            ImageIndex      =   4
            Style           =   2
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "TO-FROM"
            Object.ToolTipText     =   "View To-From Call Tree"
            ImageIndex      =   7
            Style           =   2
            Value           =   1
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "HELP"
            Object.ToolTipText     =   "About Reference Explorer"
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSComctlLib.StatusBar stbMain 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   5250
      Width           =   9780
      _ExtentX        =   17251
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
            Text            =   "Ready"
            TextSave        =   "Ready"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11438
            MinWidth        =   8819
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lsvCallList 
      Height          =   4545
      Left            =   3870
      TabIndex        =   1
      Top             =   690
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   8017
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "imgList"
      SmallIcons      =   "imgList"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "!"
         Object.Width           =   617
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Component"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Version"
         Object.Width           =   2999
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Description"
         Object.Width           =   17639
      EndProperty
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   5670
      Top             =   90
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":0894
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":0CE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":1138
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":158A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":19DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":212E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":2580
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":29D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":2AE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":2BF6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView trwCallTree 
      Height          =   4545
      Left            =   30
      TabIndex        =   0
      Top             =   690
      Width           =   3705
      _ExtentX        =   6535
      _ExtentY        =   8017
      _Version        =   393217
      Indentation     =   71
      LabelEdit       =   1
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      ImageList       =   "imgList"
      Appearance      =   1
   End
   Begin VB.Label lblListView 
      AutoSize        =   -1  'True
      Caption         =   "&List of All Child Calls:"
      Height          =   195
      Left            =   3780
      TabIndex        =   5
      Top             =   450
      Width           =   1440
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Call &Tre:"
      Height          =   195
      Left            =   60
      TabIndex        =   4
      Top             =   450
      Width           =   585
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileCreateData 
         Caption         =   "&Create Data..."
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileLoadData 
         Caption         =   "&Load Data..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileBreak1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePrinterSetup 
         Caption         =   "P&rinter Setup..."
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "&Print..."
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileBreak2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuFromToTree 
         Caption         =   "Call &From-To Tree"
         Checked         =   -1  'True
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuToFromTree 
         Caption         =   "Call &To-From Tree"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuViewBreak1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewShowStandardReferences 
         Caption         =   "Show &Standard References"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpIconInformation 
         Caption         =   "&Icon Information..."
      End
      Begin VB.Menu mnuHelpBreak1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim moclMaster As Collection
Dim mbCallToFromTree As Boolean
Dim msROOT_NODE_KEY As String
Dim msDatabaseFile As String
Dim mbShowStandardReferences As Boolean
    
Const msREG_APP_NAME As String = "Reference Explorer"
Const msREG_SECTION_SETTINGS As String = "Settings"
Const msREG_SECTION_LAST_VALUES As String = "Last Values"
Const msREG_KEY_VBP_LOCATION As String = "VBPLocation"
Const msREG_KEY_DATABASE_FILE As String = "DatabaseFile"
Const msREG_KEY_SHOW_STANDARD_REFERENCES As String = "ShowStdReferences"

Private Type udtFormResizeInfo
    TreeViewLeft As Long
End Type

Dim udFormResizeInfo As udtFormResizeInfo
Dim mbSplitStarted As Boolean

Private Sub Form_Load()

    On Error GoTo ErrorTrap


    mbCallToFromTree = False
    Call RefreshFromToViewStatus
    msDatabaseFile = vbNullString
    mbShowStandardReferences = GetSetting(msREG_APP_NAME, msREG_SECTION_SETTINGS, msREG_KEY_SHOW_STANDARD_REFERENCES, "False")
    mnuViewShowStandardReferences.Checked = mbShowStandardReferences
    
    udFormResizeInfo.TreeViewLeft = trwCallTree.Left
    mbSplitStarted = False
    
Exit Sub
ErrorTrap:
    ShowError
End Sub

Private Sub InitTree()
    Screen.MousePointer = vbHourglass
    Set moclMaster = Nothing
    If mbCallToFromTree Then
        msROOT_NODE_KEY = "Call To-From Tree"
    Else
        msROOT_NODE_KEY = "Call From-To Tree"
    End If
    
    Call BuildMasterColFromDB(GetConnectionString)
    
    Call UpdateStatus(, "Loading tree...")
    trwCallTree.Nodes.Clear
    lsvCallList.ListItems.Clear
    
    Dim oRootNode As Node
    Set oRootNode = trwCallTree.Nodes.Add(, , msROOT_NODE_KEY, msROOT_NODE_KEY, IIf(mbCallToFromTree, 7, 4), 0)
    oRootNode.Tag = 0   'Not filled
    Call FillTreeNode(oRootNode)
    oRootNode.Expanded = True
    Screen.MousePointer = vbDefault
    Call UpdateStatus(, vbNullString)
    
End Sub

Friend Sub FillTreeNode(ByVal voNode As Node)
    If voNode.Tag = 0 Then 'Not filled
        
        Screen.MousePointer = vbHourglass
        
        Dim oclCallTo As Collection
        Dim sCallTo As Variant
        Dim oNode As Node
        Dim oclChildNodes As Collection
        
        Call SilentlyRemoveTempNode(voNode)
        Set oclCallTo = GetObjectInCollection(moclMaster, (LCase(voNode.Text)))
        If Not (oclCallTo Is Nothing) Then
            Call SortCollection(oclCallTo)
            For Each sCallTo In oclCallTo
                Set oNode = trwCallTree.Nodes.Add(voNode, tvwChild, , sCallTo, 1, 2)
                oNode.Tag = 0
                Set oclCallTo = GetObjectInCollection(moclMaster, (LCase(sCallTo)))
                If Not (oclCallTo Is Nothing) Then
                    If oclCallTo.Count <> 0 Then
                        'Add dummy node
                        Call trwCallTree.Nodes.Add(oNode, tvwChild, , "")
                    Else
                        'Indicate that there are no childs
'                        oNode.Image = 4
'                        oNode.SelectedImage = 4
'                        oNode.ExpandedImage = 4
                    End If
                Else
                    'Indicate that there are no childs
'                    oNode.Image = 4
'                    oNode.SelectedImage = 4
'                    oNode.ExpandedImage = 4
                End If
            Next sCallTo
            Set oclCallTo = Nothing
        End If
        voNode.Tag = 1 'Filled
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub SilentlyRemoveTempNode(ByVal voNode As Node)
    On Error Resume Next
    Call trwCallTree.Nodes.Remove(voNode.Child.Index)
End Sub

Private Sub BuildMasterColFromDB(ByVal vsConnectionString As String)

    Dim oConn As ADODB.Connection
    Dim oRs As ADODB.Recordset
    Dim sCaller As String
    Dim sCallTo As String
    Dim oclCallTo As Collection
        
    Set oConn = New ADODB.Connection
    Set oRs = New ADODB.Recordset
    
    
    Call UpdateStatus(, "Opening database connection...")
    Call oConn.Open(vsConnectionString)
    
    Dim sSQL As String
    sSQL = "SELECT * FROM ComponentCalls"
    If mbCallToFromTree Then
        sSQL = sSQL & " ORDER BY CallTo"
    Else
        sSQL = sSQL & " ORDER BY Caller"
    End If
    
        Call UpdateStatus(, "Fetching data...")
        Call oRs.Open(sSQL, oConn, adOpenStatic, adLockReadOnly)
        
        Call UpdateStatus(, "Analysing data...")
        Dim oclRoot As Collection
        Dim sCallerWithCase As String
        Set oclRoot = New Collection
        Set moclMaster = New Collection
        If (Not oRs.EOF) And (Not oRs.BOF) Then
            oRs.MoveFirst
            Do While Not oRs.EOF
                
                If mbCallToFromTree Then
                    sCaller = LCase(oRs!callto)
                    sCallerWithCase = oRs!callto
                    sCallTo = oRs!Caller
                Else
                    sCallerWithCase = oRs!Caller
                    sCaller = LCase(oRs!Caller)
                    sCallTo = oRs!callto
                End If
                
                'If NOT (std ref is not be shown AND if caller or callto is stdole2) then
                If Not ((Not mbShowStandardReferences) And _
                    ((StrComp(sCaller, "stdole2", vbTextCompare) = 0) _
                    Or (StrComp(sCallTo, "stdole2", vbTextCompare) = 0))) Then
                
                    Set oclCallTo = GetObjectInCollection(moclMaster, sCaller)
                    If oclCallTo Is Nothing Then
                        Set oclCallTo = New Collection
                        Call moclMaster.Add(oclCallTo, sCaller)
                        Call oclRoot.Add(sCallerWithCase)
                    End If
                    Call SafeAddToCollection(oclCallTo, sCallTo, LCase(sCallTo))    'Key is used just to ensure that there is no duplicates
                    Set oclCallTo = Nothing
                    
                End If
                oRs.MoveNext
            Loop
        End If
        Call moclMaster.Add(oclRoot, LCase(msROOT_NODE_KEY))
    oRs.Close
    oConn.Close
    
    Set oConn = Nothing
    Set oRs = Nothing
    
End Sub

Private Function GetObjectInCollection(ByVal voclCol As Collection, ByVal vsKey As String) As Object
    On Error GoTo ErrorTrap
    Set GetObjectInCollection = voclCol(CStr(vsKey))
Exit Function
ErrorTrap:
    Set GetObjectInCollection = Nothing
End Function

Private Sub CreateCallTreeData(ByVal vsVBPLocation As String, ByVal vsConnectionString As String, ByVal vboolDeleteOldData As String, ByVal vboolIncludeSubFolders As Boolean)

    On Error GoTo ErrorTrap

    'For each VBP file, get referenced components
    
    Call UpdateStatus("Creating data...")
    
    Dim oConn As ADODB.Connection
    Dim oRs As ADODB.Recordset
    Dim sSQL As String
    Dim bTransactionStarted As Boolean
    
    bTransactionStarted = False
    
    Set oConn = New ADODB.Connection
    Set oRs = New ADODB.Recordset
    
    Call oConn.Open(vsConnectionString)
    oConn.CursorLocation = adUseClient  'This req due to Access OLE-DB provider bug. Otherwise, sometimes in very rare case, unpredictably, it gives almost untraceable error "Error occured"
    Call oConn.BeginTrans
    bTransactionStarted = True
    
    If vboolDeleteOldData Then
        Call oConn.Execute("DELETE FROM ComponentCalls")
    End If
    Call oRs.Open("SELECT * FROM ComponentCalls WHERE 1=2", oConn, adOpenKeyset, adLockOptimistic)
    
    Call AddVBPFilesInDirToDB(vsVBPLocation, oRs, vboolIncludeSubFolders)
    
    oRs.Close
    Set oRs = Nothing
    
    Call oConn.CommitTrans
    bTransactionStarted = False
    oConn.Close
    
    Set oConn = Nothing
        
    Call UpdateStatus("Data Created", vbNullString)

Exit Sub
ErrorTrap:
    If Not (oConn Is Nothing) Then
        If bTransactionStarted Then
            oConn.RollbackTrans
        End If
    End If
    ReRaiseError
End Sub

Private Sub AddVBPFilesInDirToDB(ByVal vsPath As String, ByVal voRS As ADODB.Recordset, ByVal vboolIncludeSubFolders As Boolean)
    
    Dim sDirToSearch As String
    Dim sVBPFileName As String
    Dim oclFileNames As Collection
    
    Call UpdateStatus(, "Searching VBP files...")
    
    'Add file that is in this dir only
    Set oclFileNames = New Collection
    sDirToSearch = GetPathWithSlash(vsPath)
    sVBPFileName = Dir$(sDirToSearch & "*.vbp")
    Do While sVBPFileName <> vbNullString
        Call oclFileNames.Add(sVBPFileName)
        sVBPFileName = Dir$
    Loop
    Call AddVBPFilesColToDB(sDirToSearch, oclFileNames, voRS)
    Set oclFileNames = Nothing

    
    If vboolIncludeSubFolders Then
        'Now for each dir inside this dir, do same
        Dim sChildDir As String
        Dim oclDirNames As Collection
        
        Set oclDirNames = New Collection
        
        sChildDir = Dir$(sDirToSearch & "*.*", vbDirectory)
        Do While sChildDir <> vbNullString
        
            ' Ignore the current directory and the encompassing directory.
            If sChildDir <> "." And sChildDir <> ".." Then
            
                ' Use bitwise comparison to make sure sChildDir is a directory.
                If IsDir(sDirToSearch & sChildDir) Then
                
                    Call oclDirNames.Add(sChildDir)
                    
                End If
                
            End If
            
            sChildDir = Dir$    ' Get next entry.
        Loop
        
        Dim sDirNameInCol As Variant
        For Each sDirNameInCol In oclDirNames
            Call AddVBPFilesInDirToDB(sDirToSearch & sDirNameInCol, voRS, vboolIncludeSubFolders)
        Next sDirNameInCol
        Set oclDirNames = Nothing
    End If

End Sub

Private Sub AddVBPFilesColToDB(ByVal vsPathWithSlash As String, ByVal voclFileNames As Collection, ByVal voRS As ADODB.Recordset)

    'On Error GoTo ErrorTrap

    Dim oVBPFile As VBPFile
    Dim sVBPFileName As Variant
    Dim oVBPRefrence As VBPReference
    Dim sDLLFileNameWithoutExt As String
    Dim sCallerVersion As String
    
    For Each sVBPFileName In voclFileNames
        Call UpdateStatus(, "Now scanning " & sVBPFileName)
        Set oVBPFile = New VBPFile
        Call oVBPFile.OpenVBP(vsPathWithSlash & sVBPFileName)
        For Each oVBPRefrence In oVBPFile.VBPRefernces
                If (StrComp(Right(oVBPRefrence.DLLName, 4), ".dll", vbTextCompare) = 0) Or (StrComp(Right(oVBPRefrence.DLLName, 4), ".tlb", vbTextCompare) = 0) Then
                    sDLLFileNameWithoutExt = Left(oVBPRefrence.DLLName, Len(oVBPRefrence.DLLName) - 4)
                Else
                    sDLLFileNameWithoutExt = oVBPRefrence.DLLName
                End If
                
                With voRS
                    .AddNew
                        !Caller = oVBPFile.ProjectName
                        !callto = sDLLFileNameWithoutExt
                        !CallToDLLName = Left(oVBPRefrence.DLLName, 250)
                        !CallToVersion = oVBPRefrence.Version
                        !CallToDescription = Left(oVBPRefrence.Description, 250)
                        !IsCallToOCX = oVBPRefrence.IsOCX
                        sCallerVersion = oVBPFile.MajorVersion & "." & oVBPFile.MinorVersion & "." & oVBPFile.BuildNumber
                        If sCallerVersion = ".." Then
                            sCallerVersion = vbNullString
                        End If
                        !CallerVersion = sCallerVersion
                        !CallerDescription = oVBPFile.ProjectDescription
                        !CallerType = oVBPFile.VBPType  'OleDll,Exe,Control
                    .Update
                End With
        Next oVBPRefrence
        Set oVBPFile = Nothing
    Next sVBPFileName
    
Exit Sub
ErrorTrap:
    Dim oConn As ADODB.Connection
    Set oConn = voRS.ActiveConnection
    'MsgBox oConn.Errors.Count
    MsgBox oConn.Errors.Item(1).Description
End Sub

Private Function GetConnectionString() As String
    If msDatabaseFile = vbNullString Then
        Err.Raise 1000, , "Database file not specified"
    End If
    GetConnectionString = "provider = microsoft.jet.oledb.4.0 ; data source=" & msDatabaseFile
End Function

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If mbSplitStarted Then
        If (x > 150) And (x < (Me.Width - 150 - udFormResizeInfo.TreeViewLeft - picSplitter.Width)) Then
            picSplitter.Left = x
        End If
    End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If mbSplitStarted Then
        mbSplitStarted = False
        Call ReleaseCapture
        Call Form_Resize
    End If
End Sub

Private Sub Form_Resize()
    
    'Resize treeview and listview heights
    Dim lNewHeight As Long
    Dim lNewWidth As Long

    lNewHeight = frmMain.Height - stbMain.Height * 2.7 - trwCallTree.Top
    If lNewHeight > 0 Then
        trwCallTree.Height = lNewHeight
    End If
    lsvCallList.Height = trwCallTree.Height
    picSplitter.Height = trwCallTree.Height
    
    lNewWidth = picSplitter.Left - trwCallTree.Left
    If lNewWidth > 0 Then
        trwCallTree.Width = lNewWidth
    End If
    lsvCallList.Left = trwCallTree.Left + trwCallTree.Width + picSplitter.Width
    lblListView.Left = lsvCallList.Left
    
    'Resize widths
    lNewWidth = frmMain.Width - lsvCallList.Left - udFormResizeInfo.TreeViewLeft - 100
    If lNewWidth > 0 Then
        lsvCallList.Width = lNewWidth
    End If
End Sub

Private Sub lsvCallList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    On Error GoTo ErrorTrap

    If lsvCallList.SortOrder = lvwAscending Then
        lsvCallList.SortOrder = lvwDescending
    Else
        lsvCallList.SortOrder = lvwAscending
    End If
    lsvCallList.SortKey = ColumnHeader.Index - 1
    lsvCallList.Sorted = False
    lsvCallList.Sorted = True
    
Exit Sub
ErrorTrap:
    ShowError
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show vbModal, Me
    
Exit Sub
ErrorTrap:
    ShowError
End Sub

Private Sub mnuFileCreateData_Click()

    On Error GoTo ErrorTrap

    Dim sVBPLocation As String
    Dim sDatabaseFile As String
    Dim bDeleteOldData As Boolean
    Dim bIncludeSubFolders As Boolean
    
    sVBPLocation = GetSetting(msREG_APP_NAME, msREG_SECTION_LAST_VALUES, msREG_KEY_VBP_LOCATION, vbNullString)
    sDatabaseFile = GetSetting(msREG_APP_NAME, msREG_SECTION_LAST_VALUES, msREG_KEY_DATABASE_FILE, vbNullString)
    bDeleteOldData = True
    bIncludeSubFolders = True
    
    If frmCreateData.DisplayForm(sVBPLocation, sDatabaseFile, bDeleteOldData, bIncludeSubFolders) Then
        Screen.MousePointer = vbHourglass
        Call SaveSetting(msREG_APP_NAME, msREG_SECTION_LAST_VALUES, msREG_KEY_VBP_LOCATION, sVBPLocation)
        Call SaveSetting(msREG_APP_NAME, msREG_SECTION_LAST_VALUES, msREG_KEY_DATABASE_FILE, sDatabaseFile)
        msDatabaseFile = sDatabaseFile
        Call CreateCallTreeData(sVBPLocation, GetConnectionString, bDeleteOldData, bIncludeSubFolders)
        Call InitTree
        Screen.MousePointer = vbDefault
    End If

Exit Sub
ErrorTrap:
    ShowError
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
    End
End Sub

Private Sub mnuFileLoadData_Click()
    On Error GoTo ErrorTrap

    Dim sDatabaseFile As String
    
    sDatabaseFile = GetSetting(msREG_APP_NAME, msREG_SECTION_LAST_VALUES, msREG_KEY_DATABASE_FILE, vbNullString)
    
    If frmLoadData.DisplayForm(sDatabaseFile) Then
        Call SaveSetting(msREG_APP_NAME, msREG_SECTION_LAST_VALUES, msREG_KEY_DATABASE_FILE, sDatabaseFile)
        msDatabaseFile = sDatabaseFile
        Call UpdateStatus("Loading Data...", vbNullString)
        Call InitTree
        Call UpdateStatus("Data Loaded", vbNullString)
    End If

Exit Sub
ErrorTrap:
    ShowError
End Sub

Private Sub mnuFilePrint_Click()
    On Error GoTo ErrorTrap
    Call frmPrint.DisplayForm(trwCallTree, lsvCallList)
Exit Sub
ErrorTrap:
    ShowError
End Sub

Private Sub mnuFilePrinterSetup_Click()
    Call cdlPrinterSetup.ShowPrinter
End Sub

Private Sub mnuFromToTree_Click()

    On Error GoTo ErrorTrap

    Call UpdateStatus("Preparing view...", vbNullString)
    mbCallToFromTree = False
    RefreshFromToViewStatus
    InitTree
    Call UpdateStatus("Ready", vbNullString)
    
Exit Sub
ErrorTrap:
    ShowError
End Sub

Private Sub RefreshFromToViewStatus()
    mnuFromToTree.Checked = Not mbCallToFromTree
    mnuToFromTree.Checked = Not mnuFromToTree.Checked
    tlbMain.Buttons("FROM-TO").Value = IIf(mnuFromToTree.Checked, tbrPressed, tbrUnpressed)
    tlbMain.Buttons("TO-FROM").Value = IIf(tlbMain.Buttons("FROM-TO").Value = tbrPressed, tbrUnpressed, tbrPressed)
End Sub

Private Sub mnuHelpIconInformation_Click()
    frmIconExplanation.Show vbModal, Me
End Sub

Private Sub mnuToFromTree_Click()

    On Error GoTo ErrorTrap

    Call UpdateStatus("Preparing view...", vbNullString)

    mbCallToFromTree = True
    RefreshFromToViewStatus
    InitTree

    Call UpdateStatus("Ready", vbNullString)

Exit Sub
ErrorTrap:
    ShowError
End Sub

Private Sub mnuViewShowStandardReferences_Click()
    mbShowStandardReferences = Not mbShowStandardReferences
    mnuViewShowStandardReferences.Checked = mbShowStandardReferences
    Call SaveSetting(msREG_APP_NAME, msREG_SECTION_SETTINGS, msREG_KEY_SHOW_STANDARD_REFERENCES, mbShowStandardReferences)
    Call InitTree
End Sub


Private Sub picSplitter_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    mbSplitStarted = True
    Call SetCapture(Me.hwnd)
End Sub

Private Sub tlbMain_ButtonClick(ByVal Button As MSComctlLib.Button)

    On Error GoTo ErrorTrap

    Select Case Button.Key
        Case "NEW"
            Call mnuFileCreateData_Click
        Case "OPEN"
            Call mnuFileLoadData_Click
        Case "FROM-TO"
            Call mnuFromToTree_Click
        Case "TO-FROM"
            Call mnuToFromTree_Click
        Case "PRINT"
            Call mnuFilePrint_Click
        Case "HELP"
            Call mnuHelpIconInformation_Click
    End Select

Exit Sub
ErrorTrap:
    ShowError
End Sub

Private Sub trwCallTree_Expand(ByVal Node As MSComctlLib.Node)

    On Error GoTo ErrorTrap

    Call FillTreeNode(Node)
    
Exit Sub
ErrorTrap:
    ShowError
End Sub

Private Sub trwCallTree_NodeClick(ByVal Node As MSComctlLib.Node)

    On Error GoTo ErrorTrap

    Call FillTreeNode(Node)
    Call FillListView(Node.Text, Node.Parent Is Nothing)

Exit Sub
ErrorTrap:
    ShowError
End Sub

Private Sub FillListView(ByVal vsNode As String, Optional ByVal vboolIsRoot As Boolean = False)
    'Get list of all dependents
    
    Screen.MousePointer = vbHourglass
    
    Call UpdateStatus("Getting all child calls...", vbNullString)
    
    Dim oclAllDependents As Collection
    Set oclAllDependents = New Collection
        
    Call AddNodeDependentsInCol(vsNode, oclAllDependents)
    
    'Now we have list of all dependents, so build a where clause for query
    Dim sSQL As String
    Dim sWhereClause As String
    Dim sGroupByClause As String
    Dim sOrderByClause As String
    Dim sNode As Variant
    
    If Not mbCallToFromTree Then
        sSQL = "SELECT CallTo as CallTo1, iif(Min(CallToVersion)=Max(CallToVersion),Max(CallToVersion),'Varies: ' + Min(CallToVersion) + ' To ' + Max(CallToVersion)) as Version1, iif(Min(CallToDescription)=Max(CallToDescription),Max(CallToDescription),'Varies: ''' + Min(CallToDescription) + ''' To ''' + Max(CallToDescription)  + '''') as Description1, Max(IsCallToOCX) as IsOCX1, "
        sSQL = sSQL & "iif((Min(CallToVersion)<>Max(CallToVersion)) or (Min(CallToDescription)<>Max(CallToDescription)),True,False) AS IsVaries, Count(*) As GroupCount, False as IsExe"
        sSQL = sSQL & " FROM ComponentCalls WHERE "
        If Not vboolIsRoot Then
            sWhereClause = vbNullString
            For Each sNode In oclAllDependents
                If sWhereClause <> vbNullString Then
                    sWhereClause = sWhereClause & " OR (CallTo=" & "'" & sNode & "')"
                Else
                    sWhereClause = "(CallTo=" & "'" & sNode & "')"
                End If
            Next sNode
            If sWhereClause = vbNullString Then
                sWhereClause = "1=2"
            End If
        Else
            sWhereClause = "1=1"
        End If
        sGroupByClause = " GROUP BY CallTo"
        sOrderByClause = " ORDER BY CallTo"
    Else
        sSQL = "SELECT Caller as CallTo1, iif(Min(CallerVersion)=Max(CallerVersion),Max(CallerVersion),'Varies: ' + Min(CallerVersion) + ' To ' + Max(CallerVersion)) as Version1, iif(Min(CallerDescription)=Max(CallerDescription),Max(CallerDescription),'Varies: ''' + Min(CallerDescription) + ''' To ''' + Max(CallerDescription)  + '''') as Description1, iif(Max(CallerType)='Control',true,false) as IsOCX1, "
        sSQL = sSQL & "iif((Min(CallerVersion)<>Max(CallerVersion)) or (Min(CallerDescription)<>Max(CallerDescription)),True,False) AS IsVaries, Count(*) As GroupCount, iif(Max(CallerType)='Exe',true,false) as IsExe"
        sSQL = sSQL & " FROM ComponentCalls WHERE "
        If Not vboolIsRoot Then
            sWhereClause = vbNullString
            For Each sNode In oclAllDependents
                If sWhereClause <> vbNullString Then
                    sWhereClause = sWhereClause & " OR (Caller=" & "'" & sNode & "')"
                Else
                    sWhereClause = "(Caller=" & "'" & sNode & "')"
                End If
            Next sNode
            If sWhereClause = vbNullString Then
                sWhereClause = "1=2"
            End If
        Else
            sWhereClause = "1=1"
        End If
        sGroupByClause = " GROUP BY Caller"
        sOrderByClause = " ORDER BY Caller"
    End If
    
    sSQL = sSQL & sWhereClause & sGroupByClause & sOrderByClause
    
    Call UpdateStatus(, "Reading database...")
    'Now get the data
    Dim rsAllNodes As ADODB.Recordset
    Set rsAllNodes = New ADODB.Recordset
    Call rsAllNodes.Open(sSQL, GetConnectionString, adOpenStatic, adLockReadOnly)
    
    Call UpdateStatus(, "Populating list...")
    'Populate the Listview
    lsvCallList.ListItems.Clear

    Dim lsiNodeInfo As ListItem
    Do While Not rsAllNodes.EOF
        Set lsiNodeInfo = lsvCallList.ListItems.Add(, , "")
        'lsiNodeInfo.SubItems(1) = rsAllNodes("Caller").Value
        lsiNodeInfo.SubItems(1) = rsAllNodes("CallTo1").Value
        lsiNodeInfo.SubItems(2) = rsAllNodes("Version1").Value & ""
        lsiNodeInfo.SubItems(3) = rsAllNodes("Description1").Value & ""
        If rsAllNodes("IsOCX1").Value = True Then
            lsiNodeInfo.SmallIcon = 6
        ElseIf rsAllNodes("IsVaries").Value = True Then
            lsiNodeInfo.SmallIcon = 5
        ElseIf rsAllNodes("IsExe").Value = True Then
            lsiNodeInfo.SmallIcon = 8
        End If
        rsAllNodes.MoveNext
    Loop
    
    rsAllNodes.Close
    Set rsAllNodes = Nothing
    
    Set oclAllDependents = Nothing
    
    Screen.MousePointer = vbDefault
    
    Call UpdateStatus("Ready", vbNullString)
    
End Sub

'Add the dependents of specified node in the collection
Public Sub AddNodeDependentsInCol(ByVal vsNode As String, ByVal voclCol As Collection)
    Dim oclDependentsForANode As Collection
    
    'Get the dependents of this node
    Set oclDependentsForANode = GetObjectInCollection(moclMaster, vsNode)
    If Not (oclDependentsForANode Is Nothing) Then
        Dim vNodeItem As Variant
        Dim sNodeItem As Variant
        For Each sNodeItem In oclDependentsForANode
            sNodeItem = CStr(sNodeItem)
            Call SafeAddToCollection(voclCol, sNodeItem, LCase(sNodeItem))
            Call AddNodeDependentsInCol(sNodeItem, voclCol)
        Next sNodeItem
    End If
    
End Sub

'This forces only one item!
Private Sub SafeAddToCollection(ByVal voclCol As Collection, ByVal vvItem As Variant, ByVal vsKey As String)
    On Error Resume Next
    Call voclCol.Add(vvItem, CStr(vsKey))
End Sub

Private Sub ShowError()
    Me.MousePointer = vbDefault
    Screen.MousePointer = vbDefault
    MsgBox "Error " & Err.Number & " : " & Err.Description
    Call UpdateStatus("Error occured", Err.Description)
End Sub

Private Sub UpdateStatus(Optional ByVal vsCatagory As Variant, Optional ByVal vsDescription As Variant)
    If Not IsMissing(vsCatagory) Then
        If StrComp(vsCatagory, "ready", vbTextCompare) <> 0 Then
            stbMain.Panels(1).Text = vsCatagory
        Else
            If lsvCallList.ListItems.Count <> 0 Then
                stbMain.Panels(1).Text = lsvCallList.ListItems.Count & " component(s)"
            Else
                stbMain.Panels(1).Text = vsCatagory
            End If
        End If
    End If
    If Not IsMissing(vsDescription) Then
        stbMain.Panels(2).Text = vsDescription
    End If
End Sub
