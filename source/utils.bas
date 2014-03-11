Attribute VB_Name = "modUtils"
Option Explicit

Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Const BIF_RETURNONLYFSDIRS = 1
Private Type BrowseInfo
    hWndOwner       As Long
    pIDLRoot        As Long
    pszDisplayName  As String
    lpszTitle       As String
    ulFlags         As Long
    lpfnCallBack    As Long
    lparam          As Long
    iImage          As Long
End Type


Private moErr As ErrObject
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Const MAX_PATH = 260


Public Declare Function WinExec Lib "kernel32" (ByVal lpCmdLine As String, ByVal nCmdShow As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lparam As Any) As Long
Public Const CB_SHOWDROPDOWN = &H14F
Public Const CB_GETDROPPEDSTATE = &H157
Public Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long

Public Const SW_SHOW = 5

Public Function CheckToBool(ByVal venmCheckBoxValue As CheckBoxConstants) As Boolean
    CheckToBool = (venmCheckBoxValue = vbChecked)
End Function

Public Function BoolToCheck(ByVal vboolValue As Boolean) As CheckBoxConstants
    BoolToCheck = IIf(vboolValue, vbChecked, vbUnchecked)
End Function

Public Function WordsToCollection(ByVal sWords As String) As Collection
    Dim oclWords As New Collection
    Dim lSpacePos As Long
    Dim sWord As String
    Const sSPACE As String = " "
    
    Do
        sWords = Trim$(sWords)
        lSpacePos = InStr(1, sWords, sSPACE)
        If lSpacePos = 0 Then
            If sWords <> vbNullString Then
                oclWords.Add sWords
            End If
        Else
            sWord = Left(sWords, lSpacePos - 1)
            oclWords.Add sWord
            sWords = Mid(sWords, lSpacePos + 1)
        End If
    Loop While (lSpacePos <> 0)
    
    Set WordsToCollection = oclWords
    Set oclWords = Nothing
End Function

Public Function IsTextBoxEmpty(ByVal voTextBox As TextBox) As Boolean
    If Trim$(voTextBox.Text) = vbNullString Then
        IsTextBoxEmpty = True
    Else
        IsTextBoxEmpty = False
    End If
End Function

Public Sub ReRaiseError()
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub

Public Function TextRightAlign(ByVal vsText As String, ByVal vlFixedlength As Long) As String
    Dim lTextLen As Long
    lTextLen = Len(vsText)
    If lTextLen < vlFixedlength Then
        TextRightAlign = Space$(vlFixedlength - lTextLen) & vsText
    Else
        TextRightAlign = vsText
    End If
End Function

Public Function ParseStringToArray(ByVal vsString As String, ByVal vsSeperator As String) As Variant
    Dim vaArray As Variant
    Dim lSeperatorPos As Long
    Dim sStringPart As String
    Dim lArrayIndex As Long
    
    Const lSTARTING_ARRAY_INDEX As Long = 1
    
    lArrayIndex = 1
    
    ReDim vaArray(lSTARTING_ARRAY_INDEX To lArrayIndex)
    sStringPart = vsString
    Do
        lSeperatorPos = InStr(1, sStringPart, vsSeperator, vbBinaryCompare)
        'No seperator found
        If lSeperatorPos = 0 Then
            vaArray(lArrayIndex) = sStringPart
        Else
            sStringPart = Left(sStringPart, lSeperatorPos - 1)
            vaArray(lArrayIndex) = sStringPart
            lArrayIndex = lArrayIndex + 1
            ReDim Preserve vaArray(lSTARTING_ARRAY_INDEX To lArrayIndex)
            sStringPart = Mid(sStringPart, lSeperatorPos)
        End If
    Loop While (lSeperatorPos <> 0)
    
    Set ParseStringToArray = vaArray
End Function

Public Sub ExecAnyFile(ByVal vsFileName As String)
    Dim lWinExecReturn As Long
    lWinExecReturn = WinExec(vsFileName, SW_SHOW)
    If lWinExecReturn <= 31 Then
        Err.Raise lWinExecReturn, , "Error executing program for " & vsFileName
    End If
End Sub

Public Sub SafeSetFocus(ByVal vctrl As Control)
 On Error Resume Next
 vctrl.SetFocus
End Sub

Public Sub SelectAllTextInControl(ByVal vctl As Control)
    vctl.SelStart = 0
    vctl.SelLen = Len(vctl.Text)
End Sub

Public Sub DropCombo(ByVal vctlCombo As ComboBox)
    Dim bIsDropped As Boolean
    
    bIsDropped = (SendMessage(vctlCombo.hwnd, CB_GETDROPPEDSTATE, 0, 0) = -1)
    
    If Not bIsDropped Then
        Call SendMessage(vctlCombo.hwnd, CB_SHOWDROPDOWN, -1, 0)
    End If
End Sub

Public Sub CloseCombo(ByVal vctlCombo As ComboBox)
    Dim bIsDropped As Boolean
    
    bIsDropped = (SendMessage(vctlCombo.hwnd, CB_GETDROPPEDSTATE, 0, 0) = -1)
    
    If bIsDropped Then
        Call SendMessage(vctlCombo.hwnd, CB_SHOWDROPDOWN, 0, 0)
    End If
End Sub

Public Function IsComboDropped(ByVal vctlCombo As ComboBox) As Boolean
    IsComboDropped = (SendMessage(vctlCombo.hwnd, CB_GETDROPPEDSTATE, 0, 0) = -1)
End Function

'Select List or combo item without causing Click event
Public Sub SilentListItemSelect(ByVal vctl As ComboBox, ByVal vlListIndex As Long)
    Const CB_SETCURSEL = &H14E
    Call SendMessage(vctl.hwnd, CB_SETCURSEL, vlListIndex, 0)
End Sub

Public Function GetPathWithSlash(ByVal vsPath As String, Optional ByVal vsSlashChar As String = "\") As String
    If Right$(vsPath, 1) <> vsSlashChar Then
        GetPathWithSlash = vsPath & vsSlashChar
    Else
        GetPathWithSlash = vsPath
    End If
End Function

Public Function ClearCollection(ByVal voclColl As Collection)
    Dim i As Long
    
    For i = voclColl.Count To 1 Step -1
        voclColl.Remove i
    Next i
    
End Function

Public Function BrowseForFolder(ByVal vsDefaultFolder As String, Optional ByVal vhWndOwner As Long = 0, Optional ByVal vsPrompt As String = vbNullString) As String
    Dim iNull As Integer
    Dim lpIDList As Long
    Dim lResult As Long
    Dim sPath As String
    Dim udtBI As BrowseInfo

    With udtBI
        .hWndOwner = vhWndOwner
        .lpszTitle = vsPrompt
        .ulFlags = BIF_RETURNONLYFSDIRS
    End With
    
    lpIDList = SHBrowseForFolder(udtBI)
    If lpIDList Then
        sPath = String$(MAX_PATH, 0)
        lResult = SHGetPathFromIDList(lpIDList, sPath)
        Call CoTaskMemFree(lpIDList)
        iNull = InStr(sPath, vbNullChar)
        If iNull Then
            sPath = Left$(sPath, iNull - 1)
        End If
    Else
        sPath = vsDefaultFolder
    End If

    BrowseForFolder = sPath
    
End Function

Public Sub SortCollection(ByVal voclValues As Collection)
    Dim lSortLevelIndex As Long
    Dim lElementIndex As Long
    Dim sTempVal1ForSwap As String
    Dim sTempVal2ForSwap As String
    
    For lSortLevelIndex = 1 To voclValues.Count - 1
        For lElementIndex = 1 To voclValues.Count - 1
            If UCase(voclValues(lElementIndex)) > UCase(voclValues(lElementIndex + 1)) Then
                sTempVal1ForSwap = voclValues(lElementIndex)
                sTempVal2ForSwap = voclValues(lElementIndex + 1)
                Call voclValues.Remove(lElementIndex)
                Call voclValues.Add(sTempVal2ForSwap, , lElementIndex)
                Call voclValues.Remove(lElementIndex + 1)
                Call voclValues.Add(sTempVal1ForSwap, , , lElementIndex)
            End If
        Next lElementIndex
    Next lSortLevelIndex
End Sub

Public Function IsNumber(ByVal vvItem As Variant) As Boolean
    On Error GoTo ErrorTrap
    
    IsNumber = False
    If Not IsEmpty(vvItem) Then
        If Not IsNull(vvItem) Then
            If Trim(CStr(vvItem)) <> vbNullString Then
                If IsNumeric(vvItem) Then
                    IsNumber = True
                End If
            End If
        End If
    End If
    
Exit Function
ErrorTrap:
    IsNumber = False
End Function


Public Function IsDir(ByVal vsDir As String) As Boolean
    On Error GoTo ErrorTrap
    IsDir = ((GetAttr(vsDir) And vbDirectory) = vbDirectory)
Exit Function
ErrorTrap:
    If Err.Number = 5 Then      'Invalid procedure call or argument - occures if file (ex. pagefile.sys) is in use
        IsDir = False
    End If
End Function


Public Sub SaveErrorObj()
    Set moErr = Nothing
    Set moErr = New ErrObject
    With moErr
        .Number = Err.Number
        .Description = Err.Description
        .HelpContext = Err.HelpContext
        .HelpFile = Err.HelpFile
        .Source = Err.Source
    End With
End Sub

Public Sub RestoreErrorObj()
    If moErr Is Nothing Then
        With Err
            .Number = moErr.Number
            .Description = moErr.Description
            .HelpContext = moErr.HelpContext
            .HelpFile = moErr.HelpFile
            .Source = moErr.Source
        End With
    End If
    Set moErr = Nothing
End Sub

