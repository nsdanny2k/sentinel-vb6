VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private moEncr As New clsEncryption
Private moR As New clsDataReader
Private moW As New clsDataWriter

Public Sub ReadListControl(ByVal pTable As String, ByVal pPublic As Boolean, ByVal pControl As Control)
Dim fnum
Dim strTmp As String
Dim strPath As String
On Error Resume Next
strPath = LogFilePath(pTable, pPublic)
If FileExists(strPath) Then
    fnum = FreeFile(1)
    Open strPath For Input As #fnum
    While Not EOF(fnum)
        Input #fnum, strTmp
        pControl.AddItem strTmp
    Wend
    Close #fnum
End If
End Sub

Public Sub SaveListControl(ByVal pTable As String, ByVal pPublic As Boolean, ByVal pControl As Control)
Dim fnum
Dim strTmp As String
Dim strPath As String
Dim n As Long
On Error Resume Next
strPath = LogFilePath(pTable, pPublic)
If pControl.ListCount > 0 Then
    fnum = FreeFile(1)
    Open strPath For Output As #fnum
    For n = 1 To pControl.ListCount
        strTmp = pControl.List(n - 1)
        Print #fnum, strTmp
    Next n
    Close #fnum
End If
End Sub

Public Function GetLogSetting(ByRef pArrIn() As String, ByVal pSetting As String, ByRef pRetval, Optional ByVal pDefault As String = "") As Boolean
Dim l As Integer
Dim u As Integer
Dim n As Integer
On Error GoTo ExitNotFound
l = LBound(pArrIn, 2)
u = UBound(pArrIn, 2)
For n = l To u
    If LCase(pArrIn(1, n)) = LCase(pSetting) Then
        pRetval = pArrIn(2, n)
        GetLogSetting = True
        Exit Function
    End If
Next
ExitNotFound:
pRetval = pDefault
GetLogSetting = False
End Function

Public Function LogDirPath(ByVal pPublic As Boolean) As String
Dim strDir As String
On Error Resume Next
If Not pPublic Or gSystemPath = "" Or LCase(Trim(gSystemPath)) = "c:\" Then
    strDir = ServerPath
Else
    strDir = gSystemPath
End If
If Right(strDir, 1) <> "\" Then strDir = strDir & "\"
LogDirPath = strDir & "logs\"
End Function

Public Function LogFilePath(ByVal pName As String, Optional ByVal pPublic As Boolean = False) As String
Dim strPath As String
Dim strDir As String
On Error Resume Next
strDir = LogDirPath(pPublic)
strPath = strDir & pName & gExtLog
CreateNewDir GetPathPart(strDir)
LogFilePath = strPath
End Function

Public Function ServerPath() As String
ServerPath = App.Path
End Function

Public Function AddLog(ByVal strLog As String, Optional ByVal pFile As String = "", Optional ByVal pPublic As Boolean = True, Optional ByVal pFloppy As Boolean = False, Optional ByVal pNewLine As Boolean = False, Optional ByVal pPath As String = "") As Boolean
Dim fnum
Dim strPath As String
On Error Resume Next
If pPath = "" Then
    If pFloppy Then
        strPath = GetFloppyDrive & pFile & ".log"
    Else
        strPath = LogFilePath(pFile, pPublic)
    End If
Else
    strPath = pPath
End If
fnum = FreeFile(1)
CreateNewDir GetPathPart(strPath)
If gHideLogFiles Then
    SetFileAttrNormal strLog
End If
Open strPath For Append As #fnum
strLog = moEncr.EncryptText(strLog)
If pNewLine Then
    strLog = strLog & vbCrLf
End If
Print #fnum, strLog
Close #fnum
If gHideLogFiles Then
    SetFileAttrHidden strLog
End If
AddLog = (Err.Number = 0)
End Function

Public Function AddSetting(ByVal pSetting As String, ByVal pValue As String, Optional ByVal pFile As String = "", Optional ByVal pPublic As Boolean = True, Optional ByVal pFloppy As Boolean = False, Optional ByVal pNewLine As Boolean = False, Optional ByVal pPath As String = "") As Boolean
Dim fnum
Dim strPath As String
Dim strLog As String
On Error Resume Next
If pPath = "" Then
    If pFloppy Then
        strPath = GetFloppyDrive & pFile & ".log"
    Else
        strPath = LogFilePath(pFile, pPublic)
    End If
Else
    strPath = pPath
End If
fnum = FreeFile(1)
CreateNewDir GetPathPart(strPath)
If gHideLogFiles Then
    SetFileAttrNormal strPath
End If
Open strPath For Append As #fnum
strLog = moEncr.EncryptText(pSetting & "=" & pValue)
Print #fnum, strLog
Close #fnum
If gHideLogFiles Then
    SetFileAttrHidden strPath
End If
AddSetting = (Err.Number = 0)
End Function

Public Function GetFileStream( _
    ByVal pPath As String, _
    Optional ByRef pData As String = "", _
    Optional ByVal pTableName As String = "", _
    Optional ByVal pBizObj As Object) As Boolean
Dim strBuff As String
Dim strTmp As String
Dim tmpArr As Variant
Dim n As Integer
Dim strData As String
Dim fnum
Dim bInit As Boolean
On Error GoTo ErrExit
n = 0
fnum = FreeFile(1)
Open pPath For Input As #fnum
If Not EOF(fnum) Then
    Input #fnum, strTmp
    strData = strTmp
End If
Do Until (EOF(fnum) = True)
    Input #fnum, strTmp
    strData = strData & vbCrLf & strTmp
Loop
Close #fnum
pData = strData
GetFileStream = True
Exit Function

ErrExit:
GetFileStream = False
On Error Resume Next
Close #fnum
Exit Function
End Function

Public Function FillLog(ByRef pArrIn() As String, ByVal pPath As String, Optional ByRef pData As String = "") As Boolean
Dim strBuff As String
Dim strTmp As String
Dim tmpArr As Variant
Dim n As Integer
Dim strData As String
Dim fnum
On Error GoTo ErrExit
n = 0
fnum = FreeFile(1)
If gHideLogFiles Then
    SetFileAttrNormal pPath
End If
Open pPath For Input As #fnum
Do Until (EOF(fnum) = True)
    Input #fnum, strTmp
    strData = strData & strTmp & vbCrLf
    strTmp = moEncr.DecryptText(strTmp)
    If InStr(1, strTmp, "=") <> 0 Then
        n = n + 1
        ReDim Preserve pArrIn(1 To 2, 1 To n)
        tmpArr = Split(strTmp, "=")
        pArrIn(1, n) = tmpArr(LBound(tmpArr))
        pArrIn(2, n) = tmpArr(UBound(tmpArr))
    End If
Loop
Close #fnum
If gHideLogFiles Then
    SetFileAttrHidden pPath
End If
pData = strData
FillLog = True
Exit Function

ErrExit:
FillLog = False
On Error Resume Next
Close #fnum
Exit Function
End Function

Public Function SaveRec(ByRef arrColData As Variant, ByRef arrColField As Variant, ByRef arrColType As Variant, ByVal pTable As String, Optional ByVal pPath As String = "") As Boolean
On Error GoTo ErrExit
Dim u As Long
Dim l As Long
Dim n As Long
Dim fnum As Long
Dim strTmp As String
Dim strPath As String
If TypeName(arrColData) = "Variant()" Or TypeName(arrColData) = "String()" Then
    u = UBound(arrColField)
    l = LBound(arrColField)
    'store text file
    For n = l To u
        strTmp = strTmp & arrColData(n) & gRecDelim
    Next
    strTmp = Left(strTmp, Len(strTmp) - 1)
    fnum = FreeFile(1)
    If pPath <> "" Then
        strPath = pPath
    Else
        strPath = LogFilePath(pTable, True)
    End If
    Open strPath For Append As #fnum
    strTmp = moEncr.EncryptText(strTmp)
    Print #fnum, strTmp
    Close #fnum
ElseIf TypeName(arrColData) = "String" Then
    strTmp = arrColData
    fnum = FreeFile(1)
    If pPath = "" Then
        strPath = pPath
    Else
        strPath = LogFilePath(pTable, True)
    End If
    Open strPath For Append As #fnum
    strTmp = moEncr.EncryptText(strTmp)
    Print #fnum, strTmp
    Close #fnum
End If
SaveRec = True
Exit Function
ErrExit:
SaveRec = False
End Function

Public Function FillList( _
    ByRef pCol As Collection, _
    ByVal pObj As Object, _
    ByVal pTable As String, _
    ByVal pPublic As Boolean, _
    Optional ByVal pPath As String = "", _
    Optional ByVal pIsSilent As Boolean = False, _
    Optional ByVal pIsRandomAccess As Boolean = False, _
    Optional ByVal pNoDB As Boolean = False) As Boolean
Dim strPath As String
Dim oDataReader As New clsDataReader
Dim pRetval As Boolean
On Error GoTo ErrExit
pRetval = False
If pPath = "" Then
    strPath = LogFilePath(pTable, pPublic)
Else
    strPath = pPath
End If
If FileExists(strPath) Then
    oDataReader.Filename = strPath
    oDataReader.IsEncrypted = True
    oDataReader.IsRandomAccess = pIsRandomAccess
    If oDataReader.OpenFile Then
        Dim strBuff As String
        Set pCol = New Collection
        While Not oDataReader.EndOfFile
            If oDataReader.ReadLine(strBuff, pObj) Then
                strBuff = pObj.GetString()
                pCol.Add strBuff
            End If
        Wend
        oDataReader.CloseFile
        pRetval = True
    Else
        GoTo ErrExit
    End If
Else
    GoTo ErrExit
End If
FillList = pRetval
Exit Function

ErrExit:
FillList = False
If Not pIsSilent Then
    If Err <> 0 And Err <> 76 Then
        frmMessage.Display Error, vbInformation, "Unable to load data"
    End If
End If
End Function

Public Function ReplicateList( _
    ByVal pTable As String, _
    Optional ByVal pIsSilent As Boolean = False, _
    Optional ByVal pNoDB As Boolean = False) As Boolean
On Error GoTo ErrExit
Dim strPubFile As String
Dim strLocFile As String
strPubFile = LogFilePath(pTable, True)
strLocFile = LogFilePath(pTable, False)
If PathExists(strPubFile) Then
    If PathExists(strLocFile) Then
        If FileDateTime(strPubFile) <> FileDateTime(strLocFile) Then
            KillPath strLocFile
            FileCopy strPubFile, strLocFile
        End If
    Else
        CreateNewDir GetPathPart(strLocFile)
        FileCopy strPubFile, strLocFile
    End If
End If
ReplicateList = True
Exit Function

ErrExit:
ReplicateList = False
If Not pIsSilent Then
    If Err <> 0 Then
        MsgBox Error, vbInformation, "Assert Failed"
    End If
End If
Exit Function
End Function
