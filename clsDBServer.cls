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

Private Const gCurDBVer As Long = 34

Private moEncr As New clsEncryption
Private moCmd As Object
Private moR As New clsDataReader
Private moW As New clsDataWriter

Public Sub ReadListControl(ByVal pTable As String, ByVal pPublic As Boolean, ByVal pControl As Control)
Dim fnum
Dim strTmp As String
Dim strPath As String
On Error Resume Next
If UseDatabaseStorage Then
    On Error GoTo ErrExit
    pControl.Clear
    Dim oRS As Object
    If InStr(1, pTable, "[") <= 0 And UCase(Left(pTable, 7)) <> "SELECT " Then
        pTable = "[" & pTable & "]"
        Set oRS = godb.GetRS("SELECT * FROM " & pTable, True)
    Else
        Set oRS = godb.GetRS(pTable, True)
    End If
    'If oRS.RecordCount <> 0 Then
    If Not oRS.EOF Then
        While Not oRS.EOF
            strTmp = oRS(0)
            pControl.AddItem strTmp
            oRS.MoveNext
        Wend
    End If
Else
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
End If
ErrExit:
End Sub

Public Sub SaveListControl(ByVal pTable As String, ByVal pPublic As Boolean, ByVal pControl As Control)
Dim fnum
Dim strTmp As String
Dim strPath As String
Dim n As Long
On Error Resume Next
If UseDatabaseStorage Then
    On Error GoTo ErrExit
    Dim rs As Object
    Set rs = godb.GetRS(pTable)
    RunSQL "DELETE FROM " & pTable
    If pControl.ListCount > 0 Then
        For n = 1 To pControl.ListCount
            strTmp = pControl.List(n - 1)
            RunSQL "INSERT INTO " & pTable & "(" & rs.Fields(0).name & ") VALUES ('" & strTmp & "')"
        Next n
    End If
    rs.Close
Else
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
End If
ErrExit:
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
AddLog = (Err.number = 0)
End Function


Public Function FillLog(ByRef pArrIn() As String, ByVal pPath As String, Optional ByRef pData As String = "") As Boolean
Dim strBuff As String
Dim strTmp As String
Dim tmpArr As Variant
Dim n As Integer
Dim strData As String
Dim fnum
On Error GoTo ErrExit
'Debug.Print pPath
n = 0
fnum = FreeFile(1)
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
pData = strData
FillLog = True
Exit Function

ErrExit:
FillLog = False
On Error Resume Next
Close #fnum
Exit Function
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
AddSetting = (Err.number = 0)
End Function

'db storage =============================

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
If UseDatabaseStorage And pTableName <> "" And Not pBizObj Is Nothing Then
    Dim oRS As Object
    Dim oEncr As New clsEncryption
    Set oRS = GetRS("SELECT * FROM " & pTableName, True)
    'strData = oRS.GetString(adClipString, , gRecDelim2, vbCrLf, "")
    'If oRS.RecordCount > 0 Then
    If Not oRS.EOF Then
        If Not oRS.EOF Then
            strBuff = goDBConn.GetString(oRS, pBizObj.GetFields)
            If pBizObj.ReadLine(strBuff) Then
                strData = oEncr.EncryptText(pBizObj.GetString)
            End If
            oRS.MoveNext
        End If
        While Not oRS.EOF
            strBuff = goDBConn.GetString(oRS, pBizObj.GetFields)
            If pBizObj.ReadLine(strBuff) Then
                strData = strData & vbCrLf & oEncr.EncryptText(pBizObj.GetString)
            End If
            oRS.MoveNext
        Wend
    End If
    oRS.Close
    Set oRS = Nothing
    GetFileStream = True
Else
    Dim oDataReader As New clsDataReader
    oDataReader.FileName = pPath
    oDataReader.IsEncrypted = False
    oDataReader.IsRandomAccess = False
    oDataReader.IsHidden = False
    'already encrypted, do not decrypt
    If oDataReader.OpenFile Then
        If Not oDataReader.EndOfFile Then
            If oDataReader.ReadLine(strTmp) Then
                strData = strTmp
            End If
        End If
        While Not oDataReader.EndOfFile
            If oDataReader.ReadLine(strTmp) Then
                strData = strData & vbCrLf & strTmp
            End If
        Wend
        oDataReader.CloseFile
    End If
    GetFileStream = True
End If
pData = strData
Exit Function

ErrExit:
GetFileStream = False
On Error Resume Next
If Not oDataReader Is Nothing Then
    oDataReader.CloseFile
End If
Exit Function
End Function

Public Function AddRec( _
    ByVal pTableName As String, _
    ByRef arrColData As Variant, _
    ByRef arrColField As Variant, _
    ByRef arrColType As Variant) As Boolean
On Error GoTo ErrExit
Dim strCmd As String
        strCmd = "INSERT INTO " & pTableName & "("
        strCmd = strCmd & goDBConn.GetFieldNames(arrColField) & ")"
        strCmd = strCmd & " VALUES ("
        strCmd = strCmd & goDBConn.GetFieldValues2(arrColData, arrColField, arrColType, "a") & ")"
RunSQL strCmd, True
'If moCmd Is Nothing Then Set moCmd = CreateDBCommand
'moCmd.ActiveConnection = goDBConn.Connection
'moCmd.CommandText = strCmd
'moCmd.Execute
AddRec = True
Exit Function
ErrExit:
MsgBox "Error:" & Error, vbInformation, "Error"
AddRec = False
End Function

Public Function SaveRec(ByRef arrColData As Variant, _
    ByRef arrColField As Variant, ByRef arrColType As Variant, _
    ByVal pTableName As String, Optional ByVal pPath As String = "") As Boolean
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
    If UseDatabaseStorage And pTableName <> "" Then
        If Not AddRec(pTableName, arrColData, arrColField, arrColType) Then
            GoTo ErrExit
        End If
    Else
        For n = l To u
            strTmp = strTmp & arrColData(n) & gRecDelim
        Next
        strTmp = Left(strTmp, Len(strTmp) - 1)
        fnum = FreeFile(1)
        If pPath <> "" Then
            strPath = pPath
        Else
            strPath = LogFilePath(pTableName, True)
        End If
        Open strPath For Append As #fnum
        strTmp = moEncr.EncryptText(strTmp)
        Print #fnum, strTmp
        Close #fnum
    End If
ElseIf TypeName(arrColData) = "String" Then
    strTmp = arrColData
    fnum = FreeFile(1)
    If pPath = "" Then
        strPath = pPath
    Else
        strPath = LogFilePath(pTableName, True)
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
    ByVal pTableName As String, _
    ByVal pPublic As Boolean, _
    Optional ByVal pPath As String = "", _
    Optional ByVal pIsSilent As Boolean = False, _
    Optional ByVal pIsRandomAccess As Boolean = False, _
    Optional ByVal pNoDB As Boolean = False, _
    Optional ByVal pRecCriteria As String = "") As Boolean
Dim strPath As String
Dim oDataReader As New clsDataReader
Dim pRetval As Boolean
Dim strBuff As String
Dim oRS As Object
On Error GoTo ErrExit
pRetval = False
If UseDatabaseStorage And pTableName <> "" Then
    Set oRS = godb.GetRS("SELECT * FROM " & pTableName & IIf(pRecCriteria = "", "", " WHERE " & pRecCriteria), True)
    Set pCol = New Collection
    'If oRS.RecordCount > 0 Then
    If Not oRS.EOF Then
         Do While Not oRS.EOF
            strBuff = goDBConn.GetString(oRS, pObj.GetFields)
            If pObj.ReadLine(strBuff) Then
                pCol.Add pObj.GetString
            End If
            oRS.MoveNext
         Loop
    End If
    pRetval = True
    oRS.Close
    Set oRS = Nothing
Else
    If pPath = "" Then
        strPath = LogFilePath(pTableName, pPublic)
    Else
        strPath = pPath
    End If
    If FileExists(strPath) Then
        oDataReader.FileName = strPath
        oDataReader.IsEncrypted = True
        oDataReader.IsRandomAccess = pIsRandomAccess
        If oDataReader.OpenFile(pObj) Then
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
End If
FillList = pRetval
Exit Function

ErrExit:
FillList = False
If Not pIsSilent Then
    If Err <> 0 Then
        MsgBox Error, vbInformation, "Unable to load data"
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
If Not UseDatabaseStorage Then
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

Public Function CreateDBConnection() As Object
Set CreateDBConnection = CreateObject("ADODB.Connection")
End Function

Public Function CreateDBCommand() As Object
Set CreateDBCommand = CreateObject("ADODB.Command")
End Function

Public Function CreateDBRecordset() As Object
Set CreateDBRecordset = CreateObject("ADODB.Recordset")
End Function

Public Function RunSQL(ByVal pSQL As String, Optional ByVal pRaiseError As Boolean = False) As Boolean
If Not pRaiseError Then
    On Error GoTo ErrExit
End If
Dim oCmd As Object
Set oCmd = CreateDBCommand
If UseDatabaseStorage Then
    oCmd.ActiveConnection = goDBConn.Connection_String
    oCmd.CommandText = pSQL
    oCmd.Execute
    Set oCmd = Nothing
Else
    GoTo ErrExit
End If
RunSQL = True
Exit Function
ErrExit:
RunSQL = False
End Function

Public Function GetRS(ByVal pSQL As String, Optional ByVal pReadOnly As Boolean = False) As Object
On Error GoTo ErrExit
Set GetRS = CreateDBRecordset
If pReadOnly Then
    GetRS.Open pSQL, _
        goDBConn.Connection, 1, 1 'adOpenKeyset, adLockReadOnly
Else
    GetRS.Open pSQL, _
        goDBConn.Connection, 2, 3 'adOpenDynamic, adLockOptimistic
End If
Exit Function
ErrExit:
Set GetRS = Nothing
Debug.Print Error
Exit Function
End Function

Public Function GetValue(ByVal pSQL As String) As Variant
On Error GoTo ErrExit
Dim rs As Object
Set rs = GetRS(pSQL)
If rs.EOF Then
    GetValue = Null
Else
    GetValue = rs(0)
End If
Exit Function
ErrExit:
GetValue = Null
Debug.Print Error
Exit Function
End Function

Private Function SetCurVer(ByRef oCmd As Object) As Boolean
On Error GoTo ErrExit
Dim strCmd As String
Dim nCurDBVer As Integer

nCurDBVer = gCurDBVer

strCmd = "UPDATE DataVer SET Version='" & nCurDBVer & "'"
RunSQL strCmd, True
'oCmd.CommandText = strCmd
'oCmd.Execute

SetCurVer = True
Exit Function

ErrExit:
SetCurVer = False
Exit Function
End Function

Public Function CheckDBVersion() As Boolean
Dim strBuff As String
Dim strCmd As String
Dim oCmd As Object
Dim nSec As Integer
Dim nCurDBVer As Integer

On Error GoTo ErrExit

nSec = 1

If UseDatabaseStorage Then

    'Set oCmd = CreateDBCommand
    'oCmd.ActiveConnection = goDBConn.Connection_String
    
    Dim oRS As Object
    Set oRS = godb.GetRS("SELECT * FROM DataVer", True)
nSec = 2
    'If oRS.RecordCount > 0 Then
    If Not oRS.EOF Then
        If Not oRS.EOF Then
            DoUpdateVer Val(oRS("Version")), oCmd
            SetCurVer oCmd
        End If
    Else
        strCmd = "INSERT INTO DataVer (Version) Values ('" & gCurDBVer & "')"
        RunSQL strCmd, True
        'oCmd.CommandText = strCmd
        'oCmd.Execute
        DoUpdateVer 0, oCmd
    End If
    oRS.Close
    Set oRS = Nothing
End If
CheckDBVersion = True
Exit Function

ErrExit:
If (Err = -2147217887 And nSec = 1) Or (Err = 91 And nSec = 2) Then
    If gUseMSDE Then
        strCmd = "CREATE TABLE DataVer (Version Char(15))"
    Else
        strCmd = "CREATE TABLE DataVer (Version TEXT(15))"
    End If
    RunSQL strCmd, True
    'oCmd.CommandText = strCmd
    'oCmd.Execute

    strCmd = "INSERT INTO DataVer (Version) Values ('" & gCurDBVer & "')"
    RunSQL strCmd, True
    'oCmd.CommandText = strCmd
    'oCmd.Execute
    DoUpdateVer 0, oCmd
    
    CheckDBVersion = True
Else
    CheckDBVersion = False
End If
Exit Function
End Function

Private Function UpdateVer1(ByRef oCmd As Object) As Boolean
On Error GoTo ErrExit
Dim strCmd As String

If gUseMSDE Then
    strCmd = "ALTER TABLE " & gFileSales & " ADD tmpRef INT, tmpStation CHAR(4)"
Else
    strCmd = "ALTER TABLE " & gFileSales & " ADD COLUMN tmpRef LONG, tmpStation TEXT(4)"
End If
RunSQL strCmd, True
'oCmd.CommandText = strCmd
'oCmd.Execute
    
If gUseMSDE Then
    strCmd = "CREATE INDEX tmpStation ON " & gFileSales & " (tmpStation)"
Else
    strCmd = "CREATE INDEX tmpStation ON " & gFileSales & " (tmpStation)"
End If
RunSQL strCmd, True
'oCmd.CommandText = strCmd
'oCmd.Execute

If gUseMSDE Then
    strCmd = "CREATE INDEX tmpRef ON " & gFileSales & " (tmpRef)"
Else
    strCmd = "CREATE INDEX tmpRef ON " & gFileSales & " (tmpRef)"
End If
RunSQL strCmd, True
'oCmd.CommandText = strCmd
'oCmd.Execute
    
UpdateVer1 = True
Exit Function

ErrExit:
UpdateVer1 = False
Exit Function
End Function


Private Function UpdateVer2(ByRef oCmd As Object) As Boolean
On Error GoTo ErrExit
Dim strCmd As String

If gUseMSDE Then
    strCmd = "CREATE TABLE " & gFileMemberGroup & " (GroupName CHAR(25), RentalDiscAmt FLOAT, RentalDiscPerHour FLOAT, RentalDiscPerc FLOAT, SalesDiscAmt FLOAT, SalesDiscPerc FLOAT, DiscMode1 CHAR(35), DiscMode2 CHAR(35), DiscMode3 CHAR(35), DiscMode4 CHAR(35), DiscMode5 CHAR(35), Descrip CHAR(255))"
Else
    strCmd = "CREATE TABLE " & gFileMemberGroup & " (GroupName TEXT(25), RentalDiscAmt DOUBLE, RentalDiscPerHour DOUBLE, RentalDiscPerc DOUBLE, SalesDiscAmt DOUBLE, SalesDiscPerc DOUBLE, DiscMode1 TEXT(35), DiscMode2 TEXT(35), DiscMode3 TEXT(35), DiscMode4 TEXT(35), DiscMode5 TEXT(35), Descrip TEXT(255))"
End If
RunSQL strCmd, True
'oCmd.CommandText = strCmd
'oCmd.Execute
    
If gUseMSDE Then
    strCmd = "CREATE UNIQUE INDEX GroupName ON " & gFileMemberGroup & " (GroupName)"
Else
    strCmd = "CREATE UNIQUE INDEX GroupName ON " & gFileMemberGroup & " (GroupName) WITH PRIMARY"
End If
RunSQL strCmd, True
'oCmd.CommandText = strCmd
'oCmd.Execute
    
If gUseMSDE Then
    strCmd = "CREATE TABLE " & gFileResource & " (ResourceName CHAR(25), AdjAmt FLOAT, AdjPerc FLOAT, AdjPerHour FLOAT, Descrip CHAR(255), Field1 CHAR(25), Field2 CHAR(25), Field3 CHAR(25), Field4 CHAR(25), Field5 CHAR(25), NumberFrom INT, NumberTo INT)"
Else
    strCmd = "CREATE TABLE " & gFileResource & " (ResourceName TEXT(25), AdjAmt DOUBLE, AdjPerc DOUBLE, AdjPerHour DOUBLE, Descrip TEXT(255), Field1 TEXT(25), Field2 TEXT(25), Field3 TEXT(25), Field4 TEXT(25), Field5 TEXT(25), NumberFrom LONG, NumberTo LONG)"
End If
RunSQL strCmd, True
'oCmd.CommandText = strCmd
'oCmd.Execute
    
If gUseMSDE Then
    strCmd = "CREATE UNIQUE INDEX ResourceName ON " & gFileResource & " (ResourceName)"
Else
    strCmd = "CREATE UNIQUE INDEX ResourceName ON " & gFileResource & " (ResourceName) WITH PRIMARY"
End If
RunSQL strCmd, True
'oCmd.CommandText = strCmd
'oCmd.Execute
    
UpdateVer2 = True
Exit Function

ErrExit:
UpdateVer2 = False
Exit Function
End Function

Private Function UpdateVer3(ByRef oCmd As Object) As Boolean
On Error GoTo ErrExit
Dim strCmd As String

If gUseMSDE Then
    strCmd = "CREATE TABLE " & gFileQueue & " (DateIn DATETIME, Display CHAR(25), [User] CHAR(25), LineOrder INT)"
Else
    strCmd = "CREATE TABLE " & gFileQueue & " (DateIn DATE, Display TEXT(25), User TEXT(25), LineOrder LONG)"
End If
RunSQL strCmd, True
'oCmd.CommandText = strCmd
'oCmd.Execute

If gUseMSDE Then
    strCmd = "CREATE INDEX DateIn ON " & gFileQueue & " (DateIn)"
Else
    strCmd = "CREATE INDEX DateIn ON " & gFileQueue & " (DateIn)"
End If
RunSQL strCmd, True
'oCmd.CommandText = strCmd
'oCmd.Execute

If gUseMSDE Then
    strCmd = "CREATE TABLE " & gFilePrepaidType & " (TypeName CHAR(25), Price FLOAT, PrepaidValue FLOAT, PrepaidType CHAR(1), [Default] CHAR(1))"
Else
    strCmd = "CREATE TABLE " & gFilePrepaidType & " (TypeName TEXT(25), Price DOUBLE, PrepaidValue DOUBLE, PrepaidType TEXT(1), Default TEXT(1))"
End If
RunSQL strCmd, True
'oCmd.CommandText = strCmd
'oCmd.Execute
    
If gUseMSDE Then
    strCmd = "CREATE UNIQUE INDEX TypeName ON " & gFilePrepaidType & " (TypeName)"
Else
    strCmd = "CREATE UNIQUE INDEX TypeName ON " & gFilePrepaidType & " (TypeName) WITH PRIMARY"
End If
RunSQL strCmd, True
'oCmd.CommandText = strCmd
'oCmd.Execute

If gUseMSDE Then
    strCmd = "ALTER TABLE " & gFileMemberList & " ADD DateValidFrom DATETIME, DateValidTo DATETIME, DateCreated DATETIME"
Else
    strCmd = "ALTER TABLE " & gFileMemberList & " ADD COLUMN DateValidFrom DATE, DateValidTo DATE, DateCreated DATE"
End If
RunSQL strCmd, True
'oCmd.CommandText = strCmd
'oCmd.Execute

If gUseMSDE Then
    strCmd = "ALTER TABLE " & gFileRateList & " ADD RoundDuration FLOAT, GracePeriod FLOAT, MinDuration FLOAT, MinAmount FLOAT"
Else
    strCmd = "ALTER TABLE " & gFileRateList & " ADD COLUMN RoundDuration DOUBLE, GracePeriod DOUBLE, MinDuration DOUBLE, MinAmount DOUBLE"
End If
RunSQL strCmd, True
'oCmd.CommandText = strCmd
'oCmd.Execute

strCmd = "UPDATE " & gFileRateList & " SET Member='0' WHERE RateType='GAMING' OR RateType='INTERNET' OR RateType='RENTAL' OR RateType='GAMES'"
RunSQL strCmd, True
'oCmd.CommandText = strCmd
'oCmd.Execute

UpdateVer3 = True
Exit Function

ErrExit:
UpdateVer3 = False
Exit Function
End Function

Private Function UpdateVer4(ByRef oCmd As Object) As Boolean
On Error GoTo ErrExit
Dim strCmd As String

Dim n As Long
If gUseMSDE Then
    strCmd = "CREATE TABLE " & gFileState & " ("
    For n = 1 To 50
        strCmd = strCmd & "c" & CStr(n) & " CHAR(100), "
    Next
    strCmd = strCmd & " Ghosted INT )"
Else
    strCmd = "CREATE TABLE " & gFileState & " ("
    For n = 1 To 50
        strCmd = strCmd & "c" & CStr(n) & " TEXT(100), "
    Next
    strCmd = strCmd & " Ghosted LONG )"
End If
RunSQL strCmd, True
'oCmd.CommandText = strCmd
'oCmd.Execute

UpdateVer4 = True
Exit Function

ErrExit:
UpdateVer4 = False
Exit Function
End Function


Private Function UpdateVer5(ByRef oCmd As Object) As Boolean
On Error GoTo ErrExit
Dim strCmd As String

If gUseMSDE Then
    strCmd = "CREATE TABLE " & gFilePrepaidTickets & " (UserName CHAR(25), FirstName CHAR(25), LastName CHAR(25), Password CHAR(25), PrepaidType CHAR(25), Price FLOAT, PrepaidValue FLOAT)"
Else
    strCmd = "CREATE TABLE " & gFilePrepaidTickets & " (UserName TEXT(25), FirstName TEXT(25), LastName TEXT(25), Password TEXT(25), PrepaidType TEXT(25), Price DOUBLE, PrepaidValue DOUBLE)"
End If
RunSQL strCmd, True
'oCmd.CommandText = strCmd
'oCmd.Execute

If gUseMSDE Then
    strCmd = "CREATE UNIQUE INDEX UserName ON " & gFilePrepaidTickets & " (UserName)"
Else
    strCmd = "CREATE UNIQUE INDEX UserName ON " & gFilePrepaidTickets & " (UserName) WITH PRIMARY"
End If
RunSQL strCmd, True
'oCmd.CommandText = strCmd
'oCmd.Execute
UpdateVer5 = True
Exit Function

ErrExit:
UpdateVer5 = False
Exit Function
End Function

Private Function UpdateVer6(ByRef oCmd As Object) As Boolean
On Error GoTo ErrExit
Dim strCmd As String

If gUseMSDE Then
    strCmd = "ALTER TABLE " & gFileState & " ADD IPAddress CHAR(25), PortNumber CHAR(25), LastConnect DATETIME"
Else
    strCmd = "ALTER TABLE " & gFileState & " ADD COLUMN IPAddress TEXT(25), PortNumber TEXT(25), LastConnect DATE"
End If
RunSQL strCmd, True
'oCmd.CommandText = strCmd
'oCmd.Execute

If gUseMSDE Then
    strCmd = "CREATE TABLE " & gFileStationHostMap & " (Station CHAR(25), HostName CHAR(25), IPAddress CHAR(25), PortNumber CHAR(25), Address1 CHAR(25), Address2 CHAR(25), Address3 CHAR(25), Address4 CHAR(26))"
Else
    strCmd = "CREATE TABLE " & gFileStationHostMap & " (Station TEXT(25), HostName TEXT(25), IPAddress TEXT(25), PortNumber TEXT(25), Address1 TEXT(25), Address2 TEXT(25), Address3 TEXT(25), Address4 TEXT(26))"
End If
RunSQL strCmd, True
'oCmd.CommandText = strCmd
'oCmd.Execute

If gUseMSDE Then
    strCmd = "CREATE UNIQUE INDEX Station ON " & gFileStationHostMap & " (Station)"
Else
    strCmd = "CREATE UNIQUE INDEX Station ON " & gFileStationHostMap & " (Station) WITH PRIMARY"
End If
RunSQL strCmd, True
'oCmd.CommandText = strCmd
'oCmd.Execute

On Error Resume Next
If gUseMSDE Then
    strCmd = "DROP INDEX " & gFileState & ".c1"
Else
    strCmd = "DROP INDEX c1 ON " & gFileState
End If
RunSQL strCmd, True
'oCmd.CommandText = strCmd
'oCmd.Execute

If gUseMSDE Then
    strCmd = "CREATE UNIQUE INDEX c1 ON " & gFileState & " (c1)"
Else
    strCmd = "CREATE UNIQUE INDEX c1 ON " & gFileState & " (c1) WITH PRIMARY"
End If
RunSQL strCmd, True
'oCmd.CommandText = strCmd
'oCmd.Execute

UpdateVer6 = True
Exit Function

ErrExit:
UpdateVer6 = False
Exit Function
End Function

Private Function UpdateVer7(ByRef oCmd As Object) As Boolean
On Error GoTo ErrExit
Dim strCmd As String

If gUseMSDE Then
    strCmd = "CREATE INDEX LogDate ON " & gFileClientUsage & " (LogDate)"
Else
    strCmd = "CREATE INDEX LogDate ON " & gFileClientUsage & " (LogDate)"
End If
RunSQL strCmd, True
'oCmd.CommandText = strCmd
'oCmd.Execute
UpdateVer7 = True
Exit Function

ErrExit:
UpdateVer7 = False
Exit Function
End Function

Private Function UpdateVer8(ByRef oCmd As Object) As Boolean
On Error GoTo ErrExit
Dim strCmd As String

If gUseMSDE Then
    strCmd = "ALTER TABLE " & gFilePrepaidTickets & " ADD Printed INT, Created DATETIME"
Else
    strCmd = "ALTER TABLE " & gFilePrepaidTickets & " ADD COLUMN Printed LONG, Created DATE"
End If
RunSQL strCmd, True
'oCmd.CommandText = strCmd
'oCmd.Execute

UpdateVer8 = True
Exit Function

ErrExit:
UpdateVer8 = False
Exit Function
End Function

Private Function UpdateVer9(ByRef oCmd As Object) As Boolean
On Error Resume Next
'UpdateVer1 oCmd
UpdateVer9 = True
End Function

Private Function UpdateVer10(ByRef oCmd As Object) As Boolean
On Error GoTo ErrExit
Dim strCmd As String

If gUseMSDE Then
    strCmd = "ALTER TABLE " & gFileState & " ADD MACAddress CHAR(25), HostName CHAR(25)"
Else
    strCmd = "ALTER TABLE " & gFileState & " ADD COLUMN MACAddress TEXT(25), HostName TEXT(25)"
End If
RunSQL strCmd, True
'oCmd.CommandText = strCmd
'oCmd.Execute

If gUseMSDE Then
    strCmd = "CREATE UNIQUE INDEX MACAddress ON " & gFileState & " (MACAddress)"
Else
    strCmd = "CREATE UNIQUE INDEX MACAddress ON " & gFileState & " (MACAddress)"
End If
RunSQL strCmd, True
'oCmd.CommandText = strCmd
'oCmd.Execute

If gUseMSDE Then
    strCmd = "CREATE UNIQUE INDEX HostName ON " & gFileState & " (HostName)"
Else
    strCmd = "CREATE UNIQUE INDEX HostName ON " & gFileState & " (HostName)"
End If
RunSQL strCmd, True
'oCmd.CommandText = strCmd
'oCmd.Execute

UpdateVer10 = True
Exit Function

ErrExit:
UpdateVer10 = False
Exit Function
End Function

Private Function UpdateVer11(ByRef oCmd As Object) As Boolean
On Error GoTo ErrExit
Dim strCmd As String
If gUseMSDE Then
    strCmd = "ALTER TABLE " & gFileReceivables & " ADD tmpRef INT, tmpStation CHAR(4)"
Else
    strCmd = "ALTER TABLE " & gFileReceivables & " ADD COLUMN tmpRef LONG, tmpStation TEXT(4)"
End If
RunSQL strCmd, True
'oCmd.CommandText = strCmd
'oCmd.Execute
    
If gUseMSDE Then
    strCmd = "CREATE INDEX tmpStation ON " & gFileReceivables & " (tmpStation)"
Else
    strCmd = "CREATE INDEX tmpStation ON " & gFileReceivables & " (tmpStation)"
End If
RunSQL strCmd, True
'oCmd.CommandText = strCmd
'oCmd.Execute

If gUseMSDE Then
    strCmd = "CREATE INDEX tmpRef ON " & gFileReceivables & " (tmpRef)"
Else
    strCmd = "CREATE INDEX tmpRef ON " & gFileReceivables & " (tmpRef)"
End If
RunSQL strCmd, True
'oCmd.CommandText = strCmd
'oCmd.Execute

If gUseMSDE Then
    strCmd = "ALTER TABLE " & gFileExpenses & " ADD tmpRef INT, tmpStation CHAR(4)"
Else
    strCmd = "ALTER TABLE " & gFileExpenses & " ADD COLUMN tmpRef LONG, tmpStation TEXT(4)"
End If
RunSQL strCmd, True
'oCmd.CommandText = strCmd
'oCmd.Execute
    
If gUseMSDE Then
    strCmd = "CREATE INDEX tmpStation ON " & gFileExpenses & " (tmpStation)"
Else
    strCmd = "CREATE INDEX tmpStation ON " & gFileExpenses & " (tmpStation)"
End If
RunSQL strCmd, True
'oCmd.CommandText = strCmd
'oCmd.Execute

If gUseMSDE Then
    strCmd = "CREATE INDEX tmpRef ON " & gFileExpenses & " (tmpRef)"
Else
    strCmd = "CREATE INDEX tmpRef ON " & gFileExpenses & " (tmpRef)"
End If
RunSQL strCmd, True
'oCmd.CommandText = strCmd
'oCmd.Execute

UpdateVer11 = True
Exit Function

ErrExit:
UpdateVer11 = False
Exit Function
End Function

Private Function UpdateVer12(ByRef oCmd As Object) As Boolean
On Error GoTo ErrExit
Dim strCmd As String
If gUseMSDE Then
    strCmd = "ALTER TABLE " & gFileRateList & " ADD Command CHAR(255)"
Else
    strCmd = "ALTER TABLE " & gFileRateList & " ADD COLUMN Command TEXT(255)"
End If
RunSQL strCmd, True
'oCmd.CommandText = strCmd
'oCmd.Execute
    
If gUseMSDE Then
    strCmd = "CREATE INDEX Command ON " & gFileRateList & " (Command)"
Else
    strCmd = "CREATE INDEX Command ON " & gFileRateList & " (Command)"
End If
RunSQL strCmd, True
'oCmd.CommandText = strCmd
'oCmd.Execute

UpdateVer12 = True
Exit Function

ErrExit:
UpdateVer12 = False
Exit Function
End Function

Private Function UpdateVer13(ByRef oCmd As Object) As Boolean
On Error GoTo ErrExit
Dim strCmd As String

If gUseMSDE Then
    strCmd = "CREATE TABLE " & gFileFirewallList & " (PortRestriction CHAR(28), PortType CHAR(1))"
Else
    strCmd = "CREATE TABLE " & gFileFirewallList & " (PortRestriction TEXT(28), PortType TEXT(1))"
End If
RunSQL strCmd, True
'oCmd.CommandText = strCmd
'oCmd.Execute

UpdateVer13 = True
Exit Function

ErrExit:
UpdateVer13 = False
Exit Function
End Function

Private Function UpdateVer14(ByRef oCmd As Object) As Boolean
On Error GoTo ErrExit
Dim strCmd As String
On Error Resume Next
If gUseMSDE Then
    strCmd = "DROP INDEX gFileFirewallList.Command"
Else
    strCmd = "DROP INDEX Command ON " & gFileFirewallList
End If
RunSQL strCmd, True
'oCmd.CommandText = strCmd
'oCmd.Execute
    
On Error GoTo ErrExit
If gUseMSDE Then
    strCmd = "CREATE UNIQUE INDEX PortRestriction ON " & gFileFirewallList & " (PortRestriction)"
Else
    strCmd = "CREATE UNIQUE INDEX PortRestriction ON " & gFileFirewallList & " (PortRestriction) WITH PRIMARY"
End If
RunSQL strCmd, True
'oCmd.CommandText = strCmd
'oCmd.Execute

UpdateVer14 = True
Exit Function

ErrExit:
UpdateVer14 = False
Exit Function
End Function

Private Function UpdateVer15(ByRef oCmd As Object) As Boolean
On Error GoTo ErrExit
Dim strCmd As String

If gUseMSDE Then
    strCmd = "CREATE TABLE " & gFileFirewallDetList & " (PortRestriction CHAR(28), RateType CHAR(50))"
Else
    strCmd = "CREATE TABLE " & gFileFirewallDetList & " (PortRestriction TEXT(28), RateType TEXT(50))"
End If
RunSQL strCmd, True
'oCmd.CommandText = strCmd
'oCmd.Execute

If gUseMSDE Then
    strCmd = "CREATE INDEX PortRestriction ON " & gFileFirewallDetList & " (PortRestriction)"
Else
    strCmd = "CREATE INDEX PortRestriction ON " & gFileFirewallDetList & " (PortRestriction)"
End If
RunSQL strCmd, True
'oCmd.CommandText = strCmd
'oCmd.Execute

If gUseMSDE Then
    strCmd = "CREATE INDEX RateType ON " & gFileFirewallDetList & " (RateType)"
Else
    strCmd = "CREATE INDEX RateType ON " & gFileFirewallDetList & " (RateType)"
End If
RunSQL strCmd, True
'oCmd.CommandText = strCmd
'oCmd.Execute

UpdateVer15 = True
Exit Function

ErrExit:
UpdateVer15 = False
Exit Function
End Function

Private Function UpdateVer16(ByRef oCmd As Object) As Boolean
On Error GoTo ErrExit
Dim strCmd As String

If gUseMSDE Then
    strCmd = "CREATE TABLE " & gFileFirewallOutList & " (PortRestriction CHAR(28), PortType CHAR(1))"
Else
    strCmd = "CREATE TABLE " & gFileFirewallOutList & " (PortRestriction TEXT(28), PortType TEXT(1))"
End If
RunSQL strCmd, True
'oCmd.CommandText = strCmd
'oCmd.Execute

If gUseMSDE Then
    strCmd = "CREATE UNIQUE INDEX PortRestriction ON " & gFileFirewallOutList & " (PortRestriction)"
Else
    strCmd = "CREATE UNIQUE INDEX PortRestriction ON " & gFileFirewallOutList & " (PortRestriction) WITH PRIMARY"
End If
RunSQL strCmd, True
'oCmd.CommandText = strCmd
'oCmd.Execute

UpdateVer16 = True
Exit Function

ErrExit:
UpdateVer16 = False
Exit Function
End Function

Private Function UpdateVer17(ByRef oCmd As Object) As Boolean
On Error GoTo ErrExit
Dim strCmd As String

If gUseMSDE Then
    strCmd = "ALTER TABLE " & gFileFirewallList & " ADD BlockTimedout CHAR(1)"
Else
    strCmd = "ALTER TABLE " & gFileFirewallList & " ADD COLUMN BlockTimedout TEXT(1)"
End If
RunSQL strCmd, True
'oCmd.CommandText = strCmd
'oCmd.Execute

strCmd = "CREATE INDEX BlockTimedout ON " & gFileFirewallList & " (BlockTimedout)"
RunSQL strCmd, True
'oCmd.CommandText = strCmd
'oCmd.Execute

UpdateVer17 = True
Exit Function

ErrExit:
UpdateVer17 = False
Exit Function
End Function

Private Function UpdateVer18(ByRef oCmd As Object) As Boolean
On Error GoTo ErrExit
Dim strCmd As String
If gUseMSDE Then
    strCmd = "ALTER TABLE " & gFileFirewallList & " ADD Descrip CHAR(255)"
Else
    strCmd = "ALTER TABLE " & gFileFirewallList & " ADD COLUMN Descrip TEXT(255)"
End If
RunSQL strCmd, True
'oCmd.CommandText = strCmd
'oCmd.Execute

UpdateVer18 = True
Exit Function

ErrExit:
UpdateVer18 = False
Exit Function
End Function

Private Function UpdateVer19(ByRef oCmd As Object) As Boolean
On Error GoTo ErrExit
Dim strCmd As String
If gUseMSDE Then
    strCmd = "ALTER TABLE " & gFileRateList2 & " ADD StockQty FLOAT, ReorderLevel FLOAT"
Else
    strCmd = "ALTER TABLE " & gFileRateList2 & " ADD COLUMN StockQty DOUBLE, ReorderLevel DOUBLE"
End If
RunSQL strCmd, True
'oCmd.CommandText = strCmd
'oCmd.Execute
If gUseMSDE Then
    strCmd = "ALTER TABLE " & gFileSales & " ADD Notes CHAR(255), AmtBefore FLOAT, AmtAfter FLOAT"
Else
    strCmd = "ALTER TABLE " & gFileSales & " ADD COLUMN Notes TEXT(255), AmtBefore DOUBLE, AmtAfter DOUBLE"
End If
RunSQL strCmd, True
'oCmd.CommandText = strCmd
'oCmd.Execute
If gUseMSDE Then
    strCmd = "ALTER TABLE " & gFileExpenses & " ADD Notes CHAR(255), AmtBefore FLOAT, AmtAfter FLOAT"
Else
    strCmd = "ALTER TABLE " & gFileExpenses & " ADD COLUMN Notes TEXT(255), AmtBefore DOUBLE, AmtAfter DOUBLE"
End If
RunSQL strCmd, True
'oCmd.CommandText = strCmd
'oCmd.Execute
If gUseMSDE Then
    strCmd = "ALTER TABLE " & gFileReceivables & " ADD Notes CHAR(255), AmtBefore FLOAT, AmtAfter FLOAT"
Else
    strCmd = "ALTER TABLE " & gFileReceivables & " ADD COLUMN Notes TEXT(255), AmtBefore DOUBLE, AmtAfter DOUBLE"
End If
RunSQL strCmd, True
'oCmd.CommandText = strCmd
'oCmd.Execute
If gUseMSDE Then
    strCmd = "ALTER TABLE " & gFileMemberPayments & " ADD Notes CHAR(255), AmtBefore FLOAT, AmtAfter FLOAT"
Else
    strCmd = "ALTER TABLE " & gFileMemberPayments & " ADD COLUMN Notes TEXT(255), AmtBefore DOUBLE, AmtAfter DOUBLE"
End If
RunSQL strCmd, True
'oCmd.CommandText = strCmd
'oCmd.Execute

UpdateVer19 = True
Exit Function

ErrExit:
UpdateVer19 = False
Exit Function
End Function

Private Function UpdateVer20(ByRef oCmd As Object) As Boolean
On Error GoTo ErrExit
Dim strCmd As String
If gUseMSDE Then
    strCmd = "ALTER TABLE " & gFileSales & " ADD MP CHAR(1)"
Else
    strCmd = "ALTER TABLE " & gFileSales & " ADD COLUMN MP TEXT(1)"
End If
RunSQL strCmd, True
'oCmd.CommandText = strCmd
'oCmd.Execute
If gUseMSDE Then
    strCmd = "ALTER TABLE " & gFileExpenses & " ADD MP CHAR(1)"
Else
    strCmd = "ALTER TABLE " & gFileExpenses & " ADD COLUMN MP TEXT(1)"
End If
RunSQL strCmd, True
'oCmd.CommandText = strCmd
'oCmd.Execute
If gUseMSDE Then
    strCmd = "ALTER TABLE " & gFileReceivables & " ADD MP CHAR(1)"
Else
    strCmd = "ALTER TABLE " & gFileReceivables & " ADD COLUMN MP TEXT(1)"
End If
RunSQL strCmd, True
'oCmd.CommandText = strCmd
'oCmd.Execute
If gUseMSDE Then
    strCmd = "ALTER TABLE " & gFileMemberPayments & " ADD MP CHAR(1)"
Else
    strCmd = "ALTER TABLE " & gFileMemberPayments & " ADD COLUMN MP TEXT(1)"
End If
RunSQL strCmd, True
'oCmd.CommandText = strCmd
'oCmd.Execute

UpdateVer20 = True
Exit Function

ErrExit:
UpdateVer20 = False
Exit Function
End Function

Private Function UpdateVer21(ByRef oCmd As Object) As Boolean
On Error GoTo ErrExit
Dim strCmd As String
If gUseMSDE Then
    strCmd = "ALTER TABLE " & gFileRateList2 & " ADD AvgCost FLOAT"
Else
    strCmd = "ALTER TABLE " & gFileRateList2 & " ADD COLUMN AvgCost DOUBLE"
End If
RunSQL strCmd, True
'oCmd.CommandText = strCmd
'oCmd.Execute

UpdateVer21 = True
Exit Function

ErrExit:
UpdateVer21 = False
Exit Function
End Function

Private Function UpdateVer22(ByRef oCmd As Object) As Boolean
On Error GoTo ErrExit
Dim strCmd As String
If gUseMSDE Then
    strCmd = "ALTER TABLE " & gFileState & " ADD GuestCount FLOAT"
Else
    strCmd = "ALTER TABLE " & gFileState & " ADD COLUMN GuestCount DOUBLE"
End If
RunSQL strCmd, True
'oCmd.CommandText = strCmd
'oCmd.Execute

UpdateVer22 = True
Exit Function

ErrExit:
UpdateVer22 = False
Exit Function
End Function

Private Function UpdateVer23(ByRef oCmd As Object) As Boolean
On Error GoTo ErrExit
Dim strCmd As String
UpdateVer23 = True
Exit Function

ErrExit:
UpdateVer23 = False
Exit Function
End Function

Private Function UpdateVer24(ByRef oCmd As Object) As Boolean
On Error GoTo ErrExit
Dim strCmd As String
If gUseMSDE Then
    strCmd = "ALTER TABLE " & gFileRateList & " ADD IconPath CHAR(200)"
Else
    strCmd = "ALTER TABLE " & gFileRateList & " ADD COLUMN IconPath TEXT(200)"
End If
RunSQL strCmd, True
'oCmd.CommandText = strCmd
'oCmd.Execute

UpdateVer24 = True
Exit Function

ErrExit:
UpdateVer24 = False
Exit Function
End Function

Private Function UpdateVer25(ByRef oCmd As Object) As Boolean
On Error GoTo ErrExit
Dim strCmd As String

If gUseMSDE Then
    strCmd = "CREATE TABLE " & gFileIconList & " (StatusType CHAR(28), IconPath CHAR(250))"
Else
    strCmd = "CREATE TABLE " & gFileIconList & " (StatusType TEXT(28), IconPath TEXT(250))"
End If
RunSQL strCmd, True
'oCmd.CommandText = strCmd
'oCmd.Execute

If gUseMSDE Then
    strCmd = "CREATE UNIQUE INDEX StatusType ON " & gFileIconList & " (StatusType)"
Else
    strCmd = "CREATE UNIQUE INDEX StatusType ON " & gFileIconList & " (StatusType) WITH PRIMARY"
End If
RunSQL strCmd, True
'oCmd.CommandText = strCmd
'oCmd.Execute

RunSQL "INSERT INTO " & gFileIconList & "(StatusType) VALUES (" & goDBConn.DataEncode("_Start", "S") & ")"
RunSQL "INSERT INTO " & gFileIconList & "(StatusType) VALUES (" & goDBConn.DataEncode("_Stop", "S") & ")"
RunSQL "INSERT INTO " & gFileIconList & "(StatusType) VALUES (" & goDBConn.DataEncode("_Warning", "S") & ")"
RunSQL "INSERT INTO " & gFileIconList & "(StatusType) VALUES (" & goDBConn.DataEncode("_Prepaid", "S") & ")"
RunSQL "INSERT INTO " & gFileIconList & "(StatusType) VALUES (" & goDBConn.DataEncode("_Member", "S") & ")"
RunSQL "INSERT INTO " & gFileIconList & "(StatusType) VALUES (" & goDBConn.DataEncode("_Open", "S") & ")"
RunSQL "INSERT INTO " & gFileIconList & "(StatusType) VALUES (" & goDBConn.DataEncode("_Pause", "S") & ")"
RunSQL "INSERT INTO " & gFileIconList & "(StatusType) VALUES (" & goDBConn.DataEncode("_Free", "S") & ")"
RunSQL "INSERT INTO " & gFileIconList & "(StatusType) VALUES (" & goDBConn.DataEncode("_Repair", "S") & ")"

UpdateVer25 = True
Exit Function

ErrExit:
UpdateVer25 = False
Exit Function
End Function

Private Function UpdateVer26(ByRef oCmd As Object) As Boolean
On Error GoTo ErrExit
Dim strCmd As String
If gUseMSDE Then
    strCmd = "ALTER TABLE " & gFileRateList & " ADD PlayAllUCan CHAR(200)"
Else
    strCmd = "ALTER TABLE " & gFileRateList & " ADD COLUMN PlayAllUCan TEXT(200)"
End If
RunSQL strCmd, True
'oCmd.CommandText = strCmd
'oCmd.Execute

UpdateVer26 = True
Exit Function

ErrExit:
UpdateVer26 = False
Exit Function
End Function

Private Function UpdateVer27(ByRef oCmd As Object) As Boolean
On Error GoTo ErrExit
Dim strCmd As String
If gUseMSDE Then
    strCmd = "ALTER TABLE " & gFileRateList & " ADD "
    strCmd = strCmd & "FreeMin Char(50),"
    'strCmd = strCmd & "GracePeriod Char(50),"
    strCmd = strCmd & "MaxFreeMin Char(50),"
    strCmd = strCmd & "MinAmt Char(50),"
    strCmd = strCmd & "MinTime Char(50),"
    strCmd = strCmd & "RoundAmount Char(50),"
    strCmd = strCmd & "RoundBy Char(50)"
Else
    strCmd = "ALTER TABLE " & gFileRateList & " ADD COLUMN "
    strCmd = strCmd & "FreeMin TEXT(50),"
    'strCmd = strCmd & "GracePeriod TEXT(50),"
    strCmd = strCmd & "MaxFreeMin TEXT(50),"
    strCmd = strCmd & "MinAmt TEXT(50),"
    strCmd = strCmd & "MinTime TEXT(50),"
    strCmd = strCmd & "RoundAmount TEXT(50),"
    strCmd = strCmd & "RoundBy TEXT(50)"
End If
RunSQL strCmd, True
'oCmd.CommandText = strCmd
'oCmd.Execute

UpdateVer27 = True
Exit Function

ErrExit:
UpdateVer27 = False
Exit Function
End Function

Private Function UpdateVer28(ByRef oCmd As Object) As Boolean
On Error GoTo ErrExit
Dim strCmd As String
If gUseMSDE Then
    strCmd = "ALTER TABLE " & gFileRateList & " ADD "
    strCmd = strCmd & "GracePd Char(50)"
Else
    strCmd = "ALTER TABLE " & gFileRateList & " ADD COLUMN "
    strCmd = strCmd & "GracePd TEXT(50)"
End If
RunSQL strCmd, True
'oCmd.CommandText = strCmd
'oCmd.Execute

UpdateVer28 = True
Exit Function

ErrExit:
UpdateVer28 = False
Exit Function
End Function

Private Function UpdateVer29(ByRef oCmd As Object) As Boolean
On Error GoTo ErrExit
Dim strCmd As String
If gUseMSDE Then
    strCmd = "ALTER TABLE " & gFileRateList & " ADD "
    strCmd = strCmd & "AutoOut Char(1),"
    strCmd = strCmd & "OpenOut Char(1)"
Else
    strCmd = "ALTER TABLE " & gFileRateList & " ADD COLUMN "
    strCmd = strCmd & "AutoOut TEXT(1),"
    strCmd = strCmd & "OpenOut TEXT(1)"
End If
RunSQL strCmd, True
'oCmd.CommandText = strCmd
'oCmd.Execute

UpdateVer29 = True
Exit Function

ErrExit:
UpdateVer29 = False
Exit Function
End Function

Private Function UpdateVer30(ByRef oCmd As Object) As Boolean
On Error GoTo ErrExit
Dim strCmd As String

RunSQL "INSERT INTO " & gFileIconList & "(StatusType) VALUES (" & goDBConn.DataEncode("_PlayAllUCan", "S") & ")"
RunSQL "INSERT INTO " & gFileIconList & "(StatusType) VALUES (" & goDBConn.DataEncode("_PrepaidWarning", "S") & ")"
RunSQL "INSERT INTO " & gFileIconList & "(StatusType) VALUES (" & goDBConn.DataEncode("_PrepaidStop", "S") & ")"
RunSQL "INSERT INTO " & gFileIconList & "(StatusType) VALUES (" & goDBConn.DataEncode("_MemberWarning", "S") & ")"
RunSQL "INSERT INTO " & gFileIconList & "(StatusType) VALUES (" & goDBConn.DataEncode("_MemberStop", "S") & ")"

UpdateVer30 = True
Exit Function

ErrExit:
UpdateVer30 = False
Exit Function
End Function

Private Function UpdateVer31(ByRef oCmd As Object) As Boolean
On Error GoTo ErrExit
Dim strCmd As String

If gUseMSDE Then
    strCmd = "CREATE TABLE " & gFilePrinterSpy & " ("
    strCmd = strCmd & "ID CHAR(30),"
    strCmd = strCmd & "DocName CHAR(250),"
    strCmd = strCmd & "PrintStatus CHAR(100),"
    strCmd = strCmd & "UserName CHAR(150),"
    strCmd = strCmd & "MachineName CHAR(150),"
    strCmd = strCmd & "PagesPrinted FLOAT,"
    strCmd = strCmd & "PrinterName CHAR(250),"
    strCmd = strCmd & "Login CHAR(50),"
    strCmd = strCmd & "EffectiveDate DateTime,"
    strCmd = strCmd & "TStamp DateTime"
    strCmd = strCmd & ")"
Else
    strCmd = "CREATE TABLE " & gFilePrinterSpy & " ("
    strCmd = strCmd & "ID TEXT(30),"
    strCmd = strCmd & "DocName TEXT(250),"
    strCmd = strCmd & "PrintStatus TEXT(100),"
    strCmd = strCmd & "UserName TEXT(150),"
    strCmd = strCmd & "MachineName TEXT(150),"
    strCmd = strCmd & "PagesPrinted DOUBLE,"
    strCmd = strCmd & "PrinterName TEXT(250),"
    strCmd = strCmd & "Login TEXT(50),"
    strCmd = strCmd & "EffectiveDate Date,"
    strCmd = strCmd & "TStamp Date"
    strCmd = strCmd & ")"
End If
RunSQL strCmd, True
'oCmd.CommandText = strCmd
'oCmd.Execute

If gUseMSDE Then
    strCmd = "CREATE UNIQUE INDEX ID ON " & gFilePrinterSpy & " (ID)"
Else
    strCmd = "CREATE UNIQUE INDEX ID ON " & gFilePrinterSpy & " (ID) WITH PRIMARY"
End If
RunSQL strCmd, True
'oCmd.CommandText = strCmd
'oCmd.Execute

If gUseMSDE Then
    strCmd = "CREATE INDEX MachineName ON " & gFilePrinterSpy & " (MachineName)"
Else
    strCmd = "CREATE INDEX MachineName ON " & gFilePrinterSpy & " (MachineName)"
End If
RunSQL strCmd, True
'oCmd.CommandText = strCmd
'oCmd.Execute

If gUseMSDE Then
    strCmd = "CREATE INDEX PrinterName ON " & gFilePrinterSpy & " (PrinterName)"
Else
    strCmd = "CREATE INDEX PrinterName ON " & gFilePrinterSpy & " (PrinterName)"
End If
RunSQL strCmd, True
'oCmd.CommandText = strCmd
'oCmd.Execute

If gUseMSDE Then
    strCmd = "CREATE INDEX Login ON " & gFilePrinterSpy & " (Login)"
Else
    strCmd = "CREATE INDEX Login ON " & gFilePrinterSpy & " (Login)"
End If
RunSQL strCmd, True
'oCmd.CommandText = strCmd
'oCmd.Execute

If gUseMSDE Then
    strCmd = "CREATE INDEX EffectiveDate ON " & gFilePrinterSpy & " (EffectiveDate)"
Else
    strCmd = "CREATE INDEX EffectiveDate ON " & gFilePrinterSpy & " (EffectiveDate)"
End If
RunSQL strCmd, True
'oCmd.CommandText = strCmd
'oCmd.Execute

If gUseMSDE Then
    strCmd = "CREATE INDEX TStamp ON " & gFilePrinterSpy & " (TStamp)"
Else
    strCmd = "CREATE INDEX TStamp ON " & gFilePrinterSpy & " (TStamp)"
End If
RunSQL strCmd, True
'oCmd.CommandText = strCmd
'oCmd.Execute

If gUseMSDE Then
    strCmd = "CREATE INDEX PrintStatus ON " & gFilePrinterSpy & " (PrintStatus)"
Else
    strCmd = "CREATE INDEX PrintStatus ON " & gFilePrinterSpy & " (PrintStatus)"
End If
RunSQL strCmd, True
'oCmd.CommandText = strCmd
'oCmd.Execute

UpdateVer31 = True
Exit Function

ErrExit:
UpdateVer31 = False
Exit Function
End Function

Private Function UpdateVer32(ByRef oCmd As Object) As Boolean
On Error GoTo ErrExit
Dim strCmd As String

RunSQL "UPDATE " & gFileRateList & " SET Member = 1 WHERE RateType = " & goDBConn.DataEncode("GAMING", "S")
RunSQL "UPDATE " & gFileRateList & " SET Member = 1 WHERE RateType = " & goDBConn.DataEncode("INTERNET", "S")
UpdateVer32 = True
Exit Function

ErrExit:
UpdateVer32 = False
Exit Function
End Function

Private Function UpdateVer33(ByRef oCmd As Object) As Boolean
On Error GoTo ErrExit
Dim strCmd As String
If gUseMSDE Then
    strCmd = "ALTER TABLE " & gFileUserList & " ADD "
    strCmd = strCmd & "RateType Char(50)"
Else
    strCmd = "ALTER TABLE " & gFileUserList & " ADD COLUMN "
    strCmd = strCmd & "RateType TEXT(50)"
End If
RunSQL strCmd, True
'oCmd.CommandText = strCmd
'oCmd.Execute

UpdateVer33 = True
Exit Function

ErrExit:
UpdateVer33 = False
Exit Function
End Function


Private Function UpdateVer34(ByRef oCmd As Object) As Boolean
On Error GoTo ErrExit
Dim strCmd As String

If gUseMSDE Then
    strCmd = "CREATE TABLE " & gFileBlockURLList & " (Pattern CHAR(250), LockMode CHAR(1))"
Else
    strCmd = "CREATE TABLE " & gFileBlockURLList & " (Pattern TEXT(250), LockMode TEXT(1))"
End If
RunSQL strCmd, True
'oCmd.CommandText = strCmd
'oCmd.Execute

If gUseMSDE Then
    strCmd = "CREATE UNIQUE INDEX Pattern ON " & gFileBlockURLList & " (Pattern)"
Else
    strCmd = "CREATE UNIQUE INDEX Pattern ON " & gFileBlockURLList & " (Pattern) WITH PRIMARY"
End If
RunSQL strCmd, True
'oCmd.CommandText = strCmd
'oCmd.Execute

UpdateVer34 = True
Exit Function

ErrExit:
UpdateVer34 = False
Exit Function
End Function

Private Sub DoUpdateVer(ByVal pVer As Long, ByRef oCmd As Object)
On Error Resume Next
If pVer < 1 Then
    UpdateVer1 oCmd
End If
If pVer < 2 Then
    UpdateVer2 oCmd
End If
If pVer < 3 Then
    UpdateVer3 oCmd
End If
If pVer < 4 Then
    UpdateVer4 oCmd
End If
If pVer < 5 Then
    UpdateVer5 oCmd
End If
If pVer < 6 Then
    UpdateVer6 oCmd
End If
If pVer < 7 Then
    UpdateVer7 oCmd
End If
If pVer < 8 Then
    UpdateVer8 oCmd
End If
If pVer < 9 Then
    UpdateVer9 oCmd
End If
If pVer < 10 Then
    UpdateVer10 oCmd
End If
If pVer < 11 Then
    UpdateVer11 oCmd
End If
If pVer < 12 Then
    UpdateVer12 oCmd
End If
If pVer < 13 Then
    UpdateVer13 oCmd
End If
If pVer < 14 Then
    UpdateVer14 oCmd
End If
If pVer < 15 Then
    UpdateVer15 oCmd
End If
'If pVer < 16 Then
'    UpdateVer16 oCmd
'End If
If pVer < 17 Then
    UpdateVer17 oCmd
End If
If pVer < 18 Then
    UpdateVer18 oCmd
End If
If pVer < 19 Then
    UpdateVer19 oCmd
End If
If pVer < 20 Then
    UpdateVer20 oCmd
End If
If pVer < 21 Then
    UpdateVer21 oCmd
End If
If pVer < 22 Then
    UpdateVer22 oCmd
End If
If pVer < 23 Then
    UpdateVer23 oCmd
End If
If pVer < 24 Then
    UpdateVer24 oCmd
End If
If pVer < 25 Then
    UpdateVer25 oCmd
End If
If pVer < 26 Then
    UpdateVer26 oCmd
End If
If pVer < 27 Then
    UpdateVer27 oCmd
End If
If pVer < 28 Then
    UpdateVer28 oCmd
End If
If pVer < 29 Then
    UpdateVer29 oCmd
End If
If pVer < 30 Then
    UpdateVer30 oCmd
End If
If pVer < 31 Then
    UpdateVer31 oCmd
End If
If pVer < 32 Then
    UpdateVer32 oCmd
End If
If pVer < 33 Then
    UpdateVer33 oCmd
End If
If pVer < 34 Then
    UpdateVer34 oCmd
End If
End Sub
