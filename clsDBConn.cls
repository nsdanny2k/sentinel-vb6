VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDBConn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public ISMSDE As Boolean

Private Const mServer As String = "(local)"
Private Const mUser As String = "sa"
Private Const mPassword As String = "password"
Private Const mOldPassword As String = ""
Private Const mLoginTimeout As Long = 60
Private Const mDatabaseName As String = "Sentinel"

Private mSQLServerPath As String
Private Const mSQLServerPath2 As String = "c:\mssql7\"
Private Const mSQLServerDataPath As String = "Data\"
Private Const mSQLServerBinPath As String = "Binn\"
Private Const mSQLServerFile As String = "sqlservr.exe"
Private Const mTempPath As String = "c:\temp\"
Private Const mDatapath As String = "Sentinel.mdf"
Private Const mDatapath2 As String = "Sentinel.ldf"

Dim moCmd As Object
Public Connection As Object

Private mvarDATABASEPASSWORD As String
Private mvarFileName As String
Private mvarPATH As String
Private mvarCONNECTION_STRING As String


Public Event Closed()
Public Event Connected()
Public Event Error(ByVal The_Message As String, ByVal ErrNum As Long)
Public Event Message(ByVal The_Message As String)

Private Function mSQLServerPath1() As String
On Error GoTo ErrExit
Dim sPath As String
sPath = GetSpecialFolder(PROGRAM_FILES)
If sPath = "" Then sPath = "c:\program files"
If Right(sPath, 1) <> "\" Then sPath = sPath & "\"
mSQLServerPath1 = sPath & "microsoft sql server\mssql\"
Exit Function
ErrExit:
mSQLServerPath1 = "c:\program files\microsoft sql server\mssql\"
End Function

Public Sub SetIsMSDE()
gUseMSDE = Not (InStr(1, mvarCONNECTION_STRING, "Microsoft Access", vbTextCompare) <> 0)
End Sub

'//For Standard Security
'oConn.Open "Provider=sqloledb;" & _
           "Data Source=myServerName;" & _
           "Initial Catalog=myDatabaseName;" & _
           "User Id=myUsername;" & _
           "Password=myPassword"
'//For a Trusted Connection
'oConn.Open "Provider=sqloledb;" & _
           "Data Source=myServerName;" & _
           "Initial Catalog=myDatabaseName;" & _
           "Integrated Security=SSPI"
'//To connect to a "Named Instance"
'oConn.Open "Provider=sqloledb;" & _
           "Data Source=myServerName\myInstanceName;" & _
           "Initial Catalog=myDatabaseName;" & _
           "User Id=myUsername;" & _
           "Password=myPassword"
'       Note: In order to connect to a SQL Server 2000 "named instance", you must have MDAC 2.6 (or greater) installed.
'//To Prompt user for username and password
'oConn.Provider = "sqloledb"
'oConn.Properties("Prompt") = adPromptAlways
'oConn.Open "Data Source=myServerName;" & _
           "Initial Catalog=myDatabaseName"
'//To connect to SQL Server running on the same computer
'oConn.Open "Provider=sqloledb;" & _
           "Data Source=(local);" & _
           "Initial Catalog=myDatabaseName;" & _
           "User ID=myUsername;" & _
           "Password=myPassword"
'//To connect to SQL Server running on a remote computer (via an IP address)
'oConn.Open "Provider=sqloledb;" & _
           "Network Library=DBMSSOCN;" & _
           "Data Source=xxx.xxx.xxx.xxx,1433;" & _
           "Initial Catalog=myDatabaseName;" & _
           "User ID=myUsername;" & _
           "Password=myPassword"
'       You can also add "Encrypt=yes" for encryption
'//DSN
'oConn.Open "DSN=mySystemDSN;" & _
           "Uid=myUsername;" & _
           "Pwd=myPassword"
'//File DSN
'oConn.Open "FILEDSN=c:\somepath\mydb.dsn;" & _
           "Uid=myUsername;" & _
           "Pwd=myPassword"
'//For Absolute Path
'oConn.Open "File Name=c:\somepath\myDatabaseName.udl"
'//For Relative Path
'oConn.Open "File Name=myDatabaseName.udl"


'//For standard security
'oConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
           "Data Source=c:\somepath\myDb.mdb;" & _
           "User Id=admin;" & _
           "Password="
'//If using a Workgroup (System Database)
'oConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
           "Data Source=c:\somepath\mydb.mdb;" & _
           "Jet OLEDB:System Database=MySystem.mdw", _
           "myUsername", "myPassword"
'    Note, remember to convert both the MDB and the MDW to the 4.0
'    database format when using the 4.0 OLE DB Provider.
'//If MDB has a database password
'oConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
           "Data Source=c:\somepath\mydb.mdb;" & _
           "Jet OLEDB:Database Password=MyDbPassword", _
           "myUsername", "myPassword"
'//If want to open up the MDB exclusively
'oConn.Mode = adModeShareExclusive
'oConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
           "Data Source=c:\somepath\myDb.mdb;" & _
           "User Id=admin;" & _
           "Password="
'// If MDB is located on a network share
'oConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
           "Data Source=\\myServer\myShare\myPath\myDb.mdb"
'//If you don't know the path to the MDB (using ASP)
'<%  ' ASP server-side code
'    oConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
              "Data Source=" & Server.MapPath(".") & "\myDb.mdb;" & _
              "User Id=admin;" & _
              "Password="
'%>
'//If you don't know the path to the MDB (using VB)
'oConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
           "Data Source=" & App.Path & "\myDb.mdb;" & _
           "User Id=admin;" & _
           "Password="
'//You can also open an Excel Spreadsheet using the JET OLE DB Provider
'oConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
           "Data Source=c:\somepath\mySpreadsheet.xls;" & _
           "Extended Properties=""Excel 8.0;HDR=Yes"""
'    Where "HDR=Yes" means that there is a header row in the cell range
'    (or named range), so the provider will not include the first row of the
'    selection into the RS.  If "HDR=No", then the provider will include
'    the first row of the cell range (or named ranged) into the RS.
'//You can also open a Text file using the JET OLE DB Provider
'oConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
       "Data Source=c:\somepath\;" & _
       "Extended Properties=""text;HDR=Yes;FMT=Delimited"""
'   Then open a RS based on a select on the actual file
'oRS.Open "Select * From MyTextFile.txt", oConn, _
         adOpenStatic, adLockReadOnly, adCmdText


Public Sub GetConnectionString(Optional ByVal pDontUseMSDE As Boolean = False)
On Error Resume Next
If Not gUseMSDE Or pDontUseMSDE Then
    mvarCONNECTION_STRING = "DRIVER={Microsoft Access Driver (*.mdb)};"
    mvarCONNECTION_STRING = mvarCONNECTION_STRING & "DBQ=" & Path_and_Filename & ";"
    mvarCONNECTION_STRING = mvarCONNECTION_STRING & "pwd=" & DatabasePassword & ";"
    mvarCONNECTION_STRING = mvarCONNECTION_STRING & "UID=Admin;"
Else
    mvarCONNECTION_STRING = "Provider=SQLOLEDB;"
    If Not gModeIsClient And gModeIsReader Then
        If gServerIP = "0" Then
            mvarCONNECTION_STRING = mvarCONNECTION_STRING & "Data Source=(local);"
        ElseIf InStr(1, gServerIP, ".") <> 0 Then
            mvarCONNECTION_STRING = mvarCONNECTION_STRING & "Network Library=DBMSSOCN;"
            mvarCONNECTION_STRING = mvarCONNECTION_STRING & "Data Source=" & gServerIP & ",1433;"
        Else
            mvarCONNECTION_STRING = mvarCONNECTION_STRING & "Data Source=" & gServerIP & ";"
        End If
    End If
        mvarCONNECTION_STRING = mvarCONNECTION_STRING & "Database=Sentinel;"
    mvarCONNECTION_STRING = mvarCONNECTION_STRING & "uid=sa;"
    mvarCONNECTION_STRING = mvarCONNECTION_STRING & "pwd=" & mPassword & ";"
End If
End Sub

Public Property Let Path(ByVal vData As String)
   mvarPATH = vData
   GetConnectionString
End Property

Public Property Get Path() As String
   Path = mvarPATH
End Property

Public Property Let FileName(ByVal vData As String)
   mvarFileName = vData
   GetConnectionString
End Property

Public Property Get FileName() As String
   FileName = mvarFileName
End Property

Public Property Let DatabasePassword(ByVal vData As String)
   mvarDATABASEPASSWORD = vData
   GetConnectionString
End Property

Public Property Get DatabasePassword() As String
   DatabasePassword = mvarDATABASEPASSWORD
End Property

Public Property Let Connection_String(ByVal vData As String)
   mvarCONNECTION_STRING = vData
End Property

Public Property Get Connection_String() As String
   Connection_String = mvarCONNECTION_STRING
End Property

Public Function Open_Database_Connection(Optional ByVal pUseMSDE As Boolean = True) As Boolean
   On Error GoTo OPEN_DATABASE_CONNECTION_ERROR
   If Not pUseMSDE Then
      GetConnectionString True
   End If
   If gSupportMSDE And pUseMSDE Then
      Me.CheckUseMSDE
   End If
   Set Connection = godb.CreateDBConnection
   SetFileAttrNormal Path_and_Filename
   If gSupportMSDE And pUseMSDE Then
      On Error Resume Next
   End If
   Connection.Open mvarCONNECTION_STRING
   If gSupportMSDE And pUseMSDE Then
      On Error GoTo OPEN_DATABASE_CONNECTION_ERROR
   End If
   DoEvents
   If Connection.State = 1 Then
      RaiseEvent Connected
      Open_Database_Connection = True
   Else
      If gSupportMSDE And pUseMSDE Then
        Me.GetConnectionString True
        Set Connection = godb.CreateDBConnection
        Connection.Open mvarCONNECTION_STRING
        DoEvents
        If Connection.State = 1 Then
           RaiseEvent Connected
           Open_Database_Connection = True
        Else
          Open_Database_Connection = False
        End If
      Else
        Open_Database_Connection = False
      End If
   End If
OPEN_DATABASE_CONNECTION_ERROR:
   If Err.number <> 0 And Connection.State <> 1 Then
      Open_Database_Connection = False
      MsgBox "OPEN DATABASE CONNECTION ERROR" & vbNewLine & _
            "Description : " & Err.Description & vbNewLine & _
            "ERROR #     : " & Str$(Err.number) & vbNewLine & _
            "LOCATION    : " & Path_and_Filename, _
            vbCritical + vbOKOnly, "Connection Error"
   Else
      Debug.Print "Database Opened : " & mvarCONNECTION_STRING
   End If
   SetIsMSDE
End Function

Private Sub Class_Initialize()
   mvarPATH = App.Path 'AddBackSlash(App.Path) & AddBackSlash("logs")
   'mvarPATH = LogDirPath(True)
   mvarFileName = gDatabaseName
   mvarDATABASEPASSWORD = gDatabasePwd
   GetConnectionString
End Sub

Private Sub Class_Terminate()
   mvarPATH = ""
   mvarFileName = ""
   mvarCONNECTION_STRING = ""
   Set Connection = Nothing
End Sub

Public Function Path_and_Filename() As String
   Dim strPath As String
   strPath = mvarPATH
   If Right(strPath, 1) <> "\" Then
        strPath = strPath & "\"
   End If
   strPath = strPath & mvarFileName
   Path_and_Filename = strPath
End Function

Private Function AddBackSlash(ByVal sPath As String) As String
   sPath = Trim$(sPath)
   If Len(sPath) > 0 Then
      sPath = sPath & IIf(Right$(sPath, 1) <> "\", "\", "")
   End If
   AddBackSlash = sPath
End Function

Public Function GetString(ByVal pRS As Object, ByVal parrFields As Variant) As String
On Error GoTo ErrExit
Dim n As Long
Dim u As Long
Dim l As Long
Dim strTmp As String
u = UBound(parrFields)
l = LBound(parrFields)
For n = l To u
    strTmp = strTmp & pRS(parrFields(n)) & gRecDelim
Next
strTmp = Left(strTmp, Len(strTmp) - 1)
GetString = strTmp
Exit Function
ErrExit:
GetString = ""
End Function

Public Function GetFieldNames(ByVal parrFields As Variant) As String
On Error GoTo ErrExit
Dim n As Long
Dim u As Long
Dim l As Long
Dim strTmp As String
u = UBound(parrFields)
l = LBound(parrFields)
For n = l To u
    If LCase(Trim(parrFields(n))) = "default" Then
        parrFields(n) = "[" & LCase(Trim(parrFields(n))) & "]"
    End If
    strTmp = strTmp & parrFields(n) & gCommaDelim
Next
strTmp = Left(strTmp, Len(strTmp) - 1)
GetFieldNames = strTmp
Exit Function
ErrExit:
GetFieldNames = ""
End Function

Public Function GetFieldValues2( _
    ByVal parrData As Variant, _
    ByVal parrFields As Variant, _
    ByVal parrTypes As Variant, _
    ByVal pType As String) As String
On Error GoTo ErrExit
Dim n As Long
Dim u As Long
Dim l As Long
Dim strTmp As String
u = UBound(parrFields)
l = LBound(parrFields)
If pType = "a" Then
    For n = l To u
        strTmp = strTmp & DataEncode(parrData(n), Left(parrTypes(n), 1)) & gCommaDelim
    Next
    strTmp = Left(strTmp, Len(strTmp) - 1)
ElseIf pType = "e" Then
    For n = l + 1 To u
        If LCase(Trim(parrFields(n))) = "default" Then
            parrFields(n) = "[" & LCase(Trim(parrFields(n))) & "]"
        End If
        strTmp = strTmp & parrFields(n) & " = " & DataEncode(parrData(n), Left(parrTypes(n), 1)) & gCommaDelim
    Next
    strTmp = Left(strTmp, Len(strTmp) - 1)
End If
GetFieldValues2 = strTmp
Exit Function
ErrExit:
GetFieldValues2 = ""
End Function

Public Function GetFieldValues( _
    ByVal pData As String, _
    ByVal parrFields As Variant, _
    ByVal parrTypes As Variant, _
    ByVal pType As String) As String
On Error GoTo ErrExit
Dim n As Long
Dim u As Long
Dim l As Long
Dim strTmp As String
Dim arrData As Variant
arrData = Split(pData, gRecDelim2)
u = UBound(parrFields)
l = LBound(parrFields)
If pType = "a" Then
    For n = l To u
        strTmp = strTmp & DataEncode(arrData(n), Left(parrTypes(n), 1)) & gCommaDelim
    Next
    strTmp = Left(strTmp, Len(strTmp) - 1)
ElseIf pType = "e" Then
    For n = l + 1 To u
        If LCase(Trim(parrFields(n))) = "default" Then
            parrFields(n) = "[" & LCase(Trim(parrFields(n))) & "]"
        End If
        strTmp = strTmp & parrFields(n) & " = " & DataEncode(arrData(n), Left(parrTypes(n), 1)) & gCommaDelim
    Next
    strTmp = Left(strTmp, Len(strTmp) - 1)
End If
GetFieldValues = strTmp
Exit Function
ErrExit:
GetFieldValues = ""
End Function

Public Function GetKeyFieldCriteria( _
    ByVal pData As String, _
    ByVal parrFields As Variant, _
    ByVal parrTypes As Variant, _
    ByVal pType As String) As String
On Error GoTo ErrExit
Dim strTmp As String
Dim arrData As Variant
arrData = Split(pData, gRecDelim2)
strTmp = parrFields(0) & " = " & DataEncode(arrData(0), Left(parrTypes(0), 1))
GetKeyFieldCriteria = strTmp
Exit Function
ErrExit:
GetKeyFieldCriteria = ""
End Function

Public Function FormatTimeField(ByVal FieldName As String) As String
If ((InStr(1, FieldName, " ") <> 0) And InStr(1, FieldName, "[") <> 0) Then FieldName = "" & FieldName & ""
If gUseMSDE Then
    'FormatField = "CONVERT(CHAR(19)," & FieldName & ",100)"
    'FormatField = "(CAST(DATEPART(Hh," & FieldName & ")) AS CHAR(2))+':'+(CAST(DATEPART(Mi," & FieldName & ")) AS CHAR(2))"
    'FormatField = "RIGHT('000'+RTRIM(CONVERT(CHAR(2),DATEPART(Hh," & FieldName & "))),2)+':'+RIGHT('000'+RTRIM(CONVERT(CHAR(2),DATEPART(Mi," & FieldName & "))),2)"
    'FormatField = "IF(DATEPART(Hh," & FieldName & ")>12,DATEPART(Hh," & FieldName & ")-12,DATEPART(Hh," & FieldName & "))"
    'FormatField = "(CASE WHEN (DATEPART(Hh," & FieldName & ")>12) THEN '1' ELSE '12')"
    '(CONVERT(INT,DATEPART(Hh," & FieldName & "))-25)<=0
    'FormatField = "(CASE WHEN 1=1 THEN 1 ELSE 12)"
    'FormatField = "(CASE WHEN DATEPART(HH," & FieldName & ") < 13 THEN 'AM' Else 'PM' END)"
    'FormatField = "(CASE WHEN DATEPART(HH," & FieldName & ") < 13 THEN DATEPART(HH," & FieldName & ") Else DATEPART(HH," & FieldName & ")-12 END)"
    'FormatField = "(CASE WHEN CONVERT(INT,DATEPART(Hh," & FieldName & "))>CONVERT(INT,12) THEN '1' ELSE '12')"
    FormatTimeField = "RIGHT('000'+RTRIM(CONVERT(CHAR(2),(CASE WHEN DATEPART(HH," & FieldName & ") < 13 THEN DATEPART(HH," & FieldName & ") Else DATEPART(HH," & FieldName & ")-12 END))),2)+':'+RIGHT('000'+RTRIM(CONVERT(CHAR(2),DATEPART(Mi," & FieldName & "))),2)+' '+(CASE WHEN DATEPART(HH," & FieldName & ") < 13 THEN 'am' Else 'pm' END)"
Else
    FormatTimeField = "Format(" & FieldName & ",'hh:nn am/pm')"
End If
End Function

Public Function FormatDateField(ByVal FieldName As String) As String
If ((InStr(1, FieldName, " ") <> 0) And InStr(1, FieldName, "[") <> 0) Then FieldName = "" & FieldName & ""
If gUseMSDE Then
    'FormatDateField = "CONVERT(CHAR(8)," & FieldName & ",101)"
    FormatDateField = "RIGHT('000'+RTRIM(CONVERT(CHAR(2),DATEPART(Mm," & FieldName & "))),2)+'/' + RIGHT('000'+RTRIM(CONVERT(CHAR(2),DATEPART(Dd," & FieldName & "))),2) + '/'+RIGHT(RTRIM(CONVERT(CHAR(4),DATEPART(Yy," & FieldName & "))),2)"
Else
    FormatDateField = "Format(" & FieldName & ",'mm/dd/yy')"
End If
End Function

Public Function FormatMonthYearField(ByVal FieldName As String) As String
If ((InStr(1, FieldName, " ") <> 0) And InStr(1, FieldName, "[") <> 0) Then FieldName = "" & FieldName & ""
If gUseMSDE Then
    FormatMonthYearField = "RIGHT('000'+RTRIM(CONVERT(CHAR(2),DATEPART(Mm," & FieldName & "))),2)+'/01/'+RIGHT(RTRIM(CONVERT(CHAR(4),DATEPART(Yy," & FieldName & "))),2)"
Else
    FormatMonthYearField = "Format(" & FieldName & ",'mm/01/yy')"
End If
End Function

Public Function FormatNZ(ByVal FieldName As String, Optional ByVal pValue As String = "0") As String
If ((InStr(1, FieldName, " ") <> 0) And InStr(1, FieldName, "[") <> 0) Then FieldName = "" & FieldName & ""
If gUseMSDE Then
    FormatNZ = "ISNULL(" & FieldName & "," & pValue & ")"
Else
    FormatNZ = "IIF(ISNULL(" & FieldName & ")," & pValue & "," & FieldName & ")"
End If
End Function

Public Function FormatIIF(ByVal IIFCriteria As String, ByVal IfTrue As String, ByVal IfFalse As String) As String
If gUseMSDE Then
    FormatIIF = "(CASE WHEN " & IIFCriteria & " THEN " & IfTrue & " Else " & IfFalse & " END)"
Else
    FormatIIF = "IIF(" & IIFCriteria & "," & IfTrue & "," & IfFalse & ")"
End If
End Function

Public Function FormatRound(ByVal FieldName As String, ByVal NDecimal As Long) As String
Dim i As Long, strBuff As String
If ((InStr(1, FieldName, " ") <> 0) And InStr(1, FieldName, "[") <> 0) Then FieldName = "" & FieldName & ""
If NDecimal > 0 Then
    For i = 1 To NDecimal
        strBuff = strBuff & "0"
    Next
    strBuff = "." & strBuff
Else
    strBuff = ""
End If
If gUseMSDE Then
    FormatRound = "Round(" & FieldName & "," & NDecimal & ")"
Else
    FormatRound = "Val(Format(" & FieldName & ",'#,###" & strBuff & "'))"
End If
End Function

Public Function CriteriaDateBetween(ByVal FieldName As String, ByVal DateFrom As Date, ByVal DateTo As Date) As String
If ((InStr(1, FieldName, " ") <> 0) And InStr(1, FieldName, "[") <> 0) Then FieldName = "" & FieldName & ""
If gUseMSDE Then
    CriteriaDateBetween = "(" & FieldName & ">='" & Format(DateFrom, "mmm dd, yyyy " & " 0:00:00") & "' AND " & FieldName & "<='" & Format(DateTo, "mmm dd, yyyy " & " 23:59:59") & "')"
Else
    CriteriaDateBetween = "(" & FieldName & ">=#" & Format(DateFrom, "mmm dd, yyyy " & " 0:00:00") & "# AND " & FieldName & "<=#" & Format(DateTo, "mmm dd, yyyy " & " 23:59:59") & "#)"
End If
End Function

Public Function DataEncode(ByVal pData As String, ByVal pType As String) As String
On Error Resume Next
Dim strRetval As String
Select Case UCase(pType)
    Case "S"
        strRetval = "'" & StringReplace(pData, "'", "''") & "'"
    Case "C"
        strRetval = Format(pData, gAmountFormat1)
    Case "D"
        Dim dBlank As Date
        If gUseMSDE Then
            If pData = "" Then
                strRetval = "'" & Format(CDate(dBlank), gDateFormat) & "'"
            Else
                strRetval = "'" & Format(CDate(pData), gDateFormat) & "'"
            End If
        Else
            If pData = "" Then
                strRetval = "#" & Format(CDate(dBlank), gDateFormat) & "#"
            Else
                strRetval = "#" & Format(CDate(pData), gDateFormat) & "#"
            End If
        End If
    Case "T"
        If gUseMSDE Then
            If pData = "" Then
                strRetval = "'" & Format(CDate(dBlank), gTimeFormat) & "'"
            Else
                strRetval = "'" & Format(CDate(pData), gTimeFormat) & "'"
            End If
        Else
            If pData = "" Then
                strRetval = "#" & Format(CDate(dBlank), gTimeFormat) & "#"
            Else
                strRetval = "#" & Format(CDate(pData), gTimeFormat) & "#"
            End If
        End If
    Case Else
        strRetval = pData
End Select
DataEncode = strRetval
End Function

Public Function AddToTable(ByVal pTableName As String, ByVal pBizObj As Object) As Boolean
On Error Resume Next
AddToTable = ExecCommand(pTableName, pBizObj, "a" & pBizObj.GetString)
End Function

Public Function ExecCommand(ByVal pTableName As String, ByVal pBizObj As Object, ByVal pCommand As String) As Boolean
On Error GoTo ErrExit
Dim n As Long
Dim l As Long
Dim u As Long
Dim strCmd As String
DoEvents
Select Case Left(pCommand, 1)
    Case "a"
        strCmd = "INSERT INTO " & pTableName & "("
        strCmd = strCmd & GetFieldNames(pBizObj.GetFields) & ")"
        strCmd = strCmd & " VALUES ("
        strCmd = strCmd & GetFieldValues(Mid(pCommand, 2), pBizObj.GetFields, pBizObj.GetFieldTypes, "a") & ")"
    Case "e"
        strCmd = "UPDATE " & pTableName & " SET "
        strCmd = strCmd & GetFieldValues(Mid(pCommand, 2), pBizObj.GetFields, pBizObj.GetFieldTypes, "e")
        strCmd = strCmd & " WHERE " & GetKeyFieldCriteria(Mid(pCommand, 2), pBizObj.GetFields, pBizObj.GetFieldTypes, "e")
    Case "d"
        strCmd = "DELETE FROM " & pTableName
        strCmd = strCmd & " WHERE " & pBizObj.GetKeyFieldName & " = "
        strCmd = strCmd & DataEncode(Mid(pCommand, 3), Left(pBizObj.GetKeyFieldType, 1))
End Select
godb.RunSQL strCmd, True
'If moCmd Is Nothing Then Set moCmd = godb.CreateDBCommand
'moCmd.ActiveConnection = Connection
'moCmd.CommandText = strCmd
'moCmd.Execute
DoEvents
ExecCommand = True
Exit Function
ErrExit:
MsgBox "Error:" & Error, vbInformation, "Error"
ExecCommand = False
End Function

Public Function IsBuffered() As Boolean
IsBuffered = (gIsDatabaseBuffered = 1)
End Function

Public Function UseTextFile() As Boolean
UseTextFile = Not UseDatabaseStorage
End Function

Public Function IsEncrypted() As Boolean
IsEncrypted = True
End Function

Public Function CompactDatabase(Optional ByVal pPath As String = "") As Boolean
On Error GoTo ErrExit
Dim strSrcPath As String
Dim strDestPath As String
Dim strSrcConn As String
Dim strDestConn As String
Dim jro As New jro.JetEngine
Dim vOldPtr As Variant
Dim strDestin As String
If pPath = "" Then
    strSrcPath = Path_and_Filename
Else
    strSrcPath = pPath
End If
If Not FileExists(strSrcPath) Then
    MsgBox "File not found '" & strSrcPath & "'!", vbExclamation, "File Not Found"
    GoTo ErrExit
End If
strDestPath = GetPathPart(TempFilePath) & "~" & Trim(GetNamePart(strSrcPath))
If FileExists(strDestPath) Then
    If MsgBox("Temporary file '" & strDestPath & " already exists!" & Chr(13) & "Do you want to delete this file?", vbYesNo + vbQuestion + vbDefaultButton2, "Delete File") = vbNo Then
        GoTo ErrExit
    End If
    KillPath strDestPath
End If

vOldPtr = Screen.MousePointer
Screen.MousePointer = vbHourglass
DoEvents

strSrcConn = "Provider=Microsoft.Jet.OLEDB.4.0;"
strSrcConn = strSrcConn & "Data Source=" & strSrcPath & ";"
strSrcConn = strSrcConn & "User ID=Admin;"
strSrcConn = strSrcConn & "Jet OLEDB:Database Password=" & DatabasePassword & ";"

strDestConn = "Provider=Microsoft.Jet.OLEDB.4.0;"
strDestConn = strDestConn & "Data Source=" & strDestPath & ";"
strDestConn = strDestConn & "Jet OLEDB:Engine Type=5;"
strDestConn = strDestConn & "User ID=Admin;"
strDestConn = strDestConn & "Jet OLEDB:Database Password=" & DatabasePassword & ";"

jro.CompactDatabase strSrcConn, strDestConn
Set jro = Nothing

'KillPath strSrcPath
FileCopy strDestPath, strSrcPath
KillPath strDestPath
Screen.MousePointer = vOldPtr
CompactDatabase = True
Exit Function
ErrExit:
Screen.MousePointer = vOldPtr
CompactDatabase = False
If Err <> 0 Then
    MsgBox Error, vbExclamation, "Error"
End If
Exit Function
End Function

Public Function TurnOffMSDE() As Boolean
On Error GoTo ErrExit
Dim oSvr As Object 'SQLDMO.SQLServer
Set oSvr = CreateObject("SQLDMO.SQLServer")
On Error GoTo StartError
oSvr.LoginTimeout = mLoginTimeout ' This needs to be high to avoid time-out errors.
oSvr.Stop
ExitSub:
TurnOffMSDE = True
Exit Function
StartError:
'If Err.number = -2147023840 Then 'This error is thrown when the server is already running, and Server.Start is executed on NT.
'    oSvr.connect mServer, mUser, mPassword
'    Resume Next
'End If
TurnOffMSDE = True
Exit Function
ErrExit:
MsgBox Error, vbInformation, "Error# " & Err
End Function

Public Function TurnOnMSDE() As Boolean
TryStart:
On Error GoTo ErrExit
Dim oSvr As Object 'SQLDMO.SQLServer
Set oSvr = CreateObject("SQLDMO.SQLServer")
On Error GoTo StartError
oSvr.LoginTimeout = mLoginTimeout ' This needs to be high to avoid time-out errors.
oSvr.start True, mServer, mUser, mPassword
ExitSub:
TurnOnMSDE = True
Exit Function
StartError:
If Err.number = -2147023840 Then 'This error is thrown when the server is already running, and Server.Start is executed on NT.
    oSvr.connect mServer, mUser, mPassword
    Resume Next
ElseIf Err.number = 18456 Then
    On Error Resume Next
    Dim oSvr2 As Object
    Set oSvr2 = CreateObject("SQLDMO.SQLServer")
    oSvr2.connect mServer, mUser, mOldPassword
    If oSvr2 Is Nothing Then
        'unknown error, password probably different
        GoTo ErrExit
    Else
        oSvr2.ExecuteImmediate "sp_password null, '" & mPassword & "', 'sa'"
        If Err <> 0 Then GoTo ErrExit
        oSvr2.Stop
        'oSvr.DisConnect
        Set oSvr2 = Nothing
        'Change Password
        GoTo TryStart
    End If
ElseIf Err.number = 1056 Then
    TurnOnMSDE = True
End If
Exit Function
ErrExit:
MsgBox Error, vbInformation, "Error# " & Err
End Function

Public Function ConnectData() As Boolean
TryStart:
On Error GoTo ErrExit
Dim oSvr As Object 'SQLDMO.SQLServer
Dim strMessage As String
Set oSvr = CreateObject("SQLDMO.SQLServer")
oSvr.connect mServer, mUser, mPassword
strMessage = oSvr.AttachDBWithSingleFile(mDatabaseName, mSQLServerPath & mSQLServerDataPath & mDatapath)
oSvr.DisConnect
Set oSvr = Nothing
ConnectData = True
Exit Function
ErrExit:
If Err.number = 18456 Then
    On Error Resume Next
    Dim oSvr2 As Object
    Set oSvr2 = CreateObject("SQLDMO.SQLServer")
    oSvr2.connect mServer, mUser, mOldPassword
    If oSvr2 Is Nothing Then
        'unknown error, password probably different
        GoTo ErrExit2
    Else
        oSvr2.ExecuteImmediate "sp_password null, '" & mPassword & "', 'sa'"
        If Err <> 0 Then GoTo ErrExit2
        oSvr2.DisConnect
        Set oSvr2 = Nothing
        GoTo TryStart
    End If
Else
ErrExit2:
    MsgBox Error, vbInformation, "Error# " & Err
End If
End Function

Public Function CheckUseMSDE()
On Error Resume Next
Dim oShell As New clsShell, strSourceDir As String
If Dir(mSQLServerPath1 & "binn\sqlservr.exe", vbHidden) <> "" Then
    mSQLServerPath = mSQLServerPath1
ElseIf Dir(mSQLServerPath2 & "binn\sqlservr.exe", vbHidden) <> "" Then
    mSQLServerPath = mSQLServerPath2
Else
    GoTo ErrExit
End If
If Dir(mSQLServerPath & mSQLServerDataPath & mDatapath, vbHidden) = "" Then
    strSourceDir = App.Path
    If Right(strSourceDir, 1) <> "\" Then strSourceDir = strSourceDir & "\"
    If Dir(strSourceDir & mDatapath, vbHidden) <> "" Then
         FileCopy strSourceDir & mDatapath, mSQLServerPath & "Data\" & mDatapath
    End If
    If Dir(strSourceDir & mDatapath2, vbHidden) <> "" Then
         FileCopy strSourceDir & mDatapath2, mSQLServerPath & "Data\" & mDatapath2
    End If
End If
If Dir(mSQLServerPath & mSQLServerDataPath & mDatapath, vbHidden) <> "" Then
    Me.TurnOnMSDE
    If (oShell.GetExeProcessID("Sqlservr.exe") <> 0) Then
        Me.ConnectData
        gUseMSDE = True
        'If Me.ConnectData Then
        '    gUseMSDE = True
        'Else
        '    gUseMSDE = False
        'End If
    Else
        gUseMSDE = False
    End If
Else
ErrExit:
    gUseMSDE = False
End If
Me.GetConnectionString
End Function

Public Function RecoverDatabase() As Boolean
On Error GoTo ErrExit
Screen.MousePointer = vbHourglass
Dim dNow As Date
Dim dDate As Date
Dim strFile As String
Dim strFName As String
Dim strSource As String
Dim sDataFile As String
Dim sDmgdFile As String
Dim n As Long
sDataFile = Path_and_Filename
strSource = gBackupDirectory
If strSource = "" Then strSource = gDefBackupDirectory
'If strSource = "" Then GoTo ErrExit
If Right(gBackupDirectory, 1) <> "\" Then
    strSource = strSource & "\"
End If
Call CreateNewDir(strSource)
dNow = Now
For n = 0 To 6 'try for 7 day backups
    dDate = DateAdd("d", n * -1, dNow)
    strFName = GetWeekDay(dDate) & "_" & gDatabaseName
    strFile = strSource & strFName
    If FileExists(strFile) Then
        If Format(FileDateTime(strFile), gDateFormat) = Format(dDate, gDateFormat) Then
            If MsgBox("Do you want to recover from the latest backup from " & IIf(n = 0, "earlier today", n & " day(s) ago") & "?", vbExclamation + vbYesNo + vbDefaultButton1, "DATABASE HAS BEEN DAMAGED !!!") = vbYes Then
                sDmgdFile = GetSaveFile(strSource, "Save damaged database as filename ?", , "MDB Files (*.MDB)", "*.MDB")
                If sDmgdFile = "" Then
                    'cancelled by user
                    GoTo ErrExit
                ElseIf FileExists(sDmgdFile) Then
                    MsgBox "File exists '" & sDmgdFile & "'!", vbExclamation, "Cannot Overwrite"
                    GoTo ErrExit
                Else
                    FileCopy sDataFile, sDmgdFile
                    If FileExists(sDmgdFile) Then
                        Kill sDataFile
                    End If
                    FileCopy strFile, sDataFile
                    If FileExists(sDataFile) Then
                        MsgBox "Database successfully recovered", vbInformation, "Continue..."
                    Else
                        MsgBox "Database recovery failed", vbExclamation, "Error"
                    End If
                    GoTo OKExit
                End If
            Else
                'cancelled by user
                GoTo ErrExit
            End If
            Exit For
        End If
    End If
Next
MsgBox "No backup database found!", vbExclamation, "Cannot Recover Database"
Screen.MousePointer = vbDefault
RecoverDatabase = False
Exit Function
ErrExit:
If Err <> 0 Then MsgBox Error, vbCritical, "Recovery Error"
Screen.MousePointer = vbDefault
RecoverDatabase = False
Exit Function
OKExit:
Screen.MousePointer = vbDefault
RecoverDatabase = True
Exit Function
End Function

Public Function BackupDatabase() As Boolean
On Error GoTo ErrExit
Screen.MousePointer = vbHourglass
Dim strFile As String
Dim strFName As String
Dim strDestin As String
strDestin = gBackupDirectory
If strDestin = "" Then strDestin = gDefBackupDirectory
'If strDestin = "" Then GoTo ErrExit
If Right(gBackupDirectory, 1) <> "\" Then
    strDestin = strDestin & "\"
End If
Call CreateNewDir(strDestin)
If gUseMSDE Then
    TurnOffMSDE
    If Dir(mSQLServerPath1 & "binn\sqlservr.exe", vbHidden) <> "" Then
        mSQLServerPath = mSQLServerPath1
    ElseIf Dir(mSQLServerPath2 & "binn\sqlservr.exe", vbHidden) <> "" Then
        mSQLServerPath = mSQLServerPath2
    Else
        GoTo ErrExit
    End If
    strFile = mSQLServerPath & mSQLServerDataPath & mDatapath
    If Not FileExists(strFile) Then
        GoTo ErrExit
    Else
        strFName = GetWeekDay(Now) & "_" & mDatapath
        FileCopy strFile, strDestin & strFName
    End If
    strFile = mSQLServerPath & mSQLServerDataPath & mDatapath2
    If Not FileExists(strFile) Then
        GoTo ErrExit
    Else
        strFName = GetWeekDay(Now) & "_" & mDatapath2
        FileCopy strFile, strDestin & strFName
    End If
Else
    If Not FileExists(Path_and_Filename) Then
        GoTo ErrExit
    Else
        strFName = GetWeekDay(Now) & "_" & gDatabaseName
        FileCopy Path_and_Filename, strDestin & strFName
    End If
End If
Screen.MousePointer = vbDefault
BackupDatabase = True
Exit Function
ErrExit:
Screen.MousePointer = vbDefault
BackupDatabase = False
End Function


