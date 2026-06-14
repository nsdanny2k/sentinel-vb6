VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDataWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const mTempPrefix As String = "~"

Public Filename As String

Public AppendOnly As Boolean

Public IsEncrypted As Boolean
Public IsHidden As Boolean
Public IsRandomAccess As Boolean
Public IsBuffered As Boolean

Private mFnum
Private moEncrypt As New clsEncryption

Private moObj As Object
Private moLen As Long
Private mCurrPos As Long

Public EncryptionMethod As Long

Public Function OpenFile(Optional ByRef pObj As Object) As Boolean
Dim pRetval As Boolean
Dim strTmp As String
On Error GoTo ErrExit
pRetval = False
mFnum = FreeFile(1)
CreateNewDir GetPathPart(Filename)
If IsHidden And FileExists(Filename) Then
    SetFileAttrNormal Filename
End If
If IsRandomAccess Then
    If IsBuffered Then
        strTmp = GetPathPart(Filename) & mTempPrefix & GetNamePart(Filename)
        FileCopy Filename, strTmp
        KillPath Filename
    End If
    Set moObj = pObj
    moLen = moObj.TLen
    Open Filename For Random As #mFnum Len = moLen
    If Not IsBuffered Then
        Dim nRec As Long
        nRec = LOF(mFnum) / moLen
        If nRec = 0 Then
            mCurrPos = 0
        Else
            mCurrPos = 1
        End If
    End If
Else
    If AppendOnly Then
        Open Filename For Append As #mFnum
    Else
        Open Filename For Output As #mFnum
    End If
End If
pRetval = True
OpenFile = pRetval
Exit Function
ErrExit:
OpenFile = False
frmMessage.Display Error, vbInformation, "Error"
On Error Resume Next
Close #mFnum
Exit Function
End Function

Public Sub CloseFile()
On Error Resume Next
Dim strTmp As String
Close #mFnum
If IsHidden Then
    SetFileAttrHidden Filename
End If
If IsBuffered Then
    strTmp = GetPathPart(Filename) & mTempPrefix & GetNamePart(Filename)
    KillPath strTmp
End If
End Sub

Public Function WriteLine(ByVal pData As String) As Boolean
Dim strBuff As String
On Error GoTo ErrExit
strBuff = pData
If IsRandomAccess Then
    If moObj Is Nothing Then
        MsgBox "Invalid Object", vbInformation, "Cannot Save"
        GoTo ErrExit
    End If
    moObj.WriteObject mFnum, pData, IsEncrypted
Else
    If IsEncrypted Then
        moEncrypt.EncryptionMethod = EncryptionMethod
        strBuff = moEncrypt.EncryptText(strBuff)
    End If
    Print #mFnum, strBuff
End If
WriteLine = True
Exit Function

ErrExit:
WriteLine = False
Exit Function
End Function

Public Function SeekRecord(ByVal pData As String) As Boolean
'look for key value = pdata
On Error GoTo ErrExit
Dim i As Long
Dim nRec As Long
Dim strTmp As String
Dim fLoc As Long
If IsRandomAccess Then
    If IsBuffered Then
        MsgBox "Database must be opened in buffered mode", vbInformation, "Cannot Seek"
        GoTo ErrExit
    End If
    If moObj Is Nothing Then
        MsgBox "Invalid Object", vbInformation, "Cannot Seek"
        GoTo ErrExit
    End If
    nRec = LOF(mFnum) / moLen
    If nRec > 0 Then
        
        pData = LCase(Trim(pData))
        For i = 1 To nRec
            
            moObj.ReadObject mFnum, strTmp, IsEncrypted
            
            fLoc = Loc(mFnum)
            
            If LCase(Trim(moObj.GetKey)) = pData Then
                Seek #mFnum, fLoc
                SeekRecord = True
                GoTo Done
            End If
        Next
Done:
        mCurrPos = fLoc
    Else
        mCurrPos = 0
    End If
Else
    SeekRecord = False
    mCurrPos = 0
End If
Exit Function
ErrExit:
SeekRecord = False
End Function

Public Function MoveNew() As Boolean
On Error GoTo ErrExit
Dim nRec As Long
Dim strTmp As String
nRec = LOF(mFnum) / moLen
If nRec = 0 Or IsBuffered Then
    mCurrPos = 0
    MoveNew = True
Else
    Seek #mFnum, nRec + 1
    mCurrPos = nRec + 1
    MoveNew = True
End If
Exit Function
ErrExit:
MoveNew = False
End Function

Public Function MoveFirst() As Boolean
On Error GoTo ErrExit
Dim nRec As Long
Dim strTmp As String
nRec = LOF(mFnum) / moLen
If nRec = 0 Or IsBuffered Then
    mCurrPos = 0
    MoveFirst = True
Else
    Seek #mFnum, 1
    moObj.ReadObject mFnum, strTmp, IsEncrypted
    Seek #mFnum, 1
    
    mCurrPos = 1
    MoveFirst = True
End If
Exit Function
ErrExit:
MoveFirst = False
End Function

Public Function MoveLast() As Boolean
On Error GoTo ErrExit
Dim nRec As Long
Dim strTmp As String
nRec = LOF(mFnum) / moLen
If nRec = 0 Or IsBuffered Then
    mCurrPos = 0
    MoveLast = True
Else
    Seek #mFnum, nRec
    moObj.ReadObject mFnum, strTmp, IsEncrypted
    Seek #mFnum, nRec
    
    mCurrPos = nRec
    MoveLast = True
End If
Exit Function
ErrExit:
MoveLast = False
End Function

Public Function MoveNext() As Boolean
On Error GoTo ErrExit
Dim nRec As Long
Dim strTmp As String
nRec = LOF(mFnum) / moLen
If nRec = 0 Or IsBuffered Then
    mCurrPos = 0
    MoveNext = False
ElseIf mCurrPos = nRec Then
    MoveNext = False
Else
    Seek #mFnum, mCurrPos + 1
    moObj.ReadObject mFnum, strTmp, IsEncrypted
    Seek #mFnum, mCurrPos + 1
    
    mCurrPos = mCurrPos + 1
    MoveNext = True
End If
Exit Function
ErrExit:
MoveNext = False
End Function

Public Function MovePrevious() As Boolean
On Error GoTo ErrExit
Dim nRec As Long
Dim strTmp As String
nRec = LOF(mFnum) / moLen
If nRec = 0 Or IsBuffered Then
    mCurrPos = 0
    MovePrevious = False
ElseIf mCurrPos = 1 Then
    MovePrevious = False
Else
    Seek #mFnum, mCurrPos - 1
    moObj.ReadObject mFnum, strTmp, IsEncrypted
    Seek #mFnum, mCurrPos - 1
    
    mCurrPos = mCurrPos - 1
    MovePrevious = True
End If
Exit Function
ErrExit:
MovePrevious = False
End Function

Public Function GetObject() As Object
On Error Resume Next
Set GetObject = moObj
End Function

Public Function GetFNum() As Long
GetFNum = mFnum
End Function

Private Sub Class_Initialize()
EncryptionMethod = 1
End Sub

