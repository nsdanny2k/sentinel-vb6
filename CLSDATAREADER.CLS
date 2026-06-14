VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDataReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Filename As String
Public EndOfFile As Boolean
Public DoNotCreate As Boolean

Public IsHidden As Boolean
Public IsEncrypted As Boolean
Public IsRandomAccess As Boolean

Private moEncrypt As New clsEncryption
Public mFnum

Private moObj As Object
Private moLen As Long
Private mCurrRec As Long

Public EncryptionMethod As Long

Public Function OpenFile(Optional ByRef pObj As Object = Nothing) As Boolean
Dim pRetval As Boolean
On Error GoTo ErrExit
pRetval = False
If Filename <> "" Then
    If FileExists(Filename) Then 'open
        mFnum = FreeFile(1)
        If IsRandomAccess Then
            If pObj Is Nothing Then
                MsgBox "Invalid object", vbInformation, "Cannot Read File"
                GoTo ErrExit
            Else
                Set moObj = pObj
                moLen = moObj.TLen
                Open Filename For Random As #mFnum Len = moLen
            End If
        Else
            Open Filename For Input As #mFnum
        End If
        pRetval = True
    Else
        If DoNotCreate Then 'cannot create
            GoTo ErrExit
        Else 'create and open
            mFnum = FreeFile(1)
            If IsRandomAccess Then
                If pObj Is Nothing Then
                    MsgBox "Invalid object", vbInformation, "Cannot Read File"
                    GoTo ErrExit
                Else
                    Set moObj = pObj
                    moLen = moObj.TLen
                    Open Filename For Random As #mFnum Len = moLen
                End If
            Else
                Open Filename For Output As #mFnum
                Close #mFnum
                mFnum = FreeFile(1)
                Open Filename For Input As #mFnum
            End If
            pRetval = True
        End If
    End If
End If
OpenFile = pRetval
Exit Function
ErrExit:
OpenFile = False
If Err <> 0 And Err <> 76 Then
    frmMessage.Display Error, vbInformation, "Error"
End If
On Error Resume Next
Close #mFnum
Exit Function
End Function

Public Sub CloseFile()
On Error Resume Next
Close #mFnum
End Sub

Public Function ReadLine(ByRef pData As String, Optional ByRef pObj As Object = Nothing) As Boolean
On Error GoTo ErrExit
Dim nRec As Long
Dim i As Long
Dim strBuff As String
pData = ""
If IsRandomAccess Then
    If pObj Is Nothing Then
        If moObj Is Nothing Then
            MsgBox "Invalid object", vbInformation, "Cannot Read File"
            GoTo ErrExit
        End If
    Else
        'ignore pobj paramater
    End If
    nRec = LOF(mFnum) / moLen
    If nRec > 0 Then
        If mCurrRec = nRec Then
            EndOfFile = True
            ReadLine = False
        Else
            mCurrRec = mCurrRec + 1
            If Not moObj Is Nothing Then
                ReadLine = moObj.ReadObject(mFnum, strBuff, IsEncrypted)
            Else
                ReadLine = True
            End If
            pData = strBuff
            EndOfFile = (nRec = mCurrRec)
        End If
    Else
        EndOfFile = True
        ReadLine = False
    End If
Else
    If (EOF(mFnum) = False) Then
        EndOfFile = False
        Input #mFnum, strBuff
        If IsEncrypted Then
            moEncrypt.EncryptionMethod = EncryptionMethod
            strBuff = moEncrypt.DecryptText(strBuff)
        End If
        If Not pObj Is Nothing Then
            ReadLine = pObj.ReadLine(strBuff)
        Else
            ReadLine = True
        End If
        pData = strBuff
    Else
        EndOfFile = True
        ReadLine = False
    End If
End If
Exit Function

ErrExit:
pData = ""
ReadLine = False
Exit Function
End Function

Public Function GetString(ByRef pData As String, ByVal pObj As Object) As Boolean
On Error GoTo ErrExit
pData = pObj.GetString
GetString = True
Exit Function

ErrExit:
pData = ""
GetString = False
Exit Function
End Function

Private Sub Class_Terminate()
On Error Resume Next
CloseFile
End Sub

Public Function GetFNum() As Long
GetFNum = mFnum
End Function

Private Sub Class_Initialize()
EncryptionMethod = 1
End Sub
