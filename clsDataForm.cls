VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDataForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event OnWriteStart(ByVal oWriter As Object)
Public Event OnWriteEnd(ByVal oWriter As Object, ByVal nRec As Long)

Public TempDataObject As Object
Public HasDefaultField As Boolean

Public IsHidden As Boolean
Public IsEncrypted As Boolean
Public IsRandomAccess As Boolean
Public IsBuffered As Boolean

Public DataList As ListBox
Public DataForm As Form

Private mDirty As Boolean
Private mCol As New Collection
Public DataPath As String

Private Sub ClearStore()
Dim strTmp
Dim n As Integer
On Error Resume Next
n = mCol.Count
If n > 0 Then
    For n = mCol.Count To 1 Step -1
        mCol.Remove n
    Next
End If
End Sub

Public Function FillList(ByVal pObj As Object, Optional ByVal pSilentMode As Boolean = False) As Boolean
Dim oDataReader As New clsDataReader
On Error GoTo ErrExit
Dim vOldPtr
vOldPtr = Screen.MousePointer
Screen.MousePointer = vbHourglass
oDataReader.Filename = DataPath
oDataReader.IsEncrypted = IsEncrypted
oDataReader.IsRandomAccess = IsRandomAccess
ClearStore
If oDataReader.OpenFile(pObj) Then
    Dim strBuff As String
    Dim strBuff2 As String
    While Not oDataReader.EndOfFile
        If oDataReader.ReadLine(strBuff, pObj) Then
            strBuff = DataForm.GetDisplayString(pObj)
            DataList.AddItem strBuff
            strBuff2 = DataForm.GetString(pObj)
            mCol.Add strBuff2
        End If
    Wend
    oDataReader.CloseFile
Else
    GoTo ErrExit
End If
Screen.MousePointer = vOldPtr
mDirty = False
FillList = True
Exit Function

ErrExit:
FillList = False
On Error Resume Next
Screen.MousePointer = vOldPtr
If Not pSilentMode Then
    If Err <> 0 And Err <> 76 Then
        frmMessage.Display Error, vbInformation, "Unable to fill the list"
    End If
End If
Exit Function
End Function

Public Function Update(ByVal pObj As Object, Optional ByVal pNoWarning As Boolean = False) As Boolean
Dim oDataWriter As New clsDataWriter
oDataWriter.IsHidden = IsHidden
On Error GoTo ErrExit
oDataWriter.Filename = DataPath
oDataWriter.IsRandomAccess = IsRandomAccess
oDataWriter.IsEncrypted = IsEncrypted
oDataWriter.IsBuffered = IsBuffered
If oDataWriter.OpenFile(pObj) Then
    RaiseEvent OnWriteStart(oDataWriter)
    If mCol.Count > 0 Then
        Dim strBuff As String
        Dim n As Integer
        For n = 1 To mCol.Count
            strBuff = mCol.item(n)
            Set pObj = DataForm.GetObject(strBuff)
            strBuff = pObj.GetString()
            oDataWriter.WriteLine strBuff
        Next
    End If
    RaiseEvent OnWriteEnd(oDataWriter, mCol.Count)
    oDataWriter.CloseFile
    mDirty = False
Else
    GoTo ErrExit
End If
Update = True
If Not pNoWarning Then
    MsgBox "Update complete.", vbInformation, "Success"
End If
Exit Function

ErrExit:
Update = False
If Not pNoWarning Then
    If Err <> 0 Then
        MsgBox Error, vbInformation, "Unable to Update"
    End If
End If
Exit Function
End Function

Public Function Delete() As Boolean
Dim n As Long
On Error Resume Next
If DataList.ListIndex >= 0 Then
    n = DataList.ListIndex
    DataList.RemoveItem n
    mCol.Remove n + 1
    mDirty = True
Else
    Beep
End If
End Function

Public Function GetItem() As Object
Dim n As Long
Dim pRetval As Object
Dim strBuff As String
On Error Resume Next
If DataList.ListIndex >= 0 Then
    n = DataList.ListIndex
    strBuff = mCol.item(n + 1)
    Set pRetval = DataForm.GetObject(strBuff)
Else
    Set pRetval = Nothing
End If
Set GetItem = pRetval
End Function

Public Function Edit(ByVal pObj As Object) As Boolean
Dim strBuff As String
Dim strBuff2 As String
Dim n As Long
On Error Resume Next
strBuff = DataForm.GetDisplayString(pObj)
strBuff2 = DataForm.GetString(pObj)
If DataList.ListIndex >= 0 Then
    n = DataList.ListIndex
    If HasDefaultField Then
        Set TempDataObject = DataForm.GetObject(strBuff2)
        If Not TempDataObject Is Nothing Then
            If TempDataObject.Field("default") = "1" Then
                ResetDefault TempDataObject
            End If
        End If
    End If
    DataList.List(n) = strBuff
    mCol.Add strBuff2, , , n + 1
    mCol.Remove n + 1
    DataList.Text = strBuff
    mDirty = True
    Edit = True
Else
    Edit = False
End If
End Function

Public Function AddNew(ByVal pObj As Object) As Boolean
On Error Resume Next
If FindKey(pObj.GetKey) Then
    MsgBox "Record is already in the list!", vbInformation, "Invalid Data"
    AddNew = False
Else
    Dim strBuff As String
    Dim strBuff2 As String
    Dim n As Long
    n = DataList.ListIndex
    strBuff = DataForm.GetDisplayString(pObj)
    strBuff2 = DataForm.GetString(pObj)
    If HasDefaultField Then
        Set TempDataObject = DataForm.GetObject(strBuff2)
        If Not TempDataObject Is Nothing Then
            If TempDataObject.Field("default") = "1" Then
                ResetDefault TempDataObject
            End If
        End If
    End If
    If n >= 0 Then
        DataList.AddItem strBuff, n + 1
        mCol.Add strBuff2, , , n + 1
    Else
        DataList.AddItem strBuff
        mCol.Add strBuff2
    End If
    DataList.Text = strBuff
    mDirty = True
    AddNew = True
End If
End Function

Private Function ResetDefault(ByVal pObj As Object) As Boolean
Dim n As Long
Dim strBuff As String
Dim strBuff2 As String
On Error Resume Next
If mCol.Count > 0 Then
    For n = 1 To mCol.Count
        strBuff = mCol.item(n)
        Set pObj = DataForm.GetObject(strBuff)
        If Not pObj Is Nothing Then
            pObj.SetDefault False
            strBuff = DataForm.GetDisplayString(pObj)
            strBuff2 = DataForm.GetString(pObj)
            DataList.RemoveItem n - 1
            DataList.AddItem strBuff, n - 1
            mCol.Add strBuff2, , n
            mCol.Remove n + 1
        End If
    Next
    mDirty = True
End If
ResetDefault = True
End Function

Private Function FindKey(ByVal pData As String)
Dim n As Integer
Dim strBuff As String
Dim oObj As Object
On Error Resume Next
If mCol.Count > 0 Then
    For n = 1 To mCol.Count
        strBuff = mCol.item(n)
        Set oObj = DataForm.GetObject(strBuff)
        If Trim(LCase(oObj.GetKey)) = Trim(LCase(pData)) Then
            FindKey = True
            Exit Function
        End If
    Next
End If
FindKey = False
End Function

Public Property Get IsDirty() As Boolean
IsDirty = mDirty
End Property

Public Function AskToSave() As Boolean
On Error Resume Next
If mDirty Then
    If MsgBox("Do you want to save your changes?", vbQuestion + vbYesNo + vbDefaultButton1, "Update?") = vbYes Then
        AskToSave = True
    Else
        AskToSave = False
    End If
Else
    AskToSave = False
End If
End Function

Public Sub Destroy()
On Error Resume Next
Set DataList = Nothing
Set DataForm = Nothing
End Sub

Private Sub Class_Initialize()
mDirty = False
IsEncrypted = True
End Sub

'Public Sub DebugIt()
'Dim n As Integer
'DataForm.List2.Clear
'If mCol.Count > 0 Then
'    For n = 1 To mCol.Count
'        DataForm.List2.AddItem mCol.Item(n)
'    Next
'End If
'End Sub
