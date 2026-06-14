VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDataForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Public EditState As Boolean

Public Event CloseForm()
Public Event BeforeSave(ByRef Cancel As Integer)
Public Event AfterSave()
Public Event AddColHeader(ByVal plvi As ListItem, ByRef pSuccess As Boolean)
Public Event AddToList(ByVal plvi As ListItem, ByRef pObject As Object, ByRef pSuccess As Boolean)
Public Event SetBuffered(ByRef pBuffered As Boolean)

Public Event AllowDeleteItem(ByVal pKey As String, ByRef Cancel As Integer)
Public Event DeleteItem(ByVal pKey As String, ByRef Cancel As Integer)
Public Event AfterDeleteItem(ByVal pKey As String)
Public Event AddItem(ByVal pBizObj As Object, ByRef Cancel As Integer)
Public Event AfterAddItem(ByVal pBizObj As Object)
Public Event EditItem(ByVal pBizObj As Object, ByRef Cancel As Integer)
Public Event AfterEditItem(ByVal pBizObj As Object)

Public AllowClose As Boolean
Public RandomAccess As Boolean
Public BizObj As Object
Public TableName As String
Public EditForm As Object
Public Filter As String
Public NoDoubleClickList As Boolean

Private WithEvents pForm As Form
Attribute pForm.VB_VarHelpID = -1
Private WithEvents pList As ListView
Attribute pList.VB_VarHelpID = -1
Private WithEvents pAdd As CommandButton
Attribute pAdd.VB_VarHelpID = -1
Private WithEvents pEdit As CommandButton
Attribute pEdit.VB_VarHelpID = -1
Private WithEvents pDelete As CommandButton
Attribute pDelete.VB_VarHelpID = -1
Private WithEvents pSave As CommandButton
Attribute pSave.VB_VarHelpID = -1
Private WithEvents pClose As CommandButton
Attribute pClose.VB_VarHelpID = -1

Private lvi As ListItem

Private moRS As Object
Private mcolCommands As Collection
Private mDataReader As clsDataReader
Private mDataWriter As clsDataWriter

Private mUseTextFile As Boolean
Private mBuffered As Boolean
Private mRandomAccess As Boolean
Private mDirty As Boolean
Private SilentMode As Boolean

Private mSortCol As Long

Private Function ReadData() As Boolean
On Error GoTo ErrExit
Set mcolCommands = New Collection
mBuffered = goDBConn.IsBuffered
RaiseEvent SetBuffered(mBuffered)
mUseTextFile = goDBConn.UseTextFile
mRandomAccess = RandomAccess
If mUseTextFile Then mBuffered = True
pSave.Visible = mBuffered
If mUseTextFile Then
    ReadData = True
Else
    ReadData = (goDBConn.Connection.State = 1) 'adStateOpen)
End If
Exit Function
ErrExit:
ReadData = False
End Function

Private Sub CloseData()
On Error Resume Next
End Sub

Public Function RefreshList()
pForm.ListView1.ListItems.Clear
FillList
End Function

Private Function FillList() As Boolean
On Error GoTo ErrExit
Dim vOldPtr
Dim strBuff As String
Dim strBuff2 As String
vOldPtr = Screen.MousePointer
Screen.MousePointer = vbHourglass
Dim strTmp As String
If Not mUseTextFile Then
    'Set moRS = godb.GetRS("SELECT * FROM " & TableName & IIf(Filter = "", "", IIf(gUseDatabase = 1, " WHERE " & Filter, "")), True)
    Set moRS = godb.GetRS("SELECT * FROM " & TableName & IIf(Filter = "", "", " WHERE " & Filter), True)
    pForm.ListView1.ListItems.Clear
    'If moRS.RecordCount > 0 Then
    If Not moRS.EOF Then
         Do While Not moRS.EOF
            strBuff = goDBConn.GetString(moRS, BizObj.GetFields)
            If BizObj.ReadLine(strBuff) Then
                Set lvi = pForm.ListView1.ListItems.Add(, "K" & BizObj.GetKey)
                lvi.Text = BizObj.GetKey
                Dim bSuccess As Boolean
                RaiseEvent AddToList(lvi, BizObj, bSuccess)
                lvi.Tag = strBuff
            End If
            moRS.MoveNext
         Loop
    End If
    moRS.Close
    Set moRS = Nothing
Else
    Set mDataReader = New clsDataReader
    mDataReader.FileName = LogFilePath(TableName, True)
    mDataReader.IsEncrypted = goDBConn.IsEncrypted
    mDataReader.IsHidden = False
    mDataReader.IsRandomAccess = mRandomAccess
    If mDataReader.OpenFile(BizObj) Then
        While Not mDataReader.EndOfFile
            If mDataReader.ReadLine(strBuff, BizObj) Then
                Set lvi = pForm.ListView1.ListItems.Add(, "K" & BizObj.GetKey)
                lvi.Text = BizObj.GetKey
                RaiseEvent AddToList(lvi, BizObj, bSuccess)
                lvi.Tag = strBuff
            End If
        Wend
        mDataReader.CloseFile
    Else
        GoTo ErrExit
    End If
    Set mDataReader = Nothing
End If
Screen.MousePointer = vOldPtr
mDirty = False
FillList = True
Exit Function
ErrExit:
Screen.MousePointer = vOldPtr
If Not SilentMode Then
    If Err <> 0 Then
        MsgBox Error, vbInformation, "Unable to fill the list"
    End If
End If
On Error Resume Next
If Not mDataReader Is Nothing Then
    mDataReader.CloseFile
End If
Set mDataReader = Nothing
Exit Function
End Function

Public Function SaveRecords() As Boolean
On Error GoTo ErrExit
Dim n As Long
Dim i As Long
Dim strTmp As String
Dim nCancel As Integer
Dim bError As Boolean
Dim bRetval As Boolean
Dim bKey As Boolean
Dim j As Long
Dim nAdj As Long
Dim nCancel2 As Integer
Dim strKey As String
If mDirty Then
    nCancel = 0
    RaiseEvent BeforeSave(nCancel)
    If nCancel = 0 Then
        bRetval = True
        If Not mUseTextFile Then
            i = mcolCommands.Count
            nAdj = 0
            If i > 0 Then
                For n = 1 To i
                    j = n - nAdj
                    strTmp = mcolCommands.Item(j)
                    If strTmp <> "" Then
                        nCancel2 = 0
                        If Left(strTmp, 1) = "d" Then
                            strKey = Trim(Mid(mcolCommands.Item(j), 3))
                            RaiseEvent DeleteItem(strKey, nCancel2)
                        End If
                        If Left(strTmp, 1) = "a" Then
                            BizObj.ReadLine Mid(mcolCommands.Item(j), 2)
                            RaiseEvent AddItem(BizObj, nCancel2)
                        End If
                        If Left(strTmp, 1) = "e" Then
                            BizObj.ReadLine Mid(mcolCommands.Item(j), 2)
                            RaiseEvent EditItem(BizObj, nCancel2)
                        End If
                        If nCancel2 <> -1 Then
                            bError = Not goDBConn.ExecCommand(TableName, BizObj, strTmp)
                        Else
                            bError = False
                        End If
                        If Not bError Then
                            '// erroneous bug? bizobj
                            If Left(strTmp, 1) = "d" Then
                                RaiseEvent AfterDeleteItem(strKey)
                            End If
                            If Left(strTmp, 1) = "a" Then
                                RaiseEvent AfterAddItem(BizObj)
                            End If
                            If Left(strTmp, 1) = "e" Then
                                RaiseEvent AfterEditItem(BizObj)
                            End If
                        End If
                        If bError Then
                            bRetval = False
                        Else
                            mcolCommands.Remove j
                            nAdj = nAdj + 1
                        End If
                    End If
                Next
            End If
        Else
            Set mDataWriter = New clsDataWriter
            mDataWriter.FileName = LogFilePath(TableName, True)
            mDataWriter.IsEncrypted = goDBConn.IsEncrypted
            mDataWriter.IsHidden = False
            mDataWriter.IsRandomAccess = mRandomAccess
            mDataWriter.IsBuffered = mBuffered
            If mDataWriter.OpenFile(BizObj) Then
                For Each lvi In pForm.ListView1.ListItems
                    strTmp = lvi.Tag
                    mDataWriter.WriteLine strTmp
                Next
                mDataWriter.CloseFile
                mDirty = False
            Else
                GoTo ErrExit
            End If
            Set mDataWriter = Nothing
        End If
        RaiseEvent AfterSave
        MsgBox "Records Saved", vbInformation, "Success"
    End If
End If
mDirty = Not bRetval
SaveRecords = bRetval
Exit Function
ErrExit:
mDirty = True
SaveRecords = False
If Not SilentMode Then
    If Err <> 0 Then
        MsgBox Error, vbInformation, "Save Error"
    End If
End If
On Error Resume Next
If Not mDataWriter Is Nothing Then
    mDataWriter.CloseFile
End If
End Function

Private Function KeyFound() As Boolean
On Error GoTo ErrExit
Dim strKey As String
Dim bRetval As Boolean
strKey = BizObj.GetKey
For Each lvi In pList.ListItems
    If LCase("K" & strKey) = LCase(lvi.Key) Then
        bRetval = True
        Exit For
    End If
Next
KeyFound = bRetval
Exit Function
ErrExit:
KeyFound = False
Exit Function
End Function

Public Function AddItemOnlyToList(ByVal oBizObj As Object) As Boolean
On Error GoTo ErrExit
Dim strStream As String
Dim strCommand As String
Set lvi = pForm.ListView1.SelectedItem
If Not lvi Is Nothing Then
    Set lvi = pForm.ListView1.ListItems.Add(lvi.Index + 1, "K" & oBizObj.GetKey)
    lvi.Text = oBizObj.GetKey
    Dim bSuccess As Boolean
    RaiseEvent AddToList(lvi, oBizObj, bSuccess)
    lvi.Tag = ""
Else
    Set lvi = pForm.ListView1.ListItems.Add(, "K" & oBizObj.GetKey)
    lvi.Text = oBizObj.GetKey
    RaiseEvent AddToList(lvi, oBizObj, bSuccess)
    lvi.Tag = ""
End If
AddItemOnlyToList = True
Exit Function
ErrExit:
AddItemOnlyToList = False
Exit Function
End Function

Public Function FlushBuffer(ByRef pMode As String) As Boolean
On Error GoTo ErrExit
Dim lvi As ListItem
Dim strStream As String
Dim strCommand As String
Dim nCancel As Integer
Dim bSuccess As Boolean
Dim bRetval As Boolean
If pMode = "a" Then
    If KeyFound Then
        MsgBox "Record exists", vbInformation, "Cannot Add"
        GoTo ErrExit
    End If
    strStream = BizObj.GetString
    strCommand = "a" & strStream
    If Not mUseTextFile Then
        If Not mBuffered Then
            nCancel = 0
            RaiseEvent AddItem(BizObj, nCancel)
            If nCancel <> -1 Then
                If Not goDBConn.ExecCommand(TableName, BizObj, strCommand) Then GoTo ErrExit
                RaiseEvent AfterAddItem(BizObj)
                mDirty = False
                bRetval = True
            End If
        Else
            mcolCommands.Add strCommand
            bRetval = True
            mDirty = True
        End If
    Else
        mDirty = True
        bRetval = True
    End If
    Set lvi = pForm.ListView1.SelectedItem
    If Not lvi Is Nothing Then
        Set lvi = pForm.ListView1.ListItems.Add(lvi.Index + 1, "K" & BizObj.GetKey)
        lvi.Text = BizObj.GetKey
        RaiseEvent AddToList(lvi, BizObj, bSuccess)
        lvi.Tag = strStream
    Else
        Set lvi = pForm.ListView1.ListItems.Add(, "K" & BizObj.GetKey)
        lvi.Text = BizObj.GetKey
        RaiseEvent AddToList(lvi, BizObj, bSuccess)
        lvi.Tag = strStream
    End If
    lvi.Selected = True
    pMode = "e"
ElseIf pMode = "e" Then
    Set lvi = pForm.ListView1.SelectedItem
    If Not lvi Is Nothing Then
        strStream = BizObj.GetString
        strCommand = "e" & strStream
        If Not mUseTextFile Then
            If Not mBuffered Then
                RaiseEvent EditItem(BizObj, nCancel)
                If nCancel <> -1 Then
                    If Not goDBConn.ExecCommand(TableName, BizObj, strCommand) Then GoTo ErrExit
                    RaiseEvent AfterEditItem(BizObj)
                    mDirty = False
                    bRetval = True
                End If
            Else
                mcolCommands.Add strCommand
                bRetval = True
                mDirty = True
            End If
        Else
            bRetval = True
            mDirty = True
        End If
        RaiseEvent AddToList(lvi, BizObj, bSuccess)
        lvi.Tag = strStream
        mDirty = True
    End If
End If
FlushBuffer = bRetval
ErrExit:
End Function

Public Function AddItem(ByVal oBizObj As Object) As Boolean
On Error GoTo ErrExit
Set BizObj = oBizObj
FlushBuffer "a"
AddItem = True
Exit Function
ErrExit:
AddItem = False
Exit Function
End Function

Private Sub pAdd_Click()
On Error GoTo ErrExit
Dim bRetval As Boolean
On Error Resume Next
Dim cMode As String
cMode = "a"
bRetval = EditForm.Display(BizObj, cMode, Me)
If Err = 450 Then
    bRetval = EditForm.Display(BizObj, cMode)
ElseIf Err = 0 Then
Else
    GoTo ErrExit
End If
On Error GoTo ErrExit
If bRetval Then
    FlushBuffer cMode
End If
Exit Sub
ErrExit:
Exit Sub
End Sub

Private Sub pEdit_Click()
On Error GoTo ErrExit
Dim bRetval As Boolean
Dim cMode As String
Dim lvi As ListItem
Set lvi = pForm.ListView1.SelectedItem
cMode = "e"
If Not lvi Is Nothing Then
    If BizObj.ReadLine(lvi.Tag) Then
        On Error Resume Next
        bRetval = EditForm.Display(BizObj, cMode, Me)
        If Err = 450 Then
            bRetval = EditForm.Display(BizObj, cMode)
        ElseIf Err = 0 Then
        Else
            GoTo ErrExit
        End If
        On Error GoTo ErrExit
        If bRetval Then
            FlushBuffer cMode
        End If
    End If
End If
Exit Sub
ErrExit:
Exit Sub
End Sub

Private Sub pDelete_Click()
On Error GoTo ErrExit
Dim strCommand As String
Dim nCancel2 As Integer
Dim nCancel3 As Integer
Dim strKey As String
Set lvi = pForm.ListView1.SelectedItem
If Not lvi Is Nothing Then
    nCancel3 = 0
    strKey = Trim(Mid(lvi.Key, 2))
    RaiseEvent AllowDeleteItem(strKey, nCancel3)
    If nCancel3 = -1 Then
        Exit Sub
    End If
    If MsgBox("Delete record '" & Trim(lvi.Text) & "' ?", _
     vbQuestion + vbYesNo + vbDefaultButton2, "Delete") = vbYes Then
        strCommand = "d" & lvi.Key
        If Not mUseTextFile Then
            If Not mBuffered Then
                nCancel2 = 0
                RaiseEvent DeleteItem(strKey, nCancel2)
                If nCancel2 <> -1 Then
                    If Not goDBConn.ExecCommand(TableName, BizObj, strCommand) Then GoTo ErrExit
                    RaiseEvent AfterDeleteItem(strKey)
                End If
            Else
                mcolCommands.Add strCommand
                mDirty = True
            End If
        Else
            mDirty = True
        End If
        pForm.ListView1.ListItems.Remove lvi.Index
    End If
End If
Exit Sub
ErrExit:
Exit Sub
End Sub

Private Sub pSave_Click()
On Error Resume Next
SaveRecords
End Sub

Public Function IsDirty() As Boolean
IsDirty = mDirty
End Function

Public Function AskToSave() As Boolean
If Not mBuffered Then
    AskToSave = True
Else
    If MsgBox("Do you want to save your changes?", vbQuestion + vbYesNo + vbDefaultButton1, "Save Changes") = vbYes Then
        AskToSave = True
    Else
        AskToSave = False
    End If
End If
End Function

Private Sub pClose_Click()
If mDirty And mBuffered Then
    If AskToSave Then
        If Not SaveRecords Then
            If MsgBox("Errors encountered while saving the records. Do you want to continue with close?", vbYesNo + vbDefaultButton2, "Errors Occurred") = vbNo Then
                Exit Sub
            End If
        End If
    End If
End If
AllowClose = True
RaiseEvent CloseForm
End Sub

Private Sub Form_Terminate()
On Error Resume Next
CloseData
End Sub

Public Function Display(ByVal pForm As Form, Optional ByVal pModal As Boolean = False, Optional ByVal pParent As Form = Nothing) As Boolean
On Error Resume Next
If Not ReadData Then
    MsgBox "The database was not found or could not be loaded.", vbCritical + vbOKOnly, "Error"
    Display = False
Else
    InitList
    FillList
    If pModal Then
        If Not pParent Is Nothing Then
            pForm.Show vbModal, pParent
        Else
            pForm.Show vbModal
        End If
    Else
        If Not pParent Is Nothing Then
            pForm.Show , pParent
        Else
            pForm.Show
        End If
    End If
    Display = True
End If
End Function

Private Sub pList_BeforeLabelEdit(Cancel As Integer)
Cancel = -1
End Sub

Private Sub pList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Dim nSortCol As Long
On Error Resume Next
nSortCol = ColumnHeader.Index - 1
pForm.ListView1.SortKey = nSortCol
If mSortCol = nSortCol Then

    'If .SortKey >= 3 Then
    '   LVSortI ListView1, .SortKey, lvDate, .SortOrder
    'Else
    '   LVSortI ListView1, .SortKey, lvNumber, .SortOrder
    'End If

    If pForm.ListView1.SortOrder = lvwDescending Then
        pForm.ListView1.SortOrder = lvwAscending
    Else
        pForm.ListView1.SortOrder = lvwDescending
    End If
Else
    pForm.ListView1.SortOrder = lvwAscending
End If
mSortCol = nSortCol
End Sub

Private Sub pList_DblClick()
If Not NoDoubleClickList Then
    pEdit_Click
End If
End Sub

Private Sub pList_KeyPress(KeyAscii As Integer)
If Not NoDoubleClickList Then
    If KeyAscii = 13 Then pEdit_Click
End If
End Sub

Private Sub InitList()
On Error Resume Next
Dim bSuccess As Boolean
pForm.ListView1.View = lvwReport
pForm.ListView1.HideColumnHeaders = False
pForm.ListView1.AllowColumnReorder = False
pForm.ListView1.FullRowSelect = True
pForm.ListView1.Arrange = lvwAutoTop
RaiseEvent AddColHeader(lvi, bSuccess)

Dim n As Long, strWidths As String, arrTmp() As String, nCol As Long
strWidths = GetSetting(App.Title, "Settings", pForm.name & " ListWidths", "")
If strWidths <> "" Then
    arrTmp = Split(strWidths, ";")
    If ((pForm.ListView1.ColumnHeaders.Count) = (UBound(arrTmp) - LBound(arrTmp) + 1)) Then
        For n = 1 To pForm.ListView1.ColumnHeaders.Count
            pForm.ListView1.ColumnHeaders(n).Width = arrTmp(n - 1)
        Next
    End If
End If

pForm.ListView1.Sorted = True
pForm.ListView1.SortOrder = lvwAscending
pForm.ListView1.SortKey = 1
End Sub

Private Sub pForm_Unload(Cancel As Integer)
Dim n As Long, strWidths As String
strWidths = ""
For n = 1 To pForm.ListView1.ColumnHeaders.Count
    strWidths = strWidths & pForm.ListView1.ColumnHeaders(n).Width & ";"
Next
strWidths = Left(strWidths, Len(strWidths) - 1)
SaveSetting App.Title, "Settings", pForm.name & " ListWidths", strWidths
End Sub

Public Sub SetForm(ByVal oForm As Form)
On Error Resume Next
Set pForm = oForm
Set pList = pForm.Controls("ListView1")
Set pAdd = pForm.Controls("cmdAdd")
Set pEdit = pForm.Controls("cmdEdit")
Set pDelete = pForm.Controls("cmdDelete")
Set pSave = pForm.Controls("cmdSave")
Set pClose = pForm.Controls("cmdClose")
End Sub

Public Sub SetDirty()
mDirty = True
End Sub
