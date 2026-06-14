VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRateLookup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mCol As New Collection
Private moRate As New clsRates
Private moTime As New clsTime

Public IsSilent As Boolean

Public Function ReadRateFile(Optional ByVal pPublic As Boolean = True) As Boolean
On Error Resume Next
ReadRateFile = godb.FillList(mCol, moRate, gFileRateList, pPublic, "", IsSilent, False)
End Function

Public Function FindRate(ByVal pRateType As String, _
    ByRef pRate1 As Double, _
    ByRef pRate2 As Double, _
    ByRef pRate3 As Double, _
    ByRef pRate4 As Double, _
    ByRef pRate5 As Double, _
    Optional ByRef pMember As Integer = 0, _
    Optional ByRef pInet As Integer = 0, _
    Optional ByRef pPassword As String = "", _
    Optional ByRef pDefDay As Integer = 0, _
    Optional ByRef pDefTime As String = "", _
    Optional ByRef pDiscMode1 As String = "", _
    Optional ByRef pDiscMode2 As String = "", _
    Optional ByRef pDiscMode3 As String = "", _
    Optional ByRef pDiscMode4 As String = "", _
    Optional ByRef pDiscMode5 As String = "", _
    Optional ByRef pDefDuration As Double = 0, _
    Optional ByRef pPrepaid As Integer = 0, _
    Optional ByRef pMembersOnly As Boolean = False, _
    Optional ByRef pCommand As String = "", _
    Optional ByRef pPlayAllUCan As String = "", _
    Optional ByRef FreeMin As String = "", Optional ByRef GracePeriod As String = "", Optional ByRef MaxFreeMin As String = "", Optional ByRef MinAmt As String = "", Optional ByRef MinTime As String = "", Optional ByRef RoundAmount As String = "", Optional ByRef RoundBy As String = "", Optional ByRef Autotimeout As String = "", Optional ByRef Opentime As String = "" _
) As Boolean
Dim pRetval As Boolean
Dim n As Integer
Dim strBuff As String
On Error GoTo ErrExit
pMembersOnly = False
If mCol.Count > 0 Then
    For n = 1 To mCol.Count
        strBuff = mCol.item(n)
        If moRate.ReadLine(strBuff) Then
            If (Trim(LCase(moRate.RateType)) = Trim(LCase(pRateType))) Then
                FindRate = True
                pRate1 = moRate.Rate1
                pRate2 = moRate.Rate2
                pRate3 = moRate.Rate3
                pRate4 = moRate.Rate4
                pRate5 = moRate.Rate5
                pMember = Val(moRate.Member)
                pInet = Val(moRate.Inet)
                pPassword = moRate.Password
                pDefDay = moRate.DefDay
                pDefTime = moRate.DefTime
                pDiscMode1 = moRate.DiscMode1
                pDiscMode2 = moRate.DiscMode2
                pDiscMode3 = moRate.DiscMode3
                pDiscMode4 = moRate.DiscMode4
                pDiscMode5 = moRate.DiscMode5
                gRateInet = pInet
                pDefDuration = Val(moRate.DefDuration)
                gRatePrepaid = Val(moRate.Prepaid)
                pPrepaid = gRatePrepaid
                pMembersOnly = (pMember = 1)
                pCommand = moRate.Command
                pPlayAllUCan = moRate.PlayAllUCan
                FreeMin = moRate.FreeMin
                GracePeriod = moRate.MaxFreeMin
                MaxFreeMin = moRate.MaxFreeMin
                MinAmt = moRate.MinAmt
                MinTime = moRate.MinTime
                RoundAmount = moRate.RoundAmount
                RoundBy = moRate.RoundBy
                Autotimeout = moRate.Autotimeout
                Opentime = moRate.Opentime
                Exit Function
            End If
        End If
    Next
End If
ErrExit:
FindRate = False
pRate1 = 0#
pRate2 = 0#
pRate3 = 0#
pRate4 = 0#
pRate5 = 0#
pMember = 0
pInet = 0
pPassword = ""
pDefDay = 0
pDefTime = ""
pDiscMode1 = ""
pDiscMode2 = ""
pDiscMode3 = ""
pDiscMode4 = ""
pDiscMode5 = ""
gRateInet = 0
pDefDuration = 0
gRatePrepaid = 0
pPrepaid = 0
pMembersOnly = False
pCommand = ""
pPlayAllUCan = ""
FreeMin = ""
GracePeriod = ""
MaxFreeMin = ""
MinAmt = ""
MinTime = ""
RoundAmount = ""
RoundBy = ""
Autotimeout = ""
Opentime = ""
End Function

Public Function FindDefaultRate() As Boolean
Dim pRetval As Boolean
Dim n As Integer
Dim strBuff As String
Dim bRetval As Boolean
On Error GoTo ErrExit
ResetDefaults
bRetval = False
If mCol.Count > 0 Then
    For n = 1 To mCol.Count
        strBuff = mCol.item(n)
        If moRate.ReadLine(strBuff) Then
            If (Trim(moRate.DefaultRate) = "1") Then
                GotDefaults
                bRetval = True
            End If
        End If
    Next
End If
FindDefaultRate = bRetval
Exit Function

ErrExit:
FindDefaultRate = False
ResetDefaults
End Function

Public Function FindTimeRate(ByVal pDate As Date) As Boolean
Dim pRetval As Boolean
Dim n As Integer
Dim strBuff As String
Dim bRetval As Boolean
Dim d1 As Integer
Dim d2 As Integer
On Error GoTo ErrExit
bRetval = False
If mCol.Count > 0 Then
    If Not frmServer.RestrictTimeRange Then
        For n = 1 To mCol.Count 'day + time
            strBuff = mCol.item(n)
            If moRate.ReadLine(strBuff) Then
                d1 = Val(moRate.DefDay)
                d2 = Weekday(Now)
                If TestWeekDay(d1, d2) And moTime.InTimeRange(moRate.DefTime, pDate) Then
                    GotDefaults
                    bRetval = True
                    GoTo Done
                End If
            End If
        Next
        For n = 1 To mCol.Count 'days
            strBuff = mCol.item(n)
            If moRate.ReadLine(strBuff) Then
                d1 = Val(moRate.DefDay)
                d2 = Weekday(Now)
                If TestWeekDay(d1, d2) And (Trim(moRate.DefTime) = "") Then
                    GotDefaults
                    bRetval = True
                    GoTo Done
                End If
            End If
        Next
        For n = 1 To mCol.Count 'time
            strBuff = mCol.item(n)
            If moRate.ReadLine(strBuff) Then
                If (Val(moRate.DefDay) = 0) _
                  And moTime.InTimeRange(moRate.DefTime, pDate) Then
                    GotDefaults
                    bRetval = True
                    GoTo Done
                End If
            End If
        Next
    End If
    For n = 1 To mCol.Count 'default
        strBuff = mCol.item(n)
        If moRate.ReadLine(strBuff) Then
            If moRate.DefaultRate Then
                GotDefaults
                bRetval = True
                GoTo Done
            End If
        End If
    Next
    'not found
    ResetDefaults 'reset
End If
Done:
FindTimeRate = bRetval
Exit Function

ErrExit:
FindTimeRate = False
End Function

Private Sub ResetDefaults()
On Error Resume Next
gRemarks = gDefRemarks
gPerHourRate = gDefPerHourRate
g45MinRate = gDef45MinRate
gHalfHourRate = gDefHalfHourRate
gQtrMinRate = gDefQtrMinRate
gPerMinRate = gDefPerMinRate
gRateMember = 0
gRateInet = 0
gRatePassword = ""
gRateDefDay = 0
gRateDefTime = ""
gDiscMode1 = ""
gDiscMode2 = ""
gDiscMode3 = ""
gDiscMode4 = ""
gDiscMode5 = ""
gRateDefDuration = 0
gRatePrepaid = 0
gRateCommand = ""
gRatePlayAllUCan = ""
gRTFreeMin = ""
gRTGracePeriod = ""
gRTMaxFreeMin = ""
gRTMinAmt = ""
gRTMinTime = ""
gRTRoundAmount = ""
gRTRoundBy = ""
End Sub

Private Sub GotDefaults()
On Error Resume Next
gRemarks = moRate.RateType
gPerHourRate = moRate.Rate1
g45MinRate = moRate.Rate2
gHalfHourRate = moRate.Rate3
gQtrMinRate = moRate.Rate4
gPerMinRate = moRate.Rate5
gRateMember = Val(moRate.Member)
gRateInet = Val(moRate.Inet)
gRatePassword = moRate.Password
gRateDefDay = moRate.DefDay
gRateDefTime = moRate.DefTime
gDiscMode1 = moRate.DiscMode1
gDiscMode2 = moRate.DiscMode2
gDiscMode3 = moRate.DiscMode3
gDiscMode4 = moRate.DiscMode4
gDiscMode5 = moRate.DiscMode5
gRateDefDuration = Val(moRate.DefDuration)
gRatePrepaid = Val(moRate.Prepaid)
gRateCommand = moRate.Command
gRatePlayAllUCan = moRate.PlayAllUCan
gRTFreeMin = moRate.FreeMin
gRTGracePeriod = moRate.GracePeriod
gRTMaxFreeMin = moRate.MaxFreeMin
gRTMinAmt = moRate.MinAmt
gRTMinTime = moRate.MinTime
gRTRoundAmount = moRate.RoundAmount
gRTRoundBy = moRate.RoundBy
End Sub

Public Sub FillRateType(ByVal pList As Control)
Dim n As Long
Dim i As Long
Dim strBuff As String
On Error GoTo ErrExit
pList.Clear
If mCol.Count > 0 Then
    For n = 1 To mCol.Count
        strBuff = mCol.item(n)
        If moRate.ReadLine(strBuff) Then
            pList.AddItem moRate.RateType
        End If
    Next
End If
ErrExit:
End Sub

Public Sub FillPublicRateType(ByVal pList As Control)
Dim n As Long
Dim i As Long
Dim strBuff As String
On Error GoTo ErrExit
pList.Clear
If mCol.Count > 0 Then
    For n = 1 To mCol.Count
        strBuff = mCol.item(n)
        If moRate.ReadLine(strBuff) Then
            If moRate.Member = 1 Then
                pList.AddItem moRate.RateType
            End If
        End If
    Next
End If
ErrExit:
End Sub

Public Function FindDefaultPublicRate() As Boolean
Dim pRetval As Boolean
Dim n As Integer
Dim strBuff As String
Dim bRetval As Boolean
On Error GoTo ErrExit
ResetDefaults
bRetval = False
If mCol.Count > 0 Then
    For n = 1 To mCol.Count
        strBuff = mCol.item(n)
        If moRate.ReadLine(strBuff) Then
            If moRate.Member = 1 Then
                If (Trim(moRate.DefaultRate) = "1") Then
                    GotDefaults
                    bRetval = True
                End If
            End If
        End If
    Next
End If
FindDefaultPublicRate = bRetval
Exit Function

ErrExit:
FindDefaultPublicRate = False
ResetDefaults
End Function

Public Sub AssertRates()
If Not gUseFreeFormRemarks Then
    gUpdateRateOut = True
    gUpdateRateIn = True
    ReadRateFile
    FindDefaultRate
End If
End Sub
