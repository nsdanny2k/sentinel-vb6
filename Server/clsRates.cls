VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public RateType As String
Public Rate1 As String
Public Rate2 As String
Public Rate3 As String
Public Rate4 As String
Public Rate5 As String
Public DefaultRate As String
Public DefDuration As String
Public Member As String
Public Inet As String
Public Password As String
Public DefDay As String
Public DefTime As String
Public DiscMode1 As String
Public DiscMode2 As String
Public DiscMode3 As String
Public DiscMode4 As String
Public DiscMode5 As String
Public Prepaid As String
Public Command As String
Public PlayAllUCan As String

Public FreeMin As String
Public GracePeriod As String
Public MaxFreeMin As String
Public MinAmt As String
Public MinTime As String
Public RoundAmount As String
Public RoundBy As String

Public AutoTimeout As String
Public OpenTime As String

Public PlayAllUCanAmount As Double
Public PlayAllUCanDay As String
Public PlayAllUCanDayCode As Long
Public PlayAllUCanTimeRange As String

Private moTime As New clsTime

Private Function ReadRate(ByVal pRate As String) As Boolean
On Error GoTo ErrExit
Dim strDay As String
Dim arrTmp() As String
Dim strTime As String
If pRate = "" Then GoTo ErrExit
arrTmp = Split(pRate, gRecDelim4)
PlayAllUCanAmount = arrTmp(0)
PlayAllUCanDay = arrTmp(1)
Dim oTime As New clsTime
PlayAllUCanDayCode = oTime.GetDayCode(PlayAllUCanDay)
PlayAllUCanTimeRange = arrTmp(2)
ReadRate = True
Exit Function

ErrExit:
PlayAllUCanAmount = 0
PlayAllUCanDay = 0
PlayAllUCanTimeRange = ""
Exit Function
End Function

Public Function GetFields() As Variant
GetFields = Array("RateType", _
                    "Rate1", _
                    "Rate2", _
                    "Rate3", _
                    "Rate4", _
                    "Rate5", _
                    "DefaultRate", _
                    "Member", _
                    "Inet", _
                    "Password", _
                    "DefDay", _
                    "DefTime", _
                    "DiscMode1", _
                    "DiscMode2", _
                    "DiscMode3", _
                    "DiscMode4", _
                    "DiscMode5", _
                    "DefDuration", _
                    "Prepaid", _
                    "Command", _
                    "PlayAllUCan", _
                    "FreeMin", "GracePd", "MaxFreeMin", "MinAmt", "MinTime", "RoundAmount", "RoundBy", "AutoOut", "OpenOut")
End Function

Public Function GetFieldTypes() As Variant
GetFieldTypes = Array("S14", _
                        "C", _
                        "C", _
                        "C", _
                        "C", _
                        "C", _
                        "S7", _
                        "S2", _
                        "S2", _
                        "S15", _
                        "S2", _
                        "S15", _
                        "S35", _
                        "S35", _
                        "S35", _
                        "S35", _
                        "S35", _
                        "S35", _
                        "S1", _
                        "S255", _
                        "S200", _
                        "S50", "S50", "S50", "S50", "S50", "S50", "S50", "S1", "S1")
End Function

Public Function GetKeyFieldType() As String
GetKeyFieldType = "S14"
End Function

Public Function GetKeyFieldName() As String
GetKeyFieldName = "RateType"
End Function

Public Function ReadLine(ByVal pData As String) As Boolean
Dim aBuff() As String
On Error GoTo ErrExit
Member = "0"
Inet = "0"
Password = ""
DefDay = "0"
DefTime = ""
DiscMode1 = ""
DiscMode2 = ""
DiscMode3 = ""
DiscMode4 = ""
DiscMode5 = ""
aBuff = Split(pData, gRecDelim)
RateType = UCase(Trim(aBuff(0)))
Rate1 = Trim(aBuff(1))
Rate2 = Trim(aBuff(2))
Rate3 = Trim(aBuff(3))
Rate4 = Trim(aBuff(4))
Rate5 = Trim(aBuff(5))
DefaultRate = Trim(aBuff(6))
On Error Resume Next
Member = Trim(aBuff(7))
If Member <> "0" And Member <> "1" Then
    Member = "0"
End If
Inet = Trim(aBuff(8))
If Inet <> "0" And Inet <> "1" Then
    Inet = "0"
End If
Password = Trim(aBuff(9))
DefDay = Trim(aBuff(10))
If DefDay <> "0" And DefDay <> "1" And DefDay <> "2" And DefDay <> "3" And DefDay <> "4" And DefDay <> "5" And DefDay <> "6" And DefDay <> "7" Then
    DefDay = "0"
End If
DefTime = Trim(aBuff(11))
DiscMode1 = Trim(aBuff(12))
DiscMode2 = Trim(aBuff(13))
DiscMode3 = Trim(aBuff(14))
DiscMode4 = Trim(aBuff(15))
DiscMode5 = Trim(aBuff(16))
DefDuration = Trim(aBuff(17))
Prepaid = Trim(aBuff(18))
If Prepaid <> "0" And Prepaid <> "1" Then
    Prepaid = "0"
End If
On Error Resume Next
Command = Trim(aBuff(19))

PlayAllUCan = ""
PlayAllUCan = Trim(aBuff(20))
If PlayAllUCan <> "" Then
    Call ReadRate(PlayAllUCan)
End If

FreeMin = Trim(aBuff(21))
GracePeriod = Trim(aBuff(22))
MaxFreeMin = Trim(aBuff(23))
MinAmt = Trim(aBuff(24))
MinTime = Trim(aBuff(25))
RoundAmount = Trim(aBuff(26))
RoundBy = Trim(aBuff(27))
 
AutoTimeout = Trim(aBuff(28))
OpenTime = Trim(aBuff(29))
 
 
ReadLine = True
Exit Function
ErrExit:
Member = "0"
Inet = "0"
Password = ""
DefDay = "0"
DefTime = ""
DiscMode1 = ""
DiscMode2 = ""
DiscMode3 = ""
DiscMode4 = ""
DiscMode5 = ""
RateType = ""
Rate1 = "0"
Rate2 = "0"
Rate3 = "0"
Rate4 = "0"
Rate5 = "0"
DefaultRate = "0"
DefDuration = "0"
Prepaid = "0"
ReadLine = False
Command = ""
PlayAllUCan = ""

FreeMin = ""
GracePeriod = ""
MaxFreeMin = ""
MinAmt = ""
MinTime = ""
RoundAmount = ""
RoundBy = ""
AutoTimeout = ""
OpenTime = ""
Exit Function
End Function

Public Function GetString() As String
GetString = UCase(Trim(RateType)) & gRecDelim & _
            Trim(Format(Rate1, gAmountFormat1)) & gRecDelim & _
            Trim(Format(Rate2, gAmountFormat1)) & gRecDelim & _
            Trim(Format(Rate3, gAmountFormat1)) & gRecDelim & _
            Trim(Format(Rate4, gAmountFormat1)) & gRecDelim & _
            Trim(Format(Rate5, gAmountFormat1)) & gRecDelim & _
            Trim(DefaultRate) & gRecDelim & _
            Trim(Member) & gRecDelim & _
            Trim(Inet) & gRecDelim & _
            Trim(Password) & gRecDelim & _
            Trim(DefDay) & gRecDelim & _
            Trim(DefTime) & gRecDelim & _
            Trim(DiscMode1) & gRecDelim & _
            Trim(DiscMode2) & gRecDelim & _
            Trim(DiscMode3) & gRecDelim & _
            Trim(DiscMode4) & gRecDelim & _
            Trim(DiscMode5) & gRecDelim & _
            Trim(DefDuration) & gRecDelim & _
            Trim(Prepaid) & gRecDelim & _
            Trim(Command) & gRecDelim & _
            Trim(PlayAllUCan) & gRecDelim & _
            Trim(FreeMin) & gRecDelim & Trim(GracePeriod) & gRecDelim & Trim(MaxFreeMin) & gRecDelim & Trim(MinAmt) & gRecDelim & Trim(MinTime) & gRecDelim & Trim(RoundAmount) & gRecDelim & Trim(RoundBy) & gRecDelim & Trim(AutoTimeout) & gRecDelim & Trim(OpenTime)
End Function

Public Function GetKey() As String
GetKey = RateType
End Function

Public Sub SetDefault(ByVal pData As Boolean)
If pData Then
    DefaultRate = "1"
Else
    DefaultRate = "0"
End If
End Sub

Public Function Field(ByVal pName As String)
Dim pRetval
Select Case LCase(pName)
    Case "ratetype"
        pRetval = RateType
    Case "rate1"
        pRetval = Rate1
    Case "rate2"
        pRetval = Rate2
    Case "rate3"
        pRetval = Rate3
    Case "rate4"
        pRetval = Rate4
    Case "rate5"
        pRetval = Rate5
    Case "defaultrate"
        pRetval = DefaultRate
    Case "default"
        pRetval = DefaultRate
    Case "member"
        pRetval = Member
    Case "inet"
        pRetval = Inet
    Case "password"
        pRetval = Password
    Case "defday"
        pRetval = DefDay
    Case "deftime"
        pRetval = DefTime
    Case "discmode1"
        pRetval = DiscMode1
    Case "discmode2"
        pRetval = DiscMode2
    Case "discmode3"
        pRetval = DiscMode3
    Case "discmode4"
        pRetval = DiscMode4
    Case "discmode5"
        pRetval = DiscMode5
    Case "defduration"
        pRetval = DefDuration
    Case "prepaid"
        pRetval = Prepaid
    Case "command"
        pRetval = Command
    Case "playallucan"
        pRetval = PlayAllUCan
    Case "freemin"
        pRetval = FreeMin
    Case "graceperiod"
        pRetval = GracePeriod
    Case "maxfreemin"
        pRetval = MaxFreeMin
    Case "minamt"
        pRetval = MinAmt
    Case "mintime"
        pRetval = MinTime
    Case "roundamount"
        pRetval = RoundAmount
    Case "roundby"
        pRetval = RoundBy
    Case "autotimeout"
        pRetval = AutoTimeout
    Case "opentime"
        pRetval = OpenTime
End Select
Field = pRetval
End Function

Public Function ValidateRateType(ByVal pData As String) As Boolean
Dim pRetval As Boolean
pData = Trim(pData)
If pData = "" Then
    MsgBox "Rate account is required", vbInformation, "Invalid Rate Account"
ElseIf Len(pData) > 15 Then
    MsgBox "Max rate account length is 15", vbInformation, "Invalid Rate Account"
ElseIf InStr(1, pData, gRecDelim, vbTextCompare) <> 0 Then
    MsgBox "Rate account cannot contain a " & gRecDelim, vbInformation, "Invalid Rate Account"
Else
    pRetval = True
End If
ValidateRateType = pRetval
End Function

Public Function ValidateRate(ByVal pData) As Boolean
Dim pRetval As Boolean
pData = Trim(pData)
If Not IsNumeric(pData) Then
    MsgBox "Rate must be numeric", vbInformation, "Invalid Rate Type"
Else
    pRetval = True
End If
ValidateRate = pRetval
End Function

Public Function ValidatePassword(ByVal pData As String) As Boolean
Dim pRetval As Boolean
pData = Trim(pData)
If Len(pData) > 15 Then
    MsgBox "Max password length is 15", vbInformation, "Invalid Password"
ElseIf InStr(1, pData, gRecDelim, vbTextCompare) <> 0 Then
    MsgBox "Password cannot contain a " & gRecDelim, vbInformation, "Invalid Password"
Else
    pRetval = True
End If
ValidatePassword = pRetval
End Function

Public Function ValidateDiscMode(ByVal pData As String) As Boolean
Dim pRetval As Boolean
pData = Trim(pData)
If Len(pData) > 35 Then
    MsgBox "Max discount mode length is 35", vbInformation, "Invalid Discount Mode"
ElseIf InStr(1, pData, gRecDelim, vbTextCompare) <> 0 Then
    MsgBox "Discount mode cannot contain a comma", vbInformation, "Invalid Discount Mode"
Else
    pRetval = True
End If
ValidateDiscMode = pRetval
End Function

Public Function ValidateDuration(ByVal pData) As Boolean
Dim pRetval As Boolean
pData = Trim(pData)
If Not IsNumeric(pData) Then
    MsgBox "Default duration must be numeric", vbInformation, "Invalid Default Duration"
Else
    pRetval = True
End If
ValidateDuration = pRetval
End Function

Public Function ValidatePlayAllUCan(ByVal pData As String) As Boolean
Dim pRetval As Boolean
pData = Trim(pData)
If Len(pData) > 35 Then
    MsgBox "Max play all u can definition length is 35", vbInformation, "Invalid Play All U Can"
ElseIf InStr(1, pData, gRecDelim, vbTextCompare) <> 0 Then
    MsgBox "Play all u can cannot contain a comma", vbInformation, "Invalid Play All U Can"
Else
    pRetval = True
End If
ValidatePlayAllUCan = pRetval
End Function
