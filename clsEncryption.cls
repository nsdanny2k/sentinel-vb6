VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEncryption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'1 byte - default
'2 simple
'3 blowfish

Private Const p1 = 1
Private Const p2 = 3
Private Const p3 = 2
Private Const p4 = 9
Private Const p5 = 9
Private Const p6 = 3
Private Const p7 = 7
Private Const p8 = 5
Private Const p9 = 3
Private Const p10 = 1

Public EncryptionMethod As Integer

Private Function Encrypt1(pc_string, p1 As Integer, p2 As Integer, p3 As Integer, p4 As Integer, p5 As Integer, p6 As Integer, p7 As Integer, p8 As Integer, p9 As Integer, p10 As Integer)
Dim ByteArray() As Byte
On Error GoTo ErrExit
ByteArray() = StrConv(pc_string, vbFromUnicode)
Call EncryptByte(ByteArray(), 100)
Encrypt1 = StrConv(ByteArray(), vbUnicode)
Exit Function
ErrExit:
Encrypt1 = ""
Exit Function
End Function

Private Function Decrypt1(pc_string, p1 As Integer, p2 As Integer, p3 As Integer, p4 As Integer, p5 As Integer, p6 As Integer, p7 As Integer, p8 As Integer, p9 As Integer, p10 As Integer)
Dim ByteArray() As Byte
On Error GoTo ErrExit
ByteArray() = StrConv(pc_string, vbFromUnicode)
Call DecryptByte(ByteArray(), 100)
Decrypt1 = StrConv(ByteArray(), vbUnicode)
Exit Function
ErrExit:
Decrypt1 = ""
Exit Function
End Function

Private Function EncryptByte(ByteArray() As Byte, Optional Key As Byte = 0)
Dim n As Integer
Dim u As Integer
Dim l As Integer
l = LBound(ByteArray)
u = UBound(ByteArray)
If l >= 0 Then
    For n = l To u
        ByteArray(n) = ByteArray(n) + Key
    Next
End If
Exit Function
ErrExit:
Exit Function
End Function

Private Function DecryptByte(ByteArray() As Byte, Optional Key As Byte = 0)
Dim n As Integer
Dim u As Integer
Dim l As Integer
l = LBound(ByteArray)
u = UBound(ByteArray)
On Error GoTo ErrExit
If l >= 0 Then
    For n = l To u
        ByteArray(n) = ByteArray(n) - Key
    Next
End If
Exit Function
ErrExit:
Exit Function
End Function

'method 2

Private Function Encrypt2(pc_string, p1 As Integer, p2 As Integer, p3 As Integer, p4 As Integer, p5 As Integer, p6 As Integer, p7 As Integer, p8 As Integer, p9 As Integer, p10 As Integer)
    Dim lc_estr, ln_long, ln_pos, pc_func
    lc_estr = ""
    ln_long = Len(pc_string)
    For ln_pos = 1 To ln_long
        lc_estr = lc_estr + Chr(Asc(Mid(pc_string, ln_pos, 1)) + ln_pos + ln_long)
    Next ln_pos
    Encrypt2 = lc_estr
End Function

Private Function Decrypt2(pc_string, p1 As Integer, p2 As Integer, p3 As Integer, p4 As Integer, p5 As Integer, p6 As Integer, p7 As Integer, p8 As Integer, p9 As Integer, p10 As Integer)
    Dim lc_estr, ln_long, ln_pos, pc_func
    lc_estr = ""
    ln_long = Len(pc_string)
    For ln_pos = 1 To ln_long
        lc_estr = lc_estr + Chr(Asc(Mid(pc_string, ln_pos, 1)) - ln_pos - ln_long)
    Next ln_pos
    Decrypt2 = lc_estr
End Function

'method 3
Private Function EncDec(inData As Variant, Optional inPW As Variant = "") As Variant
     On Error Resume Next
     Dim arrSBox(0 To 255) As Integer
     Dim arrPW(0 To 255) As Integer
     Dim bi As Integer, Bj As Integer
     Dim mKey As Integer
     Dim i As Integer, j As Integer
     Dim x As Integer, y As Integer
     Dim mCode As Byte, mCodeSeries As Variant
    
     EncDec = ""
     If Trim(inData) = "" Then
         Exit Function
     End If
    
     If inPW <> "" Then
         j = 1
         For i = 0 To 255
             arrPW(i) = Asc(Mid$(inPW, j, 1))
             j = j + 1
             If j > Len(inPW) Then
                  j = 1
             End If
         Next i
     Else
         For i = 0 To 255
             arrPW(i) = 0
         Next i
     End If
     
       ' Reseed arrSBox()
     For i = 0 To 255
         arrSBox(i) = i
     Next i
     
     j = 0
     For i = 0 To 255
         j = (arrSBox(i) + arrPW(i)) Mod 256
           ' Swap
         x = arrSBox(i)
         arrSBox(i) = arrSBox(j)
         arrSBox(j) = x
     Next i
     
     mCodeSeries = ""
     bi = 0: Bj = 0
     For i = 1 To Len(inData)
         bi = (bi + 1) Mod 256
         Bj = (Bj + arrSBox(bi)) Mod 256
           ' Swap
         x = arrSBox(bi)
         arrSBox(bi) = arrSBox(Bj)
         arrSBox(Bj) = x
            'Generate a key
         mKey = arrSBox((arrSBox(bi) + arrSBox(Bj)) Mod 256)
            'xor the key
         mCode = Asc(Mid$(inData, i, 1)) Xor mKey
         mCodeSeries = mCodeSeries & Chr(mCode)
     Next i
     EncDec = mCodeSeries
End Function

Public Function DecryptText(ByVal pString As String) As String
On Error GoTo ErrExit
Select Case EncryptionMethod
    Case 1
        DecryptText = Decrypt1(pString, p1, p2, p3, p4, p5, p6, p7, p8, p9, p10)
    Case 2
        DecryptText = Decrypt2(pString, p1, p2, p3, p4, p5, p6, p7, p8, p9, p10)
    Case 3
        DecryptText = EncDec(pString)
End Select
Exit Function
ErrExit:
DecryptText = ""
Exit Function
End Function

Public Function EncryptText(ByVal pString As String) As String
On Error GoTo ErrExit
Select Case EncryptionMethod
    Case 1
        EncryptText = Encrypt1(pString, p1, p2, p3, p4, p5, p6, p7, p8, p9, p10)
    Case 2
        EncryptText = Encrypt2(pString, p1, p2, p3, p4, p5, p6, p7, p8, p9, p10)
    Case 3
        EncryptText = EncDec(pString)
End Select
Exit Function
ErrExit:
EncryptText = ""
Exit Function
End Function

Private Sub Class_Initialize()
EncryptionMethod = 1
End Sub

Public Function RandomPW(myLength, Optional ByVal pNumeric As Boolean = True)
    'These constant are the minimum and maximum length for random
    'length passwords.  Adjust these values to your needs.
    Const minLength = 6
    Const maxLength = 20
    Dim x, y, strPW, strChar
    If myLength = 0 Then
        Randomize
        myLength = Int((maxLength * Rnd) + minLength)
    End If
    For x = 1 To myLength
        'Randomize the type of this character
        If pNumeric Then
            Do
                y = Int((3 * Rnd) + 1) '(1) Numeric, (2) Uppercase, (3) Lowercase
            Loop While y <> 1
        Else
            y = Int((3 * Rnd) + 1) '(1) Numeric, (2) Uppercase, (3) Lowercase
        End If
        Select Case y
            Case 1
                'Numeric character
                Randomize
                If x = 1 Then
                    Do
                        strChar = Chr(Int((9 * Rnd) + 48))
                    Loop While strChar = "0"
                Else
                    strChar = Chr(Int((9 * Rnd) + 48))
                End If
                strPW = strPW & strChar
            Case 2
                'Uppercase character
                Randomize
                strPW = strPW & Chr(Int((25 * Rnd) + 65))
            Case 3
                'Lowercase character
                Randomize
                strPW = strPW & Chr(Int((25 * Rnd) + 97))

        End Select
    Next
    RandomPW = Left(strPW, myLength)
End Function

Private Function GetCharEncr(ByVal c As String) As String
Dim cRetval As String
c = UCase(c)
Select Case c
Case "A"
cRetval = "M"
Case "B"
cRetval = "1"
Case "C"
cRetval = "H"
Case "D"
cRetval = "0"
Case "E"
cRetval = "L"
Case "F"
cRetval = "3"
Case "G"
cRetval = "U"
Case "H"
cRetval = "Z"
Case "I"
cRetval = "I"
Case "J"
cRetval = "T"
Case "K"
cRetval = "G"
Case "L"
cRetval = "R"
Case "M"
cRetval = "Y"
Case "N"
cRetval = "J"
Case "O"
cRetval = "X"
Case "P"
cRetval = "2"
Case "Q"
cRetval = "S"
Case "R"
cRetval = "4"
Case "S"
cRetval = "6"
Case "T"
cRetval = "N"
Case "U"
cRetval = "W"
Case "V"
cRetval = "5"
Case "W"
cRetval = "D"
Case "X"
cRetval = "E"
Case "Y"
cRetval = "C"
Case "Z"
cRetval = "Q"
Case "0"
cRetval = "V"
Case "1"
cRetval = "K"
Case "2"
cRetval = "B"
Case "3"
cRetval = "7"
Case "4"
cRetval = "F"
Case "5"
cRetval = "8"
Case "6"
cRetval = "O"
Case "7"
cRetval = "9"
Case "8"
cRetval = "A"
Case "9"
cRetval = "P"
Case Else
cRetval = c
End Select
GetCharEncr = cRetval
End Function

Private Function GetCharDecr(ByVal c As String) As String
Dim cRetval As String
c = UCase(c)
Select Case c
Case "A"
cRetval = "M"
Case "M"
cRetval = "A"
Case "1"
cRetval = "B"
Case "H"
cRetval = "C"
Case "0"
cRetval = "D"
Case "L"
cRetval = "E"
Case "3"
cRetval = "F"
Case "U"
cRetval = "G"
Case "Z"
cRetval = "H"
Case "I"
cRetval = "I"
Case "T"
cRetval = "J"
Case "G"
cRetval = "K"
Case "R"
cRetval = "L"
Case "Y"
cRetval = "M"
Case "J"
cRetval = "N"
Case "X"
cRetval = "O"
Case "2"
cRetval = "P"
Case "S"
cRetval = "Q"
Case "4"
cRetval = "R"
Case "6"
cRetval = "S"
Case "N"
cRetval = "T"
Case "W"
cRetval = "U"
Case "5"
cRetval = "V"
Case "D"
cRetval = "W"
Case "E"
cRetval = "X"
Case "C"
cRetval = "Y"
Case "Q"
cRetval = "Z"
Case "V"
cRetval = "0"
Case "K"
cRetval = "1"
Case "B"
cRetval = "2"
Case "7"
cRetval = "3"
Case "F"
cRetval = "4"
Case "8"
cRetval = "5"
Case "O"
cRetval = "6"
Case "9"
cRetval = "7"
Case "A"
cRetval = "8"
Case "P"
cRetval = "9"
Case Else
cRetval = c
End Select
GetCharDecr = cRetval
End Function

Public Function SimpleEncrypt(pc_string, Optional ByVal pMethod As String = 1)
On Error GoTo ErrExit
Dim n As Long, i As Long, s As String
n = Len(pc_string)
If n > 0 Then
    For i = 1 To n
        s = s & GetCharEncr(Mid(pc_string, i, 1))
    Next
End If
SimpleEncrypt = s
Exit Function
ErrExit:
SimpleEncrypt = ""
Exit Function
End Function

Public Function SimpleDecrypt(pc_string, Optional ByVal pMethod As String = 1)
On Error GoTo ErrExit
Dim n As Long, i As Long, s As String
n = Len(pc_string)
If n > 0 Then
    For i = 1 To n
        s = s & GetCharDecr(Mid(pc_string, i, 1))
    Next
End If
SimpleDecrypt = s
Exit Function
ErrExit:
SimpleDecrypt = ""
Exit Function
End Function
