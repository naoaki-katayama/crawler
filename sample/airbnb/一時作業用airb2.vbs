Sub test()
Dim strX1 As Variant
Dim strY1 As Variant
Dim strX2 As Variant
Dim strY2 As Variant
Dim CheckX(1 to 13) As Long
Dim CheckY(1 to 13) As Long

For i = 1 to 13
CheckX(i) = 0
CheckY(i) = 0
Next i

CheckX(1) = 1
CheckY(1) = 0

Digit = 0

strX1 = 10 ^14
strY1 = 10 ^14
strX2 = ""
strY2 = ""

For j = 1 to 13
  If CheckX(j) > 0 Then
    strX1 = strX1 + CheckX(j) * 10 ^ (14 - j)
    Debug.Print strX1
  End If
  If CheckY(j) > 0 Then
    strY1 = strY1 + CheckY(j) * 10 ^ (14 - j)
  End If
Next j

strX2 = strX1 + 10 ^ (14 - Digit)
strY2 = strY1 + 10 ^ (14 - Digit)

strX1 = Mid(strX1,1,1) & "." & Mid(strX1,2)
strY1 = Mid(strY1,1,1) & "." & Mid(strY1,2)
strX2 = Mid(strX2,1,1) & "." & Mid(strX2,2)
strY2 = Mid(strY2,1,1) & "." & Mid(strY2,2)

strX1 =strX1 + 10

Debug.Print "X1:" & 10 + strX1
Debug.Print "Y1:" & 10 + strY1
Debug.Print "X2:" & 10 + strX2
Debug.Print "Y2:" & 10 + strY2


End Sub
