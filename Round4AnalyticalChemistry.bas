Attribute VB_Name = "Round4AnalyticalChemistry"
Sub test()
    d = Array(0.001, 0.0023, 0.0013, 0.0066, 0.0035)
    v = Array(0.001234, 0.02011, 0.20113, 9.801, 123.52)
    Debug.Print vbCrLf
    
    '' 標準偏差の計算
    sd = WorksheetFunction.StDev(d)
    Debug.Print "SD: " & sd
    
    '' 定量下限値の計算
    Debug.Print "LLQ: " & LLQ(sd)
    
    '' 検出下限値の計算
    Debug.Print "LLD: " & LLD(sd)
    
    '' 数値まるめ
    Debug.Print vbCrLf & "数値丸め"
    For Each num In v
        Debug.Print num & " = " & roundSet(num, LLQ(sd), 2)
    Next

    '' 下限値以下の表記も設定
    Debug.Print vbCrLf & "表記変更"
    For Each num In v
        Debug.Print num & " = " & roundSetStyle(num, sd, 2)
    Next
    
End Sub
Function roundJIS(ByVal x As Double, Optional digits As Long = 3)
    Dim dig As Long
    Dim num As Double
    Dim m As Double
    Dim mar As String
    Dim ret As Double
    
    dig = -getDigits(x) + digits
    num = x * WorksheetFunction.Power(10, dig)
    m = num / 5 - Int(num / 5)
    If m = 0 Then
        mar = Right(Int(num / 10), 1)
        If mar Mod 2 = 0 Then
            ret = WorksheetFunction.RoundDown(x, dig - 1)
        Else
            ret = WorksheetFunction.RoundUp(x, dig - 1)
        End If
    Else
        ret = round2(x, digits)
    End If
        
    roundJIS = ret
    
End Function
Function LLD(ByVal sd As Double, Optional ByVal digits As Long = 2)
    Dim lq As Double
    Dim dig As Double
    
    lq = LLQ(sd, digits)
    dig = -(getDigits(lq) - (digits - 1))
    LLD = WorksheetFunction.Round(sd * 3, dig)
End Function
Function LLQ(ByVal sd As Double, Optional ByVal digits As Long = 2)
    LLQ = round2(sd * 10, digits)
End Function
Function print_set_digits(ByVal x As Double, ByVal digits As Long, ByVal realDig As Double) As String
    Dim dig As Double
    Dim set_style As String
    Dim ret As String
    
    dig = getFloorDigits(x, digits)
    If dig >= realDig Then
        ret = Format(x, Replace(dig, "1", "0"))
    Else
        ret = Format(x, Replace(realDig, "1", "0"))
    End If
'    dig = getDigits(x)
'    dig_last = dig - (digits - 1)
'    dig_real = getScreenDigits(x)
'
'    If (dig_last - dig_real) <> 0 Then
'        set_style = getFloorDigits(x, digits)
'        ret = Format(x, Replace(set_style, "1", "0"))
'    Else
'        ret = x
'    End If
    print_set_digits = ret
'    num = getFloorDigits(120.22234)
'    Debug.Print Format(1.000212, Replace(num, "1", "0"))
'
'    dig = getDigits(x)
'    dig_last = dig - (digits - 1)
'    dig_real = getScreenDigits(x)
'    If dig_real > dig_last Then
'        If dig <= 0 Or dig_last < 0 Then
'            set_style = "0." & String(-dig_last, "0")
'            print_set_digits = Format(x, set_style)
'        Else
'            print_set_digits = Format(x, "0")
'        End If
'    Else
'        print_set_digits = x
'    End If
End Function
Function round2(ByVal x As Double, ByVal digits As Long)
    Dim dig As Long
    
    dig = -getDigits(x) + (digits - 1)
    round2 = WorksheetFunction.Round(x, dig)
End Function
Function floor2(ByVal x As Double, ByVal y As Double, Optional ByVal digits As Long = 2)
    Dim dig As Double
    
    dig = getFloorDigits(y, digits)
    floor2 = WorksheetFunction.Floor(x, dig)
End Function
Function getFloorDigits(ByVal x As Double, Optional ByVal digits = 2) As Double
    Dim dig As Double
    
    dig = getDigits(x) - (digits - 1)
    getFloorDigits = WorksheetFunction.Power(10, dig)
End Function
Function getDigits(ByVal x As Double)
    Dim num As Double
    If x = 0 Then
        getDigits = 0
    Else
        num = Log(x) / Log(10)
        getDigits = Int(num)
    End If
End Function
Function getScreenDigits(ByVal x As Double)
    Dim dig As Double
    Dim m As Double
    Dim i As Long
    
    dig = getDigits(x)
    x = x * WorksheetFunction.Power(10, -(dig - 1))
    m = x Mod 10
    Do While m <> 0
        x = x * 10
        dig = dig - 1
        m = x Mod 10
    Loop
    getScreenDigits = dig
End Function
Function roundSet(ByVal x As Double, ByVal y As Double, _
                    ByVal x_digits As Long, Optional ByVal y_digits As Long = 2, _
                    Optional ByVal fitRoundJIS As Boolean = False)
    Dim y_dig As String
    
    If fitRoundJIS = True Then
        x = roundJIS(x, x_digits)
    ElseIf fitRoundJIS = False Then
        x = round2(x, x_digits)
    End If
    y_dig = getFloorDigits(y, y_digits)
    roundSet = print_set_digits(floor2(x, y, y_digits), x_digits, _
                y_dig)
    'roundSet = floor2(x, y, y_digits)
End Function
Function roundSetStyle(ByVal x As Double, ByVal sd As Double, _
                        Optional ByVal digits As Long = 2, Optional ByVal ll_digits As Long = 2, _
                        Optional ByVal prefix1 As String = "<[", Optional ByVal suffix1 As String = "]", _
                        Optional ByVal prefix2 As String = "(", Optional ByVal suffix2 As String = ")", _
                        Optional ByVal fitRoundJIS As Boolean = False)
    Dim ld, lq As Double
    Dim ret As String
    
    lq = LLQ(sd, ll_digits)
    ld = LLD(sd, ll_digits)

    If fitRoundJIS = True Then
        x = roundJIS(x, digits)
    ElseIf fitRoundJIS = False Then
        x = round2(x, digits)
    End If
    If x < ld Then
        ret = prefix1 & ld & suffix1
    ElseIf x < lq Then
        ret = prefix2 & roundSet(x, lq, digits, ll_digits) & suffix2
    Else
        ret = roundSet(x, lq, digits, ll_digits)
    End If
    
    roundSetStyle = ret
    
End Function
