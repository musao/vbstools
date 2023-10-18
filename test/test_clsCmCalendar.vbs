' clsCmCalendar.vbs: test.
' @import ../lib/clsCmArray.vbs
' @import ../lib/clsCmBufferedReader.vbs
' @import ../lib/clsCmBufferedWriter.vbs
' @import ../lib/clsCmCalendar.vbs
' @import ../lib/clsCmBroker.vbs
' @import ../lib/clsCompareExcel.vbs
' @import ../lib/clsFsBase.vbs
' @import ../lib/libCom.vbs

Option Explicit

'###################################################################################################
'clsCmCalendar
Sub Test_clsCmCalendar
    Dim a : Set a = new clsCmCalendar
    AssertEqual 8, VarType(a)
    AssertEqual "clsCmCalendar", TypeName(a)
End Sub

'###################################################################################################
'clsCmCalendar.getNow()/toString()
Sub Test_clsCmCalendar_getNow_toString
    Dim y,m,d,h,mm,s
    y = Right("000" & Year(now()), 4)
    m = Right("0" & Month(now()), 2)
    d = Right("0" & Day(now()), 2)
    h = Right("0" & Hour(now()), 2)
    mm = Right("0" & Minute(now()), 2)
    s = Right("0" & Second(now()), 2)
    Dim ptn : ptn = "^"&y&"/"&m&"/"&d&" "&h&":"&mm&":"&s&"\.\d{3}$"
    Dim a : Set a = (new clsCmCalendar).getNow()

    AssertMatch ptn, a.toString()
    AssertMatch ptn, a
End Sub

'###################################################################################################
'clsCmCalendar.setDateTime()/toString()
Sub Test_clsCmCalendar_setDateTime_toString
    Dim e : e = "2024/02/29 00:59:30"
    Dim a : Set a = (new clsCmCalendar).setDateTime(e)

    AssertMatch e & ".000", a.toString()
End Sub
Sub Test_clsCmCalendar_setDateTime_toString_Err
    On Error Resume Next
    Dim e : e = "2022/02/29 00:59:30"
    Dim a : Set a = (new clsCmCalendar).setDateTime(e)

    AssertEqual 13, Err.Number
    AssertEqual "å^Ç™àÍívÇµÇ‹ÇπÇÒÅB", Err.Description
    AssertEqual Empty, a
End Sub
Sub Test_clsCmCalendar_setDateTime_WithDecimal_toString
    Dim e : e = "2023/12/31 23:30:59.123456"
    Dim a : Set a = (new clsCmCalendar).setDateTime(e)

    AssertMatch mid(e,1,Len(a.toString())), a.toString()
End Sub
Sub Test_clsCmCalendar_setDateTime_WithDecimal_toString_Err
    On Error Resume Next
    Dim e : e = "2022/02/29 00:59:30.123456"
    Dim a : Set a = (new clsCmCalendar).setDateTime(e)

    AssertEqual 13, Err.Number
    AssertEqual "å^Ç™àÍívÇµÇ‹ÇπÇÒÅB", Err.Description
    AssertEqual Empty, a
End Sub

'###################################################################################################
'clsCmCalendar.formatAs()
Sub Test_clsCmCalendar_formatAs_YYYY
    Dim d,f,e,a
    d = "2024/02/29 00:59:30"
    f = "YYYY"
    e = Right("000" & Year(Left(d,19)), 4)
    Set a = (new clsCmCalendar).setDateTime(d)

    AssertEqual e, a.formatAs(f)

    f = LCase(f)
    AssertEqual e, a.formatAs(f)
End Sub
Sub Test_clsCmCalendar_formatAs_YY
    Dim d,f,e,a
    d = "2024/02/29 00:59:30.123456"
    f = "YY"
    e = Right("0" & Year(Left(d,19)), 2)
    Set a = (new clsCmCalendar).setDateTime(d)

    AssertEqual e, a.formatAs(f)

    f = LCase(f)
    AssertEqual e, a.formatAs(f)
End Sub
Sub Test_clsCmCalendar_formatAs_Month_MM_1
    Dim d,f,e,a
    d = "2024/02/29 00:59:30"
    f = "MM"
    e = Right("0" & Month(Left(d,19)), 2)
    Set a = (new clsCmCalendar).setDateTime(d)

    AssertEqual e, a.formatAs(f)
End Sub
Sub Test_clsCmCalendar_formatAs_Month_MM_2
    Dim d,f,e,a
    d = "2024/10/29 00:59:30.123456"
    f = "MM"
    e = Right("0" & Month(Left(d,19)), 2)
    Set a = (new clsCmCalendar).setDateTime(d)

    AssertEqual e, a.formatAs(f)
End Sub
Sub Test_clsCmCalendar_formatAs_Month_M_1
    Dim d,f,e,a
    d = "2024/02/29 00:59:30"
    f = "M"
    e = Cstr(Month(Left(d,19)))
    Set a = (new clsCmCalendar).setDateTime(d)

    AssertEqual e, a.formatAs(f)
End Sub
Sub Test_clsCmCalendar_formatAs_Month_M_2
    Dim d,f,e,a
    d = "2024/11/29 00:59:30.123456"
    f = "M"
    e = Cstr(Month(Left(d,19)))
    Set a = (new clsCmCalendar).setDateTime(d)

    AssertEqual e, a.formatAs(f)
End Sub
Sub Test_clsCmCalendar_formatAs_DD_1
    Dim d,f,e,a
    d = "2024/02/09 00:59:30"
    f = "DD"
    e = Right("0" & Day(Left(d,19)), 2)
    Set a = (new clsCmCalendar).setDateTime(d)

    AssertEqual e, a.formatAs(f)

    f = LCase(f)
    AssertEqual e, a.formatAs(f)
End Sub
Sub Test_clsCmCalendar_formatAs_DD_2
    Dim d,f,e,a
    d = "2024/02/29 00:59:30.123456"
    f = "DD"
    e = Right("0" & Day(Left(d,19)), 2)
    Set a = (new clsCmCalendar).setDateTime(d)

    AssertEqual e, a.formatAs(f)

    f = LCase(f)
    AssertEqual e, a.formatAs(f)
End Sub
Sub Test_clsCmCalendar_formatAs_D_1
    Dim d,f,e,a
    d = "2024/02/01 00:59:30"
    f = "D"
    e = Cstr(Day(Left(d,19)))
    Set a = (new clsCmCalendar).setDateTime(d)

    AssertEqual e, a.formatAs(f)

    f = LCase(f)
    AssertEqual e, a.formatAs(f)
End Sub
Sub Test_clsCmCalendar_formatAs_D_2
    Dim d,f,e,a
    d = "2024/02/29 00:59:30.123456"
    f = "D"
    e = Cstr(Day(Left(d,19)))
    Set a = (new clsCmCalendar).setDateTime(d)

    AssertEqual e, a.formatAs(f)

    f = LCase(f)
    AssertEqual e, a.formatAs(f)
End Sub
Sub Test_clsCmCalendar_formatAs_HH_1
    Dim d,f,e,a
    d = "2024/02/29 00:59:30"
    f = "HH"
    e = Right("0" & Hour(Left(d,19)), 2)
    Set a = (new clsCmCalendar).setDateTime(d)

    AssertEqual e, a.formatAs(f)

    f = LCase(f)
    AssertEqual e, a.formatAs(f)
End Sub
Sub Test_clsCmCalendar_formatAs_HH_2
    Dim d,f,e,a
    d = "2024/02/29 23:59:30.123456"
    f = "HH"
    e = Right("0" & Hour(Left(d,19)), 2)
    Set a = (new clsCmCalendar).setDateTime(d)

    AssertEqual e, a.formatAs(f)

    f = LCase(f)
    AssertEqual e, a.formatAs(f)
End Sub
Sub Test_clsCmCalendar_formatAs_H_1
    Dim d,f,e,a
    d = "2024/02/29 09:59:30"
    f = "H"
    e = CStr(Hour(Left(d,19)))
    Set a = (new clsCmCalendar).setDateTime(d)

    AssertEqual e, a.formatAs(f)

    f = LCase(f)
    AssertEqual e, a.formatAs(f)
End Sub
Sub Test_clsCmCalendar_formatAs_H_2
    Dim d,f,e,a
    d = "2024/02/29 10:59:30.123456"
    f = "H"
    e = CStr(Hour(Left(d,19)))
    Set a = (new clsCmCalendar).setDateTime(d)

    AssertEqual e, a.formatAs(f)

    f = LCase(f)
    AssertEqual e, a.formatAs(f)
End Sub
Sub Test_clsCmCalendar_formatAs_Minute_mm_1
    Dim d,f,e,a
    d = "2024/02/29 00:00:30"
    f = "mm"
    e = Right("0" & Minute(Left(d,19)), 2)
    Set a = (new clsCmCalendar).setDateTime(d)

    AssertEqual e, a.formatAs(f)
End Sub
Sub Test_clsCmCalendar_formatAs_Minute_mm_2
    Dim d,f,e,a
    d = "2024/02/29 00:59:30.123456"
    f = "mm"
    e = Right("0" & Minute(Left(d,19)), 2)
    Set a = (new clsCmCalendar).setDateTime(d)

    AssertEqual e, a.formatAs(f)
End Sub
Sub Test_clsCmCalendar_formatAs_Minute_m_1
    Dim d,f,e,a
    d = "2024/02/29 00:09:30"
    f = "m"
    e = Cstr(Minute(Left(d,19)))
    Set a = (new clsCmCalendar).setDateTime(d)

    AssertEqual e, a.formatAs(f)
End Sub
Sub Test_clsCmCalendar_formatAs_Minute_m_2
    Dim d,f,e,a
    d = "2024/02/29 00:10:30.123456"
    f = "m"
    e = Cstr(Minute(Left(d,19)))
    Set a = (new clsCmCalendar).setDateTime(d)

    AssertEqual e, a.formatAs(f)
End Sub
Sub Test_clsCmCalendar_formatAs_SS_1
    Dim d,f,e,a
    d = "2024/02/29 00:59:30"
    f = "SS"
    e = Right("0" & Second(Left(d,19)), 2)
    Set a = (new clsCmCalendar).setDateTime(d)

    AssertEqual e, a.formatAs(f)

    f = LCase(f)
    AssertEqual e, a.formatAs(f)
End Sub
Sub Test_clsCmCalendar_formatAs_SS_1
    Dim d,f,e,a
    d = "2024/02/29 00:59:09.123456"
    f = "SS"
    e = Right("0" & Second(Left(d,19)), 2)
    Set a = (new clsCmCalendar).setDateTime(d)

    AssertEqual e, a.formatAs(f)

    f = LCase(f)
    AssertEqual e, a.formatAs(f)
End Sub
Sub Test_clsCmCalendar_formatAs_S_1
    Dim d,f,e,a
    d = "2024/02/29 00:59:00"
    f = "S"
    e = Cstr(Second(Left(d,19)))
    Set a = (new clsCmCalendar).setDateTime(d)

    AssertEqual e, a.formatAs(f)

    f = LCase(f)
    AssertEqual e, a.formatAs(f)
End Sub
Sub Test_clsCmCalendar_formatAs_S_2
    Dim d,f,e,a
    d = "2024/02/29 00:59:10.123456"
    f = "S"
    e = Cstr(Second(Left(d,19)))
    Set a = (new clsCmCalendar).setDateTime(d)

    AssertEqual e, a.formatAs(f)

    f = LCase(f)
    AssertEqual e, a.formatAs(f)
End Sub
Sub Test_clsCmCalendar_formatAs_000000
    Dim d,f,e,a
    f = "." & "000000"

    d = "2024/02/29 00:59:10.1234567"
    e = "." & Left(Mid(d,21) & "000000", 6)
    Set a = (new clsCmCalendar).setDateTime(d)
    AssertEqualWithMessage e, a.formatAs(f), "Input = " & d

    d = "2024/02/29 00:59:10.123456"
    e = "." & Left(Mid(d,21) & "000000", 6)
    Set a = (new clsCmCalendar).setDateTime(d)
    AssertEqualWithMessage e, a.formatAs(f), "Input = " & d

    d = "2024/02/29 00:59:10.12345"
    e = "." & Left(Mid(d,21) & "000000", 6)
    Set a = (new clsCmCalendar).setDateTime(d)
    AssertEqualWithMessage e, a.formatAs(f), "Input = " & d

    d = "2024/02/29 00:59:10.1234"
    e = "." & Left(Mid(d,21) & "000000", 6)
    Set a = (new clsCmCalendar).setDateTime(d)
    AssertEqualWithMessage e, a.formatAs(f), "Input = " & d

    d = "2024/02/29 00:59:10.123"
    e = "." & Left(Mid(d,21) & "000000", 6)
    Set a = (new clsCmCalendar).setDateTime(d)
    AssertEqualWithMessage e, a.formatAs(f), "Input = " & d

    d = "2024/02/29 00:59:10.12"
    e = "." & Left(Mid(d,21) & "000000", 6)
    Set a = (new clsCmCalendar).setDateTime(d)
    AssertEqualWithMessage e, a.formatAs(f), "Input = " & d

    d = "2024/02/29 00:59:10.1"
    e = "." & Left(Mid(d,21) & "000000", 6)
    Set a = (new clsCmCalendar).setDateTime(d)
    AssertEqualWithMessage e, a.formatAs(f), "Input = " & d

    d = "2024/02/29 00:59:10"
    e = "." & "000000"
    Set a = (new clsCmCalendar).setDateTime(d)
    AssertEqualWithMessage e, a.formatAs(f), "Input = " & d
End Sub
Sub Test_clsCmCalendar_formatAs_000
    Dim d,f,e,a
    f = "." & "000"

    d = "2024/02/29 00:59:10.1234"
    e = "." & Left(Mid(d,21) & "000", 3)
    Set a = (new clsCmCalendar).setDateTime(d)
    AssertEqualWithMessage e, a.formatAs(f), "Input = " & d

    d = "2024/02/29 00:59:10.123"
    e = "." & Left(Mid(d,21) & "000", 3)
    Set a = (new clsCmCalendar).setDateTime(d)
    AssertEqualWithMessage e, a.formatAs(f), "Input = " & d

    d = "2024/02/29 00:59:10.12"
    e = "." & Left(Mid(d,21) & "000", 3)
    Set a = (new clsCmCalendar).setDateTime(d)
    AssertEqualWithMessage e, a.formatAs(f), "Input = " & d

    d = "2024/02/29 00:59:10.1"
    e = "." & Left(Mid(d,21) & "000", 3)
    Set a = (new clsCmCalendar).setDateTime(d)
    AssertEqualWithMessage e, a.formatAs(f), "Input = " & d

    d = "2024/02/29 00:59:10"
    e = "." & "000"
    Set a = (new clsCmCalendar).setDateTime(d)
    AssertEqualWithMessage e, a.formatAs(f), "Input = " & d
End Sub


' Local Variables:
' mode: Visual-Basic
' indent-tabs-mode: nil
' End:
