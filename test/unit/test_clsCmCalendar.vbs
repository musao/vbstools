' clsCmCalendar.vbs: test.
' @import ../../lib/clsAdptFile.vbs
' @import ../../lib/clsCmArray.vbs
' @import ../../lib/clsCmBroker.vbs
' @import ../../lib/clsCmBufferedReader.vbs
' @import ../../lib/clsCmBufferedWriter.vbs
' @import ../../lib/clsCmCalendar.vbs
' @import ../../lib/clsCmCharacterType.vbs
' @import ../../lib/clsCmCssGenerator.vbs
' @import ../../lib/clsCmHtmlGenerator.vbs
' @import ../../lib/clsCmReturnValue.vbs
' @import ../../lib/clsCompareExcel.vbs
' @import ../../lib/libCom.vbs

Option Explicit

'###################################################################################################
'clsCmCalendar
Sub Test_clsCmCalendar
    Dim a : Set a = new clsCmCalendar
    AssertEqual 8, VarType(a)
    AssertEqual "clsCmCalendar", TypeName(a)
End Sub

'###################################################################################################
'clsCmCalendar.serial()
Sub Test_clsCmCalendar_serial_ofNow
    Dim a,e
    Set e = (new clsCmCalendar).ofNow()
    Set a = e.clone()

    AssertWithMessage Not e Is a, "Object"
    AssertEqualWithMessage e.serial, a.serial, "serial"
    AssertEqualWithMessage e.toString, a.toString, "toString"
    AssertEqualWithMessage Left(e.formatAs("000000"),4), Left(a.formatAs("000000"),4), "microsecond"
End Sub
Sub Test_clsCmCalendar_serial_setDateTime_1
    Dim a,e
    Set e = (new clsCmCalendar).of("2023/12/31 23:30:10.4567890")
    Set a = e.clone()

    AssertWithMessage Not e Is a, "Object"
    AssertEqualWithMessage e.serial, a.serial, "serial"
    AssertEqualWithMessage e.toString, a.toString, "toString"
    AssertEqualWithMessage Left(e.formatAs("000000"),4), Left(a.formatAs("000000"),4), "microsecond"
End Sub
Sub Test_clsCmCalendar_serial_setDateTime_2
    Dim a,e
    Set e = (new clsCmCalendar).of("2023/12/31 23:30:10.5678901")
    Set a = e.clone()

    AssertWithMessage Not e Is a, "Object"
    AssertEqualWithMessage e.serial, a.serial, "serial"
    AssertEqualWithMessage e.toString, a.toString, "toString"
    AssertEqualWithMessage Left(e.formatAs("000000"),4), Left(a.formatAs("000000"),4), "microsecond"
End Sub

'###################################################################################################
'clsCmCalendar.compareTo()
Sub Test_clsCmCalendar_compareTo
    Dim d1,d2,e,a,a1,a2
    d1 = "2024/02/29 00:59:31"
    Set a1 = (new clsCmCalendar).of(d1)

    d2 = "2024/02/29 00:59:30"
    Set a2 = (new clsCmCalendar).of(d2)
    e = 1
    a = a1.compareTo(a2)
    AssertEqualWithMessage e, a, a1&" vs "&a2

    d2 = "2024/02/29 00:59:31"
    Set a2 = (new clsCmCalendar).of(d2)
    e = 0
    a = a1.compareTo(a2)
    AssertEqualWithMessage e, a, a1&" vs "&a2

    d2 = "2024/02/29 00:59:32"
    Set a2 = (new clsCmCalendar).of(d2)
    e = -1
    a = a1.compareTo(a2)
    AssertEqualWithMessage e, a, a1&" vs "&a2
End Sub
Sub Test_clsCmCalendar_compareTo_WithDecimal
    Dim d1,d2,e,a,a1,a2
    d1 = "2024/02/29 00:59:31.123456"
    Set a1 = (new clsCmCalendar).of(d1)

    d2 = "2024/02/29 00:59:31.123455"
    Set a2 = (new clsCmCalendar).of(d2)
    e = 1
    a = a1.compareTo(a2)
    AssertEqualWithMessage e, a, a1.formatAs("ss.000000")&" vs "&a2.formatAs("ss.000000")

    d2 = "2024/02/29 00:59:31.123456"
    Set a2 = (new clsCmCalendar).of(d2)
    e = 0
    a = a1.compareTo(a2)
    AssertEqualWithMessage e, a, a1.formatAs("ss.000000")&" vs "&a2.formatAs("ss.000000")

    d2 = "2024/02/29 00:59:31.123457"
    Set a2 = (new clsCmCalendar).of(d2)
    e = -1
    a = a1.compareTo(a2)
    AssertEqualWithMessage e, a, a1.formatAs("ss.000000")&" vs "&a2.formatAs("ss.000000")
End Sub
Sub Test_clsCmCalendar_compareTo_DateOnly
    Dim d1,d2,e,a,a1,a2
    d1 = "2024/02/29"
    Set a1 = (new clsCmCalendar).of(d1)

    d2 = "2024/02/28"
    Set a2 = (new clsCmCalendar).of(d2)
    e = 1
    a = a1.compareTo(a2)
    AssertEqualWithMessage e, a, a1&" vs "&a2

    d2 = "2024/02/29"
    Set a2 = (new clsCmCalendar).of(d2)
    e = 0
    a = a1.compareTo(a2)
    AssertEqualWithMessage e, a, a1&" vs "&a2

    d2 = "2024/03/01"
    Set a2 = (new clsCmCalendar).of(d2)
    e = -1
    a = a1.compareTo(a2)
    AssertEqualWithMessage e, a, a1&" vs "&a2
End Sub
Sub Test_clsCmCalendar_compareTo_TimeOnly
    Dim d1,d2,e,a,a1,a2
    d1 = "00:59:31"
    Set a1 = (new clsCmCalendar).of(d1)

    d2 = "00:59:30"
    Set a2 = (new clsCmCalendar).of(d2)
    e = 1
    a = a1.compareTo(a2)
    AssertEqualWithMessage e, a, a1&" vs "&a2

    d2 = "00:59:31"
    Set a2 = (new clsCmCalendar).of(d2)
    e = 0
    a = a1.compareTo(a2)
    AssertEqualWithMessage e, a, a1&" vs "&a2

    d2 = "00:59:32"
    Set a2 = (new clsCmCalendar).of(d2)
    e = -1
    a = a1.compareTo(a2)
    AssertEqualWithMessage e, a, a1&" vs "&a2
End Sub
Sub Test_clsCmCalendar_compareTo_Err
    On Error Resume Next
    Dim d1,e,a,a1
    d1 = "2024/02/29 00:59:31"
    Set a1 = (new clsCmCalendar).of(d1)

    e = Empty
    a = a1.compareTo(new_Dic())

    AssertEqualWithMessage e, a, "ret"
    AssertEqualWithMessage "clsCmCalendar+compareTo()", Err.Source, "Err.Source"
    AssertEqualWithMessage "That object is not a calendar class.", Err.Description, "Err.Description"
End Sub

'###################################################################################################
'clsCmCalendar.differenceFrom()
Sub Test_clsCmCalendar_differenceFrom
    Dim d1,d2,e,a,a1,a2
    d1 = "2024/02/29 00:59:31"
    Set a1 = (new clsCmCalendar).of(d1)

    d2 = "2024/02/29 00:59:30"
    Set a2 = (new clsCmCalendar).of(d2)
    e = 1
    a = a1.differenceFrom(a2)
    AssertEqualWithMessage e, a, a1&" vs "&a2

    d2 = "2024/02/29 00:59:31"
    Set a2 = (new clsCmCalendar).of(d2)
    e = 0
    a = a1.differenceFrom(a2)
    AssertEqualWithMessage e, a, a1&" vs "&a2

    d2 = "2024/02/29 00:59:32"
    Set a2 = (new clsCmCalendar).of(d2)
    e = -1
    a = a1.differenceFrom(a2)
    AssertEqualWithMessage e, a, a1&" vs "&a2
End Sub
Sub Test_clsCmCalendar_differenceFrom_WithDecimal
    Dim d1,d2,e,a,a1,a2
    d1 = "2024/02/29 00:59:31.123456"
    Set a1 = (new clsCmCalendar).of(d1)

    d2 = "2024/02/29 00:59:30.123455"
    Set a2 = (new clsCmCalendar).of(d2)
    e = 1
    a = a1.differenceFrom(a2)
    AssertEqualWithMessage e, a, a1.formatAs("ss.000000")&" vs "&a2.formatAs("ss.000000")

    d2 = "2024/02/29 00:59:31.123456"
    Set a2 = (new clsCmCalendar).of(d2)
    e = 0
    a = a1.differenceFrom(a2)
    AssertEqualWithMessage e, a, a1.formatAs("ss.000000")&" vs "&a2.formatAs("ss.000000")

    d2 = "2024/02/29 00:59:32.123457"
    Set a2 = (new clsCmCalendar).of(d2)
    e = -1
    a = a1.differenceFrom(a2)
    AssertEqualWithMessage e, a, a1.formatAs("ss.000000")&" vs "&a2.formatAs("ss.000000")
End Sub
Sub Test_clsCmCalendar_differenceFrom_Err
    On Error Resume Next
    Dim d1,e,a,a1
    d1 = "2024/02/29 00:59:31"
    Set a1 = (new clsCmCalendar).of(d1)

    e = Empty
    a = a1.differenceFrom(new_Dic())

    AssertEqualWithMessage e, a, "ret"
    AssertEqualWithMessage "clsCmCalendar+differenceFrom()", Err.Source, "Err.Source"
    AssertEqualWithMessage "That object is not a calendar class.", Err.Description, "Err.Description"
End Sub

'###################################################################################################
'clsCmCalendar.formatAs()
Sub Test_clsCmCalendar_formatAs_Normal1
    Dim d,f,e,a
    f = "YY/M/d hh:mm:ss.000"
    d = "2024/02/29 00:59:30.456789"
    e = "24/2/29 00:59:30.456"
    Set a = (new clsCmCalendar).of(d)

    AssertEqual e, a.formatAs(f)
End Sub
Sub Test_clsCmCalendar_formatAs_Normal2
    Dim d,f,e,a
    f = "fujii.txt"
    d = "2024/02/29 00:59:30.456789"
    e = f
    Set a = (new clsCmCalendar).of(d)

    AssertEqual e, a.formatAs(f)
End Sub
Sub Test_clsCmCalendar_formatAs_YYYY
    Dim d,f,e,a
    f = "YYYY"
    d = "2024/02/29 00:59:30"
    e = Right("000" & Year(Left(d,19)), 4)
    Set a = (new clsCmCalendar).of(d)

    AssertEqualWithMessage e, a.formatAs(f), "uppercase"

    f = LCase(f)
    AssertEqualWithMessage e, a.formatAs(f), "lowercase"
End Sub
Sub Test_clsCmCalendar_formatAs_YY
    Dim d,f,e,a
    f = "YY"
    d = "2024/02/29 00:59:30.123456"
    e = Right("0" & Year(Left(d,19)), 2)
    Set a = (new clsCmCalendar).of(d)

    AssertEqualWithMessage e, a.formatAs(f), "uppercase"

    f = LCase(f)
    AssertEqualWithMessage e, a.formatAs(f), "lowercase"
End Sub
Sub Test_clsCmCalendar_formatAs_Month_MM_1
    Dim d,f,e,a
    f = "MM"
    d = "2024/02/29 00:59:30"
    e = Right("0" & Month(Left(d,19)), 2)
    Set a = (new clsCmCalendar).of(d)

    AssertEqual e, a.formatAs(f)
End Sub
Sub Test_clsCmCalendar_formatAs_Month_MM_2
    Dim d,f,e,a
    f = "MM"
    d = "2024/10/29 00:59:30.123456"
    e = Right("0" & Month(Left(d,19)), 2)
    Set a = (new clsCmCalendar).of(d)

    AssertEqual e, a.formatAs(f)
End Sub
Sub Test_clsCmCalendar_formatAs_Month_M_1
    Dim d,f,e,a
    f = "M"
    d = "2024/02/29 00:59:30"
    e = Cstr(Month(Left(d,19)))
    Set a = (new clsCmCalendar).of(d)

    AssertEqual e, a.formatAs(f)
End Sub
Sub Test_clsCmCalendar_formatAs_Month_M_2
    Dim d,f,e,a
    f = "M"
    d = "2024/11/29 00:59:30.123456"
    e = Cstr(Month(Left(d,19)))
    Set a = (new clsCmCalendar).of(d)

    AssertEqual e, a.formatAs(f)
End Sub
Sub Test_clsCmCalendar_formatAs_DD_1
    Dim d,f,e,a
    f = "DD"
    d = "2024/02/09 00:59:30"
    e = Right("0" & Day(Left(d,19)), 2)
    Set a = (new clsCmCalendar).of(d)

    AssertEqualWithMessage e, a.formatAs(f), "uppercase"

    f = LCase(f)
    AssertEqualWithMessage e, a.formatAs(f), "lowercase"
End Sub
Sub Test_clsCmCalendar_formatAs_DD_2
    Dim d,f,e,a
    f = "DD"
    d = "2024/02/29 00:59:30.123456"
    e = Right("0" & Day(Left(d,19)), 2)
    Set a = (new clsCmCalendar).of(d)

    AssertEqualWithMessage e, a.formatAs(f), "uppercase"

    f = LCase(f)
    AssertEqualWithMessage e, a.formatAs(f), "lowercase"
End Sub
Sub Test_clsCmCalendar_formatAs_D_1
    Dim d,f,e,a
    f = "D"
    d = "2024/02/01 00:59:30"
    e = Cstr(Day(Left(d,19)))
    Set a = (new clsCmCalendar).of(d)

    AssertEqualWithMessage e, a.formatAs(f), "uppercase"

    f = LCase(f)
    AssertEqualWithMessage e, a.formatAs(f), "lowercase"
End Sub
Sub Test_clsCmCalendar_formatAs_D_2
    Dim d,f,e,a
    f = "D"
    d = "2024/02/29 00:59:30.123456"
    e = Cstr(Day(Left(d,19)))
    Set a = (new clsCmCalendar).of(d)

    AssertEqualWithMessage e, a.formatAs(f), "uppercase"

    f = LCase(f)
    AssertEqualWithMessage e, a.formatAs(f), "lowercase"
End Sub
Sub Test_clsCmCalendar_formatAs_HH_1
    Dim d,f,e,a
    f = "HH"
    d = "2024/02/29 00:59:30"
    e = Right("0" & Hour(Left(d,19)), 2)
    Set a = (new clsCmCalendar).of(d)

    AssertEqualWithMessage e, a.formatAs(f), "uppercase"

    f = LCase(f)
    AssertEqualWithMessage e, a.formatAs(f), "lowercase"
End Sub
Sub Test_clsCmCalendar_formatAs_HH_2
    Dim d,f,e,a
    f = "HH"
    d = "2024/02/29 23:59:30.123456"
    e = Right("0" & Hour(Left(d,19)), 2)
    Set a = (new clsCmCalendar).of(d)

    AssertEqualWithMessage e, a.formatAs(f), "uppercase"

    f = LCase(f)
    AssertEqualWithMessage e, a.formatAs(f), "lowercase"
End Sub
Sub Test_clsCmCalendar_formatAs_H_1
    Dim d,f,e,a
    f = "H"
    d = "2024/02/29 09:59:30"
    e = CStr(Hour(Left(d,19)))
    Set a = (new clsCmCalendar).of(d)

    AssertEqualWithMessage e, a.formatAs(f), "uppercase"

    f = LCase(f)
    AssertEqualWithMessage e, a.formatAs(f), "lowercase"
End Sub
Sub Test_clsCmCalendar_formatAs_H_2
    Dim d,f,e,a
    f = "H"
    d = "2024/02/29 10:59:30.123456"
    e = CStr(Hour(Left(d,19)))
    Set a = (new clsCmCalendar).of(d)

    AssertEqualWithMessage e, a.formatAs(f), "uppercase"

    f = LCase(f)
    AssertEqualWithMessage e, a.formatAs(f), "lowercase"
End Sub
Sub Test_clsCmCalendar_formatAs_Minute_mm_1
    Dim d,f,e,a
    f = "mm"
    d = "2024/02/29 00:00:30"
    e = Right("0" & Minute(Left(d,19)), 2)
    Set a = (new clsCmCalendar).of(d)

    AssertEqual e, a.formatAs(f)
End Sub
Sub Test_clsCmCalendar_formatAs_Minute_mm_2
    Dim d,f,e,a
    f = "mm"
    d = "2024/02/29 00:59:30.123456"
    e = Right("0" & Minute(Left(d,19)), 2)
    Set a = (new clsCmCalendar).of(d)

    AssertEqual e, a.formatAs(f)
End Sub
Sub Test_clsCmCalendar_formatAs_Minute_m_1
    Dim d,f,e,a
    f = "m"
    d = "2024/02/29 00:09:30"
    e = Cstr(Minute(Left(d,19)))
    Set a = (new clsCmCalendar).of(d)

    AssertEqual e, a.formatAs(f)
End Sub
Sub Test_clsCmCalendar_formatAs_Minute_m_2
    Dim d,f,e,a
    f = "m"
    d = "2024/02/29 00:10:30.123456"
    e = Cstr(Minute(Left(d,19)))
    Set a = (new clsCmCalendar).of(d)

    AssertEqual e, a.formatAs(f)
End Sub
Sub Test_clsCmCalendar_formatAs_SS_1
    Dim d,f,e,a
    f = "SS"
    d = "2024/02/29 00:59:30"
    e = Right("0" & Second(Left(d,19)), 2)
    Set a = (new clsCmCalendar).of(d)

    AssertEqualWithMessage e, a.formatAs(f), "uppercase"

    f = LCase(f)
    AssertEqualWithMessage e, a.formatAs(f), "lowercase"
End Sub
Sub Test_clsCmCalendar_formatAs_SS_1
    Dim d,f,e,a
    f = "SS"
    d = "2024/02/29 00:59:09.123456"
    e = Right("0" & Second(Left(d,19)), 2)
    Set a = (new clsCmCalendar).of(d)

    AssertEqualWithMessage e, a.formatAs(f), "uppercase"

    f = LCase(f)
    AssertEqualWithMessage e, a.formatAs(f), "lowercase"
End Sub
Sub Test_clsCmCalendar_formatAs_S_1
    Dim d,f,e,a
    f = "S"
    d = "2024/02/29 00:59:00"
    e = Cstr(Second(Left(d,19)))
    Set a = (new clsCmCalendar).of(d)

    AssertEqualWithMessage e, a.formatAs(f), "uppercase"

    f = LCase(f)
    AssertEqualWithMessage e, a.formatAs(f), "lowercase"
End Sub
Sub Test_clsCmCalendar_formatAs_S_2
    Dim d,f,e,a
    f = "S"
    d = "2024/02/29 00:59:10.123456"
    e = Cstr(Second(Left(d,19)))
    Set a = (new clsCmCalendar).of(d)

    AssertEqualWithMessage e, a.formatAs(f), "uppercase"

    f = LCase(f)
    AssertEqualWithMessage e, a.formatAs(f), "lowercase"
End Sub
Sub Test_clsCmCalendar_formatAs_000000
    Dim d,f,e,a
    f = "000000"

    d = "2024/02/29 00:59:10.1234567"
    e = Left(Mid(d,21) & "000000", 6)
    Set a = (new clsCmCalendar).of(d)
    AssertEqualWithMessage e, a.formatAs(f), "Input = " & d

    d = "2024/02/29 00:59:10.123456"
    e = Left(Mid(d,21) & "000000", 6)
    Set a = (new clsCmCalendar).of(d)
    AssertEqualWithMessage e, a.formatAs(f), "Input = " & d

    d = "2024/02/29 00:59:10.12345"
    e = Left(Mid(d,21) & "000000", 6)
    Set a = (new clsCmCalendar).of(d)
    AssertEqualWithMessage e, a.formatAs(f), "Input = " & d

    d = "2024/02/29 00:59:10.1234"
    e = Left(Mid(d,21) & "000000", 6)
    Set a = (new clsCmCalendar).of(d)
    AssertEqualWithMessage e, a.formatAs(f), "Input = " & d

    d = "2024/02/29 00:59:10.123"
    e = Left(Mid(d,21) & "000000", 6)
    Set a = (new clsCmCalendar).of(d)
    AssertEqualWithMessage e, a.formatAs(f), "Input = " & d

    d = "2024/02/29 00:59:10.12"
    e = Left(Mid(d,21) & "000000", 6)
    Set a = (new clsCmCalendar).of(d)
    AssertEqualWithMessage e, a.formatAs(f), "Input = " & d

    d = "2024/02/29 00:59:10.1"
    e = Left(Mid(d,21) & "000000", 6)
    Set a = (new clsCmCalendar).of(d)
    AssertEqualWithMessage e, a.formatAs(f), "Input = " & d

    d = "2024/02/29 00:59:10"
    e = "000000"
    Set a = (new clsCmCalendar).of(d)
    AssertEqualWithMessage e, a.formatAs(f), "Input = " & d
End Sub
Sub Test_clsCmCalendar_formatAs_000
    Dim d,f,e,a
    f = "000"

    d = "2024/02/29 00:59:10.1234"
    e = Left(Mid(d,21) & "000", 3)
    Set a = (new clsCmCalendar).of(d)
    AssertEqualWithMessage e, a.formatAs(f), "Input = " & d

    d = "2024/02/29 00:59:10.123"
    e = Left(Mid(d,21) & "000", 3)
    Set a = (new clsCmCalendar).of(d)
    AssertEqualWithMessage e, a.formatAs(f), "Input = " & d

    d = "2024/02/29 00:59:10.12"
    e = Left(Mid(d,21) & "000", 3)
    Set a = (new clsCmCalendar).of(d)
    AssertEqualWithMessage e, a.formatAs(f), "Input = " & d

    d = "2024/02/29 00:59:10.1"
    e = Left(Mid(d,21) & "000", 3)
    Set a = (new clsCmCalendar).of(d)
    AssertEqualWithMessage e, a.formatAs(f), "Input = " & d

    d = "2024/02/29 00:59:10"
    e = "000"
    Set a = (new clsCmCalendar).of(d)
    AssertEqualWithMessage e, a.formatAs(f), "Input = " & d
End Sub

'###################################################################################################
'clsCmCalendar.ofNow()/toString()
Sub Test_clsCmCalendar_ofNow_toString
    Dim n,y,m,d,h,mm,s
    n = now()
    y = Right("000" & Year(n), 4)
    m = Right("0" & Month(n), 2)
    d = Right("0" & Day(n), 2)
    h = Right("0" & Hour(n), 2)
    mm = Right("0" & Minute(n), 2)
    s = Right("0" & Second(n), 2)
    Dim ptn : ptn = "^"&y&"/"&m&"/"&d&" "&h&":"&mm&":"&s&"\.\d{3}$"
    Dim a : Set a = (new clsCmCalendar).ofNow()

    AssertMatchWithMessage ptn, a.toString(), "toString()"
    AssertMatchWithMessage ptn, a, "default"
End Sub

'###################################################################################################
'clsCmCalendar.of()/toString()
Sub Test_clsCmCalendar_setDateTime_toString
    Dim d : d = "2024/02/29 00:59:30"
    Dim e : e = d & ".000"
    Dim a : Set a = (new clsCmCalendar).of(d)

    AssertEqual e, a.toString()
End Sub
Sub Test_clsCmCalendar_setDateTime_toString_ErrAtDate
    On Error Resume Next
    Dim d : d = "2022/02/29 00:59:30"
    Dim e : e = Empty
    Dim a : Set a = (new clsCmCalendar).of(d)

    AssertEqualWithMessage e, a, "ret"
    AssertEqualWithMessage "clsCmCalendar+of()", Err.Source, "Err.Source"
    AssertMatchWithMessage "^invalid argument.*"&d, Err.Description, "Err.Description"
End Sub
Sub Test_clsCmCalendar_setDateTime_toString_ErrAtTime
    On Error Resume Next
    Dim d : d = "2024/02/29 00:ab:30"
    Dim e : e = Empty
    Dim a : Set a = (new clsCmCalendar).of(d)

    AssertEqualWithMessage e, a, "ret"
    AssertEqualWithMessage "clsCmCalendar+of()", Err.Source, "Err.Source"
    AssertMatchWithMessage "^invalid argument.*"&d, Err.Description, "Err.Description"
End Sub
Sub Test_clsCmCalendar_setDateTime_toString_WithDecimal
    Dim d : d = "2023/12/31 23:30:59.123456"
    Dim e : e = Left(d,23)
    Dim a : Set a = (new clsCmCalendar).of(d)

    AssertEqual e, a.toString()
End Sub
Sub Test_clsCmCalendar_setDateTime_toString_WithDecimal_ErrAtDate
    On Error Resume Next
    Dim d : d = "2023/13/31 00:59:30.123456"
    Dim e : e = Empty
    Dim a : Set a = (new clsCmCalendar).of(d)

    AssertEqualWithMessage e, a, "ret"
    AssertEqualWithMessage "clsCmCalendar+of()", Err.Source, "Err.Source"
    AssertMatchWithMessage "^invalid argument.*"&d, Err.Description, "Err.Description"
End Sub
Sub Test_clsCmCalendar_setDateTime_toString_WithDecimal_ErrAtTime
    On Error Resume Next
    Dim d : d = "2023/12/31 00:59:60.123456"
    Dim e : e = Empty
    Dim a : Set a = (new clsCmCalendar).of(d)

    AssertEqualWithMessage e, a, "ret"
    AssertEqualWithMessage "clsCmCalendar+of()", Err.Source, "Err.Source"
    AssertMatchWithMessage "^invalid argument.*"&d, Err.Description, "Err.Description"
End Sub
Sub Test_clsCmCalendar_setDateTime_toString_WithDecimal_ErrAtFractionalSec
    On Error Resume Next
    Dim d : d = "2023/12/31 00:59:30.12a456"
    Dim e : e = Empty
    Dim a : Set a = (new clsCmCalendar).of(d)

    AssertEqualWithMessage e, a, "ret"
    AssertEqualWithMessage "clsCmCalendar+of()", Err.Source, "Err.Source"
    AssertMatchWithMessage "^invalid argument.*"&d, Err.Description, "Err.Description"
End Sub
Sub Test_clsCmCalendar_setDateTime_toString_DateOnly
    Dim d : d = "2022/01/01"
    Dim e : e = d & " 00:00:00.000"
    Dim a : Set a = (new clsCmCalendar).of(e)

    AssertEqual e, a.toString()
End Sub
Sub Test_clsCmCalendar_setDateTime_toString_DateOnly_Err
    On Error Resume Next
    Dim d : d = "2022/00/01"
    Dim e : e = Empty
    Dim a : Set a = (new clsCmCalendar).of(d)

    AssertEqualWithMessage e, a, "ret"
    AssertEqualWithMessage "clsCmCalendar+of()", Err.Source, "Err.Source"
    AssertMatchWithMessage "^invalid argument.*"&d, Err.Description, "Err.Description"
End Sub
Sub Test_clsCmCalendar_setDateTime_toString_TimeOnly
    Dim d : d = "12:59:30"
    Dim e : e = "1900/01/01 " & d & ".000"
    Dim a : Set a = (new clsCmCalendar).of(e)

    AssertEqual e, a.toString()
End Sub
Sub Test_clsCmCalendar_setDateTime_toString_TimeOnly_Err
    On Error Resume Next
    Dim d : d = "12:60:30"
    Dim e : e = Empty
    Dim a : Set a = (new clsCmCalendar).of(d)

    AssertEqualWithMessage e, a, "ret"
    AssertEqualWithMessage "clsCmCalendar+of()", Err.Source, "Err.Source"
    AssertMatchWithMessage "^invalid argument.*"&d, Err.Description, "Err.Description"
End Sub
Sub Test_clsCmCalendar_setDateTime_toString_TimeOnly_WithDecimal
    Dim d : d = "12:59:30.4567"
    Dim e : e = "1900/01/01 " & Left(d, 12)
    Dim a : Set a = (new clsCmCalendar).of(e)

    AssertEqual e, a.toString()
End Sub
Sub Test_clsCmCalendar_setDateTime_toString_TimeOnly_WithDecimal_ErrAtTime
    On Error Resume Next
    Dim d : d = "12:60:30.4567"
    Dim e : e = Empty
    Dim a : Set a = (new clsCmCalendar).of(d)

    AssertEqualWithMessage e, a, "ret"
    AssertEqualWithMessage "clsCmCalendar+of()", Err.Source, "Err.Source"
    AssertMatchWithMessage "^invalid argument.*"&d, Err.Description, "Err.Description"
End Sub
Sub Test_clsCmCalendar_setDateTime_toString_TimeOnly_WithDecimal_ErrAtFractionalSec
    On Error Resume Next
    Dim d : d = "12:59:30.b567"
    Dim e : e = Empty
    Dim a : Set a = (new clsCmCalendar).of(d)

    AssertEqualWithMessage e, a, "ret"
    AssertEqualWithMessage "clsCmCalendar+of()", Err.Source, "Err.Source"
    AssertMatchWithMessage "^invalid argument.*"&d, Err.Description, "Err.Description"
End Sub


' Local Variables:
' mode: Visual-Basic
' indent-tabs-mode: nil
' End:
