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
Sub Test_clsCmArray_getNow_toString
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
Sub Test_clsCmArray_setDateTime_toString
    Dim e : e = "2024/02/29 00:59:30"
    Dim a : Set a = (new clsCmCalendar).setDateTime(e)

    AssertMatch e & ".000", a.toString()
    AssertMatch e & ".000", a
End Sub
Sub Test_clsCmArray_setDateTime_WithDecimal_toString
    Dim e : e = "2023/12/31 23:30:59.123456"
    Dim a : Set a = (new clsCmCalendar).setDateTime(e)

    AssertMatch mid(e,1,Len(a.toString())), a.toString()
End Sub
Sub Test_clsCmArray_setDateTime_toString_Err
    On Error Resume Next
    Dim e : e = "2022/02/29 00:59:30"
    Dim a : Set a = (new clsCmCalendar).setDateTime(e)

    AssertEqual 13, Err.Number
    AssertEqual "å^Ç™àÍívÇµÇ‹ÇπÇÒÅB", Err.Description
    AssertEqual Empty, a
End Sub


' Local Variables:
' mode: Visual-Basic
' indent-tabs-mode: nil
' End:
