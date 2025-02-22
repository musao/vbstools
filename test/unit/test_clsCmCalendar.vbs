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
'clsCmCalendar.dateTime,fractionalPartOfelapsedSeconds,elapsedSeconds,serial
Sub Test_clsCmCalendar_dateTime_fractionalPartOfelapsedSeconds_elapsedSeconds_serial_initial
    dim tg,a,ao,e
    set ao = (new clsCmCalendar)

    tg = "A.dateTime"
    e = Null
    a = ao.dateTime
    AssertEqualWithMessage e, a, tg

    tg = "B.fractionalPartOfelapsedSeconds"
    e = Null
    a = ao.fractionalPartOfelapsedSeconds
    AssertEqualWithMessage e, a, tg

    tg = "C.elapsedSeconds"
    e = Null
    a = ao.elapsedSeconds
    AssertEqualWithMessage e, a, tg

    tg = "D.serial"
    e = Null
    a = ao.serial
    AssertEqualWithMessage e, a, tg
End Sub
Sub Test_clsCmCalendar_dateTime_fractionalPartOfelapsedSeconds_elapsedSeconds_serial_elapsedSeconds_Null
    dim tg,a,ao,e,d,i,data
    d = Array ( _
            new_DicOf(Array(  "No",1 ,"date", Now()               )) _
            , new_DicOf(Array("No",2 ,"date", Date()              )) _
            , new_DicOf(Array("No",3 ,"date", Time()              )) _
            , new_DicOf(Array("No",4 ,"date", "2025/2/12 11:22:33")) _
            , new_DicOf(Array("No",5 ,"date", "2025/12/31"        )) _
            , new_DicOf(Array("No",6 ,"date", "12:34:56"          )) _
            )

    For Each i In d
        data = i.Item("date")
        set ao = (new clsCmCalendar).of(data)
        
        tg = "A.dateTime"
        e = CDate(data)
        a = ao.dateTime
        AssertEqualWithMessage e, a, tg&" No="&i.Item("No")&", data="&i.Item("date")
        
        tg = "B.fractionalPartOfelapsedSeconds"
        e = 0
        a = ao.fractionalPartOfelapsedSeconds
        AssertEqualWithMessage e, a, tg&" No="&i.Item("No")&", data="&i.Item("date")
        
        tg = "C.elapsedSeconds"
        e = Null
        a = ao.elapsedSeconds
        AssertEqualWithMessage e, a, tg&" No="&i.Item("No")&", data="&i.Item("date")
        
        tg = "D.serial"
        e = Cdbl(CDate(data))
        a = ao.serial
        AssertEqualWithMessage e, a, tg&" No="&i.Item("No")&", data="&i.Item("date")
    Next
End Sub
Sub Test_clsCmCalendar_dateTime_fractionalPartOfelapsedSeconds_elapsedSeconds_serial_elapsedSeconds_NotNull
    dim tg,a,ao,e,d,i,data
    d = Array ( _
            new_DicOf(Array(  "No",1 ,"date", Now()               , "elapsed", Timer()                 )) _
            , new_DicOf(Array("No",2 ,"date", Date()              , "elapsed", "Cal"                   )) _
            , new_DicOf(Array("No",3 ,"date", Time()              , "elapsed", "Cal"                   )) _
            , new_DicOf(Array("No",4 ,"date", "2025/2/12 11:22:33", "elapsed", 11*60*60+22*60+33+0.2345)) _
            , new_DicOf(Array("No",5 ,"date", "2025/12/31"        , "elapsed", 0.8901234               )) _
            , new_DicOf(Array("No",6 ,"date", "12:34:56"          , "elapsed", 0                       )) _
            )

    For Each i In d
        data = Array(i.Item("date"), i.Item("elapsed"))
        If data(1)="Cal" Then data(1)=(Cdbl(Cdate(data(0)))-Fix(Cdbl(Cdate(data(0)))))*24*60*60
        set ao = (new clsCmCalendar).of(data)
        
        tg = "A.dateTime"
        e = CDate(data(0))
        a = ao.dateTime
        AssertEqualWithMessage e, a, tg&" No="&i.Item("No")&", data="&cf_toString(data)
        
        tg = "B.fractionalPartOfelapsedSeconds"
        e = data(1)-Fix(data(1))
        a = ao.fractionalPartOfelapsedSeconds
        AssertWithMessage Abs(e-a)<0.0000001 Or (1-Abs(e-a))<0.0000001, tg&" No="&i.Item("No")&", data="&cf_toString(data)&", e="&cf_toString(e)&", a="&cf_toString(a)&", (e-a)="&cf_toString(e-a)
        
        tg = "C.elapsedSeconds"
        e = data(1)
        a = ao.elapsedSeconds
        AssertEqualWithMessage e, a, tg&" No="&i.Item("No")&", data="&cf_toString(data)
        
        tg = "D.serial"
        e = Cdbl(CDate(data(0)))
        a = ao.serial
        AssertEqualWithMessage e, a, tg&" No="&i.Item("No")&", data="&cf_toString(data)
    Next
End Sub

'###################################################################################################
'clsCmCalendar.toString
Sub Test_clsCmCalendar_toString_initial
    dim a,ao,e
    set ao = (new clsCmCalendar)

    e = "<clsCmCalendar><Null>"
    a = ao.toString()
    AssertEqualWithMessage e, a, "toString()"
End Sub
Sub Test_clsCmCalendar_toString
    dim a,e,d,i,data
    d = Array ( _
            new_DicOf(Array(  "No", 1,"data", Array("2025/2/12 11:22:33")        , "expect", "2025/02/12 11:22:33.000"))_
            , new_DicOf(Array("No", 2,"data", Array("2025/12/1")                 , "expect", "2025/12/01 00:00:00.000"))_
            , new_DicOf(Array("No", 3,"data", Array("12:34:56")                  , "expect", "1899/12/30 12:34:56.000"))_
            , new_DicOf(Array("No", 4,"data", Array("2025/2/12 11:22:33", 0.1234), "expect", "2025/02/12 11:22:33.123"))_
            , new_DicOf(Array("No", 5,"data", Array("2025/12/1"         , 0.9876), "expect", "2025/12/01 00:00:00.987"))_
            , new_DicOf(Array("No", 6,"data", Array("12:34:56"          , 0)     , "expect", "1899/12/30 12:34:56.000"))_
            )

    For Each i In d
        data = i.Item("data")
        e = i.Item("expect")
        a = (new clsCmCalendar).of(data).toString()
        AssertEqualWithMessage e, a, "No="&i.Item("No")&", data="&cf_toString(data)
    Next
End Sub

'###################################################################################################
'clsCmCalendar.clone()
Sub Test_clsCmCalendar_clone_initial
    dim a,ao,e,bo
    set ao = (new clsCmCalendar)
    set bo = ao.clone()

    e = 0
    a = ao.compareTo(bo)
    AssertEqualWithMessage e, a, "clone()"
End Sub
Sub Test_clsCmCalendar_clone
    dim a,e,d,i,data,ao,bo
    d = Array ( _
            new_DicOf(Array(  "No", 1,"data", Array("2025/2/12 11:22:33")        ))_
            , new_DicOf(Array("No", 2,"data", Array("2025/12/1")                 ))_
            , new_DicOf(Array("No", 3,"data", Array("12:34:56")                  ))_
            , new_DicOf(Array("No", 4,"data", Array("2025/2/12 11:22:33", 0.1234)))_
            , new_DicOf(Array("No", 5,"data", Array("2025/12/1"         , 0.9876)))_
            , new_DicOf(Array("No", 6,"data", Array("12:34:56"          , 0)     ))_
            )

    For Each i In d
        data = i.Item("data")
        Set ao = (new clsCmCalendar).of(data)
        Set bo = ao.clone()
        e = 0
        a = ao.compareTo(bo)
        AssertEqualWithMessage e, a, "No="&i.Item("No")&", data="&cf_toString(data)
    Next
End Sub

'###################################################################################################
'clsCmCalendar.compareTo()
Sub Test_clsCmCalendar_compareTo_initial
    dim a,ao,e,bo
    set ao = (new clsCmCalendar)

    set bo = (new clsCmCalendar)
    e = 0
    a = ao.compareTo(bo)
    AssertEqualWithMessage e, a, "compareTo()=0 ao="&ao.toString()&" bo="&bo.toString()

    set bo = (new clsCmCalendar).ofNow()
    e = -1
    a = ao.compareTo(bo)
    AssertEqualWithMessage e, a, "compareTo()<0 ao="&ao.toString()&" bo="&bo.toString()

    ao.ofNow()
    set bo = (new clsCmCalendar)
    e = 1
    a = ao.compareTo(bo)
    AssertEqualWithMessage e, a, "compareTo()>0 ao="&ao.toString()&" bo="&bo.toString()
End Sub
Sub Test_clsCmCalendar_compareTo
    dim a,e,d,i,data,ao,bo
    d = Array ( _
            new_DicOf(Array(  "No", 1,"data", "2024/02/29 00:59:30", "expect", 1 ))_
            , new_DicOf(Array("No", 2,"data", "2024/02/29 00:59:31", "expect", 0 ))_
            , new_DicOf(Array("No", 2,"data", "2024/02/29 00:59:32", "expect", -1))_
            )
    Set ao = (new clsCmCalendar).of("2024/02/29 00:59:31")

    For Each i In d
        data = i.Item("data")
        Set bo = (new clsCmCalendar).of(data)
        e = i.Item("expect")
        a = ao.compareTo(bo)
        AssertEqualWithMessage e, a, "No="&i.Item("No")&" ao="&ao.formatAs("YYYY/MM/DD hh:mm:ss.000000")&" data="&cf_toString(i)
    Next
End Sub
Sub Test_clsCmCalendar_compareTo_WithDecimal
    dim a,e,d,i,data,ao,bo
    d = Array ( _
            new_DicOf(Array(  "No", 1,"data", "2024/02/29 00:59:31.123455", "expect", 1 ))_
            , new_DicOf(Array("No", 2,"data", "2024/02/29 00:59:31.123456", "expect", 0 ))_
            , new_DicOf(Array("No", 2,"data", "2024/02/29 00:59:31.123457", "expect", -1))_
            )
    Set ao = (new clsCmCalendar).of("2024/02/29 00:59:31.123456")

    For Each i In d
        data = i.Item("data")
        Set bo = (new clsCmCalendar).of(data)
        e = i.Item("expect")
        a = ao.compareTo(bo)
        AssertEqualWithMessage e, a, "No="&i.Item("No")&" ao="&ao.formatAs("YYYY/MM/DD hh:mm:ss.000000")&" data="&cf_toString(i)
    Next
End Sub
Sub Test_clsCmCalendar_compareTo_DateOnly
    dim a,e,d,i,data,ao,bo
    d = Array ( _
            new_DicOf(Array(  "No", 1,"data", "2024/02/28", "expect", 1 ))_
            , new_DicOf(Array("No", 2,"data", "2024/02/29", "expect", 0 ))_
            , new_DicOf(Array("No", 2,"data", "2024/03/01", "expect", -1))_
            )
    Set ao = (new clsCmCalendar).of("2024/02/29")

    For Each i In d
        data = i.Item("data")
        Set bo = (new clsCmCalendar).of(data)
        e = i.Item("expect")
        a = ao.compareTo(bo)
        AssertEqualWithMessage e, a, "No="&i.Item("No")&" ao="&ao.formatAs("YYYY/MM/DD hh:mm:ss.000000")&" data="&cf_toString(i)
    Next
End Sub
Sub Test_clsCmCalendar_compareTo_TimeOnly
    dim a,e,d,i,data,ao,bo
    d = Array ( _
            new_DicOf(Array(  "No", 1,"data", "00:59:30", "expect", 1 ))_
            , new_DicOf(Array("No", 2,"data", "00:59:31", "expect", 0 ))_
            , new_DicOf(Array("No", 2,"data", "00:59:32", "expect", -1))_
            )
    Set ao = (new clsCmCalendar).of("00:59:31")

    For Each i In d
        data = i.Item("data")
        Set bo = (new clsCmCalendar).of(data)
        e = i.Item("expect")
        a = ao.compareTo(bo)
        AssertEqualWithMessage e, a, "No="&i.Item("No")&" ao="&ao.formatAs("YYYY/MM/DD hh:mm:ss.000000")&" data="&cf_toString(i)
    Next
End Sub
Sub Test_clsCmCalendar_compareTo_Err
    On Error Resume Next
    Dim d,e,a
    d = "2024/02/29 00:59:31"

    e = Empty
    a = (new clsCmCalendar).of(d).compareTo(new_Dic())

    AssertEqualWithMessage e, a, "ret"
    AssertEqualWithMessage "clsCmCalendar+compareTo()", Err.Source, "Err.Source"
    AssertEqualWithMessage "That object is not a calendar class.", Err.Description, "Err.Description"
End Sub

'###################################################################################################
'clsCmCalendar.differenceFrom()
Sub Test_clsCmCalendar_differenceFrom_initial
    dim a,ao,bo
    set ao = (new clsCmCalendar)

    set bo = (new clsCmCalendar)
    a = ao.differenceFrom(bo)
    AssertWithMessage a=0, "differenceFrom()=0 a="&a&" ao="&ao.toString()&" bo="&bo.toString()

    set bo = (new clsCmCalendar).ofNow()
    a = ao.differenceFrom(bo)
    AssertWithMessage a<0, "differenceFrom()<0 a="&a&" ao="&ao.toString()&" bo="&bo.toString()

    ao.ofNow()
    set bo = (new clsCmCalendar)
    a = ao.differenceFrom(bo)
    AssertWithMessage a>0, "differenceFrom()>0 a="&a&" ao="&ao.toString()&" bo="&bo.toString()
End Sub
Sub Test_clsCmCalendar_differenceFrom
    dim a,e,d,i,data,ao,bo
    d = Array ( _
            new_DicOf(Array(  "No", 1,"data", "2024/02/29 00:59:30", "expect", 1 ))_
            , new_DicOf(Array("No", 2,"data", "2024/02/29 00:59:31", "expect", 0 ))_
            , new_DicOf(Array("No", 2,"data", "2024/02/29 00:59:32", "expect", -1))_
            )
    Set ao = (new clsCmCalendar).of("2024/02/29 00:59:31")

    For Each i In d
        data = i.Item("data")
        Set bo = (new clsCmCalendar).of(data)
        e = i.Item("expect")
        a = ao.differenceFrom(bo)
        AssertEqualWithMessage e, a, "No="&i.Item("No")&" ao="&ao.formatAs("YYYY/MM/DD hh:mm:ss.000000")&" data="&cf_toString(i)
    Next
End Sub
Sub Test_clsCmCalendar_differenceFrom_WithDecimal
    dim a,e,d,i,data,ao,bo
    d = Array ( _
            new_DicOf(Array(  "No", 1,"data", "2024/02/29 00:59:31.123455", "expect", 0.000001 ))_
            , new_DicOf(Array("No", 2,"data", "2024/02/29 00:59:31.123456", "expect", 0        ))_
            , new_DicOf(Array("No", 2,"data", "2024/02/29 00:59:31.123457", "expect", -0.000001))_
            )
    Set ao = (new clsCmCalendar).of("2024/02/29 00:59:31.123456")

    For Each i In d
        data = i.Item("data")
        Set bo = (new clsCmCalendar).of(data)
        e = i.Item("expect")
        a = ao.differenceFrom(bo)
        AssertEqualWithMessage e, a, "No="&i.Item("No")&" ao="&ao.formatAs("YYYY/MM/DD hh:mm:ss.000000")&" data="&cf_toString(i)
    Next
End Sub
Sub Test_clsCmCalendar_differenceFrom_DateOnly
    dim a,e,d,i,data,ao,bo
    d = Array ( _
            new_DicOf(Array(  "No", 1,"data", "2024/02/28", "expect", 24*60*60   ))_
            , new_DicOf(Array("No", 2,"data", "2024/02/29", "expect", 0          ))_
            , new_DicOf(Array("No", 2,"data", "2024/03/01", "expect", -1*24*60*60))_
            )
    Set ao = (new clsCmCalendar).of("2024/02/29")

    For Each i In d
        data = i.Item("data")
        Set bo = (new clsCmCalendar).of(data)
        e = i.Item("expect")
        a = ao.differenceFrom(bo)
        AssertEqualWithMessage e, a, "No="&i.Item("No")&" ao="&ao.formatAs("YYYY/MM/DD hh:mm:ss.000000")&" data="&cf_toString(i)
    Next
End Sub
Sub Test_clsCmCalendar_differenceFrom_TimeOnly
    dim a,e,d,i,data,ao,bo
    d = Array ( _
            new_DicOf(Array(  "No", 1,"data", "00:59:30", "expect", 1 ))_
            , new_DicOf(Array("No", 2,"data", "00:59:31", "expect", 0 ))_
            , new_DicOf(Array("No", 2,"data", "00:59:32", "expect", -1))_
            )
    Set ao = (new clsCmCalendar).of("00:59:31")

    For Each i In d
        data = i.Item("data")
        Set bo = (new clsCmCalendar).of(data)
        e = i.Item("expect")
        a = ao.differenceFrom(bo)
        AssertEqualWithMessage e, a, "No="&i.Item("No")&" ao="&ao.formatAs("YYYY/MM/DD hh:mm:ss.000000")&" data="&cf_toString(i)
    Next
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
Sub Test_clsCmCalendar_formatAs
    dim a,e,d,i,data,format
    d = Array ( _
            new_DicOf(Array(  "No", 1,"data", "2024/02/29 00:59:30.456789" , "expect", "24/2/29 00:59:30.456","format", "YY/M/d hh:mm:ss.000"         ))_
            , new_DicOf(Array("No", 2,"data", "2024/02/29 00:59:30.456789" , "expect", "fujii.txt"           ,"format", "fujii.txt"                   ))_
            , new_DicOf(Array("No", 3,"data", "2024/02/29 00:59:30"        , "expect", "2024"                ,"format", "YYYY"               , "lcace"))_
            , new_DicOf(Array("No", 4,"data", "2024/02/29 00:59:30.123456" , "expect", "24"                  ,"format", "YY"                 , "lcace"))_
            , new_DicOf(Array("No", 5,"data", "2024/02/29 00:59:30"        , "expect", "02"                  ,"format", "MM"                          ))_
            , new_DicOf(Array("No", 6,"data", "2024/10/29 00:59:30.123456" , "expect", "10"                  ,"format", "MM"                          ))_
            , new_DicOf(Array("No", 7,"data", "2024/02/29 00:59:30.123456" , "expect", "2"                   ,"format", "M"                           ))_
            , new_DicOf(Array("No", 8,"data", "2024/11/29 00:59:30"        , "expect", "11"                  ,"format", "M"                           ))_
            , new_DicOf(Array("No", 9,"data", "2024/02/09 00:59:30"        , "expect", "09"                  ,"format", "DD"                 , "lcace"))_
            , new_DicOf(Array("No",10,"data", "2024/02/29 00:59:30.123456" , "expect", "29"                  ,"format", "DD"                 , "lcace"))_
            , new_DicOf(Array("No",11,"data", "2024/02/01 00:59:30.123456" , "expect", "1"                   ,"format", "D"                  , "lcace"))_
            , new_DicOf(Array("No",12,"data", "2024/02/29 00:59:30"        , "expect", "29"                  ,"format", "D"                  , "lcace"))_
            , new_DicOf(Array("No",13,"data", "2024/02/29 00:59:30"        , "expect", "00"                  ,"format", "HH"                 , "lcace"))_
            , new_DicOf(Array("No",14,"data", "2024/02/29 23:59:30.123456" , "expect", "23"                  ,"format", "HH"                 , "lcace"))_
            , new_DicOf(Array("No",13,"data", "2024/02/29 09:59:30.123456" , "expect", "9"                   ,"format", "H"                  , "lcace"))_
            , new_DicOf(Array("No",14,"data", "2024/02/29 10:59:30"        , "expect", "10"                  ,"format", "H"                  , "lcace"))_
            , new_DicOf(Array("No",15,"data", "2024/02/29 00:00:30"        , "expect", "00"                  ,"format", "mm"                          ))_
            , new_DicOf(Array("No",16,"data", "2024/02/29 00:59:30.123456" , "expect", "59"                  ,"format", "mm"                          ))_
            , new_DicOf(Array("No",17,"data", "2024/02/29 00:09:30.123456" , "expect", "9"                   ,"format", "m"                           ))_
            , new_DicOf(Array("No",18,"data", "2024/02/29 00:10:30"        , "expect", "10"                  ,"format", "m"                           ))_
            , new_DicOf(Array("No",19,"data", "2024/02/29 00:59:30"        , "expect", "30"                  ,"format", "SS"                 , "lcace"))_
            , new_DicOf(Array("No",20,"data", "2024/02/29 00:59:09.123456" , "expect", "09"                  ,"format", "SS"                 , "lcace"))_
            , new_DicOf(Array("No",21,"data", "2024/02/29 00:59:00.123456" , "expect", "0"                   ,"format", "S"                  , "lcace"))_
            , new_DicOf(Array("No",22,"data", "2024/02/29 00:59:10"        , "expect", "10"                  ,"format", "S"                  , "lcace"))_
            , new_DicOf(Array("No",23,"data", "2024/02/29 00:59:10.1234567", "expect", "123456"              ,"format", "000000"                      ))_
            , new_DicOf(Array("No",24,"data", "2024/02/29 00:59:10.987654" , "expect", "987654"              ,"format", "000000"                      ))_
            , new_DicOf(Array("No",25,"data", "2024/02/29 00:59:10.12345"  , "expect", "123450"              ,"format", "000000"                      ))_
            , new_DicOf(Array("No",26,"data", "2024/02/29 00:59:10.9876"   , "expect", "987600"              ,"format", "000000"                      ))_
            , new_DicOf(Array("No",27,"data", "2024/02/29 00:59:10.123"    , "expect", "123000"              ,"format", "000000"                      ))_
            , new_DicOf(Array("No",28,"data", "2024/02/29 00:59:10.98"     , "expect", "980000"              ,"format", "000000"                      ))_
            , new_DicOf(Array("No",29,"data", "2024/02/29 00:59:10.1"      , "expect", "100000"              ,"format", "000000"                      ))_
            , new_DicOf(Array("No",30,"data", "2024/02/29 00:59:10"        , "expect", "000000"              ,"format", "000000"                      ))_
            , new_DicOf(Array("No",31,"data", "2024/02/29 00:59:10.9876"   , "expect", "987"                 ,"format", "000"                         ))_
            , new_DicOf(Array("No",32,"data", "2024/02/29 00:59:10.123"    , "expect", "123"                 ,"format", "000"                         ))_
            , new_DicOf(Array("No",33,"data", "2024/02/29 00:59:10.98"     , "expect", "980"                 ,"format", "000"                         ))_
            , new_DicOf(Array("No",34,"data", "2024/02/29 00:59:10.1"      , "expect", "100"                 ,"format", "000"                         ))_
            , new_DicOf(Array("No",35,"data", "2024/02/29 00:59:10"        , "expect", "000"                 ,"format", "000"                         ))_
            )

    For Each i In d
        format = i.Item("format")
        data = i.Item("data")
        e = i.Item("expect")
        a = (new clsCmCalendar).of(data).formatAs(format)
        AssertEqualWithMessage e, a, "No="&i.Item("No")&" data="&cf_toString(i)
        If i.Exists("lcace") Then
            a = (new clsCmCalendar).of(data).formatAs(LCase(format))
            AssertEqualWithMessage e, a, "No="&i.Item("No")&"(lcase) data="&cf_toString(i)
        End If
    Next
End Sub

'###################################################################################################
'clsCmCalendar.of()
Sub Test_clsCmCalendar_of_Err
    On Error Resume Next
    Dim d
    d = "xyz"
    Call (new clsCmCalendar).of(d)

    AssertEqualWithMessage "clsCmCalendar+of()", Err.Source, "Err.Source"
    AssertMatchWithMessage "^invalid argument.*"&d, Err.Description, "Err.Description"
End Sub
Sub Test_clsCmCalendar_of_ErrImmutable
    On Error Resume Next
    Dim ao
    Set ao = (new clsCmCalendar).ofNow()
    Call ao.of("2025/2/22 22:22:22")

    AssertEqualWithMessage "clsCmCalendar+of()", Err.Source, "Err.Source"
    AssertEqualWithMessage "Because it is an immutable variable, its value cannot be changed.", Err.Description, "Err.Description"
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

    AssertEqualWithMessage e, a.toString(), "toString()"
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
