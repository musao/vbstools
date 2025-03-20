' Calendar.vbs: test.
' @import ../../lib/com/FileSystemProxy.vbs
' @import ../../lib/com/ArrayList.vbs
' @import ../../lib/com/Broker.vbs
' @import ../../lib/com/BufferedReader.vbs
' @import ../../lib/com/BufferedWriter.vbs
' @import ../../lib/com/Calendar.vbs
' @import ../../lib/com/CharacterType.vbs
' @import ../../lib/com/CssGenerator.vbs
' @import ../../lib/com/HtmlGenerator.vbs
' @import ../../lib/com/ReadOnlyObject.vbs
' @import ../../lib/com/ReturnValue.vbs
' @import ../../lib/com/libCom.vbs

Option Explicit

'###################################################################################################
'Calendar
Sub Test_Calendar
    Dim a : Set a = new Calendar
    AssertEqual 8, VarType(a)
    AssertEqual "Calendar", TypeName(a)
End Sub

'###################################################################################################
'Calendar.dateTime,fractionalPartOfElapsedSeconds,elapsedSeconds,serial
Sub Test_Calendar_dateTime_fractionalPartOfElapsedSeconds_elapsedSeconds_serial_initial
    dim tg,a,ao,e
    set ao = (new Calendar)

    tg = "A.dateTime"
    e = Null
    a = ao.dateTime
    AssertEqualWithMessage e, a, tg

    tg = "B.fractionalPartOfElapsedSeconds"
    e = Null
    a = ao.fractionalPartOfElapsedSeconds
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
Sub Test_Calendar_dateTime_fractionalPartOfElapsedSeconds_elapsedSeconds_serial_elapsedSeconds_Null
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
        set ao = (new Calendar).of(data)
        
        tg = "A.dateTime"
        e = CDate(data)
        a = ao.dateTime
        AssertEqualWithMessage e, a, tg&" No="&i.Item("No")&", data="&i.Item("date")
        
        tg = "B.fractionalPartOfElapsedSeconds"
        e = 0
        a = ao.fractionalPartOfElapsedSeconds
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
Sub Test_Calendar_dateTime_fractionalPartOfElapsedSeconds_elapsedSeconds_serial_elapsedSeconds_NotNull
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
        set ao = (new Calendar).of(data)
        
        tg = "A.dateTime"
        e = CDate(data(0))
        a = ao.dateTime
        AssertEqualWithMessage e, a, tg&" No="&i.Item("No")&", data="&cf_toString(data)
        
        tg = "B.fractionalPartOfElapsedSeconds"
        e = data(1)-Fix(data(1))
        a = ao.fractionalPartOfElapsedSeconds
        AssertWithMessage Abs(e-a)<0.0000001, tg&" No="&i.Item("No")&", data="&cf_toString(data)&", e="&cf_toString(e)&", a="&cf_toString(a)&", (e-a)="&cf_toString(e-a)
'        AssertWithMessage Abs(e-a)<0.0000001 Or (1-Abs(e-a))<0.0000001, tg&" No="&i.Item("No")&", data="&cf_toString(data)&", e="&cf_toString(e)&", a="&cf_toString(a)&", (e-a)="&cf_toString(e-a)
        
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
'Calendar.toString
Sub Test_Calendar_toString_initial
    dim a,ao,e
    set ao = (new Calendar)

    e = "<Calendar><Null>"
    a = ao.toString()
    AssertEqualWithMessage e, a, "toString()"
End Sub
Sub Test_Calendar_toString
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
        a = (new Calendar).of(data).toString()
        AssertEqualWithMessage e, a, "No="&i.Item("No")&", data="&cf_toString(data)
    Next
End Sub

'###################################################################################################
'Calendar.clone()
Sub Test_Calendar_clone_initial
    dim a,ao,e,bo
    set ao = (new Calendar)
    set bo = ao.clone()

    e = 0
    a = ao.compareTo(bo)
    AssertEqualWithMessage e, a, "clone()"
End Sub
Sub Test_Calendar_clone
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
        Set ao = (new Calendar).of(data)
        Set bo = ao.clone()
        e = 0
        a = ao.compareTo(bo)
        AssertEqualWithMessage e, a, "No="&i.Item("No")&", data="&cf_toString(data)
    Next
End Sub

'###################################################################################################
'Calendar.compareTo()
Sub Test_Calendar_compareTo_initial
    dim a,ao,e,bo
    set ao = (new Calendar)

    set bo = (new Calendar)
    e = 0
    a = ao.compareTo(bo)
    AssertEqualWithMessage e, a, "compareTo()=0 ao="&ao.toString()&" bo="&bo.toString()

    set bo = (new Calendar).ofNow()
    e = -1
    a = ao.compareTo(bo)
    AssertEqualWithMessage e, a, "compareTo()<0 ao="&ao.toString()&" bo="&bo.toString()

    ao.ofNow()
    set bo = (new Calendar)
    e = 1
    a = ao.compareTo(bo)
    AssertEqualWithMessage e, a, "compareTo()>0 ao="&ao.toString()&" bo="&bo.toString()
End Sub
Sub Test_Calendar_compareTo
    dim a,e,d,i,data,ao,bo
    d = Array ( _
            new_DicOf(Array(  "No", 1,"data", "2024/02/29 00:59:30", "expect", 1 ))_
            , new_DicOf(Array("No", 2,"data", "2024/02/29 00:59:31", "expect", 0 ))_
            , new_DicOf(Array("No", 2,"data", "2024/02/29 00:59:32", "expect", -1))_
            )
    Set ao = (new Calendar).of("2024/02/29 00:59:31")

    For Each i In d
        data = i.Item("data")
        Set bo = (new Calendar).of(data)
        e = i.Item("expect")
        a = ao.compareTo(bo)
        AssertEqualWithMessage e, a, "No="&i.Item("No")&" ao="&ao.formatAs("YYYY/MM/DD hh:mm:ss.000000")&" data="&cf_toString(i)
    Next
End Sub
Sub Test_Calendar_compareTo_WithDecimal
    dim a,e,d,i,data,ao,bo
    d = Array ( _
            new_DicOf(Array(  "No", 1,"data", "2024/02/29 00:59:31.123455", "expect", 1 ))_
            , new_DicOf(Array("No", 2,"data", "2024/02/29 00:59:31.123456", "expect", 0 ))_
            , new_DicOf(Array("No", 2,"data", "2024/02/29 00:59:31.123457", "expect", -1))_
            )
    Set ao = (new Calendar).of("2024/02/29 00:59:31.123456")

    For Each i In d
        data = i.Item("data")
        Set bo = (new Calendar).of(data)
        e = i.Item("expect")
        a = ao.compareTo(bo)
        AssertEqualWithMessage e, a, "No="&i.Item("No")&" ao="&ao.formatAs("YYYY/MM/DD hh:mm:ss.000000")&" data="&cf_toString(i)
    Next
End Sub
Sub Test_Calendar_compareTo_DateOnly
    dim a,e,d,i,data,ao,bo
    d = Array ( _
            new_DicOf(Array(  "No", 1,"data", "2024/02/28", "expect", 1 ))_
            , new_DicOf(Array("No", 2,"data", "2024/02/29", "expect", 0 ))_
            , new_DicOf(Array("No", 2,"data", "2024/03/01", "expect", -1))_
            )
    Set ao = (new Calendar).of("2024/02/29")

    For Each i In d
        data = i.Item("data")
        Set bo = (new Calendar).of(data)
        e = i.Item("expect")
        a = ao.compareTo(bo)
        AssertEqualWithMessage e, a, "No="&i.Item("No")&" ao="&ao.formatAs("YYYY/MM/DD hh:mm:ss.000000")&" data="&cf_toString(i)
    Next
End Sub
Sub Test_Calendar_compareTo_TimeOnly
    dim a,e,d,i,data,ao,bo
    d = Array ( _
            new_DicOf(Array(  "No", 1,"data", "00:59:30", "expect", 1 ))_
            , new_DicOf(Array("No", 2,"data", "00:59:31", "expect", 0 ))_
            , new_DicOf(Array("No", 2,"data", "00:59:32", "expect", -1))_
            )
    Set ao = (new Calendar).of("00:59:31")

    For Each i In d
        data = i.Item("data")
        Set bo = (new Calendar).of(data)
        e = i.Item("expect")
        a = ao.compareTo(bo)
        AssertEqualWithMessage e, a, "No="&i.Item("No")&" ao="&ao.formatAs("YYYY/MM/DD hh:mm:ss.000000")&" data="&cf_toString(i)
    Next
End Sub
Sub Test_Calendar_compareTo_Err
    On Error Resume Next
    Dim d,e,a
    d = "2024/02/29 00:59:31"

    e = Empty
    a = (new Calendar).of(d).compareTo(new_Dic())

    AssertEqualWithMessage e, a, "ret"
    AssertEqualWithMessage "Calendar+compareTo()", Err.Source, "Err.Source"
    AssertEqualWithMessage "That object is not a calendar class.", Err.Description, "Err.Description"
End Sub

'###################################################################################################
'Calendar.differenceFrom()
Sub Test_Calendar_differenceFrom_initial
    dim a,ao,bo
    set ao = (new Calendar)

    set bo = (new Calendar)
    a = ao.differenceFrom(bo)
    AssertWithMessage a=0, "differenceFrom()=0 a="&a&" ao="&ao.toString()&" bo="&bo.toString()

    set bo = (new Calendar).ofNow()
    a = ao.differenceFrom(bo)
    AssertWithMessage a<0, "differenceFrom()<0 a="&a&" ao="&ao.toString()&" bo="&bo.toString()

    ao.ofNow()
    set bo = (new Calendar)
    a = ao.differenceFrom(bo)
    AssertWithMessage a>0, "differenceFrom()>0 a="&a&" ao="&ao.toString()&" bo="&bo.toString()
End Sub
Sub Test_Calendar_differenceFrom
    dim a,e,d,i,data,ao,bo
    d = Array ( _
            new_DicOf(Array(  "No", 1,"data", "2024/02/29 00:59:30", "expect", 1 ))_
            , new_DicOf(Array("No", 2,"data", "2024/02/29 00:59:31", "expect", 0 ))_
            , new_DicOf(Array("No", 2,"data", "2024/02/29 00:59:32", "expect", -1))_
            )
    Set ao = (new Calendar).of("2024/02/29 00:59:31")

    For Each i In d
        data = i.Item("data")
        Set bo = (new Calendar).of(data)
        e = i.Item("expect")
        a = ao.differenceFrom(bo)
        AssertEqualWithMessage e, a, "No="&i.Item("No")&" ao="&ao.formatAs("YYYY/MM/DD hh:mm:ss.000000")&" data="&cf_toString(i)
    Next
End Sub
Sub Test_Calendar_differenceFrom_WithDecimal
    dim a,e,d,i,data,ao,bo
    d = Array ( _
            new_DicOf(Array(  "No", 1,"data", "2024/02/29 00:59:31.123455", "expect", 0.000001 ))_
            , new_DicOf(Array("No", 2,"data", "2024/02/29 00:59:31.123456", "expect", 0        ))_
            , new_DicOf(Array("No", 2,"data", "2024/02/29 00:59:31.123457", "expect", -0.000001))_
            )
    Set ao = (new Calendar).of("2024/02/29 00:59:31.123456")

    For Each i In d
        data = i.Item("data")
        Set bo = (new Calendar).of(data)
        e = i.Item("expect")
        a = ao.differenceFrom(bo)
        AssertEqualWithMessage e, a, "No="&i.Item("No")&" ao="&ao.formatAs("YYYY/MM/DD hh:mm:ss.000000")&" data="&cf_toString(i)
    Next
End Sub
Sub Test_Calendar_differenceFrom_DateOnly
    dim a,e,d,i,data,ao,bo
    d = Array ( _
            new_DicOf(Array(  "No", 1,"data", "2024/02/28", "expect", 24*60*60   ))_
            , new_DicOf(Array("No", 2,"data", "2024/02/29", "expect", 0          ))_
            , new_DicOf(Array("No", 2,"data", "2024/03/01", "expect", -1*24*60*60))_
            )
    Set ao = (new Calendar).of("2024/02/29")

    For Each i In d
        data = i.Item("data")
        Set bo = (new Calendar).of(data)
        e = i.Item("expect")
        a = ao.differenceFrom(bo)
        AssertEqualWithMessage e, a, "No="&i.Item("No")&" ao="&ao.formatAs("YYYY/MM/DD hh:mm:ss.000000")&" data="&cf_toString(i)
    Next
End Sub
Sub Test_Calendar_differenceFrom_TimeOnly
    dim a,e,d,i,data,ao,bo
    d = Array ( _
            new_DicOf(Array(  "No", 1,"data", "00:59:30", "expect", 1 ))_
            , new_DicOf(Array("No", 2,"data", "00:59:31", "expect", 0 ))_
            , new_DicOf(Array("No", 2,"data", "00:59:32", "expect", -1))_
            )
    Set ao = (new Calendar).of("00:59:31")

    For Each i In d
        data = i.Item("data")
        Set bo = (new Calendar).of(data)
        e = i.Item("expect")
        a = ao.differenceFrom(bo)
        AssertEqualWithMessage e, a, "No="&i.Item("No")&" ao="&ao.formatAs("YYYY/MM/DD hh:mm:ss.000000")&" data="&cf_toString(i)
    Next
End Sub
Sub Test_Calendar_differenceFrom_Err
    On Error Resume Next
    Dim d1,e,a,a1
    d1 = "2024/02/29 00:59:31"
    Set a1 = (new Calendar).of(d1)

    e = Empty
    a = a1.differenceFrom(new_Dic())

    AssertEqualWithMessage e, a, "ret"
    AssertEqualWithMessage "Calendar+differenceFrom()", Err.Source, "Err.Source"
    AssertEqualWithMessage "That object is not a calendar class.", Err.Description, "Err.Description"
End Sub

'###################################################################################################
'Calendar.formatAs()
Sub Test_Calendar_formatAs
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
        a = (new Calendar).of(data).formatAs(format)
        AssertEqualWithMessage e, a, "No="&i.Item("No")&" data="&cf_toString(i)
        If i.Exists("lcace") Then
            a = (new Calendar).of(data).formatAs(LCase(format))
            AssertEqualWithMessage e, a, "No="&i.Item("No")&"(lcase) data="&cf_toString(i)
        End If
    Next
End Sub

'###################################################################################################
'Calendar.of()
Sub Test_Calendar_of_1Arg_yyyymmdd_hhmmss
    dim a,e,d,i,data
    d = Array ( _
            new_DicOf(Array(  "No", 1,"data", "1899-10-01 00:59:30", "expect", "1899/10/01 00:59:30.000" ))_
            , new_DicOf(Array("No", 2,"data", "2024/02/29 23.00.59", "expect", "2024/02/29 23:00:59.000" ))_
            , new_DicOf(Array("No", 3,"data", "3000/06/15 12:34:00", "expect", "3000/06/15 12:34:00.000"))_
            )
    
    For Each i In d
        data = i.Item("data")
        e = i.Item("expect")
        a = (new Calendar).of(data)
        AssertEqualWithMessage e, a, "No="&i.Item("No")&" data="&cf_toString(i)
        a = (new Calendar).of(Array(data))
        AssertEqualWithMessage e, a, "No="&i.Item("No")&"(array) data="&cf_toString(i)
    Next
End Sub
Sub Test_Calendar_of_1Arg_yyyymmdd_hhmmss_Err
    Dim d
    d = Array ( _
            "2022-02-29 00:59:30" _
            , "2024/02/29 00.ab.30" _
        )

    Dim i
    For Each i In d
        Call of_Err_Detail(i)
        Call of_Err_Detail(Array(i))
    Next
End Sub
Sub Test_Calendar_of_1Arg_yyyymmdd_hhmmss_000
    dim a,e,d,i,data
    d = Array ( _
            new_DicOf(Array(  "No", 1,"data", "1900/10/01 00:59:30.123", "expect", "1900/10/01 00:59:30.123" ))_
            , new_DicOf(Array("No", 2,"data", "2024-02-29 23:00:59.987", "expect", "2024/02/29 23:00:59.987" ))_
            , new_DicOf(Array("No", 3,"data", "3000/06/15 12:34:00.001", "expect", "3000/06/15 12:34:00.001" ))_
            )
    
    For Each i In d
        data = i.Item("data")
        e = i.Item("expect")
        a = (new Calendar).of(data)
        AssertEqualWithMessage e, a, "No="&i.Item("No")&" data="&cf_toString(i)
        a = (new Calendar).of(Array(data))
        AssertEqualWithMessage e, a, "No="&i.Item("No")&"(array) data="&cf_toString(i)
    Next
End Sub
Sub Test_Calendar_of_1Arg_yyyymmdd_hhmmss_000_Err
    Dim d
    d = Array ( _
            , "1900/13/31 00:59:30.123" _
            , "2024/12/31 00:59:60.123" _
            , "3000/12/31 00:59:30.12a" _
            , "2025.02.23 12:34:56.789" _
        )

    Dim i
    For Each i In d
        Call of_Err_Detail(i)
        Call of_Err_Detail(Array(i))
    Next
End Sub
Sub Test_Calendar_of_1Arg_yyyymmdd_hhmmss_000000
    dim a,e,d,i,data
    d = Array ( _
            new_DicOf(Array(  "No", 1,"data", "1900/10/01 00:59:30.123456", "expect", "1900/10/01 00:59:30.123" ))_
            , new_DicOf(Array("No", 2,"data", "2024/02/29 23:00:59.987654", "expect", "2024/02/29 23:00:59.987" ))_
            , new_DicOf(Array("No", 3,"data", "3000-06-15 12:34:00.000001", "expect", "3000/06/15 12:34:00.000"))_
            )
    
    For Each i In d
        data = i.Item("data")
        e = i.Item("expect")
        a = (new Calendar).of(data)
        AssertEqualWithMessage e, a, "No="&i.Item("No")&" data="&cf_toString(i)
        a = (new Calendar).of(Array(data))
        AssertEqualWithMessage e, a, "No="&i.Item("No")&"(array) data="&cf_toString(i)
    Next
End Sub
Sub Test_Calendar_of_1Arg_yyyymmdd_hhmmss_000000_Err
    Dim d
    d = Array ( _
            , "1899/13/31 00:59:30.123456" _
            , "2023/12/31 00:59:60.123456" _
            , "3000/12/31 00:59:30.12a456" _
            , "2025.02.23 12:34:56.123456" _
        )

    Dim i
    For Each i In d
        Call of_Err_Detail(i)
        Call of_Err_Detail(Array(i))
    Next
End Sub
Sub Test_Calendar_of_1Arg_yyyymmdd
    dim a,e,d,i,data
    d = Array ( _
            new_DicOf(Array(  "No", 1,"data", "1899-10-01", "expect", "1899/10/01 00:00:00.000" ))_
            , new_DicOf(Array("No", 2,"data", "2024/02/29", "expect", "2024/02/29 00:00:00.000" ))_
            , new_DicOf(Array("No", 3,"data", "3000/06/15", "expect", "3000/06/15 00:00:00.000"))_
            )
    
    For Each i In d
        data = i.Item("data")
        e = i.Item("expect")
        a = (new Calendar).of(data)
        AssertEqualWithMessage e, a, "No="&i.Item("No")&" data="&cf_toString(i)
        a = (new Calendar).of(Array(data))
        AssertEqualWithMessage e, a, "No="&i.Item("No")&"(array) data="&cf_toString(i)
    Next
End Sub
Sub Test_Calendar_of_1Arg_yyyymmdd_Err
    Dim d
    d = Array ( _
            "1899/02/29" _
            , "2024/00/01" _
            , "3000/13/15" _
            , "2025.02.23" _
        )

    Dim i
    For Each i In d
        Call of_Err_Detail(i)
        Call of_Err_Detail(Array(i))
    Next
End Sub
Sub Test_Calendar_of_1Arg_hhmmss
    dim a,e,d,i,data
    d = Array ( _
            new_DicOf(Array(  "No", 1,"data", "00:59:30", "expect", "1899/12/30 00:59:30.000" ))_
            , new_DicOf(Array("No", 2,"data", "23:00:59", "expect", "1899/12/30 23:00:59.000" ))_
            , new_DicOf(Array("No", 3,"data", "12:34:00", "expect", "1899/12/30 12:34:00.000"))_
            , new_DicOf(Array("No", 4,"data", "13.24.57", "expect", "1899/12/30 13:24:57.000"))_
            )
    
    For Each i In d
        data = i.Item("data")
        e = i.Item("expect")
        a = (new Calendar).of(data)
        AssertEqualWithMessage e, a, "No="&i.Item("No")&" data="&cf_toString(i)
        a = (new Calendar).of(Array(data))
        AssertEqualWithMessage e, a, "No="&i.Item("No")&"(array) data="&cf_toString(i)
    Next
End Sub
Sub Test_Calendar_of_1Arg_hhmmss_Err
    Dim d
    d = Array ( _
            "25:59:30" _
            , "00:ab:30" _
        )

    Dim i
    For Each i In d
        Call of_Err_Detail(i)
        Call of_Err_Detail(Array(i))
    Next
End Sub
Sub Test_Calendar_of_1Arg_hhmmss_000000
    dim a,e,d,i,data
    d = Array ( _
            new_DicOf(Array(  "No", 1,"data", "00:59:30.123456", "expect", "1899/12/30 00:59:30.123" ))_
            , new_DicOf(Array("No", 2,"data", "23:00:59.987654", "expect", "1899/12/30 23:00:59.987" ))_
            , new_DicOf(Array("No", 3,"data", "12:34:00.000001", "expect", "1899/12/30 12:34:00.000"))_
            )
    
    For Each i In d
        data = i.Item("data")
        e = i.Item("expect")
        a = (new Calendar).of(data)
        AssertEqualWithMessage e, a, "No="&i.Item("No")&" data="&cf_toString(i)
        a = (new Calendar).of(Array(data))
        AssertEqualWithMessage e, a, "No="&i.Item("No")&"(array) data="&cf_toString(i)
    Next
End Sub
Sub Test_Calendar_of_1Arg_hhmmss_000000_Err
    Dim d
    d = Array ( _
            "23:59:30.12c456" _
            , "00:ab:30.98765" _
            , "23.45.01.23456" _
        )

    Dim i
    For Each i In d
        Call of_Err_Detail(i)
        Call of_Err_Detail(Array(i))
    Next
End Sub
Sub Test_Calendar_of_1Arg_hhmmss_000
    dim a,e,d,i,data
    d = Array ( _
            new_DicOf(Array(  "No", 1,"data", "00:59:30.123", "expect", "1899/12/30 00:59:30.123" ))_
            , new_DicOf(Array("No", 2,"data", "23:00:59.987", "expect", "1899/12/30 23:00:59.987" ))_
            , new_DicOf(Array("No", 3,"data", "12:34:00.001", "expect", "1899/12/30 12:34:00.001"))_
            )
    
    For Each i In d
        data = i.Item("data")
        e = i.Item("expect")
        a = (new Calendar).of(data)
        AssertEqualWithMessage e, a, "No="&i.Item("No")&" data="&cf_toString(i)
        a = (new Calendar).of(Array(data))
        AssertEqualWithMessage e, a, "No="&i.Item("No")&"(array) data="&cf_toString(i)
    Next
End Sub
Sub Test_Calendar_of_1Arg_hhmmss_000_Err
    Dim d
    d = Array ( _
            "23:59:30.12c" _
            , "00:ab:30.987" _
            , "12.34.56.789" _
        )

    Dim i
    For Each i In d
        Call of_Err_Detail(i)
        Call of_Err_Detail(Array(i))
    Next
End Sub

Sub Test_Calendar_of_2Args_yyyymmdd_hhmmss_000000
    dim a,e,d,i,data
    d = Array ( _
            new_DicOf(Array(  "No", 1,"data", Array("1900/10/01 00:59:30",  0*60*60 + 59*60 + 30+0.123456), "expect", "1900/10/01 00:59:30.123" ))_
            , new_DicOf(Array("No", 2,"data", Array("2024/02/29 23:00:59", 23*60*60 +  0*60 + 59+0.987654), "expect", "2024/02/29 23:00:59.987" ))_
            , new_DicOf(Array("No", 3,"data", Array("3000-06-15 12:34:00", 12*60*60 + 34*60 +  0+0.000001), "expect", "3000/06/15 12:34:00.000" ))_
            )
    
    For Each i In d
        data = i.Item("data")
        e = i.Item("expect")
        a = (new Calendar).of(data)
        AssertEqualWithMessage e, a, "No="&i.Item("No")&" data="&cf_toString(i)
    Next
End Sub
Sub Test_Calendar_of_2Args_yyyymmdd_hhmmss_000000_Err
    Dim d
    d = Array ( _
            Array("1899/13/31 00:59:30", 0.123456  ) _
            , Array("2023/12/31 00:59:60", 0.123456  ) _
            , Array("3000/12/31 00:59:30", "0.12a456") _
            , Array("2025.02.23 12:34:56", 0.123456  ) _
        )

    Dim i
    For Each i In d
        Call of_Err_Detail(i)
    Next
End Sub
Sub Test_Calendar_of_2Args_yyyymmdd_hhmmss_000
    dim a,e,d,i,data
    d = Array ( _
            new_DicOf(Array(  "No", 1,"data", Array("1900-10-01 00:59:30",  0*60*60 + 59*60 + 30+0.123), "expect", "1900/10/01 00:59:30.123" ))_
            , new_DicOf(Array("No", 2,"data", Array("2024/02/29 23:00:59", 23*60*60 +  0*60 + 59+0.987), "expect", "2024/02/29 23:00:59.987" ))_
            , new_DicOf(Array("No", 3,"data", Array("3000/06/15 12:34:00", 12*60*60 + 34*60 +  0+0.001), "expect", "3000/06/15 12:34:00.001" ))_
            )
    
    For Each i In d
        data = i.Item("data")
        e = i.Item("expect")
        a = (new Calendar).of(data)
        AssertEqualWithMessage e, a, "No="&i.Item("No")&" data="&cf_toString(i)
    Next
End Sub
Sub Test_Calendar_of_2Args_yyyymmdd_hhmmss_000_Err
    Dim d
    d = Array ( _
            Array("1899/13/31 00:59:30", 0.123  ) _
            , Array("2023-12-31 00:59:60", 0.123  ) _
            , Array("3000/12/31 00:59:30", "0.12a") _
            , Array("2025.02.23 12:34:56", 0.123  ) _
        )

    Dim i
    For Each i In d
        Call of_Err_Detail(i)
    Next
End Sub
Sub Test_Calendar_of_2Args_yyyymmdd_000000
    dim a,e,d,i,data
    d = Array ( _
            new_DicOf(Array(  "No", 1,"data", Array("1899/10/01", 0.000001), "expect", "1899/10/01 00:00:00.000" ))_
            , new_DicOf(Array("No", 2,"data", Array("2024/02/29", 0.123456), "expect", "2024/02/29 00:00:00.123" ))_
            , new_DicOf(Array("No", 3,"data", Array("3000-06-15", 0.987654), "expect", "3000/06/15 00:00:00.987"))_
            )
    
    For Each i In d
        data = i.Item("data")
        e = i.Item("expect")
        a = (new Calendar).of(data)
        AssertEqualWithMessage e, a, "No="&i.Item("No")&" data="&cf_toString(i)
    Next
End Sub
Sub Test_Calendar_of_2Args_yyyymmdd_000000_Err
    Dim d
    d = Array ( _
            Array("1899/02/29", 0.000001  ) _
            , Array("2024-00-01", 0.000001  ) _
            , Array("3000/13/15", 0.000001  ) _
            , Array("2025.02.23", 0.000001  ) _
            , Array("2025/02/23", "0.00xyz1") _
        )

    Dim i
    For Each i In d
        Call of_Err_Detail(i)
    Next
End Sub
Sub Test_Calendar_of_2Args_yyyymmdd_000
    dim a,e,d,i,data
    d = Array ( _
            new_DicOf(Array(  "No", 1,"data", Array("1899/10/01", 0.987), "expect", "1899/10/01 00:00:00.987" ))_
            , new_DicOf(Array("No", 2,"data", Array("2024/02/29", 0.001), "expect", "2024/02/29 00:00:00.001" ))_
            , new_DicOf(Array("No", 3,"data", Array("3000-06-15", 0.123), "expect", "3000/06/15 00:00:00.123"))_
            )
    
    For Each i In d
        data = i.Item("data")
        e = i.Item("expect")
        a = (new Calendar).of(data)
        AssertEqualWithMessage e, a, "No="&i.Item("No")&" data="&cf_toString(i)
    Next
End Sub
Sub Test_Calendar_of_2Args_yyyymmdd_000_Err
    Dim d
    d = Array ( _
            Array("1899/02/29", 0.001  ) _
            , Array("2024/00/01", 0.001  ) _
            , Array("3000/13/15", 0.001  ) _
            , Array("2025.02.23", 0.001  ) _
            , Array("2025/02/23", "0.op1") _
        )

    Dim i
    For Each i In d
        Call of_Err_Detail(i)
    Next
End Sub
Sub Test_Calendar_of_2Args_hhmmss_000000
    dim a,e,d,i,data
    d = Array ( _
            new_DicOf(Array(  "No", 1,"data", Array("00:59:30",  0*60*60 + 59*60 + 30+0.000001), "expect", "1899/12/30 00:59:30.000" ))_
            , new_DicOf(Array("No", 2,"data", Array("23:00:59", 23*60*60 +  0*60 + 59+0.123456), "expect", "1899/12/30 23:00:59.123" ))_
            , new_DicOf(Array("No", 3,"data", Array("12:34:00", 12*60*60 + 34*60 +  0+0.987654), "expect", "1899/12/30 12:34:00.987" ))_
            , new_DicOf(Array("No", 4,"data", Array("01.23.45",  1*60*60 + 23*60 + 45+0.357913), "expect", "1899/12/30 01:23:45.357" ))_
            )
    
    For Each i In d
        data = i.Item("data")
        e = i.Item("expect")
        a = (new Calendar).of(data)
        AssertEqualWithMessage e, a, "No="&i.Item("No")&" data="&cf_toString(i)
    Next
End Sub
Sub Test_Calendar_of_2Args_hhmmss_000000_Err
    Dim d
    d = Array ( _
            Array("23:59:30", "0.12c456") _
            , Array("00:ab:30", 0.98765) _
        )

    Dim i
    For Each i In d
        Call of_Err_Detail(i)
    Next
End Sub
Sub Test_Calendar_of_2Args_hhmmss_000
    dim a,e,d,i,data
    d = Array ( _
            new_DicOf(Array(  "No", 1,"data", Array("00.59.30",  0*60*60 + 59*60 + 30+0.001), "expect", "1899/12/30 00:59:30.001" ))_
            , new_DicOf(Array("No", 2,"data", Array("23:00:59", 23*60*60 +  0*60 + 59+0.357), "expect", "1899/12/30 23:00:59.357" ))_
            , new_DicOf(Array("No", 3,"data", Array("12:34:00", 12*60*60 + 34*60 +  0+0.123), "expect", "1899/12/30 12:34:00.123" ))_
            , new_DicOf(Array("No", 4,"data", Array("01:23:45",  1*60*60 + 23*60 + 45+0.987), "expect", "1899/12/30 01:23:45.987" ))_
            )
    
    For Each i In d
        data = i.Item("data")
        e = i.Item("expect")
        a = (new Calendar).of(data)
        AssertEqualWithMessage e, a, "No="&i.Item("No")&" data="&cf_toString(i)
    Next
End Sub
Sub Test_Calendar_of_2Args_hhmmss_000_Err
    Dim d
    d = Array ( _
            Array("23:59:30", "0.12c") _
            , Array("00:ab:30", 0.987) _
        )

    Dim i
    For Each i In d
        Call of_Err_Detail(i)
    Next
End Sub

Sub Test_Calendar_of_6Args_yyyymmdd_hhmmss
    dim a,e,d,i,data
    d = Array ( _
            new_DicOf(Array(  "No", 1,"data", Array(1899, 10,  1,  0, 59, 30), "expect", "1899/10/01 00:59:30.000" ))_
            , new_DicOf(Array("No", 2,"data", Array(2024,  2, 29, 23,  0, 59), "expect", "2024/02/29 23:00:59.000" ))_
            , new_DicOf(Array("No", 3,"data", Array(3000,  6, 15, 12, 34,  0), "expect", "3000/06/15 12:34:00.000"))_
            )
    
    For Each i In d
        data = i.Item("data")
        e = i.Item("expect")
        a = (new Calendar).of(data)
        AssertEqualWithMessage e, a, "No="&i.Item("No")&" data="&cf_toString(i)
    Next
End Sub
Sub Test_Calendar_of_6Args_yyyymmdd_hhmmss_Err
    Dim d
    d = Array ( _
            Array(2022,  2, 29,  0, 59, 30) _
            , Array(2024,  2, 29,  0, "ab", 30) _
        )

    Dim i
    For Each i In d
        Call of_Err_Detail(i)
    Next
End Sub

Sub Test_Calendar_of_7Args_yyyymmdd_hhmmss_000000
    dim a,e,d,i,data
    d = Array ( _
            new_DicOf(Array(  "No", 1,"data", Array(1900, 10,  1,  0, 59, 30,  0*60*60 + 59*60 + 30+0.123456), "expect", "1900/10/01 00:59:30.123" ))_
            , new_DicOf(Array("No", 2,"data", Array(2024,  2, 29, 23,  0, 59, 23*60*60 +  0*60 + 59+0.987654), "expect", "2024/02/29 23:00:59.987" ))_
            , new_DicOf(Array("No", 3,"data", Array(3000,  6, 15, 12, 34,  0, 12*60*60 + 34*60 +  0+0.000001), "expect", "3000/06/15 12:34:00.000" ))_
            )
    
    For Each i In d
        data = i.Item("data")
        e = i.Item("expect")
        a = (new Calendar).of(data)
        AssertEqualWithMessage e, a, "No="&i.Item("No")&" data="&cf_toString(i)
    Next
End Sub
Sub Test_Calendar_of_7Args_yyyymmdd_hhmmss_000000_Err
    Dim d
    d = Array ( _
            Array(  1899, 13, 31,  0, 59, 30, 0.123456  ) _
            , Array(2023, 12, 31,  0, 59, 60, 0.123456  ) _
            , Array(3000, 12, 31,  0, 59, 30, "0.12a456") _
        )

    Dim i
    For Each i In d
        Call of_Err_Detail(i)
    Next
End Sub
Sub Test_Calendar_of_7Args_yyyymmdd_hhmmss_000
    dim a,e,d,i,data
    d = Array ( _
            new_DicOf(Array(  "No", 1,"data", Array(1900, 10,  1,  0, 59, 30,  0*60*60 + 59*60 + 30+0.001), "expect", "1900/10/01 00:59:30.001" ))_
            , new_DicOf(Array("No", 2,"data", Array(2024,  2, 29, 23,  0, 59, 23*60*60 +  0*60 + 59+0.123), "expect", "2024/02/29 23:00:59.123" ))_
            , new_DicOf(Array("No", 3,"data", Array(3000,  6, 15, 12, 34,  0, 12*60*60 + 34*60 +  0+0.987), "expect", "3000/06/15 12:34:00.987" ))_
            )
    
    For Each i In d
        data = i.Item("data")
        e = i.Item("expect")
        a = (new Calendar).of(data)
        AssertEqualWithMessage e, a, "No="&i.Item("No")&" data="&cf_toString(i)
    Next
End Sub
Sub Test_Calendar_of_7Args_yyyymmdd_hhmmss_000_Err
    Dim d
    d = Array ( _
            Array(  1899, 13, 31,  0, 59, 30, 0.123  ) _
            , Array(2023, 12, 31,  0, 59, 60, 0.123  ) _
            , Array(3000, 12, 31,  0, 59, 30, "0.12a") _
        )

    Dim i
    For Each i In d
        Call of_Err_Detail(i)
    Next
End Sub

Sub Test_Calendar_of_Other_Err
    Dim d
    d = Array ( _
            Array("2025/02/23 17:09:30", 17*60*60 + 9*60 + 30 + 0.123456, "arg3") _
            , Array("2025/02/23 17:09:30", 17*60*60 + 9*60 + 30 + 0.123456, "arg3", "arg4") _
            , Array("2025/02/23 17:09:30", 17*60*60 + 9*60 + 30 + 0.123456, "arg3", "arg4", "arg5") _
            , Array("2025/02/23 17:09:30", 17*60*60 + 9*60 + 30 + 0.123456, "arg3", "arg4", "arg5", "arg6", "arg7", "arg8") _
        )

    Dim i
    For Each i In d
        Call of_Err_Detail(i)
    Next
End Sub
Sub Test_Calendar_of_ErrImmutable
    On Error Resume Next
    Dim ao
    Set ao = (new Calendar).ofNow()
    Call ao.of("2025/2/22 22:22:22")

    AssertEqualWithMessage "Calendar+of()", Err.Source, "Err.Source"
    AssertEqualWithMessage "Because it is an immutable variable, its value cannot be changed.", Err.Description, "Err.Description"
End Sub


'###################################################################################################
'Calendar.ofNow()
Sub Test_Calendar_ofNow
    Dim n,y,m,d,h,mm,s
    n = now()
    y = Right("000" & Year(n), 4)
    m = Right("0" & Month(n), 2)
    d = Right("0" & Day(n), 2)
    h = Right("0" & Hour(n), 2)
    mm = Right("0" & Minute(n), 2)
    s = Right("0" & Second(n), 2)
    Dim ptn : ptn = "^"&y&"/"&m&"/"&d&" "&h&":"&mm&":"&s&"\.\d{3}$"
    Dim a : Set a = (new Calendar).ofNow()

    AssertMatchWithMessage ptn, a.toString(), "a="&a.toString()
End Sub



'###################################################################################################
'common
Sub of_Err_Detail(arg)
    On Error Resume Next
    Dim ao : Set ao = (new Calendar).of(arg)

    Dim sSource,sDescription
    sSource = Err.Source
    sDescription = Err.Description

    AssertEqualWithMessage "Calendar+of()", sSource, "a="&ao.toString()&" Err.Source="&sSource&" arg="&cf_toString(arg)
    AssertMatchWithMessage "^invalid argument.*", sDescription, "a="&ao.toString()&" Err.Description="&sDescription&" arg="&cf_toString(arg)
    On Error Goto 0
End Sub


' Local Variables:
' mode: Visual-Basic
' indent-tabs-mode: nil
' End:
