' libCom.vbs: math_* procedure test.
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
'math_min()
Sub Test_math_min
    dim a,e,d,i,num1,num2
    d = Array ( _
            new_DicOf(Array(  "No",1 ,"Num1",1     ,"Num2",1     ,"Expected","Num1")) _
            , new_DicOf(Array("No",2 ,"Num1",2     ,"Num2",1     ,"Expected","Num2")) _
            , new_DicOf(Array("No",3 ,"Num1",-3    ,"Num2",-2    ,"Expected","Num1")) _
            , new_DicOf(Array("No",4 ,"Num1",1     ,"Num2",-2    ,"Expected","Num2")) _
            , new_DicOf(Array("No",5 ,"Num1",0.2   ,"Num2",0.3   ,"Expected","Num1")) _
            , new_DicOf(Array("No",6 ,"Num1",0.1   ,"Num2",-0.04 ,"Expected","Num2")) _
            , new_DicOf(Array("No",7 ,"Num1",-0.015,"Num2",-0.009,"Expected","Num1")) _
            )
    For Each i In d
        num1 = i.Item("Num1")
        num2 = i.Item("Num2")
        e = i.Item(i.Item("Expected"))
        a = math_min(num1,num2)
        AssertEqualWithMessage a, e, "No="&i.Item("No")&", Num1="&num1&", Num2="&num2
    Next
End Sub

'###################################################################################################
'math_max()
Sub Test_math_max
    dim a,e,d,i,num1,num2
    d = Array ( _
            new_DicOf(Array(  "No",1 ,"Num1",1     ,"Num2",1     ,"Expected","Num2")) _
            , new_DicOf(Array("No",2 ,"Num1",2     ,"Num2",1     ,"Expected","Num1")) _
            , new_DicOf(Array("No",3 ,"Num1",-3    ,"Num2",-2    ,"Expected","Num2")) _
            , new_DicOf(Array("No",4 ,"Num1",1     ,"Num2",-2    ,"Expected","Num1")) _
            , new_DicOf(Array("No",5 ,"Num1",0.2   ,"Num2",0.3   ,"Expected","Num2")) _
            , new_DicOf(Array("No",6 ,"Num1",0.1   ,"Num2",-0.04 ,"Expected","Num1")) _
            , new_DicOf(Array("No",7 ,"Num1",-0.015,"Num2",-0.009,"Expected","Num2")) _
            )
    For Each i In d
        num1 = i.Item("Num1")
        num2 = i.Item("Num2")
        e = i.Item(i.Item("Expected"))
        a = math_max(num1,num2)
        AssertEqualWithMessage a, e, "No="&i.Item("No")&", Num1="&num1&", Num2="&num2
    Next
End Sub

'###################################################################################################
'math_roundUp()
Sub Test_math_roundUp
    dim a,e,d,i,n,p
    d = Array ( _
            new_DicOf(Array(  "No",1 ,"Num",-12345.6789 ,"Place",-6 ,"Expected",0)) _
            , new_DicOf(Array("No",2 ,"Num",54545.4545 ,"Place",0   ,"Expected",54546)) _
            , new_DicOf(Array("No",3 ,"Num",10101.0101 ,"Place",4   ,"Expected",10101.0101)) _
            )
    For Each i In d
        n = i.Item("Num")
        p = i.Item("Place")
        e = i.Item("Expected")
        a = math_roundUp(n,p)
        AssertEqualWithMessage a, e, "No="&i.Item("No")&", Num="&n&", Place="&p
    Next
End Sub

'###################################################################################################
'math_round()
Sub Test_math_round
    dim a,e,d,i,n,p
    d = Array ( _
            new_DicOf(Array(  "No",1 ,"Num",-12345.6789 ,"Place",-6 ,"Expected",0)) _
            , new_DicOf(Array("No",2 ,"Num",54545.4545 ,"Place",0   ,"Expected",54545)) _
            , new_DicOf(Array("No",3 ,"Num",10101.0101 ,"Place",4   ,"Expected",10101.0101)) _
            )
    For Each i In d
        n = i.Item("Num")
        p = i.Item("Place")
        e = i.Item("Expected")
        a = math_round(n,p)
        AssertEqualWithMessage a, e, "No="&i.Item("No")&", Num="&n&", Place="&p
        If p>=0 Then AssertEqualWithMessage Round(n, p), a, "Regression No="&i.Item("No")&", Num="&n&", Place="&p
    Next
End Sub

'###################################################################################################
'math_roundDown()
Sub Test_math_roundDwon
    dim a,e,d,i,n,p
    d = Array ( _
            new_DicOf(Array(  "No",1 ,"Num",-12345.6789 ,"Place",-6 ,"Expected",0)) _
            , new_DicOf(Array("No",2 ,"Num",54545.4545 ,"Place",0   ,"Expected",54545)) _
            , new_DicOf(Array("No",3 ,"Num",10101.0101 ,"Place",4   ,"Expected",10101.0101)) _
            )
    For Each i In d
        n = i.Item("Num")
        p = i.Item("Place")
        e = i.Item("Expected")
        a = math_roundDown(n,p)
        AssertEqualWithMessage a, e, "No="&i.Item("No")&", Num="&n&", Place="&p
    Next
End Sub

'###################################################################################################
'math_rand()
Sub Test_math_rand
    dim d,m,n,p,a,i,j
    d = Array ( _
            new_DicOf(Array(  "No",1 ,"min",5 ,"max",10 ,"place",0 )) _
            , new_DicOf(Array(  "No",2 ,"min",-10 ,"max",-5 ,"place",0 )) _
            , new_DicOf(Array(  "No",3 ,"min",-5 ,"max",5 ,"place",0 )) _
            , new_DicOf(Array(  "No",4 ,"min",3.14 ,"max",12.345 ,"place",5 )) _
            , new_DicOf(Array(  "No",5 ,"min",-100.12345 ,"max",-1.23456 ,"place",5 )) _
            , new_DicOf(Array(  "No",6 ,"min",-10.246 ,"max",100.357 ,"place",3 )) _
            )
    For Each i In d
        n = i.Item("min")
        m = i.Item("max")
        p = i.Item("place")
        j=0
        Do While j<1000
            a = math_rand(n,m,p)
            If (a<n or m<a) Then
                AssertFailWithMessage "No="&i.Item("No")&", min="&n&", max="&m&", Place="&p&", math_rand(n,m,p)="&a&", j="&j
            End If
            j = j + 1
        Loop
        Assert True
    Next
End Sub

'###################################################################################################
'func_MathRound()
Sub Test_func_MathRound_d0
    dim a,e,d,i,dn,dp,dt
    dt = 0
    dn = 12345.6789
    d = Array( _
                   new_DicOf( Array("No", 1, "Place", -6, "Expected", 0) ) _
                   , new_DicOf( Array("No", 2, "Place", -5, "Expected", 0) ) _
                   , new_DicOf( Array("No", 3, "Place", -4, "Expected", 10000) ) _
                   , new_DicOf( Array("No", 4, "Place", -1, "Expected", 12340) ) _
                   , new_DicOf( Array("No", 5, "Place", 0, "Expected", 12345) ) _
                   , new_DicOf( Array("No", 6, "Place", 1, "Expected", 12345.6) ) _
                   , new_DicOf( Array("No", 7, "Place", 4, "Expected", 12345.6789) ) _
                   , new_DicOf( Array("No", 8, "Place", 5, "Expected", 12345.6789) ) _
                )
    For Each i In d
        dp = i.Item("Place")
        e = i.Item("Expected")
        a = func_MathRound(dn,dp,dt,True)
        AssertEqualWithMessage e, a, "PositiveNumber No = "&i.Item("No")& ", dn = "&dn&", dp = "&dp&", dt = "&dt
    Next

    dn = -1 * dn
    For Each i In d
        dp = i.Item("Place")
        e = -1 * i.Item("Expected")
        a = func_MathRound(dn,dp,dt,True)
        AssertEqualWithMessage e, a, "NegativeNumber No = "&i.Item("No")& ", dn = "&dn&", dp = "&dp&", dt = "&dt
    Next
End Sub
Sub Test_func_MathRound_d5
    dim a,e,d,i,dn,dp,dt
    dt = 5
    dn = 54545.4545
    d = Array( _
                   new_DicOf( Array("No", 1, "Place", -6, "Expected", 0) ) _
                   , new_DicOf( Array("No", 2, "Place", -5, "Expected", 100000) ) _
                   , new_DicOf( Array("No", 3, "Place", -4, "Expected", 50000) ) _
                   , new_DicOf( Array("No", 4, "Place", -1, "Expected", 54550) ) _
                   , new_DicOf( Array("No", 5, "Place", 0, "Expected", 54545) ) _
                   , new_DicOf( Array("No", 6, "Place", 1, "Expected", 54545.5) ) _
                   , new_DicOf( Array("No", 7, "Place", 4, "Expected", 54545.4545) ) _
                   , new_DicOf( Array("No", 8, "Place", 5, "Expected", 54545.4545) ) _
                )
    For Each i In d
        dp = i.Item("Place")
        e = i.Item("Expected")
        a = func_MathRound(dn,dp,dt,True)
        AssertEqualWithMessage e, a, "PositiveNumber No = "&i.Item("No")& ", dn = "&dn&", dp = "&dp&", dt = "&dt
        If dp>=0 Then AssertEqualWithMessage Round(dn, dp), a, "PositiveNumber Regression No = "&i.Item("No")& ", dn = "&dn&", dp = "&dp&", dt = "&dt
    Next

    dn = -1 * dn
    For Each i In d
        dp = i.Item("Place")
        e = -1 * i.Item("Expected")
        a = func_MathRound(dn,dp,dt,True)
        AssertEqualWithMessage e, a, "NegativeNumber No = "&i.Item("No")& ", dn = "&dn&", dp = "&dp&", dt = "&dt
        If dp>=0 Then AssertEqualWithMessage Round(dn, dp), a, "NegativeNumber Regression No = "&i.Item("No")& ", dn = "&dn&", dp = "&dp&", dt = "&dt
    Next
End Sub
Sub Test_func_MathRound_d9
    dim a,e,d,i,dn,dp,dt
    dt = 9
    dn = 10101.0101
    d = Array( _
                   new_DicOf( Array("No", 1, "Place", -6, "Expected", 0) ) _
                   , new_DicOf( Array("No", 2, "Place", -5, "Expected", 100000) ) _
                   , new_DicOf( Array("No", 3, "Place", -4, "Expected", 10000) ) _
                   , new_DicOf( Array("No", 4, "Place", -1, "Expected", 10110) ) _
                   , new_DicOf( Array("No", 5, "Place", 0, "Expected", 10101) ) _
                   , new_DicOf( Array("No", 6, "Place", 1, "Expected", 10101.1) ) _
                   , new_DicOf( Array("No", 7, "Place", 4, "Expected", 10101.0101) ) _
                   , new_DicOf( Array("No", 8, "Place", 5, "Expected", 10101.0101) ) _
                )
    For Each i In d
        dp = i.Item("Place")
        e = i.Item("Expected")
        a = func_MathRound(dn,dp,dt,True)
        AssertEqualWithMessage e, a, "PositiveNumber No = "&i.Item("No")& ", dn = "&dn&", dp = "&dp&", dt = "&dt
    Next

    dn = -1 * dn
    For Each i In d
        dp = i.Item("Place")
        e = -1 * i.Item("Expected")
        a = func_MathRound(dn,dp,dt,True)
        AssertEqualWithMessage e, a, "NegativeNumber PositiveNumber No = "&i.Item("No")& ", dn = "&dn&", dp = "&dp&", dt = "&dt
    Next
End Sub

'###################################################################################################
'math_log2()
Sub Test_math_log2
    dim d,i,al,e,a
    d = Array ( _
            new_DicOf(Array(  "No",1 ,"antilogarithm",1024 ,"expected",10 )) _
            , new_DicOf(Array(  "No",2 ,"antilogarithm",0.125 ,"expected",-3 )) _
            )
    For Each i In d
        al = i.Item("antilogarithm")
        e = i.Item("expected")
        a = math_log2(al)
        AssertEqualWithMessage Cstr(e), Cstr(a), "No = "&i.Item("No")&", antilogarithm = "&al
    Next
End Sub

'###################################################################################################
'func_MathLog()
Sub Test_func_MathLog
    dim d,i,b,al,e,a
    d = Array ( _
            new_DicOf(Array(  "No",1 ,"base",2 ,"antilogarithm",4 ,"expected",2 )) _
            , new_DicOf(Array(  "No",2 ,"base",10 ,"antilogarithm",0.01 ,"expected",-2 )) _
            )
    For Each i In d
        b = i.Item("base")
        al = i.Item("antilogarithm")
        e = i.Item("expected")
        a = func_MathLog(b,al)
        AssertEqualWithMessage Cstr(e), Cstr(a), "No = "&i.Item("No")& ", base = "&b&", antilogarithm = "&al
    Next
End Sub

'###################################################################################################
'math_tranc()
Sub Test_math_tranc
    dim a,e,d,i,num
    d = Array ( _
            new_DicOf(Array(  "No",1 ,"Num",1    ,"Expected",1  )) _
            , new_DicOf(Array("No",2 ,"Num",2.0  ,"Expected",2  )) _
            , new_DicOf(Array("No",3 ,"Num",-3.0 ,"Expected",-3 )) _
            , new_DicOf(Array("No",4 ,"Num",1.1  ,"Expected",1  )) _
            , new_DicOf(Array("No",5 ,"Num",2.5  ,"Expected",2  )) _
            , new_DicOf(Array("No",6 ,"Num",-3.9 ,"Expected",-3 )) _
            , new_DicOf(Array("No",7 ,"Num",-0.1 ,"Expected",0  )) _
            )
    For Each i In d
        num = i.Item("Num")
        e = i.Item("Expected")
        a = math_tranc(num)
        AssertEqualWithMessage a, e, "No="&i.Item("No")&", Num="&num
    Next
End Sub

'###################################################################################################
'math_fractional()
Sub Test_math_fractional
    dim a,e,d,i,num
    d = Array ( _
            new_DicOf(Array(  "No",1 ,"Num",1    ,"Expected",0    )) _
            , new_DicOf(Array("No",2 ,"Num",2.0  ,"Expected",0    )) _
            , new_DicOf(Array("No",3 ,"Num",-3.0 ,"Expected",0    )) _
            , new_DicOf(Array("No",4 ,"Num",1.1  ,"Expected",0.1  )) _
            , new_DicOf(Array("No",5 ,"Num",2.5  ,"Expected",0.5  )) _
            , new_DicOf(Array("No",6 ,"Num",-3.9 ,"Expected",-0.9 )) _
            , new_DicOf(Array("No",7 ,"Num",-0.1 ,"Expected",-0.1 )) _
            )
    For Each i In d
        num = i.Item("Num")
        e = CStr(i.Item("Expected"))
        a = CStr(math_fractional(num))
        AssertEqualWithMessage a, e, "No="&i.Item("No")&", Num="&num
    Next
End Sub

' Local Variables:
' mode: Visual-Basic
' indent-tabs-mode: nil
' End:
