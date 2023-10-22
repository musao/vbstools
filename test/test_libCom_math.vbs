' libCom.vbs: func_CM_Math* procedure test.
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
'math_min()
Sub Test_math_min
    dim a,e,d,i,num1,num2
    d = Array ( _
            new_DicWith(Array(  "No",1 ,"Num1",1     ,"Num2",1     ,"Expected","Num1")) _
            , new_DicWith(Array("No",2 ,"Num1",2     ,"Num2",1     ,"Expected","Num2")) _
            , new_DicWith(Array("No",3 ,"Num1",-3    ,"Num2",-2    ,"Expected","Num1")) _
            , new_DicWith(Array("No",4 ,"Num1",1     ,"Num2",-2    ,"Expected","Num2")) _
            , new_DicWith(Array("No",5 ,"Num1",0.2   ,"Num2",0.3   ,"Expected","Num1")) _
            , new_DicWith(Array("No",6 ,"Num1",0.1   ,"Num2",-0.04 ,"Expected","Num2")) _
            , new_DicWith(Array("No",7 ,"Num1",-0.015,"Num2",-0.009,"Expected","Num1")) _
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
            new_DicWith(Array(  "No",1 ,"Num1",1     ,"Num2",1     ,"Expected","Num2")) _
            , new_DicWith(Array("No",2 ,"Num1",2     ,"Num2",1     ,"Expected","Num1")) _
            , new_DicWith(Array("No",3 ,"Num1",-3    ,"Num2",-2    ,"Expected","Num2")) _
            , new_DicWith(Array("No",4 ,"Num1",1     ,"Num2",-2    ,"Expected","Num1")) _
            , new_DicWith(Array("No",5 ,"Num1",0.2   ,"Num2",0.3   ,"Expected","Num2")) _
            , new_DicWith(Array("No",6 ,"Num1",0.1   ,"Num2",-0.04 ,"Expected","Num1")) _
            , new_DicWith(Array("No",7 ,"Num1",-0.015,"Num2",-0.009,"Expected","Num2")) _
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
            new_DicWith(Array(  "No",1 ,"Num",-12345.6789 ,"Place",-6 ,"Expected",0)) _
            , new_DicWith(Array("No",2 ,"Num",54545.4545 ,"Place",0   ,"Expected",54546)) _
            , new_DicWith(Array("No",3 ,"Num",10101.0101 ,"Place",4   ,"Expected",10101.0101)) _
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
            new_DicWith(Array(  "No",1 ,"Num",-12345.6789 ,"Place",-6 ,"Expected",0)) _
            , new_DicWith(Array("No",2 ,"Num",54545.4545 ,"Place",0   ,"Expected",54545)) _
            , new_DicWith(Array("No",3 ,"Num",10101.0101 ,"Place",4   ,"Expected",10101.0101)) _
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
            new_DicWith(Array(  "No",1 ,"Num",-12345.6789 ,"Place",-6 ,"Expected",0)) _
            , new_DicWith(Array("No",2 ,"Num",54545.4545 ,"Place",0   ,"Expected",54545)) _
            , new_DicWith(Array("No",3 ,"Num",10101.0101 ,"Place",4   ,"Expected",10101.0101)) _
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
'func_MathRound()
Sub Test_func_MathRound_d0
    dim a,e,d,i,dn,dp,dt
    dt = 0
    dn = 12345.6789
    d = Array( _
                   new_DicWith( Array("No", 1, "Place", -6, "Expected", 0) ) _
                   , new_DicWith( Array("No", 2, "Place", -5, "Expected", 0) ) _
                   , new_DicWith( Array("No", 3, "Place", -4, "Expected", 10000) ) _
                   , new_DicWith( Array("No", 4, "Place", -1, "Expected", 12340) ) _
                   , new_DicWith( Array("No", 5, "Place", 0, "Expected", 12345) ) _
                   , new_DicWith( Array("No", 6, "Place", 1, "Expected", 12345.6) ) _
                   , new_DicWith( Array("No", 7, "Place", 4, "Expected", 12345.6789) ) _
                   , new_DicWith( Array("No", 8, "Place", 5, "Expected", 12345.6789) ) _
                )
    For Each i In d
        dp = i.Item("Place")
        e = i.Item("Expected")
        a = func_MathRound(dn,dp,dt)
        AssertEqualWithMessage e, a, "PositiveNumber No = "&i.Item("No")& ", dn = "&dn&", dp = "&dp&", dt = "&dt
    Next

    dn = -1 * dn
    For Each i In d
        dp = i.Item("Place")
        e = -1 * i.Item("Expected")
        a = func_MathRound(dn,dp,dt)
        AssertEqualWithMessage e, a, "NegativeNumber No = "&i.Item("No")& ", dn = "&dn&", dp = "&dp&", dt = "&dt
    Next
End Sub
Sub Test_func_MathRound_d5
    dim a,e,d,i,dn,dp,dt
    dt = 5
    dn = 54545.4545
    d = Array( _
                   new_DicWith( Array("No", 1, "Place", -6, "Expected", 0) ) _
                   , new_DicWith( Array("No", 2, "Place", -5, "Expected", 100000) ) _
                   , new_DicWith( Array("No", 3, "Place", -4, "Expected", 50000) ) _
                   , new_DicWith( Array("No", 4, "Place", -1, "Expected", 54550) ) _
                   , new_DicWith( Array("No", 5, "Place", 0, "Expected", 54545) ) _
                   , new_DicWith( Array("No", 6, "Place", 1, "Expected", 54545.5) ) _
                   , new_DicWith( Array("No", 7, "Place", 4, "Expected", 54545.4545) ) _
                   , new_DicWith( Array("No", 8, "Place", 5, "Expected", 54545.4545) ) _
                )
    For Each i In d
        dp = i.Item("Place")
        e = i.Item("Expected")
        a = func_MathRound(dn,dp,dt)
        AssertEqualWithMessage e, a, "PositiveNumber No = "&i.Item("No")& ", dn = "&dn&", dp = "&dp&", dt = "&dt
        If dp>=0 Then AssertEqualWithMessage Round(dn, dp), a, "PositiveNumber Regression No = "&i.Item("No")& ", dn = "&dn&", dp = "&dp&", dt = "&dt
    Next

    dn = -1 * dn
    For Each i In d
        dp = i.Item("Place")
        e = -1 * i.Item("Expected")
        a = func_MathRound(dn,dp,dt)
        AssertEqualWithMessage e, a, "NegativeNumber No = "&i.Item("No")& ", dn = "&dn&", dp = "&dp&", dt = "&dt
        If dp>=0 Then AssertEqualWithMessage Round(dn, dp), a, "NegativeNumber Regression No = "&i.Item("No")& ", dn = "&dn&", dp = "&dp&", dt = "&dt
    Next
End Sub
Sub Test_func_MathRound_d9
    dim a,e,d,i,dn,dp,dt
    dt = 9
    dn = 10101.0101
    d = Array( _
                   new_DicWith( Array("No", 1, "Place", -6, "Expected", 0) ) _
                   , new_DicWith( Array("No", 2, "Place", -5, "Expected", 100000) ) _
                   , new_DicWith( Array("No", 3, "Place", -4, "Expected", 10000) ) _
                   , new_DicWith( Array("No", 4, "Place", -1, "Expected", 10110) ) _
                   , new_DicWith( Array("No", 5, "Place", 0, "Expected", 10101) ) _
                   , new_DicWith( Array("No", 6, "Place", 1, "Expected", 10101.1) ) _
                   , new_DicWith( Array("No", 7, "Place", 4, "Expected", 10101.0101) ) _
                   , new_DicWith( Array("No", 8, "Place", 5, "Expected", 10101.0101) ) _
                )
    For Each i In d
        dp = i.Item("Place")
        e = i.Item("Expected")
        a = func_MathRound(dn,dp,dt)
        AssertEqualWithMessage e, a, "PositiveNumber No = "&i.Item("No")& ", dn = "&dn&", dp = "&dp&", dt = "&dt
    Next

    dn = -1 * dn
    For Each i In d
        dp = i.Item("Place")
        e = -1 * i.Item("Expected")
        a = func_MathRound(dn,dp,dt)
        AssertEqualWithMessage e, a, "NegativeNumber PositiveNumber No = "&i.Item("No")& ", dn = "&dn&", dp = "&dp&", dt = "&dt
    Next
End Sub

' Local Variables:
' mode: Visual-Basic
' indent-tabs-mode: nil
' End: