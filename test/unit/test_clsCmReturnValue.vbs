' clsCmReturnValue.vbs: test.
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
'clsCmReturnValue
Sub Test_clsCmReturnValue
    Dim a : Set a = new clsCmReturnValue
    AssertEqual 9, VarType(a)
    AssertEqual "clsCmReturnValue", TypeName(a)
End Sub

'###################################################################################################
'clsCmBroker.returnValue()
Sub Test_clsCmReturnValue_returnValue
    Dim data
    data = Array( _
        new_DicWith(Array(  "Data", Empty                , "Expect", "EqualData")) _
        , new_DicWith(Array("Data", Null                 , "Expect", "EqualData")) _
        , new_DicWith(Array("Data", new_Dic()            , "Expect", "EqualData")) _
        , new_DicWith(Array("Data", Array("a",2)         , "Expect", "EqualData")) _
        , new_DicWith(Array("Data", CInt(1)              , "Expect", "EqualData")) _
        , new_DicWith(Array("Data", CLng(999999)         , "Expect", "EqualData")) _
        , new_DicWith(Array("Data", CSng(10.1)           , "Expect", "EqualData")) _
        , new_DicWith(Array("Data", CDbl(1234.567890123) , "Expect", "EqualData")) _
        , new_DicWith(Array("Data", CCur("\1,000")       , "Expect", "EqualData")) _
        , new_DicWith(Array("Data", True                 , "Expect", "EqualData")) _
        , new_DicWith(Array("Data", CByte(0)             , "Expect", "EqualData")) _
        , new_DicWith(Array("Data", vbNullString         , "Expect", "EqualData")) _
        , new_DicWith(Array("Data", "abc"                , "Expect", "EqualData")) _
        , new_DicWith(Array("Data", "1.2"                , "Expect", "EqualData")) _
        , new_DicWith(Array("Data", "192.168.11.52"      , "Expect", "EqualData")) _
        , new_DicWith(Array("Data", "2024/01/03"         , "Expect", "EqualData")) _
        , new_DicWith(Array("Data", "５０"               , "Expect", "EqualData")) _
        , new_DicWith(Array("Data", "漢字"               , "Expect", "EqualData")) _
        )
    Dim o : Set o = new clsCmReturnValue
    
    Dim ub : ub = Ubound(data)
    Dim i,j,d,e,a
    For i=0 To ub
        cf_bind d, data(i).Item("Data")
        If cf_isSame(data(i).Item("Expect"), "EqualData") Then
            cf_bind e, d
        Else
            cf_bind e, data(i).Item("Expect")
        End if
        if IsObject(d) Then Set o.returnValue = d Else o.returnValue=d
        cf_bind a, o.returnValue

        If IsObject(e) and IsObject(a) Then
            AssertSameWithMessage e, a, "No"&i&" Data="&cf_toString(d)&" Expect="&cf_toString(e)
        ElseIf Not(IsObject(e)) and Not(IsObject(a)) Then
            If IsArray(e) Then
                AssertEqualWithMessage Ubound(e), Ubound(a), "No"&i&" Data="&Ubound(d)&" Expect="&Ubound(e)
                For j=0 To Ubound(e)
                    AssertEqualWithMessage e(j), a(j), "No"&i&"-"&j&" Data="&cf_toString(d(j))&" Expect="&cf_toString(e(j))
                Next
            Else
                AssertEqualWithMessage e, a, "No"&i&" Data="&cf_toString(d)&" Expect="&cf_toString(e)
            End If
        Else
            AssertFailWithMessage "No"&i&" Data="&cf_toString(d)&" Expect="&cf_toString(e)&" Actual="&cf_toString(a)
        End If
    Next
End Sub

'###################################################################################################
'clsCmBroker.isErr()
Sub Test_clsCmReturnValue_isErr_Normal
    Dim o : Set o = new clsCmReturnValue
    o.setValue "abc"

    Dim e,a
    e = False
    a = o.isErr()
    AssertEqual e,a
End Sub
Sub Test_clsCmReturnValue_isErr_Err
    Dim o : Set o = new clsCmReturnValue
    On Error Resume Next
    Dim ern,ers,erd : ern=9999:ers="エラー":erd="test_clsCmReturnValue.vbsのエラー"
    Err.Raise ern, ers, erd
    o.setValue "あいう"
    On Error Goto 0

    Dim e,a
    e = True
    a = o.isErr()
    AssertEqual e,a
End Sub
Sub Test_clsCmReturnValue_isErr_Initial
    Dim o : Set o = new clsCmReturnValue

    Dim e,a
    e = Empty
    a = o.isErr()
    AssertEqual e,a
End Sub

'###################################################################################################
'clsCmBroker.setValue()
Sub Test_clsCmReturnValue_setValue_Normal
    Dim o : Set o = new clsCmReturnValue
    Dim d : d = "abc"

    Dim e,a
    e = TypeName(o)
    a = TypeName(o.setValue(d))
    AssertEqualWithMessage e,a,"TypeName"

    e = d
    a = o.returnValue
    AssertEqualWithMessage e,a,"returnValue"

    e = False
    a = o.isErr()
    AssertEqualWithMessage e,a,"isErr()"

    Set e = Nothing
    Set a = o.getErr()
    AssertSameWithMessage e,a,"getErr()"
End Sub
Sub Test_clsCmReturnValue_setValue_Err
    Dim o : Set o = new clsCmReturnValue
    Dim d : Set d = new_Dic()

    Dim e,a
    On Error Resume Next
    Dim ern,ers,erd : ern=9999:ers="エラー":erd="test_clsCmReturnValue.vbsのエラー"
    Err.Raise ern, ers, erd
    e = TypeName(o)
    a = TypeName(o.setValue(d))
    On Error Goto 0
    AssertEqualWithMessage e,a,"TypeName"

    Set e = d
    Set a = o.returnValue
    AssertSameWithMessage e,a,"returnValue"

    e = True
    a = o.isErr()
    AssertEqualWithMessage e,a,"isErr()"

    Dim er : Set er = o.getErr()
    e = True
    a = cf_isAvailableObject(er)
    AssertEqualWithMessage e,a,"cf_isAvailableObject(getErr())"

    e = ern
    a = er.Item("Number")
    AssertEqualWithMessage e,a,"isErr().Item('Number')"

    e = erd
    a = er.Item("Description")
    AssertEqualWithMessage e,a,"isErr().Item('NuDescriptionmber')"

    e = ers
    a = er.Item("Source")
    AssertEqualWithMessage e,a,"isErr().Item('Source')"
End Sub
Sub Test_clsCmReturnValue_setValue_ErrToNormal
    Dim o : Set o = new clsCmReturnValue
    Dim d : Set d = new_Dic()

    Dim e,a
    On Error Resume Next
    Dim ern,ers,erd : ern=9999:ers="エラー":erd="test_clsCmReturnValue.vbsのエラー"
    Err.Raise ern, ers, erd
    o.setValue(d)
    On Error Goto 0

    d = vbNullString
    o.setValue(d)

    e = d
    a = o.returnValue
    AssertEqualWithMessage e,a,"returnValue"

    e = False
    a = o.isErr()
    AssertEqualWithMessage e,a,"isErr()"

    Set e = Nothing
    Set a = o.getErr()
    AssertSameWithMessage e,a,"getErr()"
End Sub
Sub Test_clsCmReturnValue_setValue_Initial
    Dim o : Set o = new clsCmReturnValue

    Dim e,a
    Set e = Nothing
    Set a = o.returnValue
    AssertSameWithMessage e,a,"returnValue"

    e = Empty
    a = o.isErr()
    AssertEqualWithMessage e,a,"isErr()"

    Set e = Nothing
    Set a = o.getErr()
    AssertSameWithMessage e,a,"getErr()"
End Sub

'###################################################################################################
'clsCmBroker.toString()
Sub Test_clsCmReturnValue_toString
    Dim o : Set o = new clsCmReturnValue

    Dim e,a
    e = "<clsCmReturnValue>[returnValue:<Nothing>,isErr:<Empty>,getErr:<Nothing>]"
    a = o.toString()
    AssertEqualWithMessage e,a,"TypInitial"

    Dim d : d = "abc"
    o.setValue(d)
    e = "<clsCmReturnValue>[returnValue:<String>"""&d&""",isErr:<Boolean>False,getErr:<Nothing>]"
    a = o.toString()
    AssertEqualWithMessage e,a,"Normal"

    On Error Resume Next
    Dim ern,ers,erd : ern=9999:ers="エラー":erd="test_clsCmReturnValue.vbsのエラー"
    Err.Raise ern, ers, erd
    o.setValue(d)
    On Error Goto 0
    e = "<clsCmReturnValue>[returnValue:<String>""abc"",isErr:<Boolean>True,getErr:<Err>{<String>""Number""=><Long>"&ern&",<String>""Description""=><String>"""&erd&""",<String>""Source""=><String>"""&ers&"""}]"
    a = o.toString()
    AssertEqualWithMessage e,a,"Err"

End Sub


' Local Variables:
' mode: Visual-Basic
' indent-tabs-mode: nil
' End:
