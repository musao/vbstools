' ReturnValue.vbs: test.
' @import ../../lib/com/clsAdptFile.vbs
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
'ReturnValue
Sub Test_ReturnValue
    Dim a : Set a = new ReturnValue
    AssertEqual 0, VarType(a)
    AssertEqual "ReturnValue", TypeName(a)
End Sub

'###################################################################################################
'ReturnValue.returnValue()
Sub Test_ReturnValue_returnValue
    Dim data
    data = Array( _
        new_DicOf(Array(  "Data", Empty                , "Expect", "EqualData")) _
        , new_DicOf(Array("Data", Null                 , "Expect", "EqualData")) _
        , new_DicOf(Array("Data", new_Dic()            , "Expect", "EqualData")) _
        , new_DicOf(Array("Data", Array("a",2)         , "Expect", "EqualData")) _
        , new_DicOf(Array("Data", CInt(1)              , "Expect", "EqualData")) _
        , new_DicOf(Array("Data", CLng(999999)         , "Expect", "EqualData")) _
        , new_DicOf(Array("Data", CSng(10.1)           , "Expect", "EqualData")) _
        , new_DicOf(Array("Data", CDbl(1234.567890123) , "Expect", "EqualData")) _
        , new_DicOf(Array("Data", CCur("\1,000")       , "Expect", "EqualData")) _
        , new_DicOf(Array("Data", True                 , "Expect", "EqualData")) _
        , new_DicOf(Array("Data", CByte(0)             , "Expect", "EqualData")) _
        , new_DicOf(Array("Data", vbNullString         , "Expect", "EqualData")) _
        , new_DicOf(Array("Data", "abc"                , "Expect", "EqualData")) _
        , new_DicOf(Array("Data", "1.2"                , "Expect", "EqualData")) _
        , new_DicOf(Array("Data", "192.168.11.52"      , "Expect", "EqualData")) _
        , new_DicOf(Array("Data", "2024/01/03"         , "Expect", "EqualData")) _
        , new_DicOf(Array("Data", "５０"               , "Expect", "EqualData")) _
        , new_DicOf(Array("Data", "漢字"               , "Expect", "EqualData")) _
        )
    Dim o : Set o = new ReturnValue
    
    Dim ub : ub = Ubound(data)
    Dim i,j,d,e,a
    For i=0 To ub
        cf_bind d, data(i).Item("Data")
        If cf_isSame(data(i).Item("Expect"), "EqualData") Then
            cf_bind e, d
        Else
            cf_bind e, data(i).Item("Expect")
        End if
        if IsObject(d) Then Set o.returnValue=d Else o.returnValue=d
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
'ReturnValue.isErr()
Sub Test_ReturnValue_isErr_Normal
    Dim o : Set o = new ReturnValue
    o.setValue "abc"

    Dim e,a
    e = False
    a = o.isErr()
    AssertEqual e,a
End Sub
Sub Test_ReturnValue_isErr_Err
    Dim o : Set o = new ReturnValue
    On Error Resume Next
    Dim ern,ers,erd : ern=9999:ers="エラー":erd="test_ReturnValue.vbsのエラー"
    Err.Raise ern, ers, erd
    o.setValue "あいう"
    On Error Goto 0

    Dim e,a
    e = True
    a = o.isErr()
    AssertEqual e,a
End Sub
Sub Test_ReturnValue_isErr_Initial
    Dim o : Set o = new ReturnValue

    Dim e,a
    e = Empty
    a = o.isErr()
    AssertEqual e,a
End Sub

'###################################################################################################
'ReturnValue.setValue()
Sub Test_ReturnValue_setValue_Normal
    Dim o : Set o = new ReturnValue
    Dim d : d = "abc"

    Dim e,a
    e = TypeName(o)
    a = TypeName(o.setValue(d))
    AssertEqualWithMessage e,a,"TypeName"

    e = 0
    a = Err.Number
    AssertEqualWithMessage e,a,"Err.Number"

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
Sub Test_ReturnValue_setValue_Err
    Dim o : Set o = new ReturnValue
    Dim d : Set d = new_Dic()

    Dim e,a
    On Error Resume Next
    Dim ern,ers,erd : ern=9999:ers="エラー":erd="test_ReturnValue.vbsのエラー"
    Err.Raise ern, ers, erd
    e = TypeName(o)
    a = TypeName(o.setValue(d))
    AssertEqualWithMessage e,a,"TypeName"

    e = 0
    a = Err.Number
    AssertEqualWithMessage e,a,"Err.Number"
    On Error Goto 0

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
Sub Test_ReturnValue_setValue_ErrToNormal
    Dim o : Set o = new ReturnValue
    Dim d : Set d = new_Dic()

    Dim e,a
    On Error Resume Next
    Dim ern,ers,erd : ern=9999:ers="エラー":erd="test_ReturnValue.vbsのエラー"
    Err.Raise ern, ers, erd
    o.setValue(d)

    e = 0
    a = Err.Number
    AssertEqualWithMessage e,a,"Err.Number"
    On Error Goto 0

    d = vbNullString
    o.setValue(d)

    e = 0
    a = Err.Number
    AssertEqualWithMessage e,a,"Err.Number"

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
Sub Test_ReturnValue_setValue_Initial
    Dim o : Set o = new ReturnValue

    Dim e,a
    e = Empty
    a = o.returnValue
    AssertEqualWithMessage e,a,"returnValue"

    e = Empty
    a = o.isErr()
    AssertEqualWithMessage e,a,"isErr()"

    Set e = Nothing
    Set a = o.getErr()
    AssertSameWithMessage e,a,"getErr()"
End Sub

'###################################################################################################
'ReturnValue.setValueByState()
Sub Test_ReturnValue_setValueByState_Normal_Noerr
    Dim o : Set o = new ReturnValue
    Dim normal : normal = "normal"
    Dim abnormal : abnormal = "abnormal"

    Dim e,a
    e = TypeName(o)
    a = TypeName(o.setValueByState(normal,abnormal))
    AssertEqualWithMessage e,a,"TypeName"

    e = 0
    a = Err.Number
    AssertEqualWithMessage e,a,"Err.Number"

    e = normal
    a = o.returnValue
    AssertEqualWithMessage e,a,"returnValue"

    e = False
    a = o.isErr()
    AssertEqualWithMessage e,a,"isErr()"

    Set e = Nothing
    Set a = o.getErr()
    AssertSameWithMessage e,a,"getErr()"
End Sub
Sub Test_ReturnValue_setValueByState_Normal_Err
    Dim o : Set o = new ReturnValue
    Dim normal : normal = "normal"
    Dim abnormal : abnormal = "abnormal"

    On Error Resume Next
    Dim ern,ers,erd : ern=9999:ers="エラー":erd="test_ReturnValue.vbsのエラー"
    Err.Raise ern, ers, erd
    o.setValueByState normal,abnormal

    e = 0
    a = Err.Number
    AssertEqualWithMessage e,a,"Err.Number"
    On Error Goto 0

    Dim e,a
    e = abnormal
    a = o.returnValue
    AssertEqualWithMessage e,a,"returnValue"

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

'###################################################################################################
'ReturnValue.toString()
Sub Test_ReturnValue_toString
    Dim o : Set o = new ReturnValue

    Dim e,a
    e = "<ReturnValue>[returnValue:<Empty>,isErr:<Empty>,getErr:<Nothing>]"
    a = o.toString()
    AssertEqualWithMessage e,a,"Initial"

    Dim d : d = "abc"
    o.setValue(d)
    e = "<ReturnValue>[returnValue:<String>"""&d&""",isErr:<Boolean>False,getErr:<Nothing>]"
    a = o.toString()
    AssertEqualWithMessage e,a,"Normal"

    On Error Resume Next
    Dim ern,ers,erd : ern=9999:ers="エラー":erd="test_ReturnValue.vbsのエラー"
    Err.Raise ern, ers, erd
    o.setValue(d)
    On Error Goto 0
    e = "<ReturnValue>[returnValue:<String>""abc"",isErr:<Boolean>True,getErr:<Err>{<String>""Number""=><Long>"&ern&",<String>""Description""=><String>"""&erd&""",<String>""Source""=><String>"""&ers&"""}]"
    a = o.toString()
    AssertEqualWithMessage e,a,"Err"

End Sub


' Local Variables:
' mode: Visual-Basic
' indent-tabs-mode: nil
' End:
