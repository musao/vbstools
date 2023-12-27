' libCom.vbs: cf_* procedure test.
' @import ../../lib/clsCmArray.vbs
' @import ../../lib/clsCmBroker.vbs
' @import ../../lib/clsCmBufferedReader.vbs
' @import ../../lib/clsCmBufferedWriter.vbs
' @import ../../lib/clsCmCalendar.vbs
' @import ../../lib/clsCmCharacterType.vbs
' @import ../../lib/clsCmCssGenerator.vbs
' @import ../../lib/clsCmHtmlGenerator.vbs
' @import ../../lib/clsCompareExcel.vbs
' @import ../../lib/libCom.vbs

Option Explicit

'###################################################################################################
'cf_bind()
Sub Test_cf_bind_Value
    Dim v
    cf_bind v, "Hello world."
    
    AssertEqual "Hello world.", v
End Sub
Sub Test_cf_bind_Object
    Dim v
    Dim obj: Set obj = CreateObject("Scripting.Dictionary")
    cf_bind v, obj
    
    AssertSame obj, v
End Sub

'###################################################################################################
'cf_bindAt()
Sub Test_cf_bindAt_Value
    Dim obj : Set obj = CreateObject("Scripting.Dictionary")
    cf_bindAt obj, "Value", "Hello world."
    
    AssertEqual "Hello world.", obj.Item("Value")
End Sub
Sub Test_cf_bindAt_Object
    Dim obj : Set obj = CreateObject("Scripting.Dictionary")
    cf_bindAt obj, "Object", Nothing
    
    AssertSame Nothing, obj.Item("Object")
End Sub

'###################################################################################################
'cf_push()
Sub Test_cf_push_Available
    Redim a(0)
    cf_push a, "NewValue"
    
    AssertEqual 1, Ubound(a)
    AssertEqual Empty, a(0)
    AssertEqual "NewValue", a(1)
End Sub
Sub Test_cf_push_NotAvailable
    Dim a
    cf_push a, "NewValue"
    
    AssertEqual 0, Ubound(a)
    AssertEqual "NewValue", a(0)
End Sub

'###################################################################################################
'cf_pushMulti()
Sub Test_cf_pushMulti_AddIsArray_ArrAvailable
    Dim a,d,e
    Redim a(0)
    d = Array(1,2)
    e = Array(Empty,1,2)
    cf_pushMulti a, d
    
    assertAllElements e, a
End Sub
Sub Test_cf_pushMulti_AddIsArray_ArrNotAvailable
    Dim a,d,e
    d = Array(1,2)
    e = Array(1,2)
    cf_pushMulti a, d
    
    assertAllElements e, a
End Sub
Sub Test_cf_pushMulti_AddIsArray_ArrNotAvailable2
    Dim a(),d,e
    d = Array(1,2)
    e = Array(1,2)
    cf_pushMulti a, d
    
    assertAllElements e, a
End Sub
Sub Test_cf_pushMulti_AddIsZeroArray
    Dim a,d(),e
    Redim a(0)
    e = Array(Empty)
    cf_pushMulti a, d
    
    assertAllElements e, a
End Sub
Sub Test_cf_pushMulti_AddIsNotArray_ArrAvailable
    Dim a,d,e
    Redim a(0)
    d = "a"
    e = Array(Empty,"a")
    cf_pushMulti a, d
    
    assertAllElements e, a
End Sub
Sub Test_cf_pushMulti_AddIsNotArray_ArrNotAvailable
    Dim a,d,e
    d = "a"
    e = Array("a")
    cf_pushMulti a, d
    
    assertAllElements e, a
End Sub

'###################################################################################################
'cf_tryCatch()
Sub Test_cf_tryCatch_TryOnly_Normal
    Dim oRet : Set oRet = cf_tryCatch(new_Func("a=>1/a"), 2, Nothing, Empty)
    
    AssertEqual 0, Err.Number
    AssertEqual True, oRet.Item("Result")
    AssertEqual 1/2, oRet.Item("Return")
    AssertSame Nothing, oRet.Item("Err")
End Sub
Sub Test_cf_tryCatch_TryAndCatch_Normal
    Dim oRet : Set oRet = cf_tryCatch(new_Func("a=>1/a"), 2, new_Func("a=>a"), Nothing)
    
    AssertEqual 0, Err.Number
    AssertEqual True, oRet.Item("Result")
    AssertEqual 1/2, oRet.Item("Return")
    AssertSame Nothing, oRet.Item("Err")
End Sub
Sub Test_cf_tryCatch_TryAndFinary_Normal
    Dim oRet : Set oRet = cf_tryCatch(new_Func("a=>1/a"), 2, Empty, new_Func("r=>1/2+r"))
    
    AssertEqual 0, Err.Number
    AssertEqual True, oRet.Item("Result")
    AssertEqual 1/2+1/2, oRet.Item("Return")
    AssertSame Nothing, oRet.Item("Err")
End Sub
Sub Test_cf_tryCatch_TryAndFinary_Normal_FinaryErr
    On Error Resume Next
    Dim oRet : Set oRet = cf_tryCatch(new_Func("a=>1/a"), 2, Empty, new_Func("r=>r(0)"))
    
    AssertEqual 13, Err.Number
    AssertEqual "å^Ç™àÍívÇµÇ‹ÇπÇÒÅB", Err.Description
    AssertEqual Empty, oRet
End Sub
Sub Test_cf_tryCatch_TryAndCatchAndFinary_Normal
    Dim oRet : Set oRet = cf_tryCatch(new_Func("a=>1/a"), 2, new_Func("a=>a"), new_Func("r=>1/2+r"))
    
    AssertEqual 0, Err.Number
    AssertEqual True, oRet.Item("Result")
    AssertEqual 1/2+1/2, oRet.Item("Return")
    AssertSame Nothing, oRet.Item("Err")
End Sub
Sub Test_cf_tryCatch_TryAndCatchAndFinary_Normal_FinaryErr
    On Error Resume Next
    Dim oRet : Set oRet = cf_tryCatch(new_Func("a=>1/a"), 2, new_Func("a=>a"), new_Func("r=>r(0)"))
    
    AssertEqual 13, Err.Number
    AssertEqual "å^Ç™àÍívÇµÇ‹ÇπÇÒÅB", Err.Description
    AssertEqual Empty, oRet
End Sub
Sub Test_cf_tryCatch_TryOnly_Err
    Dim oRet : Set oRet = cf_tryCatch(new_Func("a=>1/a"), 0, Empty, Empty)
    
    AssertEqual 0, Err.Number
    AssertEqual False, oRet.Item("Result")
    AssertEqual Empty, oRet.Item("Return")
    AssertEqual 11, oRet.Item("Err").Item("Number")
    AssertEqual "0 Ç≈èúéZÇµÇ‹ÇµÇΩÅB", oRet.Item("Err").Item("Description")
    AssertEqual "Microsoft VBScript é¿çséûÉGÉâÅ[", oRet.Item("Err").Item("Source")
End Sub
Sub Test_cf_tryCatch_TryAndCatch_Err
    Dim oRet : Set oRet = cf_tryCatch(new_Func("a=>1/a"), 0, new_Func("a=>a"), Empty)
    
    AssertEqual 0, Err.Number
    AssertEqual False, oRet.Item("Result")
    AssertEqual 0, oRet.Item("Return")
    AssertEqual 11, oRet.Item("Err").Item("Number")
    AssertEqual "0 Ç≈èúéZÇµÇ‹ÇµÇΩÅB", oRet.Item("Err").Item("Description")
    AssertEqual "Microsoft VBScript é¿çséûÉGÉâÅ[", oRet.Item("Err").Item("Source")
End Sub
Sub Test_cf_tryCatch_TryAndCatch_Err_CatchErr
    On Error Resume Next
    Dim oRet : Set oRet = cf_tryCatch(new_Func("a=>1/a"), 0, new_Func("a=>a(0)"), Empty)
    
    AssertEqual 13, Err.Number
    AssertEqual "å^Ç™àÍívÇµÇ‹ÇπÇÒÅB", Err.Description
    AssertEqual Empty, oRet
End Sub
Sub Test_cf_tryCatch_TryAndFinary_Err
    Dim oRet : Set oRet = cf_tryCatch(new_Func("a=>1/a"), 0, Nothing, new_Func("r=>2"))
    
    AssertEqual 0, Err.Number
    AssertEqual False, oRet.Item("Result")
    AssertEqual 2, oRet.Item("Return")
    AssertEqual 11, oRet.Item("Err").Item("Number")
    AssertEqual "0 Ç≈èúéZÇµÇ‹ÇµÇΩÅB", oRet.Item("Err").Item("Description")
    AssertEqual "Microsoft VBScript é¿çséûÉGÉâÅ[", oRet.Item("Err").Item("Source")
End Sub
Sub Test_cf_tryCatch_TryAndFinary_Err_FinaryErr
    On Error Resume Next
    Dim oRet : Set oRet = cf_tryCatch(new_Func("a=>1/a"), 0, Nothing, new_Func("r=>r(0)"))
    
    AssertEqual 13, Err.Number
    AssertEqual "å^Ç™àÍívÇµÇ‹ÇπÇÒÅB", Err.Description
    AssertEqual Empty, oRet
End Sub
Sub Test_cf_tryCatch_TryAndCatchAndFinary_Err
    Dim oRet : Set oRet = cf_tryCatch(new_Func("a=>1/a"), 0, new_Func("a=>a"), new_Func("r=>2"))
    
    AssertEqual 0, Err.Number
    AssertEqual False, oRet.Item("Result")
    AssertEqual 2, oRet.Item("Return")
    AssertEqual 11, oRet.Item("Err").Item("Number")
    AssertEqual "0 Ç≈èúéZÇµÇ‹ÇµÇΩÅB", oRet.Item("Err").Item("Description")
    AssertEqual "Microsoft VBScript é¿çséûÉGÉâÅ[", oRet.Item("Err").Item("Source")
End Sub
Sub Test_cf_tryCatch_TryAndCatchAndFinary_Err_CatchErr
    On Error Resume Next
    Dim oRet : Set oRet = cf_tryCatch(new_Func("a=>1/a"), 0, new_Func("a=>a(0)"), new_Func("r=>2"))
    
    AssertEqual 13, Err.Number
    AssertEqual "å^Ç™àÍívÇµÇ‹ÇπÇÒÅB", Err.Description
    AssertEqual Empty, oRet
End Sub
Sub Test_cf_tryCatch_TryAndCatchAndFinary_Err_FinaryErr
    On Error Resume Next
    Dim oRet : Set oRet = cf_tryCatch(new_Func("a=>1/a"), 0, new_Func("a=>a"), new_Func("r=>r(0)"))
    
    AssertEqual 13, Err.Number
    AssertEqual "å^Ç™àÍívÇµÇ‹ÇπÇÒÅB", Err.Description
    AssertEqual Empty, oRet
End Sub
Sub Test_cf_tryCatch_TryOnly_ArgEmpty
    Dim oRet : Set oRet = cf_tryCatch(new_Func("()=>1/2"), Empty, Nothing, Empty)
    
    AssertEqual 0, Err.Number
    AssertEqual True, oRet.Item("Result")
    AssertEqual 1/2, oRet.Item("Return")
    AssertSame Nothing, oRet.Item("Err")
End Sub
Sub Test_cf_tryCatch_TryAndCatch_ArgEmpty
    Dim oRet : Set oRet = cf_tryCatch(new_Func("=>1/0"), Empty, new_Func("=>1/2"), Nothing)
    
    AssertEqual 0, Err.Number
    AssertEqual False, oRet.Item("Result")
    AssertEqual 1/2, oRet.Item("Return")
    AssertEqual 11, oRet.Item("Err").Item("Number")
    AssertEqual "0 Ç≈èúéZÇµÇ‹ÇµÇΩÅB", oRet.Item("Err").Item("Description")
    AssertEqual "Microsoft VBScript é¿çséûÉGÉâÅ[", oRet.Item("Err").Item("Source")
End Sub

'###################################################################################################
'cf_isSame()
Sub Test_cf_isSame_OvsO_Same
    Dim a,da,db,e
    Set da = new_Dic()
    Set db = da
    e = True
    a = cf_isSame(da,db)

    AssertEqual e,a
End Sub
Sub Test_cf_isSame_OvsO_NotSame
    Dim a,da,db,e
    Set da = new_Dic()
    Set db = new_Dic()
    e = False
    a = cf_isSame(da,db)

    AssertEqual e,a
End Sub
Sub Test_cf_isSame_OvsV_NotSame
    Dim a,da,db,e
    Set da = new_Dic()
    db = "a"
    e = False
    a = cf_isSame(da,db)

    AssertEqual e,a
End Sub
Sub Test_cf_isSame_VvsO_NotSame
    Dim a,da,db,e
    da = 5
    Set db = new_Dic()
    e = False
    a = cf_isSame(da,db)

    AssertEqual e,a
End Sub
Sub Test_cf_isSame_VvsV_and_SvsS_Same
    Dim a,da,db,e
    da = "a"
    db = "a"
    e = True
    a = cf_isSame(da,db)

    AssertEqual e,a
End Sub
Sub Test_cf_isSame_VvsV_and_SvsS_NotSame
    Dim a,da,db,e
    da = "a"
    db = "A"
    e = False
    a = cf_isSame(da,db)

    AssertEqual e,a
End Sub
Sub Test_cf_isSame_VvsV_and_NvsN_Same
    Dim a,da,db,e
    da = 9
    db = 9
    e = True
    a = cf_isSame(da,db)

    AssertEqual e,a
End Sub
Sub Test_cf_isSame_VvsV_and_NvsN_NotSame
    Dim a,da,db,e
    da = 8
    db = 9
    e = False
    a = cf_isSame(da,db)

    AssertEqual e,a
End Sub
Sub Test_cf_isSame_VvsV_and_SvsN_NotSame
    Dim a,da,db,e
    da = "9"
    db = 9
    e = False
    a = cf_isSame(da,db)

    AssertEqual e,a
End Sub
Sub Test_cf_isSame_VvsV_and_NvsS_NotSame
    Dim a,da,db,e
    da = 500
    db = "abc"
    e = False
    a = cf_isSame(da,db)

    AssertEqual e,a
End Sub

'###################################################################################################
'cf_isValid()
Sub Test_cf_isValid_Object_Valid
    Dim a,d,e
    Set d = new_Dic()
    e = True
    a = cf_isValid(d)
    AssertEqual e,a
End Sub
Sub Test_cf_isValid_Object_InValid
    Dim a,d,e
    Set d = Nothing
    e = False
    a = cf_isValid(d)
    AssertEqual e,a
End Sub
Sub Test_cf_isValid_Array_Valid
    Dim a,d,e
    d = Array(1,"2")
    e = True
    a = cf_isValid(d)
    AssertEqual e,a
End Sub
Sub Test_cf_isValid_Array_InValid_1
    Dim a,d,e
    d = Array()
    e = False
    a = cf_isValid(d)
    AssertEqual e,a
End Sub
Sub Test_cf_isValid_Array_InValid_2
    Dim a,e
    Dim d()
    e = False
    a = cf_isValid(d)
    AssertEqual e,a
End Sub
Sub Test_cf_isValid_Variable_Valid
    Dim a,d,e
    d = "a"
    e = True
    a = cf_isValid(d)
    AssertEqual e,a
End Sub
Sub Test_cf_isValid_Variable_InValid_Empty
    Dim a,d,e
    d = Empty
    e = False
    a = cf_isValid(d)
    AssertEqual e,a
End Sub
Sub Test_cf_isValid_Variable_InValid_Null
    Dim a,d,e
    d = Null
    e = False
    a = cf_isValid(d)
    AssertEqual e,a
End Sub
Sub Test_cf_isValid_Variable_InValid_NullString
    Dim a,d,e
    d = vbNullString
    e = False
    a = cf_isValid(d)
    AssertEqual e,a
End Sub

'###################################################################################################
'cf_isAvailableObject()
Sub Test_cf_isAvailableObject
    Dim a,d,e
    Set d = new_Dic()
    e = True
    a = cf_isAvailableObject(d)
    AssertEqual e,a
End Sub
Sub Test_cf_isAvailableObject_Nothing
    Dim a,d,e
    Set d = Nothing
    e = False
    a = cf_isAvailableObject(d)
    AssertEqual e,a
End Sub
Sub Test_cf_isAvailableObject_Variable
    Dim a,d,e
    d = Empty
    e = False
    a = cf_isAvailableObject(d)
    AssertEqual e,a
End Sub

'###################################################################################################
'cf_toString()
Sub Test_cf_toString_Empty
    Dim a,d,e
    d = Empty
    e = "<Empty>"
    a = cf_toString(d)
    AssertEqual e,a
End Sub
Sub Test_cf_toString_Null
    Dim a,d,e
    d = Null
    e = "<Null>"
    a = cf_toString(d)
    AssertEqual e,a
End Sub
Sub Test_cf_toString_Integer
    Dim a,d,e
    d = CInt(100)
    e = "<Integer>100"
    a = cf_toString(d)
    AssertEqual e,a
End Sub
Sub Test_cf_toString_Long
    Dim a,d,e
    d = CLng(99999999)
    e = "<Long>99999999"
    a = cf_toString(d)
    AssertEqual e,a
End Sub
Sub Test_cf_toString_Single
    Dim a,d,e
    d = CSng(1.23)
    e = "<Single>1.23"
    a = cf_toString(d)
    AssertEqual e,a
End Sub
Sub Test_cf_toString_Double
    Dim a,d,e
    d = CDbl(1.23)
    e = "<Double>1.23"
    a = cf_toString(d)
    AssertEqual e,a
End Sub
Sub Test_cf_toString_Currency
    Dim a,d,e
    d = CCur(100.1)
    e = "<Currency>100.1"
    a = cf_toString(d)
    AssertEqual e,a
End Sub
Sub Test_cf_toString_Date
    Dim a,d,e
    d = #2023-01-24 18:12:04#
    e = "<Date>2023/01/24 18:12:04"
    a = cf_toString(d)
    AssertEqual e,a
End Sub
Sub Test_cf_toString_String
    Dim a,d,e
    d = "foo"
    e = "<String>" & Chr(34) & d & Chr(34)
    a = cf_toString(d)
    AssertEqual e,a
End Sub
Sub Test_cf_toString_String_vbNullString
    Dim a,d,e
    d = vbNullString
    e = "<String>" & Chr(34)&Chr(34)
    a = cf_toString(d)
    AssertEqual e,a
End Sub
Sub Test_cf_toString_String_ContainsDoubleQuotes
    Dim a,d,e
    d = "foo" & Chr(34) & "bar"
    e = "<String>" & Chr(34) & Replace(d,Chr(34),Chr(34)&Chr(34)) & Chr(34)
    a = cf_toString(d)
    AssertEqual e,a
End Sub
Sub Test_cf_toString_Object_Dictionary
    Dim a,d,e
    Set d = new_DicWith(Array("foo","apple","bar",5))
    e = "<Dictionary>{<String>" & Chr(34) & "foo" & Chr(34) & "=><String>" & Chr(34) & "apple" & Chr(34) & ",<String>" & Chr(34) & "bar" & Chr(34) & "=><Integer>5}"
    a = cf_toString(d)
    AssertEqual e,a
End Sub
Sub Test_cf_toString_Object_Dictionary_Empty
    Dim a,d,e
    Set d = new_Dic()
    e = "<Dictionary>{}"
    a = cf_toString(d)
    AssertEqual e,a
End Sub
Sub Test_cf_toString_Object_Nothing
    Dim a,d,e
    Set d = Nothing
    e = "<Nothing>"
    a = cf_toString(d)
    AssertEqual e,a
End Sub
Sub Test_cf_toString_Object_Other_ShellApp
    Dim a,d,e
    Set d = new_ShellApp()
    e = "<IShellDispatch6>"
    a = cf_toString(d)
    AssertEqual e,a
End Sub
Sub Test_cf_toString_Object_Other_UserDef
    Dim a,d,e
    Set d = new_Char()
    e = "<clsCmCharacterType>"
    a = cf_toString(d)
    AssertEqual e,a
End Sub
Sub Test_cf_toString_Object_Other_UserDef_Special
    Dim a,d,e
    Set d = new_DicWith(Array("__Special__", "Test", "Key", "Value"))
    e = "<Test>{<String>" & Chr(34) & "Key" & Chr(34) & "=><String>" & Chr(34) & "Value" & Chr(34) & "}"
    a = cf_toString(d)
    AssertEqual e,a
End Sub
Sub Test_cf_toString_Boolean
    Dim a,d,e
    d = True
    e = "<Boolean>True"
    a = cf_toString(d)
    AssertEqual e,a
End Sub
Sub Test_cf_toString_Byte
    Dim a,d,e
    d = CByte(1)
    e = "<Byte>1"
    a = cf_toString(d)
    AssertEqual e,a
End Sub
Sub Test_cf_toString_Array
    Dim a,d,e
    d = Array(1,"a",Array("éQ",4),new_DicWith(Array("áX",6,7,"ÇW")))
    e = "<Array>[<Integer>1,<String>" & Chr(34) & "a" & Chr(34) & ",<Array>[<String>" & Chr(34) & "éQ" & Chr(34) & ",<Integer>4],<Dictionary>{<String>" & Chr(34) & "áX" & Chr(34) & "=><Integer>6,<Integer>7=><String>" & Chr(34) & "ÇW" & Chr(34) & "}]"
    a = cf_toString(d)
    AssertEqual e,a
End Sub
Sub Test_cf_toString_Array_Empty
    Dim a,d,e
    d = Array()
    e = "<Array>[]"
    a = cf_toString(d)
    AssertEqual e,a
End Sub



'###################################################################################################
'common
Sub assertAllElements(e,a)
    AssertEqualWithMessage Ubound(e), Ubound(a), "Ubound"
    Dim i
    For i=0 To Ubound(e)
        If IsObject(e(i)) Then
            AssertSameWithMessage e(i), a(i), "Element Object"
        Else
            AssertEqualWithMessage e(i), a(i), "Element Variable"
        End If
    Next
End Sub

' Local Variables:
' mode: Visual-Basic
' indent-tabs-mode: nil
' End:
