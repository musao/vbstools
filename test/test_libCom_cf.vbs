' libCom.vbs: cf_* procedure test.
' @import ../lib/clsCmArray.vbs
' @import ../lib/clsCmBroker.vbs
' @import ../lib/clsCmBufferedReader.vbs
' @import ../lib/clsCmBufferedWriter.vbs
' @import ../lib/clsCmCalendar.vbs
' @import ../lib/clsCmCharacterType.vbs
' @import ../lib/clsCmCssGenerator.vbs
' @import ../lib/clsCmHtmlGenerator.vbs
' @import ../lib/clsCompareExcel.vbs
' @import ../lib/libCom.vbs

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
