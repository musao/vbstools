' libCom.vbs: cf_* procedure test.
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
'cf_tryCatch()
Sub Test_cf_tryCatch_TryOnly_Normal
    Dim oRet : Set oRet = cf_tryCatch(new_Func("a=>1/a"), 2, Nothing, Empty)
    
    AssertEqual 0, Err.Number
    AssertEqual True, oRet.Item("Result")
    AssertEqual 1/2, oRet.Item("Return")
    AssertSame Nothing, oRet.Item("Err")
End Sub
Sub Test_cf_tryCatch_TryAndCatch_Normal
    Dim oRet : Set oRet = cf_tryCatch(new_Func("a=>1/a"), 2, new_Func("(a,e)=>e.Item(""Description"")"), Nothing)
    
    AssertEqual 0, Err.Number
    AssertEqual True, oRet.Item("Result")
    AssertEqual 1/2, oRet.Item("Return")
    AssertSame Nothing, oRet.Item("Err")
End Sub
Sub Test_cf_tryCatch_TryAndFinary_Normal
    Dim oRet : Set oRet = cf_tryCatch(new_Func("a=>1/a"), 2, Empty, new_Func("(a,r,e)=>1/2+r"))
    
    AssertEqual 0, Err.Number
    AssertEqual True, oRet.Item("Result")
    AssertEqual 1/2+1/2, oRet.Item("Return")
    AssertSame Nothing, oRet.Item("Err")
End Sub
Sub Test_cf_tryCatch_TryAndFinary_Normal_FinaryErr
    On Error Resume Next
    Dim oRet : Set oRet = cf_tryCatch(new_Func("a=>1/a"), 2, Empty, new_Func("(a,r,e)=>r(0)"))
    
    AssertEqual 13, Err.Number
    AssertEqual "Œ^‚ªˆê’v‚µ‚Ü‚¹‚ñB", Err.Description
    AssertEqual Empty, oRet
End Sub
Sub Test_cf_tryCatch_TryAndCatchAndFinary_Normal
    Dim oRet : Set oRet = cf_tryCatch(new_Func("a=>1/a"), 2, new_Func("(a,e)=>e.Item(""Description"")"), new_Func("(a,r,e)=>1/2+r"))
    
    AssertEqual 0, Err.Number
    AssertEqual True, oRet.Item("Result")
    AssertEqual 1/2+1/2, oRet.Item("Return")
    AssertSame Nothing, oRet.Item("Err")
End Sub
Sub Test_cf_tryCatch_TryAndCatchAndFinary_Normal_FinaryErr
    On Error Resume Next
    Dim oRet : Set oRet = cf_tryCatch(new_Func("a=>1/a"), 2, new_Func("(a,e)=>e.Item(""Description"")"), new_Func("(a,r,e)=>r(0)"))
    
    AssertEqual 13, Err.Number
    AssertEqual "Œ^‚ªˆê’v‚µ‚Ü‚¹‚ñB", Err.Description
    AssertEqual Empty, oRet
End Sub
Sub Test_cf_tryCatch_TryOnly_Err
    Dim oRet : Set oRet = cf_tryCatch(new_Func("a=>1/a"), 0, Empty, Empty)
    
    AssertEqual 0, Err.Number
    AssertEqual False, oRet.Item("Result")
    AssertEqual Empty, oRet.Item("Return")
    AssertEqual 11, oRet.Item("Err").Item("Number")
    AssertEqual "0 ‚ÅœZ‚µ‚Ü‚µ‚½B", oRet.Item("Err").Item("Description")
    AssertEqual "Microsoft VBScript ÀsƒGƒ‰[", oRet.Item("Err").Item("Source")
End Sub
Sub Test_cf_tryCatch_TryAndCatch_Err
    Dim oRet : Set oRet = cf_tryCatch(new_Func("a=>1/a"), 0, new_Func("(a,e)=>e.Item(""Description"")"), Empty)
    
    AssertEqual 0, Err.Number
    AssertEqual False, oRet.Item("Result")
    AssertEqual oRet.Item("Err").Item("Description"), oRet.Item("Return")
    AssertEqual 11, oRet.Item("Err").Item("Number")
    AssertEqual "0 ‚ÅœZ‚µ‚Ü‚µ‚½B", oRet.Item("Err").Item("Description")
    AssertEqual "Microsoft VBScript ÀsƒGƒ‰[", oRet.Item("Err").Item("Source")
End Sub
Sub Test_cf_tryCatch_TryAndCatch_Err_CatchErr
    On Error Resume Next
    Dim oRet : Set oRet = cf_tryCatch(new_Func("a=>1/a"), 0, new_Func("(a,e)=>a(0)"), Empty)
    
    AssertEqual 13, Err.Number
    AssertEqual "Œ^‚ªˆê’v‚µ‚Ü‚¹‚ñB", Err.Description
    AssertEqual Empty, oRet
End Sub
Sub Test_cf_tryCatch_TryAndFinary_Err
    Dim oRet : Set oRet = cf_tryCatch(new_Func("a=>1/a"), 0, Nothing, new_Func("(a,r,e)=>2+a"))
    
    AssertEqual 0, Err.Number
    AssertEqual False, oRet.Item("Result")
    AssertEqual 2+0, oRet.Item("Return")
    AssertEqual 11, oRet.Item("Err").Item("Number")
    AssertEqual "0 ‚ÅœZ‚µ‚Ü‚µ‚½B", oRet.Item("Err").Item("Description")
    AssertEqual "Microsoft VBScript ÀsƒGƒ‰[", oRet.Item("Err").Item("Source")
End Sub
Sub Test_cf_tryCatch_TryAndFinary_Err_FinaryErr
    On Error Resume Next
    Dim oRet : Set oRet = cf_tryCatch(new_Func("a=>1/a"), 0, Nothing, new_Func("(a,r,e)=>r(0)"))
    
    AssertEqual 13, Err.Number
    AssertEqual "Œ^‚ªˆê’v‚µ‚Ü‚¹‚ñB", Err.Description
    AssertEqual Empty, oRet
End Sub
Sub Test_cf_tryCatch_TryAndCatchAndFinary_Err
    Dim oRet : Set oRet = cf_tryCatch(new_Func("a=>1/a"), 0, new_Func("(a,e)=>e.Item(""Source"")"), new_Func("(a,r,e)=>2+a"))
    
    AssertEqual 0, Err.Number
    AssertEqual False, oRet.Item("Result")
    AssertEqual 2+0, oRet.Item("Return")
    AssertEqual 11, oRet.Item("Err").Item("Number")
    AssertEqual "0 ‚ÅœZ‚µ‚Ü‚µ‚½B", oRet.Item("Err").Item("Description")
    AssertEqual "Microsoft VBScript ÀsƒGƒ‰[", oRet.Item("Err").Item("Source")
End Sub
Sub Test_cf_tryCatch_TryAndCatchAndFinary_Err_CatchErr
    On Error Resume Next
    Dim oRet : Set oRet = cf_tryCatch(new_Func("a=>1/a"), 0, new_Func("(a,e)=>a(0)"), new_Func("(a,r,e)=>2+a"))
    
    AssertEqual 13, Err.Number
    AssertEqual "Œ^‚ªˆê’v‚µ‚Ü‚¹‚ñB", Err.Description
    AssertEqual Empty, oRet
End Sub
Sub Test_cf_tryCatch_TryAndCatchAndFinary_Err_FinaryErr
    On Error Resume Next
    Dim oRet : Set oRet = cf_tryCatch(new_Func("a=>1/a"), 0, new_Func("(a,e)=>e.Item(""Source"")"), new_Func("(a,r,e)=>r(0)"))
    
    AssertEqual 13, Err.Number
    AssertEqual "Œ^‚ªˆê’v‚µ‚Ü‚¹‚ñB", Err.Description
    AssertEqual Empty, oRet
End Sub

' Local Variables:
' mode: Visual-Basic
' indent-tabs-mode: nil
' End:
