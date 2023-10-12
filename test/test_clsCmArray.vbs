' clsCmArray.vbs: new_* procedure test.
' @import ../lib/clsCmArray.vbs
' @import ../lib/clsCmBufferedReader.vbs
' @import ../lib/clsCmBufferedWriter.vbs
' @import ../lib/clsCmCalendar.vbs
' @import ../lib/clsCmPubSub.vbs
' @import ../lib/clsCompareExcel.vbs
' @import ../lib/clsFsBase.vbs
' @import ../lib/libCom.vbs

Option Explicit

'###################################################################################################
'clsCmArray.count()
Sub Test_clsCmArray_count
    Dim a : Set a = new clsCmArray
    
    AssertEqual 0, a.count
    
    a.push "hoge"
    
    AssertEqual 1, a.count
End Sub

'###################################################################################################
'clsCmArray.item()
Sub Test_clsCmArray_item
    Dim ev : ev = "hoge"
    Dim eo : Set eo = CreateObject("Scripting.Dictionary")
    Dim a : Set a = new clsCmArray
    
    a.push ev
    a.push eo
    
    AssertEqual ev, a.item(0)
    AssertEqual ev, a(0)
    AssertSame eo, a.item(1)
    AssertSame eo, a(1)
End Sub

'###################################################################################################
'clsCmArray.items()
Sub Test_clsCmArray_items
    Dim e()
    Dim a : Set a = new clsCmArray
    
    a.push CreateObject("Scripting.Dictionary")
    a.push "hoge"
    
    AssertEqual VarType(e), VarType(a.items)
    AssertEqual TypeName(e), TypeName(a.items)
    AssertEqual a.Length-1, Ubound(a.items)
End Sub

'###################################################################################################
'clsCmArray.length()
Sub Test_clsCmArray_length
    Dim a : Set a = new clsCmArray
    
    AssertEqual a.count, a.length
    
    a.push "hoge"
    
    AssertEqual a.count, a.length
End Sub

'###################################################################################################
'clsCmArray.concat()
Sub Test_clsCmArray_concat
    Dim e : e = Array(1,2,3,4,5)
    Dim e1 : e1 = Array(1,2,3)
    Dim e2 : e2 = Array(4,5)
    Dim a : Set a = new_ArrWith(e1)
    
    AssertEqual func_CM_ToString(e), func_CM_ToString(a.concat(e2))
    AssertEqual func_CM_ToString(e1), func_CM_ToString(a)
End Sub


' Local Variables:
' mode: Visual-Basic
' indent-tabs-mode: nil
' End:
