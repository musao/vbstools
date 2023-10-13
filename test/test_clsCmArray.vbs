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
Sub Test_clsCmArray_item_OutOfRangeLarge
    On Error Resume Next
    Dim a : Set a = new clsCmArray
    a.push "hoge"
    a.push CreateObject("Scripting.Dictionary")
    Dim x : x = a(2)
    
    AssertEqual 9, Err.Number
    AssertEqual "インデックスが有効範囲にありません。", Err.Description
    AssertEqual Empty, x
End Sub
Sub Test_clsCmArray_item_OutOfRangeSmall
    On Error Resume Next
    Dim a : Set a = new clsCmArray
    a.push "hoge"
    a.push CreateObject("Scripting.Dictionary")
    Dim x : x = a(-1)
    
    AssertEqual 9, Err.Number
    AssertEqual "インデックスが有効範囲にありません。", Err.Description
    AssertEqual Empty, x
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
Sub Test_clsCmArray_concat_Array
    Dim e : e = Array(1,2,3,4,5)
    Dim d1 : d1 = Array(1,2,3)
    Dim d2 : d2 = Array(4,5)
    Dim a1 : Set a1 = new_ArrWith(d1)
    Dim a2 : Set a2 = a1.concat(d2)
    
    AssertEqual Ubound(e)+1, a2.length
    AssertEqual e(0), a2(0)
    AssertEqual e(1), a2(1)
    AssertEqual e(2), a2(2)
    AssertEqual e(3), a2(3)
    AssertEqual e(4), a2(4)
    
    AssertEqual Ubound(d1)+1, a1.length
    AssertEqual d1(0), a1(0)
    AssertEqual d1(1), a1(1)
    AssertEqual d1(2), a1(2)
End Sub
Sub Test_clsCmArray_concat_Variable
    Dim e : e = Array(1,2,3,5)
    Dim d1 : d1 = Array(1,2,3)
    Dim d2 : d2 = 5
    Dim a1 : Set a1 = new_ArrWith(d1)
    Dim a2 : Set a2 = a1.concat(d2)
    
    AssertEqual Ubound(e)+1, a2.length
    AssertEqual e(0), a2(0)
    AssertEqual e(1), a2(1)
    AssertEqual e(2), a2(2)
    AssertEqual e(3), a2(3)
    
    AssertEqual Ubound(d1)+1, a1.length
    AssertEqual d1(0), a1(0)
    AssertEqual d1(1), a1(1)
    AssertEqual d1(2), a1(2)
End Sub

'###################################################################################################
'clsCmArray.every()
Sub Test_clsCmArray_every_True
    Dim d : d = Array(1,2,3)
    Dim a : Set a = new_ArrWith(d)
    
    Assert a.every(new_Func("(e,i,a)=>e<5"))
    
    AssertEqual Ubound(d)+1, a.length
    AssertEqual d(0), a(0)
    AssertEqual d(1), a(1)
    AssertEqual d(2), a(2)
End Sub
Sub Test_clsCmArray_every_False
    Dim d : d = Array(1,2,3)
    Dim a : Set a = new_ArrWith(d)
    
    Assert Not a.every(new_Func("(e,i,a)=>e<3"))
    
    AssertEqual Ubound(d)+1, a.length
    AssertEqual d(0), a(0)
    AssertEqual d(1), a(1)
    AssertEqual d(2), a(2)
End Sub

'###################################################################################################
'clsCmArray.filter()
Sub Test_clsCmArray_filter
    Dim e : e = Array(2,3)
    Dim d : d = Array(1,2,3)
    Dim a1 : Set a1 = new_ArrWith(d)
    Dim a2 : Set a2 = a1.filter(new_Func("(e,i,a)=>e>1"))
    
    AssertEqual Ubound(e)+1, a2.length
    AssertEqual e(0), a2(0)
    AssertEqual e(1), a2(1)
    
    AssertEqual Ubound(d)+1, a1.length
    AssertEqual d(0), a1(0)
    AssertEqual d(1), a1(1)
    AssertEqual d(2), a1(2)
End Sub

'###################################################################################################
'clsCmArray.find()
Sub Test_clsCmArray_find
    Dim e : e = 12
    Dim d : d = Array(5,12,8,130,44)
    Dim a1 : Set a1 = new_ArrWith(d)
    Dim a2 : a2 = a1.find(new_Func("(e,i,a)=>e>11"))
    
    AssertEqual e, a2
    
    AssertEqual Ubound(d)+1, a1.length
    AssertEqual d(0), a1(0)
    AssertEqual d(1), a1(1)
    AssertEqual d(2), a1(2)
    AssertEqual d(3), a1(3)
    AssertEqual d(4), a1(4)
End Sub

'###################################################################################################
'clsCmArray.forEach()
Sub Test_clsCmArray_forEach
    Dim e : e = Array(2,3,4)
    Dim a : Set a = new_ArrWith(Array(1,2,3))
    a.forEach(new_Func("function(e,i,a) {a(i)=a(i)+1}"))
    
    AssertEqual Ubound(e)+1, a.length
    AssertEqual e(0), a(0)
    AssertEqual e(1), a(1)
    AssertEqual e(2), a(2)
End Sub

'###################################################################################################
'clsCmArray.indexOf()
Sub Test_clsCmArray_indexOf
    Dim d : d = Array("a","b","c","b")
    Dim a : Set a = new_ArrWith(d)
    
    AssertEqual 1, a.indexOf("b")
    AssertEqual -1, a.indexOf("z")
End Sub

'###################################################################################################
'clsCmArray.joinVbs()
Sub Test_clsCmArray_joinVbs
    Dim d : d = Array(1,2,3,"testing")
    Dim e : e = Join(d, "+")
    Dim a : a = new_ArrWith(d).joinVbs("+")
    
    AssertEqual e, a
End Sub

'###################################################################################################
'clsCmArray.lastIndexOf()
Sub Test_clsCmArray_lastIndexOf
    Dim d : Set d = new_DicWith(Array(4, "five"))
    Dim a : Set a = new_ArrWith(Array("a", 2, 3.14, d, d, "End"))
    
    AssertEqual 0, a.lastIndexOf("a")
    AssertEqual 4, a.lastIndexOf(d)
    AssertEqual -1, a.lastIndexOf("Start")
    AssertEqual 1, a.lastIndexOf(2)
    AssertEqual -1, a.lastIndexOf("2")
End Sub

'###################################################################################################
'clsCmArray.map()
Sub Test_clsCmArray_map
    Dim e : e = Array(1,4,9)
    Dim d : d = Array(1,2,3)
    Dim a1 : Set a1 = new_ArrWith(d)
    Dim a2 : Set a2 = a1.map(new_Func("(e,i,a)=>e*e"))
    
    AssertEqual Ubound(e)+1, a2.length
    AssertEqual e(0), a2(0)
    AssertEqual e(1), a2(1)
    AssertEqual e(2), a2(2)
    
    AssertEqual Ubound(d)+1, a1.length
    AssertEqual d(0), a1(0)
    AssertEqual d(1), a1(1)
    AssertEqual d(2), a1(2)
End Sub

'###################################################################################################
'clsCmArray.pop()/.push()/pushMulti()
Sub Test_clsCmArray_pop_push_pushMulti
    Dim a,e
    Set a = new clsCmArray
    
    e = Array("hoge", 2, "参", Nothing)
    AssertEqual 4, a.pushMulti(e)
    
    AssertEqual Ubound(e)+1, a.length
    AssertEqual e(0), a(0)
    AssertEqual e(1), a(1)
    AssertEqual e(2), a(2)
    AssertSame e(3), a(3)
    
    AssertSame a(3), a.pop
    AssertEqual a(2), a.pop
    
    AssertEqual 2, a.length
    AssertEqual e(0), a(0)
    AssertEqual e(1), a(1)
    
    e = Array("hoge", 2, Empty, "四")
    AssertEqual 3, a.push(e(2))
    AssertEqual 4, a.push(e(3))
    
    AssertEqual Ubound(e)+1, a.length
    AssertEqual e(0), a(0)
    AssertEqual e(1), a(1)
    AssertEqual e(2), a(2)
    AssertEqual e(3), a(3)
    
    
    AssertEqual a(3), a.pop
    AssertEqual a(2), a.pop
    AssertEqual a(1), a.pop
    AssertEqual a(0), a.pop
    AssertEqual 0, a.length
End Sub

'###################################################################################################
'clsCmArray.reduce()
Sub Test_clsCmArray_reduce
    Dim e : e = 24
    Dim d : d = Array(1,2,3,4)
    Dim a1 : Set a1 = new_ArrWith(d)
    Dim a2 : a2 = a1.reduce(new_Func("(p,c,i,a)=>p*c"))
    
    AssertEqual e, a2
    
    AssertEqual Ubound(d)+1, a1.length
    AssertEqual d(0), a1(0)
    AssertEqual d(1), a1(1)
    AssertEqual d(2), a1(2)
    AssertEqual d(3), a1(3)
End Sub

'###################################################################################################
'clsCmArray.reduceRight()
Sub Test_clsCmArray_reduceRight
    Dim e : e = 3
    Dim d : d = Array(2,10,60)
    Dim a1 : Set a1 = new_ArrWith(d)
    Dim a2 : a2 = a1.reduceRight(new_Func("(p,c,i,a)=>p/c"))
    
    AssertEqual e, a2
    
    AssertEqual Ubound(d)+1, a1.length
    AssertEqual d(0), a1(0)
    AssertEqual d(1), a1(1)
    AssertEqual d(2), a1(2)
End Sub


' Local Variables:
' mode: Visual-Basic
' indent-tabs-mode: nil
' End:
