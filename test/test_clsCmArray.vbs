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
    
    assertAllElements e, a2
    
    assertAllElements d1, a1
End Sub
Sub Test_clsCmArray_concat_Variable
    Dim e : e = Array(1,2,3,5)
    Dim d1 : d1 = Array(1,2,3)
    Dim d2 : d2 = 5
    Dim a1 : Set a1 = new_ArrWith(d1)
    Dim a2 : Set a2 = a1.concat(d2)
    
    assertAllElements e, a2
    
    assertAllElements d1, a1
End Sub

'###################################################################################################
'clsCmArray.every()
Sub Test_clsCmArray_every_True
    Dim d : d = Array(1,2,3)
    Dim a : Set a = new_ArrWith(d)
    
    Assert a.every(new_Func("(e,i,a)=>e<5"))
    
    assertAllElements d, a
End Sub
Sub Test_clsCmArray_every_False
    Dim d : d = Array(1,2,3)
    Dim a : Set a = new_ArrWith(d)
    
    Assert Not a.every(new_Func("(e,i,a)=>e<3"))
    
    assertAllElements d, a
End Sub

'###################################################################################################
'clsCmArray.filter()
Sub Test_clsCmArray_filter
    Dim e : e = Array(2,3)
    Dim d : d = Array(1,2,3)
    Dim a1 : Set a1 = new_ArrWith(d)
    Dim a2 : Set a2 = a1.filter(new_Func("(e,i,a)=>e>1"))
    
    assertAllElements e, a2
    
    assertAllElements d, a1
End Sub

'###################################################################################################
'clsCmArray.find()
Sub Test_clsCmArray_find
    Dim e : e = 12
    Dim d : d = Array(5,12,8,130,44)
    Dim a1 : Set a1 = new_ArrWith(d)
    Dim a2 : a2 = a1.find(new_Func("(e,i,a)=>e>11"))
    
    AssertEqual e, a2
    
    assertAllElements d, a1
End Sub

'###################################################################################################
'clsCmArray.forEach()
Sub Test_clsCmArray_forEach
    Dim e : e = Array(2,3,4)
    Dim a : Set a = new_ArrWith(Array(1,2,3))
    a.forEach(new_Func("function(e,i,a) {a(i)=a(i)+1}"))
    
    assertAllElements e, a
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
    
    assertAllElements e, a2
    
    assertAllElements d, a1
End Sub

'###################################################################################################
'clsCmArray.pop()/.push()/pushMulti()
Sub Test_clsCmArray_pop_push_pushMulti
    Dim a,e
    Set a = new clsCmArray
    
    e = Array("hoge", 2, "参", Nothing)
    AssertEqual 4, a.pushMulti(e)
    
    assertAllElements e, a
    
    AssertSame e(3), a.pop
    AssertEqual e(2), a.pop
    
    AssertEqual 2, a.length
    AssertEqual e(0), a(0)
    AssertEqual e(1), a(1)
    
    e = Array("hoge", 2, Empty, "四")
    AssertEqual 3, a.push(e(2))
    AssertEqual 4, a.push(e(3))
    
    assertAllElements e, a
    
    AssertEqual e(3), a.pop
    AssertEqual e(2), a.pop
    AssertEqual e(1), a.pop
    AssertEqual e(0), a.pop
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
    
    assertAllElements d, a1
End Sub

'###################################################################################################
'clsCmArray.reduceRight()
Sub Test_clsCmArray_reduceRight
    Dim e : e = 3
    Dim d : d = Array(2,10,60)
    Dim a1 : Set a1 = new_ArrWith(d)
    Dim a2 : a2 = a1.reduceRight(new_Func("(p,c,i,a)=>p/c"))
    
    AssertEqual e, a2
    
    assertAllElements d, a1
End Sub

'###################################################################################################
'clsCmArray.reverse()
Sub Test_clsCmArray_reverse
    Dim e : e = Array(3,Nothing,1)
    Dim d : d = Array(1,Nothing,3)
    Dim a : Set a = new_ArrWith(d)
    a.reverse
    
    assertAllElements e, a
End Sub

'###################################################################################################
'clsCmArray.shift()/.unshift()/unshiftMulti()
Sub Test_clsCmArray_shift_unshift_unshiftMulti
    Dim a,e
    Set a = new clsCmArray
    
    e = Array("hoge", 2, "参", Nothing)
    AssertEqual 1, a.unshift(e(3))
    AssertEqual 2, a.unshift(e(2))
    AssertEqual 3, a.unshift(e(1))
    AssertEqual 4, a.unshift(e(0))
    
    assertAllElements e, a
    
    AssertEqual e(0), a.shift
    AssertEqual e(1), a.shift
    
    AssertEqual 2, a.length
    AssertEqual e(2), a(0)
    AssertSame e(3), a(1)
    
    AssertEqual 4, a.unshiftMulti(Array(Empty, "四"))
    
    e = Array(Empty, "四", "参", Nothing)
    assertAllElements e, a
    
    AssertEqual e(0), a.shift
    AssertEqual e(1), a.shift
    AssertEqual e(2), a.shift
    AssertSame e(3), a.shift
    AssertEqual 0, a.length
End Sub





'###################################################################################################
'common
Sub assertAllElements(e,a)
    AssertEqual Ubound(e)+1, a.length
    Dim i
    For i=0 To Ubound(e)
        If IsObject(e(i)) Then
            AssertSame e(i), a(i)
        Else
            AssertEqual e(i), a(i)
        End If
    Next
End Sub


' Local Variables:
' mode: Visual-Basic
' indent-tabs-mode: nil
' End:
