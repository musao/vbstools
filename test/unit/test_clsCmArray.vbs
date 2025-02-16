' clsCmArray.vbs: test.
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
'clsCmArray
Sub Test_clsCmArray
    Dim a : Set a = new clsCmArray
    AssertEqual 9, VarType(a)
    AssertEqual "clsCmArray", TypeName(a)
End Sub

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
Sub Test_clsCmArray_get_item
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
Sub Test_clsCmArray_get_item_OutOfRangeLarge
    On Error Resume Next
    Dim a : Set a = new clsCmArray
    a.push "hoge"
    a.push CreateObject("Scripting.Dictionary")
    Dim x : x = a(2)
    
    AssertEqualWithMessage "clsCmArray+item() Get", Err.Source, "Err.Source"
    AssertEqualWithMessage "Index is out of range.", Err.Description, "Err.Description"
    AssertEqualWithMessage Empty, x, "value"
End Sub
Sub Test_clsCmArray_get_item_OutOfRangeSmall
    On Error Resume Next
    Dim a : Set a = new clsCmArray
    a.push "hoge"
    a.push CreateObject("Scripting.Dictionary")
    Dim x : x = a(-1)
    
    AssertEqualWithMessage "clsCmArray+item() Get", Err.Source, "Err.Source"
    AssertEqualWithMessage "Index is out of range.", Err.Description, "Err.Description"
    AssertEqualWithMessage Empty, x, "value"
End Sub
Sub Test_clsCmArray_set_let_item
    Dim ev : ev = "hoge"
    Dim eo : Set eo = CreateObject("Scripting.Dictionary")
    Dim a : Set a = new clsCmArray
    
    a.push "fuga"
    a.push "foo"
    set a.item(0) = eo
    a(1) = ev
    
    AssertSame eo, a.item(0)
    AssertSame eo, a(0)
    AssertEqual ev, a.item(1)
    AssertEqual ev, a(1)
End Sub
Sub Test_clsCmArray_set_let_item_OutOfRangeLarge
    On Error Resume Next
    Dim a : Set a = new clsCmArray
    a.push "hoge"
    a.push CreateObject("Scripting.Dictionary")
    a(2) = "fuga"
    
    AssertEqualWithMessage "clsCmArray+item() Let", Err.Source, "Err.Source"
    AssertEqualWithMessage "Index is out of range.", Err.Description, "Err.Description"
    AssertEqualWithMessage 2, a.length, "length"
End Sub
Sub Test_clsCmArray_set_let_item_OutOfRangeSmall
    On Error Resume Next
    Dim a : Set a = new clsCmArray
    a.push "hoge"
    a.push CreateObject("Scripting.Dictionary")
    Set a(-1) = CreateObject("Scripting.Dictionary")
    
    AssertEqualWithMessage "clsCmArray+item() Set", Err.Source, "Err.Source"
    AssertEqualWithMessage "Index is out of range.", Err.Description, "Err.Description"
    AssertEqualWithMessage 2, a.length, "length"
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
    AssertEqual a.length-1, Ubound(a.items)
End Sub
Sub Test_clsCmArray_items_AllItems
    Dim e,d,ao,a
    Set ao = new clsCmArray
    d = Array(1,"b",Nothing)
    ao.pushA d

    e = d
    a = ao.items()
    assertAllElementsArray e, a

    Dim d2 : d2 = d
    d(1) = "Z"
    ao(1) = "X"
    e = d2
    assertAllElementsArray e, a
End Sub
Sub Test_clsCmArray_items_NoData
    Dim e,ao,a
    Set ao = new clsCmArray

    e = Array()
    a = ao.items()
    assertAllElementsArray e, a
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
Sub Test_clsCmArray_concat_Array_New
    Dim e : e = Array(1,2,3,4,5)
    Dim a1 : Set a1 = new_Arr()
    Dim a2 : Set a2 = a1.concat(e)
    
    assertAllElements e, a2
End Sub
Sub Test_clsCmArray_concat_Array_Add
    Dim e : e = Array(1,2,3,4,5)
    Dim d1 : d1 = Array(1,2,3)
    Dim d2 : d2 = Array(4,5)
    Dim a1 : Set a1 = new_ArrOf(d1)
    Dim a2 : Set a2 = a1.concat(d2)
    
    assertAllElements e, a2
    
    assertAllElements d1, a1
End Sub
Sub Test_clsCmArray_concat_Variable_New
    Dim e : e = Array(5)
    Dim d : d = 5
    Dim a1 : Set a1 = new_Arr()
    Dim a2 : Set a2 = a1.concat(d)
    
    assertAllElements e, a2
End Sub
Sub Test_clsCmArray_concat_Variable_Add
    Dim e : e = Array(1,2,3,5)
    Dim d1 : d1 = Array(1,2,3)
    Dim d2 : d2 = 5
    Dim a1 : Set a1 = new_ArrOf(d1)
    Dim a2 : Set a2 = a1.concat(d2)
    
    assertAllElements e, a2
    
    assertAllElements d1, a1
End Sub

'###################################################################################################
'clsCmArray.every()
Sub Test_clsCmArray_every_True
    Dim d : d = Array(1,2,3)
    Dim a : Set a = new_ArrOf(d)
    
    Assert a.every(new_Func("(e,i,a)=>e<5"))
    
    assertAllElements d, a
End Sub
Sub Test_clsCmArray_every_False
    Dim d : d = Array(1,2,3)
    Dim a : Set a = new_ArrOf(d)
    
    Assert Not a.every(new_Func("(e,i,a)=>e<3"))
    
    assertAllElements d, a
End Sub
Sub Test_clsCmArray_every_Empty
    Dim a : Set a = new_Arr()
    AssertEqual 0, a.length
    
    Assert a.every(new_Func("(e,i,a)=>e<5"))
    Assert a.every(new_Func("(e,i,a)=>e<3"))
    AssertEqual 0, a.length
End Sub

'###################################################################################################
'clsCmArray.filter()
Sub Test_clsCmArray_filter
    Dim e : e = Array(2,3)
    Dim d : d = Array(1,2,3)
    Dim a1 : Set a1 = new_ArrOf(d)
    Dim a2 : Set a2 = a1.filter(new_Func("(e,i,a)=>e>1"))
    
    assertAllElements e, a2
    
    assertAllElements d, a1
End Sub
Sub Test_clsCmArray_filter_Empty
    Dim d : d = Array(1,2,3)
    Dim a1 : Set a1 = new_ArrOf(d)
    Dim a2 : Set a2 = a1.filter(new_Func("(e,i,a)=>e>3"))
    
    AssertEqual 0, a2.length
    
    assertAllElements d, a1
End Sub

'###################################################################################################
'clsCmArray.find()
Sub Test_clsCmArray_Variable_find
    Dim e : e = 12
    Dim d : d = Array(5,12,8,130,44)
    Dim a1 : Set a1 = new_ArrOf(d)
    Dim a2 : a2 = a1.find(new_Func("(e,i,a)=>e>11"))
    
    AssertEqual e, a2
    
    assertAllElements d, a1
End Sub
Sub Test_clsCmArray_Object_find
    Dim e : Set e = Nothing
    Dim d : d = Array(0,"",vbNullString,Nothing,Empty)
    Dim a1 : Set a1 = new_ArrOf(d)
    Dim a2 : Set a2 = a1.find(new_Func("(e,i,a)=>{if isobject(e) then:return (e is Nothing):end if}"))
    
    AssertSame e, a2
    
    assertAllElements d, a1
End Sub
Sub Test_clsCmArray_find_NoHit
    Dim e : e = Empty
    Dim d : d = Array(5,12,8,130,44)
    Dim a1 : Set a1 = new_ArrOf(d)
    Dim a2 : a2 = a1.find(new_Func("(e,i,a)=>e>200"))
    
    AssertEqual e, a2
    
    assertAllElements d, a1
End Sub

'###################################################################################################
'clsCmArray.forEach()
Sub Test_clsCmArray_forEach
    Dim e : e = Array(2,3,4)
    Dim a : Set a = new_ArrOf(Array(1,2,3))
    a.forEach(new_Func("function(e,i,a) {a(i)=a(i)+1}"))
    
    assertAllElements e, a
End Sub
Sub Test_clsCmArray_forEach_Empty
    Dim a : Set a = new_Arr()
    a.forEach(new_Func("function(e,i,a) {a(i)=a(i)+1}"))
    
    AssertEqual 0, a.length
End Sub

'###################################################################################################
'clsCmArray.hasElement()
Sub Test_clsCmArray_hasElement_Ok
    Dim a
    Set a = new_Arr()
    
    Dim d1(0)
    AssertEqualWithMessage True, a.hasElement(d1), "1"
    
    Dim d2(1)
    AssertEqualWithMessage True, a.hasElement(d2), "2"
    
    Dim d3
    d3 = Array(1)
    AssertEqualWithMessage True, a.hasElement(d3), "3"
    
    d3 = Array(1,2)
    AssertEqualWithMessage True, a.hasElement(d3), "4"
    
    Redim d4(1)
    AssertEqualWithMessage True, a.hasElement(d4), "5"
    
    Redim Preserve d4(2)
    AssertEqualWithMessage True, a.hasElement(d4), "6"
    
    Redim d4(0)
    AssertEqualWithMessage True, a.hasElement(d4), "7"
End Sub
Sub Test_clsCmArray_hasElement_Ng_Variable
    Dim a
    Set a = new_Arr()
    
    Dim d
    d = Empty
    AssertEqualWithMessage False, a.hasElement(d), "1"
    
    d = vbNullString
    AssertEqualWithMessage False, a.hasElement(d), "2"
    
    Set d = Nothing
    AssertEqualWithMessage False, a.hasElement(d), "3"
    
    d = "abc"
    AssertEqualWithMessage False, a.hasElement(d), "4"
    
    d = 123
    AssertEqualWithMessage False, a.hasElement(d), "5"
End Sub
Sub Test_clsCmArray_hasElement_Ng_Array
    Dim a
    Set a = new_Arr()
    
    Dim d()
    AssertEqualWithMessage False, a.hasElement(d), "1"
    
    Dim d2: d2 = Array()
    AssertEqualWithMessage False, a.hasElement(d2), "2"
End Sub

'###################################################################################################
'clsCmArray.indexOf()
Sub Test_clsCmArray_indexOf
    Dim d : d = Array("a","b","c","b")
    Dim a : Set a = new_ArrOf(d)
    
    AssertEqual 1, a.indexOf("b")
    AssertEqual -1, a.indexOf("z")
End Sub
Sub Test_clsCmArray_indexOf_Empty
    Dim a : Set a = new_Arr()
    
    AssertEqual -1, a.indexOf("b")
    AssertEqual -1, a.indexOf("z")
End Sub

'###################################################################################################
'clsCmArray.join()
Sub Test_clsCmArray_join
    Dim d : d = Array(1,2,3,"testing")
    Dim e : e = Join(d, "+")
    Dim a : a = new_ArrOf(d).join("+")
    
    AssertEqual e, a
End Sub
Sub Test_clsCmArray_join_Empty
    Dim e : e = ""
    Dim a : a = new_Arr().join("+")
    
    AssertEqual e, a
End Sub

'###################################################################################################
'clsCmArray.lastIndexOf()
Sub Test_clsCmArray_lastIndexOf
    Dim d : Set d = new_DicOf(Array(4, "five"))
    Dim a : Set a = new_ArrOf(Array("a", 2, 3.14, d, d, "End"))
    
    AssertEqual 0, a.lastIndexOf("a")
    AssertEqual 4, a.lastIndexOf(d)
    AssertEqual -1, a.lastIndexOf("Start")
    AssertEqual 1, a.lastIndexOf(2)
    AssertEqual -1, a.lastIndexOf("2")
End Sub
Sub Test_clsCmArray_lastIndexOf_Empty
    Dim a : Set a = new_Arr()
    
    AssertEqual -1, a.lastIndexOf("a")
    AssertEqual -1, a.lastIndexOf("Start")
    AssertEqual -1, a.lastIndexOf(2)
    AssertEqual -1, a.lastIndexOf("2")
End Sub

'###################################################################################################
'clsCmArray.map()
Sub Test_clsCmArray_map
    Dim e : e = Array(1,4,9)
    Dim d : d = Array(1,2,3)
    Dim a1 : Set a1 = new_ArrOf(d)
    Dim a2 : Set a2 = a1.map(new_Func("(e,i,a)=>e*e"))
    
    assertAllElements e, a2
    
    assertAllElements d, a1
End Sub
Sub Test_clsCmArray_map_Empty
    Dim a1 : Set a1 = new_Arr()
    Dim a2 : Set a2 = a1.map(new_Func("(e,i,a)=>e*e"))
    
    AssertEqual 0, a2.length
    
    AssertEqual 0, a1.length
End Sub

'###################################################################################################
'clsCmArray.pop()/.push()/pushA()
Sub Test_clsCmArray_pop_push_pushMulti
    Dim a,e
    Set a = new clsCmArray
    
    e = Array("hoge", 2, "�Q", Nothing)
    AssertEqual 4, a.pushA(e)
    
    assertAllElements e, a
    
    AssertSame e(3), a.pop
    AssertEqual e(2), a.pop
    
    AssertEqual 2, a.length
    AssertEqual e(0), a(0)
    AssertEqual e(1), a(1)
    
    e = Array("hoge", 2, Empty, "�l")
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
Sub Test_clsCmArray_reduce_InitEmpty
    Dim e : e = 24
    Dim d : d = Array(1,2,3,4)
    Dim a1 : Set a1 = new_ArrOf(d)
    Dim a2 : a2 = a1.reduce(new_Func("(p,c,i,a)=>p*c"), Empty)
    
    AssertEqual e, a2
    
    assertAllElements d, a1
End Sub
Sub Test_clsCmArray_reduce_InitNoEmpty
    Dim e : e = 120
    Dim d : d = Array(1,2,3,4)
    Dim a1 : Set a1 = new_ArrOf(d)
    Dim a2 : a2 = a1.reduce(new_Func("(p,c,i,a)=>p*c"), 5)
    
    AssertEqual e, a2
    
    assertAllElements d, a1
End Sub
Sub Test_clsCmArray_reduce_InitEmpty_Len1
    Dim e : e = 1
    Dim d : d = Array(1)
    Dim a1 : Set a1 = new_ArrOf(d)
    Dim a2 : a2 = a1.reduce(new_Func("(p,c,i,a)=>p*c"), Empty)
    
    AssertEqual e, a2
    
    assertAllElements d, a1
End Sub
Sub Test_clsCmArray_reduce_InitNoEmpty_Len1
    Dim e : e = 10
    Dim d : d = Array(1)
    Dim a1 : Set a1 = new_ArrOf(d)
    Dim a2 : a2 = a1.reduce(new_Func("(p,c,i,a)=>p*c"), 10)
    
    AssertEqual e, a2
    
    assertAllElements d, a1
End Sub
Sub Test_clsCmArray_reduce_Err
    On Error Resume Next
    Dim a1 : Set a1 = new_Arr()
    Dim a2 : a2 = a1.reduce(new_Func("(p,c,i,a)=>p*c"), Empty)
    
    AssertEqualWithMessage "clsCmArray+reduce()", Err.Source, "Err.Source"
    AssertEqualWithMessage "Array has no elements.", Err.Description, "Err.Description"
    AssertEqualWithMessage 0, a1.lentgh, "lentgh"
    AssertEqualWithMessage Empty, a2, "value"
End Sub

'###################################################################################################
'clsCmArray.reduceRight()
Sub Test_clsCmArray_reduceRight_InitEmpty
    Dim e : e = 3
    Dim d : d = Array(2,10,60)
    Dim a1 : Set a1 = new_ArrOf(d)
    Dim a2 : a2 = a1.reduceRight(new_Func("(p,c,i,a)=>p/c"), Empty)
    
    AssertEqual e, a2
    
    assertAllElements d, a1
End Sub
Sub Test_clsCmArray_reduceRight_InitNoEmpty
    Dim e : e = 5
    Dim d : d = Array(2,10,60)
    Dim a1 : Set a1 = new_ArrOf(d)
    Dim a2 : a2 = a1.reduceRight(new_Func("(p,c,i,a)=>p/c"), 100)
    
    AssertEqual e, a2
    
    assertAllElements d, a1
End Sub
Sub Test_clsCmArray_reduceRight_InitEmpty_Len1
    Dim e : e = 2
    Dim d : d = Array(2)
    Dim a1 : Set a1 = new_ArrOf(d)
    Dim a2 : a2 = a1.reduceRight(new_Func("(p,c,i,a)=>p/c"), Empty)
    
    AssertEqual e, a2
    
    assertAllElements d, a1
End Sub
Sub Test_clsCmArray_reduceRight_InitNoEmpty_Len1
    Dim e : e = 1
    Dim d : d = Array(2)
    Dim a1 : Set a1 = new_ArrOf(d)
    Dim a2 : a2 = a1.reduceRight(new_Func("(p,c,i,a)=>p/c"), 1)
    
    AssertEqual e, a2
    
    assertAllElements d, a1
End Sub
Sub Test_clsCmArray_reduceRight_Err
    On Error Resume Next
    Dim a1 : Set a1 = new_Arr()
    Dim a2 : a2 = a1.reduceRight(new_Func("(p,c,i,a)=>p*c"), Empty)
    
    AssertEqualWithMessage "clsCmArray+reduceRight()", Err.Source, "Err.Source"
    AssertEqualWithMessage "Array has no elements.", Err.Description, "Err.Description"
    AssertEqualWithMessage 0, a1.lentgh, "lentgh"
    AssertEqualWithMessage Empty, a2, "value"
End Sub

'###################################################################################################
'clsCmArray.reverse()
Sub Test_clsCmArray_reverse
    Dim e : e = Array(3,Nothing,1)
    Dim d : d = Array(1,Nothing,3)
    Dim a : Set a = new_ArrOf(d)
    a.reverse
    
    assertAllElements e, a
End Sub
Sub Test_clsCmArray_reverse_Empty
    Dim a : Set a = new_Arr()
    AssertEqual 0, a.length
    
    a.reverse
    AssertEqual 0, a.length
End Sub

'###################################################################################################
'clsCmArray.shift()/.unshift()/unshiftA()
Sub Test_clsCmArray_shift_unshift_unshiftMulti
    Dim a,e
    Set a = new clsCmArray
    
    e = Array("hoge", 2, "�Q", Nothing)
    AssertEqualWithMessage 1, a.unshift(e(3)), "1-1"
    AssertEqualWithMessage 2, a.unshift(e(2)), "1-2"
    AssertEqualWithMessage 3, a.unshift(e(1)), "1-3"
    AssertEqualWithMessage 4, a.unshift(e(0)), "1-4"
    
    assertAllElements e, a
    
    AssertEqualWithMessage e(0), a.shift, "2-1"
    AssertEqualWithMessage e(1), a.shift, "2-2"
    
    AssertEqualWithMessage 2, a.length, "3-1"
    AssertEqualWithMessage e(2), a(0), "3-2"
    AssertSameWithMessage e(3), a(1), "3-3"
    
    AssertEqualWithMessage 4, a.unshiftA(Array(Empty, "�l")), "4-1"
    
    e = Array(Empty, "�l", "�Q", Nothing)
    assertAllElements e, a
    
    AssertEqualWithMessage e(0), a.shift, "5-1"
    AssertEqualWithMessage e(1), a.shift, "5-2"
    AssertEqualWithMessage e(2), a.shift, "5-3"
    AssertSameWithMessage e(3), a.shift, "5-4"
    AssertEqualWithMessage 0, a.length, "5-5"
End Sub

'###################################################################################################
'clsCmArray.slice
Sub Test_clsCmArray_slice_Normal
    Dim e,a,a1,d
    d = Array(1,2,3,4,5)
    Set a = new_ArrOf(d)
    
    e = Array(1,2,3)
    Set a1 = a.slice(0,3)
    assertAllElements e, a1
    
    e = Array(4,5)
    Set a1 = a.slice(3,vbNullString)
    assertAllElements e, a1
    
    e = Array(2,3,4)
    Set a1 = a.slice(1,-1)
    assertAllElements e, a1
    
    e = Array(3)
    Set a1 = a.slice(-3,-2)
    assertAllElements e, a1
End Sub
Sub Test_clsCmArray_slice_Limit_Upper
    Dim e,a,a1,d
    d = Array(1,2,3,4,5)
    Set a = new_ArrOf(d)
    
    e = Array(5)
    Set a1 = a.slice(4,vbNullString)
    assertAllElements e, a1
    
    e = Array(5)
    Set a1 = a.slice(-1,vbNullString)
    assertAllElements e, a1
    
    Set a1 = a.slice(4,4)
    AssertEqual 0, a1.length
    
    Set a1 = a.slice(4,-1)
    AssertEqual 0, a1.length
    
    Set a1 = a.slice(-1,4)
    AssertEqual 0, a1.length
    
    Set a1 = a.slice(-1,-1)
    AssertEqual 0, a1.length
End Sub
Sub Test_clsCmArray_slice_Limit_Lower
    Dim e,a,a1,d
    d = Array(1,2,3,4,5)
    Set a = new_ArrOf(d)
    
    e = Array(1)
    Set a1 = a.slice(0,1)
    assertAllElements e, a1
    
    e = Array(1)
    Set a1 = a.slice(0,-4)
    assertAllElements e, a1
    
    e = Array(1)
    Set a1 = a.slice(-5,1)
    assertAllElements e, a1
    
    e = Array(1)
    Set a1 = a.slice(-5,-4)
    
    Set a1 = a.slice(0,0)
    AssertEqual 0, a1.length
    
    Set a1 = a.slice(0,-5)
    AssertEqual 0, a1.length
    
    Set a1 = a.slice(-5,0)
    AssertEqual 0, a1.length
    
    Set a1 = a.slice(-5,-5)
    AssertEqual 0, a1.length
End Sub
Sub Test_clsCmArray_slice_Empty
    Dim a,a1
    Set a = new_Arr()
    AssertEqual 0, a.length
    
    Set a1 = a.slice(0,3)
    AssertEqual 0, a1.length
    
    Set a1 = a.slice(3,vbNullString)
    AssertEqual 0, a1.length
    
    Set a1 = a.slice(1,-1)
    AssertEqual 0, a1.length
    
    Set a1 = a.slice(-3,-2)
    AssertEqual 0, a1.length
End Sub

'###################################################################################################
'clsCmArray.some()
Sub Test_clsCmArray_some_True
    Dim d : d = Array(1,2,3)
    Dim a : Set a = new_ArrOf(d)
    
    Assert a.some(new_Func("(e,i,a)=>e>2"))
    
    assertAllElements d, a
End Sub
Sub Test_clsCmArray_some_False
    Dim d : d = Array(1,2,3)
    Dim a : Set a = new_ArrOf(d)
    
    Assert Not a.some(new_Func("(e,i,a)=>e>5"))
    
    assertAllElements d, a
End Sub
Sub Test_clsCmArray_some_Empty
    Dim a : Set a = new_Arr()
    AssertEqual 0, a.length
    
    Assert Not a.some(new_Func("(e,i,a)=>e>2"))
    Assert Not a.some(new_Func("(e,i,a)=>e>5"))
    AssertEqual 0, a.length
End Sub

'###################################################################################################
'clsCmArray.sort()
Sub Test_clsCmArray_sort_Num
    Dim e,d,a
    d = Array(5,2,9,6,4,8,7,3,0,1)
    Set a = new_ArrOf(d)
    
    e = Array(0,1,2,3,4,5,6,7,8,9)
    assertAllElements e, a.sort(True)
    
    e = Array(9,8,7,6,5,4,3,2,1,0)
    assertAllElements e, a.sort(False)
End Sub
Sub Test_clsCmArray_sort_Various
    Dim e,d,a
    d = Array("C","$","b","��","a","B","��","A","c","0")
    Set a = new_ArrOf(d)
    
    e = Array("$","0","A","B","C","a","b","c","��","��")
    assertAllElements e, a.sort(True)
    
    e = Array("��","��","c","b","a","C","B","A","0","$")
    assertAllElements e, a.sort(False)
End Sub
Sub Test_clsCmArray_sort_Empty
    Dim a
    Set a = new_Arr()
    AssertEqual 0, a.length
    
    a.sort(True)
    AssertEqual 0, a.length
    
    a.sort(False)
    AssertEqual 0, a.length
End Sub

'###################################################################################################
'clsCmArray.sortUsing()
Sub Test_clsCmArray_sortUsing_Num
    Dim e,d,a
    d = Array(5,2,9,6,4,8,7,3,0,1)
    Set a = new_ArrOf(d)
    
    e = Array(0,1,2,3,4,5,6,7,8,9)
    assertAllElements e, a.sortUsing(new_Func("(c,n)=>c>n"))
    
    e = Array(9,8,7,6,5,4,3,2,1,0)
    assertAllElements e, a.sortUsing(new_Func("(c,n)=>c<n"))
End Sub
Sub Test_clsCmArray_sortUsing_Various
    Dim e,d,a
    d = Array("C","$","b","��","a","B","��","A","c","0")
    Set a = new_ArrOf(d)
    
    e = Array("$","0","A","B","C","a","b","c","��","��")
    assertAllElements e, a.sortUsing(new_Func("(c,n)=>c>n"))
    
    e = Array("��","��","c","b","a","C","B","A","0","$")
    assertAllElements e, a.sortUsing(new_Func("(c,n)=>c<n"))
End Sub
Sub Test_clsCmArray_sortUsing_Empty
    Dim a
    Set a = new_Arr()
    AssertEqual 0, a.length
    
    a.sortUsing(new_Func("(c,n)=>c>n"))
    AssertEqual 0, a.length
    
    a.sortUsing(new_Func("(c,n)=>c<n"))
    AssertEqual 0, a.length
End Sub

'###################################################################################################
'clsCmArray.splice
Sub Test_clsCmArray_splice_Normal
    Dim e,e1,a,a1,d
    d = Array(1,2,3,4,5,6,7,8)
    Set a = new_ArrOf(d)
    
    e = Array(1,4,5,6,7,8)
    e1 = Array(2,3)
    Set a1 = a.splice(1,2,Empty)
    assertAllElements e, a
    assertAllElements e1, a1
    
    e = Array(1,5,6,7,8)
    e1 = Array(4)
    Set a1 = a.splice(1,1,Nothing)
    assertAllElements e, a
    assertAllElements e1, a1
    
    e = Array(1,2,3,5,6,7,8)
    Set a1 = a.splice(1,0,Array(2,3))
    assertAllElements e, a
    AssertEqual 0, a1.length
End Sub
Sub Test_clsCmArray_splice_Limit_Upper
    Dim e,e1,a,a1,d
    d = Array(1,2,3,4,5,6,7,8)
    
    Set a = new_ArrOf(d)
    e = d
    Set a1 = a.splice(7,0,Empty)
    assertAllElements e, a
    AssertEqual 0, a1.length
    
    Set a = new_ArrOf(d)
    e = Array(1,2,3,4,5,6,7)
    e1 = Array(8)
    Set a1 = a.splice(-1,1,Nothing)
    assertAllElements e, a
    assertAllElements e1, a1
    
    Set a = new_ArrOf(d)
    e = d
    Set a1 = a.splice(8,1,vbNullString)
    assertAllElements e, a
    AssertEqual 0, a1.length
    
    Set a = new_ArrOf(d)
    e = Array(1,2,3,4,5,6,7,11,12)
    e1 = Array(8)
    Set a1 = a.splice(7,2,Array(11,12))
    assertAllElements e, a
    assertAllElements e1, a1
    
    Set a = new_ArrOf(d)
    e = Array(1,2,3,4,5,6,7,8,21,22,23)
    Set a1 = a.splice(8,1,Array(21,22,23))
    assertAllElements e, a
    AssertEqual 0, a1.length
End Sub
Sub Test_clsCmArray_splice_Limit_Lower
    Dim e,e1,a,a1,d
    d = Array(1,2,3,4,5,6,7,8)
    
    Set a = new_ArrOf(d)
    e = d
    Set a1 = a.splice(0,0,Empty)
    assertAllElements e, a
    AssertEqual 0, a1.length
    
    Set a = new_ArrOf(d)
    e = Array(2,3,4,5,6,7,8)
    e1 = Array(1)
    Set a1 = a.splice(-8,1,Nothing)
    assertAllElements e, a
    assertAllElements e1, a1
    
    Set a = new_ArrOf(d)
    e = Array(11,12,1,2,3,4,5,6,7,8)
    Set a1 = a.splice(-9,0,Array(11,12))
    assertAllElements e, a
    AssertEqual 0, a1.length
    
    Set a = new_ArrOf(d)
    e = Array(21,22,23)
    e1 = d
    Set a1 = a.splice(-8,9,Array(21,22,23))
    assertAllElements e, a
    assertAllElements e1, a1
    
    Set a = new_ArrOf(d)
    e = Array(21,22,23)
    e1 = d
    Set a1 = a.splice(-9,8,Array(21,22,23))
    assertAllElements e, a
    assertAllElements e1, a1
End Sub
Sub Test_clsCmArray_splice_Empty
    Dim a,a1,e
    Set a = new_Arr()
    AssertEqual 0, a.length
    
    Set a1 = a.splice(1,2,Empty)
    AssertEqual 0, a.length
    AssertEqual 0, a1.length
    
    Set a1 = a.splice(1,1,Nothing)
    AssertEqual 0, a.length
    AssertEqual 0, a1.length
    
    e = Array(2,3)
    Set a1 = a.splice(1,0,Array(2,3))
    assertAllElements e, a
    AssertEqual 0, a1.length
    
End Sub

'###################################################################################################
'clsCmArray.toArray()
Sub Test_clsCmArray_toArray
    Dim e()
    Dim a : Set a = new clsCmArray
    
    a.push CreateObject("Scripting.Dictionary")
    a.push "hoge"
    
    AssertEqual VarType(e), VarType(a.toArray())
    AssertEqual TypeName(e), TypeName(a.toArray())
    AssertEqual a.length-1, Ubound(a.toArray())
End Sub
Sub Test_clsCmArray_toArray_AllItems
    Dim e,d,ao,a
    Set ao = new clsCmArray
    d = Array(1,"b",Nothing)
    ao.pushA d

    e = d
    a = ao.toArray()
    assertAllElementsArray e, a

    Dim d2 : d2 = d
    d(1) = "Z"
    ao(1) = "X"
    e = d2
    assertAllElementsArray e, a
End Sub
Sub Test_clsCmArray_toArray_NoData
    Dim e,ao,a
    Set ao = new clsCmArray

    e = Array()
    a = ao.toArray()
    assertAllElementsArray e, a
End Sub

'###################################################################################################
'clsCmArray.toString()
Sub Test_clsCmArray_toString
    Dim e,d,a
    d = Array(1,2,3)
    e = "<clsCmArray>[<Integer>1,<Integer>2,<Integer>3]"
    a = new_ArrOf(d).toString
    AssertEqual e, a
End Sub

'###################################################################################################
'clsCmArray.uniq()
Sub Test_clsCmArray_uniq_ari
    Dim e,d,a
    d = Array(1,2,3,2,3)
    e = Array(1,2,3)
    Set a = new_ArrOf(d).uniq()
    assertAllElements e, a
End Sub
Sub Test_clsCmArray_uniq_nasi
    Dim e,d,a
    d = Array(1,2,3)
    e = Array(1,2,3)
    Set a = new_ArrOf(d).uniq()
    assertAllElements e, a
End Sub
Sub Test_clsCmArray_uniq_Empty
    Dim e,d,a
    d = Array()
    e = Array()
    Set a = new_ArrOf(d).uniq()
    assertAllElements e, a
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
Sub assertAllElementsArray(e,a)
    AssertEqual Ubound(e), Ubound(a)
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
