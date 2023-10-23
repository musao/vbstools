' clsCmCalendar.vbs: test.
' @import ../lib/clsCmArray.vbs
' @import ../lib/clsCmBroker.vbs
' @import ../lib/clsCmBufferedReader.vbs
' @import ../lib/clsCmBufferedWriter.vbs
' @import ../lib/clsCmCalendar.vbs
' @import ../lib/clsCmHtmlGenerator.vbs
' @import ../lib/clsCompareExcel.vbs
' @import ../lib/clsFsBase.vbs
' @import ../lib/libCom.vbs

Option Explicit

'###################################################################################################
'clsCmHtmlGenerator
Sub Test_clsCmBroker
    Dim a : Set a = new clsCmHtmlGenerator
    AssertEqual 9, VarType(a)
    AssertEqual "clsCmHtmlGenerator", TypeName(a)
End Sub

'###################################################################################################
'clsCmHtmlGenerator.attribute/addAttribute()
Sub Test_clsCmHtmlGenerator_attribute_addAttribute_FirstTime
    Dim ao,a,ek1,ev1
    Set ao = new clsCmHtmlGenerator
    
    ek1 = "hoge" : ev1 = "fuga"
    ao.addAttribute ek1,ev1
    a = ao.attribute
    AssertEqualWithMessage 0, Ubound(a), "Ubound"
    AssertEqualWithMessage ek1, a(0).Item("key"), "key1"
    AssertEqualWithMessage ev1, a(0).Item("value"), "value1"
End Sub
Sub Test_clsCmHtmlGenerator_attribute_addAttribute_SecondTimes
    Dim ao,a,ek1,ek2,ev1,ev2
    Set ao = new clsCmHtmlGenerator
    
    ek1 = "hoge" : ev1 = "fuga"
    ao.addAttribute ek1,ev1
    ek2 = "foo" : ev2 = Empty
    ao.addAttribute ek2,ev2
    a = ao.attribute
    AssertEqualWithMessage 1, Ubound(a), "Ubound"
    AssertEqualWithMessage ek1, a(0).Item("key"), "key1"
    AssertEqualWithMessage ev1, a(0).Item("value"), "value1"
    AssertEqualWithMessage ek2, a(1).Item("key"), "key2"
    AssertEqualWithMessage ev2, a(1).Item("value"), "value2"
End Sub

'###################################################################################################
'clsCmHtmlGenerator.element()
Sub Test_clsCmHtmlGenerator_element
    Dim ao,a,d,e
    Set ao = new clsCmHtmlGenerator

    e = Empty
    a = ao.element
    AssertEqualWithMessage e, a, "1-1"

    d = "hoge"
    e = d
    ao.element = d
    a = ao.element
    AssertEqualWithMessage e, a, "1-2"

    d = "fuga"
    e = d
    ao.element = d
    a = ao.element
    AssertEqualWithMessage e, a, "1-3"
End Sub

'###################################################################################################
'clsCmHtmlGenerator.generate()
Sub Test_clsCmHtmlGenerator_generate_ElementOnly
    Dim ao,a,d,e
    Set ao = new clsCmHtmlGenerator
    
    d = "hoge"
    e = "<hoge />"
    ao.element = d
    a = ao.generate

    AssertEqualWithMessage e, a, "1"
End Sub
Sub Test_clsCmHtmlGenerator_generate_ElementAndAttribute
    Dim ao,a,de,dak1,dak2,dav1,dav2,e
    Set ao = new clsCmHtmlGenerator
    
    de = "hoge"
    dak1 = "foo" : dav1 = "bar"
    e = "<hoge foo=" & Chr(34) & "bar" & Chr(34) & " />"
    ao.element = de
    ao.addAttribute dak1,dav1
    a = ao.generate
    AssertEqualWithMessage e, a, "1"
    
    de = "hoge"
    dak2 = "woo" : dav2 = Empty
    e = "<hoge foo=" & Chr(34) & "bar" & Chr(34) & " woo />"
    ao.element = de
    ao.addAttribute dak2,dav2
    a = ao.generate
    AssertEqualWithMessage e, a, "2"
End Sub
Sub Test_clsCmHtmlGenerator_generate_Err
    Dim ao
    Set ao = new clsCmHtmlGenerator

    On Error Resume Next
    ao.generate

    AssertEqualWithMessage 17, Err.Number, "Err.Number"
    AssertEqualWithMessage "要素がないHTMLタグは生成できません。", Err.Description, "Err.Description"
End Sub

' Local Variables:
' mode: Visual-Basic
' indent-tabs-mode: nil
' End:
