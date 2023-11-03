' clsCmHtmlGenerator.vbs: test.
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
'clsCmHtmlGenerator.content/addcontent()
Sub Test_clsCmHtmlGenerator_content_addcontent_FirstTime
    Dim ao,a,d1,e1
    Set ao = new clsCmHtmlGenerator
    
    d1 = "hoge"
    e1 = d1
    ao.addcontent d1
    a = ao.content
    AssertEqualWithMessage 0, Ubound(a), "Ubound"
    AssertEqualWithMessage e1, a(0), "1"
End Sub
Sub Test_clsCmHtmlGenerator_content_addcontent_SecondTimes
    Dim ao,a,d1,d2,e1,e2
    Set ao = new clsCmHtmlGenerator
    
    d1 = "hoge" : Set d2 = new_Dic()
    e1 = d1 : Set e2 = d2
    ao.addcontent d1
    ao.addcontent d2
    a = ao.content
    AssertEqualWithMessage 1, Ubound(a), "Ubound"
    AssertEqualWithMessage e1, a(0), "1"
    AssertSameWithMessage e2, a(1), "2"
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
    de = "hoge" : ao.element = de

    dak1 = "foo" : dav1 = "bar"
    e = "<hoge foo=" & Chr(34) & "bar" & Chr(34) & " />"
    ao.addAttribute dak1,dav1
    a = ao.generate
    AssertEqualWithMessage e, a, "1"
    
    dak2 = "woo" : dav2 = Empty
    e = "<hoge foo=" & Chr(34) & "bar" & Chr(34) & " woo />"
    ao.addAttribute dak2,dav2
    a = ao.generate
    AssertEqualWithMessage e, a, "2"
End Sub
Sub Test_clsCmHtmlGenerator_generate_ElementAndContent
    Dim ao,a,de,dc1,dc2,e
    Set ao = new clsCmHtmlGenerator
    de = "hoge" : ao.element = de

    dc1 = "fuga"
    e = "<hoge>fuga</hoge>"
    ao.addContent dc1
    a = ao.generate
    AssertEqualWithMessage e, a, "1"
    
    Set dc2 = new clsCmHtmlGenerator
    dc2.element = "foo"
    dc2.addContent "bar"
    e = "<hoge>fuga<foo>bar</foo></hoge>"
    ao.addContent dc2
    a = ao.generate
    AssertEqualWithMessage e, a, "2"
End Sub
Sub Test_clsCmHtmlGenerator_generate_All
    Dim ao,a,de,dx,e
    Set dx = new clsCmHtmlGenerator
    dx.element = "fuga2"
    dx.addAttribute "foo2","bar2"
    dx.addAttribute "woo2",Empty
    dx.addContent "wao2"

    Set ao = new clsCmHtmlGenerator
    ao.element = "hoge"
    ao.addAttribute "foo1","bar1"
    ao.addAttribute "woo1",Empty
    ao.addContent "wao1"
    ao.addContent dx

    e = "<hoge foo1="&Chr(34)&"bar1"&Chr(34)&" woo1>wao1<fuga2 foo2="&Chr(34)&"bar2"&Chr(34)&" woo2>wao2</fuga2></hoge>"
    a = ao.generate
    AssertEqualWithMessage e, a, "1"
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
