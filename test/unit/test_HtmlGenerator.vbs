' HtmlGenerator.vbs: test.
' @import ../../lib/com/FileSystemProxy.vbs
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
'HtmlGenerator
Sub Test_HtmlGenerator
    Dim a : Set a = new HtmlGenerator
    AssertEqual 9, VarType(a)
    AssertEqual "HtmlGenerator", TypeName(a)
End Sub

'###################################################################################################
'HtmlGenerator.attribute/addAttribute()
Sub Test_HtmlGenerator_attribute_addAttribute_FirstTime
    Dim ao,a,ek1,ev1
    Set ao = new HtmlGenerator
    
    ek1 = "hoge" : ev1 = "fuga"
    ao.addAttribute ek1,ev1
    a = ao.attribute
    AssertEqualWithMessage 0, Ubound(a), "Ubound"
    AssertEqualWithMessage ek1, a(0).Item("key"), "key1"
    AssertEqualWithMessage ev1, a(0).Item("value"), "value1"
End Sub
Sub Test_HtmlGenerator_attribute_addAttribute_SecondTimes
    Dim ao,a,ek1,ek2,ev1,ev2
    Set ao = new HtmlGenerator
    
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
'HtmlGenerator.content/addcontent()
Sub Test_HtmlGenerator_content_addcontent_FirstTime
    Dim ao,a,d1,e1
    Set ao = new HtmlGenerator
    
    d1 = "hoge"
    e1 = d1
    ao.addcontent d1
    a = ao.content
    AssertEqualWithMessage 0, Ubound(a), "Ubound"
    AssertEqualWithMessage e1, a(0), "1"
End Sub
Sub Test_HtmlGenerator_content_addcontent_SecondTimes
    Dim ao,a,d1,d2,e1,e2
    Set ao = new HtmlGenerator
    
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
'HtmlGenerator.element()
Sub Test_HtmlGenerator_element
    Dim ao,a,d,e
    Set ao = new HtmlGenerator

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
'Sub Test_HtmlGenerator_element_Err
'    Dim ao,d
'    Set ao = new HtmlGenerator
'
'    On Error Resume Next
'    d = "Ｈｏｇｅ"
'    ao.element = d
'
'    AssertEqual 1032, Err.Number
'    AssertEqual "要素（element）には半角以外の文字を指定できません。", Err.Description
'End Sub

'###################################################################################################
'HtmlGenerator.generate()
Sub Test_HtmlGenerator_generate_ElementOnly
    Dim ao,a,d,e
    Set ao = new HtmlGenerator
    
    d = "hoge"
    e = "<hoge />"
    ao.element = d
    a = ao.generate

    AssertEqualWithMessage e, a, "1"
End Sub
Sub Test_HtmlGenerator_generate_ElementAndAttribute
    Dim ao,a,de,dak1,dak2,dav1,dav2,e
    Set ao = new HtmlGenerator
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
Sub Test_HtmlGenerator_generate_ElementAndContent
    Dim ao,a,de,dc1,dc2,e
    Set ao = new HtmlGenerator
    de = "hoge" : ao.element = de

    dc1 = "fuga"
    e = _
        "<hoge>" & vbNewLine _
        & "  fuga" & vbNewLine _
        & "</hoge>"
    ao.addContent dc1
    a = ao.generate
    AssertEqualWithMessage e, a, "1"
    
    Set dc2 = new HtmlGenerator
    dc2.element = "foo"
    dc2.addContent "bar"
    e = _
        "<hoge>" & vbNewLine _
        & "  fuga" & vbNewLine _
        & "  <foo>" & vbNewLine _
        & "    bar" & vbNewLine _
        & "  </foo>" & vbNewLine _
        & "</hoge>"
    ao.addContent dc2
    a = ao.generate
    AssertEqualWithMessage e, a, "2"
End Sub
Sub Test_HtmlGenerator_generate_All
    Dim ao,a,de,dx,e
    Set dx = new HtmlGenerator
    dx.element = "fuga2"
    dx.addAttribute "foo2","bar2"
    dx.addAttribute "woo2",Empty
    dx.addContent "wao2"

    Set ao = new HtmlGenerator
    ao.element = "hoge"
    ao.addAttribute "foo1","bar1"
    ao.addAttribute "woo1",Empty
    ao.addContent "wao1"
    ao.addContent dx

    e = _ 
        "<hoge foo1="&Chr(34)&"bar1"&Chr(34)&" woo1>" & vbNewLine _
        & "  wao1" & vbNewLine _
        & "  <fuga2 foo2="&Chr(34)&"bar2"&Chr(34)&" woo2>" & vbNewLine _
        & "    wao2" & vbNewLine _
        & "  </fuga2>" & vbNewLine _
        & "</hoge>"
    a = ao.generate
    AssertEqualWithMessage e, a, "1"
End Sub
Sub Test_HtmlGenerator_generate_Err
    Dim ao
    Set ao = new HtmlGenerator

    On Error Resume Next
    ao.generate()

    AssertEqualWithMessage "HtmlGenerator+generate()", Err.Source, "Err.Source"
    AssertEqualWithMessage "HTML tags without elements cannot be generated.", Err.Description, "Err.Description"
End Sub

'###################################################################################################
'HtmlGenerator.generate()EntityReference
Sub Test_HtmlGenerator_generate_EntityReference
    Dim ao,a,d,dc,e,i
    
    d = Array( _
        new_DicOf(Array(  "No",1 ,"Cont","'fuga"   ,"Expected","&#39;fuga")) _
        , new_DicOf(Array("No",2 ,"Cont","fu""ga"   ,"Expected","fu&quot;ga")) _
        , new_DicOf(Array("No",3 ,"Cont","fuga&"   ,"Expected","fuga&amp;")) _
        , new_DicOf(Array("No",4 ,"Cont","<fuga"   ,"Expected","&lt;fuga")) _
        , new_DicOf(Array("No",5 ,"Cont","fuga>"   ,"Expected","fuga&gt;")) _
        , new_DicOf(Array("No",6 ,"Cont","<'fu""ga&>"   ,"Expected","&lt;&#39;fu&quot;ga&amp;&gt;")) _
        )
    
    For Each i In d
        dc = i.Item("Cont")
        e = _
            "<hoge>" & vbNewLine _
            & "  " & i.Item("Expected") & vbNewLine _
            & "</hoge>"

        Set ao = new HtmlGenerator
        ao.element = "hoge"
        ao.addContent dc
        a = ao.generate

        AssertEqualWithMessage e, a, "No="&i.Item("No")&", Cont="&dc
    Next
End Sub

'###################################################################################################
'HtmlGenerator.toString()
Sub Test_HtmlGenerator_toString
    Dim ao,a,d,e
    Set ao = new HtmlGenerator
    
    d = "hoge"
    ao.element = d 
    e = ao.generate()
    a = ao.toString()

    AssertEqualWithMessage e, a, "1"
End Sub
Sub Test_HtmlGenerator_toString_Err
    Dim ao
    Set ao = new HtmlGenerator

    On Error Resume Next
    ao.toString()

    AssertEqualWithMessage "HtmlGenerator+toString()", Err.Source, "Err.Source"
    AssertEqualWithMessage "HTML tags without elements cannot be generated.", Err.Description, "Err.Description"
End Sub

' Local Variables:
' mode: Visual-Basic
' indent-tabs-mode: nil
' End:
