' CssGenerator.vbs: test.
' @import ../../lib/com/FileProxy.vbs
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
'CssGenerator
Sub Test_CssGenerator
    Dim a : Set a = new CssGenerator
    AssertEqual 9, VarType(a)
    AssertEqual "CssGenerator", TypeName(a)
End Sub

'###################################################################################################
'CssGenerator.property/addProperty()
Sub Test_CssGenerator_property_addProperty_FirstTime
    Dim ao,a,ek1,ev1
    Set ao = new CssGenerator
    
    ek1 = "hoge" : ev1 = "fuga"
    ao.addProperty ek1,ev1
    a = ao.property
    AssertEqualWithMessage 0, Ubound(a), "Ubound"
    AssertEqualWithMessage ek1, a(0).Item("key"), "key1"
    AssertEqualWithMessage ev1, a(0).Item("value"), "value1"
End Sub
Sub Test_CssGenerator_property_addProperty_SecondTimes
    Dim ao,a,ek1,ek2,ev1,ev2
    Set ao = new CssGenerator
    
    ek1 = "hoge" : ev1 = "fuga"
    ao.addProperty ek1,ev1
    ek2 = "foo" : ev2 = Empty
    ao.addProperty ek2,ev2
    a = ao.property
    AssertEqualWithMessage 1, Ubound(a), "Ubound"
    AssertEqualWithMessage ek1, a(0).Item("key"), "key1"
    AssertEqualWithMessage ev1, a(0).Item("value"), "value1"
    AssertEqualWithMessage ek2, a(1).Item("key"), "key2"
    AssertEqualWithMessage ev2, a(1).Item("value"), "value2"
End Sub

'###################################################################################################
'CssGenerator.selector()
Sub Test_CssGenerator_selector
    Dim ao,a,d,e
    Set ao = new CssGenerator

    e = Empty
    a = ao.selector
    AssertEqualWithMessage e, a, "1-1"

    d = "hoge"
    e = d
    ao.selector = d
    a = ao.selector
    AssertEqualWithMessage e, a, "1-2"

    d = "fuga"
    e = d
    ao.selector = d
    a = ao.selector
    AssertEqualWithMessage e, a, "1-3"
End Sub
'Sub Test_CssGenerator_selector_Err
'    Dim ao,a,d
'    Set ao = new CssGenerator
'
'    On Error Resume Next
'    d = "Ｈｏｇｅ"
'    ao.selector = d
'
'    AssertEqual 1032, Err.Number
'    AssertEqual "セレクタには半角以外の文字を指定できません。", Err.Description
'End Sub

'###################################################################################################
'CssGenerator.generate()
Sub Test_CssGenerator_generate_SelectorOnly
    Dim ao,a,d,e
    Set ao = new CssGenerator
    
    d = "hoge"
    e = "hoge {" & vbNewLine & "}"
    ao.selector = d
    a = ao.generate

    AssertEqualWithMessage e, a, "1"
End Sub
Sub Test_CssGenerator_generate_All
    Dim ao,a,de,dak1,dak2,dav1,dav2,e
    Set ao = new CssGenerator
    de = "hoge" : ao.selector = de

    dak1 = "foo" : dav1 = "bar"
    e = _
        "hoge {" & vbNewLine _
        & "  foo : bar ;" & vbNewLine _
        & "}"
    ao.addProperty dak1,dav1
    a = ao.generate
    AssertEqualWithMessage e, a, "1"
    
    dak2 = "woo" : dav2 = "woo"
    e = _
        "hoge {" & vbNewLine _
        & "  foo : bar ;" & vbNewLine _
        & "  woo : woo ;" & vbNewLine _
        & "}"
    ao.addProperty dak2,dav2
    a = ao.generate
    AssertEqualWithMessage e, a, "2"
End Sub
Sub Test_CssGenerator_generate_Err
    Dim ao
    Set ao = new CssGenerator

    On Error Resume Next
    ao.generate()

    AssertEqualWithMessage "CssGenerator+generate()", Err.Source, "Err.Source"
    AssertEqualWithMessage "CSS without selectors cannot be generated.", Err.Description, "Err.Description"
End Sub

'###################################################################################################
'CssGenerator.toString()
Sub Test_CssGenerator_toString
    Dim ao,a,d,e
    Set ao = new CssGenerator
    
    d = "hoge"
    ao.selector = d
    e = ao.generate()
    a = ao.toString()

    AssertEqualWithMessage e, a, "1"
End Sub
Sub Test_CssGenerator_toString_Err
    Dim ao
    Set ao = new CssGenerator

    On Error Resume Next
    ao.toString()

    AssertEqualWithMessage "CssGenerator+toString()", Err.Source, "Err.Source"
    AssertEqualWithMessage "CSS without selectors cannot be generated.", Err.Description, "Err.Description"
End Sub

' Local Variables:
' mode: Visual-Basic
' indent-tabs-mode: nil
' End:
