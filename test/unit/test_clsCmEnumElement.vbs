' clsCmBroker.vbs: test.
' @import ../../lib/clsAdptFile.vbs
' @import ../../lib/clsCmArray.vbs
' @import ../../lib/clsCmBroker.vbs
' @import ../../lib/clsCmBufferedReader.vbs
' @import ../../lib/clsCmBufferedWriter.vbs
' @import ../../lib/clsCmCalendar.vbs
' @import ../../lib/clsCmCharacterType.vbs
' @import ../../lib/clsCmCssGenerator.vbs
' @import ../../lib/clsCmEnumElement.vbs
' @import ../../lib/clsCmHtmlGenerator.vbs
' @import ../../lib/clsCmReturnValue.vbs
' @import ../../lib/clsCompareExcel.vbs
' @import ../../lib/libCom.vbs
Option Explicit

'###################################################################################################
'clsCmEnumElement
Sub Test_clsCmEnumElement
    Dim a : Set a = new clsCmEnumElement
    AssertEqual 0, VarType(a)
    AssertEqual "clsCmEnumElement", TypeName(a)
End Sub

'###################################################################################################
'clsCmEnumElement.code()
Sub Test_clsCmEnumElement_code
    Dim d : d = Array("EnumTest", "TEST", 100)
    Dim ao : Set ao = (new clsCmEnumElement).thisIs(d(0),d(1),d(2))

    Dim e : e = d(2)
    Dim a
    a = ao
    AssertEqualWithMessage e, a, "Default"
    a = ao.code()
    AssertEqualWithMessage e, a, "code()"
End Sub
Sub Test_clsCmEnumElement_code_InitialValue
    Dim ao : Set ao = (new clsCmEnumElement)
    Dim e : e = Empty
    Dim a
    a = ao
    AssertEqualWithMessage e, a, "Default"
    a = ao.code()
    AssertEqualWithMessage e, a, "code()"
End Sub

'###################################################################################################
'clsCmEnumElement.kind()
Sub Test_clsCmEnumElement_kind
    Dim d : d = Array("EnumTest", "TEST", 100)
    Dim ao : Set ao = (new clsCmEnumElement).thisIs(d(0),d(1),d(2))

    Dim e : e = d(0)
    Dim a : a = ao.kind()
    AssertEqualWithMessage e, a, "kind()"
End Sub
Sub Test_clsCmEnumElement_kind_InitialValue
    Dim ao : Set ao = (new clsCmEnumElement)
    Dim e : e = Empty
    Dim a : a = ao.kind()
    AssertEqualWithMessage e, a, "kind()"
End Sub

'###################################################################################################
'clsCmEnumElement.name()
Sub Test_clsCmEnumElement_name
    Dim d : d = Array("EnumTest", "TEST", 100)
    Dim ao : Set ao = (new clsCmEnumElement).thisIs(d(0),d(1),d(2))

    Dim e : e = d(1)
    Dim a : a = ao.name()
    AssertEqualWithMessage e, a, "name()"
End Sub
Sub Test_clsCmEnumElement_name_InitialValue
    Dim ao : Set ao = (new clsCmEnumElement)
    Dim e : e = Empty
    Dim a : a = ao.name()
    AssertEqualWithMessage e, a, "name()"
End Sub

'###################################################################################################
'clsCmEnumElement.toString()
Sub Test_clsCmEnumElement_toString
    Dim d : d = Array("EnumTest", "TEST", 0)
    Dim ao : Set ao = (new clsCmEnumElement).thisIs(d(0),d(1),d(2))
    Dim e : e = "<" & TypeName(ao) & ">(" & cf_toString(d(2)) & ":" & cf_toString(d(1)) & " of " & cf_toString(d(0)) & ")"
    Dim a : a = ao.toString()
    AssertEqualWithMessage e, a, "toString()"
End Sub
Sub Test_clsCmEnumElement_toString_Initial
    Dim ao : Set ao = (new clsCmEnumElement)
    Dim e : e = "<" & TypeName(ao) & ">(" & cf_toString(Empty) & ":" & cf_toString(Empty) & " of " & cf_toString(Empty) & ")"
    Dim a : a = ao.toString()
    AssertEqualWithMessage e, a, "toString()"
End Sub

'###################################################################################################
'clsCmEnumElement.compareTo()
Sub Test_clsCmEnumElement_compareTo_ok
    Dim kind,name,code
    kind="kind":name="name":code=10
    Dim ao : Set ao = (new clsCmEnumElement).thisIs(kind,name,code)

    Dim data : data = Array( _
        Array((new clsCmEnumElement).thisIs(kind,name,9),1) _
        , Array((new clsCmEnumElement).thisIs(kind,name,code),0) _
        , Array((new clsCmEnumElement).thisIs(kind,name,11),-1) _
        )

    Dim i,d,a,e
    For i=0 To Ubound(data)
        Set d = data(i)(0)
        e = data(i)(1)
        a = ao.compareTo(d)
        AssertEqualWithMessage e, a, "i=" & i
    Next
End Sub
Sub Test_clsCmEnumElement_compareTo_ng
    Dim kind,name,code
    kind="kind":name="name":code=10
    Dim ao : Set ao = (new clsCmEnumElement).thisIs(kind,name,code)

    Dim sou,dis
    sou="clsCmEnumElement+compareTo()":dis="The type of the argument is different"
    Dim data : data = Array( _
        Array((new clsCmEnumElement).thisIs("kind2",name,code),Array(sou,dis)) _
        , Array((new clsCmEnumElement).thisIs(kind,"name2",code),Array(sou,dis)) _
        , Array(CreateObject("Scripting.Dictionary"),Array(sou,dis)) _
        )

    On Error Resume Next
    Dim i,oup,d,a,e
    For i=0 To Ubound(data)
        Set d = data(i)(0)
        oup = data(i)(1)
        ao.compareTo(d)

        e = oup(0)
        a = Err.Source
        AssertEqualWithMessage e,a,"i=" & i & " Source"
    
        e = oup(1)
        a = Err.Description
        AssertEqualWithMessage e,a,"i=" & i & " Description"
    Next
End Sub

'###################################################################################################
'clsCmEnumElement.equals()
Sub Test_clsCmEnumElement_equals
    Dim kind,name,code
    kind="kind":name="name":code=10
    Dim ao : Set ao = (new clsCmEnumElement).thisIs(kind,name,code)

    Dim data : data = Array( _
        Array((new clsCmEnumElement).thisIs(kind,name,code),True) _
        , Array((new clsCmEnumElement).thisIs(kind,name,11),False) _
        , Array((new clsCmEnumElement).thisIs("kind2",name,code),False) _
        , Array((new clsCmEnumElement).thisIs(kind,"name2",code),True) _
        , Array((new clsCmEnumElement).thisIs(kind,name,9),False) _
        , Array(CreateObject("Scripting.Dictionary"),False) _
        )

    Dim i,d,a,e
    For i=0 To Ubound(data)
        Set d = data(i)(0)
        e = data(i)(1)
        a = ao.equals(d)
        AssertEqualWithMessage e, a, "i=" & i
    Next
End Sub


'###################################################################################################
'clsCmEnumElement.thisIs()
Sub Test_clsCmEnumElement_thisIs_Err
    Dim d : d = Array("EnumTest", "TEST", 100)
    Dim ao : Set ao = (new clsCmEnumElement).thisIs(d(0),d(1),d(2))

    Dim d2 : d2 = Array("EnumTest2", "TEST2", 200)
    On Error Resume Next
    ao.thisIs d2(0),d2(1),d2(2)

    dim e,a
    e = d(0)
    a = ao.kind
    AssertEqualWithMessage e,a,"kind"

    e = d(1)
    a = ao.name
    AssertEqualWithMessage e,a,"name"

    e = d(2)
    a = ao.code
    AssertEqualWithMessage e,a,"code"

    e = "clsCmEnumElement+thisIs()"
    a = Err.Source
    AssertEqualWithMessage e,a,"Source"

    e = "Value already set"
    a = Err.Description
    AssertEqualWithMessage e,a,"Description"
End Sub

' Local Variables:
' mode: Visual-Basic
' indent-tabs-mode: nil
' End:
