' clsCmBroker.vbs: test.
' @import ../../lib/clsAdptFile.vbs
' @import ../../lib/clsCmArray.vbs
' @import ../../lib/clsCmBroker.vbs
' @import ../../lib/clsCmBufferedReader.vbs
' @import ../../lib/clsCmBufferedWriter.vbs
' @import ../../lib/clsCmCalendar.vbs
' @import ../../lib/clsCmCharacterType.vbs
' @import ../../lib/clsCmCssGenerator.vbs
' @import ../../lib/clsCmReadOnlyObject.vbs
' @import ../../lib/clsCmHtmlGenerator.vbs
' @import ../../lib/clsCmReturnValue.vbs
' @import ../../lib/clsCompareExcel.vbs
' @import ../../lib/libCom.vbs
Option Explicit

'###################################################################################################
'clsCmReadOnlyObject
Sub Test_clsCmReadOnlyObject
    Dim a : Set a = new clsCmReadOnlyObject
    AssertEqual 0, VarType(a)
    AssertEqual "clsCmReadOnlyObject", TypeName(a)
End Sub

'###################################################################################################
'clsCmReadOnlyObject.value()
Sub Test_clsCmReadOnlyObject_value
    Dim d : d = Array("EnumTest", "TEST", 100)
    Dim ao : Set ao = (new clsCmReadOnlyObject).of(d(0),d(1),d(2))

    Dim e : e = d(2)
    Dim a
    a = ao
    AssertEqualWithMessage e, a, "Default"
    a = ao.value()
    AssertEqualWithMessage e, a, "value()"
End Sub
Sub Test_clsCmReadOnlyObject_value_InitialValue
    Dim ao : Set ao = (new clsCmReadOnlyObject)
    Dim e : e = Empty
    Dim a
    a = ao
    AssertEqualWithMessage e, a, "Default"
    a = ao.value()
    AssertEqualWithMessage e, a, "value()"
End Sub

'###################################################################################################
'clsCmReadOnlyObject.parent()
Sub Test_clsCmReadOnlyObject_parent
    Dim d : d = Array(CreateObject("Scripting.Dictionary"), "TEST", 100)
    Dim ao : Set ao = (new clsCmReadOnlyObject).of(d(0),d(1),d(2))

    Dim e : Set e = d(0)
    Dim a : Set a = ao.parent()
    AssertSameWithMessage e, a, "parent()"
End Sub
Sub Test_clsCmReadOnlyObject_parent_InitialValue
    Dim ao : Set ao = (new clsCmReadOnlyObject)
    Dim e : Set e = Nothing
    Dim a : Set a = ao.parent()
    AssertSameWithMessage e, a, "parent()"
End Sub

'###################################################################################################
'clsCmReadOnlyObject.name()
Sub Test_clsCmReadOnlyObject_name
    Dim d : d = Array("EnumTest", "TEST", 100)
    Dim ao : Set ao = (new clsCmReadOnlyObject).of(d(0),d(1),d(2))

    Dim e : e = d(1)
    Dim a : a = ao.name()
    AssertEqualWithMessage e, a, "name()"
End Sub
Sub Test_clsCmReadOnlyObject_name_InitialValue
    Dim ao : Set ao = (new clsCmReadOnlyObject)
    Dim e : e = Empty
    Dim a : a = ao.name()
    AssertEqualWithMessage e, a, "name()"
End Sub

'###################################################################################################
'clsCmReadOnlyObject.toString()
Sub Test_clsCmReadOnlyObject_toString
    Dim d : d = Array("EnumTest", "TEST", 0)
    Dim ao : Set ao = (new clsCmReadOnlyObject).of(d(0),d(1),d(2))
    Dim e : e = "<" & TypeName(ao) & ">{" & cf_toString(d(1)) & ":" & cf_toString(d(2)) & "}"
    Dim a : a = ao.toString()
    AssertEqualWithMessage e, a, "toString()"
End Sub
Sub Test_clsCmReadOnlyObject_toString_Initial
    Dim ao : Set ao = (new clsCmReadOnlyObject)
    Dim e : e = "<" & TypeName(ao) & ">{" & cf_toString(Empty) & ":" & cf_toString(Empty) & "}"
    Dim a : a = ao.toString()
    AssertEqualWithMessage e, a, "toString()"
End Sub

'###################################################################################################
'clsCmReadOnlyObject.compareTo()
Sub Test_clsCmReadOnlyObject_compareTo_ok
    Dim parent,name,value
    Set parent=CreateObject("Scripting.Dictionary"):name="name":value=10
    Dim ao : Set ao = (new clsCmReadOnlyObject).of(parent,name,value)

    Dim data : data = Array( _
        Array((new clsCmReadOnlyObject).of(parent,name,9),1) _
        , Array((new clsCmReadOnlyObject).of(parent,name,value),0) _
        , Array((new clsCmReadOnlyObject).of(parent,name,11),-1) _
        )

    Dim i,d,a,e
    For i=0 To Ubound(data)
        Set d = data(i)(0)
        e = data(i)(1)
        a = ao.compareTo(d)
        AssertEqualWithMessage e, a, "i=" & i
    Next
End Sub
Sub Test_clsCmReadOnlyObject_compareTo_ng
    Dim parent,name,value
    Set parent=CreateObject("Scripting.Dictionary"):name="name":value=10
    Dim ao : Set ao = (new clsCmReadOnlyObject).of(parent,name,value)

    Dim sou,dis
    sou="clsCmReadOnlyObject+compareTo()":dis="The type of the argument is different."
    Dim data : data = Array( _
        Array((new clsCmReadOnlyObject).of(CreateObject("Wscript.Shell"),name,value),Array(sou,dis)) _
        , Array((new clsCmReadOnlyObject).of(parent,"name2",value),Array(sou,dis)) _
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
    Err.Clear
End Sub

'###################################################################################################
'clsCmReadOnlyObject.equals()
Sub Test_clsCmReadOnlyObject_equals
    Dim parent,name,value
    Set parent=CreateObject("Scripting.Dictionary"):name="name":value=10
    Dim ao : Set ao = (new clsCmReadOnlyObject).of(parent,name,value)

    Dim data : data = Array( _
        Array((new clsCmReadOnlyObject).of(parent,name,value),True) _
        , Array((new clsCmReadOnlyObject).of(parent,name,11),False) _
        , Array((new clsCmReadOnlyObject).of(CreateObject("Wscript.Shell"),name,value),False) _
        , Array((new clsCmReadOnlyObject).of(parent,"name2",value),True) _
        , Array((new clsCmReadOnlyObject).of(parent,name,9),False) _
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
'clsCmReadOnlyObject.of()
Sub Test_clsCmReadOnlyObject_is_Err
    Dim d : d = Array(CreateObject("Scripting.Dictionary"), "TEST", 100)
    Dim ao : Set ao = (new clsCmReadOnlyObject).of(d(0),d(1),d(2))

    Dim d2 : d2 = Array(CreateObject("Wscript.Shell"), "TEST2", 200)
    On Error Resume Next
    ao.of d2(0),d2(1),d2(2)

    dim e,a
    Set e = d(0)
    Set a = ao.parent
    AssertSameWithMessage e,a,"parent"

    e = d(1)
    a = ao.name
    AssertEqualWithMessage e,a,"name"

    e = d(2)
    a = ao.value
    AssertEqualWithMessage e,a,"value"

    e = "clsCmReadOnlyObject+of()"
    a = Err.Source
    AssertEqualWithMessage e,a,"Err.Source"

    e = "Value already set."
    a = Err.Description
    AssertEqualWithMessage e,a,"Err.Description"

    Err.Clear
End Sub

' Local Variables:
' mode: Visual-Basic
' indent-tabs-mode: nil
' End:
