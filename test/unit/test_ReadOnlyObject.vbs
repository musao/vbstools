' ReadOnlyObject.vbs: test.
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
'ReadOnlyObject
Sub Test_ReadOnlyObject
    Dim a : Set a = new ReadOnlyObject
    AssertEqual 0, VarType(a)
    AssertEqual "ReadOnlyObject", TypeName(a)
End Sub

'###################################################################################################
'ReadOnlyObject.value()
Sub Test_ReadOnlyObject_value
    Dim d : d = Array("EnumTest", "TEST", 100)
    Dim ao : Set ao = (new ReadOnlyObject).of(d(0),d(1),d(2))

    Dim e : e = d(2)
    Dim a
    a = ao
    AssertEqualWithMessage e, a, "Default"
    a = ao.value()
    AssertEqualWithMessage e, a, "value()"
End Sub
Sub Test_ReadOnlyObject_value_InitialValue
    Dim ao : Set ao = (new ReadOnlyObject)
    Dim e : e = Empty
    Dim a
    a = ao
    AssertEqualWithMessage e, a, "Default"
    a = ao.value()
    AssertEqualWithMessage e, a, "value()"
End Sub

'###################################################################################################
'ReadOnlyObject.parent()
Sub Test_ReadOnlyObject_parent
    Dim d : d = Array(CreateObject("Scripting.Dictionary"), "TEST", 100)
    Dim ao : Set ao = (new ReadOnlyObject).of(d(0),d(1),d(2))

    Dim e : Set e = d(0)
    Dim a : Set a = ao.parent()
    AssertSameWithMessage e, a, "parent()"
End Sub
Sub Test_ReadOnlyObject_parent_InitialValue
    Dim ao : Set ao = (new ReadOnlyObject)
    Dim e : Set e = Nothing
    Dim a : Set a = ao.parent()
    AssertSameWithMessage e, a, "parent()"
End Sub

'###################################################################################################
'ReadOnlyObject.name()
Sub Test_ReadOnlyObject_name
    Dim d : d = Array("EnumTest", "TEST", 100)
    Dim ao : Set ao = (new ReadOnlyObject).of(d(0),d(1),d(2))

    Dim e : e = d(1)
    Dim a : a = ao.name()
    AssertEqualWithMessage e, a, "name()"
End Sub
Sub Test_ReadOnlyObject_name_InitialValue
    Dim ao : Set ao = (new ReadOnlyObject)
    Dim e : e = Empty
    Dim a : a = ao.name()
    AssertEqualWithMessage e, a, "name()"
End Sub

'###################################################################################################
'ReadOnlyObject.toString()
Sub Test_ReadOnlyObject_toString
    Dim d : d = Array("EnumTest", "TEST", 0)
    Dim ao : Set ao = (new ReadOnlyObject).of(d(0),d(1),d(2))
    Dim e : e = "<" & TypeName(ao) & ">{" & cf_toString(d(1)) & ":" & cf_toString(d(2)) & "}"
    Dim a : a = ao.toString()
    AssertEqualWithMessage e, a, "toString()"
End Sub
Sub Test_ReadOnlyObject_toString_Initial
    Dim ao : Set ao = (new ReadOnlyObject)
    Dim e : e = "<" & TypeName(ao) & ">{" & cf_toString(Empty) & ":" & cf_toString(Empty) & "}"
    Dim a : a = ao.toString()
    AssertEqualWithMessage e, a, "toString()"
End Sub

'###################################################################################################
'ReadOnlyObject.compareTo()
Sub Test_ReadOnlyObject_compareTo_ok
    Dim parent,name,value
    Set parent=CreateObject("Scripting.Dictionary"):name="name":value=10
    Dim ao : Set ao = (new ReadOnlyObject).of(parent,name,value)

    Dim data : data = Array( _
        Array((new ReadOnlyObject).of(parent,name,9),1) _
        , Array((new ReadOnlyObject).of(parent,name,value),0) _
        , Array((new ReadOnlyObject).of(parent,name,11),-1) _
        )

    Dim i,d,a,e
    For i=0 To Ubound(data)
        Set d = data(i)(0)
        e = data(i)(1)
        a = ao.compareTo(d)
        AssertEqualWithMessage e, a, "i=" & i
    Next
End Sub
Sub Test_ReadOnlyObject_compareTo_ng
    Dim parent,name,value
    Set parent=CreateObject("Scripting.Dictionary"):name="name":value=10
    Dim ao : Set ao = (new ReadOnlyObject).of(parent,name,value)

    Dim sou,dis
    sou="ReadOnlyObject+compareTo()":dis="The type of the argument is different."
    Dim data : data = Array( _
        Array((new ReadOnlyObject).of(CreateObject("Wscript.Shell"),name,value),Array(sou,dis)) _
        , Array((new ReadOnlyObject).of(parent,"name2",value),Array(sou,dis)) _
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
'ReadOnlyObject.equals()
Sub Test_ReadOnlyObject_equals
    Dim parent,name,value
    Set parent=CreateObject("Scripting.Dictionary"):name="name":value=10
    Dim ao : Set ao = (new ReadOnlyObject).of(parent,name,value)

    Dim data : data = Array( _
        Array((new ReadOnlyObject).of(parent,name,value),True) _
        , Array((new ReadOnlyObject).of(parent,name,11),False) _
        , Array((new ReadOnlyObject).of(CreateObject("Wscript.Shell"),name,value),False) _
        , Array((new ReadOnlyObject).of(parent,"name2",value),True) _
        , Array((new ReadOnlyObject).of(parent,name,9),False) _
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
'ReadOnlyObject.of()
Sub Test_ReadOnlyObject_is_Err
    Dim d : d = Array(CreateObject("Scripting.Dictionary"), "TEST", 100)
    Dim ao : Set ao = (new ReadOnlyObject).of(d(0),d(1),d(2))

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

    e = "ReadOnlyObject+of()"
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
