' libCom.vbs: ast_* procedure test.
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
Dim data : data = Array ( _
            "abc" _
            , 123 _
            , True _
            , False _
            , Null _
            , Empty _
            , vbNullString _
            , CreateObject("Scripting.Dictionary") _
            , Nothing _
            )

'###################################################################################################
'ast_argFalse()
Sub Test_ast_argFalse_ok
    ast_ok getref("ast_argFalse"), ast_filter(data,False,True)
End Sub
Sub Test_ast_argFalse_ng
    ast_ng getref("ast_argFalse"), ast_filter(data,False,False), 8193
End Sub

'###################################################################################################
'ast_argNotEmpty()
Sub Test_ast_argNotEmpty_ok
    ast_ok getref("ast_argNotEmpty"), ast_filter(data,Empty,False)
End Sub
Sub Test_ast_argNotEmpty_ng
    ast_ng getref("ast_argNotEmpty"), ast_filter(data,Empty,True), 8194
End Sub

'###################################################################################################
'ast_argNotNull()
Sub Test_ast_argNotNull_ok
    ast_ok getref("ast_argNotNull"), ast_filter(data,Null,False)
End Sub
Sub Test_ast_argNotNull_ng
    ast_ng getref("ast_argNotNull"), ast_filter(data,Null,True), 8195
End Sub

'###################################################################################################
'ast_argTrue()
Sub Test_ast_argTrue_ok
    ast_ok getref("ast_argTrue"), ast_filter(data,True,True)
End Sub
Sub Test_ast_argTrue_ng
    ast_ng getref("ast_argTrue"), ast_filter(data,True,False), 8196
End Sub

'###################################################################################################
'ast_argsAreSame()
Sub Test_ast_argsAreSame_ok
    dim d : d = Array("A", "A", "Source_ok", "Description_ok")
    ast_argsAreSame d(0),d(1),d(2),d(3)

    AssertWithMessage True, "argTrue_ok"
End Sub
Sub Test_ast_argsAreSame_ng
    On Error Resume Next
    dim d : d = Array("A", "B", "Source_ng", "Description_ng")
    ast_argsAreSame d(0),d(1),d(2),d(3)

    dim e,a
    e = 8197
    a = Err.Number
    AssertEqualWithMessage e,a,"Number"

    e = d(2)
    a = Err.Source
    AssertEqualWithMessage e,a,"Source"

    e = d(3)
    a = Err.Description
    AssertEqualWithMessage e,a,"Description"
End Sub

'###################################################################################################
'ast_argNull()
Sub Test_ast_argNull_ok
    ast_ok getref("ast_argNull"), ast_filter(data,Null,True)
End Sub
Sub Test_ast_argNull_ng
    ast_ng getref("ast_argNull"), ast_filter(data,Null,False), 8198
End Sub

'###################################################################################################
'ast_failure()
Sub Test_ast_failure_ng
    dim an,ac,ad
    On Error Resume Next
    ast_failure "Source_ng","Description_ng"
    an=Err.Number
    ac=Err.Source
    ad=Err.Description

    e = 8199
    AssertEqualWithMessage e,an,"Number"
    e = "Source_ng"
    AssertEqualWithMessage e,ac,"Source"
    e = "Description_ng"
    AssertEqualWithMessage e,ad,"Description"
End Sub

'###################################################################################################
'ast_argEmpty()
Sub Test_ast_argEmpty_ok
    ast_ok getref("ast_argEmpty"), ast_filter(data,Empty,True)
End Sub
Sub Test_ast_argEmpty_ng
    ast_ng getref("ast_argEmpty"), ast_filter(data,Empty,False), 8200
End Sub

'###################################################################################################
'ast_argNotNothing()
Sub Test_ast_argNotNothing_ok
    ast_ok getref("ast_argNotNothing"), ast_filter(data,Nothing,False)
End Sub
Sub Test_ast_argNotNothing_ng
    ast_ng getref("ast_argNotNothing"), ast_filter(data,Nothing,True), 8201
End Sub

'###################################################################################################
'ast_argNothing()
Sub Test_ast_argNothing_ok
    ast_ok getref("ast_argNothing"), ast_filter(data,Nothing,True)
End Sub
Sub Test_ast_argNothing_ng
    ast_ng getref("ast_argNothing"), ast_filter(data,Nothing,False), 8202
End Sub



'###################################################################################################
'common
Sub ast_ok(f,d)
    dim i
    For Each i In d
        f i, "Source_ok", "Description_ok"
        AssertWithMessage True, "ok "&cf_toString(i)
    Next
End Sub
Sub ast_ng(f,d,n)
    dim i,e,an,ac,ad
    For Each i In d
        On Error Resume Next

        f i,"Source_ng","Description_ng"
        an=Err.Number
        ac=Err.Source
        ad=Err.Description

        On Error Goto 0

        e = n
        AssertEqualWithMessage e,an,"Number i="&cf_toString(i)
        e = "Source_ng"
        AssertEqualWithMessage e,ac,"Source i="&cf_toString(i)
        e = "Description_ng"
        AssertEqualWithMessage e,ad,"Description i="&cf_toString(i)
    Next
End Sub
Function ast_filter(ar,tg,flg)
    Dim rt,i
    For Each i In ar
        If cf_isSame(i,tg)=flg Then cf_push rt, i
    Next
    ast_filter = rt
End Function


' Local Variables:
' mode: Visual-Basic
' indent-tabs-mode: nil
' End:
