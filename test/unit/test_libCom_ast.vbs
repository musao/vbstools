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
'ast_argFalse()
Sub Test_ast_argFalse_ok
    dim d : d = Array(False, "Source_ok", "Description_ok")
    ast_argFalse d(0),d(1),d(2)

    AssertWithMessage True, "argFalse_ok"
End Sub
Sub Test_ast_argFalse_ng
    On Error Resume Next
    dim d : d = Array(True, "Source_ng", "Description_ng")
    ast_argFalse d(0),d(1),d(2)

    dim e,a
    e = 8193
    a = Err.Number
    AssertEqualWithMessage e,a,"Number"

    e = d(1)
    a = Err.Source
    AssertEqualWithMessage e,a,"Source"

    e = d(2)
    a = Err.Description
    AssertEqualWithMessage e,a,"Description"
End Sub

'###################################################################################################
'ast_argNotEmpty()
Sub Test_ast_argNotEmpty_ok
    dim d : d = Array("test", "Source_ok", "Description_ok")
    ast_argNotEmpty d(0),d(1),d(2)

    AssertWithMessage True, "argNotEmpty_ok"
End Sub
Sub Test_ast_argNotEmpty_ng
    On Error Resume Next
    dim d : d = Array(Empty, "Source_ng", "Description_ng")
    ast_argNotEmpty d(0),d(1),d(2)

    dim e,a
    e = 8194
    a = Err.Number
    AssertEqualWithMessage e,a,"Number"

    e = d(1)
    a = Err.Source
    AssertEqualWithMessage e,a,"Source"

    e = d(2)
    a = Err.Description
    AssertEqualWithMessage e,a,"Description"
End Sub

'###################################################################################################
'ast_argNotNull()
Sub Test_ast_argNotNull_ok
    dim d : d = Array("test", "Source_ok", "Description_ok")
    ast_argNotNull d(0),d(1),d(2)

    AssertWithMessage True, "argNotNull_ok"
End Sub
Sub Test_ast_argNotNull_ng
    On Error Resume Next
    dim d : d = Array(Null, "Source_ng", "Description_ng")
    ast_argNotNull d(0),d(1),d(2)

    dim e,a
    e = 8195
    a = Err.Number
    AssertEqualWithMessage e,a,"Number"

    e = d(1)
    a = Err.Source
    AssertEqualWithMessage e,a,"Source"

    e = d(2)
    a = Err.Description
    AssertEqualWithMessage e,a,"Description"
End Sub

'###################################################################################################
'ast_argTrue()
Sub Test_ast_argTrue_ok
    dim d : d = Array(True, "Source_ok", "Description_ok")
    ast_argTrue d(0),d(1),d(2)

    AssertWithMessage True, "argTrue_ok"
End Sub
Sub Test_ast_argTrue_ng
    On Error Resume Next
    dim d : d = Array(False, "Source_ng", "Description_ng")
    ast_argTrue d(0),d(1),d(2)

    dim e,a
    e = 8196
    a = Err.Number
    AssertEqualWithMessage e,a,"Number"

    e = d(1)
    a = Err.Source
    AssertEqualWithMessage e,a,"Source"

    e = d(2)
    a = Err.Description
    AssertEqualWithMessage e,a,"Description"
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
    dim d : d = Array(Null, "Source_ok", "Description_ok")
    ast_argNull d(0),d(1),d(2)

    AssertWithMessage True, "argNotNull_ok"
End Sub
Sub Test_ast_argNull_ng
    On Error Resume Next
    dim d : d = Array("Test", "Source_ng", "Description_ng")
    ast_argNull d(0),d(1),d(2)

    dim e,a
    e = 8198
    a = Err.Number
    AssertEqualWithMessage e,a,"Number"

    e = d(1)
    a = Err.Source
    AssertEqualWithMessage e,a,"Source"

    e = d(2)
    a = Err.Description
    AssertEqualWithMessage e,a,"Description"
End Sub

'###################################################################################################
'ast_failure()
Sub Test_ast_failure_ng
    On Error Resume Next
    dim d : d = Array("Source_ng", "Description_ng")
    ast_failure d(0),d(1)

    dim e,a
    e = 8199
    a = Err.Number
    AssertEqualWithMessage e,a,"Number"

    e = d(0)
    a = Err.Source
    AssertEqualWithMessage e,a,"Source"

    e = d(1)
    a = Err.Description
    AssertEqualWithMessage e,a,"Description"
End Sub


' Local Variables:
' mode: Visual-Basic
' indent-tabs-mode: nil
' End:
