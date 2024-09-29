' clsCmBufferedReader.vbs: test.
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

' for fso.OpenTextFile
Const ForReading = 1, ForWriting = 2, ForAppending = 8

'###################################################################################################
'clsCmBufferedReader
Sub Test_clsCmBufferedReader
    Dim a : Set a = new clsCmBufferedReader
    AssertEqual 9, VarType(a)
    AssertEqual "clsCmBufferedReader", TypeName(a)
End Sub

'###################################################################################################
'clsCmBufferedReader.readSize()
Sub Test_clsCmBufferedReader_readSize
    Dim a,e
    Set a = new clsCmBufferedReader

    AssertEqual 5000, a.readSize

    e = 1
    a.readSize = e
    AssertEqualWithMessage e, a.readSize, "e="&e
    
    e = "‚P‚O"
    a.readSize = e
'    AssertEqual CLng(e), a.readSize
    AssertEqualWithMessage CDbl(e), a.readSize, "e="&e
    
    e = 2^31-1
    a.readSize = e
'    AssertEqual e, a.readSize
    AssertEqualWithMessage e, a.readSize, "e="&e
End Sub
Sub Test_clsCmBufferedReader_readSize_Err_Zero
    On Error Resume Next
    Dim a,d
    Set a = new clsCmBufferedReader
    
    e = a.readSize
    d = 0
    a.readSize = d
    
    AssertEqualWithMessage "clsCmBufferedReader+readSize() Let", Err.Source, "Err.Source"
    AssertEqualWithMessage "Not a positive integer.", Err.Description, "Err.Description"
    AssertEqualWithMessage e, a.readSize, "readSize"
End Sub
Sub Test_clsCmBufferedReader_readSize_Err_OverLower
    On Error Resume Next
    Dim a,d
    Set a = new clsCmBufferedReader
    
    e = a.readSize
    d = -1*2^31 -1
    a.readSize = d
    
    AssertEqualWithMessage "clsCmBufferedReader+readSize() Let", Err.Source, "Err.Source"
    AssertEqualWithMessage "Not a positive integer.", Err.Description, "Err.Description"
    AssertEqualWithMessage e, a.readSize, "readSize"
End Sub
Sub Test_clsCmBufferedReader_readSize_Err_OverUpper
    On Error Resume Next
    Dim a,d,e
    Set a = new clsCmBufferedReader
    
    e = 100
    a.readSize = e

    d = 2^(2^10)
    a.readSize = d
    
    AssertEqualWithMessage "clsCmBufferedReader+readSize() Let", Err.Source, "Err.Source"
    AssertEqualWithMessage "Not a positive integer.", Err.Description, "Err.Description"
    AssertEqualWithMessage e, a.readSize, "readSize"
End Sub
Sub Test_clsCmBufferedReader_readSize_Err_NotNumeric
    On Error Resume Next
    Dim a,d
    Set a = new clsCmBufferedReader
    
    e = a.readSize

    d = "abc"
    a.readSize = d
    
    AssertEqualWithMessage "clsCmBufferedReader+readSize() Let", Err.Source, "Err.Source"
    AssertEqualWithMessage "Not a positive integer.", Err.Description, "Err.Description"
    AssertEqualWithMessage e, a.readSize, "readSize"
End Sub

'###################################################################################################
'clsCmBufferedReader.textStream()/setTextStream()
Sub Test_clsCmBufferedReader_textStream_setTextStream
    Dim e : Set e = new_Fso().OpenTextFile(WScript.ScriptFullName, ForReading, False, -2)
    Dim a : Set a = New clsCmBufferedReader
    a.setTextStream(e)
    
    AssertSame e, a.textStream
    assertAll takeSnapshot(e), takeSnapshot(a)
    
    e.Close
End Sub
Sub Test_clsCmBufferedReader_textStream_setTextStream_Err_Value
    On Error Resume Next
    Dim a : Set a = New clsCmBufferedReader
    a.setTextStream(vbNullString)
    
    AssertEqualWithMessage "clsCmBufferedReader+setTextStream()", Err.Source, "Err.Source"
    AssertEqualWithMessage "Not a TextStream object.", Err.Description, "Err.Description"
    AssertSameWithMessage Nothing, a.textStream, "textStream"
End Sub
Sub Test_clsCmBufferedReader_textStream_setTextStream_Err_Object
    On Error Resume Next
    Dim a : Set a = New clsCmBufferedReader
    a.setTextStream(new_Dic())
    
    AssertEqualWithMessage "clsCmBufferedReader+setTextStream()", Err.Source, "Err.Source"
    AssertEqualWithMessage "Not a TextStream object.", Err.Description, "Err.Description"
    AssertSameWithMessage Nothing, a.textStream, "textStream"
End Sub





'###################################################################################################
'common
Function takeSnapshot(o)
    Dim ret : Set ret = new_Dic()
    with o
        ret.add "Line", .Line
        ret.add "Column", .Column
        ret.add "AtEndOfLine", .AtEndOfLine
        ret.add "AtEndOfStream", .AtEndOfStream
    end with
    Set takeSnapshot = ret
End Function
Sub assertAll(a,b)
    Dim sKey
    For Each sKey In Array("Line","Column","AtEndOfLine","AtEndOfStream")
        AssertEqualWithMessage a.Item(sKey), b.Item(sKey), sKey
    Next
End Sub


' Local Variables:
' mode: Visual-Basic
' indent-tabs-mode: nil
' End:
