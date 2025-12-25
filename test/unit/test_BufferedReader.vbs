' BufferedReader.vbs: test.
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

' @import ../../lib/libEnum.vbs

Option Explicit


'###################################################################################################
'BufferedReader
Sub Test_BufferedReader
    Dim a : Set a = new BufferedReader
    AssertEqual 9, VarType(a)
    AssertEqual "BufferedReader", TypeName(a)
End Sub

'###################################################################################################
'BufferedReader.readSize()
Sub Test_BufferedReader_readSize
    Dim a,e
    Set a = new BufferedReader

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
Sub Test_BufferedReader_readSize_Err_Zero
    On Error Resume Next
    Dim a,d
    Set a = new BufferedReader
    
    e = a.readSize
    d = 0
    a.readSize = d
    
    AssertEqualWithMessage "BufferedReader+readSize() Let", Err.Source, "Err.Source"
    AssertEqualWithMessage "Not a positive integer.", Err.Description, "Err.Description"
    AssertEqualWithMessage e, a.readSize, "readSize"
End Sub
Sub Test_BufferedReader_readSize_Err_OverLower
    On Error Resume Next
    Dim a,d
    Set a = new BufferedReader
    
    e = a.readSize
    d = -1*2^31 -1
    a.readSize = d
    
    AssertEqualWithMessage "BufferedReader+readSize() Let", Err.Source, "Err.Source"
    AssertEqualWithMessage "Not a positive integer.", Err.Description, "Err.Description"
    AssertEqualWithMessage e, a.readSize, "readSize"
End Sub
Sub Test_BufferedReader_readSize_Err_OverUpper
    On Error Resume Next
    Dim a,d,e
    Set a = new BufferedReader
    
    e = 100
    a.readSize = e

    d = 2^(2^10)
    a.readSize = d
    
    AssertEqualWithMessage "BufferedReader+readSize() Let", Err.Source, "Err.Source"
    AssertEqualWithMessage "Not a positive integer.", Err.Description, "Err.Description"
    AssertEqualWithMessage e, a.readSize, "readSize"
End Sub
Sub Test_BufferedReader_readSize_Err_NotNumeric
    On Error Resume Next
    Dim a,d
    Set a = new BufferedReader
    
    e = a.readSize

    d = "abc"
    a.readSize = d
    
    AssertEqualWithMessage "BufferedReader+readSize() Let", Err.Source, "Err.Source"
    AssertEqualWithMessage "Not a positive integer.", Err.Description, "Err.Description"
    AssertEqualWithMessage e, a.readSize, "readSize"
End Sub

'###################################################################################################
'BufferedReader.textStream()/setTextStream()
Sub Test_BufferedReader_textStream_setTextStream
    Dim e : Set e = new_Fso().OpenTextFile(WScript.ScriptFullName, tsMode.FOR_READING, False, tsFormat.USE_DEFAULT)
    Dim a : Set a = New BufferedReader
    a.setTextStream(e)
    
    AssertSame e, a.textStream
    assertAll takeSnapshot(e), takeSnapshot(a)
    
    e.Close
End Sub
Sub Test_BufferedReader_textStream_setTextStream_Err_Value
    On Error Resume Next
    Dim a : Set a = New BufferedReader
    a.setTextStream(vbNullString)
    
    AssertEqualWithMessage "BufferedReader+setTextStream()", Err.Source, "Err.Source"
    AssertEqualWithMessage "Not a TextStream object.", Err.Description, "Err.Description"
    AssertSameWithMessage Nothing, a.textStream, "textStream"
End Sub
Sub Test_BufferedReader_textStream_setTextStream_Err_Object
    On Error Resume Next
    Dim a : Set a = New BufferedReader
    a.setTextStream(new_Dic())
    
    AssertEqualWithMessage "BufferedReader+setTextStream()", Err.Source, "Err.Source"
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
