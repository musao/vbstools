' clsCmBufferedWriter.vbs: test.
' @import ../../lib/com/clsAdptFile.vbs
' @import ../../lib/com/clsCmArray.vbs
' @import ../../lib/com/clsCmBroker.vbs
' @import ../../lib/com/clsCmBufferedReader.vbs
' @import ../../lib/com/clsCmBufferedWriter.vbs
' @import ../../lib/com/clsCmCalendar.vbs
' @import ../../lib/com/clsCmCharacterType.vbs
' @import ../../lib/com/clsCmCssGenerator.vbs
' @import ../../lib/com/clsCmHtmlGenerator.vbs
' @import ../../lib/com/clsCmReadOnlyObject.vbs
' @import ../../lib/com/clsCmReturnValue.vbs
' @import ../../lib/com/libCom.vbs

Option Explicit

' for fso.OpenTextFile
Const ForReading = 1, ForWriting = 2, ForAppending = 8
Dim PsPathTempFolder,PsPathForWriting,PsPathForAppending

'###################################################################################################
'SetUp()/TearDown()
Sub SetUp()
    PsPathTempFolder = new_Fso().BuildPath(new_Fso().GetParentFolderName(WScript.ScriptFullName), "test_clsCmBufferedWriter")
    If Not(new_Fso().FolderExists(PsPathTempFolder)) Then new_Fso().CreateFolder(PsPathTempFolder)
'    fs_createFolder PsPathTempFolder
    PsPathForAppending = new_Fso().BuildPath(PsPathTempFolder, new_Fso().GetTempName())
'    PsPathForAppending = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("UTat_YYMMDD_hhmmss.000000.txt"))
    With new_Ts(PsPathForAppending, ForWriting, True, -2)
        .Write("‚ ‚¢‚¤‚¦‚¨" & vbCr)
        .Close
    End With
    PsPathForWriting = new_Fso().BuildPath(PsPathTempFolder, new_Fso().GetTempName())
'    PsPathForWriting = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("UTat_YYMMDD_hhmmss.000000.txt"))
End Sub
Sub TearDown()
    new_Fso().DeleteFolder PsPathTempFolder
'    fs_deleteFolder PsPathTempFolder
End Sub




'###################################################################################################
'clsCmBufferedWriter
Sub Test_clsCmBufferedWriter
    Dim a : Set a = new clsCmBufferedWriter
    AssertEqual 9, VarType(a)
    AssertEqual "clsCmBufferedWriter", TypeName(a)
End Sub

'###################################################################################################
'clsCmBufferedWriter.writeBufferSize()
Sub Test_clsCmBufferedWriter_writeBufferSize
    Dim a,e
    Set a = new clsCmBufferedWriter

    AssertEqualWithMessage 5000, a.writeBufferSize, "1"
    
    e = -1*2^31
    a.writeBufferSize = e
    AssertEqualWithMessage e, a.writeBufferSize, "e="&e
    
    e = -1
    a.writeBufferSize = e
    AssertEqualWithMessage e, a.writeBufferSize, "e="&e

    e = 0
    a.writeBufferSize = e
    AssertEqualWithMessage e, a.writeBufferSize, "e="&e

    e = 1
    a.writeBufferSize = e
    AssertEqualWithMessage e, a.writeBufferSize, "e="&e
    
    e = 10.5
    a.writeBufferSize = e
    AssertEqualWithMessage CDbl(e), a.writeBufferSize, "e="&e
    
    e = 2^31-1
    a.writeBufferSize = e
    AssertEqualWithMessage e, a.writeBufferSize, "e="&e
End Sub
Sub Test_clsCmBufferedWriter_writeBufferSize_Err_OverLower
    On Error Resume Next
    Dim a,d,e
    Set a = new clsCmBufferedWriter
    
    e = a.writeBufferSize
    d = -1*2^(2^1024)
    a.writeBufferSize = d
    
    AssertEqualWithMessage "clsCmBufferedWriter+writeBufferSize() Let", Err.Source, "Err.Source"
    AssertEqualWithMessage "Invalid number.", Err.Description, "Err.Description"
    AssertEqualWithMessage e, a.writeBufferSize, "e="&e
End Sub
Sub Test_clsCmBufferedWriter_writeBufferSize_Err_OverUpper
    On Error Resume Next
    Dim a,d,e
    Set a = new clsCmBufferedWriter
    
    e = 100
    a.writeBufferSize = e

    d = 2^(2^1024)
    a.writeBufferSize = d
    
    AssertEqualWithMessage "clsCmBufferedWriter+writeBufferSize() Let", Err.Source, "Err.Source"
    AssertEqualWithMessage "Invalid number.", Err.Description, "Err.Description"
    AssertEqualWithMessage e, a.writeBufferSize, "e="&e
End Sub
Sub Test_clsCmBufferedWriter_writeBufferSize_Err_NotNumeric
    On Error Resume Next
    Dim a,d,e
    Set a = new clsCmBufferedWriter
    
    e = a.writeBufferSize

    d = "abc"
    a.writeBufferSize = d
    
    AssertEqualWithMessage "clsCmBufferedWriter+writeBufferSize() Let", Err.Source, "Err.Source"
    AssertEqualWithMessage "Invalid number.", Err.Description, "Err.Description"
    AssertEqualWithMessage e, a.writeBufferSize, "e="&e
End Sub

'###################################################################################################
'clsCmBufferedWriter.writeIntervalTime()
Sub Test_clsCmBufferedWriter_writeIntervalTime
    Dim a,e
    Set a = new clsCmBufferedWriter

    AssertEqualWithMessage 0, a.writeIntervalTime, "1"
    
    e = -1*2^31
    a.writeIntervalTime = e
    AssertEqualWithMessage e, a.writeIntervalTime, "e="&e
    
    e = -1
    a.writeIntervalTime = e
    AssertEqualWithMessage e, a.writeIntervalTime, "e="&e

    e = 0
    a.writeIntervalTime = e
    AssertEqualWithMessage e, a.writeIntervalTime, "e="&e

    e = 1
    a.writeIntervalTime = e
    AssertEqualWithMessage e, a.writeIntervalTime, "e="&e
    
    e = 10.5
    a.writeIntervalTime = e
    AssertEqualWithMessage CDbl(e), a.writeIntervalTime, "e="&e
    
    e = 2^31-1
    a.writeIntervalTime = e
    AssertEqualWithMessage e, a.writeIntervalTime, "e="&e
End Sub
Sub Test_clsCmBufferedWriter_writeIntervalTime_Err_OverLower
    On Error Resume Next
    Dim a,d,e
    Set a = new clsCmBufferedWriter
    
    e = a.writeIntervalTime
    d = -1*2^(2^10)
    a.writeIntervalTime = d
    
    AssertEqualWithMessage "clsCmBufferedWriter+writeIntervalTime() Let", Err.Source, "Err.Source"
    AssertEqualWithMessage "Invalid number.", Err.Description, "Err.Description"
    AssertEqualWithMessage e, a.writeIntervalTime, "e="&e
End Sub
Sub Test_clsCmBufferedWriter_writeIntervalTime_Err_OverUpper
    On Error Resume Next
    Dim a,d,e
    Set a = new clsCmBufferedWriter
    
    e = 100
    a.writeIntervalTime = e

    d = 2^(2^10)
    a.writeIntervalTime = d
    
    AssertEqualWithMessage "clsCmBufferedWriter+writeIntervalTime() Let", Err.Source, "Err.Source"
    AssertEqualWithMessage "Invalid number.", Err.Description, "Err.Description"
    AssertEqualWithMessage e, a.writeIntervalTime, "e="&e
End Sub
Sub Test_clsCmBufferedWriter_writeIntervalTime_Err_NonNumeric
    On Error Resume Next
    Dim a,d,e
    Set a = new clsCmBufferedWriter
    
    e = a.writeIntervalTime

    d = "abc"
    a.writeIntervalTime = d
    
    AssertEqualWithMessage "clsCmBufferedWriter+writeIntervalTime() Let", Err.Source, "Err.Source"
    AssertEqualWithMessage "Invalid number.", Err.Description, "Err.Description"
    AssertEqualWithMessage e, a.writeIntervalTime, "e="&e
End Sub

'###################################################################################################
'clsCmBufferedWriter.textStream()/setTextStream()
Sub Test_clsCmBufferedWriter_textStream_setTextStream_Write
    Dim pt : pt = PsPathForWriting
    Dim e : Set e = new_Fso().OpenTextFile(pt, ForWriting, True, -2)
    Dim a : Set a = New clsCmBufferedWriter
    a.setTextStream(e)
    
    AssertSame e, a.textStream
    assertAll takeSnapshot(e), takeSnapshot(a)

    e.Close
End Sub
Sub Test_clsCmBufferedWriter_textStream_setTextStream_Append
    Dim pt : pt = PsPathForAppending
    Dim e : Set e = new_Fso().OpenTextFile(pt, ForAppending, False, -2)
    Dim a : Set a = New clsCmBufferedWriter
    a.setTextStream(e)
    
    AssertSame e, a.textStream
    assertAll takeSnapshot(e), takeSnapshot(a)

    e.Close
End Sub
Sub Test_clsCmBufferedWriter_textStream_setTextStream_Err_Value
    On Error Resume Next
    Dim a : Set a = New clsCmBufferedWriter
    a.setTextStream(vbNullString)
    
    AssertEqualWithMessage "clsCmBufferedWriter+setTextStream()", Err.Source, "Err.Source"
    AssertEqualWithMessage "Not a TextStream object.", Err.Description, "Err.Description"
    AssertSameWithMessage Nothing, a.textStream, "textStream"
End Sub
Sub Test_clsCmBufferedWriter_textStream_setTextStream_Err_Object
    On Error Resume Next
    Dim a : Set a = New clsCmBufferedWriter
    a.setTextStream(new_Dic())
    
    AssertEqualWithMessage "clsCmBufferedWriter+setTextStream()", Err.Source, "Err.Source"
    AssertEqualWithMessage "Not a TextStream object.", Err.Description, "Err.Description"
    AssertSameWithMessage Nothing, a.textStream, "textStream"
End Sub

'###################################################################################################
'clsCmBufferedWriter.currentBufferSize()
Sub Test_clsCmBufferedWriter_currentBufferSize
    Dim pt,ts,ao,a,e,d,sz
    pt = PsPathForWriting
    Set ts = new_Fso().OpenTextFile(pt, ForWriting, True, -2)
    Set ao = New clsCmBufferedWriter

    With ao
        .setTextStream(ts)
        .writeBufferSize = 10
        .writeIntervalTime = 0
        e = 0
        
        d = "ab‚ ‚¢c‚¤"
        e = e + func_CM_StrLen(d)
        .write(d)
        a = .currentBufferSize
        sz = new_Fso().GetFile(pt).Size
        AssertEqualWithMessage e, a, "1-1"
        AssertEqualWithMessage 0, sz, "1-2"
        
        d = "d‚¦e‚¨"
        e = e + func_CM_StrLen(d)
        .write(d)
        a = .currentBufferSize
        sz = new_Fso().GetFile(pt).Size
        AssertEqualWithMessage 0, a, "2-1"
        AssertEqualWithMessage e, sz, "2-2"
        
        d = "‚©f‚«g‚­h‚¯i‚±j"
        e = e + func_CM_StrLen(d)
        .write(d)
        a = .currentBufferSize
        sz = new_Fso().GetFile(pt).Size
        AssertEqualWithMessage 0, a, "3-1"
        AssertEqualWithMessage e, sz, "3-2"
    End With

    ts.Close
End Sub
Sub Test_clsCmBufferedWriter_currentBufferSize_Empty
    Dim ao,a,e
    Set ao = New clsCmBufferedWriter
    
    e = Empty
    a = ao.currentBufferSize
    AssertEqualWithMessage e, a, "1-1"
End Sub

'###################################################################################################
'clsCmBufferedWriter.lastWriteTime()
Sub Test_clsCmBufferedWriter_lastWriteTime
    Dim pt,ts,ao,a,e,d,sz,nsz
    pt = PsPathForAppending
    Set ts = new_Fso().OpenTextFile(pt, ForAppending, False, -2)
    Set ao = New clsCmBufferedWriter

    With ao
        .setTextStream(ts)
        .writeBufferSize = 5000
        .writeIntervalTime = 0.03
        sz = new_Fso().GetFile(pt).Size
        
        d = "ab‚ ‚¢c‚¤"
        e = Empty
        .write(d)
        a = .lastWriteTime
        nsz = new_Fso().GetFile(pt).Size
        AssertEqualWithMessage e, a, "1-1"
        AssertEqualWithMessage sz, nsz, "1-2"
        
        wscript.sleep 50

        d = "d‚¦e‚¨"
        e = new_Now()
        .write(d)
        a = .lastWriteTime
        nsz = new_Fso().GetFile(pt).Size
        AssertWithMessage e=<a, "2-1 ["&e&"] ["&a&"]"
        AssertWithMessage sz<nsz, "2-2"
        AssertWithMessage Not IsEmpty(a), "2-3"
        
        wscript.sleep 50
        
        d = "‚©f‚«g‚­h‚¯i‚±j"
        e = new_Now()
        .write(d)
        a = .lastWriteTime
        nsz = new_Fso().GetFile(pt).Size
        AssertWithMessage e=<a, "3-1 ["&e&"] ["&a&"]"
        AssertWithMessage sz<nsz, "3-2"
        AssertWithMessage Not IsEmpty(a), "3-3"
    End With

    ts.Close
End Sub
Sub Test_clsCmBufferedWriter_lastWriteTime_Empty
    Dim ao,a,e
    Set ao = New clsCmBufferedWriter
    
    e = Empty
    a = ao.lastWriteTime
    AssertEqualWithMessage e, a, "1-1"
End Sub

'###################################################################################################
'clsCmBufferedWriter.newLine()
Sub Test_clsCmBufferedWriter_newLine
    Dim pt,ts,ao,a,el,ec,es,sz,nsz
    pt = PsPathForWriting
    Set ts = new_Fso().OpenTextFile(pt, ForWriting, True, -2)
    Set ao = New clsCmBufferedWriter

    With ao
        .setTextStream(ts)
        .writeBufferSize = 5
        .writeIntervalTime = 0

        el = 2 : ec = 1 : es = 2
        .newLine()
        sz = new_Fso().GetFile(pt).Size
        AssertEqualWithMessage el, .line, "1-1"
        AssertEqualWithMessage ec, .column, "1-2"
        AssertEqualWithMessage es, .currentBufferSize, "1-3"
        AssertEqualWithMessage 0, sz, "1-4"

        el = 4 : ec = 1 : es = 6
        .newLine()
        .newLine()
        sz = new_Fso().GetFile(pt).Size
        AssertEqualWithMessage el, .line, "2-1"
        AssertEqualWithMessage ec, .column, "2-2"
        AssertEqualWithMessage 0, .currentBufferSize, "2-3"
        AssertEqualWithMessage es, sz, "2-4"

        el = 5 : ec = 1 : es = 8
        .newLine()
        nsz = new_Fso().GetFile(pt).Size
        AssertEqualWithMessage el, .line, "3-1"
        AssertEqualWithMessage ec, .column, "3-2"
        AssertEqualWithMessage es - sz, .currentBufferSize, "3-3"
        AssertEqualWithMessage nsz, sz, "3-4"

        .close()
    End With

    ts.Close
End Sub

'###################################################################################################
'clsCmBufferedWriter.flush()
Sub Test_clsCmBufferedWriter_flush
   Dim pt,ts,ao,at,ab,et,eb,d,sz,nsz
    pt = PsPathForAppending
    Set ts = new_Fso().OpenTextFile(pt, ForAppending, False, -2)
    Set ao = New clsCmBufferedWriter

    With ao
        .setTextStream(ts)
        .writeBufferSize = 5000
        .writeIntervalTime = 0
        sz = new_Fso().GetFile(pt).Size
        eb = 0

        d = "ab‚ ‚¢c‚¤"
        et = Empty
        eb = eb + func_CM_StrLen(d)
        .write(d)
        at = .lastWriteTime
        ab = .currentBufferSize
        nsz = new_Fso().GetFile(pt).Size
        AssertEqualWithMessage et, at, "1-1"
        AssertEqualWithMessage eb, ab, "1-2"
        AssertEqualWithMessage sz, nsz, "1-3"
        
        et = new_Now()
        .flush()
        at = .lastWriteTime
        ab = .currentBufferSize
        nsz = new_Fso().GetFile(pt).Size
        AssertWithMessage et=<at, "2-1 ["&et&"] ["&at&"]"
        AssertEqualWithMessage 0, ab, "2-2"
        AssertWithMessage sz<nsz, "2-3"
        AssertWithMessage Not IsEmpty(at), "2-4"
    End With

    ts.Close
End Sub

'###################################################################################################
'clsCmBufferedWriter.close()
Sub Test_clsCmBufferedWriter_close
    Dim pt,ts,d,ao,et,eb,at,ab,sz,nsz
    pt = PsPathForWriting
    Set ts = new_Fso().OpenTextFile(pt, ForWriting, True, -2)
    Set ao = New clsCmBufferedWriter

    With ao
        .setTextStream(ts)
        .writeBufferSize = 5000
        .writeIntervalTime = 0
        sz = 0
        eb = 0

        d = "ab‚ ‚¢c‚¤"
        et = Empty
        eb = eb + func_CM_StrLen(d)
        .write(d)
        at = .lastWriteTime
        ab = .currentBufferSize
        nsz = new_Fso().GetFile(pt).Size
        AssertEqualWithMessage et, at, "1-1"
        AssertEqualWithMessage eb, ab, "1-2"
        AssertEqualWithMessage sz, nsz, "1-3"
        sz = nsz
        
        .close()
        nsz = new_Fso().GetFile(pt).Size
        AssertEqualWithMessage eb, nsz, "2-1"
        AssertWithMessage sz<nsz, "2-2"
    End With

    ts.Close
End Sub





'###################################################################################################
'common
Function takeSnapshot(o)
    Dim ret : Set ret = new_Dic()
    with o
        ret.add "Line", .Line
        ret.add "Column", .Column
    end with
    Set takeSnapshot = ret
End Function
Sub assertAll(a,b)
    Dim sKey
    For Each sKey In Array("Line","Column")
        AssertEqualWithMessage a.Item(sKey), b.Item(sKey), sKey
    Next
End Sub


' Local Variables:
' mode: Visual-Basic
' indent-tabs-mode: nil
' End:
