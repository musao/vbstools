' clsCmBufferedReader.vbs: test.
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
Dim PsPathTempFolder,PsPathData1,PsPathData2

'###################################################################################################
'SetUp()/TearDown()
Sub SetUp()
    PsPathTempFolder = new_Fso().BuildPath(new_Fso().GetParentFolderName(WScript.ScriptFullName), "test_clsCmBufferedReader_regression")
    If Not(new_Fso().FolderExists(PsPathTempFolder)) Then new_Fso().CreateFolder(PsPathTempFolder)
'    fs_createFolder PsPathTempFolder
    PsPathData1 = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("UTat_YYMMDD_hhmmss.000000.txt"))
    With new_Ts(PsPathData1, ForWriting, True, -2)
        .Write("‚ ‚¢‚¤‚¦‚¨" & vbCrLf & vbCr & "abcde" & vbLf & vbLf & "12" & vbCr & "345")
        .Close
    End With
    PsPathData2 = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("UTat_YYMMDD_hhmmss.000000.txt"))
    With new_Ts(PsPathData2, ForWriting, True, -2)
        .Write("‚©‚«‚­‚¯‚±" & vbCr)
        .Close
    End With
End Sub
Sub TearDown()
    new_Fso().DeleteFolder PsPathTempFolder
'    fs_deleteFolder PsPathTempFolder
End Sub




'###################################################################################################
'clsCmBufferedReader.readAll()
Sub Test_clsCmBufferedReader_readAll
    reader_testCommon getref("operations_readAll"), PsPathData1
End Sub
Sub operations_readAll(v,o,r)
    cf_push v, o.ReadAll()
    cf_push r, takeSnapshot(o)
    o.Close
End Sub

'###################################################################################################
'clsCmBufferedReader.read/readLine/skip/skipLine()
Sub Test_clsCmBufferedReader_read_readLine_skip_skipLine_1
    reader_testCommon getref("operations_read_readLine_skip_skipLine_1"), PsPathData1
End Sub
Sub operations_read_readLine_skip_skipLine_1(v,o,r)
    cf_push v, o.readLine()
    cf_push r, takeSnapshot(o)

    cf_push v, o.readLine()
    cf_push r, takeSnapshot(o)

    cf_push v, o.readLine()
    cf_push r, takeSnapshot(o)

    cf_push v, o.read(1)
    cf_push r, takeSnapshot(o)

    cf_push v, o.readLine()
    cf_push r, takeSnapshot(o)

    o.Close
End Sub
Sub Test_clsCmBufferedReader_read_readLine_skip_skipLine_2
    reader_testCommon getref("operations_read_readLine_skip_skipLine_2"), PsPathData1
End Sub
Sub operations_read_readLine_skip_skipLine_2(v,o,r)
    cf_push v, o.skipLine()
    cf_push r, takeSnapshot(o)

    cf_push v, o.skipLine()
    cf_push r, takeSnapshot(o)

    cf_push v, o.skipLine()
    cf_push r, takeSnapshot(o)

    cf_push v, o.Skip(2)
    cf_push r, takeSnapshot(o)

    cf_push v, o.skipLine()
    cf_push r, takeSnapshot(o)

    o.Close
End Sub
Sub Test_clsCmBufferedReader_read_readLine_skip_skipLine_3
    reader_testCommon getref("operations_read_readLine_skip_skipLine_3"), PsPathData1
End Sub
Sub operations_read_readLine_skip_skipLine_3(v,o,r)
    cf_push v, o.skipLine()
    cf_push r, takeSnapshot(o)

    cf_push v, o.readLine()
    cf_push r, takeSnapshot(o)

    cf_push v, o.skipLine()
    cf_push r, takeSnapshot(o)

    cf_push v, o.read(5)
    cf_push r, takeSnapshot(o)

    cf_push v, o.skipLine()
    cf_push r, takeSnapshot(o)

    o.Close
End Sub
Sub Test_clsCmBufferedReader_read_readLine_skip_skipLine_4
    reader_testCommon getref("operations_read_readLine_skip_skipLine_4"), PsPathData1
End Sub
Sub operations_read_readLine_skip_skipLine_4(v,o,r)
    cf_push v, o.skipLine()
    cf_push r, takeSnapshot(o)

    cf_push v, o.read(6)
    cf_push r, takeSnapshot(o)

    cf_push v, o.readLine()
    cf_push r, takeSnapshot(o)

    o.Close
End Sub
Sub Test_clsCmBufferedReader_read_readLine_skip_skipLine_5
    reader_testCommon getref("operations_read_readLine_skip_skipLine_5"), PsPathData2
End Sub
Sub operations_read_readLine_skip_skipLine_5(v,o,r)
    cf_push v, o.Skip(5)
    cf_push r, takeSnapshot(o)

    cf_push v, o.readLine()
    cf_push r, takeSnapshot(o)

    o.Close
End Sub





'###################################################################################################
'common
Sub reader_testCommon(f,p)
    Dim j,rs
    rs = Array(5000,10,1)
    For j=0 To Ubound(rs)
        Dim ev,eo,er,av,ao,ar,i,flg
        flg=true
        For i=0 To 1
            Set eo = new_Fso().OpenTextFile(p, ForReading, False, -2)
            If flg Then
                f ev,eo,er
            Else
                Set ao = New clsCmBufferedReader
                ao.readSize = rs(j)
                ao.setTextStream(eo)
                f av,ao,ar
            End If
            flg=Not flg
        Next
        For i=0 To Ubound(ev)
            AssertEqualWithMessage ev(i), av(i), "readsize="&rs(j)&" "&i&"operation"
            assertAll er(i), ar(i)
        Next
    Next
End Sub
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
