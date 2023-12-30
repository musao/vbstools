' clsCmBufferedWriter.vbs: test.
' @import ../../lib/clsCmArray.vbs
' @import ../../lib/clsCmBroker.vbs
' @import ../../lib/clsCmBufferedReader.vbs
' @import ../../lib/clsCmBufferedWriter.vbs
' @import ../../lib/clsCmCalendar.vbs
' @import ../../lib/clsCmCharacterType.vbs
' @import ../../lib/clsCmCssGenerator.vbs
' @import ../../lib/clsCmHtmlGenerator.vbs
' @import ../../lib/clsCompareExcel.vbs
' @import ../../lib/libCom.vbs

Option Explicit

' for fso.OpenTextFile
Const ForReading = 1, ForWriting = 2, ForAppending = 8
Dim PsPathTempFolder,PsPathForWriting,PsPathForAppending

'###################################################################################################
'SetUp()/TearDown()
Sub SetUp()
    PsPathTempFolder = new_Fso().BuildPath(new_Fso().GetParentFolderName(WScript.ScriptFullName), "test_clsCmBufferedWriter")
    If Not(new_Fso().FolderExists(PsPathTempFolder)) Then fs_createFolder(PsPathTempFolder)
    PsPathForAppending = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("UTat_YYMMDD_hhmmss.000000.txt"))
    With new_Ts(PsPathForAppending, ForWriting, True, -2)
        .Write("‚ ‚¢‚¤‚¦‚¨" & vbCr)
        .Close
    End With
    PsPathForWriting = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("UTat_YYMMDD_hhmmss.000000.txt"))
End Sub
Sub TearDown()
    fs_deleteFolder PsPathTempFolder
End Sub


'###################################################################################################
'clsCmBufferedReader.write/writeBlankLines/writeLine()
Sub Test_clsCmBufferedReader_write_writeBlankLines_writeLine_Normal_Write
    writer_testCommon getref("operations_write_writeBlankLines_writeLine_Normal"), "Write"
End Sub
Sub Test_clsCmBufferedReader_write_writeBlankLines_writeLine_Normal_Append
    writer_testCommon getref("operations_write_writeBlankLines_writeLine_Normal"), "Appending"
End Sub
Sub operations_write_writeBlankLines_writeLine_Normal(o,r)
    o.write("‚ ‚¢‚¤‚¦‚¨")
    cf_push r, takeSnapshot(o)

    o.writeBlankLines(2)
    cf_push r, takeSnapshot(o)

    o.writeLine("AbcDe")
    cf_push r, takeSnapshot(o)

    o.writeBlankLines(1)
    cf_push r, takeSnapshot(o)

    o.write("‚©F‚«g‚­hI‚¯‚±k")
    cf_push r, takeSnapshot(o)

    o.Close
End Sub

Sub Test_clsCmBufferedReader_write_writeBlankLines_writeLine_NewLineOnly_Write
    writer_testCommon getref("operations_write_writeBlankLines_writeLine_NewLineOnly"), "Write"
End Sub
Sub Test_clsCmBufferedReader_write_writeBlankLines_writeLine_NewLineOnly_Append
    writer_testCommon getref("operations_write_writeBlankLines_writeLine_NewLineOnly"), "Appending"
End Sub
Sub operations_write_writeBlankLines_writeLine_NewLineOnly(o,r)
    o.write(vbCr)
    cf_push r, takeSnapshot(o)

    o.writeBlankLines(2)
    cf_push r, takeSnapshot(o)

    o.writeLine(vbLf)
    cf_push r, takeSnapshot(o)

    o.writeBlankLines(1)
    cf_push r, takeSnapshot(o)

    o.write(vbCrLf)
    cf_push r, takeSnapshot(o)

    o.Close
End Sub





'###################################################################################################
'common
Sub writer_testCommon(f,p)
    Dim j,rs
    rs = Array(5000,10,1)
    For j=0 To Ubound(rs)
        Dim ev,eo,er,av,ao,ar,i,flg,pt
        flg=true
        For i=0 To 1
            If strcomp(p, "Write", vbBinaryCompare)=0 Then
                pt = PsPathForWriting
                Set eo = new_Fso().OpenTextFile(pt, ForWriting, True, -2)
            Else
                pt = PsPathForAppending
                With new_Ts(pt, ForWriting, True, -2)
                    .Write("‚ ‚¢‚¤‚¦‚¨" & vbCr)
                    .Close
                End With
                Set eo = new_Fso().OpenTextFile(pt, ForAppending, False, -2)
            End If
            If flg Then
                f eo,er
                ev = new_Fso().GetFile(pt).Size
            Else
                Set ao = New clsCmBufferedWriter
                ao.writeBufferSize = rs(j)
                ao.setTextStream(eo)
                f ao,ar
                av = new_Fso().GetFile(pt).Size
            End If
            flg=Not flg
        Next
        AssertEqualWithMessage ev, av, "writeBufferSize="&rs(j)&" "&i&"operation"
        For i=0 To Ubound(er)
            assertAll er(i), ar(i)
        Next
    Next
End Sub
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
