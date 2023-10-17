' clsCmBufferedWriter.vbs: test.
' @import ../lib/clsCmArray.vbs
' @import ../lib/clsCmBufferedReader.vbs
' @import ../lib/clsCmBufferedWriter.vbs
' @import ../lib/clsCmCalendar.vbs
' @import ../lib/clsCmBroker.vbs
' @import ../lib/clsCompareExcel.vbs
' @import ../lib/clsFsBase.vbs
' @import ../lib/libCom.vbs

Option Explicit

' for fso.OpenTextFile
Const ForReading = 1, ForWriting = 2, ForAppending = 8
Dim PsPathTempFolder,PsPathForWriting,PsPathForAppending

'###################################################################################################
'SetUp()/TearDown()
Sub SetUp()
    PsPathTempFolder = func_CM_FsBuildPath(new_Fso().GetParentFolderName(WScript.ScriptFullName), "test_clsCmBufferedWriter")
    PsPathForAppending = func_CM_FsGetFilePathWithCreateParentFolder(PsPathTempFolder, new_Now().displayAs("UTat_YYMMDD_hhmmss.000000.txt"))
    With func_CM_FsOpenTextFile(PsPathForAppending, ForWriting, True, -2)
        .Write("あいうえお" & vbCr)
        .Close
    End With
    PsPathForWriting = func_CM_FsBuildPath(PsPathTempFolder, new_Now().displayAs("UTat_YYMMDD_hhmmss.000000.txt"))
End Sub
Sub TearDown()
    func_CM_FsDeleteFolder PsPathTempFolder
End Sub


'###################################################################################################
'clsCmBufferedReader.write/writeBlankLines/writeLine()
Sub Test_clsCmBufferedReader_write_writeBlankLines_writeLine_1
    writer_testCommon getref("operations_write_writeBlankLines_writeLine_1"), "Write"
End Sub
Sub operations_write_writeBlankLines_writeLine_1(o,r)
    o.write("あいうえお")
    cf_push r, takeSnapshot(o)

    o.writeBlankLines(2)
    cf_push r, takeSnapshot(o)

    o.writeLine("AbcDe")
    cf_push r, takeSnapshot(o)

    o.writeBlankLines(1)
    cf_push r, takeSnapshot(o)

    o.write("かFきgくhIけこk")
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
                Set eo = new_Fso().OpenTextFile(PsPathForWriting, ForWriting, True, -2)
            Else
                pt = PsPathForAppending
                Set eo = new_Fso().OpenTextFile(PsPathForAppending, ForAppending, False, -2)
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
'        AssertEqualWithMessage ev, av, "writeBufferSize="&rs(j)&" "&i&"operation"
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
