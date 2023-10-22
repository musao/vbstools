' clsCmCalendar.vbs: test.
' @import ../lib/clsCmArray.vbs
' @import ../lib/clsCmBroker.vbs
' @import ../lib/clsCmBufferedReader.vbs
' @import ../lib/clsCmBufferedWriter.vbs
' @import ../lib/clsCmCalendar.vbs
' @import ../lib/clsCmHtmlGenerator.vbs
' @import ../lib/clsCompareExcel.vbs
' @import ../lib/clsFsBase.vbs
' @import ../lib/libCom.vbs

Option Explicit

'###################################################################################################
'test_clsCmHtmlGenerator
Sub Test_clsCmBroker
    Dim a : Set a = new clsCmHtmlGenerator
    AssertEqual 9, VarType(a)
    AssertEqual "clsCmHtmlGenerator", TypeName(a)
End Sub

'###################################################################################################
'test_clsCmHtmlGenerator.xxx()
Sub Test_clsCmBroker_subscribe_publish
End Sub

' Local Variables:
' mode: Visual-Basic
' indent-tabs-mode: nil
' End:
