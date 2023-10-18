' clsCmCalendar.vbs: test.
' @import ../lib/clsCmArray.vbs
' @import ../lib/clsCmBufferedReader.vbs
' @import ../lib/clsCmBufferedWriter.vbs
' @import ../lib/clsCmCalendar.vbs
' @import ../lib/clsCmBroker.vbs
' @import ../lib/clsCompareExcel.vbs
' @import ../lib/clsFsBase.vbs
' @import ../lib/libCom.vbs

Option Explicit

'###################################################################################################
'clsCmBroker
Sub Test_clsCmBroker
    Dim a : Set a = new clsCmBroker
    AssertEqual 9, VarType(a)
    AssertEqual "clsCmBroker", TypeName(a)
End Sub

'###################################################################################################
'clsCmBroker.getNow()/toString()
Sub Test_clsCmBroker_xxxx
End Sub

' Local Variables:
' mode: Visual-Basic
' indent-tabs-mode: nil
' End:
