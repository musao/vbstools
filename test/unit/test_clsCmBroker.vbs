' clsCmBroker.vbs: test.
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

'###################################################################################################
'clsCmBroker
Sub Test_clsCmBroker
    Dim a : Set a = new clsCmBroker
    AssertEqual 9, VarType(a)
    AssertEqual "clsCmBroker", TypeName(a)
End Sub

'###################################################################################################
'clsCmBroker.subscribe()/publish()/unsubscribe()
Sub Test_clsCmBroker_subscribe_publish
    Dim ao,a,e
    Set ao = new clsCmBroker
    ao.subscribe "test1", new_Func("function(a){a=2*a}")
    ao.subscribe "test2", new_Func("function(a){a=10*a}")

    e = 2 : a = 1
    ao.publish "test1",a
    AssertEqualWithMessage e, a, "1-1 publishÅ®ubscribe"

    e = 10 : a = 1
    ao.publish "test2",a
    AssertEqualWithMessage e, a, "1-2 publishÅ®ubscribe"

    e = 1 : a = 1
    ao.publish "dummy",a
    AssertEqualWithMessage e, a, "1-3 publishÅ®Non"

    ao.unsubscribe "test1"

    e = 1 : a = 1
    ao.publish "test1",a
    AssertEqualWithMessage e, a, "2-1 publishÅ®Non"

    e = 10 : a = 1
    ao.publish "test2",a
    AssertEqualWithMessage e, a, "2-2 publishÅ®ubscribe"

    e = 1 : a = 1
    ao.publish "dummy",a
    AssertEqualWithMessage e, a, "2-3 publishÅ®Non"
End Sub

' Local Variables:
' mode: Visual-Basic
' indent-tabs-mode: nil
' End:
