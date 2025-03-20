' Broker.vbs: test.
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
Option Explicit

'###################################################################################################
'Broker
Sub Test_Broker
    Dim a : Set a = new Broker
    AssertEqual 9, VarType(a)
    AssertEqual "Broker", TypeName(a)
End Sub

'###################################################################################################
'Broker.subscribe()/publish()/unsubscribe()
Sub Test_Broker_subscribe_publish
    Dim ao,a,e
    Set ao = new Broker
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
