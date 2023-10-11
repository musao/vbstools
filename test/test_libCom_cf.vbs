' libCom.vbs: cf_* procedure test.
' @import ../lib/clsCmArray.vbs
' @import ../lib/clsCmBufferedReader.vbs
' @import ../lib/clsCmBufferedWriter.vbs
' @import ../lib/clsCmCalendar.vbs
' @import ../lib/clsCmPubSub.vbs
' @import ../lib/clsCompareExcel.vbs
' @import ../lib/clsFsBase.vbs
' @import ../lib/libCom.vbs

Option Explicit

'cf_bind()
Sub Test_cf_bind_Value
  Dim v
  cf_bind v, "Hello world."
  AssertEqual "Hello world.", v
End Sub

Sub Test_cf_bind_Object
  Dim v
  Dim obj: Set obj = CreateObject("Scripting.Dictionary")
  cf_bind v, obj
  AssertSame obj, v
End Sub


'cf_bindAt()
Sub Test_cf_bindAt_Value
  Dim obj : Set obj = CreateObject("Scripting.Dictionary")
  cf_bindAt obj, "Value", "Hello world."
  AssertEqual "Hello world.", obj.Item("Value")
End Sub

Sub Test_cf_bindAt_Object
  Dim obj : Set obj = CreateObject("Scripting.Dictionary")
  cf_bindAt obj, "Object", Nothing
  AssertSame Nothing, obj.Item("Object")
End Sub

' Local Variables:
' mode: Visual-Basic
' indent-tabs-mode: nil
' End:
