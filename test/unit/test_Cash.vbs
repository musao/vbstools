' Cash.vbs: test.
' @import ../../lib/com/FileSystemProxy.vbs
' @import ../../lib/com/ArrayList.vbs
' @import ../../lib/com/Broker.vbs
' @import ../../lib/com/BufferedReader.vbs
' @import ../../lib/com/BufferedWriter.vbs
' @import ../../lib/com/Cash.vbs
' @import ../../lib/com/Calendar.vbs
' @import ../../lib/com/CharacterType.vbs
' @import ../../lib/com/CssGenerator.vbs
' @import ../../lib/com/HtmlGenerator.vbs
' @import ../../lib/com/ReadOnlyObject.vbs
' @import ../../lib/com/ReturnValue.vbs
' @import ../../lib/com/libCom.vbs

Option Explicit
Const LOADER_PREFIX = "Loader_"

'###################################################################################################
'Cash
Sub Test_Cash
    Dim ao : Set ao = new Cash
    AssertEqualWithMessage 9, VarType(ao), "VarType(new Cash)"
    AssertEqualWithMessage "Cash", TypeName(ao), "TypeName(new Cash)"
End Sub

'###################################################################################################
'Cash.item(default),get
Sub Test_Cash_item_get_normal
    Dim ao,k,v,t,e
    Set ao = new Cash : t = -1 

    k = "hoge" : v = 123 : e = v
    ao.put k, v, t
    AssertEqualWithMessage e, ao.item(k), "ao.item('hoge')"
    AssertEqualWithMessage e, ao(k), "default ao('hoge')"
    AssertEqualWithMessage e, ao.get(k), "ao.get('hoge')"

    k = "foo" : Set v = new_Dic() : Set e = v
    ao.put k, v, t
    AssertSameWithMessage e, ao.item(k), "ao.item('foo')"
    AssertSameWithMessage e, ao(k), "default ao('foo')"
    AssertSameWithMessage e, ao.get(k), "ao.get('foo')"
End Sub
Sub Test_Cash_item_get_no_key
    Dim ao,k,e
    Set ao = new Cash

    k = "nonexistent_key" : e = True
    AssertEqualWithMessage e, IsNull(ao.item(k)), "ao.item('nonexistent_key')"
    AssertEqualWithMessage e, IsNull(ao(k)), "default ao('nonexistent_key')"
    AssertEqualWithMessage e, IsNull(ao.get(k)), "ao.get('nonexistent_key')"
End Sub
Sub Test_Cash_item_get_timeout_key
    Dim ao,k,v,t,e,waitTime
    Set ao = new Cash
    t = 0 : waitTime = 1

    k = "timeout_key" : v = 123
    ao.put k, v, t
    WScript.Sleep waitTime
    e = True
    AssertEqualWithMessage e, IsNull(ao.item(k)), "ao.item('timeout_key')"
    AssertEqualWithMessage e, IsNull(ao(k)), "default ao('timeout_key')"
    AssertEqualWithMessage e, IsNull(ao.get(k)), "ao.get('timeout_key')"
End Sub

'###################################################################################################
'Cash.getOrCompute
Sub Test_Cash_getOrCompute_normal
    Dim ao,a,k,l,t,e
    Set ao = new Cash
    t = -1 : Set l = GetRef("getOrComputeLoader")

    k = "hoge" : e = LOADER_PREFIX & k
    a = ao.getOrCompute(k, l, t)
    AssertEqualWithMessage e, a, "ao.getOrCompute('hoge', 'getOrComputeLoader', -1) first call"
    a = ao.getOrCompute(k, l, t)
    AssertEqualWithMessage e, a, "ao.getOrCompute('hoge', 'getOrComputeLoader', -1) second call"
End Sub
Sub Test_Cash_getOrCompute_timeout_key
    Dim ao,a,k,l,t,e,waitTime
    Set ao = new Cash
    t = 0 : waitTime = 1 : Set l = GetRef("getOrComputeLoader")

    k = "timeout_key" : e = LOADER_PREFIX & k
    a = ao.getOrCompute(k, l, t)
    AssertEqualWithMessage e, a, "ao.getOrCompute('hoge', 'getOrComputeLoader', 5) first call"
    WScript.Sleep waitTime
    a = ao.getOrCompute(k, l, t)
    AssertEqualWithMessage e, a, "ao.getOrCompute('hoge', 'getOrComputeLoader', 5) second call"
End Sub

'###################################################################################################
'Cash.clear
Sub Test_Cash_clear_hasitems
    Dim ao : Set ao = new Cash
    
    ao.put "hoge", 123, -1
    ao.put "foo", new_Dic(), -1

    AssertEqualWithMessage 2, ao.size(), "ao.size() before clear"

    On Error Resume Next
    ao.clear()
    AssertEqualWithMessage 0, Err.Number, "ao.clear()"
    On Error GoTo 0

    AssertEqualWithMessage 0, ao.size(), "ao.size() after clear"
End Sub
Sub Test_Cash_clear_hasnoitems
    Dim ao : Set ao = new Cash

    AssertEqualWithMessage 0, ao.size(), "ao.size() before clear"

    On Error Resume Next
    ao.clear()
    AssertEqualWithMessage 0, Err.Number, "ao.clear()"
    On Error GoTo 0

    AssertEqualWithMessage 0, ao.size(), "ao.size() after clear"
End Sub

'###################################################################################################
'Cash.delete
Sub Test_Cash_delete_deleteexistent_key
    Dim ao,k,v,t,e
    Set ao = new Cash : t = -1
    
    k = "hoge" : v = 123
    ao.put k, v, t

    e = v
    AssertEqualWithMessage e, ao(k), "ao('hoge') before delete"

    On Error Resume Next
    ao.delete k
    AssertEqualWithMessage 0, Err.Number, "ao.delete('hoge')"
    On Error GoTo 0

    e = True
    AssertEqualWithMessage e, IsNull(ao(k)), "ao('hoge') after delete"
End Sub
Sub Test_Cash_delete_deletenonexistent_key
    Dim ao,k,e
    Set ao = new Cash : k = "hoge"

    e = True
    AssertEqualWithMessage e, IsNull(ao(k)), "ao('hoge') before delete"

    On Error Resume Next
    ao.delete k
    AssertEqualWithMessage 0, Err.Number, "ao.delete('hoge')"
    On Error GoTo 0

    e = True
    AssertEqualWithMessage e, IsNull(ao(k)), "ao('hoge') after delete"
End Sub

'###################################################################################################
'Cash.has
Sub Test_Cash_has_existent_key
    Dim ao,a,k,v,t,e
    Set ao = new Cash : t = -1 

    k = "hoge" : v = 123
    ao.put k, v, t

    e = True
    a = ao.has(k)
    AssertEqualWithMessage e, a, "ao.has('hoge')"
End Sub
Sub Test_Cash_has_nonexistent_key
    Dim ao,a,k,e
    Set ao = new Cash

    k = "nonexistent_key"
    e = False
    a = ao.has(k)
    AssertEqualWithMessage e, a, "ao.has('hoge')"
End Sub
Sub Test_Cash_has_timeout_key
    Dim ao,a,k,v,t,e,waitTime
    Set ao = new Cash
    t = 50 : waitTime = 50

    k = "timeout_key" : v = 123
    ao.put k, v, t
    
    e = True
    a = ao.has(k)
    AssertEqualWithMessage e, a, "ao.has('hoge') first call"
    WScript.Sleep waitTime
    e = False
    a = ao.has(k)
    AssertEqualWithMessage e, a, "ao.has('hoge') second call"
End Sub

'###################################################################################################
'Cash.put
Sub Test_Cash_put_normal_value
    Dim ao,a,k,v,t,e
    Set ao = new Cash : t = -1
    
    k = "hoge" : v = 123 : e = v
    ao.put k, v, t
    a = ao(k)
    AssertEqualWithMessage a, e, "ao('hoge')"
End Sub
Sub Test_Cash_put_normal_object
    Dim ao,a,k,v,t,e
    Set ao = new Cash : t = -1
    
    k = "hoge" : Set v = new_Dic() : Set e = v
    ao.put k, v, t
    Set a = ao(k)
    AssertSameWithMessage a, e, "ao('hoge')"
End Sub
Sub Test_Cash_put_normal_orverride_value2object
    Dim ao,a,k,v,t,e
    Set ao = new Cash : t = 1000 : k = "hoge"

    v = 123 : e = v
    ao.put k, v, t
    a = ao(k)
    AssertEqualWithMessage a, e, "ao('hoge') first put"
    
    Set v = new_Dic() : Set e = v
    ao.put k, v, t
    Set a = ao(k)
    AssertSameWithMessage a, e, "ao('hoge') second put"
End Sub
Sub Test_Cash_put_normal_orverride_object2value
    Dim ao,a,k,v,t,e
    Set ao = new Cash : t = 1000 : k = "hoge"
    
    Set v = new_Dic() : Set e = v
    ao.put k, v, t
    Set a = ao(k)
    AssertSameWithMessage a, e, "ao('hoge') first put"
    
    v = 123 : e = v
    ao.put k, v, t
    a = ao(k)
    AssertEqualWithMessage a, e, "ao('hoge') second put"
End Sub

'###################################################################################################
'Cash.size
Sub Test_Cash_size
    Dim ao,a,t,e,waitTime
    Set ao = new Cash : t = 50 : waitTime = 50
    
    e = 0 : a = ao.size()
    AssertEqualWithMessage e, a, "ao.size() initially"
    
    ao.put "hoge", 123, -1

    e = 1 : a = ao.size()
    AssertEqualWithMessage e, a, "ao.size() after put1"

    ao.put "foo", new_Dic(), t
    
    e = 2 : a = ao.size()
    AssertEqualWithMessage e, a, "ao.size() after put2"

    ao.put "hoge", 789, -1
    
    e = 2 : a = ao.size()
    AssertEqualWithMessage e, a, "ao.size() after put3 (override existing key)"

    ao.delete "hoge"

    e = 1 : a = ao.size()
    AssertEqualWithMessage e, a, "ao.size() after delete"

    WScript.Sleep waitTime

    e = 0 : a = ao.size()
    AssertEqualWithMessage e, a, "ao.size() after sleep (timeout key expired)"
End Sub



'###################################################################################################
'common
Function getOrComputeLoader(arg)
    getOrComputeLoader = LOADER_PREFIX & CStr(arg)
End Function

' Local Variables:
' mode: Visual-Basic
' indent-tabs-mode: nil
' End:
