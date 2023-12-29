' libCom.vbs: util_* procedure test.
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

'###################################################################################################
'util_randStr()
Sub Test_util_randStr
    dim d,a,s
    s = 1000
    With new_Char()
        d = .getCharList(.typeHalfWidthNumbers)
    End With
    
    a = util_randStr(d,s)
    Dim i,j,t,flg : i=1
    Do While i<Len(a)
        t = Mid(a,i,1)
        flg = false
        For Each j In d
            If cf_isSame(t,j) Then
                flg=True
                Exit For
            End If
        Next
        If Not flg Then
            AssertFailWithMessage "util_randStr(d,s)="&a&", s="&s&", i="&i
        End If
        i=i+1
    Loop
    Assert True
End Sub

'###################################################################################################
'util_getIpAddress()
Sub Test_util_getIpAddress
    Const RE_IP4 = "^(([1-9]?[0-9]|1[0-9]{2}|2[0-4][0-9]|25[0-5])\.){3}([1-9]?[0-9]|1[0-9]{2}|2[0-4][0-9]|25[0-5])$"
    Const RE_IP6 = "^(([0-9a-fA-F]{1,4}:){7,7}[0-9a-fA-F]{1,4}|([0-9a-fA-F]{1,4}:){1,7}:|([0-9a-fA-F]{1,4}:){1,6}:[0-9a-fA-F]{1,4}|([0-9a-fA-F]{1,4}:){1,5}(:[0-9a-fA-F]{1,4}){1,2}|([0-9a-fA-F]{1,4}:){1,4}(:[0-9a-fA-F]{1,4}){1,3}|([0-9a-fA-F]{1,4}:){1,3}(:[0-9a-fA-F]{1,4}){1,4}|([0-9a-fA-F]{1,4}:){1,2}(:[0-9a-fA-F]{1,4}){1,5}|[0-9a-fA-F]{1,4}:((:[0-9a-fA-F]{1,4}){1,6})|:((:[0-9a-fA-F]{1,4}){1,7}|:)|fe80:(:[0-9a-fA-F]{0,4}){0,4}%[0-9a-zA-Z]{1,}|::(ffff(:0{1,4}){0,1}:){0,1}((25[0-5]|(2[0-4]|1{0,1}[0-9]){0,1}[0-9])\.){3,3}(25[0-5]|(2[0-4]|1{0,1}[0-9]){0,1}[0-9])|([0-9a-fA-F]{1,4}:){1,4}:((25[0-5]|(2[0-4]|1{0,1}[0-9]){0,1}[0-9])\.){3,3}(25[0-5]|(2[0-4]|1{0,1}[0-9]){0,1}[0-9]))$"
    
    dim a
    a = util_getIpAddress()

    dim i
    For i=0 To Ubound(a)
        AssertWithMessage Len(a(i).Item("Caption"))>0, "i="&i&":Caption"
        AssertMatchWithMessage RE_IP4, a(i).Item("Ip").Item("V4"), "i="&i&":IpV4"
        AssertMatchWithMessage RE_IP6, a(i).Item("Ip").Item("V6"), "i="&i&":IpV6"
    Next
End Sub

' Local Variables:
' mode: Visual-Basic
' indent-tabs-mode: nil
' End:
