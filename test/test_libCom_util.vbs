' libCom.vbs: util_* procedure test.
' @import ../lib/clsCmArray.vbs
' @import ../lib/clsCmBroker.vbs
' @import ../lib/clsCmBufferedReader.vbs
' @import ../lib/clsCmBufferedWriter.vbs
' @import ../lib/clsCmCalendar.vbs
' @import ../lib/clsCmCharacterType.vbs
' @import ../lib/clsCmCssGenerator.vbs
' @import ../lib/clsCmHtmlGenerator.vbs
' @import ../lib/clsCompareExcel.vbs
' @import ../lib/libCom.vbs

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

' Local Variables:
' mode: Visual-Basic
' indent-tabs-mode: nil
' End:
