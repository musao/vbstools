' clsCmCalendar.vbs: test.
' @import ../lib/clsCmArray.vbs
' @import ../lib/clsCmBufferedReader.vbs
' @import ../lib/clsCmBufferedWriter.vbs
' @import ../lib/clsCmCalendar.vbs
' @import ../lib/clsCmBroker.vbs
' @import ../lib/clsCompareExcel.vbs
' @import ../lib/clsCmCharacterType.vbs
' @import ../lib/clsFsBase.vbs
' @import ../lib/libCom.vbs

Option Explicit

'###################################################################################################
'clsCmCharacterType
Sub Test_clsCmCharacterType
    Dim a : Set a = new clsCmCharacterType
    AssertEqual 9, VarType(a)
    AssertEqual "clsCmCharacterType", TypeName(a)
End Sub

'###################################################################################################
'clsCmCharacterType.whatType()
Sub Test_clsCmCharacterType_whatType
    Dim ao,a,d,dc,e,i
    Set ao = new clsCmCharacterType
    d = Array( _
        new_DicWith(Array("No", "1-1", "Char", "A", "Expected", 2^0)) _
        , new_DicWith(Array("No", "1-2", "Char", "Q", "Expected", 2^0)) _
        , new_DicWith(Array("No", "2-1", "Char", "g", "Expected", 2^1)) _
        , new_DicWith(Array("No", "2-2", "Char", "z", "Expected", 2^1)) _
        , new_DicWith(Array("No", "3-1", "Char", "0", "Expected", 2^2)) _
        , new_DicWith(Array("No", "3-2", "Char", "3", "Expected", 2^2)) _
        , new_DicWith(Array("No", "4-1", "Char", " ", "Expected", 2^3)) _
        , new_DicWith(Array("No", "4-2", "Char", "~", "Expected", 2^3)) _
        , new_DicWith(Array("No", "5-1", "Char", "¶", "Expected", 2^4)) _
        , new_DicWith(Array("No", "5-2", "Char", "ﬂ", "Expected", 2^4)) _
        , new_DicWith(Array("No", "6-1", "Char", "°", "Expected", 2^5)) _
        , new_DicWith(Array("No", "6-2", "Char", "∞", "Expected", 2^5)) _
        , new_DicWith(Array("No", "7-1", "Char", "Çl", "Expected", 2^6)) _
        , new_DicWith(Array("No", "7-2", "Char", "Çy", "Expected", 2^6)) _
        , new_DicWith(Array("No", "8-1", "Char", "ÇÅ", "Expected", 2^7)) _
        , new_DicWith(Array("No", "8-2", "Char", "Çí", "Expected", 2^7)) _
        , new_DicWith(Array("No", "9-1", "Char", "ÇV", "Expected", 2^8)) _
        , new_DicWith(Array("No", "9-2", "Char", "ÇX", "Expected", 2^8)) _
        , new_DicWith(Array("No", "10-1", "Char", "Å~", "Expected", 2^9)) _
        , new_DicWith(Array("No", "10-2", "Char", "ÅÄ", "Expected", 2^9)) _
        , new_DicWith(Array("No", "11-1", "Char", "Ç«", "Expected", 2^10)) _
        , new_DicWith(Array("No", "11-2", "Char", "ÇÔ", "Expected", 2^10)) _
        , new_DicWith(Array("No", "12-1", "Char", "É~", "Expected", 2^11)) _
        , new_DicWith(Array("No", "12-2", "Char", "ÉÄ", "Expected", 2^11)) _
        , new_DicWith(Array("No", "13-1", "Char", "Éü", "Expected", 2^12)) _
        , new_DicWith(Array("No", "13-2", "Char", "Ñ`", "Expected", 2^12)) _
        , new_DicWith(Array("No", "14-1", "Char", "É÷", "Expected", 2^13)) _
        , new_DicWith(Array("No", "14-2", "Char", "Ñp", "Expected", 2^13)) _
        , new_DicWith(Array("No", "15-1", "Char", "Ñü", "Expected", 2^14)) _
        , new_DicWith(Array("No", "15-2", "Char", "Ñæ", "Expected", 2^14)) _
        , new_DicWith(Array("No", "16-1", "Char", "ì~", "Expected", 2^15)) _
        , new_DicWith(Array("No", "16-2", "Char", "ìÄ", "Expected", 2^15)) _
        , new_DicWith(Array("No", "17-1", "Char", "Ë~", "Expected", 2^16)) _
        , new_DicWith(Array("No", "17-2", "Char", "ËÄ", "Expected", 2^16)) _
        )
    
    For Each i In d
        dc = i.Item("Char")
        e = i.Item("Expected")
        a = ao.whatType(dc)
        AssertEqualWithMessage e, a, "No="&i.Item("No")&", Char="&dc
    Next
End Sub

'###################################################################################################
'clsCmCharacterType.getCharList()
Sub Test_clsCmCharacterType_getCharList
    Dim ao,a,d,e,i,j
    Set ao = new clsCmCharacterType
    
    For i=0 To 16
        d = 2^i
        For Each j In ao.getCharList(d)
            e = d
            a = ao.whatType(Chr(j))
            AssertEqualWithMessage e, a, "No="&i&", Type="&(2^i)&", Char="&Chr(j)
        Next
    Next
End Sub

' Local Variables:
' mode: Visual-Basic
' indent-tabs-mode: nil
' End:
