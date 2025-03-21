' CharacterType.vbs: test.
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
'CharacterType
Sub Test_CharacterType
    Dim a : Set a = new CharacterType
    AssertEqual 9, VarType(a)
    AssertEqual "CharacterType", TypeName(a)
End Sub

'###################################################################################################
'CharacterType.type*()
Sub Test_CharacterType_type_
    Dim ao,a,e
    Set ao = new CharacterType

    e = 2^0
    a = ao.typeHalfWidthAlphabetUppercase
    AssertEqualWithMessage e, a, "typeHalfWidthAlphabetUppercase"

    e = e*2
    a = ao.typeHalfWidthAlphabetLowercase
    AssertEqualWithMessage e, a, "typeHalfWidthAlphabetLowercase"

    e = e*2
    a = ao.typeHalfWidthNumbers
    AssertEqualWithMessage e, a, "typeHalfWidthNumbers"

    e = e*2
    a = ao.typeHalfWidthSymbol
    AssertEqualWithMessage e, a, "typeHalfWidthSymbol"

    e = e*2
    a = ao.typeHalfWidthKatakana
    AssertEqualWithMessage e, a, "typeHalfWidthKatakana"

    e = e*2
    a = ao.typeHalfWidthKatakanaSymbol
    AssertEqualWithMessage e, a, "typeHalfWidthKatakanaSymbol"

    e = e*2
    a = ao.typeFullWidthAlphabetUppercase
    AssertEqualWithMessage e, a, "typeFullWidthAlphabetUppercase"

    e = e*2
    a = ao.typeFullWidthAlphabetLowercase
    AssertEqualWithMessage e, a, "typeFullWidthAlphabetLowercase"

    e = e*2
    a = ao.typeFullWidthNumbers
    AssertEqualWithMessage e, a, "typeFullWidthNumbers"

    e = e*2
    a = ao.typeFullWidthSymbol
    AssertEqualWithMessage e, a, "typeFullWidthSymbol"

    e = e*2
    a = ao.typeFullWidthHiragana
    AssertEqualWithMessage e, a, "typeFullWidthHiragana"

    e = e*2
    a = ao.typeFullWidthKatakana
    AssertEqualWithMessage e, a, "typeFullWidthKatakana"

    e = e*2
    a = ao.typeFullWidthGreekCyrillicUppercase
    AssertEqualWithMessage e, a, "typeFullWidthGreekCyrillicUppercase"

    e = e*2
    a = ao.typeFullWidthGreekCyrillicLowercase
    AssertEqualWithMessage e, a, "typeFullWidthGreekCyrillicLowercase"

    e = e*2
    a = ao.typeFullWidthLineFrame
    AssertEqualWithMessage e, a, "typeFullWidthLineFrame"

    e = e*2
    a = ao.typeFullWidthKanjiLevel1
    AssertEqualWithMessage e, a, "typeFullWidthKanjiLevel1"

    e = e*2
    a = ao.typeFullWidthKanjiLevel2
    AssertEqualWithMessage e, a, "typeFullWidthKanjiLevel2"
End Sub
Sub Test_CharacterType_typeSum_
    Dim ao,a,e,i,stt,ed
    Set ao = new CharacterType

    stt = math_log2(ao.typeHalfWidthAlphabetUppercase)
    ed = math_log2(ao.typeFullWidthKanjiLevel2)
    e = 0
    For i=stt To ed
        e = e + 2^i
    Next
    a = ao.typeAll
    AssertEqualWithMessage e, a, "typeAll"

    stt = math_log2(ao.typeHalfWidthAlphabetUppercase)
    ed = math_log2(ao.typeHalfWidthNumbers)
    e = 0
    For i=stt To ed
        e = e + 2^i
    Next
    a = ao.typeHalfWidthAlphanumeric
    AssertEqualWithMessage e, a, "typeHalfWidthAlphanumeric"

    stt = math_log2(ao.typeHalfWidthAlphabetUppercase)
    ed = math_log2(ao.typeHalfWidthSymbol)
    e = 0
    For i=stt To ed
        e = e + 2^i
    Next
    a = ao.typeHalfWidthAlphanumericSymbols
    AssertEqualWithMessage e, a, "typeHalfWidthAlphanumericSymbols"

    stt = math_log2(ao.typeHalfWidthAlphabetUppercase)
    ed = math_log2(ao.typeHalfWidthKatakanaSymbol)
    e = 0
    For i=stt To ed
        e = e + 2^i
    Next
    a = ao.typeHalfWidthAll
    AssertEqualWithMessage e, a, "typeHalfWidthAll"

    stt = math_log2(ao.typeFullWidthAlphabetUppercase)
    ed = math_log2(ao.typeFullWidthKanjiLevel2)
    e = 0
    For i=stt To ed
        e = e + 2^i
    Next
    a = ao.typeFullWidthAll
    AssertEqualWithMessage e, a, "typeFullWidthAll"
End Sub

'###################################################################################################
'CharacterType.whatType()
Sub Test_CharacterType_whatType
    Dim ao,a,d,dc,e,i
    Set ao = new CharacterType
    d = Array( _
        new_DicOf(Array("No", "1-1", "Char", "A", "Expected", 2^0)) _
        , new_DicOf(Array("No", "1-2", "Char", "Q", "Expected", 2^0)) _
        , new_DicOf(Array("No", "2-1", "Char", "g", "Expected", 2^1)) _
        , new_DicOf(Array("No", "2-2", "Char", "z", "Expected", 2^1)) _
        , new_DicOf(Array("No", "3-1", "Char", "0", "Expected", 2^2)) _
        , new_DicOf(Array("No", "3-2", "Char", "3", "Expected", 2^2)) _
        , new_DicOf(Array("No", "4-1", "Char", " ", "Expected", 2^3)) _
        , new_DicOf(Array("No", "4-2", "Char", "~", "Expected", 2^3)) _
        , new_DicOf(Array("No", "5-1", "Char", "�", "Expected", 2^4)) _
        , new_DicOf(Array("No", "5-2", "Char", "�", "Expected", 2^4)) _
        , new_DicOf(Array("No", "6-1", "Char", "�", "Expected", 2^5)) _
        , new_DicOf(Array("No", "6-2", "Char", "�", "Expected", 2^5)) _
        , new_DicOf(Array("No", "7-1", "Char", "�l", "Expected", 2^6)) _
        , new_DicOf(Array("No", "7-2", "Char", "�y", "Expected", 2^6)) _
        , new_DicOf(Array("No", "8-1", "Char", "��", "Expected", 2^7)) _
        , new_DicOf(Array("No", "8-2", "Char", "��", "Expected", 2^7)) _
        , new_DicOf(Array("No", "9-1", "Char", "�V", "Expected", 2^8)) _
        , new_DicOf(Array("No", "9-2", "Char", "�X", "Expected", 2^8)) _
        , new_DicOf(Array("No", "10-1", "Char", "�~", "Expected", 2^9)) _
        , new_DicOf(Array("No", "10-2", "Char", "��", "Expected", 2^9)) _
        , new_DicOf(Array("No", "11-1", "Char", "��", "Expected", 2^10)) _
        , new_DicOf(Array("No", "11-2", "Char", "��", "Expected", 2^10)) _
        , new_DicOf(Array("No", "12-1", "Char", "�~", "Expected", 2^11)) _
        , new_DicOf(Array("No", "12-2", "Char", "��", "Expected", 2^11)) _
        , new_DicOf(Array("No", "13-1", "Char", "��", "Expected", 2^12)) _
        , new_DicOf(Array("No", "13-2", "Char", "�`", "Expected", 2^12)) _
        , new_DicOf(Array("No", "14-1", "Char", "��", "Expected", 2^13)) _
        , new_DicOf(Array("No", "14-2", "Char", "�p", "Expected", 2^13)) _
        , new_DicOf(Array("No", "15-1", "Char", "��", "Expected", 2^14)) _
        , new_DicOf(Array("No", "15-2", "Char", "��", "Expected", 2^14)) _
        , new_DicOf(Array("No", "16-1", "Char", "�~", "Expected", 2^15)) _
        , new_DicOf(Array("No", "16-2", "Char", "��", "Expected", 2^15)) _
        , new_DicOf(Array("No", "17-1", "Char", "�~", "Expected", 2^16)) _
        , new_DicOf(Array("No", "17-2", "Char", "�", "Expected", 2^16)) _
        )
    
    For Each i In d
        dc = i.Item("Char")
        e = i.Item("Expected")
        a = ao.whatType(dc)
        AssertEqualWithMessage e, a, "No="&i.Item("No")&", Char="&dc
    Next
End Sub

'###################################################################################################
'CharacterType.charList()
Sub Test_CharacterType_getCharList
    Dim ao,a,d,e,i,j
    Set ao = new CharacterType
    
    For i=0 To 16
        d = 2^i
        For Each j In ao.charList(d)
            e = d
            a = ao.whatType(j)
            AssertEqualWithMessage e, a, "No="&i&", Type="&(2^i)&", Char="&j
        Next
    Next
End Sub

' Local Variables:
' mode: Visual-Basic
' indent-tabs-mode: nil
' End:
