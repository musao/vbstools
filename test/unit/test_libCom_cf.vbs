' libCom.vbs: cf_* procedure test.
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

'###################################################################################################
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

'###################################################################################################
'cf_fillChar()
Sub Test_cf_fillChar
    Dim data, caseName : caseName = "fillChar_normal_"
    data = Array ( _
            new_DicOf(Array(  "Case", caseName & "1-1-1", "target", "a"     , "wordCount", 5, "toFillCharacter", "@", "isCutOut", True, "isRightAlignment", True , "expect", "@@@@a" )) _
            , new_DicOf(Array("Case", caseName & "1-1-2", "target", "aa"    , "wordCount", 5, "toFillCharacter", "@", "isCutOut", True, "isRightAlignment", True , "expect", "@@@aa" )) _
            , new_DicOf(Array("Case", caseName & "1-1-3", "target", "aaaaa" , "wordCount", 5, "toFillCharacter", "@", "isCutOut", True, "isRightAlignment", True , "expect", "aaaaa" )) _
            , new_DicOf(Array("Case", caseName & "1-1-4", "target", "aaaaaa", "wordCount", 5, "toFillCharacter", "@", "isCutOut", True, "isRightAlignment", True , "expect", "aaaaa" )) _
            , new_DicOf(Array("Case", caseName & "1-2-1", "target", "a"     , "wordCount", 5, "toFillCharacter", "@", "isCutOut", True, "isRightAlignment", False, "expect", "a@@@@" )) _
            , new_DicOf(Array("Case", caseName & "1-2-2", "target", "aa"    , "wordCount", 5, "toFillCharacter", "@", "isCutOut", True, "isRightAlignment", False, "expect", "aa@@@" )) _
            , new_DicOf(Array("Case", caseName & "1-2-3", "target", "aaaaa" , "wordCount", 5, "toFillCharacter", "@", "isCutOut", True, "isRightAlignment", False, "expect", "aaaaa" )) _
            , new_DicOf(Array("Case", caseName & "1-2-4", "target", "aaaaaa", "wordCount", 5, "toFillCharacter", "@", "isCutOut", True, "isRightAlignment", False, "expect", "aaaaa" )) _
            , new_DicOf(Array("Case", caseName & "2-1-1", "target", "a"     , "wordCount", 5, "toFillCharacter", "@", "isCutOut", False,"isRightAlignment", True , "expect", "@@@@a" )) _
            , new_DicOf(Array("Case", caseName & "2-1-2", "target", "aa"    , "wordCount", 5, "toFillCharacter", "@", "isCutOut", False,"isRightAlignment", True , "expect", "@@@aa" )) _
            , new_DicOf(Array("Case", caseName & "2-1-3", "target", "aaaaa" , "wordCount", 5, "toFillCharacter", "@", "isCutOut", False,"isRightAlignment", True , "expect", "aaaaa" )) _
            , new_DicOf(Array("Case", caseName & "2-1-4", "target", "aaaaaa", "wordCount", 5, "toFillCharacter", "@", "isCutOut", False,"isRightAlignment", True , "expect", "aaaaaa")) _
            , new_DicOf(Array("Case", caseName & "2-2-1", "target", "a"     , "wordCount", 5, "toFillCharacter", "@", "isCutOut", False,"isRightAlignment", False, "expect", "a@@@@" )) _
            , new_DicOf(Array("Case", caseName & "2-2-2", "target", "aa"    , "wordCount", 5, "toFillCharacter", "@", "isCutOut", False,"isRightAlignment", False, "expect", "aa@@@" )) _
            , new_DicOf(Array("Case", caseName & "2-2-3", "target", "aaaaa" , "wordCount", 5, "toFillCharacter", "@", "isCutOut", False,"isRightAlignment", False, "expect", "aaaaa" )) _
            , new_DicOf(Array("Case", caseName & "2-2-4", "target", "aaaaaa", "wordCount", 5, "toFillCharacter", "@", "isCutOut", False,"isRightAlignment", False, "expect", "aaaaaa")) _
            )
    
    Dim d,a,target,wordCount,toFillCharacter,isCutOut,isRightAlignment,expect
    For Each d In data
        target = d("target")
        wordCount = d("wordCount")
        toFillCharacter = d("toFillCharacter")
        isCutOut = d("isCutOut")
        isRightAlignment = d("isRightAlignment")
        expect = d("expect")
        
        a = cf_fillChar( _
                target _
                , wordCount _
                , toFillCharacter _
                , isCutOut _
                , isRightAlignment _
                )
        AssertEqualWithMessage expect, a, "data="&cf_toString(d)
    Next
End Sub

'###################################################################################################
'cf_isAvailableObject()
Sub Test_cf_isAvailableObject
    Dim a,d,e
    Set d = new_Dic()
    e = True
    a = cf_isAvailableObject(d)
    AssertEqual e,a
End Sub
Sub Test_cf_isAvailableObject_Nothing
    Dim a,d,e
    Set d = Nothing
    e = False
    a = cf_isAvailableObject(d)
    AssertEqual e,a
End Sub
Sub Test_cf_isAvailableObject_Variable
    Dim a,d,e
    d = Empty
    e = False
    a = cf_isAvailableObject(d)
    AssertEqual e,a
End Sub

'###################################################################################################
'cf_isInteger()
Sub Test_cf_isInteger
    Dim data
    data = Array( _
        new_DicOf(Array(  "Data", Empty                , "Expect", False)) _
        , new_DicOf(Array("Data", Null                 , "Expect", False)) _
        , new_DicOf(Array("Data", new_Dic()            , "Expect", False)) _
        , new_DicOf(Array("Data", Array()              , "Expect", False)) _
        , new_DicOf(Array("Data", CInt(1)              , "Expect", True )) _
        , new_DicOf(Array("Data", CLng(999999)         , "Expect", True )) _
        , new_DicOf(Array("Data", CInt(-1)             , "Expect", True )) _
        , new_DicOf(Array("Data", CLng(-999999)        , "Expect", True )) _
        , new_DicOf(Array("Data", CInt(0)              , "Expect", True )) _
        , new_DicOf(Array("Data", CSng(10.1)           , "Expect", False)) _
        , new_DicOf(Array("Data", CDbl(1234.567890123) , "Expect", False)) _
        , new_DicOf(Array("Data", CSng(-10.1)          , "Expect", False)) _
        , new_DicOf(Array("Data", CDbl(-1234.567890123), "Expect", False)) _
        , new_DicOf(Array("Data", CCur("\1,000")       , "Expect", False)) _
        , new_DicOf(Array("Data", True                 , "Expect", False)) _
        , new_DicOf(Array("Data", CByte(0)             , "Expect", False)) _
        , new_DicOf(Array("Data", vbNullString         , "Expect", False)) _
        , new_DicOf(Array("Data", "abc"                , "Expect", False)) _
        , new_DicOf(Array("Data", "1.2"                , "Expect", False)) _
        , new_DicOf(Array("Data", "-1.2"               , "Expect", False)) _
        , new_DicOf(Array("Data", "192.168.11.52"      , "Expect", False)) _
        , new_DicOf(Array("Data", "2024/01/03"         , "Expect", False)) _
        , new_DicOf(Array("Data", "ÇTÇO"               , "Expect", True )) _
        , new_DicOf(Array("Data", "Å|ÇTÇO"             , "Expect", True )) _
        , new_DicOf(Array("Data", "ÇO"                 , "Expect", True )) _
        , new_DicOf(Array("Data", "äøéö"               , "Expect", False)) _
        )
    
    Dim ub : ub = Ubound(data)
    Dim i,d,e,a
    For i=0 To ub
        cf_bind d, data(i).Item("Data")
        cf_bind e, data(i).Item("Expect")
        a = cf_isInteger(d)
        AssertEqualWithMessage e, a, "No"&i&" Data="&cf_toString(d)&" Expect="&cf_toString(e)
    Next
End Sub

'###################################################################################################
'cf_isNonNegativeNumber()
Sub Test_cf_isNonNegativeNumber
    Dim data
    data = Array( _
        new_DicOf(Array(  "Data", Empty                , "Expect", False)) _
        , new_DicOf(Array("Data", Null                 , "Expect", False)) _
        , new_DicOf(Array("Data", new_Dic()            , "Expect", False)) _
        , new_DicOf(Array("Data", Array()              , "Expect", False)) _
        , new_DicOf(Array("Data", CInt(1)              , "Expect", True )) _
        , new_DicOf(Array("Data", CLng(999999)         , "Expect", True )) _
        , new_DicOf(Array("Data", CInt(-1)             , "Expect", False)) _
        , new_DicOf(Array("Data", CLng(-999999)        , "Expect", False)) _
        , new_DicOf(Array("Data", CInt(0)              , "Expect", True )) _
        , new_DicOf(Array("Data", CSng(10.1)           , "Expect", True )) _
        , new_DicOf(Array("Data", CDbl(1234.567890123) , "Expect", True )) _
        , new_DicOf(Array("Data", CSng(-10.1)          , "Expect", False)) _
        , new_DicOf(Array("Data", CDbl(-1234.567890123), "Expect", False)) _
        , new_DicOf(Array("Data", CCur("\1,000")       , "Expect", False)) _
        , new_DicOf(Array("Data", True                 , "Expect", False)) _
        , new_DicOf(Array("Data", CByte(0)             , "Expect", False)) _
        , new_DicOf(Array("Data", vbNullString         , "Expect", False)) _
        , new_DicOf(Array("Data", "abc"                , "Expect", False)) _
        , new_DicOf(Array("Data", "1.2"                , "Expect", True )) _
        , new_DicOf(Array("Data", "-1.2"               , "Expect", False)) _
        , new_DicOf(Array("Data", "192.168.11.52"      , "Expect", False)) _
        , new_DicOf(Array("Data", "2024/01/03"         , "Expect", False)) _
        , new_DicOf(Array("Data", "ÇTÇO"               , "Expect", True )) _
        , new_DicOf(Array("Data", "Å|ÇTÇO"             , "Expect", False)) _
        , new_DicOf(Array("Data", "ÇO"                 , "Expect", True )) _
        , new_DicOf(Array("Data", "äøéö"               , "Expect", False)) _
        )
    
    Dim ub : ub = Ubound(data)
    Dim i,d,e,a
    For i=0 To ub
        cf_bind d, data(i).Item("Data")
        cf_bind e, data(i).Item("Expect")
        a = cf_isNonNegativeNumber(d)
        AssertEqualWithMessage e, a, "No"&i&" Data="&cf_toString(d)&" Expect="&cf_toString(e)
    Next
End Sub

'###################################################################################################
'cf_isNumeric()
Sub Test_cf_isNumeric
    Dim data
    data = Array( _
        new_DicOf(Array(  "Data", Empty                , "Expect", False)) _
        , new_DicOf(Array("Data", Null                 , "Expect", False)) _
        , new_DicOf(Array("Data", new_Dic()            , "Expect", False)) _
        , new_DicOf(Array("Data", Array()              , "Expect", False)) _
        , new_DicOf(Array("Data", CInt(1)              , "Expect", True )) _
        , new_DicOf(Array("Data", CLng(999999)         , "Expect", True )) _
        , new_DicOf(Array("Data", CSng(10.1)           , "Expect", True )) _
        , new_DicOf(Array("Data", CDbl(1234.567890123) , "Expect", True )) _
        , new_DicOf(Array("Data", CCur("\1,000")       , "Expect", False)) _
        , new_DicOf(Array("Data", True                 , "Expect", False)) _
        , new_DicOf(Array("Data", CByte(0)             , "Expect", False)) _
        , new_DicOf(Array("Data", vbNullString         , "Expect", False)) _
        , new_DicOf(Array("Data", "abc"                , "Expect", False)) _
        , new_DicOf(Array("Data", "1.2"                , "Expect", True)) _
        , new_DicOf(Array("Data", "192.168.11.52"      , "Expect", False)) _
        , new_DicOf(Array("Data", "2024/01/03"         , "Expect", False)) _
        , new_DicOf(Array("Data", "ÇTÇO"               , "Expect", True)) _
        , new_DicOf(Array("Data", "äøéö"               , "Expect", False)) _
        )
    
    Dim ub : ub = Ubound(data)
    Dim i,d,e,a
    For i=0 To ub
        cf_bind d, data(i).Item("Data")
        cf_bind e, data(i).Item("Expect")
        a = cf_isNumeric(d)
        AssertEqualWithMessage e, a, "No"&i&" Data="&cf_toString(d)&" Expect="&cf_toString(e)
    Next
End Sub

'###################################################################################################
'cf_isPositiveInteger()
Sub Test_cf_isPositiveInteger
    Dim data
    data = Array( _
        new_DicOf(Array(  "Data", Empty                , "Expect", False)) _
        , new_DicOf(Array("Data", Null                 , "Expect", False)) _
        , new_DicOf(Array("Data", new_Dic()            , "Expect", False)) _
        , new_DicOf(Array("Data", Array()              , "Expect", False)) _
        , new_DicOf(Array("Data", CInt(1)              , "Expect", True )) _
        , new_DicOf(Array("Data", CLng(999999)         , "Expect", True )) _
        , new_DicOf(Array("Data", CInt(-1)             , "Expect", False)) _
        , new_DicOf(Array("Data", CLng(-999999)        , "Expect", False)) _
        , new_DicOf(Array("Data", CInt(0)              , "Expect", False)) _
        , new_DicOf(Array("Data", CSng(10.1)           , "Expect", False)) _
        , new_DicOf(Array("Data", CDbl(1234.567890123) , "Expect", False)) _
        , new_DicOf(Array("Data", CSng(-10.1)          , "Expect", False)) _
        , new_DicOf(Array("Data", CDbl(-1234.567890123), "Expect", False)) _
        , new_DicOf(Array("Data", CCur("\1,000")       , "Expect", False)) _
        , new_DicOf(Array("Data", True                 , "Expect", False)) _
        , new_DicOf(Array("Data", CByte(0)             , "Expect", False)) _
        , new_DicOf(Array("Data", vbNullString         , "Expect", False)) _
        , new_DicOf(Array("Data", "abc"                , "Expect", False)) _
        , new_DicOf(Array("Data", "1.2"                , "Expect", False)) _
        , new_DicOf(Array("Data", "-1.2"               , "Expect", False)) _
        , new_DicOf(Array("Data", "192.168.11.52"      , "Expect", False)) _
        , new_DicOf(Array("Data", "2024/01/03"         , "Expect", False)) _
        , new_DicOf(Array("Data", "ÇTÇO"               , "Expect", True )) _
        , new_DicOf(Array("Data", "Å|ÇTÇO"             , "Expect", False)) _
        , new_DicOf(Array("Data", "ÇO"                 , "Expect", False)) _
        , new_DicOf(Array("Data", "äøéö"               , "Expect", False)) _
        )
    
    Dim ub : ub = Ubound(data)
    Dim i,d,e,a
    For i=0 To ub
        cf_bind d, data(i).Item("Data")
        cf_bind e, data(i).Item("Expect")
        a = cf_isPositiveInteger(d)
        AssertEqualWithMessage e, a, "No"&i&" Data="&cf_toString(d)&" Expect="&cf_toString(e)
    Next
End Sub

'###################################################################################################
'cf_isSame()
Sub Test_cf_isSame
    Dim data : data = Array ( _
                Nothing _
                , CreateObject("Scripting.Dictionary") _
                , Null _
                , Empty _
                , vbNullString _
                , "abc" _
                , 123 _
                , "2035/3/15 14:25:09" _
                , "2035/3/15" _
                , "09:59:01" _
                , True _
                , False _
                )
    Dim a,b,i,j,e
    i=-1
    For Each a In data:i=i+1
        j=-1
        For Each b In data:j=j+1
            If i=j Then e=True Else e=False
            AssertEqualWithMessage e, cf_isSame(a,b), "i="&i&" j="&j&" data(i)="&cf_toString(a)&" data(j)="&cf_toString(b)
        Next
    Next
End Sub

'###################################################################################################
'cf_isValid()
Sub Test_cf_isValid_Object_Valid
    Dim a,d,e
    Set d = new_Dic()
    e = True
    a = cf_isValid(d)
    AssertEqual e,a
End Sub
Sub Test_cf_isValid_Object_InValid
    Dim a,d,e
    Set d = Nothing
    e = False
    a = cf_isValid(d)
    AssertEqual e,a
End Sub
Sub Test_cf_isValid_Array_Valid
    Dim a,d,e
    d = Array(1,"2")
    e = True
    a = cf_isValid(d)
    AssertEqual e,a
End Sub
Sub Test_cf_isValid_Array_InValid_1
    Dim a,d,e
    d = Array()
    e = False
    a = cf_isValid(d)
    AssertEqual e,a
End Sub
Sub Test_cf_isValid_Array_InValid_2
    Dim a,e
    Dim d()
    e = False
    a = cf_isValid(d)
    AssertEqual e,a
End Sub
Sub Test_cf_isValid_Variable_Valid
    Dim a,d,e
    d = "a"
    e = True
    a = cf_isValid(d)
    AssertEqual e,a
End Sub
Sub Test_cf_isValid_Variable_InValid_Empty
    Dim a,d,e
    d = Empty
    e = False
    a = cf_isValid(d)
    AssertEqual e,a
End Sub
Sub Test_cf_isValid_Variable_InValid_Null
    Dim a,d,e
    d = Null
    e = False
    a = cf_isValid(d)
    AssertEqual e,a
End Sub
Sub Test_cf_isValid_Variable_InValid_NullString
    Dim a,d,e
    d = vbNullString
    e = False
    a = cf_isValid(d)
    AssertEqual e,a
End Sub

'###################################################################################################
'cf_push()
Sub Test_cf_push_Available
    Redim a(0)
    cf_push a, "NewValue"
    
    AssertEqual 1, Ubound(a)
    AssertEqual Empty, a(0)
    AssertEqual "NewValue", a(1)
End Sub
Sub Test_cf_push_NotAvailable
    Dim a
    cf_push a, "NewValue"
    
    AssertEqual 0, Ubound(a)
    AssertEqual "NewValue", a(0)
End Sub

'###################################################################################################
'cf_pushA()
Sub Test_cf_pushA_AddIsArray_ArrAvailable
    Dim a,d,e
    Redim a(0)
    d = Array(1,2)
    e = Array(Empty,1,2)
    cf_pushA a, d
    
    assertAllElements e, a
End Sub
Sub Test_cf_pushA_AddIsArray_ArrAvailableInitial
    Dim a,d,e
    Redim a(0)
    d = Array()
    e = Array(Empty)
    cf_pushA a, d
    
    assertAllElements e, a
End Sub
Sub Test_cf_pushA_AddIsArray_ArrNotAvailable
    Dim a,d,e
    d = Array(1,2)
    e = Array(1,2)
    cf_pushA a, d
    
    assertAllElements e, a
End Sub
Sub Test_cf_pushA_AddIsArray_ArrNotAvailable2
    Dim a(),d,e
    d = Array(1,2)
    e = Array(1,2)
    cf_pushA a, d
    
    assertAllElements e, a
End Sub
Sub Test_cf_pushA_AddIsZeroArray
    Dim a,d(),e
    Redim a(0)
    e = Array(Empty)
    cf_pushA a, d
    
    assertAllElements e, a
End Sub
Sub Test_cf_pushA_AddIsNotArray_ArrAvailable
    Dim a,d,e
    Redim a(0)
    d = "a"
    e = Array(Empty,"a")
    cf_pushA a, d
    
    assertAllElements e, a
End Sub
Sub Test_cf_pushA_AddIsNotArray_ArrNotAvailable
    Dim a,d,e
    d = "a"
    e = Array("a")
    cf_pushA a, d
    
    assertAllElements e, a
End Sub

'###################################################################################################
'cf_swap()
Sub Test_cf_swap
    Dim d
    d = Array( _
            Array("a",1) _
            , Array(new_Dic(),new_Fso()) _
            , Array("Ç†",new_Fso()) _
            )
    
    Dim i,a,da,db
    For Each i In d
        cf_bind i(0), da
        cf_bind i(1), db
        cf_swap da, db
        AssertEqualWithMessage i(0), db, i(0)&"vs"&i(1)
        AssertEqualWithMessage i(1), da, i(0)&"vs"&i(1)
    Next
End Sub

'###################################################################################################
'cf_toString()
Sub Test_cf_toString_Empty
    Dim a,d,e
    d = Empty
    e = "<Empty>"
    a = cf_toString(d)
    AssertEqual e,a
End Sub
Sub Test_cf_toString_Null
    Dim a,d,e
    d = Null
    e = "<Null>"
    a = cf_toString(d)
    AssertEqual e,a
End Sub
Sub Test_cf_toString_Integer
    Dim a,d,e
    d = CInt(100)
    e = "<Integer>100"
    a = cf_toString(d)
    AssertEqual e,a
End Sub
Sub Test_cf_toString_Long
    Dim a,d,e
    d = CLng(99999999)
    e = "<Long>99999999"
    a = cf_toString(d)
    AssertEqual e,a
End Sub
Sub Test_cf_toString_Single
    Dim a,d,e
    d = CSng(1.23)
    e = "<Single>1.23"
    a = cf_toString(d)
    AssertEqual e,a
End Sub
Sub Test_cf_toString_Double
    Dim a,d,e
    d = CDbl(1.23)
    e = "<Double>1.23"
    a = cf_toString(d)
    AssertEqual e,a
End Sub
Sub Test_cf_toString_Currency
    Dim a,d,e
    d = CCur(100.1)
    e = "<Currency>100.1"
    a = cf_toString(d)
    AssertEqual e,a
End Sub
Sub Test_cf_toString_Date
    Dim a,d,e
    d = #2023-01-24 18:12:04#
    e = "<Date>2023/01/24 18:12:04"
    a = cf_toString(d)
    AssertEqual e,a
End Sub
Sub Test_cf_toString_String
    Dim a,d,e
    d = "foo"
    e = "<String>" & Chr(34) & d & Chr(34)
    a = cf_toString(d)
    AssertEqual e,a
End Sub
Sub Test_cf_toString_String_vbNullString
    Dim a,d,e
    d = vbNullString
    e = "<String>" & Chr(34)&Chr(34)
    a = cf_toString(d)
    AssertEqual e,a
End Sub
Sub Test_cf_toString_String_ContainsDoubleQuotes
    Dim a,d,e
    d = "foo" & Chr(34) & "bar"
    e = "<String>" & Chr(34) & Replace(d,Chr(34),Chr(34)&Chr(34)) & Chr(34)
    a = cf_toString(d)
    AssertEqual e,a
End Sub
Sub Test_cf_toString_Object_Dictionary
    Dim a,d,e
    Set d = new_DicOf(Array("foo","apple","bar",5))
    e = "<Dictionary>{<String>" & Chr(34) & "foo" & Chr(34) & "=><String>" & Chr(34) & "apple" & Chr(34) & ",<String>" & Chr(34) & "bar" & Chr(34) & "=><Integer>5}"
    a = cf_toString(d)
    AssertEqual e,a
End Sub
Sub Test_cf_toString_Object_Dictionary_Empty
    Dim a,d,e
    Set d = new_Dic()
    e = "<Dictionary>{}"
    a = cf_toString(d)
    AssertEqual e,a
End Sub
Sub Test_cf_toString_Object_Nothing
    Dim a,d,e
    Set d = Nothing
    e = "<Nothing>"
    a = cf_toString(d)
    AssertEqual e,a
End Sub
Sub Test_cf_toString_Object_Other_ShellApp
    Dim a,d,e
    Set d = new_ShellApp()
    e = "<IShellDispatch6>"
    a = cf_toString(d)
    AssertEqual e,a
End Sub
Sub Test_cf_toString_Object_Other_UserDef
    Dim a,d,e
    Set d = new_Char()
    e = "<CharacterType>"
    a = cf_toString(d)
    AssertEqual e,a
End Sub
Sub Test_cf_toString_Object_Other_UserDef_Special
    Dim a,d,e
    Set d = new_DicOf(Array("__Special__", "Test", "Key", "Value"))
    e = "<Test>{<String>" & Chr(34) & "Key" & Chr(34) & "=><String>" & Chr(34) & "Value" & Chr(34) & "}"
    a = cf_toString(d)
    AssertEqual e,a
End Sub
Sub Test_cf_toString_Boolean
    Dim a,d,e
    d = True
    e = "<Boolean>True"
    a = cf_toString(d)
    AssertEqual e,a
End Sub
Sub Test_cf_toString_Byte
    Dim a,d,e
    d = CByte(1)
    e = "<Byte>1"
    a = cf_toString(d)
    AssertEqual e,a
End Sub
Sub Test_cf_toString_Array
    Dim a,d,e
    d = Array(1,"a",Array("éQ",4),new_DicOf(Array("áX",6,7,"ÇW")))
    e = "<Array>[<Integer>1,<String>" & Chr(34) & "a" & Chr(34) & ",<Array>[<String>" & Chr(34) & "éQ" & Chr(34) & ",<Integer>4],<Dictionary>{<String>" & Chr(34) & "áX" & Chr(34) & "=><Integer>6,<Integer>7=><String>" & Chr(34) & "ÇW" & Chr(34) & "}]"
    a = cf_toString(d)
    AssertEqual e,a
End Sub
Sub Test_cf_toString_Array_Empty
    Dim a,d,e
    d = Array()
    e = "<Array>[]"
    a = cf_toString(d)
    AssertEqual e,a
End Sub

'###################################################################################################
'common
Sub assertAllElements(e,a)
    AssertEqualWithMessage Ubound(e), Ubound(a), "Ubound"
    Dim i
    For i=0 To Ubound(e)
        If IsObject(e(i)) Then
            AssertSameWithMessage e(i), a(i), "Element Object"
        Else
            AssertEqualWithMessage e(i), a(i), "Element Variable"
        End If
    Next
End Sub

' Local Variables:
' mode: Visual-Basic
' indent-tabs-mode: nil
' End:
