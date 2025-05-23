' libCom.vbs: new_* procedure test.
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
'new_Adodb()
Sub Test_new_Adodb
    Dim e : Set e = CreateObject("ADODB.Stream")
    Dim a : Set a = new_Adodb()
    
    AssertEqual VarType(e), VarType(a)
    AssertEqual TypeName(e), TypeName(a)
End Sub

'###################################################################################################
'new_AdptFile()
'###################################################################################################
'new_FspOf()

'###################################################################################################
'new_Arr()
Sub Test_new_Arr
    Dim e : Set e = New ArrayList
    Dim a : Set a = new_Arr()
    
    AssertEqual VarType(e), VarType(a)
    AssertEqual TypeName(e), TypeName(a)
    AssertEqual 0, a.Length
End Sub

'###################################################################################################
'new_ArrSplit()
Sub Test_new_ArrSplit
    Dim e : Set e = New ArrayList
    Dim es : es = "one,弐,3"
    Dim ev : ev = Split(es, ",")
    Dim a : Set a = new_ArrSplit(es, ",")
    
    AssertEqual VarType(e), VarType(a)
    AssertEqual TypeName(e), TypeName(a)
    AssertEqual Ubound(ev)+1, a.Length
    AssertEqual ev(0), a(0)
    AssertEqual ev(1), a(1)
    AssertEqual ev(2), a(2)
End Sub

'###################################################################################################
'new_ArrOf()
Sub Test_new_ArrOf_Array
    Dim e : Set e = New ArrayList
    Dim ev : ev = Array(1,Nothing,"三")
    Dim a : Set a = new_ArrOf(ev)
    
    AssertEqual VarType(e), VarType(a)
    AssertEqual TypeName(e), TypeName(a)
    AssertEqual Ubound(ev)+1, a.Length
    AssertEqual ev(0), a(0)
    AssertSame ev(1), a(1)
    AssertEqual ev(2), a(2)
End Sub
Sub Test_new_ArrOf_Array_0
    Dim e : Set e = New ArrayList
    Dim ev : ev = Array()
    Dim a : Set a = new_ArrOf(ev)
    
    AssertEqual 0, a.Length
End Sub
Sub Test_new_ArrOf_Variable
    Dim ev : ev = "abc"
    Dim a : Set a = new_ArrOf(ev)
    
    AssertEqual 1, a.Length
    AssertEqual "abc", a(0)
End Sub

'###################################################################################################
'new_Broker()
Sub Test_new_Broker
    Dim e : Set e = New Broker
    Dim a : Set a = new_Broker()
    
    AssertEqual VarType(e), VarType(a)
    AssertEqual TypeName(e), TypeName(a)
End Sub

'###################################################################################################
'new_BrokerOf()
Sub Test_new_BrokerOf_1Arg
    Dim d,ao,e,a
    d = Array("test1")
    Set ao = new_BrokerOf(d)

    e = 1 : a = 1
    ao.publish "test1",a
    AssertEqualWithMessage e, a, "test1 unchanging"
End Sub
Sub Test_new_BrokerOf_2Args
    Dim d,ao,e,a
    d = Array("test1", new_Func("function(a){a=2*a}"))
    Set ao = new_BrokerOf(d)

    e = 2 : a = 1
    ao.publish "test1",a
    AssertEqualWithMessage e, a, "test1"
End Sub
Sub Test_new_BrokerOf_3Args
    Dim d,ao,e,a
    d = Array("test1", new_Func("function(a){a=2*a}"), "test2")
    Set ao = new_BrokerOf(d)

    e = 2 : a = 1
    ao.publish "test1",a
    AssertEqualWithMessage e, a, "test1"

    e = 1 : a = 1
    ao.publish "test2",a
    AssertEqualWithMessage e, a, "test2 unchanging"
End Sub
Sub Test_new_BrokerOf_4Args
    Dim d,ao,e,a
    d = Array("test1", new_Func("function(a){a=2*a}"), "test2", new_Func("function(a){a=10*a}"))
    Set ao = new_BrokerOf(d)

    e = 2 : a = 1
    ao.publish "test1",a
    AssertEqualWithMessage e, a, "test1"

    e = 10 : a = 1
    ao.publish "test2",a
    AssertEqualWithMessage e, a, "test2"
End Sub

'###################################################################################################
'new_CalAt()
Sub Test_new_CalAt
    Dim e : Set e = New Calendar
    Dim ed : ed = CDate("2024/2/29")
    Dim a : Set a = new_CalAt(ed)
    
    AssertEqual VarType(e), VarType(a)
    AssertEqual TypeName(e), TypeName(a)
    AssertEqual Cstr(DatePart("yyyy", ed)), a.formatAs("YYYY")
    AssertEqual Cstr(DatePart("m", ed)), a.formatAs("M")
    AssertEqual Cstr(DatePart("d", ed)), a.formatAs("D")
End Sub
Sub Test_new_CalAt_Err
    On Error Resume Next
    Dim a : Set a = new_CalAt(vbNullString)
    Dim e : e = Empty
    
    AssertEqualWithMessage e, a, "ret"
    AssertEqualWithMessage "Calendar+of()", Err.Source, "Err.Source"
    AssertMatchWithMessage "^invalid argument.*", Err.Description, "Err.Description"
End Sub

'###################################################################################################
'new_Char()
Sub Test_new_Char
    Dim e : Set e = New CharacterType
    Dim a : Set a = new_Char()
    
    AssertEqual VarType(e), VarType(a)
    AssertEqual TypeName(e), TypeName(a)
End Sub

'###################################################################################################
'new_CssOf()
Sub Test_new_CssOf
    Dim e : Set e = New CssGenerator
    Dim a : Set a = new_CssOf(".hoge")
    
    AssertEqual VarType(e), VarType(a)
    AssertEqual TypeName(e), TypeName(a)
End Sub
'Sub Test_new_CssOf_Err
'    On Error Resume Next
'    Dim a : Set a = new_CssOf("．Ｈｏｇｅ")
'    
'    AssertEqual 1032, Err.Number
'    AssertEqual "セレクタには半角以外の文字を指定できません。", Err.Description
'    AssertEqual Empty, a
'End Sub

'###################################################################################################
'new_Dic()
Sub Test_new_Dic
    Dim e : Set e = CreateObject("Scripting.Dictionary")
    Dim a : Set a = new_Dic()
    
    AssertEqual VarType(e), VarType(a)
    AssertEqual TypeName(e), TypeName(a)
    AssertEqual 0, a.Count
End Sub

'###################################################################################################
'new_DicOf()
Sub Test_new_DicOf_Normal
    Dim e : Set e = CreateObject("Scripting.Dictionary")
    Dim a : Set a = new_DicOf(Array(1,2,3))
    
    AssertEqual VarType(e), VarType(a)
    AssertEqual TypeName(e), TypeName(a)
End Sub
Sub Test_new_DicOf_EvenNumber
    Dim ev : ev = Array("first", "一", "Second", Nothing, "3rd", 3)
    Dim a : Set a = new_DicOf(ev)
    
    AssertEqual (Ubound(ev)+1)/2, a.Count
    AssertEqual ev(1), a.Item(ev(0))
    AssertSame ev(3), a.Item(ev(2))
    AssertEqual ev(5), a.Item(ev(4))
End Sub
Sub Test_new_DicOf_OddNumber
    Dim ev : ev = Array("first", "一", "Second", Nothing, "3rd")
    Dim a : Set a = new_DicOf(ev)
    
    AssertEqual Ubound(ev)/2+1, a.Count
    AssertEqual ev(1), a.Item(ev(0))
    AssertSame ev(3), a.Item(ev(2))
    AssertEqual Empty, a.Item(ev(4))
End Sub

'###################################################################################################
'new_DriveOf()
Sub Test_new_DriveOf
    Dim d,e,a
    d = "c"
    Set e = CreateObject("Scripting.FileSystemObject").GetDrive(d)
    Set a = new_DriveOf(d)
    
    AssertEqual VarType(e), VarType(a)
    AssertEqual TypeName(e), TypeName(a)
End Sub
Sub Test_new_DriveOf_Err
    On Error Resume Next
    Dim a : Set a = new_DriveOf(vbNullString)
    
    AssertEqual 5, Err.Number
    AssertEqual "プロシージャの呼び出し、または引数が不正です。", Err.Description
    AssertEqual Empty, a
End Sub

'###################################################################################################
'new_Enum()
Sub Test_new_Enum
    Dim def : Set def = createTestEnum

    AssertEqualWithMessage 1, GREAT_SATAN_KOSAKA.APPLE, "APPLE"
    AssertEqualWithMessage 2, GREAT_SATAN_KOSAKA.PINEAPPLE, "PINEAPPLE"
    AssertEqualWithMessage 3, GREAT_SATAN_KOSAKA.PEN, "PEN"
End Sub
Sub Test_new_Enum_valueOf_Normal
    Dim def : Set def = createTestEnum
    
    Dim i
    For Each i In def.Keys
        AssertEqualWithMessage def.Item(i), GREAT_SATAN_KOSAKA.valueOf(i), "valueOf('" & i & "')"
    Next
End Sub
Sub Test_new_Enum_valueOf_Err
    Dim def : Set def = createTestEnum
    
    On Error Resume Next
    GREAT_SATAN_KOSAKA.valueOf("ORANGE")

    Dim e,a
    e = "clsTmp_" & "[A-Za-z0-9_]{8}" & "\(GREAT_SATAN_KOSAKA\)\+valueOf\(\)"
    a = Err.Source
    AssertMatchWithMessage e,a,"Source"

    e = "There is no element with the specified name"
    a = Err.Description
    AssertEqualWithMessage e,a,"Description"
End Sub
Sub Test_new_Enum_values
    Dim def : Set def = createTestEnum

    Dim i,ar
    ar = GREAT_SATAN_KOSAKA.values
    AssertEqualWithMessage def.Count-1, Ubound(ar), "count"
    For i=0 To Ubound(ar)
        AssertEqualWithMessage def.Item(def.Keys()(i)), ar(i), "values i="&i
    Next
End Sub
Sub Test_new_Enum_toString
    Dim def : Set def = createTestEnum

    Dim ar
    cf_push ar, "<clsTmp_" & "[A-Za-z0-9_]{8}" & ">\(GREAT_SATAN_KOSAKA\){"
    cf_push ar, "<ReadOnlyObject>{<String>'APPLE':<Integer>1}"
    cf_push ar, ",<ReadOnlyObject>{<String>'PINEAPPLE':<Integer>2}"
    cf_push ar, ",<ReadOnlyObject>{<String>'PEN':<Integer>3}"
    cf_push ar, "}"
    Dim e : e = Replace(Join(ar,""), "'", """")
    Dim a : a = GREAT_SATAN_KOSAKA.toString()

    AssertMatchWithMessage e, a, "toString()"
End Sub
Sub Test_new_Enum_equals
    Dim def : Set def = createTestEnum

    AssertEqualWithMessage False, GREAT_SATAN_KOSAKA.PINEAPPLE.equals(GREAT_SATAN_KOSAKA.APPLE), "PINEAPPLE=APPLE equals()=False"
    AssertEqualWithMessage True, GREAT_SATAN_KOSAKA.PINEAPPLE.equals(GREAT_SATAN_KOSAKA.PINEAPPLE), "PINEAPPLE=PINEAPPLE equals()=True"
    AssertEqualWithMessage False, GREAT_SATAN_KOSAKA.PINEAPPLE.equals(GREAT_SATAN_KOSAKA.PEN), "PINEAPPLE=PEN equals()=False"
    AssertEqualWithMessage False, GREAT_SATAN_KOSAKA.PINEAPPLE.equals(Nothing), "Nothing equals()=False"
    AssertEqualWithMessage False, GREAT_SATAN_KOSAKA.PINEAPPLE.equals(Empty), "Empty equals()=False"
End Sub
Sub Test_new_Enum_compareTo_Normal
    Dim def : Set def = createTestEnum

    AssertEqualWithMessage 1, GREAT_SATAN_KOSAKA.PINEAPPLE.compareTo(GREAT_SATAN_KOSAKA.APPLE), "PINEAPPLEvsAPPLE compareTo()=1"
    AssertEqualWithMessage 0, GREAT_SATAN_KOSAKA.PINEAPPLE.compareTo(GREAT_SATAN_KOSAKA.PINEAPPLE), "PINEAPPLEvsPINEAPPLE compareTo()=0"
    AssertEqualWithMessage -1, GREAT_SATAN_KOSAKA.PINEAPPLE.compareTo(GREAT_SATAN_KOSAKA.PEN), "PINEAPPLEvsPEN compareTo()=-1"
End Sub
Sub Test_new_Enum_compareTo_Err
    Dim def : Set def = createTestEnum

    On Error Resume Next
    GREAT_SATAN_KOSAKA.PINEAPPLE.compareTo(Nothing)

    Dim e,a
    e = "ReadOnlyObject+compareTo()"
    a = Err.Source
    AssertEqualWithMessage e,a,"Err.Source"

    e = "The type of the argument is different."
    a = Err.Description
    AssertEqualWithMessage e,a,"Err.Description"
End Sub

'###################################################################################################
'new_FileOf()
Sub Test_new_FileOf
    Dim p,e,a
    p = WScript.ScriptFullName
    Set e = CreateObject("Scripting.FileSystemObject").GetFile(p)
    Set a = new_FileOf(p)
    
    AssertEqual VarType(e), VarType(a)
    AssertEqual TypeName(e), TypeName(a)
End Sub
Sub Test_new_FileOf_Err
    On Error Resume Next
    Dim a : Set a = new_FileOf(vbNullString)
    
    AssertEqual 5, Err.Number
    AssertEqual "プロシージャの呼び出し、または引数が不正です。", Err.Description
    AssertEqual Empty, a
End Sub

'###################################################################################################
'new_FolderItem2Of()
Sub Test_new_FolderItem2Of
    Dim p,e,a
    p = WScript.ScriptFullName
    With CreateObject("Scripting.FileSystemObject")
        Set e = CreateObject("Shell.Application").Namespace(.GetParentFolderName(p)).Items().Item(.GetFileName(p))
    End With
    Set a = new_FolderItem2Of(p)
    
    AssertEqual VarType(e), VarType(a)
    AssertEqual TypeName(e), TypeName(a)
End Sub
Sub Test_new_FolderItem2Of_Err
    On Error Resume Next
    Dim a : Set a = new_FolderItem2Of(vbNullString)
    
    AssertEqual 424, Err.Number
    AssertEqual "オブジェクトがありません。", Err.Description
    AssertEqual Empty, a
End Sub

'###################################################################################################
'new_FolderOf()
Sub Test_new_FolderOf
    Dim p,e,a
    p = new_Fso().GetParentFolderName(WScript.ScriptFullName)
    Set e = CreateObject("Scripting.FileSystemObject").GetFolder(p)
    Set a = new_FolderOf(p)
    
    AssertEqual VarType(e), VarType(a)
    AssertEqual TypeName(e), TypeName(a)
End Sub
Sub Test_new_FolderOf_Err
    On Error Resume Next
    Dim a : Set a = new_FolderOf(vbNullString)
    
    AssertEqual 5, Err.Number
    AssertEqual "プロシージャの呼び出し、または引数が不正です。", Err.Description
    AssertEqual Empty, a
End Sub

'###################################################################################################
'new_Fso()
Sub Test_new_Fso
    Dim e : Set e = CreateObject("Scripting.FileSystemObject")
    Dim a : Set a = new_Fso()
    
    AssertEqual VarType(e), VarType(a)
    AssertEqual TypeName(e), TypeName(a)
End Sub

'###################################################################################################
'new_Func()
Sub Test_new_Func_Normal_1Line_0Return
    Dim code :  code = "function () {dim x}"
    Dim e : e = Empty
    Dim a : Set a = new_Func(code)
    
    AssertEqual e, a()
End Sub
Sub Test_new_Func_Normal_1Line_1Return
    Dim code :  code = "function (a){return 'ans='&a}"
    Dim d : d = 2
    Dim e : e = "ans="&d
    Dim a : Set a = new_Func(code)
    
    AssertEqual e, a(d)
End Sub
Sub Test_new_Func_Normal_nLine_0Return
    Dim code :  code = "function (a,b) {dim y:y= _:a+b:y=a* _:b}"
    Dim e : e = Empty
    Dim a : Set a = new_Func(code)
    
    AssertEqual e, a(3,6)
End Sub
Sub Test_new_Func_Normal_nLine_1Return
    Dim code :  code = "function (a,b)  {dim y:y= _:a+b:return y* _:b}"
    Dim e : e = 80
    Dim a : Set a = new_Func(code)
    
    AssertEqual e, a(2,8)
End Sub
Sub Test_new_Func_Normal_nLine_nReturn
    Dim code :  code = "function (a,b){ if a>b Then  :return b  :else:return a :  end if}"
    Dim a : Set a = new_Func(code)
    
    AssertEqual 2, a(2,3)
    AssertEqual 5, a(5,9)
End Sub
Sub Test_new_Func_Arrow_1Line_0Return
    Dim code :  code = "=> vbNullString"
    Dim e : e = vbNullString
    Dim a : Set a = new_Func(code)
    
    AssertEqual e, a()
End Sub
Sub Test_new_Func_Arrow_1Line_1Return
    Dim code :  code = "a=>  return _:  a^2"
    Dim e : e = 9^2
    Dim a : Set a = new_Func(code)
    
    AssertEqual e, a(9)
End Sub
Sub Test_new_Func_Arrow_nLine_0Return
    Dim code :  code = "(a,b)  =>{dim z:z=a^b}"
    Dim e : e = Empty
    Dim a : Set a = new_Func(code)
    
    AssertEqual e, a(1,2)
End Sub
Sub Test_new_Func_Arrow_nLine_1Return
    Dim code :  code = "(a,b) => {dim z:z=a^b:return z+1}"
    Dim e : e = 10
    Dim a : Set a = new_Func(code)
    
    AssertEqual e, a(3,2)
End Sub
Sub Test_new_Func_Arrow_nLine_nReturn
    Dim code :  code = "(a,b)=>{if a>b Then  :return b  :else:return a :  end if}"
    Dim a : Set a = new_Func(code)
    
    AssertEqual 2, a(2,3)
    AssertEqual 5, a(5,9)
End Sub

'###################################################################################################
'new_HtmlOf()
Sub Test_new_HtmlOf
    Dim e : Set e = New HtmlGenerator
    Dim a : Set a = new_HtmlOf("hoge")
    
    AssertEqual VarType(e), VarType(a)
    AssertEqual TypeName(e), TypeName(a)
End Sub
'Sub Test_new_HtmlOf_Err
'    On Error Resume Next
'    Dim a : Set a = new_HtmlOf("Ｈｏｇｅ")
'    
'    AssertEqual 1032, Err.Number
'    AssertEqual "要素（element）には半角以外の文字を指定できません。", Err.Description
'    AssertEqual Empty, a
'End Sub

'###################################################################################################
'new_Network()
Sub Test_new_Network
    Dim e : Set e = CreateObject("WScript.Network")
    Dim a : Set a = new_Network()
    
    AssertEqual VarType(e), VarType(a)
    AssertEqual TypeName(e), TypeName(a)
End Sub

'###################################################################################################
'new_Now()
Sub Test_new_Now
    Dim e : Set e = New Calendar
    Dim ed : ed = Now()
    Dim a : Set a = new_Now()
    
    AssertEqual VarType(e), VarType(a)
    AssertEqual TypeName(e), TypeName(a)
    AssertEqual Cstr(DatePart("yyyy", ed)), a.formatAs("YYYY")
    AssertEqual Cstr(DatePart("m", ed)), a.formatAs("M")
    AssertEqual Cstr(DatePart("d", ed)), a.formatAs("D")
End Sub

'###################################################################################################
'new_Re()
Sub Test_new_Re_Normal
    Dim e : Set e = New RegExp
    Dim ptn : ptn = "pattern"
    Dim a : Set a = new_Re(ptn, "b")
    
    AssertEqual VarType(e), VarType(a)
    AssertEqual TypeName(e), TypeName(a)
    AssertEqual ptn, a.Pattern
End Sub
Sub Test_new_Re_NoOpt
    Dim opt : opt = "abCde"
    Dim a : Set a = new_Re("a", opt)
    
    AssertEqual False, a.IgnoreCase
    AssertEqual False, a.Global
    AssertEqual False, a.Multiline
End Sub
Sub Test_new_Re_OptIgnoreCaseOnly
    Dim opt : opt = "xyzi"
    Dim a : Set a = new_Re("a", opt)
    
    AssertEqual True, a.IgnoreCase
    AssertEqual False, a.Global
    AssertEqual False, a.Multiline
End Sub
Sub Test_new_Re_OptGlobalOnly
    Dim opt : opt = "DEFGH"
    Dim a : Set a = new_Re("a", opt)
    
    AssertEqual False, a.IgnoreCase
    AssertEqual True, a.Global
    AssertEqual False, a.Multiline
End Sub
Sub Test_new_Re_OptMultilineOnly
    Dim opt : opt = "m"
    Dim a : Set a = new_Re("a", opt)
    
    AssertEqual False, a.IgnoreCase
    AssertEqual False, a.Global
    AssertEqual True, a.Multiline
End Sub
Sub Test_new_Re_OptFull
    Dim opt : opt = "aBcDeFgHIJkLMNoPqRsTuVwXyZ"
    Dim a : Set a = new_Re("a", opt)
    
    AssertEqual True, a.IgnoreCase
    AssertEqual True, a.Global
    AssertEqual True, a.Multiline
End Sub

'###################################################################################################
'new_Reader()
Sub Test_new_Reader
    Dim e : Set e = New BufferedReader
    Dim ts : Set ts =  new_Fso().OpenTextFile(WScript.ScriptFullName)
    Dim a : Set a = new_Reader(ts)
    
    AssertEqual VarType(e), VarType(a)
    AssertEqual TypeName(e), TypeName(a)
    AssertSame ts, a.textStream
End Sub

'###################################################################################################
'new_ReaderOf()
Sub Test_new_ReaderOf
    Dim e : Set e = New BufferedReader
    Dim a : Set a = new_ReaderOf(WScript.ScriptFullName)
    
    AssertEqual VarType(e), VarType(a)
    AssertEqual TypeName(e), TypeName(a)
End Sub
Sub Test_new_ReaderOf_Err
    On Error Resume Next
    Dim a : Set a = new_ReaderOf(vbNullString)
    
    AssertEqual 5, Err.Number
    AssertEqual "プロシージャの呼び出し、または引数が不正です。", Err.Description
    AssertEqual Empty, a
End Sub

'###################################################################################################
'new_Ret()
Sub Test_new_Ret
    Dim e : Set e = new ReturnValue
    Dim a : Set a = new_Ret(Empty)
    
    AssertEqualWithMessage VarType(e), VarType(a), "VarType"
    AssertEqualWithMessage TypeName(e), TypeName(a), "TypeName"
End Sub

'###################################################################################################
'new_RetByState()
Sub Test_new_RetByState
    Dim e : Set e = new ReturnValue
    Dim a : Set a = new_RetByState(Empty,Nothing)

    AssertEqualWithMessage VarType(e), VarType(a), "VarType"
    AssertEqualWithMessage TypeName(e), TypeName(a), "TypeName"
End Sub

'###################################################################################################
'new_Shell()
Sub Test_new_Shell
    Dim e : Set e = CreateObject("Wscript.Shell")
    Dim a : Set a = new_Shell()
    
    AssertEqual VarType(e), VarType(a)
    AssertEqual TypeName(e), TypeName(a)
End Sub

'###################################################################################################
'new_ShellApp()
Sub Test_new_ShellApp
    Dim e : Set e = CreateObject("Shell.Application")
    Dim a : Set a = new_ShellApp()
    
    AssertEqual VarType(e), VarType(a)
    AssertEqual TypeName(e), TypeName(a)
End Sub

'###################################################################################################
'new_Ts()
Sub Test_new_Ts
    Dim e : Set e = CreateObject("Scripting.FileSystemObject").OpenTextFile(WScript.ScriptFullName, 1, False, -2)
    Dim a : Set a = new_Ts(WScript.ScriptFullName, 1, False, -2)
    
    AssertEqual VarType(e), VarType(a)
    AssertEqual TypeName(e), TypeName(a)
End Sub
Sub Test_new_WriterOf_Err
    On Error Resume Next
    Dim a : Set a = new_Ts(vbNullString, 8, False, -2)
    
    AssertEqual 5, Err.Number
    AssertEqual "プロシージャの呼び出し、または引数が不正です。", Err.Description
    AssertEqual Empty, a
End Sub

'###################################################################################################
'new_Writer()
Sub Test_new_Writer
    Dim e : Set e = New BufferedWriter
    Dim ts : Set ts =  new_Fso().OpenTextFile(WScript.ScriptFullName)
    Dim a : Set a = new_Writer(ts)
    
    AssertEqual VarType(e), VarType(a)
    AssertEqual TypeName(e), TypeName(a)
    AssertSame ts, a.textStream
End Sub

'###################################################################################################
'new_WriterOf()
Sub Test_new_WriterOf
    Dim e : Set e = New BufferedWriter
    Dim a : Set a = new_WriterOf(WScript.ScriptFullName, 8, False, -2)
    
    AssertEqual VarType(e), VarType(a)
    AssertEqual TypeName(e), TypeName(a)
End Sub
Sub Test_new_WriterOf_Err
    On Error Resume Next
    Dim a : Set a = new_WriterOf(vbNullString, 8, False, -2)
    
    AssertEqual 5, Err.Number
    AssertEqual "プロシージャの呼び出し、または引数が不正です。", Err.Description
    AssertEqual Empty, a
End Sub


'###################################################################################################
'func_NewAnalyze()
Sub Test_func_NewAnalyze_1Line
    Dim code : code = "abc"
    Dim ev : ev = Array("abc")
    Dim a : a = func_NewAnalyze(code)
    
    AssertEqual Ubound(ev), Ubound(a)
    AssertEqual ev(0), a(0)
End Sub
Sub Test_func_NewAnalyze_1Line_UnderLine
    Dim code : code = " a_b c_d_ "
    Dim ev : ev = Array("a_b c_d_")
    Dim a : a = func_NewAnalyze(code)
    
    AssertEqual Ubound(ev), Ubound(a)
    AssertEqual ev(0), a(0)
End Sub
Sub Test_func_NewAnalyze_nLine
    Dim code : code = "a b:c_: d"
    Dim ev : ev = Array("a b","c_","d")
    Dim a : a = func_NewAnalyze(code)
    
    AssertEqual Ubound(ev), Ubound(a)
    AssertEqual ev(0), a(0)
    AssertEqual ev(1), a(1)
    AssertEqual ev(2), a(2)
End Sub
Sub Test_func_NewAnalyze_nLine_UnderLine
    Dim code : code = "a: b _:c d: e "
    Dim ev : ev = Array("a","b c d", "e")
    Dim a : a = func_NewAnalyze(code)
    
    AssertEqual Ubound(ev), Ubound(a)
    AssertEqual ev(0), a(0)
    AssertEqual ev(1), a(1)
    AssertEqual ev(2), a(2)
End Sub

'###################################################################################################
'func_NewRewriteReturnPhrase()
Sub Test_func_NewRewriteReturnPhrase_Normal_1Line_0Return
    Dim fn : fn = "fn_normal"
    Dim flg : flg = False
    Dim code : code = Array("abc")
    Dim e : e = "abc"
    Dim a : a = func_NewRewriteReturnPhrase(fn, flg, code)
    
    AssertEqual e, a
End Sub
Sub Test_func_NewRewriteReturnPhrase_Normal_1Line_1Return
    Dim fn : fn = "fn_normal"
    Dim flg : flg = False
    Dim code : code = Array("ab return c")
    Dim e : e = "ab  cf_bind fn_normal, (c)"
    Dim a : a = func_NewRewriteReturnPhrase(fn, flg, code)
    
    AssertEqual e, a
End Sub
Sub Test_func_NewRewriteReturnPhrase_Normal_nLine_0Return
    Dim fn : fn = "fn_normal"
    Dim flg : flg = False
    Dim code : code = Array("a bC", "dEf", "Gh i")
    Dim e : e = "a bC:dEf:Gh i"
    Dim a : a = func_NewRewriteReturnPhrase(fn, flg, code)
    
    AssertEqual e, a
End Sub
Sub Test_func_NewRewriteReturnPhrase_Normal_nLine_1Return
    Dim fn : fn = "fn_normal"
    Dim flg : flg = False
    Dim code : code = Array("aB c", "D ef", "g return h I")
    Dim e : e = "aB c:D ef:g  cf_bind fn_normal, (h I)"
    Dim a : a = func_NewRewriteReturnPhrase(fn, flg, code)
    
    AssertEqual e, a
End Sub
Sub Test_func_NewRewriteReturnPhrase_Normal_nLine_nReturn
    Dim fn : fn = "fn_normal"
    Dim flg : flg = False
    Dim code : code = Array("Abc", "d return eF", "return g H i")
    Dim e : e = "Abc:d  cf_bind fn_normal, (eF): cf_bind fn_normal, (g H i)"
    Dim a : a = func_NewRewriteReturnPhrase(fn, flg, code)
    
    AssertEqual e, a
End Sub
Sub Test_func_NewRewriteReturnPhrase_Arrow_1Line_0Return
    Dim fn : fn = "fn_arrow"
    Dim flg : flg = True
    Dim code : code = Array("abc")
    Dim e : e = "cf_bind fn_arrow, (abc)"
    Dim a : a = func_NewRewriteReturnPhrase(fn, flg, code)
    
    AssertEqual e, a
End Sub
Sub Test_func_NewRewriteReturnPhrase_Arrow_1Line_1Return
    Dim fn : fn = "fn_arrow"
    Dim flg : flg = True
    Dim code : code = Array("a B return c")
    Dim e : e = "a B  cf_bind fn_arrow, (c)"
    Dim a : a = func_NewRewriteReturnPhrase(fn, flg, code)
    
    AssertEqual e, a
End Sub
Sub Test_func_NewRewriteReturnPhrase_Arrow_nLine_0Return
    Dim fn : fn = "fn_arrow"
    Dim flg : flg = True
    Dim code : code = Array("a b  c", "DEF", "G h  I")
    Dim e : e = "a b  c:DEF:G h  I"
    Dim a : a = func_NewRewriteReturnPhrase(fn, flg, code)
    
    AssertEqual e, a
End Sub
Sub Test_func_NewRewriteReturnPhrase_Arrow_nLine_1Return
    Dim fn : fn = "fn_arrow"
    Dim flg : flg = True
    Dim code : code = Array("return a Bc", "De f", "g  h I")
    Dim e : e = " cf_bind fn_arrow, (a Bc):De f:g  h I"
    Dim a : a = func_NewRewriteReturnPhrase(fn, flg, code)
    
    AssertEqual e, a
End Sub
Sub Test_func_NewRewriteReturnPhrase_Arrow_nLine_nReturn
    Dim fn : fn = "fn_arrow"
    Dim flg : flg = True
    Dim code : code = Array("AB return c", "D return e f", "G   HI")
    Dim e : e = "AB  cf_bind fn_arrow, (c):D  cf_bind fn_arrow, (e f):G   HI"
    Dim a : a = func_NewRewriteReturnPhrase(fn, flg, code)
    
    AssertEqual e, a
End Sub



'###################################################################################################
'common
'For Test_new_Enum*()
Function createTestEnum
    Dim def : Set def = CreateObject("Scripting.Dictionary")
    With def
        .Add "APPLE", 1
        .Add "PINEAPPLE", 2
        .Add "PEN", 3
    End With
    new_Enum "GREAT_SATAN_KOSAKA", def
    Set createTestEnum = def
End Function


' Local Variables:
' mode: Visual-Basic
' indent-tabs-mode: nil
' End:
