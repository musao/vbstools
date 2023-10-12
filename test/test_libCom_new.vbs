' libCom.vbs: new_* procedure test.
' @import ../lib/clsCmArray.vbs
' @import ../lib/clsCmBufferedReader.vbs
' @import ../lib/clsCmBufferedWriter.vbs
' @import ../lib/clsCmCalendar.vbs
' @import ../lib/clsCmPubSub.vbs
' @import ../lib/clsCompareExcel.vbs
' @import ../lib/clsFsBase.vbs
' @import ../lib/libCom.vbs

Option Explicit

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
'new_DicWith()
Sub Test_new_DicWith_Normal
    Dim e : Set e = CreateObject("Scripting.Dictionary")
    Dim a : Set a = new_DicWith(Array(1,2,3))
    
    AssertEqual VarType(e), VarType(a)
    AssertEqual TypeName(e), TypeName(a)
End Sub
Sub Test_new_DicWith_EvenNumber
    Dim ev : ev = Array("first", "��", "Second", Nothing, "3rd", 3)
    Dim a : Set a = new_DicWith(ev)
    
    AssertEqual (Ubound(ev)+1)/2, a.Count
    AssertEqual ev(1), a.Item(ev(0))
    AssertSame ev(3), a.Item(ev(2))
    AssertEqual ev(5), a.Item(ev(4))
End Sub
Sub Test_new_DicWith_OddNumber
    Dim ev : ev = Array("first", "��", "Second", Nothing, "3rd")
    Dim a : Set a = new_DicWith(ev)
    
    AssertEqual Ubound(ev)/2+1, a.Count
    AssertEqual ev(1), a.Item(ev(0))
    AssertSame ev(3), a.Item(ev(2))
    AssertEqual Empty, a.Item(ev(4))
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
    Dim e : Set e = New clsCmBufferedReader
    Dim ts : Set ts =  CreateObject("Scripting.FileSystemObject").OpenTextFile(WScript.ScriptFullName)
    Dim a : Set a = new_Reader(ts)
    
    AssertEqual VarType(e), VarType(a)
    AssertEqual TypeName(e), TypeName(a)
    AssertSame ts, a.textStream
End Sub

'###################################################################################################
'new_ReaderFrom()
Sub Test_new_ReaderFrom
    Dim e : Set e = New clsCmBufferedReader
    Dim a : Set a = new_ReaderFrom(WScript.ScriptFullName)
    
    AssertEqual VarType(e), VarType(a)
    AssertEqual TypeName(e), TypeName(a)
End Sub

'###################################################################################################
'new_Writer()
Sub Test_new_Writer
    Dim e : Set e = New clsCmBufferedWriter
    Dim ts : Set ts =  CreateObject("Scripting.FileSystemObject").OpenTextFile(WScript.ScriptFullName)
    Dim a : Set a = new_Writer(ts)
    
    AssertEqual VarType(e), VarType(a)
    AssertEqual TypeName(e), TypeName(a)
    AssertSame ts, a.textStream
End Sub

'###################################################################################################
'new_WriterTo()
Sub Test_new_WriterTo
    Dim e : Set e = New clsCmBufferedWriter
    Dim a : Set a = new_WriterTo(WScript.ScriptFullName, 8, False, -2)
    
    AssertEqual VarType(e), VarType(a)
    AssertEqual TypeName(e), TypeName(a)
End Sub

'###################################################################################################
'new_Now()
Sub Test_new_Now
    Dim e : Set e = New clsCmCalendar
    Dim ed : ed = Now()
    Dim a : Set a = new_Now()
    
    AssertEqual VarType(e), VarType(a)
    AssertEqual TypeName(e), TypeName(a)
    AssertEqual Cstr(DatePart("yyyy", ed)), a.displayAs("YYYY")
    AssertEqual Cstr(DatePart("m", ed)), a.displayAs("M")
    AssertEqual Cstr(DatePart("d", ed)), a.displayAs("D")
End Sub

'###################################################################################################
'new_CalAt()
Sub Test_new_CalAt
    Dim e : Set e = New clsCmCalendar
    Dim ed : ed = CDate("2024/2/29")
    Dim a : Set a = new_CalAt(ed)
    
    AssertEqual VarType(e), VarType(a)
    AssertEqual TypeName(e), TypeName(a)
    AssertEqual Cstr(DatePart("yyyy", ed)), a.displayAs("YYYY")
    AssertEqual Cstr(DatePart("m", ed)), a.displayAs("M")
    AssertEqual Cstr(DatePart("d", ed)), a.displayAs("D")
End Sub

'###################################################################################################
'new_Pubsub()
Sub Test_new_Pubsub
    Dim e : Set e = New clsCmPubSub
    Dim a : Set a = new_Pubsub()
    
    AssertEqual VarType(e), VarType(a)
    AssertEqual TypeName(e), TypeName(a)
End Sub

'###################################################################################################
'new_Arr()
Sub Test_new_Arr
    Dim e : Set e = New clsCmArray
    Dim a : Set a = new_Arr()
    
    AssertEqual VarType(e), VarType(a)
    AssertEqual TypeName(e), TypeName(a)
    AssertEqual 0, a.Length
End Sub

'###################################################################################################
'new_ArrWith()
Sub Test_new_ArrWith
    Dim e : Set e = New clsCmArray
    Dim ev : ev = Array(1,Nothing,"�O")
    Dim a : Set a = new_ArrWith(ev)
    
    AssertEqual VarType(e), VarType(a)
    AssertEqual TypeName(e), TypeName(a)
    AssertEqual Ubound(ev)+1, a.Length
    AssertEqual ev(0), a(0)
    AssertSame ev(1), a(1)
    AssertEqual ev(2), a(2)
End Sub

'###################################################################################################
'new_ArrSplit()
Sub Test_new_ArrSplit
    Dim e : Set e = New clsCmArray
    Dim es : es = "one,��,3"
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
'new_Func()
Sub Test_new_Func_Normal_1Line_0Return
    Dim code :  code = "function () {dim x}"
    Dim e : e = Empty
    Dim a : Set a = new_Func(code)
    
    AssertEqual e, a()
End Sub
Sub Test_new_Func_Normal_1Line_1Return
    Dim code :  code = "function (a){return a}"
    Dim e : e = 2
    Dim a : Set a = new_Func(code)
    
    AssertEqual e, a(e)
End Sub
Sub Test_new_Func_Normal_nLine_0Return
    Dim code :  code = "function (a,b) {dim y:y= _:a+b:y=a* _:b}"
    Dim e : e = Empty
    Dim a : Set a = new_Func(code)
    
    AssertEqual e, a(3,6)
End Sub
Sub Test_new_Func_Normal_nLine_1Return
    Dim code :  code = "function (a,b) {dim y:y= _:a+b:return y* _:b}"
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
    Dim code :  code = "a=>return _:  a^2"
    Dim e : e = 9^2
    Dim a : Set a = new_Func(code)
    
    AssertEqual e, a(9)
End Sub
Sub Test_new_Func_Arrow_nLine_0Return
    Dim code :  code = "(a,b)=>{dim z:z=a^b}"
    Dim e : e = Empty
    Dim a : Set a = new_Func(code)
    
    AssertEqual e, a(1,2)
End Sub
Sub Test_new_Func_Arrow_nLine_1Return
    Dim code :  code = "(a,b)=>{dim z:z=a^b:return z+1}"
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
'func_FuncAnalyze()
Sub Test_func_FuncAnalyze_1Line
    Dim code : code = "abc"
    Dim ev : ev = Array("abc")
    Dim a : a = func_FuncAnalyze(code)
    
    AssertEqual Ubound(ev), Ubound(a)
    AssertEqual ev(0), a(0)
End Sub
Sub Test_func_FuncAnalyze_1Line_UnderLine
    Dim code : code = " a_b c_d_ "
    Dim ev : ev = Array("a_b c_d_")
    Dim a : a = func_FuncAnalyze(code)
    
    AssertEqual Ubound(ev), Ubound(a)
    AssertEqual ev(0), a(0)
End Sub
Sub Test_func_FuncAnalyze_nLine
    Dim code : code = "a b:c_: d"
    Dim ev : ev = Array("a b","c_","d")
    Dim a : a = func_FuncAnalyze(code)
    
    AssertEqual Ubound(ev), Ubound(a)
    AssertEqual ev(0), a(0)
    AssertEqual ev(1), a(1)
    AssertEqual ev(2), a(2)
End Sub
Sub Test_func_FuncAnalyze_nLine_UnderLine
    Dim code : code = "a: b _:c d: e "
    Dim ev : ev = Array("a","b c d", "e")
    Dim a : a = func_FuncAnalyze(code)
    
    AssertEqual Ubound(ev), Ubound(a)
    AssertEqual ev(0), a(0)
    AssertEqual ev(1), a(1)
    AssertEqual ev(2), a(2)
End Sub

'###################################################################################################
'func_FuncRewriteReturnPhrase()
Sub Test_func_FuncRewriteReturnPhrase_Normal_1Line_0Return
    Dim fn : fn = "fn_normal"
    Dim flg : flg = False
    Dim code : code = Array("abc")
    Dim e : e = "abc"
    Dim a : a = func_FuncRewriteReturnPhrase(fn, flg, code)
    
    AssertEqual e, a
End Sub
Sub Test_func_FuncRewriteReturnPhrase_Normal_1Line_1Return
    Dim fn : fn = "fn_normal"
    Dim flg : flg = False
    Dim code : code = Array("ab return c")
    Dim e : e = "ab  cf_bind fn_normal, (c)"
    Dim a : a = func_FuncRewriteReturnPhrase(fn, flg, code)
    
    AssertEqual e, a
End Sub
Sub Test_func_FuncRewriteReturnPhrase_Normal_nLine_0Return
    Dim fn : fn = "fn_normal"
    Dim flg : flg = False
    Dim code : code = Array("a bC", "dEf", "Gh i")
    Dim e : e = "a bC:dEf:Gh i"
    Dim a : a = func_FuncRewriteReturnPhrase(fn, flg, code)
    
    AssertEqual e, a
End Sub
Sub Test_func_FuncRewriteReturnPhrase_Normal_nLine_1Return
    Dim fn : fn = "fn_normal"
    Dim flg : flg = False
    Dim code : code = Array("aB c", "D ef", "g return h I")
    Dim e : e = "aB c:D ef:g  cf_bind fn_normal, (h I)"
    Dim a : a = func_FuncRewriteReturnPhrase(fn, flg, code)
    
    AssertEqual e, a
End Sub
Sub Test_func_FuncRewriteReturnPhrase_Normal_nLine_nReturn
    Dim fn : fn = "fn_normal"
    Dim flg : flg = False
    Dim code : code = Array("Abc", "d return eF", "return g H i")
    Dim e : e = "Abc:d  cf_bind fn_normal, (eF): cf_bind fn_normal, (g H i)"
    Dim a : a = func_FuncRewriteReturnPhrase(fn, flg, code)
    
    AssertEqual e, a
End Sub
Sub Test_func_FuncRewriteReturnPhrase_Arrow_1Line_0Return
    Dim fn : fn = "fn_arrow"
    Dim flg : flg = True
    Dim code : code = Array("abc")
    Dim e : e = "cf_bind fn_arrow, (abc)"
    Dim a : a = func_FuncRewriteReturnPhrase(fn, flg, code)
    
    AssertEqual e, a
End Sub
Sub Test_func_FuncRewriteReturnPhrase_Arrow_1Line_1Return
    Dim fn : fn = "fn_arrow"
    Dim flg : flg = True
    Dim code : code = Array("a B return c")
    Dim e : e = "a B  cf_bind fn_arrow, (c)"
    Dim a : a = func_FuncRewriteReturnPhrase(fn, flg, code)
    
    AssertEqual e, a
End Sub
Sub Test_func_FuncRewriteReturnPhrase_Arrow_nLine_0Return
    Dim fn : fn = "fn_arrow"
    Dim flg : flg = True
    Dim code : code = Array("a b  c", "DEF", "G h  I")
    Dim e : e = "a b  c:DEF:G h  I"
    Dim a : a = func_FuncRewriteReturnPhrase(fn, flg, code)
    
    AssertEqual e, a
End Sub
Sub Test_func_FuncRewriteReturnPhrase_Arrow_nLine_1Return
    Dim fn : fn = "fn_arrow"
    Dim flg : flg = True
    Dim code : code = Array("return a Bc", "De f", "g  h I")
    Dim e : e = " cf_bind fn_arrow, (a Bc):De f:g  h I"
    Dim a : a = func_FuncRewriteReturnPhrase(fn, flg, code)
    
    AssertEqual e, a
End Sub
Sub Test_func_FuncRewriteReturnPhrase_Arrow_nLine_nReturn
    Dim fn : fn = "fn_arrow"
    Dim flg : flg = True
    Dim code : code = Array("AB return c", "D return e f", "G   HI")
    Dim e : e = "AB  cf_bind fn_arrow, (c):D  cf_bind fn_arrow, (e f):G   HI"
    Dim a : a = func_FuncRewriteReturnPhrase(fn, flg, code)
    
    AssertEqual e, a
End Sub



' Local Variables:
' mode: Visual-Basic
' indent-tabs-mode: nil
' End: