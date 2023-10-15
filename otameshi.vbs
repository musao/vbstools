Option Explicit


'定数
Private Const Cs_FOLDER_LIB = "lib"
Private Const Cs_FOLDER_TEMP = "tmp"

'import定義
Sub sub_import( _
    byVal asIncludeFileName _
    )
    With CreateObject("Scripting.FileSystemObject")
        Dim sParentFolderName : sParentFolderName = .GetParentFolderName(WScript.ScriptFullName)
        Dim sIncludeFilePath
        sIncludeFilePath = .BuildPath(sParentFolderName, Cs_FOLDER_LIB)
        sIncludeFilePath = .BuildPath(sIncludeFilePath, asIncludeFileName)
        ExecuteGlobal .OpenTextfile(sIncludeFilePath).ReadAll
    End With
End Sub
'import
Call sub_import("clsCmArray.vbs")
Call sub_import("clsCmBufferedWriter.vbs")
Call sub_import("clsCmCalendar.vbs")
Call sub_import("clsCmPubSub.vbs")
Call sub_import("clsCompareExcel.vbs")
Call sub_import("libCom.vbs")


'Dim a
'inputbox "", "", "vartype(a) = " & vartype(a) & vbnewline & "typename(a) = " & typename(a) & vbnewline & "isarray(a) = " & isarray(a) & vbnewline & "isempty(a) = " & isempty(a) & vbnewline & "isobject(a) = " & isobject(a)
''vartype(a) = 0 typename(a) = Empty isarray(a) = False isempty(a) = True isobject(a) = False
'
'Dim b()
'inputbox "", "", "vartype(b) = " & vartype(b) & vbnewline & "typename(b) = " & typename(b) & vbnewline & "isarray(b) = " & isarray(b) & vbnewline & "isempty(b) = " & isempty(b) & vbnewline & "isobject(b) = " & isobject(b)
''vartype(b) = 8204 typename(b) = Variant() isarray(b) = True isempty(b) = False isobject(b) = False
'
'a = array(1,2,3)
'inputbox "", "", "Vartype(a) = " & vartype(a) & vbnewline & "typename(a) = " & typename(a) & vbnewline & "ubound(a) = " & Ubound(a) & vbnewline & "isarray(a) = " & isarray(a) & vbnewline & "isempty(a) = " & isempty(a) & vbnewline & "isobject(a) = " & isobject(a)
''Vartype(a) = 8204 typename(a) = Variant() ubound(a) = 2 isarray(a) = True isempty(a) = False isobject(a) = False

''Test func_CM_UtilGetIpaddress
'inputbox "", "", func_CM_ToString(func_CM_UtilGetIpaddress())                   '{"[00000016] Hyper-V Virtual Ethernet Adapter"=>{"v4"=>"172.23.0.1","v6"=>"fe80::b763:3fce:cdd9:c0d3"},"[00000021] Hyper-V Virtual Ethernet Adapter"=>{"v4"=>"192.168.11.52","v6"=>"fe80::ba87:1e93:59ab:28f7"}}
'dim s : Set s = new_Func("a=>dim x,i:set x=new_dic():for each i in a.keys:if left(a.item(i).item(""v4""), 3)<>""172"" then:x.add i, a.item(i):end if:next:return x")(func_CM_UtilGetIpaddress())
'inputbox "", "", func_CM_ToString(s)                                            '{"[00000021] Hyper-V Virtual Ethernet Adapter"=>{"v4"=>"192.168.11.52","v6"=>"fe80::ba87:1e93:59ab:28f7"}}
'                                                                                '{"[00000021] Hyper-V Virtual Ethernet Adapter"=>{"v4"=>"192.168.11.52","v6"=>"fe80::ba87:1e93:59ab:28f7"}}
'inputbox "", "", func_CM_ToString( new_Func("a=>dim x,i:set x=new_dic():for each i in a.keys:if left(a.item(i).item(""v4""), 3)<>""172"" then:x.add i, a.item(i):end if:next:return x")(func_CM_UtilGetIpaddress()).Items()(0) )

''Test cf_tryCatch()
'Dim oFuncTry, oArguments, oFuncCatch, oFuncFinary, oReturn
'
''normal
'Set oFuncTry = new_Func("a=>msgbox(""ok"")")
'Call cf_tryCatch(oFuncTry, oArguments, oFuncCatch, oFuncFinary)           'ok
'inputbox "","",func_CM_ToStringErr()                                           '<Err> {"Number"=>0,"Description"=>"","Source"=>""}
'                                                                               '
'
''normal2
'Set oFuncTry = new_Func("a=>a(0)+a(1)")
'oArguments = Array(1,2)
'Set oReturn = cf_tryCatch(oFuncTry, oArguments, oFuncCatch, oFuncFinary)
'inputbox "","",func_CM_ToString(oReturn)                                       '{"Result"=>True,"Return"=>3,"Err"=><Nothing>}
'                                                                               '
'inputbox "","",func_CM_ToStringErr()                                           '<Err> {"Number"=>0,"Description"=>"","Source"=>""}
'                                                                               '
''normal3
'Set oFuncTry = new_Func("a=>a(0)+a(1)")
'oArguments = Array(1,2)
'Set oFuncFinary = new_Func("a=>""anser is ""&a")
'Set oReturn = cf_tryCatch(oFuncTry, oArguments, oFuncCatch, oFuncFinary)
'inputbox "","",func_CM_ToString(oReturn)                                       '{"Result"=>True,"Return"=>"anser is 3","Err"=><Nothing>}
'                                                                               '
'inputbox "","",func_CM_ToStringErr()                                           '<Err> {"Number"=>0,"Description"=>"","Source"=>""}
'                                                                               '
'
''err
'Set oFuncTry = new_Func("a=>a(0)/a(1)")
'oFuncFinary = empty
'oArguments = Array(1,0)
'Set oReturn = cf_tryCatch(oFuncTry, oArguments, oFuncCatch, oFuncFinary)
'inputbox "","",func_CM_ToString(oReturn)                                       '{"Result"=>False,"Return"=><empty>,"Err"=>{"Number"=>11,"Description"=>"0 で除算しました。","Source"=>"Microsoft VBScript 実行時エラー"}}
'                                                                               '
'inputbox "","",func_CM_ToStringErr()                                           '<Err> {"Number"=>0,"Description"=>"","Source"=>""}
'                                                                               '
'
''err2
'Set oFuncTry = new_Func("a=>a(0)/a(1)")
'oArguments = Array(1,0)
'Set oFuncCatch = new_Func("(a,e)=>a(0)+a(1)")
'Set oReturn = cf_tryCatch(oFuncTry, oArguments, oFuncCatch, oFuncFinary)
'inputbox "","",func_CM_ToString(oReturn)                                       '{"Result"=>False,"Return"=>1,"Err"=>{"Number"=>11,"Description"=>"0 で除算しました。","Source"=>"Microsoft VBScript 実行時エラー"}}
'                                                                               '
'inputbox "","",func_CM_ToStringErr()                                           '<Err> {"Number"=>0,"Description"=>"","Source"=>""}
'                                                                               '
'
''err3
'Set oFuncTry = new_Func("a=>a(0)/a(1)")
'oArguments = Array(1,0)
'Set oFuncCatch = new_Func("(a,e)=>a(0)+a(1)")
'Set oFuncFinary = new_Func("a=>""anser is ""&a")
'Set oReturn = cf_tryCatch(oFuncTry, oArguments, oFuncCatch, oFuncFinary)
'inputbox "","",func_CM_ToString(oReturn)                                       '{"Result"=>False,"Return"=>"anser is 1","Err"=>{"Number"=>11,"Description"=>"0 で除算しました。","Source"=>"Microsoft VBScript 実行時エラー"}}
'                                                                               '
'inputbox "","",func_CM_ToStringErr()                                           '<Err> {"Number"=>0,"Description"=>"","Source"=>""}
'                                                                               '
'
'wscript.quit
'

''Test new_Func()
'Dim sSoruceCode
'sSoruceCode = "function(a, b){ return (a > b) }"
'Call Msgbox(new_Func(sSoruceCode)(1,1))   'False
'Call Msgbox(new_Func(sSoruceCode)(2,1))   'True
'
'sSoruceCode = "function(a, b){ Dim c }"
'Call Msgbox(new_Func(sSoruceCode)(9,8))   '空
'
'sSoruceCode = "function(){ return ""OK"" }"
'Call Msgbox(new_Func(sSoruceCode)())      'OK
'
'sSoruceCode = "function (a, b) { Dim c" & vbNewLine & _
'                         "c = a + b" & vbNewLine & _
'                         "return c }"
'Call Msgbox(new_Func(sSoruceCode)(5,6))   '11
'
'sSoruceCode = "function(a, b){}"
'Call Msgbox(new_Func(sSoruceCode)(-4,0))  '空
'
'sSoruceCode = "a => (a + a)"
'Call Msgbox(new_Func(sSoruceCode)(-8)  )  '-16
'
'sSoruceCode = "(a, b) => b"
'Call Msgbox(new_Func(sSoruceCode)(5,6))   '6
'
'sSoruceCode = "(a, b) => { Dim c" & vbNewLine & _
'                         "c = a + b" & vbNewLine & _
'                         "return c }"
'Call Msgbox(new_Func(sSoruceCode)(7,3))   '10
'
'sSoruceCode = "a => a^2"
'Call Msgbox(new_Func(sSoruceCode)(9))     '81
'
'wscript.quit

''Test func_CM_MathRound()
'Dim dbPlus0, dbPlus1, dbPlus5 ,dbMinas0 ,dbMinas2 ,dbMinas5
'dbPlus0=14.555555
'dbPlus1=14.456789
'dbPlus5=14.432154
'dbMinas0=-14.555555
'dbMinas2=-14.501234
'dbMinas5=-14.432154
'call MsgBox( func_CM_MathRound(dbPlus5, 0, 5) )      '14.4321
'call MsgBox( func_CM_MathRound(dbPlus0, 5, 0) )      '10
'call MsgBox( func_CM_MathRound(dbPlus1, 5, 1) )      '14
'call MsgBox( func_CM_MathRound(dbPlus0, 9, 0) )      '20
'call MsgBox( func_CM_MathRound(dbPlus1, 9, 1) )      '15
'call MsgBox( func_CM_MathRound(dbPlus5, 9, 5) )      '14.4322
'call MsgBox( func_CM_MathRound(dbMinas5, 0, 5) )      '-14.4322
'call MsgBox( func_CM_MathRound(dbMinas0, 5, 0) )      '-10
'call MsgBox( func_CM_MathRound(dbMinas2, 5, 2) )      '-14.5
'call MsgBox( func_CM_MathRound(dbMinas0, 9, 0) )      '-10
'call MsgBox( func_CM_MathRound(dbMinas2, 9, 2) )      '-14.5
'call MsgBox( func_CM_MathRound(dbMinas5, 9, 5) )      '-14.4321
'
'wscript.quit

'Test func_CM_UtilGenerateRandomString
Call msgbox( func_CM_UtilGenerateRandomString(50, 15, Nothing) )        '大小数記
Call msgbox( func_CM_UtilGenerateRandomString(50, 8, Nothing)  )        '　　　記
Call msgbox( func_CM_UtilGenerateRandomString(50, 7, Nothing)  )        '大小数
Call msgbox( func_CM_UtilGenerateRandomString(50, 4, Nothing)  )        '　　数
Call msgbox( func_CM_UtilGenerateRandomString(50, 3, Nothing)  )        '大小
Call msgbox( func_CM_UtilGenerateRandomString(50, 2, Nothing)  )        '　小
Call msgbox( func_CM_UtilGenerateRandomString(50, 1, Nothing)  )        '大
Call msgbox( func_CM_UtilGenerateRandomString(50, 4, Nothing)  )        '　　数　
Call msgbox( func_CM_UtilGenerateRandomString(50, 4, Array("0", "9") ) )  '　　数　
Call msgbox( func_CM_UtilGenerateRandomString(50, 4, Array("a", "Z") ) )  '　　数　＋"a","Z"
Call msgbox( func_CM_UtilGenerateRandomString(50, 4, Array("\", "$") ) )  '　　数　＋"\","$"


wscript.quit

'Test func_CM_UtilSort～()
'Dim vArray : vArray = Array(5,2,9,6,4,8,7,3,0,1)
'Dim vArray : vArray = Array("C","$","b","漢","a","B","あ","A","c","0")
'inputbox "","",func_CM_ToString( func_CM_UtilSortHeap(vArray, new_Func("(c,n)=>c>n"), True) )
'inputbox "","",func_CM_ToString( func_CM_UtilSortHeap(vArray, new_Func("(c,n)=>c>n"), False) )
'wscript.quit
'Call msgbox( func_CM_ToString(vArray) )  '[5,2,9,6,4,8,7,3,0,1]
'private function SortTest(x,y)
'    SortTest = (x > y)
'end function
''Test func_CM_UtilSortBubble()
'Call msgbox( func_CM_ToString( func_CM_UtilSortBubble(vArray, getref("SortTest"), True) ) )  '[0,1,2,3,4,5,6,7,8,9]
'Call msgbox( func_CM_ToString( func_CM_UtilSortBubble(vArray, getref("SortTest"), False) ) ) '[9,8,7,6,5,4,3,2,1,0]

''Test func_CM_UtilSortBubble()
'Call msgbox( func_CM_ToString( func_CM_UtilSortQuick(vArray, getref("SortTest"), True) ) )  '[0,1,2,3,4,5,6,7,8,9]
'Call msgbox( func_CM_ToString( func_CM_UtilSortQuick(vArray, getref("SortTest"), False) ) ) '[9,8,7,6,5,4,3,2,1,0]

''Test func_CM_UtilSortMerge()
'Call msgbox( func_CM_ToString( func_CM_UtilSortMerge(vArray, getref("SortTest"), True) ) )  '[0,1,2,3,4,5,6,7,8,9]
'Call msgbox( func_CM_ToString( func_CM_UtilSortMerge(vArray, getref("SortTest"), False) ) ) '[9,8,7,6,5,4,3,2,1,0]

''Test func_CM_UtilSortHeap()
'Call msgbox( func_CM_ToString( func_CM_UtilSortHeap(vArray, getref("SortTest"), True) ) )  '[0,1,2,3,4,5,6,7,8,9]
'Call msgbox( func_CM_ToString( func_CM_UtilSortHeap(vArray, getref("SortTest"), False) ) ) '[9,8,7,6,5,4,3,2,1,0]
'
'wscript.quit

''Test func_CM_ArrayIsAvailable()
'Dim vArrayTest
'Call Msgbox("func_CM_ArrayIsAvailable(vArrayTest) = " & func_CM_ArrayIsAvailable(vArrayTest)) 'False
'Dim vArrayTest2()
'Call Msgbox("func_CM_ArrayIsAvailable(vArrayTest) = " & func_CM_ArrayIsAvailable(vArrayTest2)) 'False
'Redim vArrayTest2(0)
'Call Msgbox("func_CM_ArrayIsAvailable(vArrayTest) = " & func_CM_ArrayIsAvailable(vArrayTest2)) 'True
'Redim vArrayTest2(1)
'Call Msgbox("func_CM_ArrayIsAvailable(vArrayTest) = " & func_CM_ArrayIsAvailable(vArrayTest2)) 'True
'
'wscript.quit


dim arr5

''Test Concat()
'Set arr5 = new_ArrWith(Array(1,2,3,4,5,6))
'Call msgbox(func_CM_ToString(arr5))
'Call msgbox(func_CM_ToString(arr5.Concat(Array("a",9))))

''Test Every(),Some()
'private function EveryTestOk(arg, i, a)
'    EveryTestOk = (arg < 5)
'end function
'private function EveryTestNg(arg, i, a)
'    EveryTestNg = (arg < 3)
'end function
'private function EveryTestNg2(arg, i, a)
'    EveryTestNg2 = (arg < 0)
'end function
'Set arr5 = new_ArrWith(Array(1,2,3))
'Call msgbox(func_CM_ToString(arr5))
'Call msgbox( arr5.Every(getref("EveryTestOk")) )     'True
'Call msgbox( arr5.Every(getref("EveryTestNg")) )     'False
'Call msgbox( arr5.Every(getref("EveryTestNg2")) )    'False
'private function SomeTestNg(arg, i, a)
'    SomeTestNg = (arg > 5)
'end function
'private function SomeTestOk(arg, i, a)
'    SomeTestOk = (arg > 2)
'end function
'private function SomeTestNg2(arg, i, a)
'    SomeTestNg2 = True
'end function
'Call msgbox( arr5.Some(getref("SomeTestNg")) )       'False
'Call msgbox( arr5.Some(getref("SomeTestOk")) )       'True
'Set arr5 = new_Arr()
'Call msgbox( arr5.Some(getref("SomeTestNg2")) )      'False


''Test Filter()
'Set arr5 = new_ArrWith(Array(1,2,3))
'Call msgbox(func_CM_ToString(arr5))                                       '[1,2,3]
'Call msgbox( func_CM_ToString(arr5.Filter(new_Func("(e,i,a)=>(e>1)"))) )  '[2,3]

''Test ForEach()
'private function ForEachTest(arg, i, a)
'    Call msgbox(func_CM_ToString(arg))
'    Call msgbox(func_CM_ToString(i))
'    Call msgbox(func_CM_ToString(a))
'end function
'Set arr5 = new_ArrWith(Array(8, "Z"))
'Call msgbox(func_CM_ToString(arr5))
'arr5.ForEach getref("ForEachTest")

''Test IndexOf()
'Dim IndexOfTest : Set IndexOfTest = new_DicWith(Array(4, "five"))
'Set arr5 = new_Arr()
'Call msgbox( arr5.IndexOf("a") )          '-1
'Set arr5 = Nothing
'Set arr5 = new_ArrWith(Array("a", 2, 3.14, IndexOfTest, "End"))
'Call msgbox(func_CM_ToString(arr5))
'Call msgbox( arr5.IndexOf("a") )          '0
'Call msgbox( arr5.IndexOf(IndexOfTest) )  '3
'Call msgbox( arr5.IndexOf("Start") )      '-1
'Call msgbox( arr5.IndexOf(2) )            '1
'Call msgbox( arr5.IndexOf("2") )          '-1

''Test joinvbs()
'Set arr5 = new_ArrWith(Array(1, 2, 3.14, "Testing"))
'Call msgbox(func_CM_ToString(arr5))         '[1,2,3.14,"Testing"]
'Call msgbox( arr5.joinvbs("") )             '"123.14Testing"
'Call msgbox( arr5.joinvbs("+") )            '"1+2+3.14+Testing"
'Call msgbox( arr5.joinvbs("") )             '"123.14Testing"
'Call msgbox( arr5.Joinvbs("+") )            '"1+2+3.14+Testing"

''Test LastIndexOf()
'Dim LastIndexOfTest : Set LastIndexOfTest = new_DicWith(Array(4, "five"))
'Set arr5 = new_Arr()
'Call msgbox( arr5.LastIndexOf(LastIndexOfTest) )  '-1
'Set arr5 = Nothing
'Set arr5 = new_ArrWith(Array("a", 2, 3.14, LastIndexOfTest, "End"))
'Call msgbox(func_CM_ToString(arr5))
'Call msgbox( arr5.LastIndexOf("a") )          '0
'Call msgbox( arr5.LastIndexOf(LastIndexOfTest) )  '3
'Call msgbox( arr5.LastIndexOf("Start") )      '-1
'Call msgbox( arr5.LastIndexOf(2) )            '1
'Call msgbox( arr5.LastIndexOf("2") )          '-1

''Test Length(),Push(),Pop(),Shift(),Unshift()
'Set arr5 = new_Arr()
'Call msgbox( func_CM_ToString(arr5) & vbNewLine & arr5.Length )  '<clsCmArray> 0
'Set arr5 = Nothing
'Set arr5 = new_ArrWith(Array("1", 2))
'Call msgbox( func_CM_ToString(arr5) & vbNewLine & arr5.Length )  '["1",2] 2
'arr5.Concat Array(3, "Four")
'Call msgbox( func_CM_ToString(arr5) & vbNewLine & arr5.Length )  '["1",2] 2
'arr5.Push Array("th", "ree")
'Call msgbox( func_CM_ToString(arr5) & vbNewLine & arr5.Length )  '["1",2,["th","ree"]] 3
'arr5.Unshift new_DicWith(Array(4, "四"))
'Call msgbox( func_CM_ToString(arr5) & vbNewLine & arr5.Length )  '[{4=>"四"},"1",2,["th","ree"]] 4
'Call msgbox( func_CM_ToString(arr5.Pop) )                        '["th","ree"]
'Call msgbox( func_CM_ToString(arr5) & vbNewLine & arr5.Length )  '[{4=>"四"},"1",2] 3
'Call msgbox( func_CM_ToString(arr5.Shift) )                      '{4=>"四"}
'Call msgbox( func_CM_ToString(arr5) & vbNewLine & arr5.Length )  '["1",2] 2

''Test Map()
'private function MapTest(arg, i, a)
'    MapTest = arg*arg
'end function
'Set arr5 = new_ArrWith(Array(1,2,3))
'Call msgbox( func_CM_ToString(arr5) )
'Call msgbox( func_CM_ToString(arr5.Map(getref("MapTest"))) )

''Test Reduce()
'private function ReduceTest(prev, current, i, a)
'    ReduceTest = prev*current
'end function
'Set arr5 = new_ArrWith(Array(1,2,3,4))
'Call msgbox( func_CM_ToString(arr5) )
'Call msgbox( arr5.Reduce(getref("ReduceTest")) )

''Test ReduceRight()
'private function ReduceRightTest(prev, current, i, a)
'    ReduceRightTest = prev/current
'end function
'Set arr5 = new_ArrWith(Array(2,10,60))
'Call msgbox( func_CM_ToString(arr5) )
'Call msgbox( arr5.ReduceRight(getref("ReduceRightTest")) )

''Test Reverse()
'Set arr5 = new_ArrWith(Array(1,2,3))
'Call msgbox( func_CM_ToString(arr5) )                  '[1.2.3]
'arr5.Reverse
'Call msgbox( func_CM_ToString(arr5) )                  '[3,2,1]

''Test Slice()
'Set arr5 = new_ArrWith(Array(1,2,3,4,5))
'Call msgbox( func_CM_ToString(arr5) )
'Call msgbox( func_CM_ToString(arr5.Slice(0,3)) )               '[1.2.3]
'Call msgbox( func_CM_ToString(arr5.Slice(3, vbNullString)) )   '[4,5]
'Call msgbox( func_CM_ToString(arr5.Slice(1, -1)) )             '[2,3,4]
'Call msgbox( func_CM_ToString(arr5.Slice(-3, -2)) )            '[3]
'Call msgbox( func_CM_ToString(arr5.Slice(-3, -3)) )            '<clsCmArray>
'Set arr5 = new_ArrWith(Array(1))
'Call msgbox( func_CM_ToString(arr5) )
'Call msgbox( func_CM_ToString(arr5.Slice(0,2)) )               '[1]


''Test sort()
'Set arr5 = new_ArrWith(Array(5,2,9,6,4,8,7,3,0,1))
'Call msgbox( func_CM_ToString(arr5) )
'Call msgbox( func_CM_ToString(arr5.sort(True)) )
'Call msgbox( func_CM_ToString(arr5.sort(False)) )

''Test sortUsing()
'private function ArraySortTest(x,y)
'    ArraySortTest = (x > y)
'end function
'Set arr5 = new_ArrWith(Array(5,2,9,6,4,8,7,3,0,1))
'Call msgbox( func_CM_ToString(arr5) )
'Call msgbox( func_CM_ToString(arr5.sortUsing(getref("ArraySortTest"))) )
'Call msgbox( func_CM_ToString(arr5.sortUsing(new_Func("(x,y) => (x>y)"))) )

''Test Splice()
'Set arr5 = new_ArrWith(Array(1,2,3,4,5,6,7,8))
'Call msgbox( func_CM_ToString(arr5) )                          '[1,2,3,4,5,6,7,8]
'Call msgbox( func_CM_ToString(arr5.splice(1,2,Nothing)) )      '[2,3]
'Call msgbox( func_CM_ToString(arr5) )                          '[1,4,5,6,7,8]
'Call msgbox( func_CM_ToString(arr5.splice(1,1,Nothing)) )      '[4]
'Call msgbox( func_CM_ToString(arr5) )                          '[1,5,6,7,8]
'Call msgbox( func_CM_ToString(arr5.splice(1,0,Array(2,3))) )   '[]
'Call msgbox( func_CM_ToString(arr5) )                          '[1,2,3,5,6,7,8]

wscript.quit


'
'
'Call Msgbox(5 \ 3)
'
'wscript.quit
'
'private function dummy()
'    Dim cont
'    cont = "function test(arg):test = false:if arg mod 2 = 0 Then:test = true:end if:end function"
'    ExecuteGlobal(cont)
''    execute(cont)
'    Call msgbox(test(1))
'    set dummy = getref("test")
''    set dummy = getref("func_CM_ToString")
'end function
'
'Call Msgbox(dummy()(2))
'
'
'wscript.quit


private function test(arg, i, a)
    test = false
    if arg mod 2 = 0 Then test = true
end function


'dim arr2 : Set arr2 = new_Arr()
dim arr2 : Set arr2 = new_ArrWith(Array(1,2,3,4,5,6))
Call msgbox(func_CM_ToString(arr2.items))

'Call Msgbox(arr2.Length)
'Call Msgbox(arr2(2))
'arr2(2) = 10
'Set arr2(5) = new_Arr()
'Call Msgbox(arr2(2))

dim arr3
Set arr3 = arr2.filter(getref("test"))

'Call msgbox(func_CM_ToString(arr3.items))
Call Msgbox(arr3.Length)
'Call msgbox(func_CM_ToString(arr3.items))
Call Msgbox(arr3(2))

Call msgbox(func_CM_ToString(arr3.items))

Call Msgbox(arr2.find(getref("test")))

Call msgbox(func_CM_ToString(arr2.items))


'dim ele
'for each ele in arr3.items
'    Call Msgbox(ele)
'next
'
'Call msgbox(arr3.joinvbs("-"))
'
'Call msgbox(func_CM_ToString(arr3) & vbnewline & func_CM_ToString(arr3.items))

wscript.quit


dim arr : Set arr = New clsCmArray

Call Msgbox(arr.Length)

arr.push "あ"

Call Msgbox(arr.Length)
Call Msgbox(arr(0))

arr.PushMulti(array(1,"hello", #2023/9/10#))

Call Msgbox(arr.Length)
Call Msgbox(arr(3))

arr.pop

Call Msgbox(arr.Length)
Call Msgbox(arr(2))

arr.Unshift "か"

Call Msgbox(arr.Length)
Call Msgbox(arr(3))

arr.UnshiftMulti(array(9,"world", #1999/9/10#))

Call Msgbox(arr.Length)
Call Msgbox(arr(6))

arr.Shift

Call Msgbox(arr.Length)
Call Msgbox(arr(5))

wscript.quit


Call Msgbox(func_CM_ToString(1))
Call Msgbox(func_CM_ToString("Hello world."))
Call Msgbox(func_CM_ToString(#2009-03-07#))
Call Msgbox(func_CM_ToString(Array("foo", "bar", "baz")))

Dim oD : Set oD = new_Dic()
Call cf_bindAt(oD, "foo", 1)
Call cf_bindAt(oD, "bar", Nothing)
Call cf_bindAt(oD, "baz", Empty)
Call Msgbox(func_CM_ToString(oD))

Call Msgbox(func_CM_ToString(new_Re("foo", "i")))

wscript.quit


Dim sPatha
sPatha = func_CM_FsGetPrivateLogFilePath()
Dim bw
Set bw = new_Writer(func_CM_FsOpenTextFile(sPatha, 8, True, -2))

With bw
    .WriteBufferSize = 2
    Call msgbox("WriteBufferSize()='" & .WriteBufferSize() & "'" & vbNewLine _
                & "WriteIntervalTime()='" & .WriteIntervalTime() & "'" & vbNewLine _
                & "CurrentBufferSize()='" & .CurrentBufferSize() & "'" & vbNewLine _
                & "LastWriteTime()='" & .LastWriteTime() &"'" _
                )
    .WriteContents("あ")
    .newLine()
    Call msgbox("WriteBufferSize()='" & .WriteBufferSize() & "'" & vbNewLine _
                & "WriteIntervalTime()='" & .WriteIntervalTime() & "'" & vbNewLine _
                & "CurrentBufferSize()='" & .CurrentBufferSize() & "'" & vbNewLine _
                & "LastWriteTime()='" & .LastWriteTime() &"'" _
                )
    .Flush()
    Call msgbox("WriteBufferSize()='" & .WriteBufferSize() & "'" & vbNewLine _
                & "WriteIntervalTime()='" & .WriteIntervalTime() & "'" & vbNewLine _
                & "CurrentBufferSize()='" & .CurrentBufferSize() & "'" & vbNewLine _
                & "LastWriteTime()='" & .LastWriteTime() &"'" _
                )
End With

wscript.quit

Dim vMin,vMax

vMin = -1 * 2^59 / 1000
vMax = ( 2^59 - 1 ) / 1000

vMin = vMin - 0.001
'vMax = vMax + 0.001

Call msgbox(vMin & vbNewLine & ccur(vMin))
Call msgbox(vMax & vbNewLine & ccur(vMax))




''vMin = -57646075230342.3488
''vMin = -57646075230342.3516
'vMin = -922337203685477.5808   '2^63/1000
''vMin = -1 * 2^59 / 1000
'vMax = ( 2^63 - 1 ) / 1000
'
''vMin = vMin - 0.001
''vMax = vMax + 0.001
'
'Call msgbox(vMin & vbNewLine & typename(vMin))
'Call msgbox(ccur(-922337203685477.5808))
''Call msgbox(vMin & vbNewLine & ccur(vMin))
''Call msgbox(vMax & vbNewLine & ccur(vMax))

wscript.quit

dim x : x = csng(-3.402823E38)
Call msgbox(x)

dim d : set d=new_Now()
wscript.Sleep 1500
dim d2 : set d2=new_Now()


'call msgbox(now() & vbnewline & cdbl(Fix(now())) & vbnewline & timer() & vbnewline & d.GetSerial() & vbnewline & new_Now().GetSerial())
call msgbox(d.DifferenceFrom(d2))
call msgbox(d2.DifferenceFrom(d))

wscript.quit


dim oBufferedWriter : set oBufferedWriter = new clsCmBufferedWriter

call msgbox(oBufferedWriter.Outpath)
oBufferedWriter.Outpath="yahoo!"
call msgbox(oBufferedWriter.Outpath)

call msgbox(oBufferedWriter.WriteBufferSize)
oBufferedWriter.WriteBufferSize=100
call msgbox(oBufferedWriter.WriteBufferSize)
oBufferedWriter.WriteBufferSize=0
call msgbox(oBufferedWriter.WriteBufferSize)
oBufferedWriter.WriteBufferSize=-1
call msgbox(oBufferedWriter.WriteBufferSize)
oBufferedWriter.WriteBufferSize=-1
call msgbox(oBufferedWriter.WriteBufferSize)
oBufferedWriter.WriteBufferSize=-2147483648
call msgbox(oBufferedWriter.WriteBufferSize)
oBufferedWriter.WriteBufferSize=-2147483649
call msgbox(oBufferedWriter.WriteBufferSize)
oBufferedWriter.WriteBufferSize=2147483647
call msgbox(oBufferedWriter.WriteBufferSize)

call msgbox(oBufferedWriter.WriteIntervalTime)
oBufferedWriter.WriteIntervalTime=100
call msgbox(oBufferedWriter.WriteIntervalTime)
oBufferedWriter.WriteIntervalTime=0
call msgbox(oBufferedWriter.WriteIntervalTime)
oBufferedWriter.WriteIntervalTime=-1
call msgbox(oBufferedWriter.WriteIntervalTime)
oBufferedWriter.WriteIntervalTime=-2147483648
call msgbox(oBufferedWriter.WriteIntervalTime)
oBufferedWriter.WriteIntervalTime=2147483648
call msgbox(oBufferedWriter.WriteIntervalTime)
oBufferedWriter.WriteIntervalTime=2147483647
call msgbox(oBufferedWriter.WriteIntervalTime)

call msgbox(oBufferedWriter.Iomode)
oBufferedWriter.Iomode="Google"
call msgbox(oBufferedWriter.Iomode)
oBufferedWriter.Iomode="ForReading"
call msgbox(oBufferedWriter.Iomode)

call msgbox(oBufferedWriter.FileFormat)
oBufferedWriter.FileFormat="TristateFalse"
call msgbox(oBufferedWriter.FileFormat)
oBufferedWriter.FileFormat="Goo"
call msgbox(oBufferedWriter.FileFormat)


wscript.quit


'dim s : s = "C:\Users\89585\Documents\dev\vbs\otameshi.vbs"
dim t : t = "C:\Users\89585\Documents\dev\vbs\"

call msgbox(func_CM_FsIsSame(s,t))
call msgbox(func_CM_FsIsSame(s,func_CM_StrConvOnlyAlphabet(s, 2)))
call msgbox(func_CM_FsIsSame(func_CM_StrConvOnlyAlphabet(t, 1),t))
call msgbox(func_CM_FsIsSame("",t))

wscript.quit



'定数

'import定義
Sub sub_import( _
    byVal asIncludeFileName _
    )
    With CreateObject("Scripting.FileSystemObject")
        Dim sParentFolderName : sParentFolderName = .GetParentFolderName(WScript.ScriptFullName)
        Dim sIncludeFilePath
        sIncludeFilePath = .BuildPath(sParentFolderName, Cs_FOLDER_LIB)
        sIncludeFilePath = .BuildPath(sIncludeFilePath, asIncludeFileName)
        ExecuteGlobal .OpenTextfile(sIncludeFilePath).ReadAll
    End With
End Sub
'import
Call sub_import("libCom.vbs")
Call sub_import("clsCompareExcel.vbs")
Call sub_import("clsCmCalendar.vbs")


call msgbox(new_Now().displayAs("M/d/yyyy h:m:s.000000"))

Dim hoge2 : Set hoge2 = new_Now()
wscript.Sleep 3

Call Msgbox(new_Now().differenceFrom(hoge2))

wscript.quit

Dim dtHogeNow : Dim dtHogeDate : Dim dtHogeTime : Dim dbTimer : Dim dtNow
dtHogeNow = Now()
dtHogeDate = Date()
dtHogeTime = Time()
dbTimer = Timer()

dtNow = dtHogeDate + dbTimer/(60*60*24)

Call Msgbox(Cdbl(dtHogeNow) & vbCrLf & Cdbl(dtHogeDate) & vbCrLf & Cdbl(dtHogeTime)  & vbCrLf & Cdbl(dtHogeDate+dtHogeTime) & vbCrLf & Cdbl(dtNow) & vbCrLf & (dbTimer / (60*60*24)) & vbCrLf & Cdbl(dtHogeNow)+(dbTimer-Fix(dbTimer))/(60*60*24) )
Call Msgbox( ((dbTimer/(60*60*24) - dtHogeTime)*60*60*24) & vbCrLf & dbTimer-Fix(dbTimer) )
Call Msgbox( dtHogeTime*60*60*24 & vbCrLf & Fix(dbTimer) & vbCrLf & dbTimer & vbCrLf & dbTimer-Fix(dbTimer))

wscript.quit

call msgbox(Len(vbnullstring))
wscript.quit

Dim oArray1(1)
Dim oArray2(1)
Dim oArray3(1)

'Call Msgbox(func_CM_ArrayGetDimensionNumber(sArray))

Dim oDic111 : Set oDic111 = new_Dic() : oDic111.Add 1, "Dic111"
Dim oDic112 : Set oDic112 = new_Dic() : oDic112.Add 1, "Dic112"
'Dim oDic121 : Set oDic121 = new_Dic() : oDic121.Add 1, "Dic121"
'Dim oDic122 : Set oDic122 = new_Dic() : oDic122.Add 1, "Dic122"
'Dim oDic211 : Set oDic211 = new_Dic() : oDic211.Add 1, "Dic211"
'Dim oDic212 : Set oDic212 = new_Dic() : oDic212.Add 1, "Dic212"
'Dim oDic221 : Set oDic221 = new_Dic() : oDic221.Add 1, "Dic221"
'Dim oDic222 : Set oDic222 = new_Dic() : oDic222.Add 1, "Dic222"

Set oArray3(0) = oDic111
Set oArray3(1) = oDic112

oArray2(0)=oArray3
oArray1(0)=oArray2

Call Msgbox( (oArray1(0)(0)(1)).Item(1) )

wscript.quit

Dim lCnt : Dim lDimensionNum

lDimensionNum = 1
For lCnt=0 To Ubound(sArray,lDimensionNum)
    If func_CM_ArrayGetDimensionNumber(sArray) > lDimensionNum Then
        lDimensionNum = lDimensionNum + 1
        '再帰処理(lDimensionNum)
    Else
        Call Msgbox( sArray(lCnt1, lCnt2).Item(1) )
    End If
Next

wscript.quit


Call Msgbox(new_Fso().GetFile("C:\Users\89585\Documents\dev\vbs\lib\libCom.vbs").DateLastModified)
Call Msgbox(new_Fso().GetFile("C:\Users\89585\Documents\dev\vbs\lib\libCom.vbs").Item(1))

wscript.quit


Dim sPath(3)
sPath(1) = "C:\Users\89585\Documents\dev\vbs\lib\libCom.vbs"
sPath(2) = "C:\Users\89585\Documents\dev\vbs\lib"
sPath(3) = "C:\Users\89585\Documents\dev\vbs\lib.abc"

'Dim lCnt
For lCnt=1 To Ubound(sPath)
    Call Msgbox(sPath(lCnt))
    Call Msgbox("Basename : " & func_CM_FsGetGetBaseName(sPath(lCnt)) &", Extension : " & func_CM_FsGetGetExtensionName(sPath(lCnt)))
Next
wscript.quit


Dim sStr(6)
sStr(1) = "filename_221023.txt"
sStr(2) = "FILENAME_20221023_2.txt"
sStr(3) = "FileName_221023.xlsx"
sStr(4) = "fileNAME_20221023_abc.txt"
sStr(5) = "FILEname_221024.txt"
sStr(6) = "FilenamE_221024_999.txt"

Dim sBasename : sBasename = "filename"
Dim sExt : sExt = "txt"

With New RegExp
    '初期化
    .Pattern = sBasename & "_" & "(20)?(\d{2}[01]\d[0123]\d)" & "((_)(\d+))?" & "." & sExt
    .IgnoreCase = True
    .Global = True
    
'    Dim lCnt : Dim sTemp
    Call Msgbox(.Pattern)
    For lCnt=1 To Ubound(sStr)
        sTemp = sStr(lCnt)
        Call Msgbox(sTemp & " : " &  .Test(sTemp))
        If .Test(sTemp) Then
            Call Msgbox("日付 : " &  .Replace(sTemp, "$2") & ", サフィックス : " &  .Replace(sTemp, "$5"))
        End If
    Next
    
End With


Dim oEc : Set oEc = New clsCompareExcel
oEc.PathFrom = "G:\マイドライブ\30_プライベート\40_資格取得\午前Ⅰの過去問.xlsx"
oEc.PathTo = "G:\マイドライブ\30_プライベート\40_資格取得\午前Ⅰの過去問.xlsx"

If Len(oEc.PathFrom&oEc.PathTo) Then msgbox "ok"
msgbox oEc.PathFrom
msgbox oEc.PathTo
msgbox oEc.Compare()
msgbox oEc.ProcDate
msgbox oEc.StartTime
msgbox oEc.EndTime
msgbox oEc.ElapsedTime
wscript.quit

Call Msgbox(func_CM_FsGetParentFolderPath("c:\a\b") & Err.Number)
Call Msgbox(func_CM_FsGetParentFolderPath("C:\A\") & Err.Number)
Call Msgbox(func_CM_FsGetParentFolderPath("C:\a") & Err.Number)
Call Msgbox(func_CM_FsGetParentFolderPath("c:\") & Err.Number)
Call Msgbox(func_CM_FsGetParentFolderPath("C:") & Err.Number)
