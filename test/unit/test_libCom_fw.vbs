' libCom.vbs: fw_* procedure test.
' @import ../../lib/clsAdptFile.vbs
' @import ../../lib/clsCmArray.vbs
' @import ../../lib/clsCmBroker.vbs
' @import ../../lib/clsCmBufferedReader.vbs
' @import ../../lib/clsCmBufferedWriter.vbs
' @import ../../lib/clsCmCalendar.vbs
' @import ../../lib/clsCmCharacterType.vbs
' @import ../../lib/clsCmCssGenerator.vbs
' @import ../../lib/clsCmHtmlGenerator.vbs
' @import ../../lib/clsCmReturnValue.vbs
' @import ../../lib/clsCompareExcel.vbs
' @import ../../lib/libCom.vbs

Option Explicit

Const MY_NAME = "test_libCom_fw.vbs"
Const FILE_NAME = "test.txt"
Const NoArg_CONT = "NoArg"
Dim PsPathTempFolder,PsPath,PvLog

'###################################################################################################
'SetUp()/TearDown()
Sub SetUp()
    '実行スクリプト直下に当ファイル名で一時フォルダ作成
    PsPathTempFolder = new_Fso().BuildPath(new_Fso().GetParentFolderName(WScript.ScriptFullName), MY_NAME)
    If Not (new_Fso().FolderExists(PsPathTempFolder)) Then new_Fso().CreateFolder(PsPathTempFolder)
    'テスト用のファイルパス作成
    PsPath = new_Fso().BuildPath(PsPathTempFolder, FILE_NAME)
End Sub
Sub TearDown()
    '当テストで作成した一時フォルダを削除する
    new_Fso().DeleteFolder PsPathTempFolder
End Sub

'###################################################################################################
'fw_excuteSub()
Sub Test_fw_excuteSub_Arg_NoBroker_Normal
    Dim f,e,d,a,b
    Set b = Nothing : PvLog = Array()
    f = "subArg"
    d = "Arg_NoBroker_Normal"
    e = d
    
    fw_excuteSub f, d, b
    a = readFile()
    AssertEqualWithMessage e, a, "result"
    AssertEqualWithMessage 0, new_Arr().hasElement(PvLog), "log"
End Sub
Sub Test_fw_excuteSub_Arg_NoBroker_Err
    Dim f,d,b
    b = Empty : PvLog = Array()
    f = "subArg"
    d = "Arg_NoBroker_Err"
    
    fw_excuteSub f, d, b
    AssertEqualWithMessage False, new_Fso().FileExists(PsPath), "after excute file exists"
    AssertEqualWithMessage 0, new_Arr().hasElement(PvLog), "log"
End Sub
Sub Test_fw_excuteSub_NoArg_NoBroker_Normal
    Dim f,e,d,a,b
    b = Empty : PvLog = Array()
    f = "subNoArg"
    Set d = Nothing
    e = NoArg_CONT

    fw_excuteSub f, d, b
    a = readFile()
    AssertEqualWithMessage e, a, "result"
    AssertEqualWithMessage 0, new_Arr().hasElement(PvLog), "log"
End Sub
Sub Test_fw_excuteSub_NoArg_NoBroker_Err
    Dim f,d,b
    Set b = Nothing : PvLog = Array()
    f = "subNoArgErr"
    d = Empty
    
    fw_excuteSub f, d, b
    AssertEqualWithMessage False, new_Fso().FileExists(PsPath), "after excute file exists"
    AssertEqualWithMessage 0, new_Arr().hasElement(PvLog), "log"
End Sub
Sub Test_fw_excuteSub_Arg_Broker_Normal
    Dim f,e,d,a,b
    Set b = new_Broker() : b.subscribe "log", GetRef("broker") : PvLog = Array()
    f = "subArg"
    d = "Arg_Broker_Normal"
    e = d
    
    fw_excuteSub f, d, b
    a = readFile()
    AssertEqualWithMessage e, a, "result"
    assertLogs f, d, False
End Sub
Sub Test_fw_excuteSub_Arg_Broker_Err
    Dim f,d,b
    Set b = new_Broker() : b.subscribe "log", GetRef("broker") : PvLog = Array()
    f = "subArg"
    d = "Arg_Broker_Err"
    
    fw_excuteSub f, d, b
    AssertEqualWithMessage False, new_Fso().FileExists(PsPath), "after excute file exists"
    assertLogs f, d, True
End Sub
Sub Test_fw_excuteSub_NoArg_Broker_Normal
    Dim f,e,d,a,b
    Set b = new_Broker() : b.subscribe "log", GetRef("broker") : PvLog = Array()
    f = "subNoArg"
    d = Empty
    e = NoArg_CONT

    fw_excuteSub f, d, b
    a = readFile()
    AssertEqualWithMessage e, a, "result"
    assertLogs f, d, False
End Sub
Sub Test_fw_excuteSub_NoArg_Broker_Err
    Dim f,d,b
    Set b = new_Broker() : b.subscribe "log", GetRef("broker") : PvLog = Array()
    f = "subNoArgErr"
    Set d = Nothing
    
    fw_excuteSub f, d, b
    AssertEqualWithMessage False, new_Fso().FileExists(PsPath), "after excute file exists"
    assertLogs f, d, True
End Sub

'---------------------------------------------------------------------------------------------------
'stub()
Sub subArg(aArg)
    If Instr(1,aArg,"Err",vbBinaryCompare)>0 Then
        Err.Raise 9999, "エラー", "test_libCom_fw.vbsのエラーケース"
        Exit Sub
    End If
    With new_Ts(PsPath, 2, True, -2)
        .Write aArg
        .Close
    End With
End Sub
Sub subNoArg()
    With new_Ts(PsPath, 2, True, -2)
        .Write NoArg_CONT
        .Close
    End With
End Sub
Sub subNoArgErr()
    Err.Raise 9999, "エラー", "test_libCom_fw.vbsのエラーケース"
    Exit Sub
    subNoArg
End Sub

'###################################################################################################
'fw_getLogPath()
Sub Test_fw_getLogPath
    Dim a,e
    e = "^.*\\log\\TestRunner\.log$"
    a = fw_getLogPath()

    AssertMatchWithMessage e, a, "ret"

'    new_Fso().DeleteFolder(new_Fso().GetParentFolderName(a))
End Sub
Sub Test_fw_getLogPath_FolderExists
'    new_Fso().CreateFolder( _
'        new_Fso().BuildPath( _
'            new_Fso().GetParentFolderName( _
'                new_Fso().GetParentFolderName(WScript.ScriptName) _
'            ) _
'            , "log" _
'        ) _
'    )
    
    Dim a,e
    e = "^.*\\log\\TestRunner\.log$"
    a = fw_getLogPath()

    AssertMatchWithMessage e, a, "ret"

    new_Fso().DeleteFolder(new_Fso().GetParentFolderName(a))
End Sub

'###################################################################################################
'fw_getPrivatePath()
Sub Test_fw_getPrivatePath_Parent
    Dim dpf,dfn
    dpf="pfolder"
    dfn="test.txt"

    Dim a,e
    e = "^.*\\"&dpf&"\\"&dfn&"$"
    a = fw_getPrivatePath(dpf,dfn)

    AssertMatchWithMessage e, a, "ret"

'    new_Fso().DeleteFolder(new_Fso().GetParentFolderName(a))
End Sub
Sub Test_fw_getPrivatePath_ParentFolderExists
    Dim dpf,dfn
    dpf="pfolder"
    dfn="test.txt"

    Dim a,e
    e = "^.*\\"&dpf&"\\"&dfn&"$"
    a = fw_getPrivatePath(dpf,dfn)

    AssertMatchWithMessage e, a, "ret"

    new_Fso().DeleteFolder(new_Fso().GetParentFolderName(a))
End Sub
Sub Test_fw_getPrivatePath_NoParent
    Dim dpf,dfn
    dpf=vbNullString
    dfn="test.txt"

    Dim a,e
    e = "^.*\\"&dfn&"$"
    a = fw_getPrivatePath(dpf,dfn)

    AssertMatchWithMessage e, a, "ret"

End Sub

'###################################################################################################
'fw_getTempPath()
Sub Test_fw_getTempPath
    Dim a,e
    e = "^.*\\tmp\\[\d\w]{8}\.tmp$"
    a = fw_getTempPath()

    AssertMatchWithMessage e, a, "ret"

'    new_Fso().DeleteFolder(new_Fso().GetParentFolderName(a))
End Sub
Sub Test_fw_getTempPath_FolderExists
'    new_Fso().CreateFolder( _
'        new_Fso().BuildPath( _
'            new_Fso().GetParentFolderName( _
'                new_Fso().GetParentFolderName(WScript.ScriptName) _
'            ) _
'            , "tmp" _
'        ) _
'    )

    Dim a,e
    e = "^.*\\tmp\\[\d\w]{8}\.tmp$"
    a = fw_getTempPath()

    AssertMatchWithMessage e, a, "ret"

    new_Fso().DeleteFolder(new_Fso().GetParentFolderName(a))
End Sub

'###################################################################################################
'fw_logger()
Sub Test_fw_logger
    Const RE_DATE = "^\d{4}/\d{2}/\d{2} \d{2}:\d{2}:\d{2}\.\d{3}$"
    Const RE_IP4 = "(([1-9]?[0-9]|1[0-9]{2}|2[0-4][0-9]|25[0-5])\.){3}([1-9]?[0-9]|1[0-9]{2}|2[0-4][0-9]|25[0-5])"
    
    Dim ts,d,a,e
    Set ts = new_Ts(PsPath, 2, True, -2)
    d = Array("fw_logger's Test")
    e = d(0)
    
    fw_logger d, ts
    ts.Close

    a = Split(Replace(readFile(),vbNewLine,""),vbTab,-1,vbBinaryCompare)
    AssertEqualWithMessage 3, Ubound(a), "Ubound"
    AssertMatchWithMessage RE_DATE, a(0), "DateTime"
    AssertMatchWithMessage "^"&RE_IP4&"(,"&RE_IP4&")*$", a(1), "IpAddress"
    AssertWithMessage Len(a(2))>0, "HostName"
    AssertEqualWithMessage e, a(3), "data"
End Sub

'###################################################################################################
'fw_storeErr()
Sub Test_fw_storeErr_NoErr
    dim a : Set a = fw_storeErr()

    dim k,v
    k = "__Special__" : v = "Err"
    AssertEqualWithMessage v, a.Item(k), k
    k = "Number" : v = 0
    AssertEqualWithMessage v, a.Item(k), k
    k = "Description" : v = vbNullString
    AssertEqualWithMessage v, a.Item(k), k
    k = "Source" : v = vbNullString
    AssertEqualWithMessage v, a.Item(k), k
End Sub
Sub Test_fw_storeErr_Err
    Dim e : Set e = new_DicWith(Array("Number", 1234, "Description", "Val_Description", "Source", "Val_Source"))
    On Error Resume Next
    Err.Raise e.Item("Number"), e.Item("Source"), e.Item("Description")
    dim a : Set a = fw_storeErr()
    On Error Goto 0

    dim k,v
    k = "__Special__" : v = "Err"
    AssertEqualWithMessage v, a.Item(k), k
    k = "Number" : v = e.Item(k)
    AssertEqualWithMessage v, a.Item(k), k
    k = "Description" : v = e.Item(k)
    AssertEqualWithMessage v, a.Item(k), k
    k = "Source" : v = e.Item(k)
    AssertEqualWithMessage v, a.Item(k), k
End Sub

'###################################################################################################
'fw_storeArguments()
Sub Test_fw_storeArguments
    dim a : Set a = fw_storeArguments()

    dim k,v
    k = "__Special__" : v = "Arguments"
    AssertEqualWithMessage v, a.Item(k), k
    k = "All"
    AssertEqualWithMessage 0, Ubound(a.Item(k)), k
'    k = "Named"
'    AssertEqualWithMessage 0, a.Item(k).Count, k
'    k = "Unnamed"
'    AssertEqualWithMessage 0, Ubound(a.Item(k)), k
End Sub

'###################################################################################################
'fw_tryCatch()
Sub Test_cf_tryCatch_TryOnly_Normal
    Dim oRet : Set oRet = fw_tryCatch(new_Func("a=>1/a"), 2, Nothing, Empty)
    
    AssertEqual 0, Err.Number
    AssertEqual False, oRet.isErr()
    AssertEqual 1/2, oRet
End Sub
Sub Test_cf_tryCatch_TryAndCatch_Normal
    Dim oRet : Set oRet = fw_tryCatch(new_Func("a=>1/a"), 2, new_Func("a=>a"), Nothing)
    
    AssertEqual 0, Err.Number
    AssertEqual False, oRet.isErr()
    AssertEqual 1/2, oRet
End Sub
Sub Test_cf_tryCatch_TryAndFinary_Normal
    Dim oRet : Set oRet = fw_tryCatch(new_Func("a=>1/a"), 2, Empty, new_Func("r=>1/2+r"))
    
    AssertEqual 0, Err.Number
    AssertEqual False, oRet.isErr()
    AssertEqual 1/2+1/2, oRet
End Sub
Sub Test_cf_tryCatch_TryAndFinary_Normal_FinaryErr
    On Error Resume Next
    Dim oRet : Set oRet = fw_tryCatch(new_Func("a=>1/a"), 2, Empty, new_Func("r=>r(0)"))
    
    AssertEqual 13, Err.Number
    AssertEqual "型が一致しません。", Err.Description
    AssertEqual Empty, oRet
End Sub
Sub Test_cf_tryCatch_TryAndCatchAndFinary_Normal
    Dim oRet : Set oRet = fw_tryCatch(new_Func("a=>1/a"), 2, new_Func("a=>a"), new_Func("r=>1/2+r"))
    
    AssertEqual 0, Err.Number
    AssertEqual False, oRet.isErr()
    AssertEqual 1/2+1/2, oRet
End Sub
Sub Test_cf_tryCatch_TryAndCatchAndFinary_Normal_FinaryErr
    On Error Resume Next
    Dim oRet : Set oRet = fw_tryCatch(new_Func("a=>1/a"), 2, new_Func("a=>a"), new_Func("r=>r(0)"))
    
    AssertEqual 13, Err.Number
    AssertEqual "型が一致しません。", Err.Description
    AssertEqual Empty, oRet
End Sub
Sub Test_cf_tryCatch_TryOnly_Err
    Dim oRet : Set oRet = fw_tryCatch(new_Func("a=>1/a"), 0, Empty, Empty)
    
    AssertEqual 0, Err.Number
    AssertEqual True, oRet.isErr()
    AssertEqual Empty, oRet
    AssertEqual 11, oRet.getErr().Item("Number")
    AssertEqual "0 で除算しました。", oRet.getErr().Item("Description")
    AssertEqual "Microsoft VBScript 実行時エラー", oRet.getErr().Item("Source")
End Sub
Sub Test_cf_tryCatch_TryAndCatch_Err
    Dim oRet : Set oRet = fw_tryCatch(new_Func("a=>1/a"), 0, new_Func("a=>a"), Empty)
    
    AssertEqual 0, Err.Number
    AssertEqual True, oRet.isErr()
    AssertEqual 0, oRet
    AssertEqual 11, oRet.getErr().Item("Number")
    AssertEqual "0 で除算しました。", oRet.getErr().Item("Description")
    AssertEqual "Microsoft VBScript 実行時エラー", oRet.getErr().Item("Source")
End Sub
Sub Test_cf_tryCatch_TryAndCatch_Err_CatchErr
    On Error Resume Next
    Dim oRet : Set oRet = fw_tryCatch(new_Func("a=>1/a"), 0, new_Func("a=>a(0)"), Empty)
    
    AssertEqual 13, Err.Number
    AssertEqual "型が一致しません。", Err.Description
    AssertEqual Empty, oRet
End Sub
Sub Test_cf_tryCatch_TryAndFinary_Err
    Dim oRet : Set oRet = fw_tryCatch(new_Func("a=>1/a"), 0, Nothing, new_Func("r=>2"))
    
    AssertEqual 0, Err.Number
    AssertEqual True, oRet.isErr()
    AssertEqual 2, oRet
    AssertEqual 11, oRet.getErr().Item("Number")
    AssertEqual "0 で除算しました。", oRet.getErr().Item("Description")
    AssertEqual "Microsoft VBScript 実行時エラー", oRet.getErr().Item("Source")
End Sub
Sub Test_cf_tryCatch_TryAndFinary_Err_FinaryErr
    On Error Resume Next
    Dim oRet : Set oRet = fw_tryCatch(new_Func("a=>1/a"), 0, Nothing, new_Func("r=>r(0)"))
    
    AssertEqual 13, Err.Number
    AssertEqual "型が一致しません。", Err.Description
    AssertEqual Empty, oRet
End Sub
Sub Test_cf_tryCatch_TryAndCatchAndFinary_Err
    Dim oRet : Set oRet = fw_tryCatch(new_Func("a=>1/a"), 0, new_Func("a=>a"), new_Func("r=>2"))
    
    AssertEqual 0, Err.Number
    AssertEqual True, oRet.isErr()
    AssertEqual 2, oRet
    AssertEqual 11, oRet.getErr().Item("Number")
    AssertEqual "0 で除算しました。", oRet.getErr().Item("Description")
    AssertEqual "Microsoft VBScript 実行時エラー", oRet.getErr().Item("Source")
End Sub
Sub Test_cf_tryCatch_TryAndCatchAndFinary_Err_CatchErr
    On Error Resume Next
    Dim oRet : Set oRet = fw_tryCatch(new_Func("a=>1/a"), 0, new_Func("a=>a(0)"), new_Func("r=>2"))
    
    AssertEqual 13, Err.Number
    AssertEqual "型が一致しません。", Err.Description
    AssertEqual Empty, oRet
End Sub
Sub Test_cf_tryCatch_TryAndCatchAndFinary_Err_FinaryErr
    On Error Resume Next
    Dim oRet : Set oRet = fw_tryCatch(new_Func("a=>1/a"), 0, new_Func("a=>a"), new_Func("r=>r(0)"))
    
    AssertEqual 13, Err.Number
    AssertEqual "型が一致しません。", Err.Description
    AssertEqual Empty, oRet
End Sub
Sub Test_cf_tryCatch_TryOnly_ArgEmpty
    Dim oRet : Set oRet = fw_tryCatch(new_Func("()=>1/2"), Empty, Nothing, Empty)
    
    AssertEqual 0, Err.Number
    AssertEqual False, oRet.isErr()
    AssertEqual 1/2, oRet
End Sub
Sub Test_cf_tryCatch_TryAndCatch_ArgEmpty
    Dim oRet : Set oRet = fw_tryCatch(new_Func("=>1/0"), Empty, new_Func("=>1/2"), Nothing)
    
    AssertEqual 0, Err.Number
    AssertEqual True, oRet.isErr()
    AssertEqual 1/2, oRet
    AssertEqual 11, oRet.getErr().Item("Number")
    AssertEqual "0 で除算しました。", oRet.getErr().Item("Description")
    AssertEqual "Microsoft VBScript 実行時エラー", oRet.getErr().Item("Source")
End Sub

'###################################################################################################
'common
Function readFile()
    readFile = Empty
    On Error Resume Next
    With new_Ts(PsPath, 1, False, -2)
        readFile = .ReadAll
        .Close
    End With
    On Error Goto 0
    new_Fso().DeleteFile PsPath
End Function
Sub broker(arg)
    cf_push PvLog, arg
End Sub
Function assertLogs(f,d,isErr)
    Const ERR_STR = "<Err>{<String>""Number""=><Long>9999,<String>""Description""=><String>""test_libCom_fw.vbsのエラーケース"",<String>""Source""=><String>""エラー""}"
    If isErr Then
        AssertEqualWithMessage 4, Ubound(PvLog), "Ubound"
    Else
        AssertEqualWithMessage 3, Ubound(PvLog), "Ubound"
    End If
    Dim i : i=0
    AssertEqualWithMessage 5, PvLog(i)(0), i&"-0"
    AssertEqualWithMessage f, PvLog(i)(1), i&"-1"
    AssertEqualWithMessage "Start", PvLog(i)(2), i&"-2"
    i=i+1
    AssertEqualWithMessage 9, PvLog(i)(0), i&"-0"
    AssertEqualWithMessage f, PvLog(i)(1), i&"-1"
    AssertEqualWithMessage cf_toString(d), PvLog(i)(2), i&"-2"
    If isErr Then
        i=i+1
        AssertEqualWithMessage 1, PvLog(i)(0), i&"-0"
        AssertEqualWithMessage f, PvLog(i)(1), i&"-1"
        AssertEqualWithMessage ERR_STR, PvLog(i)(2), i&"-2"
    End If
    i=i+1
    AssertEqualWithMessage 5, PvLog(i)(0), i&"-0"
    AssertEqualWithMessage f, PvLog(i)(1), i&"-1"
    AssertEqualWithMessage "End", PvLog(i)(2), i&"-2"
    i=i+1
    AssertEqualWithMessage 9, PvLog(i)(0), i&"-0"
    AssertEqualWithMessage f, PvLog(i)(1), i&"-1"
    AssertEqualWithMessage cf_toString(d), PvLog(i)(2), i&"-2"
End Function



' Local Variables:
' mode: Visual-Basic
' indent-tabs-mode: nil
' End:
