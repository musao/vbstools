'***************************************************************************************************
'FILENAME                    : GeneratePassword.vbs
'Overview                    : パスワードを生成する
'Detailed Description        : 生成したパスワードはクリップボードにコピーする
'Argument                    : 以下の名前付き引数（/Key:Value 形式）のみ、名前なし引数は無視する
'                                /Length : 生成するパスワードの文字数
'                                /U      : 生成するパスワードの文字種に半角英字大文字を使用する
'                                /L      : 生成するパスワードの文字種に半角英字小文字を使用する
'                                /N      : 生成するパスワードの文字種に半角数字を使用する
'                                /S      : 生成するパスワードの文字種に記号を使用する
'                                            記号の種類   !"#$%&'()*+,-./:;<=>?@[\]^_`{|}~（32種類）
'                                /Add    : 追加指定する文字種（カンマ区切りで複数指定可能）
'Return Value
'     なし
'---------------------------------------------------------------------------------------------------
'History
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/24         Y.Fujii                  First edition
'***************************************************************************************************
Option Explicit

'lib\com import
Dim sRelativeFolderName : sRelativeFolderName = "lib\com"
With CreateObject("Scripting.FileSystemObject")
    Dim sParentFolderPath : sParentFolderPath = .GetParentFolderName(WScript.ScriptFullName)
    Dim sLibFolderPath : sLibFolderPath = .BuildPath(sParentFolderPath, sRelativeFolderName)
    Dim oLibFile
    For Each oLibFile In CreateObject("Shell.Application").Namespace(sLibFolderPath).Items
        If Not oLibFile.IsFolder Then
            If StrComp(.GetExtensionName(oLibFile.Path), "vbs", vbTextCompare)=0 Then ExecuteGlobal .OpenTextfile(oLibFile.Path).ReadAll
        End If
    Next
End With
Set oLibFile = Nothing
'lib import
sRelativeFolderName = "lib"
With new_FSO()
    sLibFolderPath = .BuildPath(sParentFolderPath, sRelativeFolderName)
    ExecuteGlobal .OpenTextfile(.BuildPath(sLibFolderPath,"libEnum.vbs")).ReadAll
End With


'ログ出力先、ブローカークラスのインスタンスの設定
Private PoTs4Log, PoBroker
Set PoTs4Log = fw_getTextstreamForLog()
Set PoBroker = new_BrokerOf(Array(topic.LOG, GetRef("this_logger")))

'Main関数実行
Call Main()

'終了処理
PoTs4Log.close()
Set PoBroker = Nothing : Set PoTs4Log = Nothing
Wscript.Quit


'***************************************************************************************************
'Processing Order            : First
'Function/Sub Name           : Main()
'Overview                    : メイン関数
'Detailed Description        : 工事中
'Argument
'     なし
'Return Value
'     なし
'---------------------------------------------------------------------------------------------------
'History
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/24         Y.Fujii                  First edition
'***************************************************************************************************
Sub Main()
    'パラメータ格納用オブジェクト宣言
    Dim oParams : Set oParams = new_Dic()
    
    '当スクリプトの引数をパラメータ格納用オブジェクトに取得する
    fw_excuteSub "this_getParameters", oParams, PoBroker
    
    'パスワードを生成する
    fw_excuteSub "this_generate", oParams, PoBroker
    
    'ログ出力をクローズ
    PoTs4Log.close()
    
    'オブジェクトを開放
    Set oParams = Nothing
End Sub

'***************************************************************************************************
'Processing Order            : 1
'Function/Sub Name           : this_getParameters()
'Overview                    : 当スクリプトの引数をパラメータ格納用オブジェクトに取得する
'Detailed Description        : 名前付き引数（/Key:Value 形式）だけを取得する
'                              Key           Value                                     Default
'                              ------------  ----------------------------------------  -------------
'                              "Param"       パラメータの解析結果
'
'                              名前付き引数（/Key:Value 形式）の構成
'                              Key           Value                                     Default
'                              ------------  ----------------------------------------  -------------
'                              "Length"      文字の長さ                                16
'                                            文字の種類                                全て含む
'                               "U"           半角英字大文字
'                               "L"           半角英字小文字
'                               "N"           半角数字
'                               "S"           全ての記号
'                              "Add"         追加指定する文字種をカンマ区切りで指定    なし
'Argument
'     aoParams               : パラメータ格納用オブジェクト
'Return Value
'     なし
'---------------------------------------------------------------------------------------------------
'History
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/24         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub this_getParameters( _
    byRef aoParams _
    )
    'オリジナルの引数を取得
    Dim oArg : Set oArg = fw_storeArguments()
    '★ログ出力
    this_logger Array(logType.TRACE, "this_getParameters()", cf_toString(oArg))
    
    '引数の内容を解析
    
    '文字の長さ
    Dim oKey, lLength
    oKey = "Length"
    If oArg.Item("Named").Exists(oKey) Then lLength = oArg.Item("Named").Item(oKey) Else lLength = 16
    
    '追加指定する文字種
    Dim vAdd : vAdd = Empty
    oKey = "Add"
    If oArg.Item("Named").Exists(oKey) Then 
        vAdd = new_ArrSplit(oArg.Item("Named").Item(oKey), ",", vbBinaryCompare).toArray()
    End If
    
    '文字の種類
    Dim oSetting, lSum, lType
    Set oSetting = new_DicOf(Array("U", 1, "L", 2, "N", 4, "S", 8))
    lSum = 0
    For Each oKey In oSetting.Keys
        If oArg.Item("Named").Exists(oKey) Then lSum = lSum + oSetting.Item(oKey)
    Next
    lType = lSum
    If lType = 0 And IsEmpty(vAdd) Then lType = 15
    
    Dim oParam : Set oParam = new_DicOf(Array("Length", lLength, "Type", lType, "Additional", vAdd))
    
    'パラメータ格納用オブジェクトに設定
    cf_bindAt aoParams, "Param", oParam
    
    Set oParam = Nothing
    Set oArg = Nothing
End Sub

'***************************************************************************************************
'Processing Order            : 2
'Function/Sub Name           : this_generate()
'Overview                    : パスワードを生成する
'Detailed Description        : 生成したパスワードはクリップボードにコピーし、InputBoxに表示する
'Argument
'     aoParams               : パラメータ格納用オブジェクト
'Return Value
'     なし
'---------------------------------------------------------------------------------------------------
'History
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/24         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub this_generate( _
    byRef aoParams _
    )
    'パスワード生成
    Dim lLength, lType, vAdd
    With aoParams.Item("Param")
        cf_bind lLength, .Item("Length")
        cf_bind lType, .Item("Type")
        cf_bind vAdd, .Item("Additional")
    End With
    Dim vCharList : vCharList = new_Char().charList(lType)
    vCharList = Filter(vCharList, " ", False, vbBinaryCompare)
    If Not IsEmpty(vAdd) Then cf_pushA vCharList, vAdd
    Dim sPw : sPw = util_randStr(vCharList, lLength)
    
    '★ログ出力
    this_logger Array(logType.INFO, "this_generate()", "Generated password is " & sPw)
    
    'ダイアログのメッセージなどを作成
    Dim sMsg, sTitle
    sMsg = "パスワードを生成しました" & vbNewLine & "OKボタンを押下するとクリップボードにコピーします"
    sTitle = new_Now() & " に作成"
    
    '★ログ出力
    this_logger Array(logType.INFO, "this_generate()", "Display Inputbox.")
    '一時ファイルのパスを作成
    Dim sPath : sPath = fw_getTempPath()
    Do Until Inputbox(sMsg, sTitle, sPw)=False
        '一時ファイルに生成したパスワードを出力
        fs_writeFileDefault sPath, sPw
        'クリップボードに一時ファイルの内容を出力
        fw_runShellSilently "cmd /c clip <" & fs_wrapInQuotes(sPath)
        '一時ファイルを削除
        fs_deleteFile sPath
        '★ログ出力
        this_logger Array(logType.INFO, "this_generate()", "Copied to clipboard.")
    Loop
    
End Sub

'***************************************************************************************************
'Processing Order            : -
'Function/Sub Name           : this_logger()
'Overview                    : ログ出力する
'Detailed Description        : fw_logger()に委譲する
'Argument
'     avParams               : 配列型のパラメータリスト
'Return Value
'     なし
'---------------------------------------------------------------------------------------------------
'History
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/03         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub this_logger( _
    byRef avParams _
    )
    fw_logger avParams, PoTs4Log
End Sub
