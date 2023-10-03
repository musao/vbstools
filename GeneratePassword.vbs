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
'                                            記号の種類   !"#$%&'()*+,-./:;<=>?[\]^_`{|}~（31種類）
'                                /Add    : 追加指定する文字種（カンマ区切りで複数指定可能）
'Return Value
'     なし
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/24         Y.Fujii                  First edition
'***************************************************************************************************
Option Explicit

'定数
Private Const Cs_FOLDER_INCLUDE = "include"
Private PoWriter, PoPubSub

'Include用関数定義
Sub sub_Include( _
    byVal asIncludeFileName _
    )
    With CreateObject("Scripting.FileSystemObject")
        Dim sParentFolderName : sParentFolderName = .GetParentFolderName(WScript.ScriptFullName)
        Dim sIncludeFilePath
        sIncludeFilePath = .BuildPath(sParentFolderName, Cs_FOLDER_INCLUDE)
        sIncludeFilePath = .BuildPath(sIncludeFilePath, asIncludeFileName)
        ExecuteGlobal .OpenTextfile(sIncludeFilePath).ReadAll
    End With
End Sub
'Include
Call sub_Include("clsCmArray.vbs")
Call sub_Include("clsCmBufferedWriter.vbs")
Call sub_Include("clsCmCalendar.vbs")
Call sub_Include("clsCmPubSub.vbs")
Call sub_Include("clsCompareExcel.vbs")
Call sub_Include("VbsBasicLibCommon.vbs")

'メイン関数実行
Call Main()
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
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/24         Y.Fujii                  First edition
'***************************************************************************************************
Sub Main()
    'ログ出力の設定
    Dim sPath : sPath = func_CM_FsGetPrivateLogFilePath()
    Set PoWriter = new_clsCmBufferedWriter(func_CM_FsOpenTextFile(sPath, 8, True, -2))
    '出版-購読型（Publish/subscribe）インスタンスの設定
    Set PoPubSub = new_clsCmPubSub()
    Call PoPubSub.Subscribe("log", GetRef("sub_GnrtPwLogger"))
    'パラメータ格納用オブジェクト宣言
    Dim oParams : Set oParams = new_Dictionary()
    
    '当スクリプトの引数をパラメータ格納用オブジェクトに取得する
    Call sub_CM_ExcuteSub("sub_GnrtPwGetParameters", oParams, PoPubSub, "log")
    
    'パスワードを生成する
    Call sub_CM_ExcuteSub("sub_GnrtPwGenerate", oParams, PoPubSub, "log")
    
    'ファイル接続をクローズする
    Call PoWriter.Close()
    
    'オブジェクトを開放
    Set oParams = Nothing
    Set PoPubSub = Nothing
    Set PoWriter = Nothing
End Sub

'***************************************************************************************************
'Processing Order            : 1
'Function/Sub Name           : sub_GnrtPwGetParameters()
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
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/24         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub sub_GnrtPwGetParameters( _
    byRef aoParams _
    )
    'オリジナルの引数を取得
    Dim oArg : Set oArg = func_CM_UtilStoringArguments()
    '★ログ出力
    Call sub_GnrtPwLogger(Array(9, "sub_GnrtPwGetParameters", func_CM_ToStringArguments()))
    
    '引数の内容を解析
    
    '文字の長さ
    Dim oKey, lLength
    oKey = "Length"
    If oArg.Item("Named").Exists(oKey) Then lLength = oArg.Item("Named").Item(oKey) Else lLength = 16
    
    '追加指定する文字種
    Dim vAdd
    oKey = "Add"
    If oArg.Item("Named").Exists(oKey) Then 
        vAdd = new_ArraySplit(oArg.Item("Named").Item(oKey), ",", vbBinaryCompare).Items
    Else
        vAdd = Empty
    End If
    
    '文字の種類
    Dim oSetting, lSum, lType
    Set oSetting = new_DictSetValues(Array("U", 1, "L", 2, "N", 4, "S", 8))
    lSum = 0
    For Each oKey In oSetting.Keys
        If oArg.Item("Named").Exists(oKey) Then lSum = lSum + oSetting.Item(oKey)
    Next
    lType = lSum
    If lType = 0 And func_CM_ArrayIsAvailable(vAdd)<>True Then lType = 15
    
    Dim oParam : Set oParam = new_DictSetValues(Array("Length", lLength, "Type", lType, "Additional", vAdd))
    
    'パラメータ格納用オブジェクトに設定
    Call sub_CM_BindAt(aoParams, "Param", oParam)
    
    Set oParam = Nothing
    Set oArg = Nothing
End Sub

'***************************************************************************************************
'Processing Order            : 2
'Function/Sub Name           : sub_GnrtPwGenerate()
'Overview                    : パスワードを生成する
'Detailed Description        : 生成したパスワードはクリップボードにコピーし、InputBoxに表示する
'Argument
'     aoParams               : パラメータ格納用オブジェクト
'Return Value
'     なし
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/24         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub sub_GnrtPwGenerate( _
    byRef aoParams _
    )
    'パスワード生成
    Dim lLength, lType, vAdd
    With aoParams.Item("Param")
        Call sub_CM_Bind(lLength, .Item("Length"))
        Call sub_CM_Bind(lType, .Item("Type"))
        Call sub_CM_Bind(vAdd, .Item("Additional"))
    End With
    Dim sPw : sPw = func_CM_UtilGenerateRandomString(lLength, lType, vAdd)
    
    '★ログ出力
    Call sub_GnrtPwLogger(Array(3, "sub_GnrtPwGenerate", "GeneratedPassword is " & sPw))
    
    'ダイアログのメッセージなどを作成
    Dim sMsg, sTitle
    sMsg = "パスワードを生成しました" & vbNewLine & "OKボタンを押下するとクリップボードにコピーします"
    sTitle = new_clsCalGetNow() & " に作成"
    
    '★ログ出力
    Call sub_GnrtPwLogger(Array(3, "sub_GnrtPwGenerate", "Display Inputbox."))
    '一時ファイルのパスを作成
    Dim sPath : sPath = func_CM_FsGetTempFilePath()
    Do Until Inputbox(sMsg, sTitle, sPw)=False
        '一時ファイルに生成したパスワードを出力
        Call sub_CM_FsWriteFile(sPath, sPw)
        'クリップボードに一時ファイルの内容を出力
        Call CreateObject("Wscript.Shell").Run("cmd /c clip <""" & sPath & """", 0, True)
        '一時ファイルを削除
        Call func_CM_FsDeleteFile(sPath)
        '★ログ出力
        Call sub_GnrtPwLogger(Array(3, "sub_GnrtPwGenerate", "Copied to clipboard."))
    Loop
    
    
End Sub

'***************************************************************************************************
'Processing Order            : -
'Function/Sub Name           : sub_GnrtPwLogger()
'Overview                    : ログ出力する
'Detailed Description        : sub_CM_UtilLogger()に委譲する
'Argument
'     avParams               : 配列型のパラメータリスト
'Return Value
'     なし
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/03         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub sub_GnrtPwLogger( _
    byRef avParams _
    )
    Call sub_CM_UtilLogger(avParams, PoWriter)
End Sub
