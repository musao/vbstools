'***************************************************************************************************
'FILENAME                    : GetPaths.vbs
'Overview                    : 引数のファイルパスをクリップボードにコピーする
'Detailed Description        : Sendtoから使用する
'Argument
'     PATH1,2...             : ファイルのパス1,2,...
'Return Value
'     なし
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2016/08/10         Y.Fujii                  First edition
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
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2016/08/10         Y.Fujii                  First edition
'***************************************************************************************************
Sub Main()
    'パラメータ格納用オブジェクト宣言
    Dim oParams : Set oParams = new_Dic()
    
    '当スクリプトの引数をパラメータ格納用オブジェクトに取得する
    fw_excuteSub "this_getParameters", oParams, PoBroker
    
    '引数のファイルパスをクリップボードに出力する
    fw_excuteSub "this_toClipbord", oParams, PoBroker
    
    'オブジェクトを開放
    Set oParams = Nothing
End Sub

'***************************************************************************************************
'Processing Order            : 1
'Function/Sub Name           : this_getParameters()
'Overview                    : 当スクリプトの引数をパラメータ格納用オブジェクトに取得する
'Detailed Description        : パラメータ格納用汎用オブジェクトにKey="Param"で格納する
'                              配列（clsCmArray型）に名前なし引数（/Key:Value 形式でない）を全て
'                              取得する
'Argument
'     aoParams               : パラメータ格納用オブジェクト
'Return Value
'     なし
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/30         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub this_getParameters( _
    byRef aoParams _
    )
    'オリジナルの引数を取得
    Dim oArg : Set oArg = fw_storeArguments()
    '★ログ出力
    this_logger Array(logType.DETAIL, "this_getParameters()", cf_toString(oArg))
    
    'パラメータ格納用オブジェクトに設定
    cf_bindAt aoParams, "Param", new_ArrOf(oArg.Item("Unnamed")).slice(0,vbNullString)
    
    Set oArg = Nothing
End Sub

'***************************************************************************************************
'Processing Order            : 2
'Function/Sub Name           : this_toClipbord()
'Overview                    : 引数のファイルパスをクリップボードに出力する
'Detailed Description        : 工事中
'Argument
'     aoParams               : パラメータ格納用オブジェクト
'Return Value
'     なし
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/30         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub this_toClipbord( _
    byRef aoParams _
    )
    'パラメータ格納用汎用オブジェクト
    Dim oParam : Set oParam = aoParams.Item("Param")
    
    '一時ファイルに連結した引数を出力
    Dim sTempFilePaths : sTempFilePaths = fw_getTempPath()
    fs_writeFileDefault sTempFilePaths, this_replaceEnvironmentStrings(oParam.join(vbNewLine))
    fw_runShellSilently "cmd /c clip <" & fs_wrapInQuotes(sTempFilePaths)
    
    '一時ファイルを削除
    fs_deleteFile sTempFilePaths
    
    Set oParam = Nothing
End Sub

'***************************************************************************************************
'Processing Order            : 2-1
'Function/Sub Name           : this_replaceEnvironmentStrings()
'Overview                    : 環境変数に置き換える
'Detailed Description        : 工事中
'Argument
'     asStr                  : 対象
'Return Value
'     環境変数に置き換えた文字列
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2024/04/06         Y.Fujii                  First edition
'***************************************************************************************************
Private Function this_replaceEnvironmentStrings( _
    byVal asStr _
    )
    Dim sSettings
    sSettings = Array("%UserProfile%")

    Dim sRet : sRet = asStr
    Dim i
    For Each i In sSettings
        sRet = Replace(sRet, new_Shell().ExpandEnvironmentStrings(i), i)
    Next

    this_replaceEnvironmentStrings = sRet
End Function

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
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/30         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub this_logger( _
    byRef avParams _
    )
    fw_logger avParams, PoTs4Log
End Sub
