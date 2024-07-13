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

'変数
Private PoWriter, PoBroker

'lib import
Private Const Cs_FOLDER_LIB = "lib"
With CreateObject("Scripting.FileSystemObject")
    Dim sParentFolderPath : sParentFolderPath = .GetParentFolderName(WScript.ScriptFullName)
    Dim sLibFolderPath : sLibFolderPath = .BuildPath(sParentFolderPath, Cs_FOLDER_LIB)
    Dim oLibFile
    For Each oLibFile In CreateObject("Shell.Application").Namespace(sLibFolderPath).Items
        If Not oLibFile.IsFolder Then
            If StrComp(.GetExtensionName(oLibFile.Path), "vbs", vbTextCompare)=0 Then ExecuteGlobal .OpenTextfile(oLibFile.Path).ReadAll
        End If
    Next
End With
Set oLibFile = Nothing

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
'2016/08/10         Y.Fujii                  First edition
'***************************************************************************************************
Sub Main()
    'ログ出力の設定
    Set PoWriter = new_WriterTo(fw_getLogPath, 8, True, -1)
    'ブローカークラスのインスタンスの設定
    Dim oBroker : Set oBroker = new_Broker()
    oBroker.subscribe publishType.LOG, GetRef("sub_GetPathsLogger")
    'パラメータ格納用オブジェクト宣言
    Dim oParams : Set oParams = new_Dic()
    
    '当スクリプトの引数をパラメータ格納用オブジェクトに取得する
    fw_excuteSub "sub_GetPathsGetParameters", oParams, oBroker
    
    '引数のファイルパスをクリップボードに出力する
    fw_excuteSub "sub_GetPathsProc", oParams, oBroker
    
    'ログ出力をクローズ
    PoWriter.close()
    
    'オブジェクトを開放
    Set oParams = Nothing
    Set oBroker = Nothing
    Set PoWriter = Nothing
End Sub

'***************************************************************************************************
'Processing Order            : 1
'Function/Sub Name           : sub_GetPathsGetParameters()
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
Private Sub sub_GetPathsGetParameters( _
    byRef aoParams _
    )
    'オリジナルの引数を取得
    Dim oArg : Set oArg = fw_storeArguments()
    '★ログ出力
    sub_GetPathsLogger Array(logType.DETAIL_INFO, "sub_GetPathsGetParameters", cf_toString(oArg))
    
    'パラメータ格納用オブジェクトに設定
    cf_bindAt aoParams, "Param", new_ArrWith(oArg.Item("Unnamed")).slice(0,vbNullString)
    
    Set oArg = Nothing
End Sub

'***************************************************************************************************
'Processing Order            : 2
'Function/Sub Name           : sub_GetPathsProc()
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
Private Sub sub_GetPathsProc( _
    byRef aoParams _
    )
    'パラメータ格納用汎用オブジェクト
    Dim oParam : Set oParam = aoParams.Item("Param")
    
    '一時ファイルに連結した引数を出力
    Dim sTempFilePaths : sTempFilePaths = fw_getTempPath()
    fs_writeFileDefault sTempFilePaths, func_GetPathsReplaceEnvironmentStrings(oParam.join(vbNewLine))
    fw_runShellSilently "cmd /c clip <" & fs_wrapInQuotes(sTempFilePaths)
    
    '一時ファイルを削除
    fs_deleteFile sTempFilePaths
    
    Set oParam = Nothing
End Sub

'***************************************************************************************************
'Processing Order            : 2-1
'Function/Sub Name           : func_GetPathsReplaceEnvironmentStrings()
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
Private Function func_GetPathsReplaceEnvironmentStrings( _
    byVal asStr _
    )
    Dim sSettings
    sSettings = Array("%UserProfile%")

    Dim sRet : sRet = asStr
    Dim i
    For Each i In sSettings
        sRet = Replace(sRet, new_Shell().ExpandEnvironmentStrings(i), i)
    Next

    func_GetPathsReplaceEnvironmentStrings = sRet
End Function

'***************************************************************************************************
'Processing Order            : -
'Function/Sub Name           : sub_GetPathsLogger()
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
Private Sub sub_GetPathsLogger( _
    byRef avParams _
    )
    fw_logger avParams, PoWriter
End Sub
