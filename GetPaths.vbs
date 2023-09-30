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
'2016/08/10         Y.Fujii                  First edition
'***************************************************************************************************
Sub Main()
    'ログ出力の設定
    Dim sPath : sPath = func_CM_FsGetPrivateLogFilePath()
    Set PoWriter = new_clsCmBufferedWriter(func_CM_FsOpenTextFile(sPath, 8, True, -2))
    '出版-購読型（Publish/subscribe）インスタンスの設定
    Set PoPubSub = new_clsCmPubSub()
    Call PoPubSub.Subscribe("log", GetRef("sub_GetPathsLogger"))
    'パラメータ格納用オブジェクト宣言
    Dim oParams : Set oParams = new_Dictionary()
    
    '当スクリプトの引数をパラメータ格納用オブジェクトに取得する
    Call sub_CM_ExcuteSub("sub_GetPathsGetParameters", oParams, PoPubSub, "log")
    
    '引数のファイルパスをクリップボードに出力する
    Call sub_CM_ExcuteSub("sub_GetPathsProc", oParams, PoPubSub, "log")
    
    'ファイル接続をクローズする
    Call PoWriter.FileClose()
    
    'オブジェクトを開放
    Set oParams = Nothing
    Set PoPubSub = Nothing
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
    Dim oArg : Set oArg = func_CM_UtilStoringArguments()
    '★ログ出力
    Call sub_GetPathsLogger(Array(9, "sub_GetPathsGetParameters", "Arguments are " & func_CM_ToStringArguments()))
    
    'パラメータ格納用オブジェクトに設定
    Call sub_CM_BindAt(aoParams, "Param", oArg.Item("Unnamed").Slice(0,vbNullString))
    
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
    Dim sTempFilePaths : sTempFilePaths = func_CM_FsGetTempFilePath() 
    Call sub_CM_FsWriteFile(sTempFilePaths, oParam.JoinVbs(vbNewLine))
    Call CreateObject("Wscript.Shell").Run("cmd /c clip <""" & sTempFilePaths & """", 0, True)
    
    '一時ファイルを削除
    Call func_CM_FsDeleteFile(sTempFilePaths)
    
    Set oParam = Nothing
End Sub

'***************************************************************************************************
'Processing Order            : -
'Function/Sub Name           : sub_GetPathsLogger()
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
'2023/09/30         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub sub_GetPathsLogger( _
    byRef avParams _
    )
    Call sub_CM_UtilLogger(avParams, PoWriter)
End Sub
