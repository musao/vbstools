'***************************************************************************************************
'FILENAME                    : GeneratePassword.vbs
'Overview                    : パスワードを生成する
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
Option Explicit

'定数
Private Const Cs_FOLDER_INCLUDE = "include"
Private PoWriter
Private PoPubSub

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
    'パラメータ格納用汎用ハッシュマップ宣言
    Dim oParams : Set oParams = new_Dictionary()
    
    '初期化
    Call sub_CM_ExcuteSub("sub_GnrtPwInitialize", oParams, PoPubSub, "log")
    
    '当スクリプトの引数取得（処理なし）
    Call sub_CM_ExcuteSub("sub_GnrtPwGetParameters", oParams, PoPubSub, "log")
    
    '比較対象ファイル入力画面の表示と取得
    Call sub_CM_ExcuteSub("sub_GnrtPwGenerate", oParams, PoPubSub, "log")
    
    '終了処理
    Call sub_CM_ExcuteSub("sub_GnrtPwTerminate", oParams, PoPubSub, "log")
    
    'ファイル接続をクローズする
    Call PoWriter.FileClose()
    
    'オブジェクトを開放
    Set oParams = Nothing
    Set PoPubSub = Nothing
    Set PoWriter = Nothing
End Sub

'***************************************************************************************************
'Processing Order            : 1
'Function/Sub Name           : sub_GnrtPwInitialize()
'Overview                    : 初期化
'Detailed Description        : 工事中
'Argument
'     aoParams               : パラメータ格納用汎用ハッシュマップ
'Return Value
'     なし
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/24         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub sub_GnrtPwInitialize( _
    byRef aoParams _
    )
End Sub

'***************************************************************************************************
'Processing Order            : 2
'Function/Sub Name           : sub_GnrtPwGetParameters()
'Overview                    : 当スクリプトの引数取得（処理なし）
'Detailed Description        : 工事中
'Argument
'     aoParams               : パラメータ格納用汎用ハッシュマップ
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
End Sub

'***************************************************************************************************
'Processing Order            : 3
'Function/Sub Name           : sub_GnrtPwGenerate()
'Overview                    : パスワード生成
'Detailed Description        : 工事中
'Argument
'     aoParams               : パラメータ格納用汎用ハッシュマップ
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
    '一時ファイルのパスを作成
    Dim sPath : sPath = func_CM_FsGetTempFilePath()
    
    'パスワード生成
    Dim sPw : sPw = func_CM_UtilGenerateRandomString(16, 15, Nothing)
    aoParams.Add "GeneratedPassword", sPw
    
    Dim sMsg, sTitle
    sMsg = "パスワードを生成しました" & vbNewLine & "OKボタンを押下するとクリップボードにコピーします"
    sTitle = new_clsCalGetNow() & " に作成"
    
    Do
        '一時ファイルに生成したパスワードを出力
        Call sub_CM_FsWriteFile(sPath, sPw)
        'クリップボードに一時ファイルの内容を出力
        Call CreateObject("Wscript.Shell").Run("cmd /c clip <""" & sPath & """", 0, True)
        '一時ファイルを削除
        Call func_CM_FsDeleteFile(sPath)
    Loop Until Inputbox(sMsg, sTitle, sPw)=False
    
    
End Sub

'***************************************************************************************************
'Processing Order            : 4
'Function/Sub Name           : sub_GnrtPwTerminate()
'Overview                    : 終了処理
'Detailed Description        : 工事中
'Argument
'     aoParams               : パラメータ格納用汎用ハッシュマップ
'Return Value
'     なし
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/03         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub sub_GnrtPwTerminate( _
    byRef aoParams _
    )
    PoWriter.Flush
End Sub

'***************************************************************************************************
'Processing Order            : -
'Function/Sub Name           : sub_GnrtPwLogger()
'Overview                    : ログ出力する
'Detailed Description        : 工事中
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
    Call sub_CM_UtilCommonLogger(avParams, PoWriter)
End Sub
