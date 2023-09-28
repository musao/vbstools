'***************************************************************************************************
'FILENAME                    : CompExcel.vbs
'Overview                    : エクセルファイルを比較する
'Detailed Description        : 工事中
'Argument
'     PATH1                  : 比較するエクセルファイルのパス1
'     PATH2                  : 比較するエクセルファイルのパス2
'Return Value
'     なし
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2017/04/26         Y.Fujii                  First edition
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
'2017/04/26         Y.Fujii                  First edition
'***************************************************************************************************
Sub Main()
    'ログ出力の設定
    Dim sPath : sPath = func_CM_FsGetPrivateLogFilePath()
    Set PoWriter = new_clsCmBufferedWriter(func_CM_FsOpenTextFile(sPath, 8, True, -2))
    '出版-購読型（Publish/subscribe）インスタンスの設定
    Set PoPubSub = new_clsCmPubSub()
    Call PoPubSub.Subscribe("log", GetRef("sub_CmpExcelLogger"))
    'パラメータ格納用汎用ハッシュマップ宣言
    Dim oParams : Set oParams = new_Dictionary()
    
    '初期化
    Call sub_CM_ExcuteSub("sub_CmpExcelInitialize", oParams, PoPubSub, "log")
    
    '当スクリプトの引数取得
    Call sub_CM_ExcuteSub("sub_CmpExcelGetParameters", oParams, PoPubSub, "log")
    
    '比較対象ファイル入力画面の表示と取得
    Call sub_CM_ExcuteSub("sub_CmpExcelDispInputFiles", oParams, PoPubSub, "log")
    
    'エクセルファイルを比較する
    Call sub_CM_ExcuteSub("sub_CmpExcelCompareFiles", oParams, PoPubSub, "log")
    
    '終了処理
    Call sub_CM_ExcuteSub("sub_CmpExcelTerminate", oParams, PoPubSub, "log")
    
    'ファイル接続をクローズする
    PoWriter.FileClose
    
    'オブジェクトを開放
    Set oParams = Nothing
    Set PoPubSub = Nothing
    Set PoWriter = Nothing
End Sub

'***************************************************************************************************
'Processing Order            : 1
'Function/Sub Name           : sub_CmpExcelInitialize()
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
'2023/09/03         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub sub_CmpExcelInitialize( _
    byRef aoParams _
    )
End Sub

'***************************************************************************************************
'Processing Order            : 2
'Function/Sub Name           : sub_CmpExcelGetParameters()
'Overview                    : 当スクリプトの引数取得
'Detailed Description        : パラメータ格納用汎用ハッシュマップにKey="Param"で格納する
'                              パラメータ格納用ハッシュマップの構成
'                              Key                      Value
'                              -------------------      --------------------------------------------
'                              Seq(1,2)                 比較するエクセルファイルのパス
'                              引数があり存在するファイルパスのみ取得する
'Argument
'     aoParams               : パラメータ格納用汎用ハッシュマップ
'Return Value
'     なし
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2017/04/26         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub sub_CmpExcelGetParameters( _
    byRef aoParams _
    )
    'オリジナルの引数を取得
    Dim oArg : Set oArg = func_CM_UtilStoringArguments()
    '★ログ出力
    Call sub_CmpExcelLogger(Array(9, "sub_CmpExcelGetParameters", "Arguments are " & func_CM_ToStringArguments()))
    
    'パラメータ格納用オブジェクトに設定
    Call sub_CM_BindAt(aoParams, "Param", oArg.Item("Unnamed").Slice(0,2))
    
    Set oArg = Nothing
End Sub

'***************************************************************************************************
'Processing Order            : 3
'Function/Sub Name           : sub_CmpExcelDispInputFiles()
'Overview                    : 比較対象ファイル入力画面の表示と取得
'Detailed Description        : パラメータ格納用汎用ハッシュマップにKey="Parameter"で格納する
'                              パラメータ格納用ハッシュマップの構成
'                              Key                      Value
'                              -------------------      --------------------------------------------
'                              Seq(1,2)                 比較するエクセルファイルのパス
'                              比較するエクセルファイルのパスが2つ未満の場合に不足分を取得格納する
'Argument
'     aoParams               : パラメータ格納用汎用ハッシュマップ
'Return Value
'     なし
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2017/04/26         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub sub_CmpExcelDispInputFiles( _
    byRef aoParams _
    )
    'パラメータ格納用汎用ハッシュマップ
    Dim oParam : Set oParam = aoParams.Item("Param")
    
    Const Cs_TITLE_EXCEL = "比較対象ファイルを開く"
    
    If oParam.Length > 1 Then
    'パラメータが2個以上だったら関数を抜ける
        Exit Sub
    End If
    
    With CreateObject("Excel.Application")
        Dim sPath
        Do Until oParam.Length > 1
            
            sPath = .GetOpenFilename( , , Cs_TITLE_EXCEL, , False)
            If sPath = False Then
            'ファイル選択キャンセルの場合は当スクリプトを終了する
                Call sub_CM_ExcuteSub("sub_CmpExcelTerminate", aoParams, PoPubSub, "log")
                PoWriter.FileClose
                Wscript.Quit
            End If
            If func_CM_FsFileExists(sPath) Then
            'ファイルが存在する場合パラメータを取得
                oParam.Push sPath
            End If
        Loop
        
        .Quit
    End With
    
    'オブジェクトを開放
    Set oParam = Nothing
End Sub

'***************************************************************************************************
'Processing Order            : 4
'Function/Sub Name           : sub_CmpExcelCompareFiles()
'Overview                    : エクセルファイルを比較する
'Detailed Description        : エラーは無視する
'Argument
'     aoParams               : パラメータ格納用汎用ハッシュマップ
'Return Value
'     なし
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2017/04/26         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub sub_CmpExcelCompareFiles( _
    byRef aoParams _
    )
    'パラメータ格納用汎用ハッシュマップ
    Dim oParam : Set oParam = aoParams.Item("Param")
    
    '4-1 比較するファイルを古い順（最終更新日昇順）に並べ替える
    Call sub_CM_ExcuteSub("sub_CmpExcelSortByDateLastModified", aoParams, PoPubSub, "log")
    
    '4-2 比較
    With New clsCompareExcel
        .PathFrom = oParam(0)
        .PathTo = oParam(1)
        .Compare()
    End With
    
    'オブジェクトを開放
    Set oParam = Nothing
End Sub

'***************************************************************************************************
'Processing Order            : 4-1
'Function/Sub Name           : sub_CmpExcelSortByDateLastModified()
'Overview                    : 比較するファイルを古い順（最終更新日昇順）に並べ替える
'Detailed Description        : 工事中
'Argument
'     aoParams               : パラメータ格納用汎用ハッシュマップ
'Return Value
'     なし
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2017/04/26         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub sub_CmpExcelSortByDateLastModified( _
    byRef aoParams _
    )
    'パラメータ格納用汎用ハッシュマップ
    Dim oParam : Set oParam = aoParams.Item("Param")
    
    Dim oDateTimeA : Set oDateTimeA = new_clsCalSetDate(func_CM_FsGetFile(oParam(0)).DateLastModified)
    Dim oDateTimeB : Set oDateTimeB = new_clsCalSetDate(func_CM_FsGetFile(oParam(1)).DateLastModified)
    If oDateTimeA.CompareTo(oDateTimeB) > 0 Then
    '最初のファイルの方が新しい（最終更新日が大きい）場合、順番を入れ替える
        oParam.Reverse
    End If
    
    'オブジェクトを開放
    Set oParam = Nothing
    Set oDateTimeA = Nothing
    Set oDateTimeB = Nothing
End Sub

'***************************************************************************************************
'Processing Order            : 5
'Function/Sub Name           : sub_CmpExcelTerminate()
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
Private Sub sub_CmpExcelTerminate( _
    byRef aoParams _
    )
    PoWriter.Flush
End Sub

'***************************************************************************************************
'Processing Order            : -
'Function/Sub Name           : sub_CmpExcelLogger()
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
Private Sub sub_CmpExcelLogger( _
    byRef avParams _
    )
    Call sub_CM_UtilCommonLogger(avParams, PoWriter)
End Sub
