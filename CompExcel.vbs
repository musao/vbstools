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
    Set PoPubSub = Nothing
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
    Call PoWriter.FileClose()
    
    'オブジェクトを開放
    Set oParams = Nothing
    Set PoWriter = Nothing
    Set PoPubSub = Nothing
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
    'ログ出力の設定
    Dim sPath : sPath = func_CM_FsBuildPath( _
                    func_CM_FsGetPrivateFolder("log") _
                    , func_CM_FsGetGetBaseName(WScript.ScriptName) & new_clsCalGetNow().DisplayFormatAs("_YYMMDD_hhmmss.000.log") _
                    )
    Set PoWriter = new_clsCmBufferedWriter(func_CM_FsOpenTextFile(sPath, 8, True, -2))
    
    '出版-購読型（Publish/subscribe）インスタンスの設定
    Set PoPubSub = new_clsCmPubSub()
    Call PoPubSub.Subscribe("log", GetRef("sub_CmpExcelLogger"))
'    Call sub_CM_BindAt( aoParams, "PubSub", oPubSub)
    
End Sub

'***************************************************************************************************
'Processing Order            : 2
'Function/Sub Name           : sub_CmpExcelGetParameters()
'Overview                    : 当スクリプトの引数取得
'Detailed Description        : パラメータ格納用汎用ハッシュマップにKey="Parameter"で格納する
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
    'パラメータ格納用ハッシュマップ
    Dim oParameter : Set oParameter = new_Dictionary()
    
    Dim lCnt : lCnt = 0
    Dim sParam
    For Each sParam In WScript.Arguments
        If func_CM_FsFileExists(sParam) Then
        'ファイルが存在する場合パラメータを取得
            lCnt = lCnt + 1
            Call sub_CM_BindAt(oParameter, lCnt, sParam)
        End If
    Next
    
    Call sub_CM_BindAt(aoParams, "Parameter", oParameter)
    
    'オブジェクトを開放
    Set oParameter = Nothing
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
    Dim oParameter : Set oParameter = aoParams.Item("Parameter")
    
    Const Cs_TITLE_EXCEL = "比較対象ファイルを開く"
    
    If oParameter.Count > 1 Then
    'パラメータが2個以上だったら関数を抜ける
        Exit Sub
    End If
    
    With CreateObject("Excel.Application")
        Dim sPath
        Do Until oParameter.Count > 1
            
            sPath = .GetOpenFilename( , , Cs_TITLE_EXCEL, , False)
            If sPath = False Then
            'ファイル選択キャンセルの場合は当スクリプトを終了する
                Call sub_CM_ExcuteSub("sub_CmpExcelTerminate", aoParams, "log")
                Wscript.Quit
            End If
            If func_CM_FsFileExists(sPath) Then
            'ファイルが存在する場合パラメータを取得
                Call oParameter.Add(oParameter.Count+1, sPath)
            End If
        Loop
        
        .Quit
    End With
    
    'オブジェクトを開放
    Set oParameter = Nothing
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
    
    '4-1 比較するファイルを古い順（最終更新日昇順）に並べ替える
    Call sub_CM_ExcuteSub("sub_CmpExcelSortByDateLastModified", aoParams, PoPubSub, "log")
    
    '4-2 比較
    With New clsCompareExcel
'        Call sub_CM_Bind(.PubSub, aoParams.Item("PubSub"))
        .PathFrom = aoParams.Item("Parameter").Item(1)
        .PathTo = aoParams.Item("Parameter").Item(2)
        .Compare()
    End With

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
    Dim oParameter : Set oParameter = aoParams.Item("Parameter")
    
    With oParameter
        Dim oDateTimeA : Set oDateTimeA = new_clsCalSetDate(func_CM_FsGetFile(.Item(1)).DateLastModified)
        Dim oDateTimeB : Set oDateTimeB = new_clsCalSetDate(func_CM_FsGetFile(.Item(2)).DateLastModified)
        If oDateTimeA.CompareTo(oDateTimeB) <= 0 Then
        '最初のファイルの方が古い（最終更新日が小さい）場合、処理を抜ける
            Exit Sub
        End If
        
        '値を入れ替える
        Dim sValue1, sValue2
        sValue1 = .Item(1)
        sValue2 = .Item(2)
        
        .RemoveAll
        
        Call .Add(1, sValue2)
        Call .Add(2, sValue1)
    End With
    
    'オブジェクトを開放
    Set oParameter = Nothing
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
'    'ファイル接続をクローズする
'    Call PoWriter.FileClose()
'    Set PoWriter = Nothing
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
    With PoWriter
        Dim sCont : sCont = new_clsCalGetNow()
        sCont = sCont & vbTab & avParams(0)
        sCont = sCont & vbTab & avParams(1)
        sCont = sCont & vbTab & avParams(2)
        .WriteContents(sCont)
        .newLine()
    End With
End Sub
