'***************************************************************************************************
'FILENAME                    : clsFsBaseTest.vbs
'Overview                    : ファイル・フォルダ共通クラスのテスト
'Detailed Description        : 工事中
'Argument
'     なし
'Return Value
'     なし
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/11/03         Y.Fujii                  First edition
'***************************************************************************************************
Option Explicit

'定数
Private Const Cs_FOLDER_LIB = "lib"
Private Const Cs_UTLIB_FILE = "VbsUtLib.vbs"
Private Const Cs_UTAST_FILE = "clsUtAssistant.vbs"
Private Const Cs_COMMON_FILE = "libCom.vbs"
Private Const Cs_TEST_FILE = "clsFsBase.vbs"

With CreateObject("Scripting.FileSystemObject")
    '単体テスト用ライブラリ読み込み
    Dim sIncludeFolderPath : sIncludeFolderPath = .BuildPath(.GetParentFolderName(WScript.ScriptFullName), Cs_FOLDER_LIB)
    ExecuteGlobal .OpenTextfile(.BuildPath(sIncludeFolderPath, Cs_UTLIB_FILE)).ReadAll
    ExecuteGlobal .OpenTextfile(.BuildPath(sIncludeFolderPath, Cs_UTAST_FILE)).ReadAll
    '共通ライブラリ読み込み
    sIncludeFolderPath = .BuildPath(.GetParentFolderName(.GetParentFolderName(WScript.ScriptFullName)), Cs_FOLDER_LIB)
    ExecuteGlobal .OpenTextfile(.BuildPath(sIncludeFolderPath, Cs_COMMON_FILE)).ReadAll
    '単体テスト対象ソース読み込み
    Dim sParentFolderPath : sParentFolderPath = .GetParentFolderName(.GetParentFolderName(WScript.ScriptFullName))
    sIncludeFolderPath = .BuildPath(sParentFolderPath, Cs_FOLDER_LIB)
    ExecuteGlobal .OpenTextfile(.BuildPath(sIncludeFolderPath, Cs_TEST_FILE)).ReadAll
End With

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
'2022/11/03         Y.Fujii                  First edition
'***************************************************************************************************
Sub Main()
    Dim oUtAssistant : Set oUtAssistant = New clsUtAssistant
    
    'ノーマルケースのテスト
    Call sub_clsFsBaseTest_1(oUtAssistant)
    
    '結果出力
    Call sub_UtResultOutput(oUtAssistant)
    
    Set oUtAssistant = Nothing
    
End Sub

'***************************************************************************************************
'Processing Order            : Last
'Function/Sub Name           : sub_OutputReport()
'Overview                    : 結果出力
'Detailed Description        : 工事中
'Argument
'     aoUtAssistant          : 単体テスト用アシスタントクラスのインスタンス
'Return Value
'     なし
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/11/03         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub sub_OutputReport( _
    byRef aoUtAssistant _
    )
    Call sub_UtWriteFile(func_UtGetThisLogFilePath(), aoUtAssistant.OutputReportInTsvFormat())
End Sub

'***************************************************************************************************
'Processing Order            : 1
'Function/Sub Name           : func_clsFsBaseTest_1()
'Overview                    : ノーマルケースのテスト
'Detailed Description        : 下記を汎用パターンとして指定する
'                              ・FileSystemObjectオブジェクトの設定有無それぞれについて検証する
'                              ・ファイル/フォルダそれぞれについて検証する
'Argument
'     aoUtAssistant          : 単体テスト用アシスタントクラスのインスタンス
'Return Value
'     なし
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/11/03         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub sub_clsFsBaseTest_1( _
    byRef aoUtAssistant _
    )
    'FileSystemObjectオブジェクトの設定有無パターン
    Dim boSetFsoFlgs : boSetFsoFlgs = Array(True, False)
    'ファイル/フォルダのパターン
    Dim boTargetIsFiles : boTargetIsFiles = Array(True, False)
    
    Call sub_clsFsBaseTest_1_1(aoUtAssistant, Array(boSetFsoFlgs, boTargetIsFiles))
    Call sub_clsFsBaseTest_1_2(aoUtAssistant, Array(boSetFsoFlgs, boTargetIsFiles))
    
End Sub

'***************************************************************************************************
'Processing Order            : 1-1
'Function/Sub Name           : sub_clsFsBaseTest_1_1()
'Overview                    : clsFsBaseの全属性、取得項目の確からしさを確認する
'Detailed Description        : 上位から指定されたケースの汎用パターン分実行指示する
'Argument
'     aoUtAssistant          : 単体テスト用アシスタントクラスのインスタンス
'     avGeneralPatterns      : ケースの汎用パターン（配列の配列）
'Return Value
'     なし
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/11/18         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub sub_clsFsBaseTest_1_1( _
    byRef aoUtAssistant _
    , byRef avGeneralPatterns _
    )
    Dim vIndividualPatterns : Dim oCreateArgumentFunc : Dim sCaseChildNum
    
    '1-1-1 全属性の値が正しいこと
    sCaseChildNum = "1"
    '全属性のパターン
    vIndividualPatterns = Array("Attributes", "DateCreated", "DateLastAccessed", "DateLastModified", _
                                "Drive", "Name", "ParentFolder", "Path", "ShortName", "ShortPath", _
                                "Size", "Type")
    Set oCreateArgumentFunc = GetRef("func_clsFsBaseTestCreateArgumentFor_1_1_" & sCaseChildNum & "_x")
    Call sub_clsFsBaseTest_1_1_x(aoUtAssistant, avGeneralPatterns, vIndividualPatterns, oCreateArgumentFunc, sCaseChildNum)
    
    '1-1-2 キャッシュ使用可否が変わっていないこと
    sCaseChildNum = "2"
    'キャッシュ使用可否のパターン
    vIndividualPatterns = Array(True, False)
    Set oCreateArgumentFunc = GetRef("func_clsFsBaseTestCreateArgumentFor_1_1_" & sCaseChildNum & "_x")
    Call sub_clsFsBaseTest_1_1_x(aoUtAssistant, avGeneralPatterns, vIndividualPatterns, oCreateArgumentFunc, sCaseChildNum)
    
    '1-1-3 キャッシュ有効期間（秒数）が変わっていないこと
    sCaseChildNum = "3"
    'キャッシュ有効期間（秒数）のパターン
    vIndividualPatterns = Array(0,1,2147483647,-1,-2147483648)
    Set oCreateArgumentFunc = GetRef("func_clsFsBaseTestCreateArgumentFor_1_1_" & sCaseChildNum & "_x")
    Call sub_clsFsBaseTest_1_1_x(aoUtAssistant, avGeneralPatterns, vIndividualPatterns, oCreateArgumentFunc, sCaseChildNum)
    
    '1-1-4 最終キャッシュ確認時間が変わっていないこと
    sCaseChildNum = "4"
    '全属性のパターン
    vIndividualPatterns = Array("Attributes", "DateCreated", "DateLastAccessed", "DateLastModified", _
                                "Drive", "Name", "ParentFolder", "Path", "ShortName", "ShortPath", _
                                "Size", "Type")
    Set oCreateArgumentFunc = GetRef("func_clsFsBaseTestCreateArgumentFor_1_1_" & sCaseChildNum & "_x")
    Call sub_clsFsBaseTest_1_1_x(aoUtAssistant, avGeneralPatterns, vIndividualPatterns, oCreateArgumentFunc, sCaseChildNum)
    
    '1-1-5 最終キャッシュ更新時間が変わっていること
    sCaseChildNum = "5"
    '全属性のパターン
    vIndividualPatterns = Array("Attributes", "DateCreated", "DateLastAccessed", "DateLastModified", _
                                "Drive", "Name", "ParentFolder", "Path", "ShortName", "ShortPath", _
                                "Size", "Type")
    Set oCreateArgumentFunc = GetRef("func_clsFsBaseTestCreateArgumentFor_1_1_" & sCaseChildNum & "_x")
    Call sub_clsFsBaseTest_1_1_x(aoUtAssistant, avGeneralPatterns, vIndividualPatterns, oCreateArgumentFunc, sCaseChildNum)
    
    Set oCreateArgumentFunc = Nothing
End Sub

'***************************************************************************************************
'Processing Order            : 1-1-1
'Function/Sub Name           : func_clsFsBaseTestCreateArgumentFor_1_1_1_x()
'Overview                    : sub_clsFsBaseTest_1_1_x()用の引数情報ハッシュマップを作成
'Detailed Description        : 処理はfunc_clsFsBaseTestCreateArgumentFor_1_1_x()に委譲する
'                              ケースの詳細
'                              上位から指定されたケースのパターン分下記を実行する
'                              実施条件
'                              ・キャッシュ使用可否は否
'                              ・キャッシュ有効期間は0秒
'                              ・全属性の値を1回取得
'                              期待値
'                              ・全属性の値が正しいこと
'Argument
'     avArguments            : ケースごとの引数のパターン
'Return Value
'     引数情報ハッシュマップ
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/12/02         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_clsFsBaseTestCreateArgumentFor_1_1_1_x( _
    byRef avArguments _
    )
    'サブタイトル名の作成
    Dim sSubTitle : sSubTitle = func_clsFsBaseTestCaseDescriptionFso(avArguments(0)) _
                              & "-" & func_clsFsBaseTestCaseDescriptionIsFile(avArguments(1)) _
                              & "-" & avArguments(2)
    
    'sub_clsFsBaseTest_1_1_1()用の引数情報ハッシュマップを作成
    Set func_clsFsBaseTestCreateArgumentFor_1_1_1_x = func_clsFsBaseTestCreateArgumentFor_1_1_x( _
                                                                                sSubTitle _
                                                                                , avArguments(1) _
                                                                                , False _
                                                                                , 0 _
                                                                                , avArguments(0) _
                                                                                , avArguments(2) _
                                                                                , False _
                                                                                , False _
                                                                                , vbNullString _
                                                                                , vbNullString _
                                                                                )
End Function

'***************************************************************************************************
'Processing Order            : 1-1-2
'Function/Sub Name           : func_clsFsBaseTestCreateArgumentFor_1_1_2_x()
'Overview                    : sub_clsFsBaseTest_1_1_x()用の引数情報ハッシュマップを作成
'Detailed Description        : 処理はfunc_clsFsBaseTestCreateArgumentFor_1_1_x()に委譲する
'                              ケースの詳細
'                              上位から指定されたケースのパターン分下記を実行する
'                              実施条件
'                              ・キャッシュ有効期間は0秒
'                              ・任意の属性の値を1回取得
'                              期待値
'                              ・キャッシュ使用可否が変わっていないこと
'Argument
'     avArguments            : ケースごとの引数のパターン
'Return Value
'     引数情報ハッシュマップ
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/12/02         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_clsFsBaseTestCreateArgumentFor_1_1_2_x( _
    byRef avArguments _
    )
    'サブタイトル名の作成
    Dim sUseCache : If avArguments(2) Then sUseCache="キャッシュ使用あり" Else sUseCache="キャッシュ使用なし"
    Dim sSubTitle : sSubTitle = func_clsFsBaseTestCaseDescriptionFso(avArguments(0)) _
                              & "-" & func_clsFsBaseTestCaseDescriptionIsFile(avArguments(1)) _
                              & "-" & sUseCache
    
    'sub_clsFsBaseTest_1_1_2()用の引数情報ハッシュマップを作成
    Set func_clsFsBaseTestCreateArgumentFor_1_1_2_x = func_clsFsBaseTestCreateArgumentFor_1_1_x( _
                                                                                sSubTitle _
                                                                                , avArguments(1) _
                                                                                , avArguments(2) _
                                                                                , 0 _
                                                                                , avArguments(0) _
                                                                                , "Attributes" _
                                                                                , True _
                                                                                , False _
                                                                                , vbNullString _
                                                                                , vbNullString _
                                                                                )
End Function

'***************************************************************************************************
'Processing Order            : 1-1-3
'Function/Sub Name           : func_clsFsBaseTestCreateArgumentFor_1_1_3_x()
'Overview                    : sub_clsFsBaseTest_1_1_x()用の引数情報ハッシュマップを作成
'Detailed Description        : 処理はfunc_clsFsBaseTestCreateArgumentFor_1_1_x()に委譲する
'                              ケースの詳細
'                              上位から指定されたケースのパターン分下記を実行する
'                              実施条件
'                              ・キャッシュ使用可否は可
'                              ・任意の属性の値を1回取得
'                              期待値
'                              ・キャッシュ有効期間（秒数）が変わっていないこと
'Argument
'     avArguments            : ケースごとの引数のパターン
'Return Value
'     引数情報ハッシュマップ
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/12/02         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_clsFsBaseTestCreateArgumentFor_1_1_3_x( _
    byRef avArguments _
    )
    'サブタイトル名の作成
    Dim sValidPeriod : Select Case avArguments(2)
        Case 0
            sValidPeriod = "キャッシュ有効期間がゼロ"
        Case 1
            sValidPeriod = "キャッシュ有効期間が１"
        Case 2147483647
            sValidPeriod = "キャッシュ有効期間が最大"
        Case -1
            sValidPeriod = "キャッシュ有効期間が－１"
        Case -2147483648
            sValidPeriod = "キャッシュ有効期間が最小"
    End Select
    Dim sSubTitle : sSubTitle = func_clsFsBaseTestCaseDescriptionFso(avArguments(0)) _
                              & "-" & func_clsFsBaseTestCaseDescriptionIsFile(avArguments(1)) _
                              & "-" & sValidPeriod
    
    'sub_clsFsBaseTest_1_1_3()用の引数情報ハッシュマップを作成
    Set func_clsFsBaseTestCreateArgumentFor_1_1_3_x = func_clsFsBaseTestCreateArgumentFor_1_1_x( _
                                                                                sSubTitle _
                                                                                , avArguments(1) _
                                                                                , True _
                                                                                , avArguments(2) _
                                                                                , avArguments(0) _
                                                                                , "Attributes" _
                                                                                , False _
                                                                                , True _
                                                                                , vbNullString _
                                                                                , vbNullString _
                                                                                )
End Function

'***************************************************************************************************
'Processing Order            : 1-1-4
'Function/Sub Name           : func_clsFsBaseTestCreateArgumentFor_1_1_4_x()
'Overview                    : sub_clsFsBaseTest_1_1_x()用の引数情報ハッシュマップを作成
'Detailed Description        : 処理はfunc_clsFsBaseTestCreateArgumentFor_1_1_x()に委譲する
'                              ケースの詳細
'                              上位から指定されたケースのパターン分下記を実行する
'                              実施条件
'                              ・キャッシュ使用可否は否
'                              ・キャッシュ有効期間は0秒
'                              ・全属性の値を1回取得
'                              期待値
'                              ・全属性の値が正しいこと
'                              ・最終キャッシュ確認時間が変わっていないこと
'Argument
'     avArguments            : ケースごとの引数のパターン
'Return Value
'     引数情報ハッシュマップ
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/01/07         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_clsFsBaseTestCreateArgumentFor_1_1_4_x( _
    byRef avArguments _
    )
    'サブタイトル名の作成
    Dim sSubTitle : sSubTitle = func_clsFsBaseTestCaseDescriptionFso(avArguments(0)) _
                              & "-" & func_clsFsBaseTestCaseDescriptionIsFile(avArguments(1)) _
                              & "-" & avArguments(2)
    
    'sub_clsFsBaseTest_1_1_1()用の引数情報ハッシュマップを作成
    Set func_clsFsBaseTestCreateArgumentFor_1_1_4_x = func_clsFsBaseTestCreateArgumentFor_1_1_x( _
                                                                                sSubTitle _
                                                                                , avArguments(1) _
                                                                                , False _
                                                                                , 0 _
                                                                                , avArguments(0) _
                                                                                , avArguments(2) _
                                                                                , False _
                                                                                , False _
                                                                                , True _
                                                                                , vbNullString _
                                                                                )
End Function

'***************************************************************************************************
'Processing Order            : 1-1-5
'Function/Sub Name           : func_clsFsBaseTestCreateArgumentFor_1_1_5_x()
'Overview                    : sub_clsFsBaseTest_1_1_x()用の引数情報ハッシュマップを作成
'Detailed Description        : 処理はfunc_clsFsBaseTestCreateArgumentFor_1_1_x()に委譲する
'                              ケースの詳細
'                              上位から指定されたケースのパターン分下記を実行する
'                              実施条件
'                              ・キャッシュ使用可否は否
'                              ・キャッシュ有効期間は0秒
'                              ・全属性の値を1回取得
'                              期待値
'                              ・全属性の値が正しいこと
'                              ・最終キャッシュ更新時間が変わっていること
'Argument
'     avArguments            : ケースごとの引数のパターン
'Return Value
'     引数情報ハッシュマップ
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/01/07         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_clsFsBaseTestCreateArgumentFor_1_1_5_x( _
    byRef avArguments _
    )
    'サブタイトル名の作成
    Dim sSubTitle : sSubTitle = func_clsFsBaseTestCaseDescriptionFso(avArguments(0)) _
                              & "-" & func_clsFsBaseTestCaseDescriptionIsFile(avArguments(1)) _
                              & "-" & avArguments(2)
    
    'sub_clsFsBaseTest_1_1_1()用の引数情報ハッシュマップを作成
    Set func_clsFsBaseTestCreateArgumentFor_1_1_5_x = func_clsFsBaseTestCreateArgumentFor_1_1_x( _
                                                                                sSubTitle _
                                                                                , avArguments(1) _
                                                                                , False _
                                                                                , 0 _
                                                                                , avArguments(0) _
                                                                                , avArguments(2) _
                                                                                , False _
                                                                                , False _
                                                                                , vbNullString _
                                                                                , False _
                                                                                )
End Function

'***************************************************************************************************
'Processing Order            : 1-1-x
'Function/Sub Name           : sub_clsFsBaseTest_1_1_x()
'Overview                    : ケース1-1-x汎用実行関数
'Detailed Description        : 汎用パターン＋個別パターン分の検証を行う
'                              ケースの詳細、引数情報ハッシュマップ作成は引数指定された関数に委譲する
'Argument
'     aoUtAssistant          : 単体テスト用アシスタントクラスのインスタンス
'     avGeneralPatterns      : ケースの汎用パターン（配列の配列）
'     avIndividualPatterns   : ケースの個別パターン（配列の配列）
'     aoCreateArgumentFunc   : 引数情報ハッシュマップ作成の関数
'     asCaseChildNum         : ケース子番号
'Return Value
'     なし
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/11/03         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub sub_clsFsBaseTest_1_1_x( _
    byRef aoUtAssistant _
    , byRef avGeneralPatterns _
    , byRef avIndividualPatterns _
    , byRef aoCreateArgumentFunc _
    , byVal asCaseChildNum _
    )
    'ケースの汎用パターンにケースの個別パターンを追加
    Dim vPatterns : vPatterns = avGeneralPatterns
    Call cf_push(vPatterns, avIndividualPatterns)
    
    '階層構造（配列の入れ子）のパターン情報格納用ハッシュマップ作成
    Dim vPatternInfos : vPatternInfos = func_clsFsBaseTestCreateaoHierarchicalPatterns( _
                                            vPatterns _
                                            , 0 _
                                            , aoCreateArgumentFunc _
                                            , vbNullString _
                                        )
    
    'ケース実行
    Call aoUtAssistant.RunWithMultiplePatterns( _
                                "func_clsFsBaseTest_1_1_" & asCaseChildNum & "_" _
                                , vPatternInfos _
                            )
    
End Sub

'***************************************************************************************************
'Processing Order            : 1-1-x
'Function/Sub Name           : func_clsFsBaseTestCreateArgumentFor_1_1_x()
'Overview                    : func_clsFsBaseTest_1_1_x()用の引数情報ハッシュマップを作成
'Detailed Description        : func_clsFsBaseTestCreateArgument()を参照
'Argument
'     asSubTitle             : ケースのサブ名称
'     aboTargetIsFile        : 対象はファイルか否か
'     aboUseCache            : キャッシュ使用可否
'     alValidPeriod          : キャッシュ有効期間（秒数）
'     boSetFsoFlg            : FileSystemObjectオブジェクトの設定有無
'     asPropName1            : 1回目に取得する属性名（2回目がない場合は値を検証する）
'     aboDontChgUc           : 最後にキャッシュ使用可否が変わっていないことを検証するか否か
'     aboDontChgVp           : 最後にキャッシュ有効期間（秒数）が変わっていないことを検証するか否か
'     aboIsUpdLcct           : 最終キャッシュ確認時間が最後の属性取得の直前から変わっているか否か
'     aboIsUpdLcut           : 最終キャッシュ更新時間が最後の属性取得の直前から変わっているか否か
'Return Value
'     引数情報ハッシュマップ
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/11/03         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_clsFsBaseTestCreateArgumentFor_1_1_x( _
    byVal asSubTitle _
    , byVal aboTargetIsFile _
    , byVal aboUseCache _
    , byVal alValidPeriod _
    , byVal boSetFsoFlg _
    , byVal asPropName1 _
    , byVal aboDontChgUc _
    , byVal aboDontChgVp _
    , byVal aboIsUpdLcct _
    , byVal aboIsUpdLcut _
    )
    Set func_clsFsBaseTestCreateArgumentFor_1_1_x = _
        func_clsFsBaseTestCreateArgument( _
            asSubTitle _
            , aboTargetIsFile _
            , aboUseCache _
            , alValidPeriod _
            , boSetFsoFlg _
            , False _
            , False _
            , 0 _
            , asPropName1 _
            , vbNullString _
            , aboDontChgUc _
            , aboDontChgVp _
            , aboIsUpdLcct _
            , aboIsUpdLcut _
            )
End Function

'***************************************************************************************************
'Processing Order            : 1-2
'Function/Sub Name           : sub_clsFsBaseTest_1_2()
'Overview                    : clsFsBaseのキャッシュの確からしさを確認する
'Detailed Description        : 上位から指定されたケースの汎用パターン分実行指示する
'Argument
'     aoUtAssistant          : 単体テスト用アシスタントクラスのインスタンス
'     avGeneralPatterns      : ケースの汎用パターン（配列の配列）
'Return Value
'     なし
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/12/05         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub sub_clsFsBaseTest_1_2( _
    byRef aoUtAssistant _
    , byRef avGeneralPatterns _
    )
    Dim vProps : Dim oCreateArgumentFunc : Dim sCaseChildNum
    '全属性のパターン
    vProps = Array("Attributes", "DateCreated", "DateLastAccessed", "DateLastModified", _
                                "Drive", "Name", "ParentFolder", "Path", "ShortName", "ShortPath", _
                                "Size", "Type")
    
    '1-2-1 同じ属性を2回取得する際、キャッシュ有効期間中はキャッシュを使用し属性の値が正しいこと
    sCaseChildNum = "1"
    Set oCreateArgumentFunc = GetRef("func_clsFsBaseTestCreateArgumentFor_1_2_" & sCaseChildNum & "_x")
    Call sub_clsFsBaseTest_1_2_x(aoUtAssistant, avGeneralPatterns, vProps, oCreateArgumentFunc, sCaseChildNum, True)
    
    '1-2-2 キャッシュがない属性を取得する際、キャッシュを使用ぜず属性の値が正しいこと
    sCaseChildNum = "2"
    Set oCreateArgumentFunc = GetRef("func_clsFsBaseTestCreateArgumentFor_1_2_" & sCaseChildNum & "_x")
    Call sub_clsFsBaseTest_1_2_x(aoUtAssistant, avGeneralPatterns, vProps, oCreateArgumentFunc, sCaseChildNum, False)
    
    Set oCreateArgumentFunc = Nothing
End Sub

'***************************************************************************************************
'Processing Order            : 1-2-1
'Function/Sub Name           : func_clsFsBaseTestCreateArgumentFor_1_2_1_x()
'Overview                    : sub_clsFsBaseTest_1_2_x()用の引数情報ハッシュマップを作成
'Detailed Description        : 処理はfunc_clsFsBaseTestCreateArgumentFor_1_2_x()に委譲する
'                              ケースの詳細
'                              上位から指定されたケースのパターン分下記を実行する
'                              実施条件
'                              ・キャッシュ使用可否は可
'                              ・キャッシュ有効期間は3600秒
'                              ・同じ属性の値を2回取得
'                              ・1回目の属性取得の直後にスリープする時間は10ミリ秒
'                              期待値
'                              ・属性の値が正しいこと
'                              ・キャッシュ確認あり（最終キャッシュ確認時間が1回目取得後から変わっていること）
'                              ・キャッシュ更新なし（最終キャッシュ更新時間が1回目取得後から変わっていないこと）
'Argument
'     avArguments            : ケースごとの引数のパターン
'Return Value
'     引数情報ハッシュマップ
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/01/02         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_clsFsBaseTestCreateArgumentFor_1_2_1_x( _
    byRef avArguments _
    )
    'サブタイトル名の作成
    Dim sSubTitle : sSubTitle = func_clsFsBaseTestCaseDescriptionFso(avArguments(0)) _
                              & "-" & func_clsFsBaseTestCaseDescriptionIsFile(avArguments(1)) _
                              & "-" & avArguments(2) _
                              & "-" & avArguments(3)
    
    'sub_clsFsBaseTest_1_2_1()用の引数情報ハッシュマップを作成
    Set func_clsFsBaseTestCreateArgumentFor_1_2_1_x = func_clsFsBaseTestCreateArgumentFor_1_2_x( _
                                                                                sSubTitle _
                                                                                , avArguments(1) _
                                                                                , True _
                                                                                , 3600 _
                                                                                , avArguments(0) _
                                                                                , False _
                                                                                , 10 _
                                                                                , avArguments(2) _
                                                                                , avArguments(3) _
                                                                                , False _
                                                                                , True _
                                                                                )
End Function

'***************************************************************************************************
'Processing Order            : 1-2-2
'Function/Sub Name           : func_clsFsBaseTestCreateArgumentFor_1_2_2_x()
'Overview                    : sub_clsFsBaseTest_1_2_x()用の引数情報ハッシュマップを作成
'Detailed Description        : 処理はfunc_clsFsBaseTestCreateArgumentFor_1_2_x()に委譲する
'                              ケースの詳細
'                              上位から指定されたケースのパターン分下記を実行する
'                              実施条件
'                              ・キャッシュ使用可否は可
'                              ・キャッシュ有効期間は3600秒
'                              ・1回目と異なる属性の値を取得
'                              ・1回目の属性取得の直後にスリープする時間は10ミリ秒
'                              期待値
'                              ・属性の値が正しいこと
'                              ・キャッシュ確認なし（最終キャッシュ確認時間が1回目取得後から変わっていないこと）
'                              ・キャッシュ更新あり（最終キャッシュ更新時間が1回目取得後から変わっていること）
'Argument
'     avArguments            : ケースごとの引数のパターン
'Return Value
'     引数情報ハッシュマップ
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/01/02         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_clsFsBaseTestCreateArgumentFor_1_2_2_x( _
    byRef avArguments _
    )
    'サブタイトル名の作成
    Dim sSubTitle : sSubTitle = func_clsFsBaseTestCaseDescriptionFso(avArguments(0)) _
                              & "-" & func_clsFsBaseTestCaseDescriptionIsFile(avArguments(1)) _
                              & "-" & avArguments(2) _
                              & "-" & avArguments(3)
    
    'sub_clsFsBaseTest_1_2_1()用の引数情報ハッシュマップを作成
    Set func_clsFsBaseTestCreateArgumentFor_1_2_2_x = func_clsFsBaseTestCreateArgumentFor_1_2_x( _
                                                                                sSubTitle _
                                                                                , avArguments(1) _
                                                                                , True _
                                                                                , 3600 _
                                                                                , avArguments(0) _
                                                                                , False _
                                                                                , 10 _
                                                                                , avArguments(2) _
                                                                                , avArguments(3) _
                                                                                , True _
                                                                                , False _
                                                                                )
End Function

'***************************************************************************************************
'Processing Order            : 1-2-x
'Function/Sub Name           : sub_clsFsBaseTest_1_2_x()
'Overview                    : ケース1-2-x汎用実行関数
'Detailed Description        : 汎用パターン分の検証を行う
'                              ケースの詳細、引数情報ハッシュマップ作成は引数指定された関数に委譲する
'Argument
'     aoUtAssistant          : 単体テスト用アシスタントクラスのインスタンス
'     avGeneralPatterns      : ケースの汎用パターン（配列の配列）
'     avProps                : 属性のパターン（配列の配列）
'     aboIsSameProp          : 2回目に参照する属性を1回目と同じとするか
'     asCaseChildNum         : ケース子番号
'     aboSamaPropRead        : 同じ属性の値を取得するか否か
'Return Value
'     なし
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/12/05         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub sub_clsFsBaseTest_1_2_x( _
    byRef aoUtAssistant _
    , byRef avGeneralPatterns _
    , byRef avProps _
    , byRef aoCreateArgumentFunc _
    , byVal asCaseChildNum _
    , byVal aboSamaPropRead _
    )
    'ケースの汎用パターンにケースの個別パターンを追加
    Dim vPatterns : vPatterns = avGeneralPatterns
    Call cf_push(vPatterns, avProps)
    Call cf_push(vPatterns, avProps)
    
    '階層構造（配列の入れ子）のパターン情報格納用ハッシュマップ作成
    Dim vPatternInfos : vPatternInfos = func_clsFsBaseTestCreateaoHierarchicalPatternsEx( _
                                            vPatterns _
                                            , 0 _
                                            , aoCreateArgumentFunc _
                                            , vbNullString _
                                            , aboSamaPropRead _
                                        )
    
    'ケース実行
    Call aoUtAssistant.RunWithMultiplePatterns( _
                                "func_clsFsBaseTest_1_2_" & asCaseChildNum & "_" _
                                , vPatternInfos _
                            )
    
End Sub

'***************************************************************************************************
'Processing Order            : 1-2-x
'Function/Sub Name           : func_clsFsBaseTestCreateArgumentFor_1_2_x()
'Overview                    : func_clsFsBaseTest_1_2_x()用の引数情報ハッシュマップを作成
'Detailed Description        : func_clsFsBaseTestCreateArgument()を参照
'Argument
'     asSubTitle             : ケースのサブ名称
'     aboTargetIsFile        : 対象はファイルか否か
'     aboUseCache            : キャッシュ使用可否
'     alValidPeriod          : キャッシュ有効期間（秒数）
'     boSetFsoFlg            : FileSystemObjectオブジェクトの設定有無
'     aboIsRecreate          : 2回目の属性取得の直前に対象ファイル/フォルダを再作成するか否か
'     alSleepMSecond         : 1回目の属性取得の直後にスリープする時間（ミリ秒）
'     asPropName1            : 1回目に取得する属性名（2回目がない場合は値を検証する）
'     asPropName2            : 2回目に取得する属性名、値を検証する
'     aboIsUpdLcct           : 最終キャッシュ確認時間が最後の属性取得の直前から変わっているか否か
'     aboIsUpdLcut           : 最終キャッシュ更新時間が最後の属性取得の直前から変わっているか否か
'Return Value
'     引数情報ハッシュマップ
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/01/02         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_clsFsBaseTestCreateArgumentFor_1_2_x( _
    byVal asSubTitle _
    , byVal aboTargetIsFile _
    , byVal aboUseCache _
    , byVal alValidPeriod _
    , byVal boSetFsoFlg _
    , byVal aboIsRecreate _
    , byVal alSleepMSecond _
    , byVal asPropName1 _
    , byVal asPropName2 _
    , byVal aboIsUpdLcct _
    , byVal aboIsUpdLcut _
    )
    Set func_clsFsBaseTestCreateArgumentFor_1_2_x = _
        func_clsFsBaseTestCreateArgument( _
            asSubTitle _
            , aboTargetIsFile _
            , aboUseCache _
            , alValidPeriod _
            , boSetFsoFlg _
            , True _
            , aboIsRecreate _
            , alSleepMSecond _
            , asPropName1 _
            , asPropName2 _
            , False _
            , False _
            , aboIsUpdLcct _
            , aboIsUpdLcut _
            )
End Function

'***************************************************************************************************
'Processing Order            : none
'Function/Sub Name           : func_clsFsBaseTestCreateArgument()
'Overview                    : ケースパターン情報格納用ハッシュマップに登録する引数情報ハッシュマップを作成する
'                              Key                      Value
'                              -------------------      --------------------------------------------
'                              "Conditions"             実施条件のハッシュマップ
'                              "Inspections"            検証内容のハッシュマップ
'
'                              実施条件のハッシュマップの内容
'                              Key                      Value
'                              -------------------      --------------------------------------------
'                              "TargetIsFile"           対象はファイルか否か
'                              "UseCache"               キャッシュ使用可否
'                              "ValidPeriod"            キャッシュ有効期間（秒数）
'                              "SetFsoFlg"              FileSystemObjectオブジェクトの設定有無
'                              "DoItTwice"              属性取得を2回するか否か
'                              "IsRecreate"             2回目の属性取得の直前に対象ファイル/フォルダを再作成するか否か
'                              "SleepMSecond"           1回目の属性取得の直後にスリープする時間（ミリ秒）
'
'                              検証内容のハッシュマップの内容
'                              Key                      Value
'                              -------------------      --------------------------------------------
'                              "PropName1"              1回目に取得する属性名（2回目がない場合は値を検証する）
'                              "PropName2"              2回目に取得する属性名、値を検証する
'                              "DontChgUc"              最後にキャッシュ使用可否が変わっていないことを検証するか否か
'                              "DontChgVp"              最後にキャッシュ有効期間（秒数）が変わっていないことを検証するか否か
'                              "IsUpdLcct"              最終キャッシュ確認時間が最後の属性取得の直前から変わっているか否か
'                              "IsUpdLcut"              最終キャッシュ更新時間が最後の属性取得の直前から変わっているか否か
'Argument
'     asSubTitle             : ケースのサブ名称
'     aboTargetIsFile        : 実施条件のハッシュマップの"TargetIsFile"と同じ
'     aboUseCache            : 実施条件のハッシュマップの"UseCache"と同じ
'     alValidPeriod          : 実施条件のハッシュマップの"ValidPeriod"と同じ
'     aboSetFsoFlg           : 実施条件のハッシュマップの"SetFsoFlg"と同じ
'     aboDoItTwice           : 実施条件のハッシュマップの"DoItTwice"と同じ
'     aboIsRecreate          : 実施条件のハッシュマップの"IsRecreate"と同じ
'     alSleepMSecond         : 実施条件のハッシュマップの"SleepMSecond"と同じ
'     asPropName1            : 検証内容のハッシュマップの"PropName1"と同じ
'     asPropName2            : 検証内容のハッシュマップの"PropName2"と同じ
'     aboDontChgUc           : 検証内容のハッシュマップの"DontChgUc"と同じ
'     aboDontChgVp           : 検証内容のハッシュマップの"DontChgVp"と同じ
'     aboIsUpdLcct           : 検証内容のハッシュマップの"IsUpdLcct"と同じ
'     aboIsUpdLcut           : 検証内容のハッシュマップの"IsUpdLcut"と同じ
'Return Value
'     引数情報ハッシュマップ
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/11/12         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_clsFsBaseTestCreateArgument( _
    byVal asSubTitle _
    , byVal aboTargetIsFile _
    , byVal aboUseCache _
    , byVal alValidPeriod _
    , byVal aboSetFsoFlg _
    , byVal aboDoItTwice _
    , byVal aboIsRecreate _
    , byVal alSleepMSecond _
    , byVal asPropName1 _
    , byVal asPropName2 _
    , byVal aboDontChgUc _
    , byVal aboDontChgVp _
    , byVal aboIsUpdLcct _
    , byVal aboIsUpdLcut _
    )
    Dim oConditions : Set oConditions = new_Dic()
    With oConditions
        .Add "TargetIsFile", aboTargetIsFile
        .Add "UseCache", aboUseCache
        .Add "ValidPeriod", alValidPeriod
        .Add "SetFsoFlg", aboSetFsoFlg
        .Add "DoItTwice", aboDoItTwice
        .Add "IsRecreate", aboIsRecreate
        .Add "SleepMSecond", alSleepMSecond
    End With
    
    Dim oInspections : Set oInspections = new_Dic()
    With oInspections
        .Add "PropName1", asPropName1
        .Add "PropName2", asPropName2
        .Add "DontChgUc", aboDontChgUc
        .Add "DontChgVp", aboDontChgVp
        .Add "IsUpdLcct", aboIsUpdLcct
        .Add "IsUpdLcut", aboIsUpdLcut
    End With
    
    Dim oArgument : Set oArgument = new_Dic()
    With oArgument
        .Add "SubTitle", asSubTitle
        .Add "Conditions", oConditions
        .Add "Inspections", oInspections
    End With
    
    Set func_clsFsBaseTestCreateArgument = oArgument
    Set oInspections = Nothing
    Set oConditions = Nothing
    Set oArgument = Nothing
End Function

'***************************************************************************************************
'Processing Order            : none
'Function/Sub Name           : func_clsFsBaseTestCreateaoHierarchicalPatterns()
'Overview                    : 階層構造（配列の入れ子）のパターン情報格納用ハッシュマップを作成する
'Detailed Description        : 引数のパターン（配列の配列）を網羅するパターン情報を作成する
'                              パターン情報の作成は引数のパターン情報格納用ハッシュマップを作成する
'                              関数に委譲する
'Argument
'     avHierarchicalPatterns : ケースのパターン（配列の配列）
'     alLayerNum             : 階層の位置（パターン（配列の配列）のインデックス）
'     aoFunc                 : 引数情報格納用ハッシュマップを作成する関数のポインタ
'     avFuncArguments        : 上記関数の引数、ケースごとの引数のパターン
'Return Value
'     階層構造（配列の入れ子）のパターン情報格納用ハッシュマップ
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/12/02         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_clsFsBaseTestCreateaoHierarchicalPatterns( _
    byRef avHierarchicalPatterns _
    , byVal alLayerNum _
    , byRef aoFunc _
    , byRef avFuncArguments _
    )
    Dim vArray : Dim vFuncArguments : Dim vItem
    For Each vItem In avHierarchicalPatterns(alLayerNum)
        '引数パターンの作成
        vFuncArguments = avFuncArguments
        Call cf_push(vFuncArguments, vItem)
        
        If Ubound(avHierarchicalPatterns)=alLayerNum Then
        'ケースのパターンの最下層の場合
        '引数情報格納用ハッシュマップを作成する関数の戻り値を配列に追加する
            Call cf_push(vArray, aoFunc(vFuncArguments))
        Else
        '最下層でない場合
        '一階層下（ケースのパターン配列の次）の情報を取得する、自身を再帰呼び出し
            Call cf_push(vArray, _
                func_clsFsBaseTestCreateaoHierarchicalPatterns(avHierarchicalPatterns, alLayerNum+1, aoFunc, vFuncArguments)_
                )
        End If
    Next
    func_clsFsBaseTestCreateaoHierarchicalPatterns = vArray
End Function

'***************************************************************************************************
'Processing Order            : none
'Function/Sub Name           : func_clsFsBaseTestCreateaoHierarchicalPatternsEx()
'Overview                    : func_clsFsBaseTestCreateaoHierarchicalPatterns()と同様
'                              最下層と最下層から2番目の要素の比較結果によって追加判定をする
'Detailed Description        : func_clsFsBaseTestCreateaoHierarchicalPatterns()と同様
'Argument
'     avHierarchicalPatterns : ケースのパターン（配列の配列）
'     alLayerNum             : 階層の位置（パターン（配列の配列）のインデックス）
'     aoFunc                 : 引数情報格納用ハッシュマップを作成する関数のポインタ
'     avFuncArguments        : 上記関数の引数、ケースごとの引数のパターン
'     aboCompareBottom       : 最下層と最下層から2番目の要素の比較結果
'Return Value
'     階層構造（配列の入れ子）のパターン情報格納用ハッシュマップ
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/01/02         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_clsFsBaseTestCreateaoHierarchicalPatternsEx( _
    byRef avHierarchicalPatterns _
    , byVal alLayerNum _
    , byRef aoFunc _
    , byRef avFuncArguments _
    , byRef aboCompareBottom _
    )
    Dim vArray : Dim vFuncArguments : Dim vItem
    For Each vItem In avHierarchicalPatterns(alLayerNum)
        '引数パターンの作成
        vFuncArguments = avFuncArguments
        Call cf_push(vFuncArguments, vItem)
        
        If Ubound(avHierarchicalPatterns)=alLayerNum Then
        'ケースのパターンの最下層の場合
            If aboCompareBottom = (vItem = avFuncArguments(Ubound(avFuncArguments))) Then
            '最下層と最下層から2番目の要素の比較結果が引数（aboCompareBottom）と等しい場合
                '引数情報格納用ハッシュマップを作成する関数の戻り値を配列に追加する
                Call cf_push(vArray, aoFunc(vFuncArguments))
            End If
        Else
        '最下層でない場合
        '一階層下（ケースのパターン配列の次）の情報を取得する、自身を再帰呼び出し
            Call cf_push(vArray, _
                func_clsFsBaseTestCreateaoHierarchicalPatternsEx(avHierarchicalPatterns, alLayerNum+1, aoFunc, vFuncArguments, aboCompareBottom)_
                )
        End If
    Next
    func_clsFsBaseTestCreateaoHierarchicalPatternsEx = vArray
End Function

'***************************************************************************************************
'Processing Order            : none
'Function/Sub Name           : 割愛
'Overview                    : ケースごとにノーマルケース汎用実行に委譲する関数
'Detailed Description        : func_clsFsBaseTestCreateArgumentFor_x_x()を参照
'Argument
'     aoArgument             : 引数情報のハッシュマップ
'Return Value
'     結果 True,Flase
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/11/03         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_clsFsBaseTest_1_1_1_( _
    byRef aoArgument _
    )
    func_clsFsBaseTest_1_1_1_ = func_clsFsBaseTestNormalBase(aoArgument)
End Function
Private Function func_clsFsBaseTest_1_1_2_( _
    byRef aoArgument _
    )
    func_clsFsBaseTest_1_1_2_ = func_clsFsBaseTestNormalBase(aoArgument)
End Function
Private Function func_clsFsBaseTest_1_1_3_( _
    byRef aoArgument _
    )
    func_clsFsBaseTest_1_1_3_ = func_clsFsBaseTestNormalBase(aoArgument)
End Function
Private Function func_clsFsBaseTest_1_1_4_( _
    byRef aoArgument _
    )
    func_clsFsBaseTest_1_1_4_ = func_clsFsBaseTestNormalBase(aoArgument)
End Function
Private Function func_clsFsBaseTest_1_1_5_( _
    byRef aoArgument _
    )
    func_clsFsBaseTest_1_1_5_ = func_clsFsBaseTestNormalBase(aoArgument)
End Function
Private Function func_clsFsBaseTest_1_2_1_( _
    byRef aoArgument _
    )
    func_clsFsBaseTest_1_2_1_ = func_clsFsBaseTestNormalBase(aoArgument)
End Function
Private Function func_clsFsBaseTest_1_2_2_( _
    byRef aoArgument _
    )
    func_clsFsBaseTest_1_2_2_ = func_clsFsBaseTestNormalBase(aoArgument)
End Function

'***************************************************************************************************
'Processing Order            : none
'Function/Sub Name           : func_clsFsBaseTestNormalBase()
'Overview                    : ノーマルケース汎用実行
'Detailed Description        : 引数情報ハッシュマップの構成はfunc_clsFsBaseTestCreateArgument()を参照
'Argument
'     aoArgument             : 引数情報のハッシュマップ
'Return Value
'     結果 True,Flase
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/11/03         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_clsFsBaseTestNormalBase( _
    byRef aoArgument _
    )
    '引数情報の取得
    With aoArgument.Item("Conditions")
    '実施条件
        Dim boTargetIsFile : boTargetIsFile = .Item("TargetIsFile")
        Dim boUseCache : boUseCache = .Item("UseCache")
        Dim dbValidPeriod : dbValidPeriod = .Item("ValidPeriod")
        Dim boSetFsoFlg : boSetFsoFlg = .Item("SetFsoFlg")
        Dim boDoItTwice : boDoItTwice = .Item("DoItTwice")
        Dim boIsRecreate : boIsRecreate = .Item("IsRecreate")
        Dim lSleepMSecond : lSleepMSecond = .Item("SleepMSecond")
    End With
    With aoArgument.Item("Inspections")
    '検証内容
        Dim sPropName1 : sPropName1 = .Item("PropName1")
        Dim sPropName2 : sPropName2 = .Item("PropName2")
        Dim boDontChgUc : boDontChgUc = .Item("DontChgUc")
        Dim boDontChgVp : boDontChgVp = .Item("DontChgVp")
        Dim boIsUpdLcct : boIsUpdLcct = .Item("IsUpdLcct")
        Dim boIsUpdLcut : boIsUpdLcut = .Item("IsUpdLcut")
    End With
    
    '前処理 一時ファイル/フォルダ作成
    '期待値は属性取得が1回か、2回目の属性取得直前に一時ファイル/フォルダを再作成しない場合に取得する
    Dim oExpect
    Dim boResult : boResult = True
    Dim sPath : sPath = func_UtGetThisTempFilePath()
    If boTargetIsFile Then
        Call CreateObject("Scripting.FileSystemObject").CreateTextFile(sPath)
        If Not(func_CM_FsFileExists(sPath)) Then Exit Function
        If Not (boDoItTwice And boIsRecreate) Then Set oExpect = func_clsFsBaseTestGetExpectedValue(func_CM_FsGetFile(sPath))
    Else
        Call func_CM_FsCreateFolder(sPath)
        If Not(func_CM_FsFolderExists(sPath)) Then Exit Function
        If Not (boDoItTwice And boIsRecreate) Then Set oExpect = func_clsFsBaseTestGetExpectedValue(func_CM_FsGetFolder(sPath))
    End If
    
    With New clsFsBase
        'テスト対象クラスに条件を指定
        .UseCache = boUseCache
        .ValidPeriod = dbValidPeriod
        .Path = sPath
        If boSetFsoFlg Then .Fso = CreateObject("Scripting.FileSystemObject")
        
        '属性取得（1回目）
        Dim lLastCacheConfirmationTime : lLastCacheConfirmationTime = .LastCacheConfirmationTime
        Dim lLastCacheUpdateTime : lLastCacheUpdateTime = .LastCacheUpdateTime
        Dim vProp : Call cf_bind(vProp, .Prop(sPropName1))
'        Dim vProp : Call sub_CM_TransferBetweenVariables(.Prop(sPropName1), vProp)
        
        If Not(boDoItTwice) Then
        '属性取得が1回の場合
            '属性の値を検証
            boResult = func_CM_IsSame(vProp, oExpect.Item(sPropName1))
            'キャッシュ使用可否が変わっていないことの検証
            If boDontChgUc Then boResult = (boUseCache = .UseCache)
            'キャッシュ有効期間（秒数）が変わっていないことの検証
            If boDontChgVp Then boResult = (dbValidPeriod = .ValidPeriod)
            '最終キャッシュ確認時間が最後の属性取得の直前から変わっているか
            If boIsUpdLcct<>vbNullString Then
                boResult = (boIsUpdLcct = (.LastCacheConfirmationTime=lLastCacheConfirmationTime))
            End If
            '最終キャッシュ更新時間が最後の属性取得の直前から変わっているか
            If boIsUpdLcut<>vbNullString Then
                boResult = (boIsUpdLcut = (.LastCacheUpdateTime=lLastCacheUpdateTime))
            End If
            
            '後処理 一時ファイル/フォルダ削除
            If boTargetIsFile Then Call func_CM_FsDeleteFile(sPath) Else Call func_CM_FsDeleteFolder(sPath)
            Set oExpect = Nothing
            
            '結果返却
            func_clsFsBaseTestNormalBase = boResult
            Exit Function
        End If
        '以降属性取得が2回の場合
        
        'スリープ
        WScript.Sleep lSleepMSecond
        
        If boIsRecreate Then
        '2回目の属性取得の直前に対象ファイル/フォルダを再作成する
            '一時ファイル/フォルダ削除
            If boTargetIsFile Then Call func_CM_FsDeleteFile(sPath) Else Call func_CM_FsDeleteFolder(sPath)
            '一時ファイル/フォルダ作成、期待値の取得
            If boTargetIsFile Then
                Call CreateObject("Scripting.FileSystemObject").CreateTextFile(sPath)
                If Not(func_CM_FsFileExists(sPath)) Then Exit Function
                Set oExpect = func_clsFsBaseTestGetExpectedValue(func_CM_FsGetFile(sPath))
            Else
                Call func_CM_FsCreateFolder(sPath)
                If Not(func_CM_FsFolderExists(sPath)) Then Exit Function
                Set oExpect = func_clsFsBaseTestGetExpectedValue(func_CM_FsGetFolder(sPath))
            End If
        End If
        
        '属性取得（2回目）
        lLastCacheConfirmationTime = .LastCacheConfirmationTime
        lLastCacheUpdateTime = .LastCacheUpdateTime
        Call cf_bind(vProp, .Prop(sPropName2))
'        Call sub_CM_TransferBetweenVariables(.Prop(sPropName2), vProp)
        
        '属性の値を検証
        boResult = func_CM_IsSame(vProp, oExpect.Item(sPropName2))
        'キャッシュ使用可否が変わっていないことの検証
        If boDontChgUc Then boResult = (boUseCache = .UseCache)
        'キャッシュ有効期間（秒数）が変わっていないことの検証
        If boDontChgVp Then boResult = (dbValidPeriod = .ValidPeriod)
        '最終キャッシュ確認時間が最後の属性取得の直前から変わっているか
        If boIsUpdLcct<>vbNullString Then
            boResult = (boIsUpdLcct = (.LastCacheConfirmationTime=lLastCacheConfirmationTime))
        End If
        '最終キャッシュ更新時間が最後の属性取得の直前から変わっているか
        If boIsUpdLcut<>vbNullString Then
            boResult = (boIsUpdLcut = (.LastCacheUpdateTime=lLastCacheUpdateTime))
        End If
        
    End With
    
    '後処理 一時ファイル/フォルダ削除
    If boTargetIsFile Then Call func_CM_FsDeleteFile(sPath) Else Call func_CM_FsDeleteFolder(sPath)
    Set oExpect = Nothing
    
    '結果返却
    func_clsFsBaseTestNormalBase = boResult
End Function

'***************************************************************************************************
'Processing Order            : none
'Function/Sub Name           : func_clsFsBaseTestGetExpectedValue()
'Overview                    : 期待値の取得
'Detailed Description        : 工事中
'Argument
'     aoSomeObject           : File/Folderオブジェクト
'Return Value
'     期待値のハッシュマップ
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/11/03         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_clsFsBaseTestGetExpectedValue( _
    byRef aoSomeObject _
    )
    
    Dim oExpect : Set oExpect = new_Dic()
    With aoSomeObject
        oExpect.Add "Attributes", .Attributes
        oExpect.Add "DateCreated", .DateCreated
        oExpect.Add "DateLastAccessed", .DateLastAccessed
        oExpect.Add "DateLastModified", .DateLastModified
        oExpect.Add "Drive", .Drive
        oExpect.Add "Name", .Name
        oExpect.Add "ParentFolder", .ParentFolder
        oExpect.Add "Path", .Path
        oExpect.Add "ShortName", .ShortName
        oExpect.Add "ShortPath", .ShortPath
        oExpect.Add "Size", .Size
        oExpect.Add "Type", .Type
    End With
    
    Set func_clsFsBaseTestGetExpectedValue = oExpect
    Set oExpect = Nothing
    
End Function

'***************************************************************************************************
'Processing Order            : none
'Function/Sub Name           : func_clsFsBaseTestCaseDescriptionFso()
'Overview                    : FSO有無のケース説明
'Detailed Description        : 工事中
'Argument
'     avKey                  : キー
'Return Value
'     引数情報ハッシュマップ
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/12/04         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_clsFsBaseTestCaseDescriptionFso( _
    byVal avKey _
    )
    Dim sDescription: If avKey Then sDescription="FSOあり" Else sDescription="FSOなし"
    func_clsFsBaseTestCaseDescriptionFso = sDescription
End Function

'***************************************************************************************************
'Processing Order            : none
'Function/Sub Name           : func_clsFsBaseTestCaseDescriptionIsFile()
'Overview                    : ファイル/フォルダのケース説明
'Detailed Description        : 工事中
'Argument
'     avKey                  : キー
'Return Value
'     引数情報ハッシュマップ
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/12/04         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_clsFsBaseTestCaseDescriptionIsFile( _
    byVal avKey _
    )
    Dim sDescription: If avKey Then sDescription="ファイル" Else sDescription="フォルダ"
    func_clsFsBaseTestCaseDescriptionIsFile = sDescription
End Function







''***************************************************************************************************
''Processing Order            : 1-2
''Function/Sub Name           : func_clsFsBaseTest_1_2()
''Overview                    : 各プロパティの値の取得の正当性（2回目、キャッシュ無効）
''Detailed Description        : 実施条件
''                              ・キャッシュ使用可否は否
''                              ・キャッシュ有効期間は3600秒
''                              ・全プロパティの値を2回取得
''                              期待値
''                              ・2回目に取得した全プロパティの値が正しいこと
''                              ・キャッシュ使用可否、同有効期間が変わらないこと
''                              ・キャッシュ確認なし（最終キャッシュ確認時間が1回目取得後から変わっていないこと）
''                              ・キャッシュ使用なし（最終キャッシュ更新時間が1回目取得後から変わっていること）
''Argument
''     aoUtAssistant          : 単体テスト用アシスタントクラスのインスタンス
''Return Value
''     結果 True,Flase
''---------------------------------------------------------------------------------------------------
''Histroy
''Date               Name                     Reason for Changes
''----------         ----------------------   -------------------------------------------------------
''2022/11/03         Y.Fujii                  First edition
''***************************************************************************************************
'Private Function func_clsFsBaseTest_1_2( _
'    )
'    Dim boResult : boResult = True
'    
'    '実施条件
'    Dim boUseCache : boUseCache = False
'    Dim dbValidPeriod : dbValidPeriod = 3600
'    
'    'テスト対象
'    Dim oSut : Set oSut = New clsFsBase
'    With oSut
'        '一時ファイル作成、期待値取得
'        Dim sPath : sPath = func_UtGetThisTempFilePath()
'        Call CreateObject("Scripting.FileSystemObject").CreateTextFile(sPath)
'        If Not(func_CM_FsFileExists(sPath)) Then Exit Function
'        Dim oExpect : Set oExpect = func_clsFsBaseTestGetExpectedValue(func_CM_FsGetFile(sPath))
'        
'        'テスト対象クラスに条件を指定
'        .UseCache = boUseCache
'        .ValidPeriod = dbValidPeriod
'        .Path = sPath
'        
'        '全プロパティの値を取得（1回目）
'        Call func_clsFsBaseTestValidateAllItems(oSut, oExpect)
'        Dim lLastCacheConfirmationTime : lLastCacheConfirmationTime = .LastCacheConfirmationTime
'        Dim lLastCacheUpdateTime : lLastCacheUpdateTime = .LastCacheUpdateTime
'        
'        '10msスリープ
'        WScript.Sleep 10
'        
'        '全プロパティの値を取得（2回目）
'        boResult = func_clsFsBaseTestValidateAllItems(oSut, oExpect)
'        
'        '検証
'        If .UseCache <> boUseCache Then boResult = False
'        If .ValidPeriod <> dbValidPeriod Then boResult = False
'        If .LastCacheConfirmationTime <> lLastCacheConfirmationTime Then boResult = False
'        If .LastCacheUpdateTime = lLastCacheUpdateTime Then boResult = False
'        
'        '一時ファイル削除
'        Call func_CM_FsDeleteFile(sPath)
'    End With
'    
'    '実施結果
'    func_clsFsBaseTest_1_2 = boResult
'    Set oExpect = Nothing
'    Set oSut = Nothing
'End Function
'
''***************************************************************************************************
''Processing Order            : 1-3
''Function/Sub Name           : func_clsFsBaseTest_1_3()
''Overview                    : 各プロパティの値の取得の正当性（2回目、キャッシュ有効期間超過かつファイル更新なし）
''Detailed Description        : 実施条件
''                              ・キャッシュ使用可否は可
''                              ・キャッシュ有効期間は0秒
''                              ・全プロパティの値を2回取得
''                              ・1回目と2回目でファイルの最終更新日が変わっていない
''                              期待値
''                              ・2回目に取得した全プロパティの値が正しいこと
''                              ・キャッシュ使用可否、同有効期間が変わらないこと
''                              ・キャッシュ確認あり（最終キャッシュ確認時間が1回目取得後から変わっていること）
''                              ・キャッシュ使用あり（最終キャッシュ更新時間が1回目取得後から変わっていないこと）
''Argument
''     aoUtAssistant          : 単体テスト用アシスタントクラスのインスタンス
''Return Value
''     結果 True,Flase
''---------------------------------------------------------------------------------------------------
''Histroy
''Date               Name                     Reason for Changes
''----------         ----------------------   -------------------------------------------------------
''2022/11/03         Y.Fujii                  First edition
''***************************************************************************************************
'Private Function func_clsFsBaseTest_1_3( _
'    )
'    Dim boResult : boResult = True
'    
'    '実施条件
'    Dim boUseCache : boUseCache = True
'    Dim dbValidPeriod : dbValidPeriod = 0
'    
'    'テスト対象
'    Dim oSut : Set oSut = New clsFsBase
'    With oSut
'        '一時ファイル作成、期待値取得
'        Dim sPath : sPath = func_UtGetThisTempFilePath()
'        Call CreateObject("Scripting.FileSystemObject").CreateTextFile(sPath)
'        If Not(func_CM_FsFileExists(sPath)) Then Exit Function
'        Dim oExpect : Set oExpect = func_clsFsBaseTestGetExpectedValue(func_CM_FsGetFile(sPath))
'        
'        'テスト対象クラスに条件を指定
'        .UseCache = boUseCache
'        .ValidPeriod = dbValidPeriod
'        .Path = sPath
'        
'        '全プロパティの値を取得（1回目）
'        Call func_clsFsBaseTestValidateAllItems(oSut, oExpect)
'        Dim lLastCacheConfirmationTime : lLastCacheConfirmationTime = .LastCacheConfirmationTime
'        Dim lLastCacheUpdateTime : lLastCacheUpdateTime = .LastCacheUpdateTime
'        
'        '10msスリープ
'        WScript.Sleep 10
'        
'        '全プロパティの値を取得（2回目）
'        boResult = func_clsFsBaseTestValidateAllItems(oSut, oExpect)
'        
'        '検証
'        If .UseCache <> boUseCache Then boResult = False
'        If .ValidPeriod <> dbValidPeriod Then boResult = False
'        If .LastCacheConfirmationTime = lLastCacheConfirmationTime Then boResult = False
'        If .LastCacheUpdateTime <> lLastCacheUpdateTime Then boResult = False
'        
'        '一時ファイル削除
'        Call func_CM_FsDeleteFile(sPath)
'    End With
'    
'    '実施結果
'    func_clsFsBaseTest_1_3 = boResult
'    Set oExpect = Nothing
'    Set oSut = Nothing
'End Function
'
''***************************************************************************************************
''Processing Order            : 1-4
''Function/Sub Name           : func_clsFsBaseTest_1_4()
''Overview                    : 各プロパティの値の取得の正当性（2回目、キャッシュ有効期間超過かつファイル更新あり）
''Detailed Description        : 実施条件
''                              ・キャッシュ使用可否は可
''                              ・キャッシュ有効期間は0秒
''                              ・全プロパティの値を2回取得
''                              ・1回目と2回目でファイルの最終更新日が変わっていない
''                              期待値
''                              ・2回目に取得した全プロパティの値が正しいこと
''                              ・キャッシュ使用可否、同有効期間が変わらないこと
''                              ・キャッシュ確認あり（最終キャッシュ確認時間が1回目取得後から変わっていること）
''                              ・キャッシュ使用なし（最終キャッシュ更新時間が1回目取得後から変わっていること）
''Argument
''     aoUtAssistant          : 単体テスト用アシスタントクラスのインスタンス
''Return Value
''     結果 True,Flase
''---------------------------------------------------------------------------------------------------
''Histroy
''Date               Name                     Reason for Changes
''----------         ----------------------   -------------------------------------------------------
''2022/11/03         Y.Fujii                  First edition
''***************************************************************************************************
'Private Function func_clsFsBaseTest_1_4( _
'    )
'    Dim boResult : boResult = True
'    
'    '実施条件
'    Dim boUseCache : boUseCache = True
'    Dim dbValidPeriod : dbValidPeriod = 0
'    
'    'テスト対象
'    Dim oSut : Set oSut = New clsFsBase
'    With oSut
'        '一時ファイル作成、期待値取得
'        Dim sPath : sPath = func_UtGetThisTempFilePath()
'        Call CreateObject("Scripting.FileSystemObject").CreateTextFile(sPath)
'        If Not(func_CM_FsFileExists(sPath)) Then Exit Function
'        Dim oExpect : Set oExpect = func_clsFsBaseTestGetExpectedValue(func_CM_FsGetFile(sPath))
'        
'        'テスト対象クラスに条件を指定
'        .UseCache = boUseCache
'        .ValidPeriod = dbValidPeriod
'        .Path = sPath
'        
'        '全プロパティの値を取得（1回目）
'        Call func_clsFsBaseTestValidateAllItems(oSut, oExpect)
'        Dim lLastCacheConfirmationTime : lLastCacheConfirmationTime = .LastCacheConfirmationTime
'        Dim lLastCacheUpdateTime : lLastCacheUpdateTime = .LastCacheUpdateTime
'        
'        '10msスリープ
'        WScript.Sleep 10
'        
'        '一時ファイル削除＆再作成、期待値の取得
'        Call func_CM_FsDeleteFile(sPath)
'        Call CreateObject("Scripting.FileSystemObject").CreateTextFile(sPath)
'        If Not(func_CM_FsFileExists(sPath)) Then Exit Function
'        oExpect.RemoveAll
'        Set oExpect = func_clsFsBaseTestGetExpectedValue(func_CM_FsGetFile(sPath))
'        
'        '全プロパティの値を取得（2回目）
'        boResult = func_clsFsBaseTestValidateAllItems(oSut, oExpect)
'        
'        '検証
'        If .UseCache <> boUseCache Then boResult = False
'        If .ValidPeriod <> dbValidPeriod Then boResult = False
'        If .LastCacheConfirmationTime = lLastCacheConfirmationTime Then boResult = False
'        If .LastCacheUpdateTime = lLastCacheUpdateTime Then boResult = False
'        
'        '一時ファイル削除
'        Call func_CM_FsDeleteFile(sPath)
'    End With
'    
'    '実施結果
'    func_clsFsBaseTest_1_4 = boResult
'    Set oExpect = Nothing
'    Set oSut = Nothing
'End Function
'
''***************************************************************************************************
''Processing Order            : none
''Function/Sub Name           : func_clsFsBaseTestValidateAllItems()
''Overview                    : 全項目の検証を行う
''Detailed Description        : 工事中
''Argument
''     aoSut                  : テスト対象クラス
''     aoExpect               : 期待値のハッシュマップ
''Return Value
''     結果 True,Flase
''---------------------------------------------------------------------------------------------------
''Histroy
''Date               Name                     Reason for Changes
''----------         ----------------------   -------------------------------------------------------
''2022/11/03         Y.Fujii                  First edition
''***************************************************************************************************
'Private Function func_clsFsBaseTestValidateAllItems( _
'    byRef aoSut _
'    , byRef aoExpect _
'    )
'    Dim boFlg : boFlg = True
'    
'    With aoExpect
'        Dim sKey
'        For Each sKey In .Keys
'            If IsObject(.Item(sKey)) Then
'                If Not (aoSut.Prop(sKey) Is .Item(sKey)) Then boFlg = False
'            Else
'                If aoSut.Prop(sKey) <> .Item(sKey) Then boFlg = False
'            End If
'        Next
'    End With
'    
'    func_clsFsBaseTestValidateAllItems = boFlg
'    
'End Function
