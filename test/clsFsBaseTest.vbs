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
Private Const Cs_FOLDER_INCLUDE = "include"
Private Const Cs_UTLIB_FILE = "VbsUtLib.vbs"
Private Const Cs_UTAST_FILE = "clsUtAssistant.vbs"
Private Const Cs_COMMON_FILE = "VbsBasicLibCommon.vbs"
Private Const Cs_TEST_FILE = "clsFsBase.vbs"

With CreateObject("Scripting.FileSystemObject")
    '単体テスト用ライブラリ読み込み
    Dim sIncludeFolderPath : sIncludeFolderPath = .BuildPath(.GetParentFolderName(WScript.ScriptFullName), Cs_FOLDER_INCLUDE)
    ExecuteGlobal .OpenTextfile(.BuildPath(sIncludeFolderPath, Cs_UTLIB_FILE)).ReadAll
    ExecuteGlobal .OpenTextfile(.BuildPath(sIncludeFolderPath, Cs_UTAST_FILE)).ReadAll
    '共通ライブラリ読み込み
    sIncludeFolderPath = .BuildPath(.GetParentFolderName(.GetParentFolderName(WScript.ScriptFullName)), Cs_FOLDER_INCLUDE)
    ExecuteGlobal .OpenTextfile(.BuildPath(sIncludeFolderPath, Cs_COMMON_FILE)).ReadAll
    '単体テスト対象ソース読み込み
    Dim sParentFolderPath : sParentFolderPath = .GetParentFolderName(.GetParentFolderName(WScript.ScriptFullName))
    sIncludeFolderPath = .BuildPath(sParentFolderPath, Cs_FOLDER_INCLUDE)
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
Private Sub sub_clsFsBaseTest_1( _
    byRef aoUtAssistant _
    )
    
    Call sub_clsFsBaseTest_1_1(aoUtAssistant)
    
End Sub

'***************************************************************************************************
'Processing Order            : 1-1
'Function/Sub Name           : func_clsFsBaseTest_1_1()
'Overview                    : clsFsBaseの全属性の確からしさを確認する
'Detailed Description        : 工事中
'Argument
'     aoUtAssistant          : 単体テスト用アシスタントクラスのインスタンス
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
    )
    
    Call sub_clsFsBaseTest_1_1_1(aoUtAssistant)
'    Call sub_clsFsBaseTest_1_1_2(aoUtAssistant)
'    Call sub_clsFsBaseTest_1_1_3(aoUtAssistant)
    
End Sub

'***************************************************************************************************
'Processing Order            : 1-1-1
'Function/Sub Name           : sub_clsFsBaseTest_1_1_1()
'Overview                    : 各属性の取得の正当性を確認する
'Detailed Description        : 実施条件
'                              ・FileSystemObjectオブジェクトの設定有無それぞれについて検証する
'                              ・ファイル/フォルダそれぞれについて検証する
'                              ・キャッシュ使用可否は否
'                              ・キャッシュ有効期間は0秒
'                              ・全属性の値を1回取得
'                              期待値
'                              ・全属性の値が正しいこと
'Argument
'     aoUtAssistant          : 単体テスト用アシスタントクラスのインスタンス
'Return Value
'     結果 True,Flase
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/11/03         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub sub_clsFsBaseTest_1_1_1( _
    byRef aoUtAssistant _
    )
    'FileSystemObjectオブジェクトの設定有無パターン
    Dim boSetFsoFlg : Dim boSetFsoFlgs
    boSetFsoFlgs = Array(True, False)
    'ファイル/フォルダのパターン
    Dim boTargetIsFile : Dim boTargetIsFiles
    boTargetIsFiles = Array(True, False)
    '各属性のパターン
    Dim sPropName : Dim sPropNames
    sPropNames = Array("Attributes", "DateCreated", "DateLastAccessed", "DateLastModified", "Drive", _
        "Name", "ParentFolder", "Path", "ShortName", "ShortPath", "Size", "Type")
    
    Dim vArray1 : Dim vArray2 : Dim vArray3
    For Each boSetFsoFlg In boSetFsoFlgs
        If IsArray(vArray2) Then Erase vArray2
        For Each boTargetIsFile In boTargetIsFiles
            If IsArray(vArray3) Then Erase vArray3
            For Each sPropName In sPropNames
                Call sub_CM_ArrayAddItem(vArray3, func_clsFsBaseTest_1_1_1_C(boSetFsoFlg&boTargetIsFile&sPropName, boSetFsoFlg, boTargetIsFile, sPropName))
            Next
            Call sub_CM_ArrayAddItem(vArray2, vArray3)
        Next
        Call sub_CM_ArrayAddItem(vArray1, vArray2)
    Next
    
    Call aoUtAssistant.RunWithMultiplePatterns("func_clsFsBaseTest_1_1_1_", vArray1)
    
End Sub

'***************************************************************************************************
'Processing Order            : 1-1-1
'Function/Sub Name           : sub_clsFsBaseTest_1_1_1()
'Overview                    : 各属性の取得の正当性を確認する
'Detailed Description        : 実施条件
'                              ・ファイル/フォルダそれぞれについて検証する
'                              ・キャッシュ使用可否は否
'                              ・キャッシュ有効期間は0秒
'                              ・全属性の値を1回取得
'                              期待値
'                              ・全属性の値が正しいこと
'Argument
'     aoUtAssistant          : 単体テスト用アシスタントクラスのインスタンス
'Return Value
'     結果 True,Flase
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/11/03         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_clsFsBaseTest_1_1_1_C( _
    byVal asSubTitle _
    , byVal aboSetFsoFlg _
    , byVal aboTargetIsFile _
    , byVal asPropName _
    )
    Set func_clsFsBaseTest_1_1_1_C = _
        func_clsFsBaseTestCreateArgumentFor_1_1_x(asSubTitle, aboTargetIsFile, False, 0, aboSetFsoFlg, asPropName, False, False)
End Function
'***************************************************************************************************
'Processing Order            : 1-1-2
'Function/Sub Name           : sub_clsFsBaseTest_1_1_2()
'Overview                    : キャッシュ使用可否が変わっていないことを確認する
'Detailed Description        : 実施条件
'                              ・ファイル/フォルダそれぞれについて検証する
'                              ・キャッシュ有効期間は0秒
'                              ・任意の属性の値を1回取得
'                              期待値
'                              ・キャッシュ使用可否が変わっていないこと
'Argument
'     aoUtAssistant          : 単体テスト用アシスタントクラスのインスタンス
'Return Value
'     結果 True,Flase
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/11/18         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub sub_clsFsBaseTest_1_1_2( _
    byRef aoUtAssistant _
    )
    'ファイル/フォルダのパターン
    Dim boTargetIsFile : Dim boTargetIsFiles(2)
    boTargetIsFiles(1) = True : boTargetIsFiles(1) = False
    'キャッシュ使用可否のパターン
    Dim boUseCaches
    boUseCaches = Array(True, False)
    
    Dim oPatterns : Set oPatterns = CreateObject("Scripting.Dictionary")
    Dim sPropName : sPropName = "Attributes"
    Dim lCntOut : Dim lCntIn : Dim boUseCache
    
    'ファイル/フォルダそれぞれについて
    For lCntOut=1 To Ubound(boTargetIsFiles)
        boTargetIsFile = boTargetIsFiles(lCntOut)
        oPatterns.RemoveAll
        
        'キャッシュ使用可否が変わっていないことの検証
        For lCntIn = 0 To Ubound(boUseCaches)
            boUseCache = boUseCaches(lCntIn)
            oPatterns.Add boUseCache, func_clsFsBaseTestCreateArgumentFor_1_1_x(boTargetIsFile, boUseCache, 0, sPropName, True, False)
        Next
        Call aoUtAssistant.RunWithMultiplePatterns("func_clsFsBaseTest_1_1_2_" & Cstr(lCntOut) & "_", oPatterns)
    Next
    
    Set oPatterns = Nothing
End Sub

'***************************************************************************************************
'Processing Order            : 1-1-3
'Function/Sub Name           : sub_clsFsBaseTest_1_1_3()
'Overview                    : キャッシュ有効期間（秒数）が変わっていないことを確認する
'Detailed Description        : 実施条件
'                              ・ファイル/フォルダそれぞれについて検証する
'                              ・キャッシュ使用可否は可
'                              ・任意の属性の値を1回取得
'                              期待値
'                              ・キャッシュ有効期間（秒数）が変わっていないこと
'Argument
'     aoUtAssistant          : 単体テスト用アシスタントクラスのインスタンス
'Return Value
'     結果 True,Flase
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/11/18         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub sub_clsFsBaseTest_1_1_3( _
    byRef aoUtAssistant _
    )
    'ファイル/フォルダのパターン
    Dim boTargetIsFile : Dim boTargetIsFiles(2)
    boTargetIsFiles(1) = True : boTargetIsFiles(1) = False
    'キャッシュ有効期間（秒数）のパターン
    Dim lValidPeriods
    lValidPeriods = Array(0,1,2147483647,-1,-2147483648)
    
    Dim oPatterns : Set oPatterns = CreateObject("Scripting.Dictionary")
    Dim sPropName : sPropName = "Attributes"
    Dim lCntOut : Dim lCntIn : Dim lValidPeriod
    
    'ファイル/フォルダそれぞれについて
    For lCntOut=1 To Ubound(boTargetIsFiles)
        boTargetIsFile = boTargetIsFiles(lCntOut)
        oPatterns.RemoveAll
        
        'ファイルの属性の検証
        For lCntIn = 0 To Ubound(lValidPeriods)
            lValidPeriod = lValidPeriods(lCntIn)
            oPatterns.Add Cstr(lValidPeriod), func_clsFsBaseTestCreateArgumentFor_1_1_x(True, True, lValidPeriod, sPropName, False, True)
        Next
        Call aoUtAssistant.RunWithMultiplePatterns("func_clsFsBaseTest_1_1_3_" & Cstr(lCntOut) & "_", oPatterns)
    Next
    
    Set oPatterns = Nothing
End Sub

'***************************************************************************************************
'Processing Order            : 1-1-x
'Function/Sub Name           : func_clsFsBaseTestCreateArgumentFor_1_1_x()
'Overview                    : func_clsFsBaseTest_1_1()用の引数情報ハッシュマップを作成
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
            , vbNullString _
            , vbNullString _
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
    Dim oConditions : Set oConditions = CreateObject("Scripting.Dictionary")
    With oConditions
        .Add "TargetIsFile", aboTargetIsFile
        .Add "UseCache", aboUseCache
        .Add "ValidPeriod", alValidPeriod
        .Add "SetFsoFlg", aboSetFsoFlg
        .Add "DoItTwice", aboDoItTwice
        .Add "IsRecreate", aboIsRecreate
        .Add "SleepMSecond", alSleepMSecond
    End With
    
    Dim oInspections : Set oInspections = CreateObject("Scripting.Dictionary")
    With oInspections
        .Add "PropName1", asPropName1
        .Add "PropName2", asPropName2
        .Add "DontChgUc", aboDontChgUc
        .Add "DontChgVp", aboDontChgVp
        .Add "IsUpdLcct", aboIsUpdLcct
        .Add "IsUpdLcut", aboIsUpdLcut
    End With
    
    Dim oArgument : Set oArgument = CreateObject("Scripting.Dictionary")
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
Private Function func_clsFsBaseTest_1_1_1_1_( _
    byRef aoArgument _
    )
    func_clsFsBaseTest_1_1_1_1_ = func_clsFsBaseTestNormalBase(aoArgument)
End Function
Private Function func_clsFsBaseTest_1_1_1_2_( _
    byRef aoArgument _
    )
    func_clsFsBaseTest_1_1_1_2_ = func_clsFsBaseTestNormalBase(aoArgument)
End Function
Private Function func_clsFsBaseTest_1_1_2_1_( _
    byRef aoArgument _
    )
    func_clsFsBaseTest_1_1_2_1_ = func_clsFsBaseTestNormalBase(aoArgument)
End Function
Private Function func_clsFsBaseTest_1_1_2_2_( _
    byRef aoArgument _
    )
    func_clsFsBaseTest_1_1_2_2_ = func_clsFsBaseTestNormalBase(aoArgument)
End Function
Private Function func_clsFsBaseTest_1_1_3_1_( _
    byRef aoArgument _
    )
    func_clsFsBaseTest_1_1_3_1_ = func_clsFsBaseTestNormalBase(aoArgument)
End Function
Private Function func_clsFsBaseTest_1_1_3_2_( _
    byRef aoArgument _
    )
    func_clsFsBaseTest_1_1_3_2_ = func_clsFsBaseTestNormalBase(aoArgument)
End Function

'***************************************************************************************************
'Processing Order            : none
'Function/Sub Name           : func_clsFsBaseTestNormalBase()
'Overview                    : ノーマルケース汎用実行
'Detailed Description        : 引数情報ハッシュマップの構成はfunc_clsFsBaseTestCreateArgument()を参照
'                              本関数で使用する項目に限定して記載する
'                              実施条件のハッシュマップの内容
'                              Key                      Value
'                              -------------------      --------------------------------------------
'                              "TargetIsFile"           対象はファイルか否か
'                              "UseCache"               キャッシュ使用可否
'                              "ValidPeriod"            キャッシュ有効期間（秒数）
'
'                              検証内容のハッシュマップの内容
'                              Key                      Value
'                              -------------------      --------------------------------------------
'                              "PropName1"              1回目に取得する属性名（2回目がない場合は値を検証する）
'                              "DontChgUc"              最後にキャッシュ使用可否が変わっていないことを検証するか否か
'                              "DontChgVp"              最後にキャッシュ有効期間（秒数）が変わっていないことを検証するか否か
'
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
    End With
    With aoArgument.Item("Inspections")
    '検証内容
        Dim sPropName : sPropName = .Item("PropName1")
        Dim boDontChgUc : boDontChgUc = .Item("DontChgUc")
        Dim boDontChgVp : boDontChgVp = .Item("DontChgVp")
    End With
    
    '前処理 一時ファイル/フォルダ作成、期待値取得
    Dim oExpect
    Dim boResult : boResult = True
    Dim sPath : sPath = func_UtGetThisTempFilePath()
    If boTargetIsFile Then
        Call CreateObject("Scripting.FileSystemObject").CreateTextFile(sPath)
        If Not(func_CM_FsFileExists(sPath)) Then Exit Function
        Set oExpect = func_clsFsBaseTestGetExpectedValue(func_CM_FsGetFile(sPath))
    Else
        Call func_CM_FsCreateFolder(sPath)
        If Not(func_CM_FsFolderExists(sPath)) Then Exit Function
        Set oExpect = func_clsFsBaseTestGetExpectedValue(func_CM_FsGetFolder(sPath))
    End If
    
    With New clsFsBase
        'テスト対象クラスに条件を指定
        .UseCache = boUseCache
        .ValidPeriod = dbValidPeriod
        .Path = sPath
        
        '指定したプロパティの値を検証
        If IsObject(oExpect.Item(sPropName)) Then
            If Not (.Prop(sPropName) Is oExpect.Item(sPropName)) Then boResult = False
        Else
            If .Prop(sPropName) <> oExpect.Item(sPropName) Then boResult = False
        End If
        
        'キャッシュ使用可否が変わっていないことの検証
        If (boDontChgUc=True) Then boResult = (boUseCache = .UseCache)
        
        'キャッシュ有効期間（秒数）が変わっていないことの検証
        If (boDontChgVp=True) Then boResult = (dbValidPeriod = .ValidPeriod)
        
    End With
    
    '後処理 一時ファイル/フォルダ削除
    If boTargetIsFile Then Call func_CM_FsDeleteFile(sPath) Else Call func_CM_FsDeleteFolder(sPath)
    Set oExpect = Nothing
    
    '結果返却
    func_clsFsBaseTestNormalBase = boResult
End Function

''***************************************************************************************************
''Processing Order            : 1-1
''Function/Sub Name           : func_clsFsBaseTest_1_1()
''Overview                    : 各プロパティの値の取得の正当性（1回目）
''Detailed Description        : 実施条件
''                              ・キャッシュ使用可否は可
''                              ・キャッシュ有効期間は3600秒
''                              ・全プロパティの値を1回取得
''                              期待値
''                              ・全プロパティの値が正しいこと
''                              ・キャッシュ使用可否、同有効期間が変わらないこと
''                              ・キャッシュ確認あり（最終キャッシュ確認時間が初期値でないこと）
''                              ・キャッシュ使用なし（最終キャッシュ更新時間が初期値でないこと）
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
'Private Function func_clsFsBaseTest_1_1( _
'    )
'    Dim boResult : boResult = True
'    
'    '実施条件
'    Dim boUseCache : boUseCache = True
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
'        boResult = func_clsFsBaseTestValidateAllItems(oSut, oExpect)
'        
'        '検証
'        If .UseCache <> boUseCache Then boResult = False
'        If .ValidPeriod <> dbValidPeriod Then boResult = False
'        If .LastCacheConfirmationTime = 0 Then boResult = False
'        If .LastCacheUpdateTime = 0 Then boResult = False
'        
'        '一時ファイル削除
'        Call func_CM_FsDeleteFile(sPath)
'    End With
'    
'    '実施結果
'    func_clsFsBaseTest_1_1 = boResult
'    Set oExpect = Nothing
'    Set oSut = Nothing
'End Function

'***************************************************************************************************
'Processing Order            : 1-2
'Function/Sub Name           : func_clsFsBaseTest_1_2()
'Overview                    : 各プロパティの値の取得の正当性（2回目、キャッシュ無効）
'Detailed Description        : 実施条件
'                              ・キャッシュ使用可否は否
'                              ・キャッシュ有効期間は3600秒
'                              ・全プロパティの値を2回取得
'                              期待値
'                              ・2回目に取得した全プロパティの値が正しいこと
'                              ・キャッシュ使用可否、同有効期間が変わらないこと
'                              ・キャッシュ確認なし（最終キャッシュ確認時間が1回目取得後から変わっていないこと）
'                              ・キャッシュ使用なし（最終キャッシュ更新時間が1回目取得後から変わっていること）
'Argument
'     aoUtAssistant          : 単体テスト用アシスタントクラスのインスタンス
'Return Value
'     結果 True,Flase
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/11/03         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_clsFsBaseTest_1_2( _
    )
    Dim boResult : boResult = True
    
    '実施条件
    Dim boUseCache : boUseCache = False
    Dim dbValidPeriod : dbValidPeriod = 3600
    
    'テスト対象
    Dim oSut : Set oSut = New clsFsBase
    With oSut
        '一時ファイル作成、期待値取得
        Dim sPath : sPath = func_UtGetThisTempFilePath()
        Call CreateObject("Scripting.FileSystemObject").CreateTextFile(sPath)
        If Not(func_CM_FsFileExists(sPath)) Then Exit Function
        Dim oExpect : Set oExpect = func_clsFsBaseTestGetExpectedValue(func_CM_FsGetFile(sPath))
        
        'テスト対象クラスに条件を指定
        .UseCache = boUseCache
        .ValidPeriod = dbValidPeriod
        .Path = sPath
        
        '全プロパティの値を取得（1回目）
        Call func_clsFsBaseTestValidateAllItems(oSut, oExpect)
        Dim lLastCacheConfirmationTime : lLastCacheConfirmationTime = .LastCacheConfirmationTime
        Dim lLastCacheUpdateTime : lLastCacheUpdateTime = .LastCacheUpdateTime
        
        '10msスリープ
        WScript.Sleep 10
        
        '全プロパティの値を取得（2回目）
        boResult = func_clsFsBaseTestValidateAllItems(oSut, oExpect)
        
        '検証
        If .UseCache <> boUseCache Then boResult = False
        If .ValidPeriod <> dbValidPeriod Then boResult = False
        If .LastCacheConfirmationTime <> lLastCacheConfirmationTime Then boResult = False
        If .LastCacheUpdateTime = lLastCacheUpdateTime Then boResult = False
        
        '一時ファイル削除
        Call func_CM_FsDeleteFile(sPath)
    End With
    
    '実施結果
    func_clsFsBaseTest_1_2 = boResult
    Set oExpect = Nothing
    Set oSut = Nothing
End Function

'***************************************************************************************************
'Processing Order            : 1-3
'Function/Sub Name           : func_clsFsBaseTest_1_3()
'Overview                    : 各プロパティの値の取得の正当性（2回目、キャッシュ有効期間超過かつファイル更新なし）
'Detailed Description        : 実施条件
'                              ・キャッシュ使用可否は可
'                              ・キャッシュ有効期間は0秒
'                              ・全プロパティの値を2回取得
'                              ・1回目と2回目でファイルの最終更新日が変わっていない
'                              期待値
'                              ・2回目に取得した全プロパティの値が正しいこと
'                              ・キャッシュ使用可否、同有効期間が変わらないこと
'                              ・キャッシュ確認あり（最終キャッシュ確認時間が1回目取得後から変わっていること）
'                              ・キャッシュ使用あり（最終キャッシュ更新時間が1回目取得後から変わっていないこと）
'Argument
'     aoUtAssistant          : 単体テスト用アシスタントクラスのインスタンス
'Return Value
'     結果 True,Flase
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/11/03         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_clsFsBaseTest_1_3( _
    )
    Dim boResult : boResult = True
    
    '実施条件
    Dim boUseCache : boUseCache = True
    Dim dbValidPeriod : dbValidPeriod = 0
    
    'テスト対象
    Dim oSut : Set oSut = New clsFsBase
    With oSut
        '一時ファイル作成、期待値取得
        Dim sPath : sPath = func_UtGetThisTempFilePath()
        Call CreateObject("Scripting.FileSystemObject").CreateTextFile(sPath)
        If Not(func_CM_FsFileExists(sPath)) Then Exit Function
        Dim oExpect : Set oExpect = func_clsFsBaseTestGetExpectedValue(func_CM_FsGetFile(sPath))
        
        'テスト対象クラスに条件を指定
        .UseCache = boUseCache
        .ValidPeriod = dbValidPeriod
        .Path = sPath
        
        '全プロパティの値を取得（1回目）
        Call func_clsFsBaseTestValidateAllItems(oSut, oExpect)
        Dim lLastCacheConfirmationTime : lLastCacheConfirmationTime = .LastCacheConfirmationTime
        Dim lLastCacheUpdateTime : lLastCacheUpdateTime = .LastCacheUpdateTime
        
        '10msスリープ
        WScript.Sleep 10
        
        '全プロパティの値を取得（2回目）
        boResult = func_clsFsBaseTestValidateAllItems(oSut, oExpect)
        
        '検証
        If .UseCache <> boUseCache Then boResult = False
        If .ValidPeriod <> dbValidPeriod Then boResult = False
        If .LastCacheConfirmationTime = lLastCacheConfirmationTime Then boResult = False
        If .LastCacheUpdateTime <> lLastCacheUpdateTime Then boResult = False
        
        '一時ファイル削除
        Call func_CM_FsDeleteFile(sPath)
    End With
    
    '実施結果
    func_clsFsBaseTest_1_3 = boResult
    Set oExpect = Nothing
    Set oSut = Nothing
End Function

'***************************************************************************************************
'Processing Order            : 1-4
'Function/Sub Name           : func_clsFsBaseTest_1_4()
'Overview                    : 各プロパティの値の取得の正当性（2回目、キャッシュ有効期間超過かつファイル更新あり）
'Detailed Description        : 実施条件
'                              ・キャッシュ使用可否は可
'                              ・キャッシュ有効期間は0秒
'                              ・全プロパティの値を2回取得
'                              ・1回目と2回目でファイルの最終更新日が変わっていない
'                              期待値
'                              ・2回目に取得した全プロパティの値が正しいこと
'                              ・キャッシュ使用可否、同有効期間が変わらないこと
'                              ・キャッシュ確認あり（最終キャッシュ確認時間が1回目取得後から変わっていること）
'                              ・キャッシュ使用なし（最終キャッシュ更新時間が1回目取得後から変わっていること）
'Argument
'     aoUtAssistant          : 単体テスト用アシスタントクラスのインスタンス
'Return Value
'     結果 True,Flase
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/11/03         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_clsFsBaseTest_1_4( _
    )
    Dim boResult : boResult = True
    
    '実施条件
    Dim boUseCache : boUseCache = True
    Dim dbValidPeriod : dbValidPeriod = 0
    
    'テスト対象
    Dim oSut : Set oSut = New clsFsBase
    With oSut
        '一時ファイル作成、期待値取得
        Dim sPath : sPath = func_UtGetThisTempFilePath()
        Call CreateObject("Scripting.FileSystemObject").CreateTextFile(sPath)
        If Not(func_CM_FsFileExists(sPath)) Then Exit Function
        Dim oExpect : Set oExpect = func_clsFsBaseTestGetExpectedValue(func_CM_FsGetFile(sPath))
        
        'テスト対象クラスに条件を指定
        .UseCache = boUseCache
        .ValidPeriod = dbValidPeriod
        .Path = sPath
        
        '全プロパティの値を取得（1回目）
        Call func_clsFsBaseTestValidateAllItems(oSut, oExpect)
        Dim lLastCacheConfirmationTime : lLastCacheConfirmationTime = .LastCacheConfirmationTime
        Dim lLastCacheUpdateTime : lLastCacheUpdateTime = .LastCacheUpdateTime
        
        '10msスリープ
        WScript.Sleep 10
        
        '一時ファイル削除＆再作成、期待値の取得
        Call func_CM_FsDeleteFile(sPath)
        Call CreateObject("Scripting.FileSystemObject").CreateTextFile(sPath)
        If Not(func_CM_FsFileExists(sPath)) Then Exit Function
        oExpect.RemoveAll
        Set oExpect = func_clsFsBaseTestGetExpectedValue(func_CM_FsGetFile(sPath))
        
        '全プロパティの値を取得（2回目）
        boResult = func_clsFsBaseTestValidateAllItems(oSut, oExpect)
        
        '検証
        If .UseCache <> boUseCache Then boResult = False
        If .ValidPeriod <> dbValidPeriod Then boResult = False
        If .LastCacheConfirmationTime = lLastCacheConfirmationTime Then boResult = False
        If .LastCacheUpdateTime = lLastCacheUpdateTime Then boResult = False
        
        '一時ファイル削除
        Call func_CM_FsDeleteFile(sPath)
    End With
    
    '実施結果
    func_clsFsBaseTest_1_4 = boResult
    Set oExpect = Nothing
    Set oSut = Nothing
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
    
    Dim oExpect : Set oExpect = CreateObject("Scripting.Dictionary")
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
'Function/Sub Name           : func_clsFsBaseTestValidateAllItems()
'Overview                    : 全項目の検証を行う
'Detailed Description        : 工事中
'Argument
'     aoSut                  : テスト対象クラス
'     aoExpect               : 期待値のハッシュマップ
'Return Value
'     結果 True,Flase
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/11/03         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_clsFsBaseTestValidateAllItems( _
    byRef aoSut _
    , byRef aoExpect _
    )
    Dim boFlg : boFlg = True
    
    With aoExpect
        Dim sKey
        For Each sKey In .Keys
            If IsObject(.Item(sKey)) Then
                If Not (aoSut.Prop(sKey) Is .Item(sKey)) Then boFlg = False
            Else
                If aoSut.Prop(sKey) <> .Item(sKey) Then boFlg = False
            End If
        Next
    End With
    
    func_clsFsBaseTestValidateAllItems = boFlg
    
End Function

