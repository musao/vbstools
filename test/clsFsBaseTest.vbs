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
    Call func_clsFsBaseTest_1(oUtAssistant)
    
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
Private Sub func_clsFsBaseTest_1( _
    byRef aoUtAssistant _
    )
    
    Call func_clsFsBaseTest_1_1(aoUtAssistant)
'    Call aoUtAssistant.Run("func_clsFsBaseTest_1_1")
'    Call aoUtAssistant.Run("func_clsFsBaseTest_1_2")
'    Call aoUtAssistant.Run("func_clsFsBaseTest_1_3")
'    Call aoUtAssistant.Run("func_clsFsBaseTest_1_4")
'    
End Sub

'***************************************************************************************************
'Processing Order            : 1-1
'Function/Sub Name           : func_clsFsBaseTest_1_1()
'Overview                    : 各プロパティの値の取得の正当性（1回目）
'Detailed Description        : 実施条件
'                              ・キャッシュ使用可否は可
'                              ・キャッシュ有効期間は3600秒
'                              ・全プロパティの値を1回取得
'                              期待値
'                              ・全プロパティの値が正しいこと
'                              ・キャッシュ使用可否、同有効期間が変わらないこと
'                              ・キャッシュ確認あり（最終キャッシュ確認時間が初期値でないこと）
'                              ・キャッシュ使用なし（最終キャッシュ更新時間が初期値でないこと）
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
Private Sub func_clsFsBaseTest_1_1( _
    byRef aoUtAssistant _
    )
    Dim oPatterns : Set oPatterns = CreateObject("Scripting.Dictionary")
    Dim lNum : lNum = 0
    Dim sPropName
    
    lNum = lNum + 1 : sPropName = "Attributes" : oPatterns.Add lNum & "_" & sPropName, func_clsFsBaseTest_1_1_CreateArgument(sPropName)
    lNum = lNum + 1 : sPropName = "DateCreated" : oPatterns.Add lNum & "_" & sPropName, func_clsFsBaseTest_1_1_CreateArgument(sPropName)
    lNum = lNum + 1 : sPropName = "DateLastAccessed" : oPatterns.Add lNum & "_" & sPropName, func_clsFsBaseTest_1_1_CreateArgument(sPropName)
    lNum = lNum + 1 : sPropName = "DateLastModified" : oPatterns.Add lNum & "_" & sPropName, func_clsFsBaseTest_1_1_CreateArgument(sPropName)
    lNum = lNum + 1 : sPropName = "Drive" : oPatterns.Add lNum & "_" & sPropName, func_clsFsBaseTest_1_1_CreateArgument(sPropName)
    lNum = lNum + 1 : sPropName = "Name" : oPatterns.Add lNum & "_" & sPropName, func_clsFsBaseTest_1_1_CreateArgument(sPropName)
    lNum = lNum + 1 : sPropName = "ParentFolder" : oPatterns.Add lNum & "_" & sPropName, func_clsFsBaseTest_1_1_CreateArgument(sPropName)
    lNum = lNum + 1 : sPropName = "Path" : oPatterns.Add lNum & "_" & sPropName, func_clsFsBaseTest_1_1_CreateArgument(sPropName)
    lNum = lNum + 1 : sPropName = "ShortName" : oPatterns.Add lNum & "_" & sPropName, func_clsFsBaseTest_1_1_CreateArgument(sPropName)
    lNum = lNum + 1 : sPropName = "ShortPath" : oPatterns.Add lNum & "_" & sPropName, func_clsFsBaseTest_1_1_CreateArgument(sPropName)
    lNum = lNum + 1 : sPropName = "Size" : oPatterns.Add lNum & "_" & sPropName, func_clsFsBaseTest_1_1_CreateArgument(sPropName)
    lNum = lNum + 1 : sPropName = "Type" : oPatterns.Add lNum & "_" & sPropName, func_clsFsBaseTest_1_1_CreateArgument(sPropName)
    
    Call aoUtAssistant.RunWithMultiplePatterns("func_clsFsBaseTest_1_1_", oPatterns)
    
    Set oPatterns = Nothing
End Sub

'***************************************************************************************************
'Processing Order            : 1-1
'Function/Sub Name           : func_clsFsBaseTest_1_1()
'Overview                    : 各プロパティの値の取得の正当性（1回目）
'Detailed Description        : 実施条件
'                              ・キャッシュ使用可否は可
'                              ・キャッシュ有効期間は3600秒
'                              ・全プロパティの値を1回取得
'                              期待値
'                              ・全プロパティの値が正しいこと
'                              ・キャッシュ使用可否、同有効期間が変わらないこと
'                              ・キャッシュ確認あり（最終キャッシュ確認時間が初期値でないこと）
'                              ・キャッシュ使用なし（最終キャッシュ更新時間が初期値でないこと）
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
Private Function func_clsFsBaseTest_1_1_CreateArgument( _
    byVal asPropName _
    )
    Dim oArgument : Set oArgument = CreateObject("Scripting.Dictionary")
    Dim oConditions : Set oConditions = CreateObject("Scripting.Dictionary")
    Dim oInspections : Set oInspections = CreateObject("Scripting.Dictionary")
    
    oConditions.Add "UseCache", False
    oConditions.Add "ValidPeriod", 0
    
    oInspections.Add "PropName", asPropName
    
    oArgument.Add "Conditions", oConditions
    oArgument.Add "Inspections", oInspections
    
    Set func_clsFsBaseTest_1_1_CreateArgument = oArgument
    
    Set oInspections = Nothing
    Set oConditions = Nothing
    Set oArgument = Nothing
End Function

'***************************************************************************************************
'Processing Order            : 1-1-x
'Function/Sub Name           : func_clsFsBaseTest_1_1_()
'Overview                    : 引数で指定した属性の値の正当性を確認する
'Detailed Description        : 引数情報のハッシュマップの内容
'                              Key                      Value
'                              -------------------      --------------------------------------------
'                              "Conditions"             実施条件のハッシュマップ
'                              "Inspections"            検証内容のハッシュマップ
'
'                              実施条件のハッシュマップの内容
'                              Key                      Value
'                              -------------------      --------------------------------------------
'                              "UseCache"               キャッシュ使用可否
'                              "ValidPeriod"            キャッシュ有効期間（秒数）
'
'                              検証内容のハッシュマップの内容
'                              Key                      Value
'                              -------------------      --------------------------------------------
'                              "PropName"               検証対象の属性名
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
Private Function func_clsFsBaseTest_1_1_( _
    byRef aoArgument _
    )
    '引数情報の取得
    With aoArgument.Item("Conditions")
    '実施条件
        Dim boUseCache : boUseCache = .Item("UseCache")
        Dim dbValidPeriod : dbValidPeriod = .Item("ValidPeriod")
    End With
    With aoArgument.Item("Inspections")
    '検証内容
        Dim sPropName : sPropName = .Item("PropName")
    End With
    
    Dim boResult : boResult = True
    
    '一時ファイル作成、期待値取得
    Dim sPath : sPath = func_UtGetThisTempFilePath()
    Call CreateObject("Scripting.FileSystemObject").CreateTextFile(sPath)
    If Not(func_CM_FsFileExists(sPath)) Then Exit Function
    Dim oExpect : Set oExpect = func_clsFsBaseTestGetExpectedValue(func_CM_FsGetFile(sPath))
    
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
    End With
    
    '一時ファイル削除
    Call func_CM_FsDeleteFile(sPath)
    
    '実施結果
    func_clsFsBaseTest_1_1_ = boResult
    Set oExpect = Nothing
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

