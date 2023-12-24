'***************************************************************************************************
'FILENAME                    : clsCompareExcel.vbs
'Overview                    : エクセルファイルの比較を行う
'Detailed Description        : 共通関数ライブラリを読み込んでから使用すること
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/10/15         Y.Fujii                  First edition
'***************************************************************************************************
Class clsCompareExcel
    'クラス内変数、定数
    Private PsPathFrom, PsPathTo, PoBroker
    
    '***************************************************************************************************
    'Function/Sub Name           : Class_Initialize()
    'Overview                    : コンストラクタ
    'Detailed Description        : 内部変数の初期化
    'Argument
    '     なし
    'Return Value
    '     なし
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/03         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub Class_Initialize()
        '初期化
        PsPathFrom = ""
        PsPathTo = ""
        Set PoBroker = Nothing
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : Class_Terminate()
    'Overview                    : デストラクタ
    'Detailed Description        : 終了処理
    'Argument
    '     なし
    'Return Value
    '     なし
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/03         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub Class_Terminate()
        Set PoBroker = Nothing
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Let pathFrom()
    'Overview                    : 比較元エクセルファイルのパスを設定する
    'Detailed Description        : 工事中
    'Argument
    '     asPath                 : 比較するエクセルファイルのパス
    'Return Value
    '     なし
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2022/10/15         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Let pathFrom( _
        byVal asPath _
        )
        If new_Fso().FileExists(asPath) Then PsPathFrom = asPath Else PsPathFrom = ""
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get pathFrom()
    'Overview                    : 比較元エクセルファイルのパスを返す
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     比較元エクセルファイルのパス
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2022/10/15         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get pathFrom()
        pathFrom = PsPathFrom
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Let pathTo()
    'Overview                    : 比較先エクセルファイルのパスを設定する
    'Detailed Description        : 工事中
    'Argument
    '     asPath                 : 比較するエクセルファイルのパス
    'Return Value
    '     なし
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2022/10/15         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Let pathTo( _
        byVal asPath _
        )
        If new_Fso().FileExists(asPath) Then PsPathTo = asPath Else PsPathTo = ""
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get pathTo()
    'Overview                    : 比較先エクセルファイルのパスを返す
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     比較先エクセルファイルのパス
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2022/10/15         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get pathTo()
        pathTo = PsPathTo
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Set broker()
    'Overview                    : 出版-購読型（Publish/Subscribe）クラスのオブジェクトを設定する
    'Detailed Description        : 工事中
    'Argument
    '     aoBroker               : 出版-購読型（Publish/Subscribe）クラスのインスタンス
    'Return Value
    '     なし
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/03         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Set broker( _
        byRef aoBroker _
        )
        Set PoBroker = aoBroker
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get broker()
    'Overview                    : 出版-購読型（Publish/Subscribe）クラスのオブジェクトを返す
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     出版-購読型（Publish/Subscribe）クラスのインスタンス
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/03         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get broker()
        Set broker = PoBroker
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : compare()
    'Overview                    : エクセルファイルを比較する
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     結果 True:正常完了 / False:失敗
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2022/10/15         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function compare( _
        )
        Dim sMyName : sMyName = "+compare"
        '★ログ出力
        Call sub_CmpExcelPublish("log", 5, sMyName, "Start")
        Call sub_CmpExcelPublish("log", 9, sMyName, "PsPathFrom = " & cf_toString(PsPathFrom) & ", PsPathTo = " & cf_toString(PsPathTo))
        
        compare = False
        
        '比較結果用の新規ワークブックを作成
        With CreateObject("Excel.Application")
            .DisplayAlerts = False
            .ScreenUpdating = False
            .AutomationSecurity = 3                               'msoAutomationSecurityForceDisable = 3
            Dim oWorkbookForResults
            Set oWorkbookForResults = .Workbooks.Add(-4167)      '新規ワークブック xlWBATWorksheet=-4167
        End With
        '★ログ出力
        Call sub_CmpExcelPublish("log", 3, sMyName, "Create a new workbook for comparison.")
        
        Dim oParams : Set oParams = new_DicWith(Array("WorkbookForResults", oWorkbookForResults))
        
        '比較対象ファイルの全シートを比較結果用ワークブックにコピーする
        Call sub_CmpExcelCopyAllSheetsToWorkbookForResults(oParams)
        
        'エクセルファイルを比較する
        Call sub_CmpExcelCompare(oParams)
        
        '★ログ出力
        Call sub_CmpExcelPublish("log", 5, sMyName, "End")
        
        '終了
        Set oParams = Nothing
        Set oWorkbookForResults = Nothing
        compare = True
    End Function
    
    
    
    
    
    '***************************************************************************************************
    'Function/Sub Name           : sub_CmpExcelCopyAllSheetsToWorkbookForResults()
    'Overview                    : 比較対象ファイルの全シートを比較結果用ワークブックにコピーする
    'Detailed Description        : パラメータ格納用汎用オブジェクトに格納する
    '                              ワークシートのリネーム情報のハッシュマップの構成
    '                              Key                       Value
    '                              --------------------      -------------------------------------------
    '                              "WorkbookForResults"      比較結果用のワークブック
    '                              "From"                    比較元ワークシートのリネーム情報（clsCmArray型）
    '                              "To"                      比較先ワークシートのリネーム情報（clsCmArray型）
    'Argument
    '     aoParams               : パラメータ格納用汎用オブジェクト
    'Return Value
    '     なし
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2017/04/26         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub sub_CmpExcelCopyAllSheetsToWorkbookForResults( _
        byRef aoParams _
        )
        Dim sMyName : sMyName = "-sub_CmpExcelCopyAllSheetsToWorkbookForResults"
        '★ログ出力
        Call sub_CmpExcelPublish("log", 5, sMyName, "Start")
        Call sub_CmpExcelPublish("log", 9, sMyName, cf_toString(aoParams))
        
        'パラメータ格納用汎用オブジェクトから必要な要素を取り出す
        Dim oWorkbookForResults : Call cf_bind(oWorkbookForResults, aoParams.Item("WorkbookForResults"))
        
        Dim sPath : Dim sFromToString
        '比較元ファイルのコピー
        sPath = PsPathFrom : sFromToString = "From" 
        Call aoParams.Add(sFromToString, _
            func_CmpExcelCopyAllSheetsToWorkbookForResultsDetail(oWorkbookForResults, sPath, sFromToString))
        
        '★ログ出力
        Call sub_CmpExcelPublish("log", 3, sMyName, "Source file copy completed.")
        Call sub_CmpExcelPublish("log", 9, sMyName, cf_toString(aoParams))
        
        '比較先ファイルのコピー
        sPath = PsPathTo : sFromToString = "To"
        Call aoParams.Add(sFromToString, _
            func_CmpExcelCopyAllSheetsToWorkbookForResultsDetail(oWorkbookForResults, sPath, sFromToString))
        
        '★ログ出力
        Call sub_CmpExcelPublish("log", 5, sMyName, "End")
        Call sub_CmpExcelPublish("log", 9, sMyName, cf_toString(aoParams))
        
        Set oWorkbookForResults = Nothing
    End Sub

    '***************************************************************************************************
    'Function/Sub Name           : func_CmpExcelCopyAllSheetsToWorkbookForResultsDetail()
    'Overview                    : 比較対象ファイルの全シートを比較結果用ワークブックにコピーする
    'Detailed Description        : 比較対象の全シートを比較結果用ワークブックにコピーした上で、
    '                              シートごとの変更前後のシート名を格納したオブジェクト（以下参照）
    '                              の配列（clsCmArray型）を返す
    '                              Key                      Value
    '                              -------------------      --------------------------------------------
    '                              "Before"                 変更前のワークシート名
    '                              "After"                  変更後のワークシート名
    'Argument
    '     aoWorkbookForResults   : 比較結果用のワークブック
    '     asPath                 : 比較対象ファイルのパス
    '     asFromToString         : 比較元先を識別する文字列 "From","To"
    'Return Value
    '     シートごとの変更前後のシート名を格納したオブジェクトの配列（clsCmArray型）
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2017/04/26         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function func_CmpExcelCopyAllSheetsToWorkbookForResultsDetail( _
        byRef aoWorkbookForResults _
        , byVal asPath _
        , byVal asFromToString _
        )
        Dim sMyName : sMyName = "-func_CmpExcelCopyAllSheetsToWorkbookForResultsDetail"
        
        '★ログ出力
        Call sub_CmpExcelPublish("log", 5, sMyName, "Start")
        Call sub_CmpExcelPublish("log", 9, sMyName, "aoWorkbookForResults = " & cf_toString(aoWorkbookForResults) & ", asPath = " & cf_toString(asPath)& ", asFromToString = " & cf_toString(asFromToString))

        '比較対象ファイルを開く
        Dim oExcel : Set oExcel = aoWorkbookForResults.Parent
        Dim oWorkBook : Set oWorkBook = func_CM_ExcelOpenFile(oExcel, asPath)
        '★ログ出力
        Call sub_CmpExcelPublish("log", 3, sMyName, "Opened Excel file, file path is " & cf_toString(asPath) )
        
        Dim sTempPath : sTempPath = ""
        If oWorkBook.HasVBProject Then
        'マクロありの場合は別名で保存した上で再度開く
            sTempPath = func_CM_FsGetTempFilePath()
            Call sub_CM_ExcelSaveAs(oWorkBook, sTempPath, vbNullString)
            Set oWorkBook = func_CM_ExcelOpenFile( oExcel, sTempPath)
            '★ログ出力
            Call sub_CmpExcelPublish("log", 3, sMyName, "It was Excel with a macro, so save it with a different name and reopen it.")
        End If

        '★ログ出力
        Call sub_CmpExcelPublish("log", 3, sMyName, "Attempt to unprotect Excel file." )
        '文書の保護を解除する
        Call sub_CmpExcelTryCatchAfterProc(cf_tryCatch(new_Func("a=>a.Unprotect"), oWorkBook, empty, empty), sMyName)
        
        With oWorkBook
            'ワークシートのリネーム情報格納用配列（clsCmArray型）
            Dim oWorkSheetRenameInfo : Set oWorkSheetRenameInfo = new_Arr()
            'タブの色変換用ハッシュマップ定義
            Dim oStringToThemeColor : Set oStringToThemeColor = new_DicWith(Array("From", 2, "To", 8))
            
            Dim oWorksheet, sNewSheetName
            For Each oWorksheet In .Worksheets
                If oWorksheet.Visible=True Then
                '全ての見えるシートを比較結果用ワークブックにコピーする
                    '★ログ出力
                    Call sub_CmpExcelPublish("log", 3, sMyName, "Start processing sheet " & cf_toString(oWorksheet.Name) & "." )
                    
                    'シート保護の解除
                    Call sub_CmpExcelPublish("log", 3, sMyName, "Try to unprotect a sheet.")
                    Call sub_CmpExcelTryCatchAfterProc(cf_tryCatch(new_Func("a=>{If a.ProtectContents Then:a.Unprotect(vbNullString):End If}"), oWorksheet, empty, empty), sMyName)
                    
                    'オートフィルタの解除
                    Call sub_CmpExcelPublish("log", 3, sMyName, "Try to clear the AutoFilter.")
                    Call sub_CmpExcelTryCatchAfterProc(cf_tryCatch(new_Func("a=>{If a.AutoFilterMode Then:a.Cells(1,1).AutoFilter:End If}"), oWorksheet, empty, empty), sMyName)
                    
                    'ワークシート名取得および変更する名称を決める
                    sNewSheetName = func_CmpExcelMakeSheetName(oWorkSheetRenameInfo.Length+1, asFromToString)
                    oWorkSheetRenameInfo.Push new_DicWith( Array("Before", oWorksheet.Name, "After", sNewSheetName) )
                    '★ログ出力
                    Call sub_CmpExcelPublish("log", 9, sMyName, "oWorkSheetRenameInfo = " & cf_toString(oWorkSheetRenameInfo) )
                    
                    'シート名変更＆タブの色を変更
                    oWorksheet.Name = sNewSheetName
                    oWorksheet.Tab.ThemeColor = oStringToThemeColor.Item(asFromToString)
                    oWorksheet.Tab.TintAndShade = 0
                    'シートの表示を整える
                    oWorksheet.Activate
                    .Windows(1).View = 1                      'xlNormalView 標準
                    .Windows(1).Zoom = 25                     '表示倍率
                    .Windows(1).ScrollColumn = 1              '列1が左端になるようにウィンドウをスクロール
                    .Windows(1).ScrollRow = 1                 '行1が上端になるようにウィンドウをスクロール
                    .Windows(1).FreezePanes = False           'ウィンドウ枠の固定解除
                    
                    '★ログ出力
                    Call sub_CmpExcelPublish("log", 3, sMyName, "Start copying sheets to a new workbook for comparison results.")
                    'シートを比較結果用の新規ワークブックにコピー
                    Call oWorksheet.Copy(, aoWorkbookForResults.Worksheets(aoWorkbookForResults.Worksheets.Count))
                    '★ログ出力
                    Call sub_CmpExcelPublish("log", 3, sMyName, "Copy Complete.")
                End If
            Next

            '比較対象ファイルを閉じる
            Call .Close(False)
            '★ログ出力
            Call sub_CmpExcelPublish("log", 3, sMyName, "Close the file being compared." )
        End With
        
        If Len(sTempPath) Then
        'マクロありの場合に別名で保存したファイルがあったら削除する
            fs_deleteFile sTempPath
            '★ログ出力
            Call sub_CmpExcelPublish("log", 3, sMyName, "Delete file saved with a different name.")
        End If

        'サマリーシートのカラム位置変換用ハッシュマップ定義
        Dim oStringToColumn : Set oStringToColumn = new_DicWith(Array("From", 1, "To", 2))
        'サマリーシートに比較対象ファイルの情報を出力
        Dim lRow : Dim lColumn : Dim oItem
        lColumn = oStringToColumn.Item(asFromToString)
        With aoWorkbookForResults.Worksheets.Item(1)
            'ファイルパス
            lRow = 1
            .Cells(lRow, lColumn).Value = asPath
            'シート名
            For Each oItem In oWorkSheetRenameInfo.Map(new_Func( "(e,i,a)=>e.Item(""Before"")" ) ).Items
                lRow = lRow + 1
                .Cells(lRow, lColumn).Value = oItem
            Next
        End With
        '★ログ出力
        Call sub_CmpExcelPublish("log", 3, sMyName, "Output the information of the files to be compared in the summary sheet.")

        'ワークシートのリネーム情報を返却
        Set func_CmpExcelCopyAllSheetsToWorkbookForResultsDetail = oWorkSheetRenameInfo
        '★ログ出力
        Call sub_CmpExcelPublish("log", 5, sMyName, "End")
        Call sub_CmpExcelPublish("log", 9, sMyName, "func_CmpExcelCopyAllSheetsToWorkbookForResultsDetail = " & cf_toString(oWorkSheetRenameInfo))
        
        'オブジェクトを開放
        Set oStringToColumn = Nothing
        Set oItem = Nothing
        Set oWorksheet = Nothing
        Set oStringToThemeColor = Nothing
        Set oWorkSheetRenameInfo = Nothing
        Set oWorkBook = Nothing
        Set oExcel = Nothing
        
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : func_CmpExcelMakeSheetName()
    'Overview                    : シート名作成
    'Detailed Description        : 工事中
    'Argument
    '     alCnt                  : シートの先頭からの番号
    '     asFromToString         : 比較元先を識別する文字列 "From","To"
    'Return Value
    '     シート名
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2017/04/26         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function func_CmpExcelMakeSheetName( _
        byVal alCnt _
        , byVal asFromToString _
        )
        func_CmpExcelMakeSheetName = "【" & asFromToString & "_" & CStr(alCnt) & "シート目】"
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : sub_CmpExcelCompare()
    'Overview                    : エクセルファイルを比較する
    'Detailed Description        : 工事中
    'Argument
    '     aoParams               : パラメータ格納用汎用オブジェクト
    'Return Value
    '     なし
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2017/04/26         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub sub_CmpExcelCompare( _
        byRef aoParams _
        )
        Dim sMyName : sMyName = "-sub_CmpExcelCompare"
        '★ログ出力
        Call sub_CmpExcelPublish("log", 5, sMyName, "Start")
        Call sub_CmpExcelPublish("log", 9, sMyName, "aoParams = " & cf_toString(aoParams))
        
        'パラメータ格納用汎用オブジェクトから必要な要素を取り出す
        Dim oWorkbookForResults : Call cf_bind(oWorkbookForResults, aoParams.Item("WorkbookForResults"))
        Dim oFrom : Call cf_bind(oFrom, aoParams.Item("From"))
        Dim oTo : Call cf_bind(oTo, aoParams.Item("To"))

        Dim lCnt
        For lCnt = 0 To math_min(oFrom.Length, oTo.Length)-1
        '比較元先の各シートに差分が分かる書式設定をする
            '★ログ出力
            Call sub_CmpExcelPublish("log", 3, sMyName, "Comparison of " & lCnt+1 & "th sheets.")
            
            '比較元（To）のシートに対し比較先（From）との差分が分かるようにする
            Call sub_CmpExcelSetFormatToUnderstandDifference(_
                    oWorkbookForResults, oFrom(lCnt).Item("After"), oTo(lCnt).Item("After"))
            '★ログ出力
            Call sub_CmpExcelPublish("log", 3, sMyName, "to see the difference from the comparison destination (" & oFrom(lCnt).Item("Before") & ") to the source sheet (" & oTo(lCnt).Item("Before") & ").")
            
            '比較先（From）のシートに対し比較元（To）との差分が分かるようにする
            Call sub_CmpExcelSetFormatToUnderstandDifference( _
                    oWorkbookForResults, oTo(lCnt).Item("After"), oFrom(lCnt).Item("After"))
            '★ログ出力
            Call sub_CmpExcelPublish("log", 3, sMyName, "to see the difference from the comparison source (" & oTo(lCnt).Item("Before") & ") to the comparison destination sheet (" & oFrom(lCnt).Item("Before") & ").")
            
        Next
        
        '★ログ出力
        Call sub_CmpExcelPublish("log", 3, sMyName, "Arrange the Window so that you can see the difference.")
        '同じブックの新しいウィンドウを開く
        oWorkbookForResults.Worksheets(oFrom(0).Item("After")).Activate
        With oWorkbookForResults.Windows(1).NewWindow
            Dim sCaption : sCaption = .Caption
            Dim oWorksheet
            For Each oWorksheet In .Parent.Worksheets
                oWorksheet.Activate
                .Zoom = 25
            Next
        End With
        oWorkbookForResults.Worksheets(oTo(0).Item("After")).Activate
        '並べて比較
        oWorkbookForResults.Activate
        With oWorkbookForResults.Parent
            .Windows.CompareSideBySideWith(sCaption)
            Call .Windows.Arrange(-4166, True)               'xlVertical = -4166
            .DisplayAlerts = True
            .ScreenUpdating = True
            .AutomationSecurity = 2                     'msoAutomationSecurityByUI = 2 [ セキュリティ] ダイアログ ボックスで指定されたセキュリティ設定を使用
            .Visible = True
        End With
        
        '★ログ出力
        Call sub_CmpExcelPublish("log", 5, sMyName, "End")
        
        'オブジェクトを開放
        Set oWorksheet = Nothing
        Set oTo = Nothing
        Set oFrom = Nothing
        Set oWorkbookForResults = Nothing

    End Sub

    '***************************************************************************************************
    'Function/Sub Name           : sub_CmpExcelSetFormatToUnderstandDifference()
    'Overview                    : asSheetNameAのシートにasSheetNameBシートとの差分が分かる書式設定をする
    'Detailed Description        : 工事中
    'Argument
    '     aoWorkbookForResults   : 比較結果用のワークブック
    '     asSheetNameA           : 比較元のシート名
    '     asSheetNameB           : 比較先のシート名
    'Return Value
    '     なし
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2017/04/26         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub sub_CmpExcelSetFormatToUnderstandDifference( _
        byRef aoWorkbookForResults _
        , byVal asSheetNameA _
        , byVal asSheetNameB _
        )
        Dim sMyName : sMyName = "-sub_CmpExcelSetFormatToUnderstandDifference"

        'セルの比較
        aoWorkbookForResults.Worksheets(asSheetNameA).Activate
        aoWorkbookForResults.Worksheets(asSheetNameA).UsedRange.Select
        Dim oExcel : Set oExcel = aoWorkbookForResults.Parent
        Call oExcel.Selection.FormatConditions.Add( _
                2 _
                , _
                , "=EXACT(OFFSET($A$1,ROW()-1,COLUMN()-1),OFFSET('" _
                & asSheetNameB _
                & "'!$A$1,ROW()-1,COLUMN()-1))=TRUE" _
                )    'xlExpression=2（数式）
        oExcel.Selection.FormatConditions(oExcel.Selection.FormatConditions.Count).SetFirstPriority

        With oExcel.Selection.FormatConditions(1).Interior
            .Pattern = 1                        '実践 xlSolid
            .PatternColorIndex = -4105          '自動 xlAutomatic
            .ThemeColor = 1                     '濃色 xlThemeColorDark1
            .TintAndShade = -0.149998474074526  '色を明るくするかまたは暗くする
            .PatternTintAndShade = 0            '濃色と網掛けパターン
        End With

        With oExcel.Selection.FormatConditions(1).Font
            .ThemeColor = 1                     '濃色 xlThemeColorDark1
            .TintAndShade = -0.499984740745262  '色を明るくするかまたは暗くする
        End With

        aoWorkbookForResults.Worksheets(asSheetNameA).Range("A1").Select
        '★ログ出力
        Call sub_CmpExcelPublish("log", 3, sMyName, "Cell comparison complete.")

        'オートシェイプの比較
        Dim oAutoshapeA, oAutoshapeB, oRet, sTextA
        For Each oAutoshapeA In aoWorkbookForResults.Worksheets(asSheetNameA).Shapes
            Set oRet = cf_tryCatch(new_Func("(a)=>a(0).Item(a(1))"), Array(aoWorkbookForResults.Worksheets(asSheetNameB).Shapes, oAutoshapeA.Name), Empty, Empty)
            If oRet.Item("Result") Then
                Set oAutoshapeB = oRet.Item("Return")
                Set oRet = cf_tryCatch(Getref("func_CM_ExcelGetTextFromAutoshape"), oAutoshapeA, Empty, Empty)
                If oRet.Item("Result") Then
                    sTextA = oRet.Item("Return")
                    Set oRet = cf_tryCatch(Getref("func_CM_ExcelGetTextFromAutoshape"), oAutoshapeB, Empty, Empty)
                End If
                If oRet.Item("Result") Then
                    If cf_isSame(sTextA, oRet.Item("Return")) Then
                    'オートシェイプの名前とテキストが一致する（差異がない）場合は灰色にする
                        sub_CmpExcelSetAutoshapeColor oAutoshapeA
                    End If
                End If
            End If
        Next

        '★ログ出力
        Call sub_CmpExcelPublish("log", 3, sMyName, "AutoShape comparison complete.")

        'オブジェクトを開放
        Set oRet = Nothing
        Set oAutoshapeA = Nothing
        Set oAutoshapeB = Nothing
        Set oExcel = Nothing
    End Sub

    '***************************************************************************************************
    'Function/Sub Name           : sub_CmpExcelSetAutoshapeColor()
    'Overview                    : オートシェイプの色を灰色にする
    'Detailed Description        : エラーは無視する
    'Argument
    '     aoAutoshape            : オートシェイプ
    'Return Value
    '     なし
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2017/04/26         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub sub_CmpExcelSetAutoshapeColor( _
        byRef aoAutoshape _
        )
        On Error Resume Next
        With aoAutoshape.Fill
            .Visible = True                          'msoTrue
            .ForeColor.ObjectThemeColor = 14         '背景１テーマの色 msoThemeColorBackground1
            .ForeColor.TintAndShade = 0              '色を明るくするかまたは暗くする単精度浮動小数点型 (Single) の値
            .ForeColor.Brightness = -0.150000006     '明度
            .Transparency = 0                        '塗りつぶしの透明度を示す 0.0 (不透明) から 1.0 (透明) までの値
            .Solid                                   '塗りつぶしを均一な色に設定
        End With
        On Error Goto 0
    End Sub

    '***************************************************************************************************
    'Function/Sub Name           : sub_CmpExcelTryCatchAfterProc()
    'Overview                    : TryCatchでエラー時の処理
    'Detailed Description        : 工事中
    'Argument
    '     aoRet                  : cf_tryCatch()の戻り値
    '     asYourName             : 処理を実行した関数名
    'Return Value
    '     なし
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/28         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub sub_CmpExcelTryCatchAfterProc( _
        byRef aoRet _
        , byVal asYourName _
        )
        If aoRet.Item("Result") Then Exit Sub
        sub_CmpExcelPublish "log", 3, asYourName, "It couldn't."
        sub_CmpExcelPublish "log", 9, asYourName, "<Err> " & cf_toString(aoRet.Item("Err"))
    End Sub

    '***************************************************************************************************
    'Function/Sub Name           : sub_CmpExcelPublish()
    'Overview                    : 出版（Publish）処理
    'Detailed Description        : 出版-購読型（Publish/subscribe）クラスがあれば出版（Publish）処理する
    'Argument
    '     asTopic                : トピック
    '     alLevel                : レベル
    '     asFuncName             : 関数名
    '     asCont                 : 内容
    'Return Value
    '     なし
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/28         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub sub_CmpExcelPublish( _
        byVal asTopic _
        , byVal alLevel _
        , byVal asFuncName _
        , byVal asCont _
        )
        If PoBroker Is Nothing Then Exit Sub
        PoBroker.Publish asTopic, Array(alLevel, TypeName(Me)&asFuncName, asCont)
    End Sub
    
End Class
