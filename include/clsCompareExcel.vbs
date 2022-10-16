'***************************************************************************************************
'FILENAME                    : clsCompareExcel.vbs
'Overview                    : エクセルファイルの比較を行う
'Detailed Description        : 共通関数ライブラリ（VbsBasicLibCommon.vbs）を読み込んでから使用すること
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/10/15         Y.Fujii                  First edition
'***************************************************************************************************
Class clsCompareExcel
    'クラス内変数、定数
    Private PdtNow
    Private PdtDate
    Private PdtStart
    Private PdtEnd
    Private PsPathFrom
    Private PsPathTo
    Private Cs_FOLDER_TEMP
    
    'コンストラクタ
    Private Sub Class_Initialize()
        '初期化
        PasPathA = ""
        PasPathB = ""
        Cs_FOLDER_TEMP = "tmp"
    End Sub
    'デストラクタ
    Private Sub Class_Terminate()
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Let PathFrom()
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
    Public Property Let PathFrom( _
        byVal asPath _
        )
        If func_CM_FsFileExists(asPath) Then PsPathFrom = asPath Else PsPathFrom = ""
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get PathFrom()
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
    Public Property Get PathFrom()
        PathFrom = PsPathFrom
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Let PathTo()
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
    Public Property Let PathTo( _
        byVal asPath _
        )
        If func_CM_FsFileExists(asPath) Then PsPathTo = asPath Else PsPathTo = ""
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get PathTo()
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
    Public Property Get PathTo()
        PathTo = PsPathTo
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get ProcDate()
    'Overview                    : 処理実施日時を返す
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     処理日時
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2022/10/13         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get ProcDate()
        ProcDate = PdtNow
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get StartTime()
    'Overview                    : 処理開始日時を返す
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     処理開始日時
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2022/10/13         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get StartTime()
        StartTime = func_CM_GetDateInMilliseconds(PdtDate, PdtStart)
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get EndTime()
    'Overview                    : 処理終了日時を返す
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     処理の終了日時
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2022/10/13         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get EndTime()
        EndTime = func_CM_GetDateInMilliseconds(PdtDate, PdtEnd)
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get ElapsedTime()
    'Overview                    : 処理にかかった時間を返す
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     処理にかかった時間
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2022/10/13         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get ElapsedTime()
       ElapsedTime = PdtEnd - PdtStart
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Compare()
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
    Public Function Compare( _
        )
        Compare = False
        
        '開始日時の取得
        PdtNow = Now
        PdtDate = Date
        PdtStart = Timer
        
        '比較結果用の新規ワークブックを作成
        With CreateObject("Excel.Application")
            .DisplayAlerts = False
            .ScreenUpdating = False
            .AutomationSecurity = 3                               'msoAutomationSecurityForceDisable = 3
            Dim oWorkbookForResults
            Set oWorkbookForResults = .Workbooks.Add(-4167)      '新規ワークブック xlWBATWorksheet=-4167
        End With
        
        Dim oParams : Set oParams = CreateObject("Scripting.Dictionary")
        
        '比較対象ファイルの全シートを比較結果用ワークブックにコピーする
        Call sub_CmpExcelCopyAllSheetsToWorkbookForResults(oWorkbookForResults, oParams)
        
        'エクセルファイルを比較する
        Call sub_CmpExcelCompare(oWorkbookForResults, oParams)
        
        '終了
        Set oParams = Nothing
        Set oWorkbookForResults = Nothing
        PdtEnd = Timer
        Compare = True
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : sub_CmpExcelCopyAllSheetsToWorkbookForResults()
    'Overview                    : 比較対象ファイルの全シートを比較結果用ワークブックにコピーする
    'Detailed Description        : パラメータ格納用汎用ハッシュマップに格納する
    '                              ワークシートのリネーム情報のハッシュマップの構成
    '                              Key                      Value
    '                              -------------------      --------------------------------------------
    '                              "From"                   比較元のワークシートのリネーム情報のハッシュマップ
    '                              "To"                     比較先のワークシートのリネーム情報のハッシュマップ
    'Argument
    '     aoParams               : パラメータ格納用汎用ハッシュマップ
    '     aoWorkbookForResults   : 比較結果用のワークブック
    'Return Value
    '     なし
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2017/04/26         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub sub_CmpExcelCopyAllSheetsToWorkbookForResults( _
        byRef aoWorkbookForResults _
        , byRef aoParams _
        )
        
        Dim sPath : Dim sFromToString
        '比較元ファイルのコピー
        sPath = PsPathFrom : sFromToString = "From" 
        Call aoParams.Add(sFromToString, _
            func_CmpExcelCopyAllSheetsToWorkbookForResultsDetail(aoWorkbookForResults, sPath, sFromToString))

        '比較先ファイルのコピー
        sPath = PsPathTo : sFromToString = "To"
        Call aoParams.Add(sFromToString, _
            func_CmpExcelCopyAllSheetsToWorkbookForResultsDetail(aoWorkbookForResults, sPath, sFromToString))

    End Sub

    '***************************************************************************************************
    'Function/Sub Name           : func_CmpExcelCopyAllSheetsToWorkbookForResultsDetail()
    'Overview                    : 比較対象ファイルの全シートを比較結果用ワークブックにコピーする
    'Detailed Description        : ワークシートのリネーム情報のハッシュマップの構成
    '                              Key                      Value
    '                              -------------------      --------------------------------------------
    '                              Seq(1,2,3…)              変更前後のワークシート名格納用ハッシュマップ
    'Argument
    '     aoWorkbookForResults   : 比較結果用のワークブック
    '     asPath                 : 比較対象ファイルのパス
    '     asFromToString         : 比較元先を識別する文字列 "From","To"
    'Return Value
    '     ワークシートのリネーム情報のハッシュマップ
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

        '比較対象ファイルを開く
        Dim oExcel : Set oExcel = aoWorkbookForResults.Parent
        Dim oWorkBook : Set oWorkBook = func_CM_ExcelOpenFile(oExcel, asPath)
        Dim sTempPath : sTempPath = ""
        If oWorkBook.HasVBProject Then
        'マクロありの場合は別名で保存した上で再度開く）
            sTempPath = func_CmpExcelGetTempFilePath()
            Call sub_CM_ExcelSaveAs(oWorkBook, sTempPath, vbNullString)
            Set oWorkBook = func_CM_ExcelOpenFile( oExcel, sTempPath)
        End If

        '文書の保護を解除する
        Call sub_CM_OfficeUnprotect(oWorkBook, vbNullString)
        
        With oWorkBook
            'ワークシートのリネーム情報格納用ハッシュマップ定義
            Dim oWorkSheetRenameInfo : Set oWorkSheetRenameInfo = CreateObject("Scripting.Dictionary")
            'タブの色変換用ハッシュマップ定義
            Dim oStringToThemeColor : Set oStringToThemeColor = CreateObject("Scripting.Dictionary")
            Call oStringToThemeColor.Add("From",2)
            Call oStringToThemeColor.Add("To",8)

            Dim oWorksheet
            Dim lCnt : lCnt = 0
            For Each oWorksheet In .Worksheets
                If oWorksheet.Visible Then
                '全ての見えるシートを比較結果用ワークブックにコピーする
                    
                    '変更前後のワークシート名を取得
                    lCnt = lCnt + 1
                    Call oWorkSheetRenameInfo.Add( _
                                    lCnt, func_CmpExcelGetMapWorkSheetRenameInfo( _
                                                            oWorksheet.Name _
                                                            , func_CmpExcelMakeSheetName( _
                                                                                    lCnt _
                                                                                    , asFromToString _
                                                                                    ) _
                                                            ) _
                                    )
                    
                    'シートの表示を整える
                    If oWorksheet.AutoFilterMode Then
                    'オートフィルタが設定されていたら解除する
                         oWorksheet.Cells(1,1).AutoFilter
                    End If
                    oWorksheet.Activate
                    .Windows(1).View = 1                      'xlNormalView 標準
                    .Windows(1).Zoom = 25                     '表示倍率
                    .Windows(1).ScrollColumn = 1              '列1が左端になるようにウィンドウをスクロール
                    .Windows(1).ScrollRow = 1                 '行1が上端になるようにウィンドウをスクロール
                    .Windows(1).FreezePanes = False           'ウィンドウ枠の固定解除

                    'シート名を変更、タブの色を変更
                    oWorksheet.Name = oWorkSheetRenameInfo.Item(lCnt).Item("After")
                    oWorksheet.Tab.ThemeColor = oStringToThemeColor.Item(asFromToString)
                    oWorksheet.Tab.TintAndShade = 0

                    'シートを比較結果用の新規ワークブックにコピー
                    Call oWorksheet.Copy(, aoWorkbookForResults.Worksheets(aoWorkbookForResults.Worksheets.Count))
                End If
            Next

            '比較対象ファイルを閉じる
            Call .Close(False)
        End With
        
        If Len(sTempPath) Then
        'マクロありの場合に別名で保存したファイルがあったら削除する
            Call func_CM_FsDeleteFile(sTempPath)
        End If

        'サマリーシートのカラム位置変換用ハッシュマップ定義
        Dim oStringToColumn : Set oStringToColumn = CreateObject("Scripting.Dictionary")
        Call oStringToColumn.Add("From",1)
        Call oStringToColumn.Add("To",2)
        
        'サマリーシートに比較対象ファイルの情報を出力
        Dim lRow : Dim lColumn : Dim oItem
        lColumn = oStringToColumn.Item(asFromToString)
        With aoWorkbookForResults.Worksheets.Item(1)
            'ファイルパス
            lRow = 1
            .Cells(lRow, lColumn).Value = asPath
            'シート名
            For Each oItem In oWorkSheetRenameInfo.Items
                lRow = lRow + 1
                .Cells(lRow, lColumn).Value = oItem.Item("Before")
            Next
        End With

        'ワークシートのリネーム情報を返却
        Set func_CmpExcelCopyAllSheetsToWorkbookForResultsDetail = oWorkSheetRenameInfo

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
    'Function/Sub Name           : func_CmpExcelGetMapWorkSheetRenameInfo()
    'Overview                    : 変更前後のワークシート名格納用ハッシュマップ作成
    'Detailed Description        : 変更前後のワークシート名格納用ハッシュマップの構成
    '                              Key                      Value
    '                              -------------------      --------------------------------------------
    '                              "Before"                 変更前のシート名
    '                              "After"                  変更後のシート名
    'Argument
    '     asBefore               : 変更前のシート名
    '     asAfter                : 変更後のシート名
    'Return Value
    '     変更前後のワークシート名格納用ハッシュマップ
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2017/04/26         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function func_CmpExcelGetMapWorkSheetRenameInfo( _
        byVal asBefore _
        , byVal asAfter _
        )
        Dim oTemp : Set oTemp = CreateObject("Scripting.Dictionary")
        Call oTemp.Add("Before", asBefore)
        Call oTemp.Add("After", asAfter)
        Set func_CmpExcelGetMapWorkSheetRenameInfo = oTemp
        Set oTemp = Nothing
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : sub_CmpExcelCompare()
    'Overview                    : エクセルファイルを比較する
    'Detailed Description        : 工事中
    'Argument
    '     aoParams               : パラメータ格納用汎用ハッシュマップ
    '     aoWorkbookForResults   : 比較結果用のワークブック
    'Return Value
    '     なし
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2017/04/26         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub sub_CmpExcelCompare( _
        byRef aoWorkbookForResults _
        , byRef aoParams _
        )
        'ワークシートごとのシート名リネーム情報用ハッシュマップ
        Dim oFrom : Set oFrom = aoParams.Item("From")
        Dim oTo : Set oTo = aoParams.Item("To")

        Dim lCnt
        For lCnt = 1 To func_CM_MathMin(oFrom.Count, oTo.Count)
        '比較元先の各シートに差分が分かる書式設定をする
            '比較元（To）のシートに対し比較先（From）との差分が分かるようにする
            Call sub_CmpExcelSetFormatToUnderstandDifference(_
                    aoWorkbookForResults, oFrom.Item(lCnt).Item("After"), oTo.Item(lCnt).Item("After"))        
            '比較先（From）のシートに対し比較元（To）との差分が分かるようにする
            Call sub_CmpExcelSetFormatToUnderstandDifference( _
                    aoWorkbookForResults, oTo.Item(lCnt).Item("After"), oFrom.Item(lCnt).Item("After"))        
        Next

        '同じブックの新しいウィンドウを開く
        aoWorkbookForResults.Worksheets(oFrom.Item(1).Item("After")).Activate
        With aoWorkbookForResults.Windows(1).NewWindow
            Dim sCaption : sCaption = .Caption
            Dim oWorksheet
            For Each oWorksheet In .Parent.Worksheets
                oWorksheet.Activate
                .Zoom = 25
            Next
        End With
        aoWorkbookForResults.Worksheets(oTo.Item(1).Item("After")).Activate
        '並べて比較
        aoWorkbookForResults.Activate
        With aoWorkbookForResults.Parent
            .Windows.CompareSideBySideWith(sCaption)
            Call .Windows.Arrange(-4166, True)               'xlVertical = -4166
            .DisplayAlerts = True
            .ScreenUpdating = True
            .AutomationSecurity = 2                     'msoAutomationSecurityByUI = 2 [ セキュリティ] ダイアログ ボックスで指定されたセキュリティ設定を使用
            .Visible = True
        End With

        'オブジェクトを開放
        Set oWorksheet = Nothing
        Set oTo = Nothing
        Set oFrom = Nothing

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

        'オートシェイプの比較
        Dim oAutoshapeA : Dim oAutoshapeB
        For Each oAutoshapeA In aoWorkbookForResults.Worksheets(asSheetNameA).Shapes
            Set oAutoshapeB = func_CM_GetObjectByIdFromCollection(aoWorkbookForResults.Worksheets(asSheetNameA).Shapes, oAutoshapeA.Id)
            If Trim(func_CM_ExcelGetTextFromAutoshape(oAutoshapeA)) _
               = Trim(func_CM_ExcelGetTextFromAutoshape(oAutoshapeB)) Then
            'オートシェイプのIDとテキストが一致する（差異がない）場合は灰色にする
                Call sub_CmpExcelSetAutoshapeColor(oAutoshapeA)
            End If
        Next

        'オブジェクトを開放
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
            .ForeColor.ObjectTjemeColor = 14         '背景１テーマの色 msoThemeColorBackground1
            .ForeColor.TintAndShade = 0              '色を明るくするかまたは暗くする単精度浮動小数点型 (Single) の値
            .ForeColor.Brightness = -0.150000006     '明度
            .Transparency = 0                        '塗りつぶしの透明度を示す 0.0 (不透明) から 1.0 (透明) までの値
            .Solid                                   '塗りつぶしを均一な色に設定
        End With
        If Err.Number Then
            Err.Clear
        End If
    End Sub

    '***************************************************************************************************
    'Function/Sub Name           : func_CmpExcelGetTempFilePath()
    'Overview                    : 一時ファイルのフルパスを取得
    'Detailed Description        : 実行中のスクリプトファイルがあるフォルダの下にある
    '                              Cs_FOLDER_TEMP以下の一時ファイルのパスを返す
    '                              Cs_FOLDER_TEMPフォルダがない場合は作成する
    'Argument
    '     なし
    'Return Value
    '     一時ファイルのフルパス
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2017/04/26         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function func_CmpExcelGetTempFilePath()
        Dim sParentFolderPath : sParentFolderPath = func_CM_FsGetParentFolderPath(WScript.ScriptFullName)
        Dim sFolderPath : sFolderPath = func_CM_FsBuildPath(sParentFolderPath, Cs_FOLDER_TEMP)
        If Not(func_CM_FsFolderExists(sFolderPath)) Then func_CM_FsCreateFolder(sFolderPath)
        func_CmpExcelGetTempFilePath = func_CM_FsBuildPath(sFolderPath, func_CM_FsGetTempFileName())
    End Function

End Class
