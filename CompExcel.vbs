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
Private Const Cs_FOLDER_TEMP = "tmp"

'Include用関数定義
Sub sub_Include( _
    byVal asIncludeFileName _
    )
    Dim oFso : Set oFso = CreateObject("Scripting.FileSystemObject")
    Dim sParentFolderName : sParentFolderName = oFso.GetParentFolderName(WScript.ScriptFullName)
    Dim sIncludeFilePath
    sIncludeFilePath = oFso.BuildPath(sParentFolderName, Cs_FOLDER_INCLUDE)
    sIncludeFilePath = oFso.BuildPath(sIncludeFilePath, asIncludeFileName)
    ExecuteGlobal oFso.OpenTextfile(sIncludeFilePath).ReadAll
    Set oFso = Nothing
End Sub
'Include
Call sub_Include("VbsBasicLibCommon.vbs")

Main

Wscript.Quit


'***************************************************************************************************
'Processing Order            : 0
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
    
    Dim oParams : Set oParams = CreateObject("Scripting.Dictionary")
    
    '当スクリプトの引数取得
    Call sub_CmpExcelGetParameters( _
                            oParams _
                             )
    
    '比較対象ファイル入力画面の表示と取得
    Call sub_CmpExcelDispInputFiles( _
                            oParams _
                             )
    
    'エクセルファイルを比較する
    Call sub_CmpExcelCompareFiles( _
                            oParams _
                             )
    
    'オブジェクトを開放
    Set oParams = Nothing
    
End Sub

'***************************************************************************************************
'Processing Order            : 1
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
    Dim oParameter : Set oParameter = CreateObject("Scripting.Dictionary")
    Dim lCnt : lCnt = 0
    Dim sParam
    For Each sParam In WScript.Arguments
        If func_CM_FsFileExists(sParam) Then
        'ファイルが存在する場合パラメータを取得
            lCnt = lCnt + 1
            Call oParameter.Add(lCnt, sParam)
        End If
    Next
    
    Call aoParams.Add("Parameter", oParameter)
End Sub

'***************************************************************************************************
'Processing Order            : 2
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
    
    Dim oExcel : Set oExcel = CreateObject("Excel.Application")
    With oExcel
        Dim sPath
        Do Until oParameter.Count > 1
            
            sPath = .GetOpenFilename( , , Cs_TITLE_EXCEL, , False)
            If sPath = False Then
            'ファイル選択キャンセルの場合は当スクリプトを終了する
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
    Set oExcel = Nothing
    Set oParameter = Nothing
End Sub

'***************************************************************************************************
'Processing Order            : 3
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
 '   On Error Resume Next
    
    '3-1 比較するファイルを古い順（最終更新日昇順）に並べ替える
    Call sub_CmpExcelSortByDateLastModified(aoParams)
    
    '比較結果用の新規ワークブックを作成
    Dim oWorkbookForResults
    With CreateObject("Excel.Application")
        .DisplayAlerts = False
        .ScreenUpdating = False
        .AutomationSecurity = 3                              'msoAutomationSecurityForceDisable = 3
        Set oWorkbookForResults = .Workbooks.Add(-4167)      '新規ワークブック xlWBATWorksheet=-4167
    End With
    
    '3-2 比較対象ファイルの全シートを比較結果用ワークブックにコピーする
    Call sub_CmpExcelCopyAllSheetsToWorkbookForResults(aoParams, oWorkbookForResults)
    
    '3-3 エクセルファイルを比較する
    Call sub_CmpExcelCompare(aoParams, oWorkbookForResults)
    
    'オブジェクトを開放
    Set oWorkbookForResults = Nothing

End Sub

'***************************************************************************************************
'Processing Order            : 3-1
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
    
    If func_CM_FsGetFile(oParameter.Item(1)).DateLastModified _
        <= _
        func_CM_FsGetFile(oParameter.Item(2)).DateLastModified _
        Then
    '最初のファイルの方が古い（最終更新日が小さい）場合、処理を抜ける
        Exit Sub
    End If
    
    '値を入れ替える
    With oParameter
        Dim sValue1 : Dim sValue2
        sValue1 = .Item(1)
        sValue2 = .Item(2)
        
        .RemoveAll
        
        Call .Add(1, sValue2)
        Call .Add(2, sValue1)
    End With
    
    'オブジェクトを開放
    Set oParameter = Nothing
End Sub

'***************************************************************************************************
'Processing Order            : 3-2
'Function/Sub Name           : sub_CmpExcelCopyAllSheetsToWorkbookForResults()
'Overview                    : 比較対象ファイルの全シートを比較結果用ワークブックにコピーする
'Detailed Description        : パラメータ格納用汎用ハッシュマップにKey="WorkSheetRenameInfos"で格納する
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
    byRef aoParams _
    , byRef aoWorkbookForResults _
    )
    'パラメータ格納用汎用ハッシュマップ
    Dim oParameter : Set oParameter = aoParams.Item("Parameter")
    'ワークシートごとのシート名リネーム情報格納用ハッシュマップ
    Dim oWorkSheetRenameInfos : Set oWorkSheetRenameInfos = CreateObject("Scripting.Dictionary")
    
    Dim sPath : Dim sFromToString
    '3-2-1 比較元ファイルのコピー
    sPath = oParameter.Item(1) : sFromToString = "From" 
    Call oWorkSheetRenameInfos.Add(sFromToString, _
        func_CmpExcelCopyAllSheetsToWorkbookForResultsDetail(aoWorkbookForResults, sPath, sFromToString))

    '3-2-1 比較先ファイルのコピー
    sPath = oParameter.Item(2) : sFromToString = "To"
    Call oWorkSheetRenameInfos.Add(sFromToString, _
        func_CmpExcelCopyAllSheetsToWorkbookForResultsDetail(aoWorkbookForResults, sPath, sFromToString))
    
    'ワークシートごとのシート名リネーム情報格納用ハッシュマップをパラメータ格納用汎用ハッシュマップに格納
    Call aoParams.Add("WorkSheetRenameInfos", oWorkSheetRenameInfos)

    'オブジェクトを開放
    Set oWorkSheetRenameInfos = Nothing
    Set oParameter = Nothing
End Sub

'***************************************************************************************************
'Processing Order            : 3-2-1
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
'Processing Order            : 3-2-1
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
'Processing Order            : 3-2-1
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
'Processing Order            : 3-3
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
    byRef aoParams _
    , byRef aoWorkbookForResults _
    )
    'ワークシートごとのシート名リネーム情報用ハッシュマップ
    Dim oWorkSheetRenameInfos : Set oWorkSheetRenameInfos = aoParams.Item("WorkSheetRenameInfos")
    Dim oFrom : Set oFrom = oWorkSheetRenameInfos.Item("From")
    Dim oTo : Set oTo = oWorkSheetRenameInfos.Item("To")

    Dim lCnt
    For lCnt = 1 To func_CM_MathMin(oFrom.Count, oTo.Count)
    '比較元先の各シートに差分が分かる書式設定をする
        '3-3-1 比較元（To）のシートに対し比較先（From）との差分が分かるようにする
        Call sub_CmpExcelSetFormatToUnderstandDifference(_
                aoWorkbookForResults, oFrom.Item(lCnt).Item("After"), oTo.Item(lCnt).Item("After"))        
        '3-3-1 比較先（From）のシートに対し比較元（To）との差分が分かるようにする
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
    Set oWorkSheetRenameInfos = Nothing

End Sub

'***************************************************************************************************
'Processing Order            : 3-3-1
'Function/Sub Name           : sub_CmpExcelSetFormatToUnderstandDifference()
'Overview                    : asSheetNameAのシートにasSheetNameBシートとの差分が分かる書式設定をする
'Detailed Description        : 工事中
'Argument
'     aoWorkbookForResults   : 比較結果用のワークブック
'     asSheetNameA           : 比較元のシート名
'     asSheetNameA           : 比較先のシート名
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
'Processing Order            : 3-3-1
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
'Processing Order            : -
'Function/Sub Name           : func_CmpExcelGetTempFilePath()
'Overview                    : 一時ファイルのフルパスを取得
'Detailed Description        : 工事中
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
    func_CmpExcelGetTempFilePath = func_CM_FsBuildPath(sFolderPath, func_CM_FsGetTempFileName())
End Function
