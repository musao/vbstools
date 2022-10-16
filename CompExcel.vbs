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
Call sub_Include("VbsBasicLibCommon.vbs")
Call sub_Include("clsCompareExcel.vbs")


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
    
    '3-1 比較するファイルを古い順（最終更新日昇順）に並べ替える
    Call sub_CmpExcelSortByDateLastModified(aoParams)
    
    '3-2 比較
    With New clsCompareExcel
        .PathFrom = aoParams.Item("Parameter").Item(1)
        .PathTo = aoParams.Item("Parameter").Item(2)
        .Compare()
    End With

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
