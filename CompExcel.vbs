'***************************************************************************************************
'FILENAME                    : CompExcel.vbs
'Overview                    : エクセルファイルを比較する
'Detailed Description        : 引数で指定されたエクセルファイルを比較対象とする
'                              指定がないまたは1つだけの場合は、ダイアログで比較対象の入力を求める
'Argument                    : 名前なし引数（/Key:Value 形式でない）のみ
'                                1,2番目   : 比較するエクセルファイルのパス（ともに省略可能）
'                                3番目以降 : 無視する
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
Private Const Cs_FOLDER_LIB = "lib"
Private PoWriter, PoBroker

'import定義
Sub sub_import( _
    byVal asIncludeFileName _
    )
    With CreateObject("Scripting.FileSystemObject")
        Dim sParentFolderName : sParentFolderName = .GetParentFolderName(WScript.ScriptFullName)
        Dim sIncludeFilePath
        sIncludeFilePath = .BuildPath(sParentFolderName, Cs_FOLDER_LIB)
        sIncludeFilePath = .BuildPath(sIncludeFilePath, asIncludeFileName)
        ExecuteGlobal .OpenTextfile(sIncludeFilePath).ReadAll
    End With
End Sub
'import
sub_import "clsCmArray.vbs"
sub_import "clsCmBroker.vbs"
sub_import "clsCmBufferedReader.vbs"
sub_import "clsCmBufferedWriter.vbs"
sub_import "clsCmCalendar.vbs"
sub_import "clsCmCharacterType.vbs"
sub_import "clsCmCssGenerator.vbs"
sub_import "clsCmHtmlGenerator.vbs"
sub_import "libCom.vbs"
sub_import "clsCompareExcel.vbs"

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
    Set PoWriter = new_WriterTo(func_CM_FsGetPrivateLogFilePath(), 8, True, -1)
    'ブローカークラスのインスタンスの設定
    Set PoBroker = new_Broker()
    PoBroker.subscribe "log", GetRef("sub_CmpExcelLogger")
    'パラメータ格納用汎用オブジェクト宣言
    Dim oParams : Set oParams = new_Dic()
    
    '当スクリプトの引数取得
    fw_excuteSub "sub_CmpExcelGetParameters", oParams, PoBroker
    
    '比較対象ファイル入力画面の表示と取得
    fw_excuteSub "sub_CmpExcelDispInputFiles", oParams, PoBroker
    
    'エクセルファイルを比較する
    fw_excuteSub "sub_CmpExcelCompareFiles", oParams, PoBroker
    
    'ログ出力をクローズ
    PoWriter.close
    
    'オブジェクトを開放
    Set oParams = Nothing
    Set PoBroker = Nothing
    Set PoWriter = Nothing
End Sub

'***************************************************************************************************
'Processing Order            : 1
'Function/Sub Name           : sub_CmpExcelGetParameters()
'Overview                    : 当スクリプトの引数取得
'Detailed Description        : パラメータ格納用汎用オブジェクトにKey="Param"で格納する
'                              配列（clsCmArray型）に名前なし引数（/Key:Value 形式でない）があるれば
'                              2番目まで取得する
'                              名前なし引数の3番目以降あるいは名前付き引数（/Key:Value 形式）は無視する
'                              Index   Contents
'                              -----   -------------------------------------------------------------
'                              0       名前なし引数の1番目
'                              1       名前なし引数の2番目
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
Private Sub sub_CmpExcelGetParameters( _
    byRef aoParams _
    )
    'オリジナルの引数を取得
    Dim oArg : Set oArg = fw_storeArguments()
    '★ログ出力
    sub_CmpExcelLogger Array(9, "sub_CmpExcelGetParameters", cf_toString(oArg))
    
    'パラメータ格納用オブジェクトに設定
    cf_bindAt aoParams, "Param", new_ArrWith(oArg.Item("Unnamed")).slice(0,2)
    
    Set oArg = Nothing
End Sub

'***************************************************************************************************
'Processing Order            : 2
'Function/Sub Name           : sub_CmpExcelDispInputFiles()
'Overview                    : 比較対象ファイル入力画面の表示と取得
'Detailed Description        : 引数で比較するエクセルファイルの指定がない場合、Excel.Applicationの
'                              ダイアログを表示してユーザにファイルを選択させる
'                              Index   Contents
'                              -----   -------------------------------------------------------------
'                              0       Excel.Applicationのダイアログで選択したファイルパスを設定する
'                              1       同上
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
Private Sub sub_CmpExcelDispInputFiles( _
    byRef aoParams _
    )
    Dim oParam : Set oParam = aoParams.Item("Param")
    If oParam.length > 1 Then
    'パラメータが2個以上だったら関数を抜ける
        '★ログ出力
        sub_CmpExcelLogger Array(3, "sub_CmpExcelDispInputFiles", "No dialog required.")
        Exit Sub
    End If
    
    'パラメータ格納用汎用オブジェクト
    Const Cs_TITLE_EXCEL = "比較対象ファイルを開く"
    With CreateObject("Excel.Application")
        Dim sPath
        Do Until oParam.Length > 1
            
            sPath = .GetOpenFilename( , , Cs_TITLE_EXCEL, , False)
            If sPath = False Then
            'ファイル選択キャンセルの場合は当スクリプトを終了する
                '★ログ出力
                sub_CmpExcelLogger Array(3, "sub_CmpExcelDispInputFiles", "Dialog input canceled.")
                
                PoWriter.close
                Set oParam = Nothing
                Wscript.Quit
            End If
            '選択したファイルのパスを取得
            oParam.push sPath
        Loop
        
        .Quit
    End With
    
    'オブジェクトを開放
    Set oParam = Nothing
End Sub

'***************************************************************************************************
'Processing Order            : 3
'Function/Sub Name           : sub_CmpExcelCompareFiles()
'Overview                    : エクセルファイルを比較する
'Detailed Description        : エラーは無視する
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
Private Sub sub_CmpExcelCompareFiles( _
    byRef aoParams _
    )
    'パラメータ格納用汎用オブジェクト
    Dim oParam : Set oParam = aoParams.Item("Param")
    
    'ファイルの最終更新日昇順に並べ替える
    oParam.sortUsing new_Func("(c,n)=>new_CalAt(new_FileOf(c).DateLastModified).compareTo(new_CalAt(new_FileOf(n).DateLastModified))>0")
    '★ログ出力
    sub_CmpExcelLogger Array(3, "sub_CmpExcelCompareFiles", "aoParams sorted.")
    sub_CmpExcelLogger Array(9, "sub_CmpExcelCompareFiles", "aoParams is " & cf_toString(aoParams))
    
    '比較
    With New clsCompareExcel
        Set .broker = PoBroker
        .pathFrom = oParam(0)
        .pathTo = oParam(1)
        .compare()
    End With
    
    'オブジェクトを開放
    Set oParam = Nothing
End Sub

'***************************************************************************************************
'Processing Order            : -
'Function/Sub Name           : sub_CmpExcelLogger()
'Overview                    : ログ出力する
'Detailed Description        : fw_logger()に委譲する
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
    fw_logger avParams, PoWriter
End Sub
