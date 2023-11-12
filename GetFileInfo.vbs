'***************************************************************************************************
'FILENAME                    : GetFileInfo.vbs
'Overview                    : 引数のファイルの情報をHTMLで出力する
'Detailed Description        : Sendtoから使用する
'Argument
'     PATH1,2...             : ファイルのパス1,2,...
'Return Value
'     なし
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/11/05         Y.Fujii                  First edition
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
'2023/11/05         Y.Fujii                  First edition
'***************************************************************************************************
Sub Main()
    'ログ出力の設定
    Set PoWriter = new_WriterTo(func_CM_FsGetPrivateLogFilePath, 8, True, -2)
    'ブローカークラスのインスタンスの設定
    Dim oBroker : Set oBroker = new_Broker()
    oBroker.subscribe "log", GetRef("sub_GetFileInfoLogger")
    'パラメータ格納用オブジェクト宣言
    Dim oParams : Set oParams = new_Dic()
    
    '当スクリプトの引数をパラメータ格納用オブジェクトに取得する
    sub_CM_ExcuteSub "sub_GetFileInfoGetParameters", oParams, oBroker
    
    'ファイル情報の取得
    sub_CM_ExcuteSub "sub_GetFileInfoProc", oParams, oBroker
    
    'ログ出力をクローズ
    PoWriter.close()
    
    'オブジェクトを開放
    Set oParams = Nothing
    Set oBroker = Nothing
    Set PoWriter = Nothing
End Sub

'***************************************************************************************************
'Processing Order            : 1
'Function/Sub Name           : sub_GetFileInfoGetParameters()
'Overview                    : 当スクリプトの引数をパラメータ格納用オブジェクトに取得する
'Detailed Description        : パラメータ格納用汎用オブジェクトにKey="Param"で格納する
'                              配列（clsCmArray型）に名前なし引数（/Key:Value 形式でない）を全て
'                              取得する
'Argument
'     aoParams               : パラメータ格納用オブジェクト
'Return Value
'     なし
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/11/05         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub sub_GetFileInfoGetParameters( _
    byRef aoParams _
    )
    'オリジナルの引数を取得
    Dim oArg : Set oArg = func_CM_UtilStoringArguments()
    '★ログ出力
    sub_GetFileInfoLogger Array(9, "sub_GetFileInfoGetParameters", func_CM_ToStringArguments())
    
    'パラメータ格納用オブジェクトに設定
    Dim oParam, oRet, oItem
    Set oParam = new_Arr()
    For Each oItem In oArg.Item("Unnamed").Items()
        Set oRet = cf_tryCatch(Getref("new_FileOf"), oItem, Empty, Empty)
        If Not oRet.Item("Result") Then Set oRet = cf_tryCatch(Getref("new_FolderOf"), oItem, Empty, Empty)
        If oRet.Item("Result") Then oParam.push oRet.Item("Return")
    Next
    cf_bindAt aoParams, "Param", oParam
    
    Set oItem = Nothing
    Set oRet = Nothing
    Set oParam = Nothing
    Set oArg = Nothing
End Sub

'***************************************************************************************************
'Processing Order            : 2
'Function/Sub Name           : sub_GetFileInfoProc()
'Overview                    : 引数のファイルパスをクリップボードに出力する
'Detailed Description        : 工事中
'Argument
'     aoParams               : パラメータ格納用オブジェクト
'Return Value
'     なし
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/11/05         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub sub_GetFileInfoProc( _
    byRef aoParams _
    )
    'パラメータ格納用汎用オブジェクト
    Dim oParam : Set oParam = aoParams.Item("Param").slice(0,vbNullString)

    'ファイル情報を取得
    Dim oList : Set oList = new_Arr()
    Do While oParam.length>0
        oList.pushMulti func_GetFileInfoProcGetFilesRecursion(oParam.pop)
    Loop

    '重複を排除してpath順にソートする
    cf_bindAt aoParams, "List", oList.uniq().sortUsing(new_Func("(c,n)=>c.Path>n.Path"))

    Set oList = Nothing
    Set oParam = Nothing
End Sub

'***************************************************************************************************
'Processing Order            : 2-1
'Function/Sub Name           : func_GetFileInfoProcGetFilesRecursion()
'Overview                    : フォルダ配下の全ファイルを取得する
'Detailed Description        : 工事中
'Argument
'     aoItem                 : ファイル/フォルダオブジェクト
'Return Value
'     ファイルオブジェクトの配列
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/11/12         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_GetFileInfoProcGetFilesRecursion( _
    byRef aoItem _
    )
    If cf_isSame(TypeName(aoItem), "Folder") Then
    'フォルダの場合
        Dim oEle, vRet
        'ファイルの取得
        For Each oEle In aoItem.Files
            cf_push vRet, oEle
        Next
        'フォルダの取得
        For Each oEle In aoItem.SubFolders
            cf_pushMulti vRet, func_GetFileInfoProcGetFilesRecursion(oEle)
        Next
        func_GetFileInfoProcGetFilesRecursion = vRet
    Else
    'ファイルの場合
        func_GetFileInfoProcGetFilesRecursion = Array(aoItem)
    End If

End Function


'***************************************************************************************************
'Processing Order            : -
'Function/Sub Name           : sub_GetFileInfoLogger()
'Overview                    : ログ出力する
'Detailed Description        : sub_CM_UtilLogger()に委譲する
'Argument
'     avParams               : 配列型のパラメータリスト
'Return Value
'     なし
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/11/05         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub sub_GetFileInfoLogger( _
    byRef avParams _
    )
    sub_CM_UtilLogger avParams, PoWriter
End Sub
