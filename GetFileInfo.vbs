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
sub_import "clsAdptFile.vbs"
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
    PoWriter.writeBufferSize=100000
    'ブローカークラスのインスタンスの設定
    Dim oBroker : Set oBroker = new_Broker()
    oBroker.subscribe "log", GetRef("sub_GetFileInfoLogger")
    'パラメータ格納用オブジェクト宣言
    Dim oParams : Set oParams = new_Dic()
    
    '当スクリプトの引数をパラメータ格納用オブジェクトに取得する
    sub_CM_ExcuteSub "sub_GetFileInfoGetParameters", oParams, oBroker
    
    'ファイル情報の取得
    sub_CM_ExcuteSub "sub_GetFileInfoProc", oParams, oBroker
    
    '結果出力
    sub_CM_ExcuteSub "sub_GetFileInfoReport", oParams, oBroker
    
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
    
    '実在するパスだけパラメータ格納用オブジェクトに設定
    Dim oParam, oRet, oItem
    Set oParam = new_Arr()
    For Each oItem In oArg.Item("Unnamed").Items()
        Set oRet = cf_tryCatch(Getref("new_FileOf"), oItem, Empty, Empty)
        If Not oRet.Item("Result") Then Set oRet = cf_tryCatch(Getref("new_FolderOf"), oItem, Empty, Empty)
        If oRet.Item("Result") Then
            oParam.push oRet.Item("Return")
        Else
            '★ログ出力
            sub_GetFileInfoLogger Array(3, "sub_GetFileInfoGetParameters", oItem & "is an invalid argument.")
        End If
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
'Overview                    : ファイル情報の取得
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

    '★ログ出力
    sub_GetFileInfoLogger Array(3, "sub_GetFileInfoProc", "Before getting list of files.")
    'ファイルオブジェクトのリストを取得
    Dim oList : Set oList = new_Arr()
    Do While oParam.length>0
'        oList.pushMulti fs_getAllFiles(oParam.pop().Path)
        oList.pushMulti fs_getAllFilesByShell(oParam.pop().Path)
    Loop

    '★ログ出力
    sub_GetFileInfoLogger Array(3, "sub_GetFileInfoProc", "Before sorting.")
    '重複を排除してpath順にソートする
    cf_bindAt aoParams, "List", oList.uniq().sortUsing(new_Func("(c,n)=>c.ParentFolder&c.Path>n.ParentFolder&n.Path"))

    Set oList = Nothing
    Set oParam = Nothing
End Sub

'***************************************************************************************************
'Processing Order            : 3
'Function/Sub Name           : sub_GetFileInfoReport()
'Overview                    : 結果出力
'Detailed Description        : 工事中
'Argument
'     aoParams               : パラメータ格納用オブジェクト
'Return Value
'     なし
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/11/12         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub sub_GetFileInfoReport( _
    byRef aoParams _
    )
    If aoParams.Item("List").length=0 Then
    'パラメータ格納用汎用オブジェクトが空の場合
        '★ログ出力
        sub_GetFileInfoLogger Array(3, "sub_GetFileInfoReport", "There was no files.")
        '何もせず処理を抜ける
        Exit Sub
    End If

    'レポートの作成
    With new_HtmlOf("html")
        .addContent func_GetFileInfoReportHtmlHead(aoParams)
        .addContent func_GetFileInfoReportHtmlBody(aoParams)
    
        '★ログ出力
        sub_GetFileInfoLogger Array(3, "sub_GetFileInfoReport", "Before reportfile output.")
        'レポートをファイルに出力
        Dim sPath
        sPath = func_CM_FsGetPrivateFilePath("report", new_Fso().GetBaseName(WScript.ScriptName) & new_Now().formatAs("_YYMMDD_HHmmSS_000") & ".html")
        sub_CM_FsWriteFile sPath, .generate
    End With

    '★ログ出力
    sub_GetFileInfoLogger Array(3, "sub_GetFileInfoReport", "Before open reportfile.")
    'レポートを開く
    CreateObject("WScript.Shell").Run sPath, 1
    
End Sub

'***************************************************************************************************
'Processing Order            : 3-1
'Function/Sub Name           : func_GetFileInfoReportHtmlHead()
'Overview                    : 結果HTMLのheadタグ内の編集
'Detailed Description        : 工事中
'Argument
'     aoParams               : パラメータ格納用オブジェクト
'Return Value
'     なし
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/11/12         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_GetFileInfoReportHtmlHead( _
    byRef aoParams _
    )
    'パラメータ格納用汎用オブジェクト
    Dim oList : Set oList = aoParams.Item("List").slice(0,vbNullString)
    
    Dim oStyle : Set oStyle = _
        new_HtmlOf("style") _
            .addAttribute("type", "text/css") _
            .addAttribute("media", "all")
    
    With oStyle
        .addContent new_CssOf(".table_wrap") _
            .addProperty("overflow", "auto") _
            .addProperty("height", "100%")
        
        .addContent new_CssOf("table.table01") _
            .addProperty("width", "1000px") _
            .addProperty("min-width", "635px") _
            .addProperty("margin", "10px 0") _
            .addProperty("font-size", "1.4rem") _
            .addProperty("border-spacing", "0px") _
            .addProperty("border-collapse", "separate")
        
        .addContent new_CssOf("table.table01 th") _
            .addProperty("background-color", "#6b6b6b") _
            .addProperty("color", "#fff") _
            .addProperty("padding", "10px") _
            .addProperty("border-bottom", "1px solid #E0E1E3") _
            .addProperty("border-right", "1px solid #E0E1E3") _
            .addProperty("position", "sticky") _
            .addProperty("top", "0") _
            .addProperty("left", "0") _
            .addProperty("z-index", "1")
        
        .addContent new_CssOf("table.table01 thead table th") _
            .addProperty("border-top", "1px solid #E0E1E3")
        
        .addContent new_CssOf("table.table01 thead tr:first-of-type th:first-of-type") _
            .addProperty("z-index", "2")
        
        .addContent new_CssOf("table.table01 tbody td") _
            .addProperty("padding", "10px") _
            .addProperty("font-weight", "normal") _
            .addProperty("border-bottom", "1px solid #E0E1E3") _
            .addProperty("border-right", "1px solid #E0E1E3")
    End With

    Dim oHead : Set oHead = new_HtmlOf("head")
    oHead.addContent oStyle

    Set func_GetFileInfoReportHtmlHead = oHead
    Set oStyle = Nothing
    Set oHead = Nothing
End Function

'***************************************************************************************************
'Processing Order            : 3-2
'Function/Sub Name           : func_GetFileInfoReportHtmlBody()
'Overview                    : 結果HTMLのbodyタグ内の編集
'Detailed Description        : 工事中
'Argument
'     aoParams               : パラメータ格納用オブジェクト
'Return Value
'     なし
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/11/12         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_GetFileInfoReportHtmlBody( _
    byRef aoParams _
    )
    'パラメータ格納用汎用オブジェクト
    Dim oList : Set oList = aoParams.Item("List").slice(0,vbNullString)
    
    'thead
    Dim oTr : Set oTr = new_HtmlOf("tr")
    oTr.addContent new_HtmlOf("th").addContent("Seq")
    oTr.addContent new_HtmlOf("th").addContent("DateLastModified")
    oTr.addContent new_HtmlOf("th").addContent("Name")
    oTr.addContent new_HtmlOf("th").addContent("Path")
    oTr.addContent new_HtmlOf("th").addContent("ParentFolder")
    oTr.addContent new_HtmlOf("th").addContent("Size")
    oTr.addContent new_HtmlOf("th").addContent("Type")
    Dim oThead : Set oThead = new_HtmlOf("thead")
    oThead.addContent oTr

    'tbody
    Dim oTbody : Set oTbody = new_HtmlOf("tbody")
    Dim lSeq : lSeq=1
    Do While oList.length>0
        Set oTr = new_HtmlOf("tr")
        With oList.shift
            oTr.addContent new_HtmlOf("th").addContent(lSeq)
            oTr.addContent new_HtmlOf("td").addContent(.DateLastModified)
            oTr.addContent new_HtmlOf("td").addContent(.Name)
            oTr.addContent new_HtmlOf("td").addContent(.Path)
            oTr.addContent new_HtmlOf("td").addContent(.ParentFolder)
            oTr.addContent new_HtmlOf("td").addContent(.Size)
            oTr.addContent new_HtmlOf("td").addContent(.FileType)
        End With
        oTbody.addContent oTr
        lSeq = lSeq+1
    Loop
    Dim oTable : Set oTable = new_HtmlOf("table").addAttribute("class", "table01")
    oTable.addContent oThead
    oTable.addContent oTbody

    Dim oDiv : Set oDiv = new_HtmlOf("div").addAttribute("class", "table_wrap")
    oDiv.addContent oTable

    Dim oBody : Set oBody = new_HtmlOf("body")
    oBody.addContent oDiv

    Set func_GetFileInfoReportHtmlBody = oBody
    Set oTr = Nothing
    Set oThead = Nothing
    Set oTbody = Nothing
    Set oTable = Nothing
    Set oBody = Nothing
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
