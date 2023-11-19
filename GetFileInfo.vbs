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
    
    '結果出力
    sub_GetFileInfoReport oParams
'    sub_CM_ExcuteSub "sub_GetFileInfoReport", oParams, oBroker
    
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

    'ファイルオブジェクトのリストを取得
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
'Overview                    : フォルダ配下のファイルオブジェクトを取得する
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

    'レポートの作成
    Dim oHtml : Set oHtml = new_HtmlOf("html")
    oHtml.addContent func_GetFileInfoReportHtmlHead(aoParams)
    oHtml.addContent func_GetFileInfoReportHtmlBody(aoParams)

    'ファイル出力
    Dim sPath
    sPath = func_CM_FsGetPrivateFilePath("report", new_Fso().GetBaseName(WScript.ScriptName) & new_Now().formatAs("_YYMMDD_HHmmSS_000") & ".html")
    sub_CM_FsWriteFile sPath, oHtml.generate

    Set oHtml = Nothing
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
    
    Dim oStyle : Set oStyle = new_HtmlOf("style").addAttribute("type", "text/css").addAttribute("media", "all")
    With oStyle
        .addContent new_CssOf(".table_wrap").addProperty("overflow", "auto").addProperty("height", "300px")
        .addContent new_CssOf("table.table01").addProperty("width", "1000px").addProperty("min-width", "635px").addProperty("margin", "10px 0").addProperty("font-size", "1.4rem").addProperty("border-spacing", "0px").addProperty("border-collapse", "separate")
        .addContent new_CssOf("table.table01 th").addProperty("background-color", "#6b6b6b").addProperty("color", "#fff").addProperty("padding", "10px").addProperty("border-bottom", "1px solid #E0E1E3").addProperty("border-right", "1px solid #E0E1E3").addProperty("position", "sticky").addProperty("top", "0").addProperty("left", "0").addProperty("z-index", "1")
        .addContent new_CssOf("table.table01 thead table th").addProperty("border-top", "1px solid #E0E1E3")
        .addContent new_CssOf("table.table01 thead tr:first-of-type th:first-of-type").addProperty("z-index", "2")
        .addContent new_CssOf("table.table01 tbody td").addProperty("padding", "10px").addProperty("font-weight", "normal").addProperty("border-bottom", "1px solid #E0E1E3").addProperty("border-right", "1px solid #E0E1E3")
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
    
    'theader
    Dim oTr : Set oTr = new_HtmlOf("tr")
    oTr.addContent new_HtmlOf("th").addContent("Seq")
    oTr.addContent new_HtmlOf("th").addContent("Attributes")
    oTr.addContent new_HtmlOf("th").addContent("DateCreated")
    oTr.addContent new_HtmlOf("th").addContent("DateLastAccessed")
    oTr.addContent new_HtmlOf("th").addContent("DateLastModified")
    oTr.addContent new_HtmlOf("th").addContent("Drive")
    oTr.addContent new_HtmlOf("th").addContent("Name")
    oTr.addContent new_HtmlOf("th").addContent("ParentFolder")
    oTr.addContent new_HtmlOf("th").addContent("Path")
    oTr.addContent new_HtmlOf("th").addContent("ShortName")
    oTr.addContent new_HtmlOf("th").addContent("ShortPath")
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
            oTr.addContent new_HtmlOf("td").addContent(.Attributes)
            oTr.addContent new_HtmlOf("td").addContent(.DateCreated)
            oTr.addContent new_HtmlOf("td").addContent(.DateLastAccessed)
            oTr.addContent new_HtmlOf("td").addContent(.DateLastModified)
            oTr.addContent new_HtmlOf("td").addContent(.Drive)
            oTr.addContent new_HtmlOf("td").addContent(.Name)
            oTr.addContent new_HtmlOf("td").addContent(.ParentFolder)
            oTr.addContent new_HtmlOf("td").addContent(.Path)
            oTr.addContent new_HtmlOf("td").addContent(.ShortName)
            oTr.addContent new_HtmlOf("td").addContent(.ShortPath)
            oTr.addContent new_HtmlOf("td").addContent(.Size)
            oTr.addContent new_HtmlOf("td").addContent(.Type)
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
