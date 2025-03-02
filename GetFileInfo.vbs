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

'変数
Private PoWriter, PoBroker

'lib import
Private Const Cs_FOLDER_LIB = "lib"
With CreateObject("Scripting.FileSystemObject")
    Dim sParentFolderPath : sParentFolderPath = .GetParentFolderName(WScript.ScriptFullName)
    Dim sLibFolderPath : sLibFolderPath = .BuildPath(sParentFolderPath, Cs_FOLDER_LIB)
    Dim oLibFile
    For Each oLibFile In CreateObject("Shell.Application").Namespace(sLibFolderPath).Items
        If Not oLibFile.IsFolder Then
            If StrComp(.GetExtensionName(oLibFile.Path), "vbs", vbTextCompare)=0 Then ExecuteGlobal .OpenTextfile(oLibFile.Path).ReadAll
        End If
    Next
End With
Set oLibFile = Nothing

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
    Set PoWriter = new_WriterTo(fw_getLogPath, 8, True, -1)
    PoWriter.writeBufferSize=100000
    'ブローカークラスのインスタンスの設定
    Dim oBroker : Set oBroker = new_Broker()
    oBroker.subscribe topic.LOG, GetRef("this_logger")
    'パラメータ格納用オブジェクト宣言
    Dim oParams : Set oParams = new_Dic()
    
    '当スクリプトの引数をパラメータ格納用オブジェクトに取得する
    fw_excuteSub "this_getParameters", oParams, oBroker
    
    'ファイル情報の取得
    fw_excuteSub "this_getFileInfomations", oParams, oBroker
    
    '結果出力
    fw_excuteSub "this_makeReport", oParams, oBroker
    
    'ログ出力をクローズ
    PoWriter.close()
    
    'オブジェクトを開放
    Set oParams = Nothing
    Set oBroker = Nothing
    Set PoWriter = Nothing
End Sub

'***************************************************************************************************
'Processing Order            : 1
'Function/Sub Name           : this_getParameters()
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
Private Sub this_getParameters( _
    byRef aoParams _
    )
    'オリジナルの引数を取得
    Dim oArg : Set oArg = fw_storeArguments()
    '★ログ出力
    this_logger Array(logType.DETAIL, "this_getParameters()", cf_toString(oArg))
    
    '実在するパスだけパラメータ格納用オブジェクトに設定
    Dim oParam, oRet, oItem
    Set oParam = new_Arr()
    For Each oItem In oArg.Item("Unnamed")
        Set oRet = fw_tryCatch(Getref("new_FileOf"), oItem, Empty, Empty)
        If oRet.isErr() Then Set oRet = fw_tryCatch(Getref("new_FolderOf"), oItem, Empty, Empty)
        If Not oRet.isErr() Then
            oParam.push oRet.returnValue
        Else
            '★ログ出力
            this_logger Array(logType.WARNING, "this_getParameters()", oItem & "is an invalid argument.")
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
'Function/Sub Name           : this_getFileInfomations()
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
Private Sub this_getFileInfomations( _
    byRef aoParams _
    )
    'パラメータ格納用汎用オブジェクト
    Dim oParam : Set oParam = aoParams.Item("Param").slice(0,vbNullString)

    '★ログ出力
    this_logger Array(logType.INFO, "this_getFileInfomations()", "Before getting list of files.")
    'ファイルオブジェクトのリストを取得
    Dim oList : Set oList = new_Arr()
    Do While oParam.length>0
        oList.pushA fs_getAllFiles(oParam.pop().Path)
    Loop

    '★ログ出力
    this_logger Array(logType.INFO, "this_getFileInfomations()", "Before sorting.")
    '重複を排除してpath順にソートする
    cf_bindAt aoParams, "List", oList.uniq().sortUsing(getref("this_sort"))

    Set oList = Nothing
    Set oParam = Nothing
End Sub

'***************************************************************************************************
'Processing Order            : 3
'Function/Sub Name           : this_makeReport()
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
Private Sub this_makeReport( _
    byRef aoParams _
    )
    If aoParams.Item("List").length=0 Then
    'パラメータ格納用汎用オブジェクトが空の場合
        '★ログ出力
        this_logger Array(logType.WARNING, "this_makeReport()", "There was no files.")
        '何もせず処理を抜ける
        Exit Sub
    End If

    'レポートの作成
    With new_HtmlOf("html")
        .addContent this_makeReportHtmlHeader(aoParams)
        .addContent this_makeReportHtmlBody(aoParams)
    
        '★ログ出力
        this_logger Array(logType.INFO, "this_makeReport()", "Before reportfile output.")
        'レポートをファイルに出力
        Dim sPath
        sPath = fw_getPrivatePath("report", new_Fso().GetBaseName(WScript.ScriptName) & new_Now().formatAs("_YYMMDD_HHmmSS_000") & ".html")
        fs_writeFile sPath, .generate
    End With

    '★ログ出力
    this_logger Array(logType.INFO, "this_makeReport()", "Before open reportfile.")
    'レポートを開く
    fw_runShellSilently fs_wrapInQuotes(sPath)
    
End Sub

'***************************************************************************************************
'Processing Order            : 3-1
'Function/Sub Name           : this_makeReportHtmlHeader()
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
Private Function this_makeReportHtmlHeader( _
    byRef aoParams _
    )
    
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
            .addProperty("border-right", "1px solid #E0E1E3") _
            .addProperty("white-space", "nowrap")
    End With

    Dim oHead : Set oHead = new_HtmlOf("head")
    oHead.addContent oStyle

    Set this_makeReportHtmlHeader = oHead
    Set oStyle = Nothing
    Set oHead = Nothing
End Function

'***************************************************************************************************
'Processing Order            : 3-2
'Function/Sub Name           : this_makeReportHtmlBody()
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
Private Function this_makeReportHtmlBody( _
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

    Set this_makeReportHtmlBody = oBody
    Set oTr = Nothing
    Set oThead = Nothing
    Set oTbody = Nothing
    Set oTable = Nothing
    Set oBody = Nothing
    Set oList = Nothing
End Function

'***************************************************************************************************
'Processing Order            : -
'Function/Sub Name           : this_logger()
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
'2023/11/05         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub this_logger( _
    byRef avParams _
    )
    fw_logger avParams, PoWriter
End Sub
'***************************************************************************************************
'Processing Order            : -
'Function/Sub Name           : this_sort()
'Overview                    : 要素の比較結果を返す
'Detailed Description        : ファイル情報リストのソートで使用する
'Argument
'     aoCurrentValue         : 配列の要素
'     aoNextValue            : 次の配列の要素
'Return Value
'     ソート後の配列
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/12/21         Y.Fujii                  First edition
'***************************************************************************************************
Private Function this_sort( _
    byRef aoCurrentValue _
    , byRef aoNextValue _
    )
    this_sort = aoCurrentValue.ParentFolder&aoCurrentValue.Path > aoNextValue.ParentFolder&aoNextValue.Path
End Function
