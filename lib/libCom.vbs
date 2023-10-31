'***************************************************************************************************
'FILENAME                    : libCom.vbs
'Overview                    : 共通関数ライブラリ
'Detailed Description        : 工事中
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/09/27         Y.Fujii                  First edition
'***************************************************************************************************


'###################################################################################################
'カスタム関数
'###################################################################################################

'***************************************************************************************************
'Function/Sub Name           : cf_bind()
'Overview                    : 変数間の項目移送
'Detailed Description        : 移送する値または変数がオブジェクトか否かによるVBS構文の違い（Setの有無）を吸収する
'                              移送先がコレクションのメンバーの場合は動作しない
'                              移送先が変数の場合に使用できる
'Argument
'     avTo                   : 移送先の変数
'     avValue                : 移送する値または変数
'Return Value
'     なし
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/11/06         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub cf_bind( _
    byRef avTo _
    , byRef avValue _
    )
    If IsObject(avValue) Then Set avTo = avValue Else avTo = avValue
End Sub

'***************************************************************************************************
'Function/Sub Name           : cf_bindAt()
'Overview                    : 変数間の項目移送
'Detailed Description        : 移送する値または変数がオブジェクトか否かによるVBS構文の違い（Setの有無）を吸収する
'                              移送先がコレクションの場合は当関数を使用する
'Argument
'     aoCollection           : 移送先のコレクション
'     asKey                  : 移送先のコレクションのキー
'     avValue                : 移送する値または変数
'Return Value
'     なし
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/11/06         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub cf_bindAt( _
    byRef aoCollection _
    , byVal asKey _
    , byRef avValue _
    )
    If IsObject(avValue) Then Set aoCollection.Item(asKey) = avValue Else aoCollection.Item(asKey) = avValue
End Sub

'***************************************************************************************************
'Function/Sub Name           : cf_push()
'Overview                    : 配列に要素を追加する
'Detailed Description        : 工事中
'Argument
'     avArr                  : 配列
'     aoEle                  : 追加する要素
'Return Value
'     配列の次元数
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/11/23         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub cf_push( _
    byRef avArr _ 
    , byRef aoEle _ 
    )
    If new_Arr().hasElement(avArr) Then
'    If func_CM_ArrayIsAvailable(avArr) Then
        Redim Preserve avArr(Ubound(avArr)+1)
    Else
        Redim avArr(0)
    End If
    cf_bind avArr(Ubound(avArr)), aoEle

'    cf_tryCatch Getref("func_CM_ArrayAddElement"), avArr, Getref("func_CM_ArrayInitialize"), Empty
'    cf_bind avArr(Ubound(avArr)), aoEle

End Sub

'***************************************************************************************************
'Function/Sub Name           : cf_tryCatch()
'Overview                    : 処理の実行とエラー発生時の処理実行
'Detailed Description        : 他の言語のtry-chatch文に準拠
'Argument
'     aoTry                  : 実行する処理（tryブロックの処理）
'     aoArgs                 : 実行する処理の引数
'     aoCatch                : エラー発生時の処理（catchブロックの処理）
'     aoFinary               : エラーの有無に依らず最後に実行する処理（finaryブロックの処理）
'Return Value
'     処理結果
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/10/01         Y.Fujii                  First edition
'***************************************************************************************************
Private Function cf_tryCatch( _
    byRef aoTry _
    , byRef aoArgs _
    , byRef aoCatch _
    , byRef aoFinary _
    )
    Dim oRet, oErr, boFlg
    Set oErr = Nothing : boFlg = True
    
    'tryブロックの処理
    On Error Resume Next
    cf_bind oRet, aoTry(aoArgs)
    If Err.Number<>0 Then
        boFlg = False
        Set oErr = func_CM_UtilStoringErr()
    End If
    On Error GoTo 0

    'catchブロックの処理
    If Not boFlg And func_CM_UtilIsAvailableObject(aoCatch) Then
        cf_bind oRet, aoCatch(aoArgs, oErr)
    End If
    
    'finaryブロックの処理
    If func_CM_UtilIsAvailableObject(aoFinary) Then
        cf_bind oRet, aoFinary(aoArgs, oRet, oErr)
    End If
    
    '結果を返却
    Set cf_tryCatch = new_DicWith(Array("Result", boFlg, "Return", oRet, "Err", oErr))
    Set oRet = Nothing
    Set oErr = Nothing
End Function


'###################################################################################################
'インスタンス生成関数
'###################################################################################################

'***************************************************************************************************
'Function/Sub Name           : new_Dic()
'Overview                    : Dictionaryオブジェクト生成関数
'Detailed Description        : 工事中
'Argument
'     なし
'Return Value
'     生成したDictionaryオブジェクトのインスタンス
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/03         Y.Fujii                  First edition
'***************************************************************************************************
Private Function new_Dic( _
    )
    Set new_Dic = CreateObject("Scripting.Dictionary")
End Function

'***************************************************************************************************
'Function/Sub Name           : new_DicWith()
'Overview                    : Dictionaryオブジェクトを生成し初期値を設定する
'Detailed Description        : 工事中
'Argument
'     avParams               : 初期値奇数（1,3,5,...）はKey、偶数（2,4,6,...）はValue
'                              Keyだけの場合は値にEmptyを設定する。
'Return Value
'     生成したDictionaryオブジェクトのインスタンス
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/03         Y.Fujii                  First edition
'***************************************************************************************************
Private Function new_DicWith( _
    byVal avParams _
    )
    Dim oDict, vItem, vKey, boIsKey
    
    boIsKey = True
    Set oDict = new_Dic()
    
    For Each vItem In avParams
        If boIsKey Then
            cf_bind vKey, vItem
            cf_bindAt oDict, vKey, Empty
        Else
            cf_bindAt oDict, vKey, vItem
        End If
        boIsKey = Not boIsKey
    Next
    
    Set new_DicWith = oDict
    Set oDict = Nothing
End Function

'***************************************************************************************************
'Function/Sub Name           : new_Fso()
'Overview                    : FileSystemObjectオブジェクト生成関数
'Detailed Description        : 工事中
'Argument
'     なし
'Return Value
'     生成したFileSystemObjectオブジェクトのインスタンス
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/10/13         Y.Fujii                  First edition
'***************************************************************************************************
Private Function new_Fso( _
    )
    Set new_Fso = CreateObject("Scripting.FileSystemObject")
End Function

'***************************************************************************************************
'Function/Sub Name           : new_Re()
'Overview                    : 正規表現オブジェクト生成関数
'Detailed Description        : 工事中
'Argument
'     asPattern              : 正規表現のパターン
'     asOptions              : この引数内にある文字の有無で正規表現の以下のプロパティをTrueにする
'                                "i":大文字と小文字を区別する（.IgnoreCase = True）
'                                "g"文字列全体を検索する（.Global = True）
'                                "m"文字列を複数行として扱う（.Multiline = True）
'Return Value
'     生成した正規表現オブジェクトのインスタンス
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/03         Y.Fujii                  First edition
'***************************************************************************************************
Private Function new_Re( _
    byVal asPattern _
    , byVal asOptions _
    )
    Dim oRe, sOpts
    
    Set oRe = New RegExp
    oRe.Pattern = asPattern
    
    sOpts = LCase(asOptions)
    If InStr(sOpts, "i") > 0 Then oRe.IgnoreCase = True
    If InStr(sOpts, "g") > 0 Then oRe.Global = True
    If InStr(sOpts, "m") > 0 Then oRe.Multiline = True
    
    Set new_Re = oRe
    Set oRe = Nothing
End Function

'***************************************************************************************************
'Function/Sub Name           : new_Reader()
'Overview                    : ファイル読込バッファリング処理クラスのインスタンス生成関数
'Detailed Description        : 工事中
'Argument
'     aoTextStream           : テキストストリームオブジェクト
'Return Value
'     生成したファイル読込バッファリング処理クラスのインスタンス
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/10/02         Y.Fujii                  First edition
'***************************************************************************************************
Private Function new_Reader( _
    byRef aoTextStream _
    )
    Set new_Reader = (New clsCmBufferedReader).setTextStream(aoTextStream)
End Function

'***************************************************************************************************
'Function/Sub Name           : new_ReaderFrom()
'Overview                    : ファイル読込バッファリング処理クラスのインスタンス生成関数
'Detailed Description        : 工事中
'Argument
'     asPath                 : 読み込むファイルのパス
'Return Value
'     生成したファイル読込バッファリング処理クラスのインスタンス
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/10/09         Y.Fujii                  First edition
'***************************************************************************************************
Private Function new_ReaderFrom( _
    byVal asPath _
    )
    Set new_ReaderFrom = (New clsCmBufferedReader).setTextStream(func_CM_FsOpenTextFile(asPath, 1, False, -2))
End Function

'***************************************************************************************************
'Function/Sub Name           : new_Writer()
'Overview                    : ファイル出力バッファリング処理クラスのインスタンス生成関数
'Detailed Description        : 工事中
'Argument
'     aoTextStream           : テキストストリームオブジェクト
'Return Value
'     生成したファイル出力バッファリング処理クラスのインスタンス
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/08/27         Y.Fujii                  First edition
'***************************************************************************************************
Private Function new_Writer( _
    byRef aoTextStream _
    )
    Set new_Writer = (New clsCmBufferedWriter).setTextStream(aoTextStream)
End Function

'***************************************************************************************************
'Function/Sub Name           : new_WriterTo()
'Overview                    : ファイル出力バッファリング処理クラスのインスタンス生成関数
'Detailed Description        : 工事中
'Argument
'     asPath                 : 書き込むファイルのパス
'     alIomode               : 出力モード 2:ForWriting,8:ForAppending
'     aboCreate              : asPathが存在しない場合 True:新しいファイルを作成する、False:作成しない
'     alFileFormat           : ファイルの形式 -2:TristateUseDefault,-1:TristateTrue,0:TristateFalse
'Return Value
'     生成したファイル出力バッファリング処理クラスのインスタンス
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/10/09         Y.Fujii                  First edition
'***************************************************************************************************
Private Function new_WriterTo( _
    byVal asPath _
    , byVal alIomode _
    , byVal aboCreate _
    , byVal alFileFormat _
    )
    Set new_WriterTo = (New clsCmBufferedWriter).setTextStream(func_CM_FsOpenTextFile(asPath, alIomode, aboCreate, alFileFormat))
End Function

'***************************************************************************************************
'Function/Sub Name           : new_Now()
'Overview                    : インスタンス生成関数
'Detailed Description        : 今の日付時刻で生成した日付クラスのインスタンスを返す
'Argument
'     なし
'Return Value
'     日付クラスのインスタンス
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/01/04         Y.Fujii                  First edition
'***************************************************************************************************
Private Function new_Now( _
    )
    Set new_Now = (New clsCmCalendar).getNow()
End Function

'***************************************************************************************************
'Function/Sub Name           : new_CalAt()
'Overview                    : インスタンス生成関数
'Detailed Description        : 指定した日付時刻で生成した日付クラスのインスタンスを返す
'Argument
'     avDateTime             : 設定する日付時刻
'Return Value
'     日付クラスのインスタンス
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/03         Y.Fujii                  First edition
'***************************************************************************************************
Private Function new_CalAt( _
    ByVal avDateTime _
    )
    Set new_CalAt = (New clsCmCalendar).setDateTime(avDateTime)
End Function

'***************************************************************************************************
'Function/Sub Name           : new_Broker()
'Overview                    : インスタンス生成関数
'Detailed Description        : 出版-購読型（Publish/Subscribe）クラスのインスタンスを返す
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
Private Function new_Broker( _
    )
    Set new_Broker = (New clsCmBroker)
End Function

'***************************************************************************************************
'Function/Sub Name           : new_Arr()
'Overview                    : インスタンス生成関数
'Detailed Description        : 生成した同クラスのインスタンスを返す
'Argument
'     なし
'Return Value
'     同クラスのインスタンス
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/08         Y.Fujii                  First edition
'***************************************************************************************************
Private Function new_Arr( _
    )
    Set new_Arr = (New clsCmArray)
End Function

'***************************************************************************************************
'Function/Sub Name           : new_ArrWith()
'Overview                    : インスタンス生成関数
'Detailed Description        : 引数で指定した要素を含んだ同クラスのインスタンスを返す
'Argument
'     avArr                  : 配列に追加する要素（配列）
'Return Value
'     同クラスのインスタンス
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/08         Y.Fujii                  First edition
'***************************************************************************************************
Private Function new_ArrWith( _
    byRef avArr _
    )
    Dim oArr : Set oArr = new_Arr()
    oArr.PushMulti avArr
    Set new_ArrWith = oArr
    Set oArr = Nothing
End Function

'***************************************************************************************************
'Function/Sub Name           : new_ArrSplit()
'Overview                    : インスタンス生成関数
'Detailed Description        : vbscriptのSplit関数と同等の機能、同クラスのインスタンスを返す
'Argument
'     asTarget               : 部分文字列と区切り文字を含む文字列表現
'     asDelimiter            : 区切り文字
'Return Value
'     同クラスのインスタンス
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/08         Y.Fujii                  First edition
'***************************************************************************************************
Private Function new_ArrSplit( _
    byVal asTarget _
    , byVal asDelimiter _
    )
    Set new_ArrSplit = new_ArrWith(Split(asTarget, asDelimiter, -1, vbBinaryCompare))
End Function

'***************************************************************************************************
'Function/Sub Name           : new_Char()
'Overview                    : インスタンス生成関数
'Detailed Description        : 文字種類管理クラスのインスタンスを返す
'Argument
'     なし
'Return Value
'     文字種類管理クラスのインスタンス
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/10/31         Y.Fujii                  First edition
'***************************************************************************************************
Private Function new_Char( _
    )
    Set new_Char = (New clsCmCharacterType)
End Function

'***************************************************************************************************
'Function/Sub Name           : new_Func()
'Overview                    : 関数のインスタンスを生成する
'Detailed Description        : javascriptの無名関数に準拠（vbscriptの仕様上仮の名前はつける）
'Argument
'     asSoruceCode           : 生成する関数のソースコード
'                              以下のいずれかの様式とし、function（subではない）を生成する
'                              1.通常
'                               function (①) {②}
'                                ①引数をカンマ区切りで指定する
'                                ②vbscriptの構文に準拠する、戻り値は"return hoge"と表記する
'                                  "return"句がない場合は戻り値はなしとする
'                              2.Arrow関数
'                               ① => ②
'                                ①引数をカンマ区切りで指定する、複数の場合は()で囲む
'                                ②単一行の場合はそのまま戻り値とする、複数行の場合は1.通常の②と同じ
'Return Value
'     生成した関数のインスタンス
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/27         Y.Fujii                  First edition
'***************************************************************************************************
Private Function new_Func( _
    byVal asSoruceCode _
    )
    '生成する関数のソースコードの改行を:に変換
    Dim sSoruceCode
    sSoruceCode = Replace(asSoruceCode, vbCrLf, ":")
    sSoruceCode = Replace(sSoruceCode, vbLf, ":")
    sSoruceCode = Replace(sSoruceCode, vbCr, ":")
    '生成する関数のソースコードの'（シングルクォーテーション）を"（ダブルクォーテーション）に変換
    sSoruceCode = Replace(sSoruceCode, "'", """")
    
    '関数名（仮名）を作る
    Dim sFuncName : sFuncName = "anonymous_" & func_CM_UtilGenerateRandomString(10, 5, Array("_"))
    
    Dim sPattern, oRegExp, sArgStr, sProcStr
    '生成する関数のソースコードの様式が「1.通常」の場合
    sPattern = "function\s*\((.*)\)\s*{(.*)}"
    Set oRegExp = new_Re(sPattern, "igm")
    If oRegExp.Test(sSoruceCode) Then
        sArgStr = oRegExp.Replace(sSoruceCode, "$1")
        sProcStr = oRegExp.Replace(sSoruceCode, "$2")
        
        'return句があれば関数名で書き換える
        sProcStr = func_FuncRewriteReturnPhrase(sFuncName, False, func_FuncAnalyze(sProcStr) )
        
        '関数の生成
        Set new_Func = func_FuncGenerate(sFuncName, sArgStr, sProcStr)
        Set oRegExp = Nothing
        Exit Function
    End If
    
    '生成する関数のソースコードの様式が「2.Arrow関数」の場合
    sPattern = "(.*)\s*=>\s*(.*)\s*"
    Set oRegExp = new_Re(sPattern, "igm")
    If oRegExp.Test(sSoruceCode) Then
        sArgStr = oRegExp.Replace(sSoruceCode, "$1")
        sProcStr = oRegExp.Replace(sSoruceCode, "$2")
        
        'それぞれ前後の括弧があれば除去
        sPattern = "\(\s?(.*)\s?\)"
        sArgStr = new_Re(sPattern, "igm").Replace(sArgStr, "$1")
        sPattern = "{\s?(.*)\s?}"
        sProcStr = new_Re(sPattern, "igm").Replace(sProcStr, "$1")
        
        'return句があれば関数名で書き換える
        sProcStr = func_FuncRewriteReturnPhrase(sFuncName, True, func_FuncAnalyze(sProcStr) )
        
        '関数の生成
        Set new_Func = func_FuncGenerate(sFuncName, sArgStr, sProcStr)
    End If
    Set oRegExp = Nothing
End Function

'***************************************************************************************************
'Function/Sub Name           : func_FuncAnalyze()
'Overview                    : ソースコードを解釈する
'Detailed Description        : new_Func()から使用する
'                              _（アンダーライン）は行を結合する
'Argument
'     asCode                 : ソースコード
'Return Value
'     ソースコードを行ごとに分解した配列
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/10/12         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_FuncAnalyze( _
    byVal asCode _
    )
    Dim sRow, sPtn, oCode, sTemp
    Set oCode = new_Dic()
    sTemp= ""
    For Each sRow In Split(asCode, ":", -1, vbBinaryCompare)
        If Len(Trim(sRow))>0 Then
            sPtn = "^(.*)\s_\s*$"
            If new_Re(sPtn, "ig").Test(sRow) Then
                sTemp = sTemp & Trim(new_Re(sPtn, "ig").Replace(sRow, "$1"))
            Else
                sRow = sTemp & " " & Trim(sRow)
                sTemp = ""
                oCode.Add oCode.Count, Trim(sRow)
            End If
        End If
    Next
    
    func_FuncAnalyze = oCode.Items()
    Set oCode = Nothing
End Function

'***************************************************************************************************
'Function/Sub Name           : func_FuncRewriteReturnPhrase()
'Overview                    : return句を書き換える
'Detailed Description        : new_Func()から使用する
'                              Arrow関数で1行の場合はその行全体をreturnする
'Argument
'     asFuncName             : 関数名
'     aboArrowFlg            : Arrow関数か否かのフラグ
'     avCode                 : ソースコードを行ごとに分解した配列
'Return Value
'     書き換えたソースの処理内容部分のソースコード
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/27         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_FuncRewriteReturnPhrase( _
    byVal asFuncName _
    , byVal aboArrowFlg _
    , byRef avCode _
    )
    Dim sPtnRet : sPtnRet = "^(.*\s+)?return\s+(.*)\s{0,}$"
    
    If Ubound(avCode)=0 And aboArrowFlg=True Then
    'Arrow関数で1行の場合
        Dim sCode : sCode = avCode(0)
        If new_Re(sPtnRet, "ig").Test(sCode) Then
        'return句がある場合
            func_FuncRewriteReturnPhrase = new_Re(sPtnRet, "ig").Replace(sCode, "$1 cf_bind " & asFuncName & ", ($2)")
        Else
        'return句がない場合
            func_FuncRewriteReturnPhrase = "cf_bind " & asFuncName & ", (" & sCode & ")"
        End If
        Exit Function
    End If
    
    Dim lCnt, sPtn, sRow
    For lCnt=0 To Ubound(avCode)
        sRow = avCode(lCnt)
        If new_Re(sPtnRet, "ig").Test(sRow) Then
            avCode(lCnt) = new_Re(sPtnRet, "ig").Replace(sRow, "$1 cf_bind " & asFuncName & ", ($2)")
        End If
    Next
    
    func_FuncRewriteReturnPhrase = Join(avCode, ":")
    
End Function

'***************************************************************************************************
'Function/Sub Name           : func_FuncGenerate()
'Overview                    : 引数の情報で関数のインスタンスを生成する
'Detailed Description        : new_Func()から使用する
'Argument
'     asFuncName             : 関数名
'     asArgStr               : ソースの引数部分のソースコード
'     asProcStr              : ソースの処理内容部分のソースコード
'Return Value
'     生成した関数のインスタンス
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/27         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_FuncGenerate( _
    byVal asFuncName _
    , byVal asArgStr _
    , byVal asProcStr _
    )
    Dim sCode
    'ソースコード作成
    sCode = _
        "Private Function " & asFuncName & "(" & asArgStr & ")" & vbNewLine _
        & asProcStr & vbNewLine _
        & "End Function"
    
'inputbox "","",sCode
    '関数の生成
    ExecuteGlobal sCode
    Set func_FuncGenerate = Getref(asFuncName)
End Function

'###################################################################################################
'数学系の関数
'###################################################################################################

'***************************************************************************************************
'Function/Sub Name           : math_min()
'Overview                    : 最小値を求める
'Detailed Description        : 工事中
'Argument
'     al1                    : 数値1
'     al2                    : 数値2
'Return Value
'     al1とal2の値が小さい方
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/09/27         Y.Fujii                  First edition
'***************************************************************************************************
Private Function math_min( _
    byVal al1 _ 
    , byVal al2 _
    )
    Dim lRet
    If al1 < al2 Then lRet = al1 Else lRet = al2
    math_min = lRet
End Function

'***************************************************************************************************
'Function/Sub Name           : math_max()
'Overview                    : 最大値を求める
'Detailed Description        : 工事中
'Argument
'     al1                    : 数値1
'     al2                    : 数値2
'Return Value
'     al1とal2の値が大きい方
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/09/27         Y.Fujii                  First edition
'***************************************************************************************************
Private Function math_max( _
    byVal al1 _ 
    , byVal al2 _
    )
    Dim lRet
    If al1 > al2 Then lRet = al1 Else lRet = al2
    math_max = lRet
End Function

'***************************************************************************************************
'Function/Sub Name           : math_roundUp()
'Overview                    : 切り上げする
'Detailed Description        : func_MathRound()に委譲する
'Argument
'     adbNum                 : 数値
'     alPlace                : 小数の位、切り上げする端数の位置を小数の位で表す
'Return Value
'     切り上げした値
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/09/24         Y.Fujii                  First edition
'***************************************************************************************************
Private Function math_roundUp( _
    byVal adbNum _ 
    , byVal alPlace _
    )
    math_roundUp = func_MathRound(adbNum, alPlace, 9, True)
End Function

'***************************************************************************************************
'Function/Sub Name           : math_round()
'Overview                    : 四捨五入する
'Detailed Description        : func_MathRound()に委譲する
'Argument
'     adbNum                 : 数値
'     alPlace                : 小数の位、四捨五入する端数の位置を小数の位で表す
'Return Value
'     四捨五入した値
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/09/24         Y.Fujii                  First edition
'***************************************************************************************************
Private Function math_round( _
    byVal adbNum _ 
    , byVal alPlace _
    )
    math_round = func_MathRound(adbNum, alPlace, 5, True)
End Function

'***************************************************************************************************
'Function/Sub Name           : math_roundDown()
'Overview                    : 切り捨てする
'Detailed Description        : func_MathRound()に委譲する
'Argument
'     adbNum                 : 数値
'     alPlace                : 小数の位、切り捨てする端数の位置を小数の位で表す
'Return Value
'     切り捨てした値
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/09/24         Y.Fujii                  First edition
'***************************************************************************************************
Private Function math_roundDown( _
    byVal adbNum _ 
    , byVal alPlace _
    )
    math_roundDown = func_MathRound(adbNum, alPlace, 0, True)
End Function

'***************************************************************************************************
'Function/Sub Name           : math_rand()
'Overview                    : 乱数を生成する
'Detailed Description        : 工事中
'Argument
'     adbMin                 : 生成する乱数の最小値
'     adbMax                 : 生成する乱数の最大値
'     alPlace                : 小数の位、切り上げする端数の位置を小数の位で表す
'Return Value
'     生成した乱数
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/24         Y.Fujii                  First edition
'***************************************************************************************************
Private Function math_rand( _
    byVal adbMin _
    , byVal adbMax _
    , byVal alPlace _
    )
    Randomize
    math_rand = adbMin + Fix( ((adbMax-adbMin)*(10^alPlace)+1)*Rnd )*10^(-1*alPlace)
End Function

'***************************************************************************************************
'Function/Sub Name           : func_MathRound()
'Overview                    : 数値を丸める
'Detailed Description        : 符号を無視して絶対値を丸める
'                              引数のalPlaceは丸めたい小数の位を指定する、例えば第一位を場合は0を指定する
'                              小数第二位を四捨五入する場合、alPlaceに1、alThresholdに5を指定する
'                              一の位、十の位、・・・の場合は-1,-2,…のように負値を指定する
'                                例）１８２．７３２
'                                　　↑　　↑　↑　 ↑
'                                   -3  -1　0　 2
'Argument
'     adbNum                 : 数値
'     alPlace                : 小数の位、処理する端数の位置を小数の位で表す
'     alThreshold            : 閾値
'                               0：切り捨て
'                               5：四捨五入
'                               9：切り上げ
'     aboMode                : 端数処理の方法
'                               True  ：符号を無視して絶対値を丸める（正負で丸める方向が異なる）
'                               False ：正数の場合と増減を同じ向きに丸める
'                              https://ja.wikipedia.org/wiki/%E7%AB%AF%E6%95%B0%E5%87%A6%E7%90%86
'Return Value
'     丸めた値
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/09/24         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_MathRound( _
    byVal adbNum _ 
    , byVal alPlace _
    , byVal alThreshold _
    , byVal aboMode _
    )
    Dim lThreshold : lThreshold = alThreshold
    If adbNum<0 Then lThreshold = -1*lThreshold

    Dim dbTemp
    dbTemp = Cstr((adbNum+lThreshold*10^(-1*(alPlace+1))) * 10^(alPlace))

    If aboMode Then
        func_MathRound = Cdbl( Cstr( Fix(dbTemp) * 10^(-1*alPlace) ) )
    Else
        func_MathRound = Cdbl( Cstr( Int(dbTemp) * 10^(-1*alPlace) ) )
    End If
End Function


'###################################################################################################
'エクセル系
'###################################################################################################

'***************************************************************************************************
'Function/Sub Name           : sub_CM_ExcelSaveAs()
'Overview                    : エクセルファイルを別名で保存して閉じる
'Detailed Description        : 工事中
'Argument
'     aoWorkBook             : エクセルのワークブック
'     asPath                 : 保存するファイルのフルパス
'     alFileformat           : XlFileFormat 列挙体（デフォルトはxlOpenXMLWorkbook 51 Excelブック）
'Return Value
'     なし
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/09/27         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub sub_CM_ExcelSaveAs( _
    byRef aoWorkBook _
    , byVal asPath _
    , byVal alFileformat _
    )
    If Not(IsNumeric(alFileformat)) Then
        alFileformat = 51                  'xlOpenXMLWorkbook 51 Excelブック
    End If
    Call aoWorkBook.SaveAs( _
                            asPath _
                            , alFileformat _
                            , , _
                            , False _
                            , False _
                            )
    Call aoWorkBook.Close(False)
End Sub

'***************************************************************************************************
'Function/Sub Name           : func_CM_ExcelOpenFile()
'Overview                    : エクセルファイルを読み取り専用／ダイアログなしで開く
'Detailed Description        : 工事中
'Argument
'     aoExcel                : エクセル
'     asPath                 : エクセルファイルのフルパス
'Return Value
'     開いたエクセルのワークブック
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/09/27         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_ExcelOpenFile( _
    byRef aoExcel _
    , byVal asPath _
    )    
    Set func_CM_ExcelOpenFile = aoExcel.Workbooks.Open( _
                                                        asPath _
                                                        , 0 _
                                                        , True _
                                                        , , , _
                                                        , True _
                                                        )
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_ExcelGetTextFromAutoshape()
'Overview                    : エクセルのオートシェイプのテキストを取り出す
'Detailed Description        : エラーは無視する
'Argument
'     aoAutoshape            : オートシェイプ
'Return Value
'     オートシェイプのテキスト
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/09/27         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_ExcelGetTextFromAutoshape( _
    byRef aoAutoshape _
    )
    func_CM_ExcelGetTextFromAutoshape = aoAutoshape.TextFrame.Characters.Text
End Function


'###################################################################################################
'ファイル操作系
'###################################################################################################

'***************************************************************************************************
'Function/Sub Name           : func_CM_FsDeleteFile()
'Overview                    : ファイルを削除する
'Detailed Description        : FileSystemObjectのDeleteFile()と同等
'Argument
'     asPath                 : 削除するファイルのフルパス
'Return Value
'     結果 True:成功 / False:失敗
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/09/27         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FsDeleteFile( _
    byVal asPath _
    ) 
    If Not func_CM_FsFileExists(asPath) Then func_CM_FsDeleteFile = False
    func_CM_FsDeleteFile = cf_tryCatch(new_Func("a=>a(0).DeleteFile(a(1))"), Array(new_Fso(), asPath), Empty, Empty).Item("Result")
    
'    On Error Resume Next
'    new_Fso().DeleteFile(asPath)
'    func_CM_FsDeleteFile = True
'    If Err.Number Then
'        Err.Clear
'        func_CM_FsDeleteFile = False
'    End If
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_FsDeleteFolder()
'Overview                    : フォルダを削除する
'Detailed Description        : FileSystemObjectのDeleteFolder()と同等
'Argument
'     asPath                 : 削除するフォルダのフルパス
'Return Value
'     結果 True:成功 / False:失敗
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/11/12         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FsDeleteFolder( _
    byVal asPath _
    ) 
    If Not func_CM_FsFolderExists(asPath) Then func_CM_FsDeleteFolder = False
    On Error Resume Next
    new_Fso().DeleteFolder(asPath)
    func_CM_FsDeleteFolder = True
    If Err.Number Then
        Err.Clear
        func_CM_FsDeleteFolder = False
    End If
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_FsDeleteFsObject()
'Overview                    : ファイルかフォルダを削除する
'Detailed Description        : func_CM_FsDeleteFile()とfunc_CM_FsDeleteFolder()に委譲
'Argument
'     asPath                 : パス
'Return Value
'     結果 True:成功 / False:失敗
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/08/19         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FsDeleteFsObject( _
    byVal asPath _
    )
    func_CM_FsDeleteFsObject = False
    If func_CM_FsFileExists(asPath) Then func_CM_FsDeleteFsObject = func_CM_FsDeleteFile(asPath)
    If func_CM_FsFolderExists(asPath) Then func_CM_FsDeleteFsObject = func_CM_FsDeleteFolder(asPath)
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_FsCopyFile ()
'Overview                    : ファイルをコピーする
'Detailed Description        : FileSystemObjectのCopyFile()と同等
'Argument
'     asPathFrom             : コピー元ファイルのフルパス
'     asPathTo               : コピー先ファイルのフルパス
'Return Value
'     結果 True:成功 / False:失敗
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/10/24         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FsCopyFile( _
    byVal asPathFrom _
    , byVal asPathTo _
    ) 
    If Not func_CM_FsFileExists(asPathFrom) Then func_CM_FsCopyFile = False
    On Error Resume Next
    Call new_Fso().CopyFile(asPathFrom, asPathTo)
    func_CM_FsCopyFile = True
    If Err.Number Then
        Err.Clear
        func_CM_FsCopyFile = False
    End If
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_FsCopyFolder ()
'Overview                    : フォルダをコピーする
'Detailed Description        : FileSystemObjectのCopyFolder()と同等
'Argument
'     asPathFrom             : コピー元フォルダのフルパス
'     asPathTo               : コピー先フォルダのフルパス
'Return Value
'     結果 True:成功 / False:失敗
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/08/19         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FsCopyFolder( _
    byVal asPathFrom _
    , byVal asPathTo _
    ) 
    If Not func_CM_FsFolderExists(asPathFrom) Then func_CM_FsCopyFolder = False
    On Error Resume Next
    Call new_Fso().CopyFolder(asPathFrom, asPathTo)
    func_CM_FsCopyFolder = True
    If Err.Number Then
        Err.Clear
        func_CM_FsCopyFolder = False
    End If
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_FsCopyFsObject()
'Overview                    : ファイルかフォルダをコピーする
'Detailed Description        : func_CM_FsCopyFile()とfunc_CM_FsCopyFolder()に委譲
'Argument
'     asPathFrom             : コピー元ファイル/フォルダのフルパス
'     asPathTo               : コピー先のフルパス
'Return Value
'     結果 True:成功 / False:失敗
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/08/19         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FsCopyFsObject( _
    byVal asPathFrom _
    , byVal asPathTo _
    )
    func_CM_FsCopyFsObject = False
    If func_CM_FsFileExists(asPathFrom) Then func_CM_FsCopyFsObject = func_CM_FsCopyFile(asPathFrom, asPathTo)
    If func_CM_FsFolderExists(asPathFrom) Then func_CM_FsCopyFsObject = func_CM_FsCopyFolder(asPathFrom, asPathTo)
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_FsMoveFile ()
'Overview                    : ファイルを移動する
'Detailed Description        : FileSystemObjectのMoveFile()と同等
'Argument
'     asPathFrom             : 移動元ファイルのフルパス
'     asPathTo               : 移動先ファイルのフルパス
'Return Value
'     結果 True:成功 / False:失敗
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/08/19         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FsMoveFile( _
    byVal asPathFrom _
    , byVal asPathTo _
    ) 
    If Not func_CM_FsFileExists(asPathFrom) Then func_CM_FsMoveFile = False
    On Error Resume Next
    Call new_Fso().MoveFile(asPathFrom, asPathTo)
    func_CM_FsMoveFile = True
    If Err.Number Then
        Err.Clear
        func_CM_FsMoveFile = False
    End If
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_FsMoveFolder ()
'Overview                    : フォルダを移動する
'Detailed Description        : FileSystemObjectのMoveFolder()と同等
'Argument
'     asPathFrom             : 移動元フォルダのフルパス
'     asPathTo               : 移動先フォルダのフルパス
'Return Value
'     結果 True:成功 / False:失敗
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/08/19         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FsMoveFolder( _
    byVal asPathFrom _
    , byVal asPathTo _
    ) 
    If Not func_CM_FsFolderExists(asPathFrom) Then func_CM_FsMoveFolder = False
    On Error Resume Next
    Call new_Fso().MoveFolder(asPathFrom, asPathTo)
    func_CM_FsMoveFolder = True
    If Err.Number Then
        Err.Clear
        func_CM_FsMoveFolder = False
    End If
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_FsMoveFsObject()
'Overview                    : ファイルかフォルダを移動する
'Detailed Description        : func_CM_FsMoveFile()とfunc_CM_FsMoveFolder()に委譲
'Argument
'     asPathFrom             : 移動元ファイル/フォルダのフルパス
'     asPathTo               : 移動先のフルパス
'Return Value
'     結果 True:成功 / False:失敗
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/08/19         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FsMoveFsObject( _
    byVal asPathFrom _
    , byVal asPathTo _
    )
    func_CM_FsMoveFsObject = False
    If func_CM_FsFileExists(asPathFrom) Then func_CM_FsMoveFsObject = func_CM_FsMoveFile(asPathFrom, asPathTo)
    If func_CM_FsFolderExists(asPathFrom) Then func_CM_FsMoveFsObject = func_CM_FsMoveFolder(asPathFrom, asPathTo)
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_FsGetParentFolderPath()
'Overview                    : 親フォルダパスの取得
'Detailed Description        : FileSystemObjectのGetParentFolderName()と同等
'Argument
'     asPath                 : ファイルのパス
'Return Value
'     親フォルダパス
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/09/27         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FsGetParentFolderPath( _
    byVal asPath _
    ) 
    func_CM_FsGetParentFolderPath = new_Fso().GetParentFolderName(asPath)
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_FsGetGetBaseName()
'Overview                    : ファイル名（拡張子を除く）の取得
'Detailed Description        : FileSystemObjectのGetBaseName()と同等
'Argument
'     asPath                 : ファイルのパス
'Return Value
'     ファイル名（拡張子を除く）
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/10/23         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FsGetGetBaseName( _
    byVal asPath _
    ) 
    func_CM_FsGetGetBaseName = new_Fso().GetBaseName(asPath)
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_FsGetGetExtensionName()
'Overview                    : ファイルの拡張子の取得
'Detailed Description        : FileSystemObjectのGetExtensionName()と同等
'Argument
'     asPath                 : ファイルのパス
'Return Value
'     ファイルの拡張子
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/10/23         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FsGetGetExtensionName( _
    byVal asPath _
    ) 
    func_CM_FsGetGetExtensionName = new_Fso().GetExtensionName(asPath)
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_FsBuildPath()
'Overview                    : ファイルパスの連結
'Detailed Description        : FileSystemObjectのBuildPath()と同等
'Argument
'     asFolderPath           : パス
'     asItemName             : asFolderPathに連結するフォルダ名またはファイル名
'Return Value
'     連結したファイルパス
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/09/27         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FsBuildPath( _
    byVal asFolderPath _
    , byVal asItemName _
    ) 
    func_CM_FsBuildPath = new_Fso().BuildPath(asFolderPath, asItemName)
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_FsFileExists()
'Overview                    : ファイルの存在確認
'Detailed Description        : FileSystemObjectのFileExists()と同等
'Argument
'     asPath                 : パス
'Return Value
'     結果 True:存在する / False:存在しない
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/09/27         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FsFileExists( _
    byVal asPath _
    ) 
    func_CM_FsFileExists = new_Fso().FileExists(asPath)
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_FsFolderExists()
'Overview                    : フォルダの存在確認
'Detailed Description        : FileSystemObjectのFolderExists()と同等
'Argument
'     asPath                 : パス
'Return Value
'     結果 True:存在する / False:存在しない
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/10/16         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FsFolderExists( _
    byVal asPath _
    ) 
    func_CM_FsFolderExists = new_Fso().FolderExists(asPath)
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_FsGetFile()
'Overview                    : ファイルオブジェクトの取得
'Detailed Description        : FileSystemObjectのGetFile()と同等
'Argument
'     asPath                 : パス
'Return Value
'     Fileオブジェクト
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/09/27         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FsGetFile( _
    byVal asPath _
    ) 
    Set func_CM_FsGetFile = new_Fso().GetFile(asPath)
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_FsGetFolder()
'Overview                    : フォルダオブジェクトの取得
'Detailed Description        : FileSystemObjectのGetFolder()と同等
'Argument
'     asPath                 : パス
'Return Value
'     Folderオブジェクト
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/10/23         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FsGetFolder( _
    byVal asPath _
    ) 
    Set func_CM_FsGetFolder = new_Fso().GetFolder(asPath)
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_FsGetFsObject()
'Overview                    : ファイルかフォルダオブジェクトの取得
'Detailed Description        : func_CM_FsGetFile()とfunc_CM_FsGetFolder()に委譲
'Argument
'     asPath                 : パス
'Return Value
'     File/Folderオブジェクト
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/08/19         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FsGetFsObject( _
    byVal asPath _
    )
    Set func_CM_FsGetFsObject = Nothing
    If func_CM_FsFileExists(asPath) Then Set func_CM_FsGetFsObject = func_CM_FsGetFile(asPath)
    If func_CM_FsFolderExists(asPath) Then Set func_CM_FsGetFsObject = func_CM_FsGetFolder(asPath)
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_FsGetFiles()
'Overview                    : 指定したフォルダ以下のFilesコレクションを取得する
'Detailed Description        : FileSystemObjectのFolderオブジェクトのFilesコレクションと同等
'Argument
'     asPath                 : パス
'Return Value
'     Filesコレクション
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/10/23         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FsGetFiles( _
    byVal asPath _
    ) 
    Set func_CM_FsGetFiles = new_Fso().GetFolder(asPath).Files
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_FsGetFolders()
'Overview                    : 指定したフォルダ以下のFoldersコレクションを取得する
'Detailed Description        : FileSystemObjectのFolderオブジェクトのSubFoldersコレクションと同等
'Argument
'     asPath                 : パス
'Return Value
'     Foldersコレクション
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/10/23         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FsGetFolders( _
    byVal asPath _
    ) 
    Set func_CM_FsGetFolders = new_Fso().GetFolder(asPath).SubFolders
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_FsGetFsObjects()
'Overview                    : 指定したフォルダ以下のFilesコレクションとFoldersコレクションを取得する
'Detailed Description        : func_CM_FsGetFiles()とfunc_CM_FsGetFolders()に委譲
'Argument
'     asPath                 : パス
'Return Value
'     FilesコレクションとFoldersコレクション
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/08/19         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FsGetFsObjects( _
    byVal asPath _
    )
    Set func_CM_FsGetFsObjects = Nothing
    If Not func_CM_FsFolderExists(asPath) Then Exit Function
    Dim oTemp : Set oTemp = new_Dic()
    With oTemp
        .Add "Filse", func_CM_FsGetFiles(asPath)
        .Add "Folders", func_CM_FsGetFolders(asPath)
    End With
    Set func_CM_FsGetFsObjects = oTemp
    Set oTemp = Nothing
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_FsGetTempFileName()
'Overview                    : ランダムに生成された一時ファイルまたはフォルダーの名前の取得
'Detailed Description        : FileSystemObjectのGetTempName()と同等
'Argument
'     asPath                 : パス
'Return Value
'     一時ファイルまたはフォルダーの名前
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/09/27         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FsGetTempFileName()
    func_CM_FsGetTempFileName = new_Fso().GetTempName()
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_FsGetPrivateFilePath()
'Overview                    : 実行中のスクリプトがあるフォルダからのパスを返す
'Detailed Description        : 上位フォルダが存在しない場合は作成する
'Argument
'     asParentFolderName     : 親フォルダ名
'     asFileName             : ファイル名
'Return Value
'     ファイルのフルパス
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/26         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FsGetPrivateFilePath( _
    byVal asParentFolderName _
    , byVal asFileName _
    )
    Dim sParentFolderPath : sParentFolderPath = func_CM_FsGetParentFolderPath(WScript.ScriptFullName)
    If Len(asParentFolderName)>0 Then
    '引数で指定したディレクトリ名がある場合
        sParentFolderPath = func_CM_FsBuildPath(sParentFolderPath ,asParentFolderName)
    End If
    func_CM_FsGetPrivateFilePath = func_CM_FsGetFilePathWithCreateParentFolder(sParentFolderPath, asFileName)
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_FsGetTempFilePath()
'Overview                    : 一時ファイルのパスを返す
'Detailed Description        : 実行中のスクリプトがあるフォルダのtmpフォルダ以下に作成する
'                              上位フォルダが存在しない場合は作成する
'Argument
'     なし
'Return Value
'     ファイルのフルパス
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/26         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FsGetTempFilePath( _
    )
    func_CM_FsGetTempFilePath = func_CM_FsGetPrivateFilePath("tmp", func_CM_FsGetTempFileName())
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_FsGetPrivateLogFilePath()
'Overview                    : 実行中のスクリプトのログファイルパスを返す
'Detailed Description        : 実行中のスクリプトがあるフォルダのlogフォルダ以下に
'                              スクリプトファイル名＋".log"形式のファイル名で作成する
'                              上位フォルダが存在しない場合は作成する
'Argument
'     なし
'Return Value
'     ファイルのフルパス
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/26         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FsGetPrivateLogFilePath( _
    )
    func_CM_FsGetPrivateLogFilePath = func_CM_FsGetPrivateFilePath("log", func_CM_FsGetGetBaseName(WScript.ScriptName) & ".log" )
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_FsGetFilePathWithCreateParentFolder()
'Overview                    : ファイルのパスを取得
'Detailed Description        : 上位フォルダが存在しない場合は作成する
'Argument
'     asParentFolderPath     : 親フォルダのパス
'     asFileName             : ファイル名
'Return Value
'     ファイルのフルパス
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/26         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FsGetFilePathWithCreateParentFolder( _
    byVal asParentFolderPath _
    , byVal asFileName _
    )
    If Not(func_CM_FsFolderExists(asParentFolderPath)) Then func_CM_FsCreateFolder(asParentFolderPath)
    func_CM_FsGetFilePathWithCreateParentFolder = func_CM_FsBuildPath(asParentFolderPath, asFileName)
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_FsCreateFolder()
'Overview                    : フォルダを作成する
'Detailed Description        : FileSystemObjectのCreateFolder()と同等
'Argument
'     asPath                 : パス
'Return Value
'     作成したフォルダのフルパス
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/09/27         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FsCreateFolder( _
    byVal asPath _
    )
    func_CM_FsCreateFolder = new_Fso().CreateFolder(asPath)
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_FsOpenTextFile()
'Overview                    : ファイルを開きTextStreamオブジェクトを返す
'Detailed Description        : FileSystemObjectのOpenTextFile()と同等
'Argument
'     asPath                 : パス
'     alIomode               : 入力/出力モード 1:ForReading,2:ForWriting,8:ForAppending
'     aboCreate              : asPathが存在しない場合 True:新しいファイルを作成する、False:作成しない
'     alFileFormat           : ファイルの形式 -2:TristateUseDefault,-1:TristateTrue,0:TristateFalse
'Return Value
'     TextStreamオブジェクト
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/01/09         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FsOpenTextFile( _
    byVal asPath _
    , byVal alIomode _
    , byVal aboCreate _
    , byVal alFileFormat _
    )
    'ファイルを開く
    Set func_CM_FsOpenTextFile = new_Fso().OpenTextFile( _
                                                              asPath _
                                                              , alIomode _
                                                              , aboCreate _
                                                              , alFileFormat _
                                                              )
End Function

'***************************************************************************************************
'Function/Sub Name           : sub_CM_FsWriteFile()
'Overview                    : ファイル出力する
'Detailed Description        : エラーは無視する
'Argument
'     asPath                 : 出力先のフルパス
'     asCont                 : 出力する内容
'     なし
'Return Value
'     作成したフォルダのフルパス
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/10/16         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub sub_CM_FsWriteFile( _
    byVal asPath _
    , byVal asCont _
    )
    On Error Resume Next
    'ファイルを開く（存在しない場合は作成する）
    With func_CM_FsOpenTextFile(asPath, 2, True, -2)
        Call .WriteLine(asCont)
        Call .Close
    End With
    If Err.Number Then
        Err.Clear
    End If
End Sub

'***************************************************************************************************
'Function/Sub Name           : func_CM_FsIsSame()
'Overview                    : 指定したパスが同じファイル/フォルダか検査する
'Detailed Description        : 工事中
'Argument
'     asPathA                : ファイル/フォルダのフルパス
'     asPathB                : ファイル/フォルダのフルパス
'Return Value
'     結果 True:同一 / False:同一でない
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/08/19         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FsIsSame( _
    byVal asPathA _
    , byVal asPathB _
    )
    func_CM_FsIsSame = (func_CM_FsGetFsObject(asPathA) Is func_CM_FsGetFsObject(asPathB))
End Function


'###################################################################################################
'文字列操作系
'###################################################################################################

'***************************************************************************************************
'Function/Sub Name           : func_CM_StrConvOnlyAlphabet()
'Overview                    : 英字だけ大文字／小文字に変換する
'Detailed Description        : 工事中
'Argument
'     asTarget               : 変換する文字列
'     alConversion           : 実行する変換の種類 1:UpperCase,2:LowerCase
'Return Value
'     変換した文字列
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/08/19         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_StrConvOnlyAlphabet( _
    byVal asTarget _
    , byVal alConversion _
    )
    Dim sChar, asTargetTmp
    
    '1文字ずつ判定する
    Dim asTargetNew : asTargetNew = asTarget
    Dim lPos : lPos = 1
    Do While Len(asTargetNew) >= lPos
        '変換対象の1文字を取得
        sChar = Mid(asTargetNew, lPos, 1)
        
        If new_Char().whatType(sChar)<3 Then
        '変換対象が英字の場合のみ変換する
            asTargetTmp = ""
            
            '変換対象の文字までの文字列
            If lPos > 1 Then
                asTargetTmp = Mid(asTargetNew, 1, lPos-1)
            End If
            
            '変換した1文字を結合
            sChar = func_CM_StrConv(sChar, alConversion)
            asTargetTmp = asTargetTmp & sChar
            
            '変換対象の文字移行の文字列を結合
            If lPos < Len(asTargetNew) Then
                asTargetTmp = asTargetTmp & Mid(asTargetNew, lPos+1, Len(asTargetNew)-lPos)
            End If
            
            '変換後の文字列を格納
            asTargetNew = asTargetTmp
        End If
        
        'カウントアップ
        lPos = lPos+1
    Loop
    
    func_CM_StrConvOnlyAlphabet = asTargetNew
    
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_StrConv()
'Overview                    : 文字列を指定のとおり変換する
'Detailed Description        : 工事中
'Argument
'     asTarget               : 変換する文字列
'     alConversion           : 実行する変換の種類（現時点で1,2のみ実装）
'                                 1:文字列を大文字に変換
'                                 2:文字列を小文字に変換
'                                 3:文字列内のすべての単語の最初の文字を大文字に変換
'                                 4:文字列内の狭い (1 バイト) 文字をワイド (2 バイト) 文字に変換
'                                 8:文字列内のワイド (2 バイト) 文字を狭い (1 バイト) 文字に変換
'                                16:文字列内のひらがな文字をカタカナ文字に変換
'                                32:文字列内のカタカナ文字をひらがな文字に変換
'Return Value
'     変換した文字列
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/08/19         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_StrConv( _
    byVal asTarget _
    , byVal alalConversion _
    )
    Dim sChar, asTargetTmp
    func_CM_StrConv = asTarget
    Select Case alalConversion
        Case 1:
            func_CM_StrConv = UCase(asTarget)
        Case 2:
            func_CM_StrConv = LCase(asTarget)
    End Select
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_StrLen()
'Overview                    : 全角は2文字、半角は1文字として文字数を返す
'Detailed Description        : 工事中
'Argument
'     asTarget               : 文字列
'Return Value
'     文字数
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/08/19         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_StrLen( _
    byVal asTarget _
    )
    '1文字ずつ判定する
    Dim sChar
    Dim lLength : lLength = 0
    Dim lPos : lPos = 1
    Do While Len(asTarget) >= lPos
        '1文字を取得
        sChar = Mid(asTarget, lPos, 1)
        
        If (Asc(sChar) And &HFF00) <> 0 Then
            lLength = lLength+2
        Else
            lLength = lLength+1
        End If
        
        'カウントアップ
        lPos = lPos+1
    Loop
    
    func_CM_StrLen = lLength
End Function

'###################################################################################################
'配列系
'###################################################################################################

'***************************************************************************************************
'Function/Sub Name           : func_CM_ArrayGetDimensionNumber()
'Overview                    : 配列の次元数を求める
'Detailed Description        : 工事中
'Argument
'     avArray                : 配列
'Return Value
'     配列の次元数
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/11/19         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_ArrayGetDimensionNumber( _
    byRef avArray _ 
    )
   If Not IsArray(avArray) Then Exit Function
   On Error Resume Next
   Dim lNum : lNum = 0
   Dim lTemp
   Do
       lNum = lNum + 1
       lTemp = UBound(avArray, lNum)
   Loop Until Err.Number <> 0
   Err.Clear
   func_CM_ArrayGetDimensionNumber = lNum - 1
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_ArrayAddElement()
'Overview                    : 配列の要素を追加する
'Detailed Description        : 工事中
'Argument
'     avArr                  : 配列
'Return Value
'     なし
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/10/15         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_ArrayAddElement( _
    byRef avArr _
    )
    Redim Preserve avArr(Ubound(avArr)+1)
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_ArrayInitialize()
'Overview                    : 配列を初期化する
'Detailed Description        : 工事中
'Argument
'     avArr                  : 配列
'     avErr                  : エラー情報
'Return Value
'     なし
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/10/15         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_ArrayInitialize( _
    byRef avArr _
    , byRef avErr _
    )
    Redim avArr(0)
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_ArrayUbound()
'Overview                    : 配列のインデックスの最大数を返す
'Detailed Description        : 工事中
'Argument
'     avArr                  : 配列
'Return Value
'     なし
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/10/15         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_ArrayUbound( _
    byRef avArr _
    )
    func_CM_ArrayUbound = Ubound(avArr)
End Function

'###################################################################################################
'チェック系
'###################################################################################################

'***************************************************************************************************
'Function/Sub Name           : func_CM_ValidationlIsWithinTheRangeOf()
'Overview                    : 数値型の範囲内にあるか検査する
'Detailed Description        : 工事中
'Argument
'     avNumber               : 数値
'     alType                 : 変数の型
'                                1:整数型（Integer）
'                                2:長整数型（Long）
'                                3:バイト型（Byte）
'                                4:単精度浮動小数点型（Single）
'                                5:倍精度浮動小数点型（Double）
'                                6:通貨型（Currency）
'Return Value
'     整形した浮動小数点型
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/08/26         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_ValidationlIsWithinTheRangeOf( _
    byVal avNumber _
    , byVal alType _
    )
    Dim vMin,vMax
    Select Case alType
        Case 1:                   '整数型（Integer）
            vMin = -1 * 2^15
            vMax = 2^15 - 1
        Case 2:                   '長整数型（Long）
            vMin = -1 * 2^31
            vMax = 2^31 - 1
        Case 3:                   'バイト型（Byte）
            vMin = 0
            vMax = 2^8 - 1
        Case 4:                   '単精度浮動小数点型（Single）
            vMin = -3.402823E38
            vMax = 3.402823E38
        Case 5:                   '倍精度浮動小数点型（Double）
            vMin = -1.79769313486231E308
            vMax = 1.79769313486231E308
        Case 6:                   '通貨型（Currency）
            vMin = -1 * 2^59 / 1000
            vMax = ( 2^59 - 1 ) / 1000
    End Select
    
    func_CM_ValidationlIsWithinTheRangeOf = False
    If vMin<=avNumber And avNumber<=vMax Then
        func_CM_ValidationlIsWithinTheRangeOf = True
    End If
End Function


'###################################################################################################
'その他
'###################################################################################################

'***************************************************************************************************
'Function/Sub Name           : func_CM_GetObjectByNameFromCollection()
'Overview                    : コレクションから指定したnameのメンバーを取得する
'Detailed Description        : エラー処理は行わない
'Argument
'     aoArr                  : 0番目　コレクション、1番目　name
'Return Value
'     該当するメンバー
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/09/27         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_GetObjectByNameFromCollection( _
    byRef aoArr _
    )
    cf_bind func_CM_GetObjectByNameFromCollection, aoArr(0).Item(aoArr(1))
End Function

'***************************************************************************************************
'Function/Sub Name           : sub_CM_Swap()
'Overview                    : 変数の値を入れ替える
'Detailed Description        : 移送処理はcf_bind()を使用する
'Argument
'     avA                    : 値を入れ替える変数
'     avB                    : 値を入れ替える変数
'Return Value
'     なし
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/21         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub sub_CM_Swap( _
    byRef avA _
    , byRef avB _
    )
    Dim oTemp
    Call cf_bind(oTemp, avA)
    Call cf_bind(avA, avB)
    Call cf_bind(avB, oTemp)
    Set oTemp = Nothing
End Sub

'***************************************************************************************************
'Function/Sub Name           : func_CM_FillInTheCharacters()
'Overview                    : 文字を埋める
'Detailed Description        : 対象文字の不足桁を指定したアライメントで指定した文字の1文字目で埋める
'                              対象文字に不足桁がない場合は、指定した文字数で切り取る
'Argument
'     asTarget               : 対象文字列
'     alWordCount            : 文字数
'     asToFillCharacter      : 埋める文字
'     aboIsCutOut            : 文字数で切り取り（True：する/False：しない）
'     aboIsRightAlignment    : アライメント（True：右寄せ/False：左寄せ）
'Return Value
'     埋めた文字列
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/01/04         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FillInTheCharacters( _
    byVal asTarget _
    , byVal alWordCount _
    , byVal asToFillCharacter _
    , byVal aboIsCutOut _
    , byVal aboIsRightAlignment _
    )
    
    '切り取りなしで対象文字列が文字数より大きい場合は処理を抜ける
    Dim lTargetLen : lTargetLen = Len(asTarget)
    If Not(aboIsCutOut) And lTargetLen>=alWordCount Then
        func_CM_FillInTheCharacters = asTarget
        Exit Function
    End If
    
    '埋める文字列の作成
    Dim sFillStrings : sFillStrings = ""
    If alWordCount-lTargetLen > 0 Then
        sFillStrings = String(alWordCount-lTargetLen , asToFillCharacter)
    End If
    
    Dim sResult
    'アライメント指定によって文字列を埋める
    If aboIsRightAlignment Then
        sResult = sFillStrings & asTarget
    Else
        sResult = asTarget & sFillStrings
    End If
    
    '切り取りなしの場合は処理を抜ける
    If Not(aboIsCutOut) Then
        func_CM_FillInTheCharacters = sResult
        Exit Function
    End If
    
    'アライメント指定によって文字列を切り取る
    If aboIsRightAlignment Then
        sResult = Right(sResult, alWordCount)
    Else
        sResult = Left(sResult, alWordCount)
    End If
    func_CM_FillInTheCharacters = sResult
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_FormatDecimalNumber()
'Overview                    : 浮動小数点型を整形する
'Detailed Description        : 工事中
'Argument
'     adbNumber              : 浮動小数点型の数値
'     alDecimalPlaces        : 小数の桁数
'Return Value
'     整形した浮動小数点型
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/08/26         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FormatDecimalNumber( _
    byVal adbNumber _
    , byVal alDecimalPlaces _
    )
    func_CM_FormatDecimalNumber = Fix(adbNumber) & "." _
                             & func_CM_FillInTheCharacters( _
                                                          Abs(Fix( (adbNumber - Fix(adbNumber))*10^alDecimalPlaces )) _
                                                          , alDecimalPlaces _
                                                          , "0" _
                                                          , False _
                                                          , True _
                                                          )
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_ToString()
'Overview                    : 引数の数値・文字列やオブジェクトの中身を可読な表示に変換する
'Detailed Description        : 配列やディクショナリのようなオブジェクトだったら中身を表示し、
'                              そうでない場合はVarTypeでオブジェクトのクラスを表示する
'Argument
'     avTarget               : 対象
'Return Value
'     変換した文字列
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/03         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_ToString( _
    byRef avTarget _
    )
    Dim oEscapingDoubleQuote, sRet
    Set oEscapingDoubleQuote = new_Re("""", "g")
    sRet = ""
    
    Err.Clear
    On Error Resume Next
    
    If VarType(avTarget) = vbString Then
        sRet = """" & oEscapingDoubleQuote.Replace(avTarget, """""") & """"
    ElseIf IsArray(avTarget) Then
        sRet = func_CM_ToStringArray(avTarget)
    ElseIf IsObject(avTarget) Then
        sRet = func_CM_ToStringObject(avTarget)
    ElseIf IsEmpty(avTarget) Then
        sRet = "<empty>"
    ElseIf IsNull(avTarget) Then
        sRet = "<null>"
    Else
        sRet = func_CM_ToStringOther(avTarget)
    End If
    
    If Err.Number <> 0 Then
        Err.Clear
        sRet = func_CM_ToStringUnknown(avTarget)
    End If
    
    func_CM_ToString = sRet
    
    Set oEscapingDoubleQuote = Nothing
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_ToStringArray()
'Overview                    : 配列の中身を可読な表示に変換する
'Detailed Description        : 工事中
'Argument
'     avTarget               : 対象
'Return Value
'     変換した文字列
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/03         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_ToStringArray( _
    byRef avTarget _
    )
    Dim oTemp(), vItem
    
    For Each vItem In avTarget
        Call cf_push(oTemp, func_CM_ToString(vItem))
    Next
    func_CM_ToStringArray = "[" & Join(oTemp, ",") & "]"
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_ToStringDictionary()
'Overview                    : ディクショナリの中身を可読な表示に変換する
'Detailed Description        : 工事中
'Argument
'     avTarget               : 対象
'Return Value
'     変換した文字列
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/03         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_ToStringDictionary( _
    byRef avTarget _
    )
    Dim oTemp(), vKey
    
    For Each vKey In avTarget.Keys
        Call cf_push(oTemp, func_CM_ToString(vKey) & "=>" & func_CM_ToString(avTarget.Item(vKey)))
    Next
    func_CM_ToStringDictionary = "{" & Join(oTemp, ",") & "}"
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_ToStringObject()
'Overview                    : オブジェクトの中身を可読な表示に変換する
'Detailed Description        : 工事中
'Argument
'     avTarget               : 対象
'Return Value
'     変換した文字列
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/03         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_ToStringObject( _
    byRef avTarget _
    )
    Dim sRet
    
    Err.Clear
    On Error Resume Next
    
    sRet = func_CM_ToStringDictionary(avTarget)
    
    If Err.Number <> 0 Then
        Err.Clear
        sRet = func_CM_ToStringArray(avTarget)
    End If
    
    If Err.Number <> 0 Then
        Err.Clear
        sRet = func_CM_ToStringArray(avTarget.Items)
    End If
    
    If Err.Number <> 0 Then
        Err.Clear
        sRet = "<" & TypeName(avTarget) & ">"
    End If
    
    func_CM_ToStringObject = sRet
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_ToStringOther()
'Overview                    : その他オブジェクトの中身を可読な表示に変換する
'Detailed Description        : 工事中
'Argument
'     avTarget               : 対象
'Return Value
'     変換した文字列
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/03         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_ToStringOther( _
    byRef avTarget _
    )
    Dim sRet
    
    Err.Clear
    On Error Resume Next
    
    sRet = CStr(avTarget)
    
    If Err.Number <> 0 Then
        Err.Clear
        sRet = func_CM_ToStringArray(avTarget)
    End If
    
    If Err.Number <> 0 Then
        Err.Clear
        sRet = func_CM_ToStringDictionary(avTarget)
    End If
    
    If Err.Number <> 0 Then
        Err.Clear
        sRet = func_CM_ToStringUnknown(avTarget)
    End If
    
    func_CM_ToStringOther = sRet
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_ToStringUnknown()
'Overview                    : 引数の型が不明な場合に可読な表示に変換する
'Detailed Description        : 工事中
'Argument
'     avTarget               : 対象
'Return Value
'     変換した文字列
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/03         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_ToStringUnknown( _
    byRef avTarget _
    )
    func_CM_ToStringUnknown = "<unknown:" & VarType(avTarget) & " " & TypeName(avTarget) & ">"
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_ToStringErr()
'Overview                    : Errオブジェクトの内容を可読な表示に変換する
'Detailed Description        : 工事中
'Argument
'     なし
'Return Value
'     変換した文字列
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/25         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_ToStringErr( _
    )
    func_CM_ToStringErr = "<Err> " & func_CM_ToString(func_CM_UtilStoringErr())
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_ToStringArguments()
'Overview                    : Argumentsオブジェクトの内容を可読な表示に変換する
'Detailed Description        : 工事中
'Argument
'     なし
'Return Value
'     変換した文字列
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/26         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_ToStringArguments( _
    )
    func_CM_ToStringArguments = "<Arguments> " & func_CM_ToString(func_CM_UtilStoringArguments())
End Function

'***************************************************************************************************
'Function/Sub Name           : sub_CM_ExcuteSub()
'Overview                    : 関数を実行する
'Detailed Description        : 工事中
'Argument
'     asSubName              : 実行する関数名
'     aoArgument             : 実行する関数に渡す引数
'     aoBroker               : 出版-購読型（Publish/subscribe）クラスのオブジェクト
'Return Value
'     なし
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/03         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub sub_CM_ExcuteSub( _
    byVal asSubName _
    , byRef aoArgument _
    , byRef aoBroker _
    )
    Const Cs_TOPIC = "log"
    
    '出版（Publish） 開始
    If Not aoBroker Is Nothing Then
        aoBroker.Publish Cs_TOPIC, Array(5 ,asSubName ,"Start")
        aoBroker.Publish Cs_TOPIC, Array(9 ,asSubName ,func_CM_ToString(aoArgument))
    End If
    
    '関数の実行
    Dim oFunc, oRet
    Set oFunc = GetRef(asSubName)
    If aoArgument Is Nothing Then
        Set oRet = cf_tryCatch( new_Func("function(a){a()}"), oFunc, Empty, Empty )
    Else
        Set oRet = cf_tryCatch( oFunc, aoArgument, Empty, Empty )
    End If
    
    '出版（Publish） 終了
    If Not aoBroker Is Nothing Then
        If oRet.Item("Result")=False Then
        'エラー
            aoBroker.Publish Cs_TOPIC, Array(1, asSubName, func_CM_ToString(oRet.Item("Err")))
        Else
        '正常
            aoBroker.Publish Cs_TOPIC, Array(5, asSubName, "End")
        End If
        aoBroker.Publish Cs_TOPIC, Array(9, asSubName, func_CM_ToString(aoArgument))
    End If
    
    Set oRet = Nothing
    Set oFunc = Nothing
End Sub

'###################################################################################################
'ユーティリティ系
'###################################################################################################

'***************************************************************************************************
'Function/Sub Name           : func_CM_UtilSortBubble()
'Overview                    : バブルソート
'Detailed Description        : 計算回数はO(N^2)
'                              配列（avArr）が無効な配列の場合は配列（avArr）をそのまま返す
'                              引数の関数の引数は以下のとおり
'                                currentValue :配列の要素
'                                nextValue    :次の配列の要素
'Argument
'     avArr                  : 配列
'     aoFunc                 : 関数
'     aboFlg                 : 判定方法
'                                True  :昇順（関数の結果がTrueの場合に入れ替える）
'                                False :降順（関数の結果がFalseの場合に入れ替える）
'Return Value
'     ソート後の配列
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/18         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_UtilSortBubble( _
    byRef avArr _
    , byRef aoFunc _
    , byVal aboFlg _
    )
    func_CM_UtilSortBubble = avArr
    If Not new_Arr().hasElement(avArr) Then Exit Function
'    If Not func_CM_ArrayIsAvailable(avArr) Then Exit Function
    If Ubound(avArr)=0 Then Exit Function
    
    Dim lEnd, lPos
    lEnd = Ubound(avArr)
    Do While lEnd>0
        For lPos=0 To lEnd-1
            If aoFunc(avArr(lPos), avArr(lPos+1))=aboFlg Then
            'lPos番目の要素と(lPos+1)番目の要素を入れ替える
                Call sub_CM_Swap(avArr(lPos), avArr(lPos+1))
            End If
        Next
        lEnd = lEnd-1
    Loop
    func_CM_UtilSortBubble = avArr
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_UtilSortQuick()
'Overview                    : クイックソート
'Detailed Description        : 計算回数は平均O(N*logN)、最悪はO(N^2)
'                              配列（avArr）が無効な配列の場合は配列（avArr）をそのまま返す
'                              引数の関数の引数は以下のとおり
'                                currentValue :配列の要素
'                                nextValue    :次の配列の要素
'Argument
'     avArr                  : 配列
'     aoFunc                 : 関数
'     aboFlg                 : 判定方法
'                                True  :昇順（関数の結果がTrueの場合に入れ替える）
'                                False :降順（関数の結果がFalseの場合に入れ替える）
'Return Value
'     ソート後の配列
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/18         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_UtilSortQuick( _
    byRef avArr _
    , byRef aoFunc _
    , byVal aboFlg _
    )
    func_CM_UtilSortQuick = avArr
    If Not new_Arr().hasElement(avArr) Then Exit Function
'    If Not func_CM_ArrayIsAvailable(avArr) Then Exit Function
    If Ubound(avArr)=0 Then Exit Function
    
    '0番目の要素をピボットに決める
    Dim oPivot : Call cf_bind(oPivot, avArr(0))
    
    'ピボットと要素を関数で判定し判定方法に合致するグループをRight、そうでないグループをLeftとする
    Dim lPos, vRight, vLeft
    For lPos=1 To Ubound(avArr)
        If aoFunc(avArr(lPos), oPivot)=aboFlg Then
            Call cf_push(vRight, avArr(lPos))
        Else
            Call cf_push(vLeft, avArr(lPos))
        End If
    Next
    
    '上述で分けたRight、Leftのグループごとに再帰処理する
    vLeft = func_CM_UtilSortQuick(vLeft, aoFunc, aboFlg)
    vRight = func_CM_UtilSortQuick(vRight, aoFunc, aboFlg)
    
    'Leftにピボット＋Rightを結合する
    Call cf_push(vLeft, oPivot)
    If new_Arr().hasElement(vRight) Then
'    If func_CM_ArrayIsAvailable(vRight) Then
        For lPos=0 To Ubound(vRight)
            Call cf_push(vLeft, vRight(lPos))
        Next
    End If
    
    func_CM_UtilSortQuick = vLeft
    Set oPivot = Nothing
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_UtilSortMerge()
'Overview                    : マージソート
'Detailed Description        : 計算回数はO(N*logN)
'                              配列（avArr）が無効な配列の場合は配列（avArr）をそのまま返す
'                              マージ処理はfunc_CM_UtilSortMergeMerge()に委譲する
'Argument
'     avArr                  : 配列
'     aoFunc                 : 関数
'     aboFlg                 : 判定方法
'                                True  :昇順（関数の結果がTrueの場合に入れ替える）
'                                False :降順（関数の結果がFalseの場合に入れ替える）
'Return Value
'     ソート後の配列
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/18         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_UtilSortMerge( _
    byRef avArr _
    , byRef aoFunc _
    , byVal aboFlg _
    )
    func_CM_UtilSortMerge = avArr
    If Not new_Arr().hasElement(avArr) Then Exit Function
'    If Not func_CM_ArrayIsAvailable(avArr) Then Exit Function
    If Ubound(avArr)=0 Then Exit Function
    
    '2つの配列に分解する
    Dim lLength, lMedian
    lLength = Ubound(avArr) - Lbound(avArr) + 1
    lMedian = math_roundUp(lLength/2, 0)
'    lMedian = math_roundUp(lLength/2, 1)
    Dim lPos, vFirst, vSecond
    For lPos=Lbound(avArr) To lMedian-1
        Call cf_push(vFirst, avArr(lPos))
    Next
    For lPos=lMedian To Ubound(avArr)
        Call cf_push(vSecond, avArr(lPos))
    Next
    
    '再帰処理で配列の要素が1つになるまで分解する
    vFirst = func_CM_UtilSortMerge(vFirst, aoFunc, aboFlg)
    vSecond = func_CM_UtilSortMerge(vSecond, aoFunc, aboFlg)
    
    'マージをしながら上位に戻す
    func_CM_UtilSortMerge = func_CM_UtilSortMergeMerge(vFirst, vSecond, aoFunc, aboFlg)
    
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_UtilSortMergeMerge()
'Overview                    : マージソートのマージ処理
'Detailed Description        : func_CM_UtilSortMerge()から呼び出す
'                              引数の関数の引数は以下のとおり
'                                currentValue :配列の要素
'                                nextValue    :次の配列の要素
'Argument
'     avFirst                : マージするソート済みの配列
'     avSecond               : マージするソート済みの配列
'     aoFunc                 : 関数
'     aboFlg                 : 判定方法
'                                True  :昇順（関数の結果がTrueの場合に入れ替える）
'                                False :降順（関数の結果がFalseの場合に入れ替える）
'Return Value
'     マージ済の配列
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/18         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_UtilSortMergeMerge( _
    byRef avFirst _
    , byRef avSecond _
    , byRef aoFunc _
    , byVal aboFlg _
    )
    Dim lPosF, lPosS, lEndF, lEndS
    lPosF = Lbound(avFirst) : lPosS = Lbound(avSecond)
    lEndF = Ubound(avFirst) : lEndS = Ubound(avSecond)
    
    '双方の配列の先頭の要素同士を関数で判定して戻り値の配列に追加する
    Dim vRet
    Do While lPosF<=lEndF And lPosS<=lEndS
        If aoFunc(avFirst(lPosF), avSecond(lPosS))=aboFlg Then
            Call cf_push(vRet, avSecond(lPosS))
            lPosS = lPosS + 1
        Else
            Call cf_push(vRet, avFirst(lPosF))
            lPosF = lPosF + 1
        End If
    Loop
    
    'それぞれ残っている方の配列の要素を追加する
    Dim lPos
    If lPosF<=lEndF Then
        For lPos=lPosF To lEndF
            Call cf_push(vRet, avFirst(lPos))
        Next
    End If
    If lPosS<=lEndS Then
        For lPos=lPosS To lEndS
            Call cf_push(vRet, avSecond(lPos))
        Next
    End If
    
    'マージ済の配列を返す
    func_CM_UtilSortMergeMerge = vRet
    
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_UtilSortHeap()
'Overview                    : ヒープソート
'Detailed Description        : 計算回数はO(N*logN)
'                              配列（avArr）が無効な配列の場合は配列（avArr）をそのまま返す
'Argument
'     avArr                  : 配列
'     aoFunc                 : 関数
'     aboFlg                 : 判定方法
'                                True  :昇順（関数の結果がTrueの場合に入れ替える）
'                                False :降順（関数の結果がFalseの場合に入れ替える）
'Return Value
'     ソート後の配列
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/21         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_UtilSortHeap( _
    byRef avArr _
    , byRef aoFunc _
    , byVal aboFlg _
    )
    func_CM_UtilSortHeap = avArr
    If Not new_Arr().hasElement(avArr) Then Exit Function
'    If Not func_CM_ArrayIsAvailable(avArr) Then Exit Function
    If Ubound(avArr)=0 Then Exit Function
    
    'ヒープの作成
    Dim lLb, lUb, lSize, lParent
    lLb = Lbound(avArr) : lUb = Ubound(avArr)
    lSize = lUb - lLb + 1
    '子を持つ最下部のノードから上位に向けて順番にノード単位の処理を行う
    For lParent=lSize\2-1 To lLb Step -1
        Call sub_CM_UtilSortHeapPerNodeProc(avArr, lSize, lParent, aoFunc, aboFlg)
    Next
    
    'ヒープの先頭（最大/最小値）を順番に取り出す
    Do While lSize>0
        'ヒープの先頭と末尾を入れ替える
        Call sub_CM_Swap(avArr(lLb), avArr(lSize-1))
        'ヒープサイズを１つ減らして再作成
        lSize = lSize - 1
        Call sub_CM_UtilSortHeapPerNodeProc(avArr, lSize, 0, aoFunc, aboFlg)
    Loop
    
    'ソート済の配列を返す
    func_CM_UtilSortHeap = avArr
    
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_UtilSortHeapPerNodeProc()
'Overview                    : ヒープソートのノード単位の処理
'Detailed Description        : func_CM_UtilSortHeap()から呼び出す
'                              引数の関数の引数は以下のとおり
'                                currentValue :配列の要素
'                                nextValue    :次の配列の要素
'Argument
'     avArr                  : 配列
'     alSize                 : ヒープのサイズ
'     alParent               : ノードの親の配列番号
'     aoFunc                 : 関数
'     aboFlg                 : 判定方法
'                                True  :昇順（関数の結果がTrueの場合に入れ替える）
'                                False :降順（関数の結果がFalseの場合に入れ替える）
'Return Value
'     ソート後の配列
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/21         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub sub_CM_UtilSortHeapPerNodeProc( _
    byRef avArr _
    , byVal alSize _
    , byVal alParent _
    , byRef aoFunc _
    , byVal aboFlg _
    )
    Dim lRight, lLeft, lToSwap
    lLeft = alParent*2 + 1
    lRight = lLeft + 1
    lToSwap = alParent
    
    If lRight<alSize Then
    '右側の子がある場合
        If aoFunc(avArr(lRight), avArr(alParent))=aboFlg Then
        '親と右側の子の要素を関数で判定し判定方法に合致する場合は入れ替える
            lToSwap = lRight
        End If
    End If
    
    If lLeft<alSize Then
    '左側の子がある場合
        If aoFunc(avArr(lLeft), avArr(lToSwap))=aboFlg Then
        '親と右側の子の勝者と左側の子の要素を関数で判定し判定方法に合致する場合は入れ替える
            lToSwap = lLeft
        End If
    End If
    
    If lToSwap<>alParent Then
        '親と子の要素を入れ替える
        Call sub_CM_Swap(avArr(alParent), avArr(lToSwap))
        '入れ替えた子の要素以下のノードを再処理する
        Call sub_CM_UtilSortHeapPerNodeProc(avArr, alSize, lToSwap, aoFunc, aboFlg)
    End If
    
End Sub

'***************************************************************************************************
'Function/Sub Name           : func_CM_UtilSortDefaultFunc()
'Overview                    : 要素の比較結果を返す
'Detailed Description        : ソート関数群で使うデフォルトの関数
'Argument
'     aoCurrentValue         : 配列の要素
'     aoNextValue            : 次の配列の要素
'Return Value
'     ソート後の配列
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/10/15         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_UtilSortDefaultFunc( _
    byRef aoCurrentValue _
    , byRef aoNextValue _
    )
    func_CM_UtilSortDefaultFunc = aoCurrentValue>aoNextValue
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_UtilGenerateRandomString()
'Overview                    : ランダムな文字列を生成する
'Detailed Description        : 指定した長さ、文字の種類でランダムな文字列を生成する
'Argument
'     alLength               : 文字の長さ
'     alType                 : 文字の種類（複数指定する場合は以下の和を設定する）
'                                    1:半角英字大文字
'                                    2:半角英字小文字
'                                    4:半角数字
'                                    8:半角記号
'                                   16:半角カタカナ
'                                   32:半角カタカナ記号
'                                   64:全角英字大文字
'                                  128:全角英字小文字
'                                  256:全角数字
'                                  512:全角記号
'                                 1024:全角ひらがな
'                                 2048:全角カタカナ
'                                 4096:全角ギリシャ、キリル文字の大文字
'                                 8192:全角ギリシャ、キリル文字の小文字
'                                16384:全角線枠
'                                32768:全角漢字 第1水準(16区～47区)
'                                65536:全角漢字 第2水準(48区～84区)
'     avAdditional           : 配列で指定する文字種、前述の文字の種類と重複する場合は追加しない
'                              指定がない場合はNothingなど配列以外を指定する
'Return Value
'     生成した文字列
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/24         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_UtilGenerateRandomString( _
    byVal alLength _
    , byVal alType _
    , byRef avAdditional _
    )
    
    '文字の種類（alType）で指定した文字のリストを作成する
    Dim vSettings : vSettings = Array( _
          Array( Array("A", "Z") ) _
          , Array( Array("a", "z") ) _
          , Array( Array("0", "9") ) _
          , Array( Array("!", "/"), Array(":", "@"), Array("[", "`"), Array("{", "~") ) _
          , Array( Array("ｦ", "ｯ"), Array("ｱ", "ﾟ") ) _
          , Array( Array("｡", "･"), Array("ｰ", "ｰ") ) _
          , Array( Array("Ａ", "Ｚ") ) _
          , Array( Array("ａ", "ｚ") ) _
          , Array( Array("０", "９") ) _
          , Array( Array("、", "〓"), Array("∈", "∩"), Array("∧", "∃"), Array("∠", "∬"), Array("Å", "¶"), Array("◯", "◯") ) _
          , Array( Array("ぁ", "ん") ) _
          , Array( Array("ァ", "ヶ") ) _
          , Array( Array("Α", "Ω"), Array("А", "Я") ) _
          , Array( Array("α", "ω"), Array("а", "я") ) _
          , Array( Array("─", "╂") ) _
          , Array( Array("亜", "腕") ) _
          , Array( Array("弌", "滌"), Array("漾", "熙") ) _
          )
    
    Dim lType : lType = alType
    Dim lPowerOf2 : lPowerOf2 = 16          '2^16 = 65536 <= alTypeの最大値
    Dim oChars : Set oChars = new_Dic()
    Dim lQuotient,lDivide, vSetting, vItem, bCode
'    Dim lQuotient,lDivide, vSetting, vItem, bCode, sCodeHex
    Do Until lPowerOf2<0
        lDivide = 2^lPowerOf2
        lQuotient = lType \ lDivide
        lType = lType Mod lDivide
        
        If lQuotient>0 Then
            vSetting = vSettings(lPowerOf2)
            For Each vItem In vSetting
                For bCode = Asc(vItem(0)) To Asc(vItem(1))
                    oChars.Add bCode, Chr(bCode)
'                    sCodeHex = Right(Hex(bCode),2)
'                    If bCode>0 Or (sCodeHex<>"7F" And ("3F"<sCodeHex And sCodeHex<"FD")) Then
'                        oChars.Add bCode, Chr(bCode)
'                    End If
                Next
            Next
        End If
        
        lPowerOf2 = lPowerOf2 - 1
    Loop
    
    'sjis使用範囲外のコードを除外する
    Dim sCodeHex
    For Each bCode In oChars.Keys()
        If bCode<0 Then
            sCodeHex = Right(Hex(bCode),2)
            If sCodeHex="7F" Or sCodeHex<="3F" Or "FD"<=sCodeHex Then
                oChars.Remove bCode
            End If
        End If
    Next
    
    '配列で指定する文字種（avAdditional）を追加する
    If Not IsObject(avAdditional) Then
        If IsArray(avAdditional) And (Not IsEmpty(avAdditional)) Then
            Dim sChar
            For Each sChar In avAdditional
                If Not oChars.Exists(Asc(sChar)) Then
                    oChars.Add Asc(sChar), sChar
                End If
            Next
        End If
    End If

    '上述で作成した文字のリストを使ってランダムな文字列を生成する
    Dim lPos, sRet
    sRet = ""
    For lPos = 1 To alLength
        sRet = sRet & oChars.Items()( math_rand(0, oChars.Count - 1, 0) )
    Next
    func_CM_UtilGenerateRandomString = sRet
    
    Set oChars = Nothing
End Function

'***************************************************************************************************
'Function/Sub Name           : sub_CM_UtilLogger()
'Overview                    : ログ出力する
'Detailed Description        : 引数の情報にタイムスタンプを付加してファイル出力する
'Argument
'     avParams               : 配列型のパラメータリスト
'     aoWriter               : ファイル出力バッファリング処理クラスのインスタンス
'Return Value
'     なし
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/26         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub sub_CM_UtilLogger( _
    byRef avParams _
    , byRef aoWriter _
    )
    Dim oCont, sIp
    sIp = new_ArrWith(func_CM_UtilGetIpaddress()).filter(new_Func("(e,i,a)=>left(e.item(""Ip"").item(""V4""),3)<>""172"""))(0).Item("Ip").Item("V4")
    Set oCont = new_ArrWith(Array(new_Now(), sIp, func_CM_UtilGetComputerName()))
    
    With aoWriter
        .Write(oCont.Concat(avParams).join(vbTab))
        .newLine()
    End With
    Set oCont = Nothing
End Sub

'***************************************************************************************************
'Function/Sub Name           : func_CM_UtilStoringErr()
'Overview                    : Errオブジェクトの内容をオブジェクトに変換する
'Detailed Description        : 変換したオブジェクトの構成
'                              Key             Value                     例
'                              --------------  ------------------------  ---------------------------
'                              "Number"        Err.Numberの内容          11
'                              "Description"   Err.Descriptionのの内容   0 で除算しました。
'                              "Source"        Err.Sourceの内容          Microsoft VBScript 実行時エラー
'Argument
'     なし
'Return Value
'     変換したオブジェクト
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/10/01         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_UtilStoringErr( _
    )
    Dim oRet : Set oRet = new_Dic()
    oRet.Add "Number", Err.Number
    oRet.Add "Description", Err.Description
    oRet.Add "Source", Err.Source
    Set func_CM_UtilStoringErr = oRet
    Set oRet = Nothing
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_UtilStoringArguments()
'Overview                    : Argumentsオブジェクトの内容をオブジェクトに変換する
'Detailed Description        : 変換したオブジェクトの構成
'                              例は引数が a /X /Hoge:Fuga, b の場合
'                              Key         Value                                        例
'                              ----------  -------------------------------------------  -------------
'                              "All"       WScript.Arguments以下のItemの内容            a /X /Hoge:Fuga, b
'                              "Named"     WScript.Arguments.Named以下のItemの内容      X: Hoge:Fuga
'                              "Unnamed"   WScript.Arguments.Unnamed以下のItemの内容    a b
'Argument
'     なし
'Return Value
'     変換したオブジェクト
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/26         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_UtilStoringArguments( _
    )
    Dim oRet : Set oRet = new_Dic()
    Dim oTemp, oEle, oKey
    
    'All
    Set oTemp = new_Arr()
    For Each oEle In WScript.Arguments
        oTemp.Push oEle
    Next
    oRet.Add "All", oTemp
    
    'Named
    Set oTemp = new_Dic()
    For Each oKey In WScript.Arguments.Named
        oTemp.Add oKey, WScript.Arguments.Named.Item(oKey)
    Next
    oRet.Add "Named", oTemp
    
    'Unnamed
    Set oTemp = new_Arr()
    For Each oEle In WScript.Arguments.Unnamed
        oTemp.Push oEle
    Next
    oRet.Add "Unnamed", oTemp
    
    Set func_CM_UtilStoringArguments = oRet
    
    Set oKey = Nothing
    Set oEle = Nothing
    Set oTemp = Nothing
    Set oRet = Nothing
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_UtilGetIpaddress()
'Overview                    : 自身のIPアドレスを取得する
'Detailed Description        : IPアドレスを格納したオブジェクトを返す
'Argument
'     なし
'Return Value
'     IPアドレスを格納したオブジェクトの配列
'                              内容は以下のとおり
'                               Key             Value                   例
'                               --------------  ----------------------  ----------------------------
'                               "Caption"       Adapter名               -
'                               "Ip"            以下オブジェクト        -
'                              
'                              IP Addressを格納したオブジェクト
'                               Key             Value                   例
'                               --------------  ----------------------  ----------------------------
'                               "V4"            IP Address(v4)          192.168.11.52
'                               "V6"            IP Address(v6)          fe80::ba87:1e93:59ab:28f7%18
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/10/10         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_UtilGetIpaddress( _
    )
    Dim sMyComp, oAdapter, oAddress, oRet, oIpv4, oIpv6
    
    sMyComp = "."
    Set oRet = new_Arr()
    For Each oAdapter in GetObject("winmgmts:\\"&sMyComp&"\root\cimv2").ExecQuery("Select * From Win32_NetworkAdapterConfiguration Where IPEnabled = True")
         For Each oAddress in oAdapter.IPAddress
             If new_ArrSplit(oAddress, ".").length=4 Then
             'IPv4
                 cf_bind oIpv4, oAddress
             Else
             'IPv6
                 cf_bind oIpv6, oAddress
             End If
         Next
         oRet.push new_DicWith(Array("Caption", oAdapter.Caption, "Ip", new_DicWith(Array("V4", oIpv4, "V6", oIpv6))))
    Next
    func_CM_UtilGetIpaddress = oRet.items
    
    Set oAddress = Nothing
    Set oAdapter = Nothing
    Set oRet = Nothing
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_UtilGetComputerName()
'Overview                    : 自身のコンピュータ名を取得する
'Detailed Description        : 工事中
'Argument
'     なし
'Return Value
'     自身のコンピュータ名
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/10/10         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_UtilGetComputerName( _
    )
    func_CM_UtilGetComputerName = CreateObject("WScript.Network").ComputerName
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_UtilIsSame()
'Overview                    : 同一か判定する
'Detailed Description        : 工事中
'Argument
'     aoA                    : 比較対象
'     aoB                    : 比較対象
'Return Value
'     結果 True:同一 / False:同一でない
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/10/15         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_UtilIsSame( _
    byRef aoA _
    , byRef aoB _
    )
    Dim boFlg : boFlg = False
    If IsObject(aoA) And IsObject(aoB) Then
        If aoA Is aoB Then boFlg = True
    ElseIf Not IsObject(aoA) And Not IsObject(aoB) Then
        If VarType(aoA) = vbString And VarType(aoB) = vbString Then
            If Strcomp(aoA, aoB, vbBinaryCompare)=0 Then boFlg = True
        Else
            If aoA = aoB Then boFlg = True
        End If
    End If
    func_CM_UtilIsSame = boFlg
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_UtilIsAvailableObject()
'Overview                    : オブジェクトが利用可能か判定する
'Detailed Description        : 工事中
'Argument
'     aoObj                  : オブジェクト
'Return Value
'     結果 True:利用可能 / False:利用不可
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/10/15         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_UtilIsAvailableObject( _
    byRef aoObj _
    )
    Dim boFlg : boFlg = False
    If IsObject(aoObj) Then
        If Not aoObj Is Nothing Then boFlg = True
    End If
    func_CM_UtilIsAvailableObject = boFlg
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_UtilIsTextStream()
'Overview                    : オブジェクトがTextStreamか判定する
'Detailed Description        : 工事中
'Argument
'     aoObj                  : オブジェクト
'Return Value
'     結果 True:TextStreamである / False:TextStreamでない
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/10/15         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_UtilIsTextStream( _
    byRef aoObj _
    )
    Dim boFlg : boFlg = False
    If Vartype(aoObj)=9 And Strcomp(Typename(aoObj),"TextStream",vbBinaryCompare)=0 Then boFlg = True
    func_CM_UtilIsTextStream = boFlg
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_UtilJoin()
'Overview                    : Join関数
'Detailed Description        : vbscriptのJoin関数と同等の機能
'Argument
'     avArr                  : 配列
'     asDel                  : 区切り文字
'Return Value
'     配列の各要素を連結した文字列
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/10/17         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_UtilJoin( _
    byRef avArr _
    , byVal asDel _
)
    func_CM_UtilJoin = Join(avArr, asDel)
End Function
