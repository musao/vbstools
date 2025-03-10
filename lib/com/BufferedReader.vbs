'***************************************************************************************************
'FILENAME                    : BufferedReader.vbs
'Overview                    : ファイル読込バッファリング処理クラス
'Detailed Description        : 工事中
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/10/02         Y.Fujii                  First edition
'***************************************************************************************************
Class BufferedReader
    'クラス内変数、定数
    Private PoTextStream, PoOutbound, PoInbound, PoBuffer, PlReadSize
    
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
    '2023/10/02         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub Class_Initialize()
        PlReadSize = 5000                 'デフォルトは5000バイト
        Set PoTextStream = Nothing
        Dim vArr : vArr = Array("line", Empty, "column", Empty, "atEndOfLine", Empty, "atEndOfStream", Empty)
        Set PoOutbound = new_DicOf(vArr)
        Set PoInbound = new_DicOf(vArr)
        Set PoBuffer = new_DicOf(Array("buffer", Empty, "pointer", Empty, "length", Empty))
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : Class_Terminate()
    'Overview                    : デストラクタ
    'Detailed Description        : バッファの未出力分を出力してから終了処理を行う
    'Argument
    '     なし
    'Return Value
    '     なし
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/02         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub Class_Terminate()
        this_close
        Set PoOutbound = Nothing
        Set PoInbound = Nothing
        Set PoBuffer = Nothing
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Let readSize()
    'Overview                    : 読込サイズを設定する
    'Detailed Description        : 読込要求時に読込バッファのサイズがこれを超えた場合
    '                              ファイルを読込む
    'Argument
    '     alReadSize             : 読込サイズ
    'Return Value
    '     なし
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/02         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Let readSize( _
        byVal alReadSize _
        )
        ast_argTrue cf_isPositiveInteger(alReadSize), TypeName(Me)&"+readSize() Let", "Not a positive integer."
        PlReadSize = CLng(alReadSize)
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get readSize()
    'Overview                    : 読込サイズを返す
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     読込サイズ
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/02         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get readSize()
        readSize = PlReadSize
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get textStream()
    'Overview                    : テキストストリームを返す
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     テキストストリームオブジェクト
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/02         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get textStream()
        Set textStream = PoTextStream
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get line()
    'Overview                    : 現在の行番号を返す
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     現在の行番号
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/02         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get line()
        line = PoOutbound.Item("line")
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get column()
    'Overview                    : 現在の列番号を返す
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     現在の列番号
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/02         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get column()
        column = PoOutbound.Item("column")
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get atEndOfStream()
    'Overview                    : 行末の場合にTrueを返す
    'Detailed Description        : 現行末マーカーの直前にファイル ポインターがある場合は true を返し、
    '                              そうでない場合は false を返す。
    'Argument
    '     なし
    'Return Value
    '     結果 True:行末 / False:行末以外
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/02         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get atEndOfStream()
        atEndOfStream = PoOutbound.Item("atEndOfStream")
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get atEndOfLine()
    'Overview                    : ファイルの終端の場合にTrueを返す
    'Detailed Description        : 最後にファイル ポインターがある場合は true を返し、
    '                              そうでない場合は false を返す。
    'Argument
    '     なし
    'Return Value
    '     結果 True:ファイルの終端 / False:ファイルの終端以外
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/02         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get atEndOfLine()
        atEndOfLine = PoOutbound.Item("atEndOfLine")
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : setTextStream()
    'Overview                    : テキストストリームを設定する
    'Detailed Description        : 工事中
    'Argument
    '     aoTextStream           : テキストストリームオブジェクト
    'Return Value
    '     自身のインスタンス
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/02         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function setTextStream( _
        byRef aoTextStream _
        )
        ast_argTrue this_isTextStream(aoTextStream), TypeName(Me)&"+setTextStream()", "Not a TextStream object."

        Set PoTextStream = aoTextStream
        Set setTextStream = Me
        'Inbound、Outboundなどの情報を初期化する
        this_initialize
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : read()
    'Overview                    : ファイルから指定した文字数だけ読み込む
    'Detailed Description        : this_read()に委譲する
    'Argument
    '     alLength               : 読み込む文字数
    'Return Value
    '     読み込んだ文字列
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/02         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function read( _
        byVal alLength _
        )
        read = this_read(alLength)
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : readLine()
    'Overview                    : ファイルから1行読み込む
    'Detailed Description        : this_readLine()に委譲する
    'Argument
    '     なし
    'Return Value
    '     読み込んだ文字列
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/02         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function readLine( _
        )
        readLine = this_readLine()
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : readAll()
    'Overview                    : ファイル全体を読み込む
    'Detailed Description        : this_readAll()に委譲する
    'Argument
    '     なし
    'Return Value
    '     読み込んだ文字列
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/02         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function readAll( _
        )
        readAll = this_readAll()
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : skip()
    'Overview                    : ファイルから指定した文字数だけスキップする
    'Detailed Description        : this_read()に委譲する
    'Argument
    '     alLength               : スキップする文字数
    'Return Value
    '     なし
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/02         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Sub skip( _
        byVal alLength _
        )
        this_read alLength
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : skipLine()
    'Overview                    : ファイルから1行スキップする
    'Detailed Description        : this_readLine()に委譲する
    'Argument
    '     なし
    'Return Value
    '     なし
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/02         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Sub skipLine( _
        )
        this_readLine
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : close()
    'Overview                    : ファイル接続をクローズする
    'Detailed Description        : this_close()に委譲する
    'Argument
    '     なし
    'Return Value
    '     なし
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/02         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Sub close( _
        )
        this_close
    End Sub
    
    
    
    
    
    
    '***************************************************************************************************
    'Function/Sub Name           : this_read()
    'Overview                    : ファイルから指定した文字数だけ読み込む
    'Detailed Description        : 工事中
    'Argument
    '     alLength               : 読み込む文字数
    'Return Value
    '     読み取った文字列
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/02         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_read( _
        byVal alLength _
        )
        If PoTextStream Is Nothing Then Exit Function
        
        'バッファがなければファイルを読み取る
        Do While PoInbound.Item("atEndOfStream")=False And (PoBuffer.Item("length")-PoBuffer.Item("pointer")+1)<alLength
        'インバウンドが読み出し可能（atEndOfStream=False）かつバッファの未読み出し部分の長さが読み込む文字数未満の場合
            '読込バッファサイズだけ読み取る
            this_readFile False
        Loop
        
        'バッファから指定した文字数取り出す
        Dim sRet : sRet = Mid(PoBuffer.Item("buffer"), PoBuffer.Item("pointer"), alLength)
        
        'ポインタを更新
        PoBuffer.Item("pointer") = PoBuffer.Item("pointer")+Len(sRet)
        
        'アウトバウンドの情報を更新
        Dim boFlg : If PoBuffer.Item("pointer")>PoBuffer.Item("length") Then boFlg = True Else boFlg = False
        If boFlg Then
        'ポインタがバッファの外にある場合
            'インバウンドの状態をアウトバウンドにコピーする
            this_copyInboundStateToOutbound
        Else
        'ポインタがバッファ内にある場合
            Dim oArr : Set oArr = new_ArrSplit(Mid(PoBuffer.Item("buffer"), 1, PoBuffer.Item("pointer") - 1), vbLf)
            oArr.Reverse()
            With PoOutbound
                .Item("line") = oArr.length
                .Item("column") = Len(oArr(0))+1
                .Item("atEndOfStream") = PoInbound.Item("atEndOfStream") And (PoBuffer.Item("pointer") > PoBuffer.Item("length"))
                .Item("atEndOfLine") = .Item("atEndOfStream") Or (StrComp(Mid(PoBuffer.Item("buffer"), PoBuffer.Item("pointer"), 1), vbLf, vbBinaryCompare)=0)
            End With
        End If
        
        '戻り値を返す
        this_read = sRet
        Set oArr = Nothing
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : this_readLine()
    'Overview                    : ファイルから1行読み込む
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     読み取った文字列
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/02         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_readLine( _
        )
        If PoTextStream Is Nothing Then Exit Function
        
        'バッファがなければファイルを読み取る
        Do While PoInbound.Item("atEndOfStream")=False And InStr(PoBuffer.Item("pointer"), PoBuffer.Item("buffer"), vbLf, vbBinaryCompare)=0
        'インバウンドが読み出し可能（atEndOfStream=False）かつポインタのある行がバッファの最終行の場合
            '読込バッファサイズだけ読み取る
            this_readFile False
        Loop
        
        '行末（vbLf）を検索する
        Dim lPosRowEnd : lPosRowEnd = InStr(PoBuffer.Item("pointer"), PoBuffer.Item("buffer"), vbLf, vbBinaryCompare)
        Dim sRet
        If lPosRowEnd=0 Then
        '行末（vbLf）が見つからなかった＝ファイルの終端の場合
            'ポインタ以降全ての文字を返す
            sRet = Mid(PoBuffer.Item("buffer"), PoBuffer.Item("pointer"))
            'ファイルの終端にポインタを更新
            PoBuffer.Item("pointer") = PoBuffer.Item("length")+1
        Else
        '行末（vbLf）が見つかった＝ファイルの終端でない場合
            'ポインタから次の改行文字（vbLf）まで（改行文字を含まない）を返す
            sRet = Mid(PoBuffer.Item("buffer"), PoBuffer.Item("pointer"), lPosRowEnd-PoBuffer.Item("pointer"))
           '最後がvbCrの場合は削除する
           If StrComp(Right(sRet, 1), vbCr, vbBinaryCompare)=0 Then sRet = Mid(sRet, 1, Len(sRet)-1)
           '次の行の行頭（見つかった行末の位置+1）にポインタを更新
           PoBuffer.Item("pointer") = lPosRowEnd+1
        End If
        
        'アウトバウンドの情報を更新
        Dim boFlg : If PoBuffer.Item("pointer")>PoBuffer.Item("length") Then boFlg = True Else boFlg = False
        If boFlg Then
        'ポインタがバッファの外にある場合
            'インバウンドの状態をアウトバウンドにコピーする
            this_copyInboundStateToOutbound
        Else
        'ポインタがバッファ内にある場合
            With PoOutbound
                '次の行の行頭に更新する
                .Item("line") = .Item("line")+1
                .Item("column") = 1
                .Item("atEndOfStream") = False
                .Item("atEndOfLine") = (StrComp(Mid(PoBuffer.Item("buffer"), PoBuffer.Item("pointer"), 1), vbLf, vbBinaryCompare)=0)
            End With
        End If
        
        '戻り値を返す
        this_readLine = sRet
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : this_readAll()
    'Overview                    : ファイル全体を読み取る
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     読み取った文字列
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/02         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_readAll( _
        )
        If PoTextStream Is Nothing Then Exit Function
        
        'ファイル全体を読み取る
        Dim sRet : sRet = this_readFile(True)
        
        'インバウンドの状態をアウトバウンドにコピーする
        this_copyInboundStateToOutbound
        'ポインタを更新
        PoBuffer.Item("pointer") = PoBuffer.Item("length")+1
        
        '戻り値を返す
        this_readAll = sRet
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : this_readFile()
    'Overview                    : 指定した方法でファイルを読み込んでバッファに書き込む
    'Detailed Description        : 読込んだ後にインバウンドの状態を取得する
    'Argument
    '     aboIsReadAll           : ファイルの読み取り方法
    '                                True :ファイル全体を読み取る
    '                                False:読込バッファサイズだけ読み取る
    'Return Value
    '     読み取った文字列
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/02         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_readFile( _
        byVal aboIsReadAll _
        )
        
        'ファイルを読込む
        Dim sText : sText = ""
        If aboIsReadAll Then
            sText = PoTextStream.ReadAll
        Else
            sText = PoTextStream.Read(PlReadSize)
        End If
        'バッファの更新
        With PoBuffer
            .Item("buffer") = .Item("buffer") & sText
            .Item("length") = Len(.Item("buffer"))
        End With
        'インバウンドの状態を取得する
        this_getInboundStatus
        '戻り値を返す
        this_readFile = sText
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : this_close()
    'Overview                    : ファイル接続をクローズする
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     なし
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/02         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub this_close( _
        )
        If PoTextStream Is Nothing Then Exit Sub
        
        PoTextStream.Close
        Set PoTextStream = Nothing
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : this_getInboundStatus()
    'Overview                    : インバウンドの状態を取得する
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     なし
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/02         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub this_getInboundStatus( _
        )
        With PoTextStream
            'インバウンドの状態を取得する
            Set PoInbound = new_DicOf(Array("line", .Line, "column", .Column, "atEndOfLine", .AtEndOfLine, "atEndOfStream", .AtEndOfStream))
        End With
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : this_copyInboundStateToOutbound()
    'Overview                    : インバウンドの状態をアウトバウンドにコピーする
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     なし
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/02         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub this_copyInboundStateToOutbound( _
        )
        With PoInbound
            'アウトバウンドの状態にインバウンドの状態をコピーする
            Dim sKey, oOutbound
            Set oOutbound = new_Dic()
            For Each sKey In Array("line", "column", "atEndOfLine", "atEndOfStream")
                oOutbound.Add sKey, .Item(sKey)
            Next
        End With
        Set PoOutbound = oOutbound
        Set oOutbound = Nothing
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : this_initialize()
    'Overview                    : Inbound、Outboundなどの情報を初期化する
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     なし
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/02         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub this_initialize( _
        )
        'インバウンドの状態を取得する
        this_getInboundStatus
        'インバウンドの状態をアウトバウンドにコピーする
        this_copyInboundStateToOutbound
        'ポインタの初期化
        Set PoBuffer = new_DicOf(Array("pointer", 1, "buffer", "", "length", 0))
    End Sub

    '***************************************************************************************************
    'Function/Sub Name           : this_isTextStream()
    'Overview                    : オブジェクトがTextStreamか検査する
    'Detailed Description        : 工事中
    'Argument
    '     aoObj                  : オブジェクト
    'Return Value
    '     結果 True:TextStreamである / False:TextStreamでない
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2024/09/29         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_isTextStream( _
        byRef aoObj _
        )
        this_isTextStream = _
                cf_isSame(Vartype(aoObj),vbObject) _
                And _
                cf_isSame(Typename(aoObj),Typename(new_Ts(WScript.ScriptFullName,1,False,-2)))
    End Function
    
End Class
