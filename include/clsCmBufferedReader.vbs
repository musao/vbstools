'***************************************************************************************************
'FILENAME                    : clsCmBufferedReader.vbs
'Overview                    : ファイル読込バッファリング処理クラス
'Detailed Description        : 工事中
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/10/02         Y.Fujii                  First edition
'***************************************************************************************************

'***************************************************************************************************
'Function/Sub Name           : new_clsCmBufferedReader()
'Overview                    : インスタンス生成関数
'Detailed Description        : 工事中
'Argument
'     aoTextStream           : テキストストリームオブジェクト
'Return Value
'     ファイル読込バッファリング処理クラスのインスタンス
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/10/02         Y.Fujii                  First edition
'***************************************************************************************************
Private Function new_clsCmBufferedReader( _
    byRef aoTextStream _
    )
    Set new_clsCmBufferedReader = (New clsCmBufferedReader).SetTextStream(aoTextStream)
End Function

Class clsCmBufferedReader
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
        Dim vArr : vArr = Array("Line", Empty, "Column", Empty, "AtEndOfLine", Empty, "AtEndOfStream", Empty)
        Set PoOutbound = new_DictSetValues(vArr)
        Set PoInbound = new_DictSetValues(vArr)
        Set PoBuffer = new_DictSetValues(Array("Buffer", Empty, "Pointer", Empty, "Length", Empty))
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
        Call sub_CmBufferedReaderClose()
        Set PoOutbound = Nothing
        Set PoInbound = Nothing
        Set PoBuffer = Nothing
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Let ReadSize()
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
    Public Property Let ReadSize( _
        byVal alReadSize _
        )
        If func_CM_ValidationlIsWithinTheRangeOf(alReadSize, 2) Then
            PlReadSize = CLng(alReadSize)
        End If
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get ReadSize()
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
    Public Property Get ReadSize()
        ReadSize = PlReadSize
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get TextStream()
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
    Public Property Get TextStream()
        Set TextStream = aoTextStream
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get Line()
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
    Public Property Get Line()
        Line = PoOutbound.Item("Line")
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get Column()
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
    Public Property Get Column()
        Column = PoOutbound.Item("Column")
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get AtEndOfStream()
    'Overview                    : 行末の場合にTrueを返す
    'Detailed Description        : 現行末マーカーの直前にファイル ポインターがある場合は true を返し、
    '                              そうでない場合は false を返します。
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
    Public Property Get AtEndOfStream()
        AtEndOfStream = PoOutbound.Item("AtEndOfStream")
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get AtEndOfLine()
    'Overview                    : ファイルの終端の場合にTrueを返す
    'Detailed Description        : 最後にファイル ポインターがある場合は true を返し、
    '                              そうでない場合は false を返します。
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
    Public Property Get AtEndOfLine()
        AtEndOfLine = PoOutbound.Item("AtEndOfLine")
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : SetTextStream()
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
    Public Function SetTextStream( _
        byRef aoTextStream _
        )
        Set PoTextStream = aoTextStream
        Set SetTextStream = Me
        
        'Inbound、Outboundなどの情報を初期化する
        Call sub_CmBufferedReaderInitialize()
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : Read()
    'Overview                    : ファイルから指定した文字数だけ読み込む
    'Detailed Description        : func_CmBufferedReaderRead()に委譲する
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
    Public Function Read( _
        byVal alLength _
        )
        Read = func_CmBufferedReaderRead(alLength)
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : ReadLine()
    'Overview                    : ファイルから1行読み込む
    'Detailed Description        : func_CmBufferedReaderReadLine()に委譲する
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
    Public Function ReadLine( _
        )
        ReadLine = func_CmBufferedReaderReadLine()
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : ReadAll()
    'Overview                    : ファイル全体を読み込む
    'Detailed Description        : func_CmBufferedReaderReadAll()に委譲する
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
    Public Function ReadAll( _
        )
        ReadAll = func_CmBufferedReaderReadAll()
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : Skip()
    'Overview                    : ファイルから指定した文字数だけスキップする
    'Detailed Description        : func_CmBufferedReaderRead()に委譲する
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
    Public Sub Skip( _
        byVal alLength _
        )
        Call func_CmBufferedReaderRead(alLength)
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : SkipLine()
    'Overview                    : ファイルから1行スキップする
    'Detailed Description        : func_CmBufferedReaderReadLine()に委譲する
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
    Public Sub SkipLine( _
        )
        Call func_CmBufferedReaderReadLine()
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : Close()
    'Overview                    : ファイル接続をクローズする
    'Detailed Description        : sub_CmBufferedReaderClose()に委譲する
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
    Public Sub Close( _
        )
        Call sub_CmBufferedReaderClose()
    End Sub
    
    
    
    
    
    
    '***************************************************************************************************
    'Function/Sub Name           : func_CmBufferedReaderRead()
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
    Private Function func_CmBufferedReaderRead( _
        byVal alLength _
        )
        If PoTextStream Is Nothing Then Exit Function
        
        'バッファがなければファイルを読み取る
        Do While PoInbound.Item("AtEndOfStream")=False And (PoBuffer.Item("Length")-PoBuffer.Item("Pointer")+1)<alLength
        'インバウンドが読み出し可能（AtEndOfStream=False）かつバッファの未読み出し部分の長さが読み込む文字数未満の場合
            '読込バッファサイズだけ読み取る
            Call func_CmBufferedReaderReadFile(False)
        Loop
        
        'バッファから指定した文字数取り出す
        Dim sRet : sRet = Mid(PoBuffer.Item("Buffer"), PoBuffer.Item("Pointer"), alLength)
        
        'ポインタを更新
        PoBuffer.Item("Pointer") = PoBuffer.Item("Pointer")+Len(sRet)
        
        'アウトバウンドの情報を更新
        Dim oArr : Set oArr = new_ArraySplit(Mid(PoBuffer.Item("Buffer"), 1, PoBuffer.Item("Pointer") - 1), vbLf)
        oArr.Reverse()
        With PoOutbound
            .Item("Line") = oArr.Length
            .Item("Column") = Len(oArr(0))+1
            .Item("AtEndOfStream") = PoInbound.Item("AtEndOfStream") And (PoBuffer.Item("Pointer") > PoBuffer.Item("Length"))
            .Item("AtEndOfLine") = .Item("AtEndOfStream") Or (StrComp(Mid(PoBuffer.Item("Buffer"), PoBuffer.Item("Pointer"), 1), vbLf, vbBinaryCompare)=0)
        End With
        
        '戻り値を返す
        func_CmBufferedReaderRead = sRet
        Set oArr = Nothing
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : func_CmBufferedReaderReadLine()
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
    Private Function func_CmBufferedReaderReadLine( _
        )
        If PoTextStream Is Nothing Then Exit Function
        
        'バッファがなければファイルを読み取る
        Do While PoInbound.Item("AtEndOfStream")=False And InStr(PoBuffer.Item("Pointer"), PoBuffer.Item("Buffer"), vbLf, vbBinaryCompare)=0
        'インバウンドが読み出し可能（AtEndOfStream=False）かつポインタのある行がバッファの最終行の場合
            '読込バッファサイズだけ読み取る
            Call func_CmBufferedReaderReadFile(False)
        Loop
        
        '行末（vbLf）を検索する
        Dim lPosRowEnd : lPosRowEnd = InStr(PoBuffer.Item("Pointer"), PoBuffer.Item("Buffer"), vbLf, vbBinaryCompare)
        Dim sRet
        If lPosRowEnd=0 Then
        '行末（vbLf）が見つからなかった＝ファイルの終端の場合
            'ポインタ以降全ての文字を返す
            sRet = Mid(PoBuffer.Item("Buffer"), PoBuffer.Item("Pointer"))
            'ファイルの終端にポインタを更新
            PoBuffer.Item("Pointer") = PoBuffer.Item("Length")+1
        Else
        '行末（vbLf）が見つかった＝ファイルの終端でない場合
            'ポインタから次の改行文字（vbLf）まで（改行文字を含まない）を返す
            sRet = Mid(PoBuffer.Item("Buffer"), PoBuffer.Item("Pointer"), lPosRowEnd-PoBuffer.Item("Pointer"))
           '最後がvbCrの場合は削除する
           If StrComp(Right(sRet, 1), vbCr, vbBinaryCompare)=0 Then sRet = Mid(sRet, 1, Len(sRet)-1)
           '次の行の行頭（見つかった行末の位置+1）にポインタを更新
           PoBuffer.Item("Pointer") = lPosRowEnd+1
        End If
        
        'アウトバウンドの情報を更新
        With PoOutbound
            Dim boEof : If .Item("Line")+1>PoInbound.Item("Line") Then boEof = True Else boEof = False
            If boEof Then
            'ファイルの終端まで読み出した場合
                'インバウンドの状態をアウトバウンドにコピーする
                Call sub_CmBufferedReaderCopyInboundStateToOutbound()
            Else
            'ファイルの終端まで読み出してない場合
                '次の行の行頭に更新する
                .Item("Line") = .Item("Line")+1
                .Item("Column") = 1
                .Item("AtEndOfStream") = False
                .Item("AtEndOfLine") = (StrComp(Mid(PoBuffer.Item("Buffer"), PoBuffer.Item("Pointer"), 1), vbLf, vbBinaryCompare)=0)
            End If
        End With
        
        '戻り値を返す
        func_CmBufferedReaderReadLine = sRet
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : func_CmBufferedReaderReadAll()
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
    Private Function func_CmBufferedReaderReadAll( _
        )
        If PoTextStream Is Nothing Then Exit Function
        
        'ファイル全体を読み取る
        Dim sRet : sRet = func_CmBufferedReaderReadFile(True)
        
        'インバウンドの状態をアウトバウンドにコピーする
        Call sub_CmBufferedReaderCopyInboundStateToOutbound()
        'ポインタを更新
        PoBuffer.Item("Pointer") = PoBuffer.Item("Length")+1
        
        '戻り値を返す
        func_CmBufferedReaderReadAll = sRet
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : func_CmBufferedReaderReadFile()
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
    Private Function func_CmBufferedReaderReadFile( _
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
            .Item("Buffer") = .Item("Buffer") & sText
            .Item("Length") = Len(.Item("Buffer"))
        End With
        'インバウンドの状態を取得する
        Call sub_CmBufferedReaderGetInboundStatus()
        '戻り値を返す
        func_CmBufferedReaderReadFile = sText
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : sub_CmBufferedReaderClose()
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
    Private Sub sub_CmBufferedReaderClose( _
        )
        If PoTextStream Is Nothing Then Exit Sub
        
        Call PoTextStream.Close
        Set PoTextStream = Nothing
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : sub_CmBufferedReaderGetInboundStatus()
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
    Private Sub sub_CmBufferedReaderGetInboundStatus( _
        )
        With PoTextStream
            'インバウンドの状態を取得する
            Set PoInbound = new_DictSetValues(Array("Line", .Line, "Column", .Column, "AtEndOfLine", .AtEndOfLine, "AtEndOfStream", .AtEndOfStream))
        End With
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : sub_CmBufferedReaderCopyInboundStateToOutbound()
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
    Private Sub sub_CmBufferedReaderCopyInboundStateToOutbound( _
        )
        With PoInbound
            'アウトバウンドの状態にインバウンドの状態をコピーする
            Dim sKey, oOutbound
            Set oOutbound = new_Dictionary()
            For Each sKey In Array("Line", "Column", "AtEndOfLine", "AtEndOfStream")
                oOutbound.Add sKey, .Item(sKey)
            Next
        End With
        Set PoOutbound = oOutbound
        Set oOutbound = Nothing
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : sub_CmBufferedReaderInitialize()
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
    Private Sub sub_CmBufferedReaderInitialize( _
        )
        'インバウンドの状態を取得する
        Call sub_CmBufferedReaderGetInboundStatus()
        'インバウンドの状態をアウトバウンドにコピーする
        Call sub_CmBufferedReaderCopyInboundStateToOutbound()
        'ポインタの初期化
        Set PoBuffer = new_DictSetValues(Array("Pointer", 1, "Buffer", "", "Length", 0))
    End Sub
    
End Class
