'***************************************************************************************************
'FILENAME                    : clsCmBufferedWriter.vbs
'Overview                    : ファイル出力バッファリング処理クラス
'Detailed Description        : 工事中
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/01/07         Y.Fujii                  First edition
'***************************************************************************************************
Class clsCmBufferedWriter
    'クラス内変数、定数
    Private PoTextStream, PlWriteBufferSize, PlWriteIntervalTime, PoOutbound, PoInbound, PoBuffer
    
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
    '2023/08/20         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub Class_Initialize()
        Set PoTextStream = Nothing
        PlWriteBufferSize = 5000                 'デフォルトは5000バイト
        PlWriteIntervalTime = 0                  'デフォルトは0秒
        
        Dim vArr : vArr = Array("line", Empty, "column", Empty)
        Set PoOutbound = new_DicWith(vArr)
        Set PoInbound = new_DicWith(vArr)
        Set PoBuffer = new_DicWith(Array("buffer", Empty, "length", Empty, "lastWriteTime", Empty, "firstRequestTime", Empty))
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
    '2023/08/20         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub Class_Terminate()
        sub_CmBufferedWriterClose
        Set PoOutbound = Nothing
        Set PoInbound = Nothing
        Set PoBuffer = Nothing
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Let writeBufferSize()
    'Overview                    : 出力バッファサイズを設定する
    'Detailed Description        : 出力要求時に出力バッファのサイズがこれを超えた場合
    '                              ファイルに出力する
    'Argument
    '     alWriteBufferSize      : 出力バッファサイズ
    'Return Value
    '     なし
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/08/20         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Let writeBufferSize( _
        byVal alWriteBufferSize _
        )
        If func_CM_ValidationlIsWithinTheRangeOf(alWriteBufferSize, 2) Then
            PlWriteBufferSize = CLng(alWriteBufferSize)
        Else
            Err.Raise 1031, "clsCmBufferedWriter.vbs:clsCmBufferedWriter+writeBufferSize()", "不正な数字です。"
        End If
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get writeBufferSize()
    'Overview                    : 出力バッファサイズを返す
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     出力バッファサイズ
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/08/20         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get writeBufferSize()
        writeBufferSize = PlWriteBufferSize
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Let writeIntervalTime()
    'Overview                    : 出力間隔時間（秒）を設定する
    'Detailed Description        : 出力要求時に前回出力してから出力間隔時間を超えた場合
    '                              出力バッファの内容がサイズ未満でもファイルに出力する
    'Argument
    '     alWriteIntervalTime    : 出力間隔時間（秒）
    'Return Value
    '     なし
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/08/20         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Let writeIntervalTime( _
        byVal alWriteIntervalTime _
        )
        If func_CM_ValidationlIsWithinTheRangeOf(alWriteIntervalTime, 2) Then
            PlWriteIntervalTime = CLng(alWriteIntervalTime)
        Else
            Err.Raise 1031, "clsCmBufferedWriter.vbs:clsCmBufferedWriter+writeIntervalTime()", "不正な数字です。"
        End If
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get writeIntervalTime()
    'Overview                    : 出力間隔時間（秒）を返す
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     出力間隔時間（秒）
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/08/20         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get writeIntervalTime()
        writeIntervalTime = PlWriteIntervalTime
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
    '2023/08/27         Y.Fujii                  First edition
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
    '2023/10/17         Y.Fujii                  First edition
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
    '2023/10/17         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get column()
        column = PoOutbound.Item("column")
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get currentBufferSize()
    'Overview                    : 今のバッファサイズを返す
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     今のバッファサイズ
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/08/20         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get currentBufferSize()
        currentBufferSize = PoBuffer.Item("length")
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get lastWriteTime()
    'Overview                    : 最終ファイル出力日時を返す
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     最終ファイル出力日時
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/08/20         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get lastWriteTime()
        If IsEmpty(PoBuffer.Item("lastWriteTime")) Then
            lastWriteTime = Empty
        Else
            lastWriteTime = PoBuffer.Item("lastWriteTime")
        End If
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
    '2023/08/27         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function setTextStream( _
        byRef aoTextStream _
        )
        If Not func_CM_UtilIsTextStream(aoTextStream) Then
            Err.Raise 438, "clsCmBufferedWriter.vbs:clsCmBufferedWriter+setTextStream()", "オブジェクトでサポートされていないプロパティまたはメソッドです。"
            Exit Function
        End If

        Set PoTextStream = aoTextStream
        Set setTextStream = Me
        'Inbound、Outboundを最新化する
        sub_CmBufferedWriterUpdateStatus
        'バッファの初期化
        Set PoBuffer = new_DicWith(Array("buffer", "", "length", 0, "lastWriteTime", Empty, "firstRequestTime", Empty))
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : write()
    'Overview                    : 指定したテキストをファイルに書き込む
    'Detailed Description        : sub_CmBufferedWriterWrite()に委譲する
    'Argument
    '     asCont                 : 内容
    'Return Value
    '     なし
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/01/07         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Sub write( _
        byVal asCont _
        )
        sub_CmBufferedWriterWriteBuffer asCont
        sub_CmBufferedWriterWrite
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : writeBlankLines()
    'Overview                    : 指定した数の改行文字をファイルに書き込む
    'Detailed Description        : sub_CmBufferedWriterWriteFile()に委譲する
    'Argument
    '     alLines                : ファイルに書き込む改行文字の数
    'Return Value
    '     なし
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/09         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Sub writeBlankLines( _
        byVal alLines _
        )
        Dim sTmp, lIdx
        sTmp = ""
        For lIdx=1 To alLines 
            sTmp = sTmp & vbNewLine
        Next
        sub_CmBufferedWriterWriteBuffer sTmp
        sub_CmBufferedWriterWrite
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : writeLine()
    'Overview                    : 指定したテキストと改行をファイルに書き込む
    'Detailed Description        : sub_CmBufferedWriterWrite()に委譲する
    'Argument
    '     asCont                 : 内容
    'Return Value
    '     なし
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/09         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Sub writeLine( _
        byVal asCont _
        )
        sub_CmBufferedWriterWriteBuffer asCont & vbNewLine
        sub_CmBufferedWriterWrite
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : newLine()
    'Overview                    : 改行文字をファイルに書き込む
    'Detailed Description        : sub_CmBufferedWriterWriteFile()に委譲する
    'Argument
    '     なし
    'Return Value
    '     なし
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/08/27         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Sub newLine( _
        )
        sub_CmBufferedWriterWriteBuffer vbNewLine
        sub_CmBufferedWriterWrite
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : flush()
    'Overview                    : バッファに溜めた内容をファイルに出力する
    'Detailed Description        : sub_CmBufferedWriterWriteFile()に委譲する
    'Argument
    '     なし
    'Return Value
    '     なし
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/08/27         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Sub flush( _
        )
        sub_CmBufferedWriterWriteFile
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : close()
    'Overview                    : ファイル接続をクローズする
    'Detailed Description        : sub_CmBufferedWriterClose()に委譲する
    'Argument
    '     なし
    'Return Value
    '     なし
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/08/27         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Sub close( _
        )
        sub_CmBufferedWriterClose
    End Sub
    
    
    
    
    
    '***************************************************************************************************
    'Function/Sub Name           : sub_CmBufferedWriterWrite()
    'Overview                    : ファイル出力する
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     なし
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/01/09         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub sub_CmBufferedWriterWrite( _
        )
        'ファイル出力判定＆ファイル出力
        If func_CmBufferedWriterDecideToWrite() Then Call sub_CmBufferedWriterWriteFile()
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : func_CmBufferedWriterDecideToWrite()
    'Overview                    : ファイル出力するか判断する
    'Detailed Description        : 以下の条件で判断する
    '                              ・バッファのサイズが出力バッファサイズを超える
    '                              ・出力日時から出力間隔時間（秒）を経過した
    'Argument
    '     なし
    'Return Value
    '     結果 True:ファイルに出力する / False:ファイルに出力しない
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/08/20         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function func_CmBufferedWriterDecideToWrite( _
        )
        func_CmBufferedWriterDecideToWrite=False
        If PoTextStream Is Nothing Then Exit Function
        
        '戻り値の初期化
        Dim boRet : boRet=False

        If IsEmpty(PoBuffer.Item("firstRequestTime")) Then
        '初回の出力日時がない場合、初回リクエスト日時を設定する
            Set PoBuffer.Item("firstRequestTime") = new_Now()
        End If
        
        'バッファサイズの判定
        If PoBuffer.Item("length")>=PlWriteBufferSize Then boRet=True
        
        If boRet Or PlWriteIntervalTime<=0 Then
        'バッファのサイズが出力バッファサイズを超えたか出力日時から出力間隔時間（秒）が0以下（＝不要）の場合は関数を抜ける
            func_CmBufferedWriterDecideToWrite=boRet
            Exit Function
        End If
        
        '比較用日時の取得
        Dim oForComparison
        If IsEmpty(PoBuffer.Item("lastWriteTime")) Then
        '前回の出力日時がない場合、初回リクエスト日時を使用する
            Set oForComparison = PoBuffer.Item("firstRequestTime")
        Else
        '前回の出力日時がある場合、最終ファイル出力日時を使用する
            Set oForComparison = PoBuffer.Item("lastWriteTime")
        End If
        
        '出力日時の判定
        If Abs(oForComparison.differenceFrom(new_Now()))>=PlWriteIntervalTime Then boRet=True
        
        '戻り値を返す
        func_CmBufferedWriterDecideToWrite=boRet
        
        Set oForComparison = Nothing
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : sub_CmBufferedWriterWriteFile()
    'Overview                    : バッファの内容をファイルに出力する
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     なし
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/08/20         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub sub_CmBufferedWriterWriteFile( _
        )
        If PoTextStream Is Nothing Then Exit Sub
        
        'ファイルに出力
        PoTextStream.Write PoBuffer.Item("buffer")
        'Inbound、Outboundを最新化する
        sub_CmBufferedWriterUpdateStatus
        With PoBuffer
            .Item("buffer") = ""                      'バッファのクリア
            .Item("length") = 0                       'バッファ長を0にする
            Set .Item("lastWriteTime") = new_Now()    '出力日時を記録
        End With
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : sub_CmBufferedWriterClose()
    'Overview                    : ファイル接続をクローズする
    'Detailed Description        : バッファの未出力分を出力後にファイル接続をクローズする
    'Argument
    '     なし
    'Return Value
    '     なし
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/08/20         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub sub_CmBufferedWriterClose( _
        )
        If PoTextStream Is Nothing Then Exit Sub
        
        'バッファが残っていたら出力する
        If PoBuffer.Item("length")<>0 Then Call sub_CmBufferedWriterWriteFile()
        'テキストストリームをクローズする
        PoTextStream.Close
        Set PoTextStream = Nothing
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : sub_CmBufferedWriterWriteBuffer()
    'Overview                    : バッファに書き込む
    'Detailed Description        : 工事中
    'Argument
    '     asCont                 : 内容
    'Return Value
    '     なし
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/17         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub sub_CmBufferedWriterWriteBuffer( _
        byVal asCont _
        )
        Dim oArr

        With PoBuffer
            .Item("buffer") = .Item("buffer") & asCont
            .Item("length") = func_CM_StrLen(.Item("buffer"))
        End With

        Set oArr = new_ArrSplit(asCont, vbLf)
        oArr.reverse()
        With PoOutbound
            .Item("line") = .Item("line") + oArr.length-1
            If oArr.length=1 Then
                .Item("column") = .Item("column") + Len(oArr(0))
            Else
                .Item("column") = Len(oArr(0))+1
            End If
        End With

        Set oArr = Nothing
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : sub_CmBufferedWriterGetInboundStatus()
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
    '2023/10/17         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub sub_CmBufferedWriterGetInboundStatus( _
        )
        With PoTextStream
            'インバウンドの状態を取得する
            Set PoInbound = new_DicWith(Array("line", .Line, "column", .Column))
        End With
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : sub_CmBufferedWriterCopyInboundStateToOutbound()
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
    '2023/10/17         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub sub_CmBufferedWriterCopyInboundStateToOutbound( _
        )
        With PoInbound
            'アウトバウンドの状態にインバウンドの状態をコピーする
            Dim sKey, oOutbound
            Set oOutbound = new_Dic()
            For Each sKey In Array("line", "column")
                oOutbound.Add sKey, .Item(sKey)
            Next
        End With
        Set PoOutbound = oOutbound
        Set oOutbound = Nothing
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : sub_CmBufferedWriterUpdateStatus()
    'Overview                    : Inbound、Outboundを最新化する
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     なし
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/17         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub sub_CmBufferedWriterUpdateStatus( _
        )
        'インバウンドの状態を取得する
        sub_CmBufferedWriterGetInboundStatus
        'インバウンドの状態をアウトバウンドにコピーする
        sub_CmBufferedWriterCopyInboundStateToOutbound
    End Sub
    
End Class
