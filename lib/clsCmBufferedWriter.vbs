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
    Private PoTextStream, PoWriteDateTime, PoRequestFirstDateTime, PlWriteBufferSize, PlWriteIntervalTime, PsBuffer
    
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
        Set PoWriteDateTime = Nothing
        Set PoRequestFirstDateTime = Nothing
        PsBuffer = ""
        
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
        Call sub_CmBufferedWriterClose()
        Set PoWriteDateTime = Nothing
        Set PoRequestFirstDateTime = Nothing
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
        currentBufferSize = func_CM_StrLen(PsBuffer)
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
        If PoWriteDateTime Is Nothing Then
            lastWriteTime=""
        Else
            lastWriteTime = PoWriteDateTime
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
        Set PoTextStream = aoTextStream
        Set setTextStream = Me
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : write()
    'Overview                    : 指定したテキストをファイルに書き込む
    'Detailed Description        : sub_CmBufferedWriterWrite()に委譲する
    'Argument
    '     asContents             : 内容
    'Return Value
    '     なし
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/01/07         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Sub write( _
        byVal asContents _
        )
        PsBuffer = PsBuffer & asContents
        Call sub_CmBufferedWriterWrite()
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
        PsBuffer = PsBuffer & String(alLines, vbNewLine)
        Call sub_CmBufferedWriterWrite()
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : writeLine()
    'Overview                    : 指定したテキストと改行をファイルに書き込む
    'Detailed Description        : sub_CmBufferedWriterWrite()に委譲する
    'Argument
    '     asContents             : 内容
    'Return Value
    '     なし
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/09         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Sub writeLine( _
        byVal asContents _
        )
        PsBuffer = PsBuffer & asContents & vbNewLine
        Call sub_CmBufferedWriterWrite()
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
        PsBuffer = PsBuffer & vbNewLine
        Call sub_CmBufferedWriterWrite()
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
        Call sub_CmBufferedWriterWriteFile()
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
        Call sub_CmBufferedWriterClose()
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
        If func_CmBufferedWriterDetermineToWrite() Then Call sub_CmBufferedWriterWriteFile()
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : func_CmBufferedWriterDetermineToWrite()
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
    Private Function func_CmBufferedWriterDetermineToWrite( _
        )
        func_CmBufferedWriterDetermineToWrite=False
        If PoTextStream Is Nothing Then Exit Function
        
        '戻り値の初期化
        Dim boReturn : boReturn=False
        
        'バッファサイズの判定
        If func_CM_StrLen(PsBuffer)>=PlWriteBufferSize Then boReturn=True
        
        If boReturn Or PlWriteIntervalTime<=0 Then
        'バッファのサイズが出力バッファサイズを超えたか出力日時から出力間隔時間（秒）が0以下（＝不要）の場合は関数を抜ける
            func_CmBufferedWriterDetermineToWrite=boReturn
            Exit Function
        End If
        
        If PoWriteDateTime Is Nothing And PoRequestFirstDateTime Is Nothing Then
        '前回と初回の出力日時がない場合、本リクエスト（＝初回リクエスト）日時を取得して関数を抜ける
            Set PoRequestFirstDateTime = new_Now()
            func_CmBufferedWriterDetermineToWrite=boReturn
            Exit Function
        End If
        
        '比較用日時の取得
        Dim oForComparison
        Set oForComparison = PoWriteDateTime
        If oForComparison Is Nothing Then
        '前回の出力日時がない場合、初回リクエスト日時を使用する
            Set oForComparison = PoRequestFirstDateTime
        End If
        
        '出力日時の判定
        If Abs(oForComparison.differenceFrom(new_Now()))>=PlWriteIntervalTime Then boReturn=True
        
        '戻り値を返す
        func_CmBufferedWriterDetermineToWrite=boReturn
        
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
        Call PoTextStream.Write(PsBuffer)
        'バッファのクリア
        PsBuffer = ""
        '出力日時を記録
        Set PoWriteDateTime = new_Now()
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
        If func_CM_StrLen(PsBuffer)<>0 Then Call sub_CmBufferedWriterWriteFile()
        'テキストストリームをクローズする
        Call PoTextStream.Close
        Set PoTextStream = Nothing
    End Sub

End Class
