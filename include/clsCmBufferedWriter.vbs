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
    Private PoTextStream
    Private PoWriteDateTime
    Private PoRequestFirstDateTime
    Private PsPath
    Private PsPathAlreadyOpened
    Private PlWriteBufferSize
    Private PlWriteIntervalTime
    Private PsBuffer
    Private PoIomodeLst
    Private PsIomode              '入力/出力モード
    Private PoFileFormatLst
    Private PsFileFormat          'ファイルの形式
    
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
        PsPath = ""
        PsPathAlreadyOpened = ""
        PlWriteBufferSize = 5000                 'デフォルトは5000バイト
        PlWriteIntervalTime = 0                  'デフォルトは0秒
        Set PoWriteDateTime = Nothing
        Set PoRequestFirstDateTime = Nothing
        PsBuffer = ""
        
        Set PoIomodeLst = CreateObject("Scripting.Dictionary")
        With PoIomodeLst
            .Add "ForReading", 1
            .Add "ForWriting", 2
            .Add "ForAppending", 8
            PsIomode = .Keys()(2)                'デフォルトはForAppending
        End With
        
        Set PoFileFormatLst = CreateObject("Scripting.Dictionary")
        With PoFileFormatLst
            .Add "TristateUseDefault", -2
            .Add "TristateTrue", -1
            .Add "TristateFalse", 0
            PsFileFormat = .Keys()(0)            'デフォルトはTristateUseDefault
        End With
        
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
        Call sub_CmBufferedWriterCloseFile()
        Set PoWriteDateTime = Nothing
        Set PoRequestFirstDateTime = Nothing
        Set PoWriteDateTime = Nothing
        Set PoFormatLst = Nothing
        Set PoIomodeLst = Nothing
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Let Outpath()
    'Overview                    : 出力先ファイルのパスを設定する
    'Detailed Description        : 工事中
    'Argument
    '     asPath                 : ファイルのパス
    'Return Value
    '     なし
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/01/07         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Let Outpath( _
        byVal asPath _
        )
        PsPath = asPath
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get Outpath()
    'Overview                    : 出力先ファイルのパスを返す
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     出力先ファイルのパス
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/01/07         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get Outpath()
        Outpath = PsPath
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Let WriteBufferSize()
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
    Public Property Let WriteBufferSize( _
        byVal alWriteBufferSize _
        )
        If -2147483648<=alWriteBufferSize And alWriteBufferSize<=2147483647 Then
        'Long型の範囲（-2,147,483,648 ～ 2,147,483,647）の場合
            PlWriteBufferSize = CLng(alWriteBufferSize)
        End If
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get WriteBufferSize()
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
    Public Property Get WriteBufferSize()
        WriteBufferSize = PlWriteBufferSize
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Let WriteIntervalTime()
    'Overview                    : 出力間隔時間（秒）を設定する
    'Detailed Description        : 出力要求時に前回出力してから出力間隔時間を超えた場合
    '                              出力バッファの内容がサイズ未満でもファイルに出力する
    '                              設定値が0以下の場合はこの判断をしない
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
    Public Property Let WriteIntervalTime( _
        byVal alWriteIntervalTime _
        )
        If -2147483648<=alWriteIntervalTime And alWriteIntervalTime<=2147483647 Then
        'Long型の範囲（-2,147,483,648 ～ 2,147,483,647）の場合
            PlWriteIntervalTime = CLng(alWriteIntervalTime)
        End If
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get WriteIntervalTime()
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
    Public Property Get WriteIntervalTime()
        WriteIntervalTime = PlWriteIntervalTime
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Let Iomode()
    'Overview                    : 入力/出力モードを設定する
    'Detailed Description        : 工事中
    'Argument
    '     asIomode               : 入力/出力モード "ForReading","ForWriting","ForAppending"
    'Return Value
    '     なし
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/01/09         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Let Iomode( _
        byVal asIomode _
        )
        If PoIomodeLst.Exists(asIomode) Then
            PsIomode = asIomode
        End If
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get Iomode()
    'Overview                    : 入力/出力モードを返す
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     出力先ファイルのパス
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/01/09         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get Iomode()
        Iomode = PsIomode
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Let FileFormat()
    'Overview                    : ファイルの形式を設定する
    'Detailed Description        : 工事中
    'Argument
    '     asFileFormat           : ファイルの形式 "TristateUseDefault","TristateTrue","TristateFalse"
    'Return Value
    '     なし
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/01/09         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Let FileFormat( _
        byVal asFileFormat _
        )
        If PoFileFormatLst.Exists(asFileFormat) Then
            PsFileFormat = asFileFormat
        End If
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get FileFormat()
    'Overview                    : ファイルの形式を返す
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     出力先ファイルのパス
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/01/09         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get FileFormat()
        FileFormat = PsFileFormat
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get CurrentBufferSize()
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
    Public Property Get CurrentBufferSize()
        CurrentBufferSize = func_CM_StrLen(PsBuffer)
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get LastWriteDateTime()
    'Overview                    : 最終ファイル出力日時
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
    Public Property Get LastWriteDateTime()
        If PoWriteDateTime Is Nothing Then
            LastWriteDateTime=""
        Else
            LastWriteDateTime = PoWriteDateTime.DisplayFormatAs("YYYY/MM/DD hh:mm:ss.000")
        End If
    End Property
    '***************************************************************************************************
    'Function/Sub Name           : WriteContents()
    'Overview                    : ファイル出力する
    'Detailed Description        : sub_CmBufferedWriterWriteFile()に委譲する
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
    Public Sub WriteContents( _
        byVal asContents _
        )
        PsBuffer = PsBuffer & asContents
        Call sub_CmBufferedWriterWriteContents()
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : WriteLineContents()
    'Overview                    : 1行ファイル出力する
    'Detailed Description        : sub_CmBufferedWriterWriteFile()に委譲する
    'Argument
    '     asContents             : 内容
    'Return Value
    '     なし
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/08/20         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Sub WriteLineContents( _
        byVal asContents _
        )
        PsBuffer = PsBuffer & asContents & vbNewLine
        Call sub_CmBufferedWriterWriteContents()
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : sub_CmBufferedWriterWriteContents()
    'Overview                    : ファイル出力する
    'Detailed Description        : sub_CmBufferedWriterWriteContents()に委譲する
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
    Private Sub sub_CmBufferedWriterWriteContents( _
        byVal asContents _
        )
        'ファイル出力判定＆ファイル出力
        If func_CmBufferedWriterDetermineToWrite() Then Call sub_CmBufferedWriterWriteFile()
    End Sub
    
    
    '***************************************************************************************************
    'Function/Sub Name           : sub_CmBufferedWriterCreateTextStream()
    'Overview                    : テキストストリームを作成する
    'Detailed Description        : 工事中
    'Argument
    '     asContents             : 内容
    'Return Value
    '     なし
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/08/19         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub sub_CmBufferedWriterCreateTextStream( _
        )
        If Not PoTextStream Is Nothing Then
            If Len(PsPath)=0 Or Not func_CM_FsIsSame(PsPath, PsPathAlreadyOpened) Then
            '今のPoTextStreamの未出力分を処理した上で、クローズする
                Call sub_CmBufferedWriterCloseFile()
            End If
        End If
        
        If Len(PsPath)>0 Then
            If PoTextStream Is Nothing Or Not func_CM_FsIsSame(PsPath, PsPathAlreadyOpened) Then
            'PoTextStreamを新規作成する
                If Not func_CM_FsFileExists(PsPath) Then
                '出力先ファイルのパスが存在しない かつ 出力先ファイルの親フォルダが存在しない場合、フォルダを作成
                    Dim sParentFolderPath : sParentFolderPath = func_CM_FsGetParentFolderPath(PsPath)
                    If Not func_CM_FsFolderExists(sParentFolderPath) Then
                        Call func_CM_FsCreateFolder(sParentFolderPath)
                    End If
                End If
                
                'PoTextStreamを作成（ファイルがなければ新規作成）
                Set PoTextStream = func_CM_FsOpenTextFile(PsPath, PoIomodeLst.Item(PsIomode) _
                                                      , True, PoFileFormatLst.Item(PsFileFormat))
                '出力先ファイルパスを退避する
                PsPathAlreadyOpened = PsPath
            End If
        End If
        
    End Sub
    
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
        'テキストストリームを作成する
        Call sub_CmBufferedWriterCreateTextStream()
        
        If PoTextStream Is Nothing Then Exit Sub
        
        'ファイルに出力
        Call PoTextStream.Write(PsBuffer)
        'バッファのクリア
        PsBuffer = ""
        '出力日時を記録
        Set PoWriteDateTime = new_clsCmCalendar()
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : sub_CmBufferedWriterCloseFile()
    'Overview                    : ファイル出力を完了する
    'Detailed Description        : バッファの未出力分を出力の上、ファイル接続をクローズする
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
    Private Sub sub_CmBufferedWriterCloseFile( _
        )
        If PoTextStream Is Nothing Then Exit Sub
        'バッファが残っていたら出力する
        If func_CM_StrLen(PsBuffer)<>0 Then Call sub_CmBufferedWriterWriteFile()
        'テキストストリームをクローズする
        Call PoTextStream.Close
        Set PoTextStream = Nothing
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
    '     なし
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
        
        Dim boReturn : boReturn=False
        
        'バッファサイズ
        If func_CM_StrLen(PsBuffer)>=PlWriteBufferSize Then boReturn=True
        
        '出力日時
        If PoWriteDateTime Is Nothing Then
        '初回書き込み前は初回リクエスト時からの経過時間で判断する
            If PoRequestFirstDateTime Is Nothing Then
                Set PoRequestFirstDateTime = new_clsCmCalendar()
            Else
                If Abs(PoRequestFirstDateTime.DifferenceInScondsFrom(new_clsCmCalendar()))>=PlWriteIntervalTime Then boReturn=True
            End If
        Else
            If Abs(PoWriteDateTime.DifferenceInScondsFrom(new_clsCmCalendar()))>=PlWriteIntervalTime Then boReturn=True
        End If
        func_CmBufferedWriterDetermineToWrite=boReturn
    End Function

End Class
