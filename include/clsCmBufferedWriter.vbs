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
    Private PsPath
    Private PsPathAlreadyOpened
    Private PsBuffer
    Private PoIomodeLst
    Private PsIomode              '入力/出力モード
    Private PoFileFormatLst
    Private PsFileFormat          'ファイルの形式
    
    'コンストラクタ
    Private Sub Class_Initialize()
        
        Set PoTextStream = Nothing
        PsPath = ""
        PsPathAlreadyOpened = ""
        PsBuffer = ""
        
        Set PoIomodeLst = CreateObject("Scripting.Dictionary")
        With PoIomodeLst
            .Add "ForReading", 1
            .Add "ForWriting", 2
            .Add "ForAppending", 8
        End With
        PsIomode = PoIomodeLst.Item(1)           'デフォルトはForAppending
        
        Set PoFileFormatLst = CreateObject("Scripting.Dictionary")
        With PoFileFormatLst
            .Add "TristateUseDefault", -2
            .Add "TristateTrue", -1
            .Add "TristateFalse", 0
        End With
        PsFileFormat = PoFileFormatLst.Item(1)   'デフォルトはTristateUseDefault
    End Sub
    
    'デストラクタ
    Private Sub Class_Terminate()
        Set PoFormatLst = Nothing
        Set PoIomodeLst = Nothing
        Set PoTextStream = Nothing
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
    'Function/Sub Name           : WriteContents()
    'Overview                    : ファイル出力する
    'Detailed Description        : sub_CmBufferedWriterWriteContents()に委譲する
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
        Call sub_CmBufferedWriterWriteContents(asContents)
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : sub_CmBufferedWriterWriteContents()
    'Overview                    : ファイル出力する
    'Detailed Description        : sub_CmBufferedWriterWriteContents()に委譲する
    'Argument
    '     asContents             : 内容
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
        
        'テキストストリームを作成する
        Call sub_CmBufferedWriterGetTextStream()
        
        PsBuffer = PsBuffer & vbCrLf & asContents
    End Sub
    
    
    '***************************************************************************************************
    'Function/Sub Name           : sub_CmBufferedWriterGetTextStream()
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
    Private Sub sub_CmBufferedWriterGetTextStream( _
        )
        
        If Len(PsPath)=0 Then Exit Sub
        If func_CM_FsIsSame(PsPath, PsPathAlreadyOpened) Then
        
        
        If PoTextStream Is Nothing Then
        'PoTextStreamがなければ作成する
            Dim boFileExists : boFileExists = func_CM_FsFileExists(PsPath)
            If Not boFileExists Then
            '出力先ファイルのパスが存在しない場合
                Dim sParentFolderPath : sParentFolderPath = func_CM_FsGetParentFolderPath(PsPath)
                Dim boParentFolderExists : boParentFolderExists = func_CM_FsFolderExists(sParentFolderPath)
                If Not boParentFolderExists Then
                '出力先ファイルの親フォルダが存在しない場合、フォルダを作成
                    Call func_CM_FsCreateFolder(sParentFolderPath)
                End If
            End If
            
            'ファイルを開く
            Set PoTextStream = func_CM_FsOpenTextFile(PsPath, PoIomodeLst.Item(PsIomode) _
                                                  , True, PoFileFormatLst.Item(PsFileFormat))
            PsPathAlreadyOpened = PsPath
        End If
    End Sub

End Class
