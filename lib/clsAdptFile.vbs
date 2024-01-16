'***************************************************************************************************
'FILENAME                    : clsAdptFile.vbs
'Overview                    : Fileオブジェクトのアダプタークラス
'Detailed Description        : Fileオブジェクトと同じIFを提供する
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/11/26         Y.Fujii                  First edition
'***************************************************************************************************
Class clsAdptFile
    'クラス内変数、定数
    Private PsTypeName,PoFile,PsPath
    
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
    '2023/11/26         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub Class_Initialize()
        PsTypeName = "FolderItem2"
        Set PoFile = Nothing
        PsPath = Empty
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : Class_Terminate()
    'Overview                    : デストラクタ
    'Detailed Description        : 終了処理
    'Argument
    '     なし
    'Return Value
    '     なし
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/11/26         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub Class_Terminate()
        Set PoFile = Nothing
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get DateLastModified()
    'Overview                    : ファイルの最終更新日時を返す
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     ファイルの最終更新日時
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/11/26         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get DateLastModified()
        DateLastModified = PoFile.ModifyDate
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get Name()
    'Overview                    : ファイルの名前を返す
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     ファイルの名前
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/11/26         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get Name()
        Name = new_Fso().GetFileName(PsPath)
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get ParentFolder()
    'Overview                    : ファイルの親フォルダーのフルパスを返す
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     ファイルの親フォルダーのフルパス
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/11/26         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get ParentFolder()
'        ParentFolder = PoFile.Parent.Self.Path
        ParentFolder = new_Fso().GetParentFolderName(PsPath)
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get Path()
    'Overview                    : ファイルのパスを返す
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     ファイルのパス
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/11/26         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Default Property Get Path()
        Path = PsPath
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get Size()
    'Overview                    : ファイルのサイズをバイト単位で返す
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     ファイルのサイズ（バイト単位）
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/11/26         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get Size()
        Size = PoFile.Size
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get FileType()
    'Overview                    : ファイルの種類を返す
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     ファイルの種類
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/11/26         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get [Type]()
        [Type] = PoFile.Type
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : setFileObject()
    'Overview                    : ファイルのオブジェクトを設定する
    'Detailed Description        : 工事中
    'Argument
    '     aoFile                 : FolderItem2オブジェクト
    'Return Value
    '     自身のインスタンス
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/11/26         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function setFileObject( _
        byRef aoFile _
        )
        If Not cf_isSame(PsTypeName, Typename(aoFile)) Then
            PsPath = Empty
            Set PoFile = Nothing
            Err.Raise 438, "clsAdptFile.vbs:clsAdptFile+setFileObject()", "オブジェクトでサポートされていないプロパティまたはメソッドです。"
            Exit Function
        End If
        
        PsPath = aoFile.Path
        Set PoFile = aoFile

        Set setFileObject = Me
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : setFilePath()
    'Overview                    : ファイルのパスからオブジェクトを設定する
    'Detailed Description        : 工事中
    'Argument
    '     asPath                 : ファイルのパス
    'Return Value
    '     自身のインスタンス
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/12/19         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function setFilePath( _
        byVal asPath _
        )
        If Not new_Fso().FileExists(asPath) Then
            PsPath = Empty
            Set PoFile = Nothing
            Err.Raise 76, "clsAdptFile.vbs:clsAdptFile+setFilePath()", "パスが見つかりません。"
            Exit Function
        End If
        
        PsPath = asPath
        Set PoFile = new_ShellApp().Namespace(new_Fso().GetParentFolderName(asPath)).Items().Item(new_Fso().GetFileName(asPath))

        Set setFilePath = Me
    End Function

End Class
