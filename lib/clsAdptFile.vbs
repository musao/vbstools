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
    Private PoCacheInfo,PoCache,PoFile,PsTypeName
    
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
'        Set PoCacheInfo = new_DicWith(Array("ValidityPeriod", 3, "LastReferencedDateTime", Empty))
'        sub_AdptFileInitCache
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
'        Set PoCacheInfo = Nothing
'        Set PoCache = Nothing
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
'        Name = PoFile.Name
        Name = new_Fso().GetFileName(PoFile.Path)
'        Name = func_AdptFileGet("Name")
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
        ParentFolder = new_Fso().GetParentFolderName(PoFile.Path)
'        ParentFolder = func_AdptFileGet("ParentFolder")
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
        Path = PoFile.Path
'        Path = func_AdptFileGet("Path")
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
'        Size = func_AdptFileGet("Size")
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
'        [Type] = func_AdptFileGet("Type")
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
            Err.Raise 438, "clsAdptFile.vbs:clsAdptFile+setFileObject()", "オブジェクトでサポートされていないプロパティまたはメソッドです。"
            Exit Function
        End If
        
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
            Err.Raise 76, "clsAdptFile.vbs:clsAdptFile+setFilePath()", "パスが見つかりません。"
            Exit Function
        End If
        
        Set PoFile = new_ShellApp().Namespace(new_Fso().GetParentFolderName(asPath)).Items().Item(new_Fso().GetFileName(asPath))
        Set setFilePath = Me
    End Function


    
'    '***************************************************************************************************
'    'Function/Sub Name           : sub_AdptFileInitCache()
'    'Overview                    : キャッシュを初期化する
'    'Detailed Description        : 工事中
'    'Argument
'    '     なし
'    'Return Value
'    '     なし
'    '---------------------------------------------------------------------------------------------------
'    'Histroy
'    'Date               Name                     Reason for Changes
'    '----------         ----------------------   -------------------------------------------------------
'    '2024/01/13         Y.Fujii                  First edition
'    '***************************************************************************************************
'    Public Sub sub_AdptFileInitCache( _
'        )
'        Set PoCache = new_DicWith(Array("DateLastModified", Empty, "Name", Empty, "ParentFolder", Empty, "Path", Empty, "Size", Empty, "Type", Empty))
'    End Sub
'    
'    '***************************************************************************************************
'    'Function/Sub Name           : func_AdptFileGet()
'    'Overview                    : 指定したプロパティを取得する
'    'Detailed Description        : 工事中
'    'Argument
'    '     asProp                 : プロパティを指定する文字列
'    'Return Value
'    '     プロパティの内容
'    '---------------------------------------------------------------------------------------------------
'    'Histroy
'    'Date               Name                     Reason for Changes
'    '----------         ----------------------   -------------------------------------------------------
'    '2024/01/13         Y.Fujii                  First edition
'    '***************************************************************************************************
'    Public Function func_AdptFileGet( _
'        byVal asProp _
'        )
'        'キャッシュ利用判定
'        Dim boUseCache : boUseCache=False
'        If Not IsEmpty(PoCacheInfo.Item("LastReferencedDateTime")) Then
'        '最終参照日時が空でない場合
'            If new_Now().differenceFrom(PoCacheInfo.Item("LastReferencedDateTime"))<PoCacheInfo.Item("ValidityPeriod") Then
'            '最終参照日時からキャッシュ有効期間を経過していない場合、対象のキャッシュがある
'                If Not IsEmpty(PoCache.Item(asProp)) Then boUseCache=True
'            End If
'        End If
'
'        If boUseCache Then
'        'キャッシュを使う場合
'            cf_bind func_AdptFileGet, PoCache.Item(asProp)
'            Exit Function
'        End If
'
'        'キャッシュを使用しない場合
'        sub_AdptFileInitCache
'        Select Case asProp
'            Case "DateLastModified"
'                PoCache.Item(asProp) = vRet
'            Case "Name","ParentFolder","Path"
'                Dim sPath : sPath = PoFile.Path
'                PoCache.Item("Name") = new_Fso().GetFileName(sPath)
'                PoCache.Item("ParentFolder") = new_Fso().GetParentFolderName(sPath)
'                PoCache.Item("Path") = sPath
'            Case "Size"
'                PoCache.Item(asProp) = PoFile.Size
'            Case "Type"
'                PoCache.Item(asProp) = PoFile.Type
'        End Select
'        cf_bind func_AdptFileGet, PoCache.Item(asProp)
'        Set PoCacheInfo.Item("LastReferencedDateTime") = new_Now()
'    End Function

End Class
