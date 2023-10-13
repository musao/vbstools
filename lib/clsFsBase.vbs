'***************************************************************************************************
'FILENAME                    : clsFsBase.vbs
'Overview                    : ファイル・フォルダ共通クラス
'Detailed Description        : 工事中
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/11/03         Y.Fujii                  First edition
'***************************************************************************************************
Class clsFsBase
    'クラス内変数、定数
    Private PoFso                                          'FileSystemObjectオブジェクト
    Private PoProp                                         '属性格納用ハッシュマップ
    Private PboUseCache                                    'キャッシュ使用可否（最新を取得するかどうか）
    Private PdbLastCacheConfirmationTime                   '最終キャッシュ確認時間（Timer関数の値）
    Private PdbLastCacheUpdateTime                         '最終キャッシュ更新時間（Timer関数の値）
    Private PdbValidPeriod                                 'キャッシュ有効期間（秒数）、最終キャッシュ確認時間からの経過時間
    
    'コンストラクタ
    Private Sub Class_Initialize()
        '初期化
        Set PoFso = Nothing
        PboUseCache = True
        PdbLastCacheConfirmationTime = 0
        PdbLastCacheUpdateTime = 0
        PdbValidPeriod = 1
        
        Set PoProp = new_Dic()
        With PoProp
            .Add "Attributes", vbNullString                '属性
            .Add "DateCreated", vbNullString               '作成された日付と時刻
            .Add "DateLastAccessed", vbNullString          '最後にアクセスした日付と時刻
            .Add "DateLastModified", vbNullString          '最終更新日時
            .Add "Drive", vbNullString                     'ファイルまたはフォルダーがあるドライブのドライブ文字
            .Add "Name", vbNullString                      '名前
            .Add "ParentFolder", vbNullString              '親のフォルダーオブジェクト
            .Add "Path", vbNullString                      'パス
            .Add "ShortName", vbNullString                 '短い名前(8.3 名前付け規則)
            .Add "ShortPath", vbNullString                 '短いパス(8.3 名前付け規則)
            .Add "Size", vbNullString                      'サイズ（バイト単位）
            .Add "Type", vbNullString                      '種類
        End With
    End Sub
    'デストラクタ
    Private Sub Class_Terminate()
        Set PoFso = Nothing
        Set PoProp = Nothing
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Let Patha()
    'Overview                    : パスを設定する
    'Detailed Description        : 工事中
    'Argument
    '     asPath                 : パス
    'Return Value
    '     なし
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2022/11/03         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Let Path( _
        byVal asPath _
        )
        PoProp.Item("Path") = asPath
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Let Fso()
    'Overview                    : 本インスタンスが使用するFileSystemObjectオブジェクトを設定する
    'Detailed Description        : 工事中
    'Argument
    '     aoFso                  : FileSystemObjectオブジェクト
    'Return Value
    '     なし
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2022/11/03         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Let Fso( _
        byRef aoFso _
        )
        Set PoFso = aoFso
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Let UseCache()
    'Overview                    : キャッシュ使用可否（最新を取得するかどうか）を設定する
    'Detailed Description        : 工事中
    'Argument
    '     aboUseCache            : キャッシュ使用可否
    'Return Value
    '     なし
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2022/11/03         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Let UseCache( _
        byVal aboUseCache _
        )
        PboUseCache = aboUseCache
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get UseCache()
    'Overview                    : キャッシュ使用可否（最新を取得するかどうか）を返す
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     キャッシュ使用可否
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2022/11/03         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get UseCache()
       UseCache = PboUseCache
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Let ValidPeriod()
    'Overview                    : キャッシュ有効期間（秒数）を設定する
    'Detailed Description        : 工事中
    'Argument
    '     adbValidPer            : キャッシュ有効期間（秒数）
    'Return Value
    '     なし
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2022/11/03         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Let ValidPeriod( _
        byVal adbValidPeriod _
        )
        PdbValidPeriod = adbValidPeriod
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get ValidPeriod()
    'Overview                    : キャッシュ有効期間（秒数）を返す
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     キャッシュ有効期間（秒数）
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2022/11/03         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get ValidPeriod()
       ValidPeriod = PdbValidPeriod
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get LastCacheConfirmationTime()
    'Overview                    : 最終キャッシュ確認時間（Timer関数の値）を返す
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     最終キャッシュ確認時間（Timer関数の値）
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2022/11/03         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get LastCacheConfirmationTime()
       LastCacheConfirmationTime = PdbLastCacheConfirmationTime
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get LastCacheUpdateTime()
    'Overview                    : 最終キャッシュ更新時間（Timer関数の値）を返す
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     最終キャッシュ更新時間（Timer関数の値）
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2022/11/03         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get LastCacheUpdateTime()
       LastCacheUpdateTime = PdbLastCacheUpdateTime
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Prop()
    'Overview                    : File/Folderの属性を返す
    'Detailed Description        : 引数で指定した属性の値を返却する
    '                               "Attributes"        属性
    '                               "DateCreated"       作成された日付と時刻
    '                               "DateLastAccessed"  最後にアクセスした日付と時刻
    '                               "DateLastModified"  最終更新日時
    '                               "Drive"             ファイルまたはフォルダーがあるドライブのドライブ文字
    '                               "Name"              名前
    '                               "ParentFolder"      親のフォルダーオブジェクト
    '                               "Path"              パス
    '                               "ShortName"         短い名前(8.3 名前付け規則)
    '                               "ShortPath"         短いパス(8.3 名前付け規則)
    '                               "Size"              サイズ（バイト単位）
    '                               "Type"              種類
    'Argument
    '     asKey                  : 属性を指定するキー
    'Return Value
    '     引数で指定した属性の値
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2022/11/05         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function Prop( _
        byVal asKey _
        )
        If Not(PoProp.Exists(asKey)) Then Exit Function
        
        If func_FsBaseIsGetObjectValue(PoProp.Item(asKey)) Then
            With func_FsBaseGetObject()
                Select Case asKey
                Case "Attributes"                          '属性
                    Call cf_bindAt(PoProp, asKey, .Attributes)
'                    Call sub_CM_TransferToCollection(.Attributes, PoProp, asKey)
                Case "DateCreated"                         '作成された日付と時刻
                    Call cf_bindAt(PoProp, asKey, .DateCreated)
'                    Call sub_CM_TransferToCollection(.DateCreated, PoProp, asKey)
                Case "DateLastAccessed"                    '最後にアクセスした日付と時刻
                    Call cf_bindAt(PoProp, asKey, .DateLastAccessed)
'                    Call sub_CM_TransferToCollection(.DateLastAccessed, PoProp, asKey)
                Case "DateLastModified"                    '最終更新日時
                    '最終更新日は常に設定するため、ここでは何もしない
                Case "Drive"                               'ファイルまたはフォルダーがあるドライブのドライブ文字
                    Call cf_bindAt(PoProp, asKey, .Drive)
'                    Call sub_CM_TransferToCollection(.Drive, PoProp, asKey)
                Case "Name"                                '名前
                    Call cf_bindAt(PoProp, asKey, .Name)
'                    Call sub_CM_TransferToCollection(.Name, PoProp, asKey)
                Case "ParentFolder"                        '親のフォルダーオブジェクト
                    Call cf_bindAt(PoProp, asKey, .ParentFolder)
'                    Call sub_CM_TransferToCollection(.ParentFolder, PoProp, asKey)
                Case "Path"                                'パス
                    Call cf_bindAt(PoProp, asKey, .Path)
'                    Call sub_CM_TransferToCollection(.Path, PoProp, asKey)
                Case "ShortName"                           '短い名前(8.3 名前付け規則)
                    Call cf_bindAt(PoProp, asKey, .ShortName)
'                    Call sub_CM_TransferToCollection(.ShortName, PoProp, asKey)
                Case "ShortPath"                           '短いパス(8.3 名前付け規則)
                    Call cf_bindAt(PoProp, asKey, .ShortPath)
'                    Call sub_CM_TransferToCollection(.ShortPath, PoProp, asKey)
                Case "Size"                                'サイズ（バイト単位）
                    Call cf_bindAt(PoProp, asKey, .Size)
'                    Call sub_CM_TransferToCollection(.Size, PoProp, asKey)
                Case "Type"                                '種類
                    Call cf_bindAt(PoProp, asKey, .Type)
'                    Call sub_CM_TransferToCollection(.Type, PoProp, asKey)
                End Select
                '最終更新日時 と 最終キャッシュ更新時間（Timer関数の値） の更新
                Call cf_bindAt(PoProp, "DateLastModified", .DateLastModified)
'                Call sub_CM_TransferToCollection(.DateLastModified, PoProp, "DateLastModified")
                PdbLastCacheUpdateTime = Timer()
            End With
        End If
        
        '値を返却
        Call cf_bind(Prop, PoProp.Item(asKey))
'        Call sub_CM_TransferBetweenVariables(PoProp.Item(asKey), Prop)
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : func_FsBaseIsGetObjectValue()
    'Overview                    : File/Folderオブジェクトから値を取得するか判断する
    'Detailed Description        : 下記いずれかに該当する場合はオブジェクトを参照する
    '                              ・キャッシュがない（参照する値がvbNullString）
    '                              ・上記以外で、キャッシュを使用しない
    '                              ・上記以外で、有効期間を超過し当該オブジェクトの最終更新日が変わった
    'Argument
    '     avSomeValue            : 参照する値
    'Return Value
    '     結果 True:要 / False:否
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2022/11/03         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function func_FsBaseIsGetObjectValue( _
        byRef avSomeValue _
        )
        func_FsBaseIsGetObjectValue = True
        If avSomeValue = vbNullString Then Exit Function
        If Not(PboUseCache) Then Exit Function
        If Abs(Timer() - PdbLastCacheConfirmationTime) > PdbValidPeriod Then
            PdbLastCacheConfirmationTime = Timer()                   '最終キャッシュ確認時間（Timer関数の値）の更新
            If PoProp.Item("DateLastModified") <> func_FsBaseGetObject().DateLastModified Then Exit Function
        End If
        func_FsBaseIsGetObjectValue = False
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : func_FsBaseGetObject()
    'Overview                    : File/Folderオブジェクトを取得する
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     File/Folderオブジェクト
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2022/11/03         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function func_FsBaseGetObject( _
        )
        Set func_FsBaseGetObject = Nothing
        Dim sPath : sPath = PoProp.Item("Path")
        With func_FsBaseGetFso()
            If .FileExists(sPath) Then Set func_FsBaseGetObject = .GetFile(sPath)
            If .FolderExists(sPath) Then Set func_FsBaseGetObject = .GetFolder(sPath)
        End With
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : func_FsBaseGetFso()
    'Overview                    : FileSystemObjectオブジェクトを取得する
    'Detailed Description        : Nothingだったら作成する
    'Argument
    '     なし
    'Return Value
    '     FileSystemObjectオブジェクト
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2022/11/03         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function func_FsBaseGetFso( _
        )
        If PoFso Is Nothing Then Set PoFso = new_Fso()
        Set func_FsBaseGetFso = PoFso
    End Function

End Class
