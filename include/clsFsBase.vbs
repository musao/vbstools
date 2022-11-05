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
    Private PoFso                              'FileSystemObjectオブジェクト
    Private PboUseCache                        'キャッシュ使用可否（最新を取得するかどうか）
    Private PdbMostRecentReference             'キャッシュ情報取得時間（Timer関数の値）
    Private PdbValidPeriod                     'キャッシュ有効期間（秒数）
    
    Private PlAttributes                       '属性
    Private PdtDateCreated                     '作成された日付と時刻
    Private PdtDateLastAccessed                '最後にアクセスした日付と時刻
    Private PdtDateLastModified                '最終更新日時
    Private PsDrive                            'ファイルまたはフォルダーがあるドライブのドライブ文字
    Private PsName                             '名前
    Private PoParentFolder                     '親のフォルダーオブジェクト
    Private PsPath                             'パス
    Private PsShortName                        '短い名前(8.3 名前付け規則)
    Private PsShortPath                        '短いパス(8.3 名前付け規則)
    Private PlSize                             'サイズ（バイト単位）
    Private PsType                             '種類
    
    'コンストラクタ
    Private Sub Class_Initialize()
        '初期化
        Set PoFso = Nothing
        PboUseCache = True
        PdbMostRecentReference = 0
        PdbValidPeriod = 1
        
        PlAttributes = vbNullString
        PdtDateCreated = vbNullString
        PdtDateLastAccessed = vbNullString
        PdtDateLastModified = vbNullString
        PsDrive = vbNullString
        PsName = vbNullString
        PoParentFolder = vbNullString
        PsPath = vbNullString
        PsShortName = vbNullString
        PsShortPath = vbNullString
        PlSize = vbNullString
        PsType = vbNullString
    End Sub
    'デストラクタ
    Private Sub Class_Terminate()
        Set PoFso = Nothing
        Set PoParentFolder = Nothing
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get Attributes()
    'Overview                    : 属性を返す
    'Detailed Description        : File/Folderオブジェクトの同名プロパティと同様
    'Argument
    '     なし
    'Return Value
    '     属性
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2022/11/03         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get Attributes( _
        )
        If func_FsBaseIsGetObjectValue(PlAttributes) Then
            Dim oObject : Set oObject = func_FsBaseGetObject()
            PlAttributes = oObject.Attributes
            Call sub_FsBaseRecordCacheAcquisition(oObject)
            Set oObject = Nothing
        End If
        Attributes = PlAttributes
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get DateCreated()
    'Overview                    : 作成された日付と時刻を返す
    'Detailed Description        : File/Folderオブジェクトの同名プロパティと同様
    'Argument
    '     なし
    'Return Value
    '     作成された日付と時刻
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2022/11/03         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get DateCreated( _
        )
        If func_FsBaseIsGetObjectValue(PdtDateCreated) Then
            Dim oObject : Set oObject = func_FsBaseGetObject()
            PdtDateCreated = oObject.DateCreated
            Call sub_FsBaseRecordCacheAcquisition(oObject)
            Set oObject = Nothing
        End If
        DateCreated = PdtDateCreated
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get DateLastAccessed()
    'Overview                    : 最後にアクセスした日付と時刻を返す
    'Detailed Description        : File/Folderオブジェクトの同名プロパティと同様
    'Argument
    '     なし
    'Return Value
    '     最後にアクセスした日付と時刻
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2022/11/03         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get DateLastAccessed( _
        )
        If func_FsBaseIsGetObjectValue(PdtDateLastAccessed) Then
            Dim oObject : Set oObject = func_FsBaseGetObject()
            PdtDateLastAccessed = oObject.DateLastAccessed
            Call sub_FsBaseRecordCacheAcquisition(oObject)
            Set oObject = Nothing
        End If
        DateLastAccessed = PdtDateLastAccessed
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get DateLastModified()
    'Overview                    : 最終更新日時を返す
    'Detailed Description        : File/Folderオブジェクトの同名プロパティと同様
    'Argument
    '     なし
    'Return Value
    '     最終更新日時
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2022/11/03         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get DateLastModified( _
        )
        If func_FsBaseIsGetObjectValue(PdtDateLastModified) Then
            Dim oObject : Set oObject = func_FsBaseGetObject()
            Call sub_FsBaseRecordCacheAcquisition(oObject)
            Set oObject = Nothing
        End If
        DateLastModified = PdtDateLastModified
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get Drive()
    'Overview                    : ファイルまたはフォルダーがあるドライブのドライブ文字を返す
    'Detailed Description        : File/Folderオブジェクトの同名プロパティと同様
    'Argument
    '     なし
    'Return Value
    '     ファイルまたはフォルダーがあるドライブのドライブ文字
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2022/11/03         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get Drive( _
        )
        If func_FsBaseIsGetObjectValue(PsDrive) Then
            Dim oObject : Set oObject = func_FsBaseGetObject()
            PsDrive = oObject.Drive
            Call sub_FsBaseRecordCacheAcquisition(oObject)
            Set oObject = Nothing
        End If
        Drive = PsDrive
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get Name()
    'Overview                    : 名前を返す
    'Detailed Description        : File/Folderオブジェクトの同名プロパティと同様
    'Argument
    '     なし
    'Return Value
    '     名前
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2022/11/03         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get Name( _
        )
        If func_FsBaseIsGetObjectValue(PsName) Then
            Dim oObject : Set oObject = func_FsBaseGetObject()
            PsName = oObject.Name
            Call sub_FsBaseRecordCacheAcquisition(oObject)
            Set oObject = Nothing
        End If
        Name = PsName
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get ParentFolder()
    'Overview                    : 親のフォルダーオブジェクトを返す
    'Detailed Description        : File/Folderオブジェクトの同名プロパティと同様
    'Argument
    '     なし
    'Return Value
    '     親のフォルダーオブジェクト
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2022/11/03         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get ParentFolder( _
        )
        If func_FsBaseIsGetObjectValue(PoParentFolder) Then
            Dim oObject : Set oObject = func_FsBaseGetObject()
            Set PoParentFolder = oObject.ParentFolder
            Call sub_FsBaseRecordCacheAcquisition(oObject)
            Set oObject = Nothing
        End If
        Set ParentFolder = PoParentFolder
    End Property
    
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
        PsPath = asPath
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get Path()
    'Overview                    : パスを返す
    'Detailed Description        : File/Folderオブジェクトの同名プロパティと同様
    'Argument
    '     なし
    'Return Value
    '     パス
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2022/11/03         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get Path( _
        )
        If func_FsBaseIsGetObjectValue(PsPath) Then
            Dim oObject : Set oObject = func_FsBaseGetObject()
            PsPath = oObject.Path
            Call sub_FsBaseRecordCacheAcquisition(oObject)
            Set oObject = Nothing
        End If
        Path = PsPath
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get ShortName()
    'Overview                    : 短い名前(8.3 名前付け規則)を返す
    'Detailed Description        : File/Folderオブジェクトの同名プロパティと同様
    'Argument
    '     なし
    'Return Value
    '     短い名前(8.3 名前付け規則)
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2022/11/03         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get ShortName( _
        )
        If func_FsBaseIsGetObjectValue(PsShortName) Then
            Dim oObject : Set oObject = func_FsBaseGetObject()
            PsShortName = oObject.ShortName
            Call sub_FsBaseRecordCacheAcquisition(oObject)
            Set oObject = Nothing
        End If
        ShortName = PsShortName
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get ShortPath()
    'Overview                    : 短いパス(8.3 名前付け規則)を返す
    'Detailed Description        : File/Folderオブジェクトの同名プロパティと同様
    'Argument
    '     なし
    'Return Value
    '     短いパス(8.3 名前付け規則)
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2022/11/03         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get ShortPath( _
        )
        If func_FsBaseIsGetObjectValue(PsShortPath) Then
            Dim oObject : Set oObject = func_FsBaseGetObject()
            PsShortPath = oObject.ShortPath
            Call sub_FsBaseRecordCacheAcquisition(oObject)
            Set oObject = Nothing
        End If
        ShortPath = PsShortPath
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get Size()
    'Overview                    : サイズ（バイト単位）を返す
    'Detailed Description        : File/Folderオブジェクトの同名プロパティと同様
    'Argument
    '     なし
    'Return Value
    '     サイズ（バイト単位）
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2022/11/03         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get Size( _
        )
        If func_FsBaseIsGetObjectValue(PlSize) Then
            Dim oObject : Set oObject = func_FsBaseGetObject()
            PlSize = oObject.Size
            Call sub_FsBaseRecordCacheAcquisition(oObject)
            Set oObject = Nothing
        End If
        Size = PlSize
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get FileFolderType()
    'Overview                    : 種類を返す
    'Detailed Description        : File/Folderオブジェクトの同名プロパティと同様
    'Argument
    '     なし
    'Return Value
    '     種類
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2022/11/03         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get FileFolderType( _
        )
        If func_FsBaseIsGetObjectValue(PsType) Then
            Dim oObject : Set oObject = func_FsBaseGetObject()
            PsType = oObject.Type
            Call sub_FsBaseRecordCacheAcquisition(oObject)
            Set oObject = Nothing
        End If
        FileFolderType = PsType
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
    'Function/Sub Name           : Property Get MostRecentReference()
    'Overview                    : キャッシュ情報取得時間（Timer関数の値）を返す
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     キャッシュ情報取得時間（Timer関数の値）
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2022/11/03         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get MostRecentReference()
       MostRecentReference = PdbMostRecentReference
    End Property
    
    
    
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
        If Abs(Timer() - PdbMostRecentReference) > PdbValidPeriod Then
            If PdtDateLastModified <> func_FsBaseGetObject().DateLastModified Then Exit Function
        End If
        func_FsBaseIsGetObjectValue = False
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : sub_FsBaseRecordCacheAcquisition()
    'Overview                    : キャッシュ取得時の情報を記録する
    'Detailed Description        : 下記を記録する
    '                              ・最終更新日時
    '                              ・キャッシュ情報取得時間（Timer関数の値）
    'Argument
    '     aoSomeObject           : File/Folderオブジェクト
    'Return Value
    '     なし
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2022/11/03         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub sub_FsBaseRecordCacheAcquisition( _
        byRef aoSomeObject _
        )
        With func_FsBaseGetObject()
            PdtDateLastModified = .DateLastModified
            PdbMostRecentReference = Timer()
        End With
    End Sub
    
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
        With func_FsBaseGetFso()
            If .FileExists(PsPath) Then Set func_FsBaseGetObject = .GetFile(PsPath)
            If .FolderExists(PsPath) Then Set func_FsBaseGetObject = .GetFolder(PsPath)
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
        If PoFso Is Nothing Then Set PoFso = CreateObject("Scripting.FileSystemObject")
        Set func_FsBaseGetFso = PoFso
    End Function

End Class
