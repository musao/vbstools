'***************************************************************************************************
'FILENAME                    : FileSystemProxy.vbs
'Overview                    : File/Folderオブジェクトのプロキシクラス
'Detailed Description        : File/Folderオブジェクトと同じIFを提供する
'---------------------------------------------------------------------------------------------------
'History
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/11/26         Y.Fujii                  First edition
'***************************************************************************************************
Class FileSystemProxy
    'クラス内変数、定数
    Private PoFolderItem, PoFolder, PoParent, PsActualPath, PsVirtualPath, PeEntryType, PePathType
    
    '***************************************************************************************************
    'Function/Sub Name           : Class_Initialize()
    'Overview                    : コンストラクタ
    'Detailed Description        : 内部変数の初期化
    'Argument
    '     なし
    'Return Value
    '     なし
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/11/26         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub Class_Initialize()
        PsActualPath = vbNullString
        PsVirtualPath = vbNullString
        Set PoFolderItem = Nothing
        Set PoFolder = Nothing
        Set PoParent = Nothing
        Set PeEntryType = new_DicOf( _
            Array( _
                "ENTRY", 0 _
                , "FILE", 2 _
                , "FILE_EXCLUDING_ARCHIVE", 3 _
                , "FOLDER", 4 _
                , "CONTAINER", 5 _
            ) _
        )
        Set PePathType = new_DicOf( _
            Array( _
                "ACTUAL", 0 _
                , "VIRTUAL", 1 _
            ) _
        )
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
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/11/26         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub Class_Terminate()
        Set PoFolderItem = Nothing
        Set PoFolder = Nothing
        Set PoParent = Nothing
        Set PeEntryType = Nothing
        Set PePathType = Nothing
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get actualPath()
    'Overview                    : 仮想でない実際のフルパスを返す
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     仮想でない実際のフルパス
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2026/01/24         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get actualPath()
        actualPath = this_path(PePathType("ACTUAL"))
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get allContainers()
    'Overview                    : 配下の全てのフォルダーとアーカイブのリストを取得する
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     当クラスのインスタンスの配列
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/04/27         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get allContainers()
        allContainers = this_entries(PeEntryType("CONTAINER"), False, True)
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get allContainersIncludingSelf()
    'Overview                    : 自身を含む配下の全てのフォルダーとアーカイブのリストを取得する
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     当クラスのインスタンスの配列
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/04/25         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get allContainersIncludingSelf()
        allContainersIncludingSelf = this_entries(PeEntryType("CONTAINER"), True, True)
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get allEntries()
    'Overview                    : 配下の全てのエントリー（ファイル、アーカイブ、フォルダー）のリストを取得する
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     当クラスのインスタンスの配列
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/03/20         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get allEntries()
        allEntries = this_entries(PeEntryType("ENTRY"), False, True)
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get allEntriesIncludingSelf()
    'Overview                    : 自身を含む配下の全てのエントリー（ファイル、アーカイブ、フォルダー）のリストを取得する
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     当クラスのインスタンスの配列
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/04/17         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get allEntriesIncludingSelf()
        allEntriesIncludingSelf = this_entries(PeEntryType("ENTRY"), True, True)
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get allFilesExcludingArchives()
    'Overview                    : 配下の全てのアーカイブを除くファイルのリストを取得する
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     当クラスのインスタンスの配列
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/04/27         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get allFilesExcludingArchives()
        allFilesExcludingArchives = this_entries(PeEntryType("FILE_EXCLUDING_ARCHIVE"), False, True)
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get allFilesExcludingArchivesIncludingSelf()
    'Overview                    : 自身を含む配下の全てのアーカイブを除くファイルのリストを取得する
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     当クラスのインスタンスの配列
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/04/25         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get allFilesExcludingArchivesIncludingSelf()
        allFilesExcludingArchivesIncludingSelf = this_entries(PeEntryType("FILE_EXCLUDING_ARCHIVE"), True, True)
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get baseName()
    'Overview                    : 拡張子を除いた名前を返す
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     拡張子を除いた名前
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/03/20         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get baseName()
        baseName = this_baseName()
    End Property

    '***************************************************************************************************
    'Function/Sub Name           : Property Get containers()
    'Overview                    : 配下のコンテナ（フォルダーかアーカイブ）のリストを取得する
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     当クラスのインスタンスの配列
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/04/27         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get containers()
        containers = this_entries(PeEntryType("CONTAINER"), False, False)
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get dateLastModified()
    'Overview                    : 最終更新日時を返す
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     最終更新日時
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/11/26         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get dateLastModified()
        dateLastModified = this_dateLastModified()
    End Property

    '***************************************************************************************************
    'Function/Sub Name           : Property Get entries()
    'Overview                    : 配下のエントリー（ファイル、アーカイブ、フォルダー）のリストを取得する
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     当クラスのインスタンスの配列
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/03/20         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get entries()
        entries = this_entries(PeEntryType("ENTRY"), False, False)
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get extension()
    'Overview                    : 拡張子を返す
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     拡張子
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/03/20         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get extension()
        extension = this_extension()
    End Property

    '***************************************************************************************************
    'Function/Sub Name           : Property Get filesExcludingArchives()
    'Overview                    : 配下のアーカイブを除くファイルのリストを取得する
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     当クラスのインスタンスの配列
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/04/27         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get filesExcludingArchives()
        filesExcludingArchives = this_entries(PeEntryType("FILE_EXCLUDING_ARCHIVE"), False, False)
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get hasContainers()
    'Overview                    : 配下にコンテナ（フォルダーかアーカイブ）を1つ以上持つか返す
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     結果 True:配下にコンテナ（フォルダーかアーカイブ）を1つ以上持つ / False:持たない
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/04/27         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get hasContainers()
        hasContainers = this_hasEntries(PeEntryType("CONTAINER"))
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get hasEntries()
    'Overview                    : 配下にエントリー（ファイル、アーカイブ、フォルダー）を1つ以上持つか返す
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     結果 True:配下にエントリー（ファイル、アーカイブ、フォルダー）を1つ以上持つ / False:持たない
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/03/29         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get hasEntries()
        hasEntries = this_hasEntries(PeEntryType("ENTRY"))
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get hasFilesExcludingArchives()
    'Overview                    : 配下にアーカイブを除くファイルを1つ以上持つか返す
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     結果 True:配下にアーカイブを除くファイルを1つ以上持つ / False:持たない
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/04/26         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get hasFilesExcludingArchives()
        hasFilesExcludingArchives = this_hasEntries(PeEntryType("FILE_EXCLUDING_ARCHIVE"))
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get isArchive()
    'Overview                    : アーカイブかどうかを返す
    'Detailed Description        : ファイルシステムの場合
    '                                Folderでないのにファイルシステムでない場合FolderItemのisFolderがTrueならアーカイブと判断する
    '                              ファイルシステムでない場合
    '                                拡張子がzipの場合アーカイブと判断する
    'Argument
    '     なし
    'Return Value
    '     アーカイブかどうか
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2026/01/25         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get isArchive()
        isArchive = this_isArchive()
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get isBrowsable()
    'Overview                    : ブラウザーまたはWindowsエクスプローラーフレーム内でアイテムをホストできるかどうかを返す
    'Detailed Description        : https://learn.microsoft.com/ja-jp/windows/win32/shell/folderitem
    'Argument
    '     なし
    'Return Value
    '     ブラウザーまたはWindowsエクスプローラーフレーム内でアイテムをホストできるかどうか
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/03/16         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get isBrowsable()
        isBrowsable = this_isBrowsable()
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get isContainer()
    'Overview                    : コンテナ（フォルダーまたはアーカイブ）かどうかを返す
    'Detailed Description        : FolderItemのisFolderがTrueまたは拡張子がzipの場合、
    '                              コンテナ（フォルダーまたはアーカイブ）と判断する
    'Argument
    '     なし
    'Return Value
    '     コンテナかどうか
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2026/01/25         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get isContainer()
        isContainer = this_isContainer()
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get isFile()
    'Overview                    : ファイル（アーカイブを含む）かどうかを返す
    'Detailed Description        : FolderItemのisFolderがFalseの場合、ファイル（アーカイブを含む）と判断する
    'Argument
    '     なし
    'Return Value
    '     ファイル（アーカイブを含む）かどうか
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2026/01/25         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get isFile()
        isFile = this_isFile()
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get isFileSystem()
    'Overview                    : 項目がファイルシステムの一部であるかどうかを返す
    'Detailed Description        : https://learn.microsoft.com/ja-jp/windows/win32/shell/folderitem
    'Argument
    '     なし
    'Return Value
    '     項目がファイルシステムの一部であるかどうか
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/03/16         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get isFileSystem()
        isFileSystem = this_isFileSystem()
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get isFolder()
    'Overview                    : フォルダであるかどうかを返す
    'Detailed Description        : ファイルシステムの場合はFileSystemObjectのFolderExists()と同じ
    '                              ファイルシステムでない場合FolderItemのisFolder()と同じ
    'Argument
    '     なし
    'Return Value
    '     アイテムがフォルダであるかどうか
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/03/16         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get isFolder()
        isFolder = this_isFolder()
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get isLink()
    'Overview                    : 項目がショートカットであるかどうかを返す
    'Detailed Description        : https://learn.microsoft.com/ja-jp/windows/win32/shell/folderitem
    'Argument
    '     なし
    'Return Value
    '     項目がショートカットであるかどうか
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/03/16         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get isLink()
        isLink = this_isLink()
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get name()
    'Overview                    : 名前を返す
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     名前
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/11/26         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get name()
        name = this_name()
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get parentFolder()
    'Overview                    : 親フォルダを返す
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     親フォルダのインスタンス
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/11/26         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get parentFolder()
        cf_bind parentFolder, this_parentFolder()
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get path()
    'Overview                    : フルパスを返す
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     フルパス
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/11/26         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Default Property Get path()
        path = this_path(PePathType("VIRTUAL"))
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get size()
    'Overview                    : サイズをバイト単位で返す
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     バイト単位のサイズ
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/11/26         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get size()
        size = this_size()
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get toString()
    'Overview                    : 当インスタンスの内容を文字列に変換する
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     文字列
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/03/16         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get toString()
        toString = "<"&TypeName(Me)&">"&this_path(PePathType("VIRTUAL"))
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get type()
    'Overview                    : 種類を返す
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     種類
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/11/26         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get [type]()
        [type] = this_type()
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : of()
    'Overview                    : エントリー（ファイル、アーカイブ、フォルダー）を設定する
    'Detailed Description        : 工事中
    'Argument
    '     asPath                 : エントリー（ファイル、アーカイブ、フォルダー）のパス
    'Return Value
    '     自身のインスタンス
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/12/19         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function of( _
        byVal asPath _
        )
        this_setData asPath, TypeName(Me)&"+of()"
        Set of = Me
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : setParent()
    'Overview                    : 親フォルダのインスタンスを設定する
    'Detailed Description        : 工事中
    'Argument
    '     aoParent               : 親フォルダのインスタンス
    'Return Value
    '     自身のインスタンス
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/03/23         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function setParent( _
        byRef aoParent _
        )
        this_setParent aoParent, TypeName(Me)&"+setParent()"
        Set setParent = Me
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : setVirtualPath()
    'Overview                    : 仮想パスを設定する
    'Detailed Description        : 工事中
    'Argument
    '     asVirtualPath          : 仮想パス
    'Return Value
    '     自身のインスタンス
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2026/01/24         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function setVirtualPath( _
        byVal asVirtualPath _
        )
        this_setVirtualPath asVirtualPath, TypeName(Me)&"+setVirtualPath()"
        Set setVirtualPath = Me
    End Function



    
    '***************************************************************************************************
    'Function/Sub Name           : this_baseName()
    'Overview                    : 拡張子を除いた名前を返す
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     拡張子を除いた名前
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/03/20         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_baseName()
        this_baseName = Null
        If Not this_isInitial() Then this_baseName = new_Fso().GetBaseName(PsVirtualPath)
    End Function
   
    '***************************************************************************************************
    'Function/Sub Name           : this_dateLastModified()
    'Overview                    : 最終更新日時を返す
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     最終更新日時
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/03/15         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_dateLastModified()
        this_dateLastModified = Null
        If Not this_isInitial() Then this_dateLastModified = PoFolderItem.ModifyDate
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : this_entries()
    'Overview                    : フォルダー内のエントリーの配列を返す
    'Detailed Description        : 工事中
    'Argument
    '     alEntryType            : エントリー（ファイル、アーカイブ、フォルダーなど）の種類
    '     aboIncludingSelf       : True:自身を含める / False:自身を含めない
    '     aboRecursive           : True:再帰処理する / False:再帰処理しない
    'Return Value
    '     当クラスのインスタンスの配列
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/03/20         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_entries( _
        byVal alEntryType _
        , byVal aboIncludingSelf _
        , byVal aboRecursive _
        )
        this_entries = Null
        If this_isInitial() Then Exit Function
        
        Dim vRet, boHasEntries
        vRet = Array()
        boHasEntries = this_hasEntries(PeEntryType("ENTRY"))

        If aboIncludingSelf Then
            Dim boFlg : boFlg = (this_existsFolder() Or boHasEntries)
            Select Case alEntryType 
                Case PeEntryType("FILE_EXCLUDING_ARCHIVE")
                    If Not boFlg Then vRet=Array(Me)
                Case PeEntryType("CONTAINER")
                    If boFlg Then vRet=Array(Me)
                Case Else
                    vRet=Array(Me)
            End Select
        End If
        this_entries = vRet
        If Not boHasEntries Then Exit Function

        If this_existsFolder() Then
        'フォルダの場合
            pushA vRet, this_entriesForFolder(alEntryType, aboRecursive)
        Else
        'zipの場合
            pushA vRet, this_entriesForZip(alEntryType, aboRecursive)
        End If
        this_entries = vRet
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : this_entriesForFolder()
    'Overview                    : フォルダー内のエントリーの配列を返す
    'Detailed Description        : フォルダの場合
    'Argument
    '     alEntryType            : エントリー（ファイル、アーカイブ、フォルダーなど）の種類
    '     aboRecursive           : True:再帰処理する / False:再帰処理しない
    'Return Value
    '     当クラスのインスタンスの配列
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/03/20         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_entriesForFolder( _
        byVal alEntryType _
        , byVal aboRecursive _
        )
        Dim oEle,vRet()
        With PoFolder
            'ファイルの取得
            For Each oEle In .Files
                this_entriesGetEntries alEntryType, aboRecursive, oEle.Path, vRet
            Next
            
            'フォルダの取得
            If aboRecursive Or alEntryType<>PeEntryType("FILE_EXCLUDING_ARCHIVE") Then
            '再帰処理するかファイルのみ対象以外フォルダを取得する
                For Each oEle In .SubFolders
                    this_entriesGetEntries alEntryType, aboRecursive, oEle.Path, vRet
                Next
            End If
        End With

        this_entriesForFolder = vRet
        Set oEle = Nothing
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : this_entriesForZip()
    'Overview                    : アーカイブ内のエントリーの配列を返す
    'Detailed Description        : zipの場合
    'Argument
    '     aboRecursive           : True:再帰処理する / False:再帰処理しない
    '     alEntryType            : エントリー（ファイル、アーカイブ、フォルダーなど）の種類
    'Return Value
    '     当クラスのインスタンスの配列
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/03/20         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_entriesForZip( _
        byVal alEntryType _
        , byVal aboRecursive _
        )
        Dim oEle,vRet()
        For Each oEle In PoFolderItem.GetFolder.Items
            this_entriesGetEntries alEntryType, aboRecursive, oEle.Path, vRet
        Next

        this_entriesForZip = vRet
        Set oEle = Nothing
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : this_entriesGetEntries()
    'Overview                    : エントリーを取得する
    'Detailed Description        : 再帰処理する場合は下位のエントリーも取得する
    'Argument
    '     alEntryType            : エントリー（ファイル、アーカイブ、フォルダーなど）の種類
    '     aboRecursive           : True:再帰処理する / False:再帰処理しない
    '     asPath                 : パス
    '     avAr                   : 取得したエントリーを格納する配列
    'Return Value
    '     なし
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/03/23         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub this_entriesGetEntries( _
        byVal alEntryType _
        , byVal aboRecursive _
        , byVal asPath _
        , byRef avAr _
        )
        Dim oNewItem : Set oNewItem = new_FspOf(asPath).setParent(Me)

        If aboRecursive Then
        '再帰処理する場合
            Select Case alEntryType
                Case PeEntryType("FILE_EXCLUDING_ARCHIVE")
                    cf_pushA avAr, oNewItem.allFilesExcludingArchivesIncludingSelf()
                Case PeEntryType("CONTAINER")
                    cf_pushA avAr, oNewItem.allContainersIncludingSelf()
                Case Else
                    cf_pushA avAr, oNewItem.allEntriesIncludingSelf()
            End Select
        Else
        '再帰処理しない場合
            Select Case alEntryType
                Case PeEntryType("FILE_EXCLUDING_ARCHIVE"),PeEntryType("CONTAINER")
                    Dim boFlg : If alEntryType=PeEntryType("FILE_EXCLUDING_ARCHIVE") Then boFlg=False Else boFlg=True
                    If (oNewItem.isFolder Or oNewItem.hasEntries)=boFlg Then cf_push avAr, oNewItem
                Case Else
                    cf_push avAr, oNewItem
            End Select
        End If

        Set oNewItem = Nothing
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : this_existsFolder()
    'Overview                    : 自身がファイルシステム上のフォルダであるかどうかを返す
    'Detailed Description        : FileSystemObjectのFolderExists()と同じ
    'Argument
    '     なし
    'Return Value
    '     自身がファイルシステム上のフォルダであるかどうか
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2026/01/25         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_existsFolder()
        this_existsFolder = Null
        If Not this_isInitial() Then this_existsFolder = new_Fso().FolderExists(PsActualPath)
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : this_extension()
    'Overview                    : 拡張子を返す
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     拡張子
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/03/20         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_extension()
        this_extension = Null
        If Not this_isInitial() Then this_extension = new_Fso().GetExtensionName(PsVirtualPath)
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : this_hasEntries()
    'Overview                    : 配下に引数で指定したエントリー（ファイル、アーカイブ、フォルダー）を1つ以上持つか返す
    'Detailed Description        : 工事中
    'Argument
    '     alEntryType            : エントリー（ファイル、アーカイブ、フォルダーなど）の種類
    'Return Value
    '     結果 True:配引数で指定したエントリーを1つ以上持つ / False:持たない
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/03/29         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_hasEntries( _
        byVal alEntryType _
        )
        this_hasEntries = Null
        If this_isInitial() Then Exit Function

        this_hasEntries = False
        Select Case alEntryType 
            Case PeEntryType("FILE_EXCLUDING_ARCHIVE"),PeEntryType("CONTAINER")
            '対象がファイルのみかフォルダーのみの場合
                If this_existsFolder() Then
                '自身がフォルダの場合
                    If alEntryType=PeEntryType("FILE_EXCLUDING_ARCHIVE") Then
                    '対象がファイルのみの場合
                        this_hasEntries=(PoFolder.Files.Count>0)
                    Else
                    '対象がフォルダーのみの場合
                        this_hasEntries=(PoFolder.SubFolders.Count>0)
                    End If
                ElseIf PoFolderItem.IsFolder Then
                '自身がzipの場合
                    Dim oEle,boFlg
                    If alEntryType=PeEntryType("FILE_EXCLUDING_ARCHIVE") Then boFlg=False Else boFlg=True
                    For Each oEle In PoFolderItem.GetFolder.Items
                        If oEle.IsFolder=boFlg Then
                            this_hasEntries=True
                            Exit For
                        End If
                    Next
                End If
            Case Else
            '上記以外の場合
                If PoFolderItem.IsFolder Then this_hasEntries=(PoFolderItem.GetFolder.Items.Count>0)
        End Select
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : this_isArchive()
    'Overview                    : アーカイブかどうかを返す
    'Detailed Description        : ファイルシステムの場合
    '                                Folderでないのにファイルシステムでない場合FolderItemのisFolderがTrueならアーカイブと判断する
    '                              ファイルシステムでない場合
    '                                拡張子がzipの場合アーカイブと判断する
    'Argument
    '     なし
    'Return Value
    '     アーカイブかどうか
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2026/01/25         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_isArchive()
        this_isArchive = Null
        If this_isInitial() Then Exit Function
        If this_isFileSystem() Then
        'ファイルシステムの場合、FolderでなくFolderItemのisFolderがTrueならアーカイブと判断する
            this_isArchive = False
            If Not this_existsFolder() And PoFolderItem.IsFolder Then this_isArchive = True
        Else
        'ファイルシステムでない場合、拡張子で判断する
            this_isArchive = cf_isSame(LCase(this_extension), "zip")
        End If
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : this_isBrowsable()
    'Overview                    : ブラウザーまたはWindowsエクスプローラーフレーム内でアイテムをホストできるかどうかを返す
    'Detailed Description        : https://learn.microsoft.com/ja-jp/windows/win32/shell/folderitem
    'Argument
    '     なし
    'Return Value
    '     ブラウザーまたはWindowsエクスプローラーフレーム内でアイテムをホストできるかどうか
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/03/16         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_isBrowsable()
        this_isBrowsable = Null
        If Not this_isInitial() Then this_isBrowsable = PoFolderItem.IsBrowsable
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : this_isContainer()
    'Overview                    : コンテナ（フォルダーまたはアーカイブ）かどうかを返す
    'Detailed Description        : FolderItemのisFolderがTrueまたは拡張子がzipの場合、
    '                              コンテナ（フォルダーまたはアーカイブ）と判断する
    'Argument
    '     なし
    'Return Value
    '     コンテナかどうか
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2026/01/25         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_isContainer()
        this_isContainer = Null
        If this_isInitial() Then Exit Function
        this_isContainer = PoFolderItem.IsFolder Or cf_isSame(LCase(this_extension), "zip")
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : this_isFile()
    'Overview                    : ファイル（アーカイブを含む）かどうかを返す
    'Detailed Description        : FolderItemのisFolderがFalseの場合、ファイル（アーカイブを含む）と判断する
    'Argument
    '     なし
    'Return Value
    '     ファイル（アーカイブを含む）かどうか
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2026/01/25         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_isFile()
        this_isFile = Null
        If this_isInitial() Then Exit Function
        this_isFile = Not PoFolderItem.IsFolder
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : this_isFileSystem()
    'Overview                    : 項目がファイルシステムの一部であるかどうかを返す
    'Detailed Description        : https://learn.microsoft.com/ja-jp/windows/win32/shell/folderitem
    'Argument
    '     なし
    'Return Value
    '     項目がファイルシステムの一部であるかどうか
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/03/16         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_isFileSystem()
        this_isFileSystem = Null
        If Not this_isInitial() Then this_isFileSystem = PoFolderItem.IsFileSystem
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : this_isFolder()
    'Overview                    : アイテムがフォルダであるかどうかを返す
    'Detailed Description        : ファイルシステムの場合はFileSystemObjectのFolderExists()と同じ
    '                              ファイルシステムでない場合FolderItemのisFolder()と同じ
    'Argument
    '     なし
    'Return Value
    '     アイテムがフォルダであるかどうか
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/03/16         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_isFolder()
        this_isFolder = Null
        If this_isInitial() Then Exit Function
        If this_isFileSystem() Then
        'ファイルシステムの場合
            this_isFolder = this_existsFolder()
        Else
        'ファイルシステムでない場合
            this_isFolder = PoFolderItem.IsFolder
        End If
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : this_isInitial()
    'Overview                    : 当インスタンスが初期状態か返す
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     結果 True:初期状態 / False:初期状態でない
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/12/06         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_isInitial()
        this_isInitial = (PoFolderItem Is Nothing)
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : this_isLink()
    'Overview                    : 項目がショートカットであるかどうかを返す
    'Detailed Description        : https://learn.microsoft.com/ja-jp/windows/win32/shell/folderitem
    'Argument
    '     なし
    'Return Value
    '     項目がショートカットであるかどうか
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/03/16         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_isLink()
        this_isLink = Null
        If Not this_isInitial() Then this_isLink = PoFolderItem.IsLink
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : this_name()
    'Overview                    : 名前を返す
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     名前
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/03/15         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_name()
        this_name = Null
        If Not this_isInitial() Then this_name = new_Fso().GetFileName(PsVirtualPath)
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : this_parentFolder()
    'Overview                    : 親フォルダのフルパスを返す
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     親フォルダのインスタンス
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/03/15         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_parentFolder()
        this_parentFolder = Null
        If this_isInitial() Then Exit Function
        If PoParent Is Nothing Then Set PoParent = new_FspOf(new_Fso().GetParentFolderName(PsActualPath))
        Set this_parentFolder = PoParent
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : this_path()
    'Overview                    : フルパスを返す
    'Detailed Description        : 工事中
    'Argument
    '     asPathType             : パスの種類（実パス、仮想パス）
    'Return Value
    '     フルパス
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/03/15         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_path( _
        byVal asPathType _
        )
        this_path = Null
        If this_isInitial() Then Exit Function
        
        If asPathType=PePathType("ACTUAL") Then
            this_path = PsActualPath
        Else
            this_path = PsVirtualPath
        End If
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : this_setData()
    'Overview                    : データを設定する
    'Detailed Description        : 工事中
    'Argument
    '     asPath                 : フルパス
    '     asSource               : ソース
    'Return Value
    '     なし
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/03/16         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub this_setData( _
        byVal asPath _
        , byVal asSource _
        )
        ast_argTrue this_isInitial() , asSource, "Because it is an immutable variable, its value cannot be changed."

        Dim oFolderItem : Set oFolderItem = Nothing
        With fw_try(Getref("new_FolderItem2Of"), asPath)
            If .isErr() Then
                ast_argNotNothing oFolderItem , asSource, "invalid argument. " & cf_toString(asPath)
            Else
                Set oFolderItem = .returnValue
            End If
        End With

        this_setFolderItem oFolderItem, asSource
        this_setPath asPath
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : this_setFolderItem()
    'Overview                    : オブジェクト（FolderItem2）を設定する
    'Detailed Description        : 工事中
    'Argument
    '     aoFolderItem           : オブジェクト
    '     asSource               : ソース
    'Return Value
    '     なし
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/03/16         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub this_setFolderItem( _
        byRef aoFolderItem _
        , byVal asSource _
        )
        ast_argsAreSame "FolderItem2", TypeName(aoFolderItem), asSource, "This is not FolderItem2."
        Set PoFolderItem = aoFolderItem
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : this_setParent()
    'Overview                    : 親フォルダを設定する
    'Detailed Description        : 工事中
    'Argument
    '     aoParent               : 親フォルダのインスタンス
    '     asSource               : ソース
    'Return Value
    '     なし
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/03/23         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub this_setParent( _
        byRef aoParent _
        , byVal asSource _
        )
        ast_argFalse this_isInitial() , asSource, "Please set the value before setting the parent folder."
        ast_argsAreSame TypeName(Me), TypeName(aoParent), asSource, "This is not " & TypeName(Me) &"."
        ast_argsAreSame new_Fso().GetParentFolderName(PsActualPath), aoParent.path, asSource, "This is not a parent folder."

        Set PoParent = aoParent
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : this_setPath()
    'Overview                    : パスを設定する
    'Detailed Description        : 仮想パスにも同じパスを設定する
    'Argument
    '     asPath                 : フルパス
    'Return Value
    '     なし
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/03/16         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub this_setPath( _
        byVal asPath _
        )
        PsActualPath = asPath
        PsVirtualPath = asPath
        If this_existsFolder() Then Set PoFolder = new_FolderOf(PsActualPath)
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : this_setVirtualPath()
    'Overview                    : 仮想パスを設定する
    'Detailed Description        : 仮想パスが空文字の場合は実パスを設定する
    'Argument
    '     asVirtualPath          : 仮想パス
    '     asSource               : ソース
    'Return Value
    '     なし
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2026/01/24         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub this_setVirtualPath( _
        byVal asVirtualPath _
        , byVal asSource _
        )
        ast_argFalse this_isInitial() , asSource, "Please set the value before setting the virtual path."
        If asVirtualPath="" Then
            PsVirtualPath = PsActualPath
        Else
            PsVirtualPath = asVirtualPath
        End If
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : this_size()
    'Overview                    : サイズを返す
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     バイト単位のサイズ
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/03/15         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_size()
        this_size = Null
        If this_isInitial() Then Exit Function

        If this_existsFolder() Then
        'フォルダの場合
            this_size = PoFolder.Size
        Else
        'フォルダ以外の場合
            this_size = PoFolderItem.Size
        End If
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : this_type()
    'Overview                    : 種類を返す
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     種類
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/03/15         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_type()
        this_type = Null
        If Not this_isInitial() Then this_type = PoFolderItem.Type
    End Function

End Class
