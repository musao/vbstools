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
    Private PoFolderItem, PoParent, PsPath, PeEntryType
    
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
        PsPath = vbNullString
        Set PoFolderItem = Nothing
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
        Set PoParent = Nothing
        Set PeEntryType = Nothing
    End Sub
    
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
        allContainers = this_entries(True, PeEntryType("CONTAINER"))
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
        allContainersIncludingSelf = this_allEntriesIncludingSelf(PeEntryType("CONTAINER"))
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
        allEntries = this_entries(True, PeEntryType("ENTRY"))
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
        allEntriesIncludingSelf = this_allEntriesIncludingSelf(PeEntryType("ENTRY"))
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
        allFilesExcludingArchives = this_entries(True, PeEntryType("FILE_EXCLUDING_ARCHIVE"))
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
        allFilesExcludingArchivesIncludingSelf = this_allEntriesIncludingSelf(PeEntryType("FILE_EXCLUDING_ARCHIVE"))
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
        containers = this_entries(False, PeEntryType("CONTAINER"))
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
        entries = this_entries(False, PeEntryType("ENTRY"))
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
        filesExcludingArchives = this_entries(False, PeEntryType("FILE_EXCLUDING_ARCHIVE"))
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
    'Detailed Description        : FolderItem2オブジェクトのIsFolder()ではなく
    '                              FileSystemObjectのFolderExists()と同じ
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
        path = this_path()
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
        toString = "<"&TypeName(Me)&">"&this_path()
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
    'Function/Sub Name           : this_allEntriesIncludingSelf()
    'Overview                    : 自身とフォルダー内のエントリーの配列を返す
    'Detailed Description        : 工事中
    'Argument
    '     alEntryType            : エントリー（ファイル、アーカイブ、フォルダーなど）の種類
    'Return Value
    '     当クラスのインスタンスの配列
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/04/17         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_allEntriesIncludingSelf( _
        byVal alEntryType _
        )
        this_allEntriesIncludingSelf = Null
        If this_isInitial() Then Exit Function

        Dim vRet : vRet = Array()
        Dim boFlg : boFlg = (this_isFolder() Or this_hasEntries(PeEntryType("ENTRY")))
        Select Case alEntryType 
            Case PeEntryType("FILE_EXCLUDING_ARCHIVE")
                If Not boFlg Then vRet=Array(Me)
            Case PeEntryType("CONTAINER")
                If boFlg Then vRet=Array(Me)
            Case Else
                vRet=Array(Me)
        End Select
        
        cf_pushA vRet, this_entries(True, alEntryType)
        this_allEntriesIncludingSelf = vRet
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
        If Not this_isInitial() Then this_baseName = new_Fso().GetBaseName(PsPath)
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
    '     aboRecursiveFlg        : True:再帰処理する / False:再帰処理しない
    '     alEntryType            : エントリー（ファイル、アーカイブ、フォルダーなど）の種類
    'Return Value
    '     当クラスのインスタンスの配列
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/03/20         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_entries( _
        byVal aboRecursiveFlg _
        , byVal alEntryType _
        )
        this_entries = Null
        If this_isInitial() Then Exit Function

        this_entries = Array()
        If Not this_hasEntries(PeEntryType("ENTRY")) Then Exit Function

        If this_isFolder() Then
        'フォルダの場合
            this_entries = this_entriesForFolder(aboRecursiveFlg, alEntryType)
        Else
        'zipの場合
            this_entries = this_entriesForZip(aboRecursiveFlg, alEntryType)
        End If
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : this_entriesForFolder()
    'Overview                    : フォルダー内のエントリーの配列を返す
    'Detailed Description        : フォルダの場合
    'Argument
    '     aboRecursiveFlg        : True:再帰処理する / False:再帰処理しない
    '     alEntryType            : エントリー（ファイル、アーカイブ、フォルダーなど）の種類
    'Return Value
    '     当クラスのインスタンスの配列
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/03/20         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_entriesForFolder( _
        byVal aboRecursiveFlg _
        , byVal alEntryType _
        )
        Dim oEle,vRet()
        With new_FolderOf(PsPath)
            'ファイルの取得
            For Each oEle In .Files
                this_entriesGetEntries vRet,oEle.Path,aboRecursiveFlg,alEntryType
            Next
            
            'フォルダの取得
            If aboRecursiveFlg Or alEntryType<>PeEntryType("FILE_EXCLUDING_ARCHIVE") Then
            '再帰処理するかファイルのみ対象以外フォルダを取得する
                For Each oEle In .SubFolders
                    this_entriesGetEntries vRet,oEle.Path,aboRecursiveFlg,alEntryType
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
    '     aboRecursiveFlg        : True:再帰処理する / False:再帰処理しない
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
        byVal aboRecursiveFlg _
        , byVal alEntryType _
        )
        Dim oEle,vRet()
        For Each oEle In PoFolderItem.GetFolder.Items
            this_entriesGetEntries vRet,oEle.Path,aboRecursiveFlg,alEntryType
        Next

        this_entriesForZip = vRet
        Set oEle = Nothing
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : this_entriesGetEntries()
    'Overview                    : エントリーを取得する
    'Detailed Description        : 再帰処理する場合は下位のエントリーも取得する
    'Argument
    '     avAr                   : 取得したエントリーを格納する配列
    '     asPath                 : パス
    '     aboRecursiveFlg        : True:再帰処理する / False:再帰処理しない
    '     alEntryType            : エントリー（ファイル、アーカイブ、フォルダーなど）の種類
    'Return Value
    '     なし
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/03/23         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub this_entriesGetEntries( _
        byRef avAr _
        , byVal asPath _
        , byVal aboRecursiveFlg _
        , byVal alEntryType _
        )
        Dim oNewItem : Set oNewItem = new_FspOf(asPath).setParent(Me)

        If aboRecursiveFlg Then
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
        If Not this_isInitial() Then this_extension = new_Fso().GetExtensionName(PsPath)
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
                If this_isFolder() Then
                '自身がフォルダの場合
                    If alEntryType=PeEntryType("FILE_EXCLUDING_ARCHIVE") Then
                    '対象がファイルのみの場合
                        this_hasEntries=(new_FolderOf(PsPath).Files.Count>0)
                    Else
                    '対象がフォルダーのみの場合
                        this_hasEntries=(new_FolderOf(PsPath).SubFolders.Count>0)
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
    'Detailed Description        : FolderItem2オブジェクトのIsFolder()ではなく
    '                              FileSystemObjectのFolderExists()と同じ
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
        If Not this_isInitial() Then this_isFolder = new_Fso().FolderExists(PsPath)
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
        If Not this_isInitial() Then this_name = new_Fso().GetFileName(PsPath)
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
        If PoParent Is Nothing Then Set PoParent = new_FspOf(new_Fso().GetParentFolderName(PsPath))
        Set this_parentFolder = PoParent
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : this_path()
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
    '2025/03/15         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_path()
        this_path = Null
        If Not this_isInitial() Then this_path = PsPath
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
        ast_argNothing PoFolderItem , asSource, "Because it is an immutable variable, its value cannot be changed."

        Dim oFolderItem : Set oFolderItem = Nothing
        On Error Resume Next
        Set oFolderItem = new_FolderItem2Of(asPath)
        On Error Goto 0
        ast_argNotNothing oFolderItem , asSource, "invalid argument. " & cf_toString(asPath)

        If oFolderItem Is Nothing Then Exit Sub

        this_setFolderItem oFolderItem, asSource
        this_setPath asPath, asSource
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
        ast_argNotNothing PoFolderItem, asSource, "Please set the value before setting the parent folder."
        ast_argsAreSame TypeName(Me), TypeName(aoParent), asSource, "This is not " & TypeName(Me) &"."
        ast_argsAreSame new_Fso().GetParentFolderName(PsPath), aoParent.path, asSource, "This is not a parent folder."

        Set PoParent = aoParent
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : this_setPath()
    'Overview                    : パスを設定する
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
    Private Sub this_setPath( _
        byVal asPath _
        , byVal asSource _
        )
        PsPath = asPath
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

        If this_isFolder() Then
        'フォルダの場合
            this_size = new_FolderOf(PsPath).Size
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
