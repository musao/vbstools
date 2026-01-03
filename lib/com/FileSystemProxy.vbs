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
    Private PoFolderItem, PoParent, PsPath, PeItemType
    
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
        Set PeItemType = new_DicOf( _
            Array( _
                "ALL", 0 _
                , "FILE", 1 _
                , "FILE_EXCLUDE_ARCHIVE", 3 _
                , "FOLDER", 4 _
                , "FOLDER_INCLUDE_ARCHIVE", 5 _
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
        Set PeItemType = Nothing
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get allFiles()
    'Overview                    : フォルダー以下の全てのファイルの配列を返す
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
    Public Property Get allFiles()
        allFiles = this_items(True, PeItemType("FILE_EXCLUDE_ARCHIVE"))
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get allFolders()
    'Overview                    : フォルダー以下の全てのファイルの配列を返す
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
    Public Property Get allFolders()
        allFolders = this_items(True, PeItemType("FOLDER_INCLUDE_ARCHIVE"))
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get allItems()
    'Overview                    : フォルダー以下の全てのアイテムの配列を返す
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
    Public Property Get allItems()
        allItems = this_items(True, PeItemType("ALL"))
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get baseName()
    'Overview                    : ファイル／フォルダの拡張子を除いた名前を返す
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     ファイル／フォルダの拡張子を除いた名前
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
    'Function/Sub Name           : Property Get dateLastModified()
    'Overview                    : ファイル／フォルダの最終更新日時を返す
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     ファイル／フォルダの最終更新日時
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
    'Function/Sub Name           : Property Get extension()
    'Overview                    : ファイル／フォルダの拡張子を返す
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     ファイル／フォルダの拡張子
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
    'Function/Sub Name           : Property Get files()
    'Overview                    : フォルダー内のファイルの配列を返す
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
    Public Property Get files()
        files = this_items(False, PeItemType("FILE_EXCLUDE_ARCHIVE"))
    End Property

    '***************************************************************************************************
    'Function/Sub Name           : Property Get folders()
    'Overview                    : フォルダー内のフォルダーの配列を返す
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
    Public Property Get folders()
        folders = this_items(False, PeItemType("FOLDER_INCLUDE_ARCHIVE"))
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get hasFile()
    'Overview                    : 配下にファイルを1つ以上持つか返す
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     配下にファイルを1つ以上持つか
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/04/26         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get hasFile()
        hasFile = this_hasItem(PeItemType("FILE_EXCLUDE_ARCHIVE"))
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get hasFolder()
    'Overview                    : 配下にファイルを1つ以上持つか返す
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     配下にファイルを1つ以上持つか
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/04/27         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get hasFolder()
        hasFolder = this_hasItem(PeItemType("FOLDER_INCLUDE_ARCHIVE"))
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get hasItem()
    'Overview                    : 配下にアイテムを1つ以上持つか返す
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     配下にアイテムを1つ以上持つか
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/03/29         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get hasItem()
        hasItem = this_hasItem(PeItemType("ALL"))
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
    'Function/Sub Name           : Property Get items()
    'Overview                    : フォルダー内のアイテムの配列を返す
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
    Public Property Get items()
        items = this_items(False, PeItemType("ALL"))
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get name()
    'Overview                    : ファイル／フォルダの名前を返す
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     ファイル／フォルダの名前
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
    'Overview                    : ファイル／フォルダの親フォルダのフルパスを返す
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     ファイル／フォルダの親フォルダの当クラスのインスタンス
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
    'Overview                    : ファイル／フォルダのフルパスを返す
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     ファイル／フォルダのフルパス
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
    'Function/Sub Name           : Property Get selfAndAllFiles()
    'Overview                    : 自身と配下の全てファイルの配列を返す
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
    Public Property Get selfAndAllFiles()
        selfAndAllFiles = this_selfAndAllItems(PeItemType("FILE_EXCLUDE_ARCHIVE"))
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get selfAndAllFolders()
    'Overview                    : 自身と配下の全てフォルダーの配列を返す
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
    Public Property Get selfAndAllFolders()
        selfAndAllFolders = this_selfAndAllItems(PeItemType("FOLDER_INCLUDE_ARCHIVE"))
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get selfAndAllItems()
    'Overview                    : 自身と配下の全てアイテムの配列を返す
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
    Public Property Get selfAndAllItems()
        selfAndAllItems = this_selfAndAllItems(PeItemType("ALL"))
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get size()
    'Overview                    : ファイル／フォルダのサイズをバイト単位で返す
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     ファイル／フォルダのサイズ（バイト単位）
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
    'Overview                    : オブジェクトを文字列に変換する
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     変換した文字列
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
    'Overview                    : ファイル／フォルダの種類を返す
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     ファイル／フォルダの種類
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
    'Overview                    : ファイル／フォルダのパスを設定する
    'Detailed Description        : 工事中
    'Argument
    '     asPath                 : ファイル／フォルダのパス
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
    'Function/Sub Name           : this_baseName()
    'Overview                    : ファイル／フォルダの拡張子を除いた名前を返す
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     ファイル／フォルダの拡張子を除いた名前
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
    'Overview                    : ファイル／フォルダの最終更新日時を返す
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     ファイル／フォルダの最終更新日時
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
    'Function/Sub Name           : this_extension()
    'Overview                    : ファイル／フォルダの拡張子を返す
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     ファイル／フォルダの拡張子
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
    'Function/Sub Name           : this_hasItem()
    'Overview                    : 配下に引数で指定したアイテムを1つ以上持つか返す
    'Detailed Description        : 工事中
    'Argument
    '     alItemType             : Cl_FILE:ファイルのみ / Cl_FOLDER:フォルダーのみ / 左記以外:全て
    'Return Value
    '     配下にアイテムを1つ以上持つか
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/03/29         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_hasItem( _
        byVal alItemType _
        )
        this_hasItem = Null
        If this_isInitial() Then Exit Function

        this_hasItem = False
        Select Case alItemType  
            Case PeItemType("FILE_EXCLUDE_ARCHIVE"),PeItemType("FOLDER_INCLUDE_ARCHIVE")
            '対象がファイルのみかフォルダーのみの場合
                If new_Fso().FolderExists(PsPath) Then
                '自身がフォルダの場合
                    If alItemType=PeItemType("FILE_EXCLUDE_ARCHIVE") Then
                    '対象がファイルのみの場合
                        this_hasItem=(new_FolderOf(PsPath).Files.Count>0)
                    Else
                    '対象がフォルダーのみの場合
                        this_hasItem=(new_FolderOf(PsPath).SubFolders.Count>0)
                    End If
                ElseIf PoFolderItem.IsFolder Then
                '自身がzipの場合
                    Dim oEle,boFlg
                    If alItemType=PeItemType("FILE_EXCLUDE_ARCHIVE") Then boFlg=False Else boFlg=True
                    For Each oEle In PoFolderItem.GetFolder.Items
                        If oEle.IsFolder=boFlg Then
                            this_hasItem=True
                            Exit For
                        End If
                    Next
                End If
            Case Else
            '上記以外の場合
                If PoFolderItem.IsFolder Then this_hasItem=(PoFolderItem.GetFolder.Items.Count>0)
        End Select
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : this_items()
    'Overview                    : フォルダー内のアイテムの配列を返す
    'Detailed Description        : 工事中
    'Argument
    '     aboRecursiveFlg        : True:再帰処理する / False:再帰処理しない
    '     alItemType             : Cl_FILE:ファイルのみ / Cl_FOLDER:フォルダーのみ / 左記以外:全て
    'Return Value
    '     当クラスのインスタンスの配列
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/03/20         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_items( _
        byVal aboRecursiveFlg _
        , byVal alItemType _
        )
        this_items = Null
        If this_isInitial() Then Exit Function

        this_items = Array()
        If Not this_hasItem(PeItemType("ALL")) Then Exit Function

        If new_Fso().FolderExists(PsPath) Then
        'フォルダの場合
            this_items = this_itemsForFolder(aboRecursiveFlg, alItemType)
'            this_items = this_itemsByDir(aboRecursiveFlg, alItemType)
        Else
        'zipの場合
            this_items = this_itemsForZip(aboRecursiveFlg, alItemType)
        End If
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : this_itemsByDir()
    'Overview                    : フォルダー内のアイテムの配列を返す
    'Detailed Description        : cmdのdir版
    'Argument
    '     aboRecursiveFlg        : True:再帰処理する / False:再帰処理しない
    '     alItemType             : Cl_FILE:ファイルのみ / Cl_FOLDER:フォルダーのみ / 左記以外:全て
    'Return Value
    '     当クラスのインスタンスの配列
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/04/26         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_itemsByDir( _
        byVal aboRecursiveFlg _
        , byVal alItemType _
        )
        Dim sFlg,sDir
        sFlg="" : If aboRecursiveFlg Then sFlg="/S "
        sDir = "dir /B " & sFlg & fs_wrapInQuotes(PsPath)
        Dim sTmpPath : sTmpPath = fw_getTempPath()
        
        fw_runShellSilently "cmd /U /C " & sDir & " > " & fs_wrapInQuotes(sTmpPath)
        Dim vArrList : vArrList = Split(fs_readFile(sTmpPath), vbNewLine)
        fs_deleteFile sTmpPath
        Redim Preserve vArrList(Ubound(vArrList)-1)

        Dim oParents : Set oParents = new_DicOf(Array(PsPath,Me))
        Dim boFlg : If alItemType=PeItemType("FILE_EXCLUDE_ARCHIVE") Then boFlg=False Else boFlg=True
        Dim sPath,sEle,sParentPath,oNewItem,vRet()
        For Each sEle In vArrList
            If aboRecursiveFlg Then sPath=sEle Else sPath=new_Fso().BuildPath(PsPath,sEle)
            Set oNewItem = new_FspOf(sPath)
            sParentPath = new_Fso().GetParentFolderName(sPath)
            If oParents.Exists(sParentPath) Then oNewItemoFsp.setParent oParents(sParentPath)

            Select Case alItemType
                Case PeItemType("FILE_EXCLUDE_ARCHIVE"),PeItemType("FOLDER_INCLUDE_ARCHIVE")
                    If (oNewItem.isFolder Or oNewItem.hasItem)=boFlg Then cf_push vRet, oNewItem
                Case Else
                    cf_push vRet, oNewItem
            End Select

            If aboRecursiveFlg And oNewItem.isFolder And Not oParents.Exists(sPath) Then oParents.Add sPath, oNewItem
        Next

        this_itemsByDir = vRet
        Set oFsp = Nothing
        Set oParents = Nothing
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : this_itemsForFolder()
    'Overview                    : フォルダー内のアイテムの配列を返す
    'Detailed Description        : フォルダの場合
    'Argument
    '     aboRecursiveFlg        : True:再帰処理する / False:再帰処理しない
    '     alItemType             : Cl_FILE:ファイルのみ / Cl_FOLDER:フォルダーのみ / 左記以外:全て
    'Return Value
    '     当クラスのインスタンスの配列
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/03/20         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_itemsForFolder( _
        byVal aboRecursiveFlg _
        , byVal alItemType _
        )
        Dim oEle,vRet()
        With new_FolderOf(PsPath)
            'ファイルの取得
            For Each oEle In .Files
                this_itemsGetItems vRet,oEle.Path,aboRecursiveFlg,alItemType
            Next
            
            'フォルダの取得
            If aboRecursiveFlg Or alItemType<>PeItemType("FILE_EXCLUDE_ARCHIVE") Then
            '再帰処理するかファイルのみ対象以外フォルダを取得する
                For Each oEle In .SubFolders
                    this_itemsGetItems vRet,oEle.Path,aboRecursiveFlg,alItemType
                Next
            End If
        End With

        this_itemsForFolder = vRet
        Set oEle = Nothing
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : this_itemsForZip()
    'Overview                    : フォルダー内のアイテムの配列を返す
    'Detailed Description        : zipの場合
    'Argument
    '     aboRecursiveFlg        : True:再帰処理する / False:再帰処理しない
    '     alItemType             : Cl_FILE:ファイルのみ / Cl_FOLDER:フォルダーのみ / 左記以外:全て
    'Return Value
    '     当クラスのインスタンスの配列
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/03/20         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_itemsForZip( _
        byVal aboRecursiveFlg _
        , byVal alItemType _
        )
        Dim oEle,vRet()
        For Each oEle In PoFolderItem.GetFolder.Items
            this_itemsGetItems vRet,oEle.Path,aboRecursiveFlg,alItemType
        Next

        this_itemsForZip = vRet
        Set oEle = Nothing
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : this_itemsGetItems()
    'Overview                    : アイテムを取得する
    'Detailed Description        : 再帰処理する場合は下位のアイテムも取得する
    'Argument
    '     avAr                   : 取得したアイテムを格納する配列
    '     asPath                 : パス
    '     aboRecursiveFlg        : True:再帰処理する / False:再帰処理しない
    '     alItemType             : Cl_FILE:ファイルのみ / Cl_FOLDER:フォルダーのみ / 左記以外:全て
    'Return Value
    '     なし
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/03/23         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub this_itemsGetItems( _
        byRef avAr _
        , byVal asPath _
        , byVal aboRecursiveFlg _
        , byVal alItemType _
        )
        Dim oNewItem : Set oNewItem = new_FspOf(asPath).setParent(Me)

        If aboRecursiveFlg Then
        '再帰処理する場合
            Select Case alItemType
                Case PeItemType("FILE_EXCLUDE_ARCHIVE")
                    cf_pushA avAr, oNewItem.selfAndAllFiles()
                Case PeItemType("FOLDER_INCLUDE_ARCHIVE")
                    cf_pushA avAr, oNewItem.selfAndAllFolders()
                Case Else
                    cf_pushA avAr, oNewItem.selfAndAllItems()
            End Select
        Else
        '再帰処理しない場合
            Select Case alItemType
                Case PeItemType("FILE_EXCLUDE_ARCHIVE"),PeItemType("FOLDER_INCLUDE_ARCHIVE")
                    Dim boFlg : If alItemType=PeItemType("FILE_EXCLUDE_ARCHIVE") Then boFlg=False Else boFlg=True
                    If (oNewItem.isFolder Or oNewItem.hasItem)=boFlg Then cf_push avAr, oNewItem
                Case Else
                    cf_push avAr, oNewItem
            End Select
        End If

        Set oNewItem=Nothing
    End Sub

    '***************************************************************************************************
    'Function/Sub Name           : this_name()
    'Overview                    : ファイル／フォルダの名前を返す
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     ファイル／フォルダの名前
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
    'Overview                    : ファイル／フォルダの親フォルダのフルパスを返す
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     ファイル／フォルダの親フォルダの当クラスのインスタンス
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
    'Overview                    : ファイル／フォルダのフルパスを返す
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     ファイル／フォルダのフルパス
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
    'Function/Sub Name           : this_setParent()
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
    'Function/Sub Name           : this_selfAndAllItems()
    'Overview                    : 自身とフォルダー内のアイテムの配列を返す
    'Detailed Description        : 工事中
    'Argument
    '     alItemType             : Cl_FILE:ファイルのみ / Cl_FOLDER:フォルダーのみ / 左記以外:全て
    'Return Value
    '     当クラスのインスタンスの配列
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/04/17         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_selfAndAllItems( _
        byVal alItemType _
        )
        this_selfAndAllItems = Null
        If this_isInitial() Then Exit Function

        Dim vRet : vRet = Array()
        Dim boFlg : boFlg = (this_isFolder() Or this_hasItem(PeItemType("ALL")))
        Select Case alItemType  
            Case PeItemType("FILE_EXCLUDE_ARCHIVE")
                If Not boFlg Then vRet=Array(Me)
            Case PeItemType("FOLDER_INCLUDE_ARCHIVE")
                If boFlg Then vRet=Array(Me)
            Case Else
                vRet=Array(Me)
        End Select
        
        cf_pushA vRet, this_items(True, alItemType)
        this_selfAndAllItems = vRet
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : this_size()
    'Overview                    : ファイル／フォルダのサイズを返す
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     ファイル／フォルダのサイズ（バイト単位）
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/03/15         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_size()
        this_size = Null
        If this_isInitial() Then Exit Function

        If new_Fso().FolderExists(PsPath) Then
        'フォルダの場合
            this_size = new_FolderOf(PsPath).Size
        Else
        'フォルダ以外の場合
            this_size = PoFolderItem.Size
        End If
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : this_type()
    'Overview                    : ファイル／フォルダの種類を返す
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     ファイル／フォルダの種類
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
    'Function/Sub Name           : this_setData()
    'Overview                    : データを設定する
    'Detailed Description        : 工事中
    'Argument
    '     asPath                 : ファイル／フォルダのパス
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
    'Function/Sub Name           : this_setPath()
    'Overview                    : ファイル／フォルダのパスを設定する
    'Detailed Description        : 工事中
    'Argument
    '     asPath                 : ファイル／フォルダのパス
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

End Class
