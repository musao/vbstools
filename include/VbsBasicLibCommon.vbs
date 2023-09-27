'***************************************************************************************************
'FILENAME                    : VbsBasicLibCommon.vbs
'Overview                    : 共通関数ライブラリ
'Detailed Description        : 工事中
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/09/27         Y.Fujii                  First edition
'***************************************************************************************************


'オフィス全般

'***************************************************************************************************
'Function/Sub Name           : sub_CM_OfficeUnprotect()
'Overview                    : 文書の保護を解除する
'Detailed Description        : エラーは無視する
'                              引数のパスワードを指定しない場合は、呼び出し側でvbNullStringを設定すること
'Argument
'     aoOffice               : オフィスのインスタンス、エクセルの場合はワークブック
'     asPassword             : パスワード
'Return Value
'     なし
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/09/27         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub sub_CM_OfficeUnprotect( _
    byRef aoOffice _
    , byVal asPassword _
    )
    On Error Resume Next
    aoOffice.Unprotect(asPassword)
    If Err.Number Then
        Err.Clear
    End If
End Sub



'エクセル系

'***************************************************************************************************
'Function/Sub Name           : sub_CM_ExcelSaveAs()
'Overview                    : エクセルファイルを別名で保存して閉じる
'Detailed Description        : 工事中
'Argument
'     aoWorkBook             : エクセルのワークブック
'     asPath                 : 保存するファイルのフルパス
'     alFileformat           : XlFileFormat 列挙体（デフォルトはxlOpenXMLWorkbook 51 Excelブック）
'Return Value
'     なし
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/09/27         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub sub_CM_ExcelSaveAs( _
    byRef aoWorkBook _
    , byVal asPath _
    , byVal alFileformat _
    )
    If Not(IsNumeric(alFileformat)) Then
        alFileformat = 51                  'xlOpenXMLWorkbook 51 Excelブック
    End If
    Call aoWorkBook.SaveAs( _
                            asPath _
                            , alFileformat _
                            , , _
                            , False _
                            , False _
                            )
    Call aoWorkBook.Close(False)
End Sub

'***************************************************************************************************
'Function/Sub Name           : func_CM_ExcelOpenFile()
'Overview                    : エクセルファイルを読み取り専用／ダイアログなしで開く
'Detailed Description        : 工事中
'Argument
'     aoExcel                : エクセル
'     asPath                 : エクセルファイルのフルパス
'Return Value
'     開いたエクセルのワークブック
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/09/27         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_ExcelOpenFile( _
    byRef aoExcel _
    , byVal asPath _
    )    
    Set func_CM_ExcelOpenFile = aoExcel.Workbooks.Open( _
                                                        asPath _
                                                        , 0 _
                                                        , True _
                                                        , , , _
                                                        , True _
                                                        )
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_ExcelGetTextFromAutoshape()
'Overview                    : エクセルのオートシェイプのテキストを取り出す
'Detailed Description        : エラーは無視する
'Argument
'     aoAutoshape            : オートシェイプ
'Return Value
'     オートシェイプのテキスト
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/09/27         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_ExcelGetTextFromAutoshape( _
    byRef aoAutoshape _
    )
    On Error Resume Next
    func_CM_ExcelGetTextFromAutoshape = aoAutoshape.TextFrame.Characters.Text
    If Err.Number Then
        Err.Clear
    End If
End Function


'ファイル操作系

'***************************************************************************************************
'Function/Sub Name           : func_CM_FsDeleteFile()
'Overview                    : ファイルを削除する
'Detailed Description        : FileSystemObjectのDeleteFile()と同等
'Argument
'     asPath                 : 削除するファイルのフルパス
'Return Value
'     結果 True:成功 / False:失敗
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/09/27         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FsDeleteFile( _
    byVal asPath _
    ) 
    If Not func_CM_FsFileExists(asPath) Then func_CM_FsDeleteFile = False
    On Error Resume Next
    CreateObject("Scripting.FileSystemObject").DeleteFile(asPath)
    func_CM_FsDeleteFile = True
    If Err.Number Then
        Err.Clear
        func_CM_FsDeleteFile = False
    End If
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_FsDeleteFolder()
'Overview                    : フォルダを削除する
'Detailed Description        : FileSystemObjectのDeleteFolder()と同等
'Argument
'     asPath                 : 削除するフォルダのフルパス
'Return Value
'     結果 True:成功 / False:失敗
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/11/12         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FsDeleteFolder( _
    byVal asPath _
    ) 
    If Not func_CM_FsFolderExists(asPath) Then func_CM_FsDeleteFolder = False
    On Error Resume Next
    CreateObject("Scripting.FileSystemObject").DeleteFolder(asPath)
    func_CM_FsDeleteFolder = True
    If Err.Number Then
        Err.Clear
        func_CM_FsDeleteFolder = False
    End If
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_FsDeleteFsObject()
'Overview                    : ファイルかフォルダを削除する
'Detailed Description        : func_CM_FsDeleteFile()とfunc_CM_FsDeleteFolder()に委譲
'Argument
'     asPath                 : パス
'Return Value
'     結果 True:成功 / False:失敗
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/08/19         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FsDeleteFsObject( _
    byVal asPath _
    )
    func_CM_FsDeleteFsObject = False
    If func_CM_FsFileExists(asPath) Then func_CM_FsDeleteFsObject = func_CM_FsDeleteFile(asPath)
    If func_CM_FsFolderExists(asPath) Then func_CM_FsDeleteFsObject = func_CM_FsDeleteFolder(asPath)
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_FsCopyFile ()
'Overview                    : ファイルをコピーする
'Detailed Description        : FileSystemObjectのCopyFile()と同等
'Argument
'     asPathFrom             : コピー元ファイルのフルパス
'     asPathTo               : コピー先ファイルのフルパス
'Return Value
'     結果 True:成功 / False:失敗
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/10/24         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FsCopyFile( _
    byVal asPathFrom _
    , byVal asPathTo _
    ) 
    If Not func_CM_FsFileExists(asPathFrom) Then func_CM_FsCopyFile = False
    On Error Resume Next
    Call CreateObject("Scripting.FileSystemObject").CopyFile(asPathFrom, asPathTo)
    func_CM_FsCopyFile = True
    If Err.Number Then
        Err.Clear
        func_CM_FsCopyFile = False
    End If
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_FsCopyFolder ()
'Overview                    : フォルダをコピーする
'Detailed Description        : FileSystemObjectのCopyFolder()と同等
'Argument
'     asPathFrom             : コピー元フォルダのフルパス
'     asPathTo               : コピー先フォルダのフルパス
'Return Value
'     結果 True:成功 / False:失敗
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/08/19         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FsCopyFolder( _
    byVal asPathFrom _
    , byVal asPathTo _
    ) 
    If Not func_CM_FsFolderExists(asPathFrom) Then func_CM_FsCopyFolder = False
    On Error Resume Next
    Call CreateObject("Scripting.FileSystemObject").CopyFolder(asPathFrom, asPathTo)
    func_CM_FsCopyFolder = True
    If Err.Number Then
        Err.Clear
        func_CM_FsCopyFolder = False
    End If
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_FsCopyFsObject()
'Overview                    : ファイルかフォルダをコピーする
'Detailed Description        : func_CM_FsCopyFile()とfunc_CM_FsCopyFolder()に委譲
'Argument
'     asPathFrom             : コピー元ファイル/フォルダのフルパス
'     asPathTo               : コピー先のフルパス
'Return Value
'     結果 True:成功 / False:失敗
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/08/19         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FsCopyFsObject( _
    byVal asPathFrom _
    , byVal asPathTo _
    )
    func_CM_FsCopyFsObject = False
    If func_CM_FsFileExists(asPathFrom) Then func_CM_FsCopyFsObject = func_CM_FsCopyFile(asPathFrom, asPathTo)
    If func_CM_FsFolderExists(asPathFrom) Then func_CM_FsCopyFsObject = func_CM_FsCopyFolder(asPathFrom, asPathTo)
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_FsMoveFile ()
'Overview                    : ファイルを移動する
'Detailed Description        : FileSystemObjectのMoveFile()と同等
'Argument
'     asPathFrom             : 移動元ファイルのフルパス
'     asPathTo               : 移動先ファイルのフルパス
'Return Value
'     結果 True:成功 / False:失敗
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/08/19         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FsMoveFile( _
    byVal asPathFrom _
    , byVal asPathTo _
    ) 
    If Not func_CM_FsFileExists(asPathFrom) Then func_CM_FsMoveFile = False
    On Error Resume Next
    Call CreateObject("Scripting.FileSystemObject").MoveFile(asPathFrom, asPathTo)
    func_CM_FsMoveFile = True
    If Err.Number Then
        Err.Clear
        func_CM_FsMoveFile = False
    End If
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_FsMoveFolder ()
'Overview                    : フォルダを移動する
'Detailed Description        : FileSystemObjectのMoveFolder()と同等
'Argument
'     asPathFrom             : 移動元フォルダのフルパス
'     asPathTo               : 移動先フォルダのフルパス
'Return Value
'     結果 True:成功 / False:失敗
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/08/19         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FsMoveFolder( _
    byVal asPathFrom _
    , byVal asPathTo _
    ) 
    If Not func_CM_FsFolderExists(asPathFrom) Then func_CM_FsMoveFolder = False
    On Error Resume Next
    Call CreateObject("Scripting.FileSystemObject").MoveFolder(asPathFrom, asPathTo)
    func_CM_FsMoveFolder = True
    If Err.Number Then
        Err.Clear
        func_CM_FsMoveFolder = False
    End If
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_FsMoveFsObject()
'Overview                    : ファイルかフォルダを移動する
'Detailed Description        : func_CM_FsMoveFile()とfunc_CM_FsMoveFolder()に委譲
'Argument
'     asPathFrom             : 移動元ファイル/フォルダのフルパス
'     asPathTo               : 移動先のフルパス
'Return Value
'     結果 True:成功 / False:失敗
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/08/19         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FsMoveFsObject( _
    byVal asPathFrom _
    , byVal asPathTo _
    )
    func_CM_FsMoveFsObject = False
    If func_CM_FsFileExists(asPathFrom) Then func_CM_FsMoveFsObject = func_CM_FsMoveFile(asPathFrom, asPathTo)
    If func_CM_FsFolderExists(asPathFrom) Then func_CM_FsMoveFsObject = func_CM_FsMoveFolder(asPathFrom, asPathTo)
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_FsGetParentFolderPath()
'Overview                    : 親フォルダパスの取得
'Detailed Description        : FileSystemObjectのGetParentFolderName()と同等
'Argument
'     asPath                 : ファイルのパス
'Return Value
'     親フォルダパス
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/09/27         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FsGetParentFolderPath( _
    byVal asPath _
    ) 
    func_CM_FsGetParentFolderPath = CreateObject("Scripting.FileSystemObject").GetParentFolderName(asPath)
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_FsGetGetBaseName()
'Overview                    : ファイル名（拡張子を除く）の取得
'Detailed Description        : FileSystemObjectのGetBaseName()と同等
'Argument
'     asPath                 : ファイルのパス
'Return Value
'     ファイル名（拡張子を除く）
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/10/23         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FsGetGetBaseName( _
    byVal asPath _
    ) 
    func_CM_FsGetGetBaseName = CreateObject("Scripting.FileSystemObject").GetBaseName(asPath)
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_FsGetGetExtensionName()
'Overview                    : ファイルの拡張子の取得
'Detailed Description        : FileSystemObjectのGetExtensionName()と同等
'Argument
'     asPath                 : ファイルのパス
'Return Value
'     ファイルの拡張子
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/10/23         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FsGetGetExtensionName( _
    byVal asPath _
    ) 
    func_CM_FsGetGetExtensionName = CreateObject("Scripting.FileSystemObject").GetExtensionName(asPath)
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_FsBuildPath()
'Overview                    : ファイルパスの連結
'Detailed Description        : FileSystemObjectのBuildPath()と同等
'Argument
'     asFolderPath           : パス
'     asItemName             : asFolderPathに連結するフォルダ名またはファイル名
'Return Value
'     連結したファイルパス
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/09/27         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FsBuildPath( _
    byVal asFolderPath _
    , byVal asItemName _
    ) 
    func_CM_FsBuildPath = CreateObject("Scripting.FileSystemObject").BuildPath(asFolderPath, asItemName)
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_FsFileExists()
'Overview                    : ファイルの存在確認
'Detailed Description        : FileSystemObjectのFileExists()と同等
'Argument
'     asPath                 : パス
'Return Value
'     結果 True:存在する / False:存在しない
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/09/27         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FsFileExists( _
    byVal asPath _
    ) 
    func_CM_FsFileExists = CreateObject("Scripting.FileSystemObject").FileExists(asPath)
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_FsFolderExists()
'Overview                    : フォルダの存在確認
'Detailed Description        : FileSystemObjectのFolderExists()と同等
'Argument
'     asPath                 : パス
'Return Value
'     結果 True:存在する / False:存在しない
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/10/16         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FsFolderExists( _
    byVal asPath _
    ) 
    func_CM_FsFolderExists = CreateObject("Scripting.FileSystemObject").FolderExists(asPath)
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_FsGetFile()
'Overview                    : ファイルオブジェクトの取得
'Detailed Description        : FileSystemObjectのGetFile()と同等
'Argument
'     asPath                 : パス
'Return Value
'     Fileオブジェクト
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/09/27         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FsGetFile( _
    byVal asPath _
    ) 
    Set func_CM_FsGetFile = CreateObject("Scripting.FileSystemObject").GetFile(asPath)
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_FsGetFolder()
'Overview                    : フォルダオブジェクトの取得
'Detailed Description        : FileSystemObjectのGetFolder()と同等
'Argument
'     asPath                 : パス
'Return Value
'     Folderオブジェクト
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/10/23         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FsGetFolder( _
    byVal asPath _
    ) 
    Set func_CM_FsGetFolder = CreateObject("Scripting.FileSystemObject").GetFolder(asPath)
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_FsGetFsObject()
'Overview                    : ファイルかフォルダオブジェクトの取得
'Detailed Description        : func_CM_FsGetFile()とfunc_CM_FsGetFolder()に委譲
'Argument
'     asPath                 : パス
'Return Value
'     File/Folderオブジェクト
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/08/19         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FsGetFsObject( _
    byVal asPath _
    )
    Set func_CM_FsGetFsObject = Nothing
    If func_CM_FsFileExists(asPath) Then Set func_CM_FsGetFsObject = func_CM_FsGetFile(asPath)
    If func_CM_FsFolderExists(asPath) Then Set func_CM_FsGetFsObject = func_CM_FsGetFolder(asPath)
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_FsGetFiles()
'Overview                    : 指定したフォルダ以下のFilesコレクションを取得する
'Detailed Description        : FileSystemObjectのFolderオブジェクトのFilesコレクションと同等
'Argument
'     asPath                 : パス
'Return Value
'     Filesコレクション
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/10/23         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FsGetFiles( _
    byVal asPath _
    ) 
    Set func_CM_FsGetFiles = CreateObject("Scripting.FileSystemObject").GetFolder(asPath).Files
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_FsGetFolders()
'Overview                    : 指定したフォルダ以下のFoldersコレクションを取得する
'Detailed Description        : FileSystemObjectのFolderオブジェクトのSubFoldersコレクションと同等
'Argument
'     asPath                 : パス
'Return Value
'     Foldersコレクション
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/10/23         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FsGetFolders( _
    byVal asPath _
    ) 
    Set func_CM_FsGetFolders = CreateObject("Scripting.FileSystemObject").GetFolder(asPath).SubFolders
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_FsGetFsObjects()
'Overview                    : 指定したフォルダ以下のFilesコレクションとFoldersコレクションを取得する
'Detailed Description        : func_CM_FsGetFiles()とfunc_CM_FsGetFolders()に委譲
'Argument
'     asPath                 : パス
'Return Value
'     FilesコレクションとFoldersコレクション
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/08/19         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FsGetFsObjects( _
    byVal asPath _
    )
    Set func_CM_FsGetFsObjects = Nothing
    If Not func_CM_FsFolderExists(asPath) Then Exit Function
    Dim oTemp : Set oTemp = new_Dictionary()
    With oTemp
        .Add "Filse", func_CM_FsGetFiles(asPath)
        .Add "Folders", func_CM_FsGetFolders(asPath)
    End With
    Set func_CM_FsGetFsObjects = oTemp
    Set oTemp = Nothing
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_FsGetTempFileName()
'Overview                    : ランダムに生成された一時ファイルまたはフォルダーの名前の取得
'Detailed Description        : FileSystemObjectのGetTempName()と同等
'Argument
'     asPath                 : パス
'Return Value
'     一時ファイルまたはフォルダーの名前
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/09/27         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FsGetTempFileName()
    func_CM_FsGetTempFileName = CreateObject("Scripting.FileSystemObject").GetTempName()
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_FsGetPrivateFilePath()
'Overview                    : 実行中のスクリプトがあるフォルダからのパスを返す
'Detailed Description        : 上位フォルダが存在しない場合は作成する
'Argument
'     asParentFolderName     : 親フォルダ名
'     asFileName             : ファイル名
'Return Value
'     ファイルのフルパス
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/26         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FsGetPrivateFilePath( _
    byVal asParentFolderName _
    , byVal asFileName _
    )
    Dim sParentFolderPath : sParentFolderPath = func_CM_FsGetParentFolderPath(WScript.ScriptFullName)
    If Len(asParentFolderName)>0 Then
    '引数で指定したディレクトリ名がある場合
        sParentFolderPath = func_CM_FsBuildPath(sParentFolderPath ,asParentFolderName)
    End If
    func_CM_FsGetPrivateFilePath = func_CM_FsGetFilePathWithCreateParentFolder(sParentFolderPath, asFileName)
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_FsGetTempFilePath()
'Overview                    : 一時ファイルのパスを返す
'Detailed Description        : 実行中のスクリプトがあるフォルダのtmpフォルダ以下に作成する
'                              上位フォルダが存在しない場合は作成する
'Argument
'     なし
'Return Value
'     ファイルのフルパス
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/26         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FsGetTempFilePath( _
    )
    func_CM_FsGetTempFilePath = func_CM_FsGetPrivateFilePath("tmp", func_CM_FsGetTempFileName())
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_FsGetPrivateLogFilePath()
'Overview                    : 実行中のスクリプトのログファイルパスを返す
'Detailed Description        : 実行中のスクリプトがあるフォルダのlogフォルダ以下に
'                              スクリプトファイル名＋".log"形式のファイル名で作成する
'                              上位フォルダが存在しない場合は作成する
'Argument
'     なし
'Return Value
'     ファイルのフルパス
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/26         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FsGetPrivateLogFilePath( _
    )
    func_CM_FsGetPrivateLogFilePath = func_CM_FsGetPrivateFilePath("log", func_CM_FsGetGetBaseName(WScript.ScriptName) & ".log" )
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_FsGetFilePathWithCreateParentFolder()
'Overview                    : ファイルのパスを取得
'Detailed Description        : 上位フォルダが存在しない場合は作成する
'Argument
'     asParentFolderPath     : 親フォルダのパス
'     asFileName             : ファイル名
'Return Value
'     ファイルのフルパス
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/26         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FsGetFilePathWithCreateParentFolder( _
    byVal asParentFolderPath _
    , byVal asFileName _
    )
    If Not(func_CM_FsFolderExists(asParentFolderPath)) Then func_CM_FsCreateFolder(asParentFolderPath)
    func_CM_FsGetFilePathWithCreateParentFolder = func_CM_FsBuildPath(asParentFolderPath, asFileName)
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_FsCreateFolder()
'Overview                    : フォルダを作成する
'Detailed Description        : FileSystemObjectのCreateFolder()と同等
'Argument
'     asPath                 : パス
'Return Value
'     作成したフォルダのフルパス
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/09/27         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FsCreateFolder( _
    byVal asPath _
    )
    func_CM_FsCreateFolder = CreateObject("Scripting.FileSystemObject").CreateFolder(asPath)
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_FsOpenTextFile()
'Overview                    : ファイルを開きTextStreamオブジェクトを返す
'Detailed Description        : FileSystemObjectのOpenTextFile()と同等
'Argument
'     asPath                 : パス
'     alIomode               : 入力/出力モード 1:ForReading,2:ForWriting,8:ForAppending
'     aboCreate              : asPathが存在しない場合 True:新しいファイルを作成する、False:作成しない
'     asFileFormat           : ファイルの形式 -2:TristateUseDefault,-1:TristateTrue,0:TristateFalse
'Return Value
'     TextStreamオブジェクト
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/01/09         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FsOpenTextFile( _
    byVal asPath _
    , byVal alIomode _
    , byVal aboCreate _
    , byVal asFileFormat _
    )
    'ファイルを開く
    Set func_CM_FsOpenTextFile = CreateObject("Scripting.FileSystemObject").OpenTextFile( _
                                                              asPath _
                                                              , alIomode _
                                                              , aboCreate _
                                                              , asFileFormat _
                                                              )
End Function

'***************************************************************************************************
'Function/Sub Name           : sub_CM_FsWriteFile()
'Overview                    : ファイル出力する
'Detailed Description        : エラーは無視する
'Argument
'     asPath                 : 出力先のフルパス
'     asCont                 : 出力する内容
'     なし
'Return Value
'     作成したフォルダのフルパス
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/10/16         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub sub_CM_FsWriteFile( _
    byVal asPath _
    , byVal asCont _
    )
    On Error Resume Next
    'ファイルを開く（存在しない場合は作成する）
    With func_CM_FsOpenTextFile(asPath, 2, True, -2)
'    With CreateObject("Scripting.FileSystemObject").OpenTextFile(asPath, 2, True)
        Call .WriteLine(asCont)
        Call .Close
    End With
    If Err.Number Then
        Err.Clear
    End If
End Sub

'***************************************************************************************************
'Function/Sub Name           : func_CM_FsIsSame()
'Overview                    : 指定したパスが同じファイル/フォルダか検査する
'Detailed Description        : 工事中
'Argument
'     asPathA                : ファイル/フォルダのフルパス
'     asPathB                : ファイル/フォルダのフルパス
'Return Value
'     結果 True:同一 / False:同一でない
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/08/19         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FsIsSame( _
    byVal asPathA _
    , byVal asPathB _
    )
    func_CM_FsIsSame = (func_CM_FsGetFsObject(asPathA) Is func_CM_FsGetFsObject(asPathB))
End Function


'文字列操作系

'***************************************************************************************************
'Function/Sub Name           : func_CM_StrConvOnlyAlphabet()
'Overview                    : 英字だけ大文字／小文字に変換する
'Detailed Description        : 工事中
'Argument
'     asTarget               : 変換する文字列
'     alConversion           : 実行する変換の種類 1:UpperCase,2:LowerCase
'Return Value
'     変換した文字列
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/08/19         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_StrConvOnlyAlphabet( _
    byVal asTarget _
    , byVal alConversion _
    )
    Dim sChar, asTargetTmp
    
    '1文字ずつ判定する
    Dim asTargetNew : asTargetNew = asTarget
    Dim lPos : lPos = 1
    Do While Len(asTargetNew) >= lPos
        '変換対象の1文字を取得
        sChar = Mid(asTargetNew, lPos, 1)
        
        If func_CM_StrDetermineCharacterType(sChar, 1) Then
        '変換対象が英字の場合のみ変換する
            asTargetTmp = ""
            
            '変換対象の文字までの文字列
            If lPos > 1 Then
                asTargetTmp = Mid(asTargetNew, 1, lPos-1)
            End If
            
            '変換した1文字を結合
            sChar = func_CM_StrConv(sChar, alConversion)
            asTargetTmp = asTargetTmp & sChar
            
            '変換対象の文字移行の文字列を結合
            If lPos < Len(asTargetNew) Then
                asTargetTmp = asTargetTmp & Mid(asTargetNew, lPos+1, Len(asTargetNew)-lPos)
            End If
            
            '変換後の文字列を格納
            asTargetNew = asTargetTmp
        End If
        
        'カウントアップ
        lPos = lPos+1
    Loop
    
    func_CM_StrConvOnlyAlphabet = asTargetNew
    
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_StrDetermineCharacterType()
'Overview                    : 文字種を判定する
'Detailed Description        : 工事中
'Argument
'     asChar                 : 文字種を判定する文字
'     alType                 : 文字種 1:半角のAlphabet
'Return Value
'     結果 True:合致する / False:合致しない
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/08/19         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_StrDetermineCharacterType( _
    byVal asChar _
    , byVal alType _
    )
    func_CM_StrDetermineCharacterType = False
    Select Case alType
        Case 1:
            If _
                    Asc("A") <= Asc(asChar) And Asc(asChar) <= Asc("Z") _
                Or Asc("a") <= Asc(asChar) And Asc(asChar) <= Asc("z")  _
                Then
                func_CM_StrDetermineCharacterType = True
            End If
    End Select
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_StrConv()
'Overview                    : 文字列を指定のとおり変換する
'Detailed Description        : 工事中
'Argument
'     asTarget               : 変換する文字列
'     alConversion           : 実行する変換の種類（現時点で1,2のみ実装）
'                                 1:文字列を大文字に変換
'                                 2:文字列を小文字に変換
'                                 3:文字列内のすべての単語の最初の文字を大文字に変換
'                                 4:文字列内の狭い (1 バイト) 文字をワイド (2 バイト) 文字に変換
'                                 8:文字列内のワイド (2 バイト) 文字を狭い (1 バイト) 文字に変換
'                                16:文字列内のひらがな文字をカタカナ文字に変換
'                                32:文字列内のカタカナ文字をひらがな文字に変換
'Return Value
'     変換した文字列
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/08/19         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_StrConv( _
    byVal asTarget _
    , byVal alalConversion _
    )
    Dim sChar, asTargetTmp
    func_CM_StrConv = asTarget
    Select Case alalConversion
        Case 1:
            func_CM_StrConv = UCase(asTarget)
        Case 2:
            func_CM_StrConv = LCase(asTarget)
    End Select
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_StrLen()
'Overview                    : 全角は2文字、半角は1文字として文字数を返す
'Detailed Description        : 工事中
'Argument
'     asTarget               : 文字列
'Return Value
'     文字数
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/08/19         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_StrLen( _
    byVal asTarget _
    )
    '1文字ずつ判定する
    Dim sChar
    Dim lLength : lLength = 0
    Dim lPos : lPos = 1
    Do While Len(asTarget) >= lPos
        '1文字を取得
        sChar = Mid(asTarget, lPos, 1)
        
        If (Asc(sChar) And &HFF00) <> 0 Then
            lLength = lLength+2
        Else
            lLength = lLength+1
        End If
        
        'カウントアップ
        lPos = lPos+1
    Loop
    
    func_CM_StrLen = lLength
End Function

'数学系

'***************************************************************************************************
'Function/Sub Name           : func_CM_MathMin()
'Overview                    : 最小値を求める
'Detailed Description        : 工事中
'Argument
'     al1                    : 数値1
'     al2                    : 数値2
'Return Value
'     al1とal2の値が小さい方
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/09/27         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_MathMin( _
    byVal al1 _ 
    , byVal al2 _
    )
    Dim lRet
    If al1 < al2 Then lRet = al1 Else lRet = al2
    func_CM_MathMin = lRet
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_MathMax()
'Overview                    : 最大値を求める
'Detailed Description        : 工事中
'Argument
'     al1                    : 数値1
'     al2                    : 数値2
'Return Value
'     al1とal2の値が大きい方
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/09/27         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_MathMax( _
    byVal al1 _ 
    , byVal al2 _
    )
    Dim lRet
    If al1 > al2 Then lRet = al1 Else lRet = al2
    func_CM_MathMax = lRet
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_MathRoundUp()
'Overview                    : 切り上げする
'Detailed Description        : 工事中
'Argument
'     adbNumber              : 数値
'     aiPlace                : 切り上げする小数点以下の桁数
'Return Value
'     切り上げした値
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/09/24         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_MathRoundUp( _
    byVal adbNumber _ 
    , byVal aiPlace _
    )
    func_CM_MathRoundUp = func_CM_MathRound(adbNumber, aiPlace, 9)
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_MathRoundOff()
'Overview                    : 四捨五入する
'Detailed Description        : 工事中
'Argument
'     adbNumber              : 数値
'     aiPlace                : 四捨五入する小数点以下の桁数
'Return Value
'     四捨五入した値
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/09/24         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_MathRoundOff( _
    byVal adbNumber _ 
    , byVal aiPlace _
    )
    func_CM_MathRoundOff = func_CM_MathRound(adbNumber, aiPlace, 5)
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_MathRoundDown()
'Overview                    : 切り捨てする
'Detailed Description        : 工事中
'Argument
'     adbNumber              : 数値
'     aiPlace                : 切り捨てする小数点以下の桁数
'Return Value
'     切り捨てした値
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/09/24         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_MathRoundDown( _
    byVal adbNumber _ 
    , byVal aiPlace _
    )
    func_CM_MathRoundDown = func_CM_MathRound(adbNumber, aiPlace, 0)
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_MathRound()
'Overview                    : 数値を丸める
'Detailed Description        : 工事中
'Argument
'     adbNumber              : 数値
'     aiPlace                : 丸める小数点以下の桁数
'     alThreshold            : 閾値
'                               0：切り捨て
'                               5：四捨五入
'                               9：切り上げ
'Return Value
'     丸めた値
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/09/24         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_MathRound( _
    byVal adbNumber _ 
    , byVal aiPlace _
    , byVal alThreshold _
    )
    Dim lMultiply, lReverse
    lReverse = 10^(aiPlace-1)
    lMultiply = 10^(-1*aiPlace)
    func_CM_MathRound = Int((adbNumber + alThreshold*lMultiply)*lReverse)/lReverse
End Function


'配列系

'***************************************************************************************************
'Function/Sub Name           : func_CM_ArrayIsAvailable()
'Overview                    : 有効な配列か検査する
'Detailed Description        : 初期状態ではなく要素を1つ以上含む配列
'Argument
'     avArray                : 検査対象の配列
'Return Value
'     結果 True:有効 / False:無効
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/18         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_ArrayIsAvailable( _
    byRef avArray _
    )
    func_CM_ArrayIsAvailable = False
    On Error Resume Next
    If IsArray(avArray) And (Not IsEmpty(avArray)) Then
        Ubound(avArray)
        If Err.Number=0 Then
            func_CM_ArrayIsAvailable = True
        Else
            Err.Clear
        End If
    End If
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_ArrayGetDimensionNumber()
'Overview                    : 配列の次元数を求める
'Detailed Description        : 工事中
'Argument
'     avArray                : 配列
'Return Value
'     配列の次元数
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/11/19         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_ArrayGetDimensionNumber( _
    byRef avArray _ 
    )
   If Not IsArray(avArray) Then Exit Function
   On Error Resume Next
   Dim lNum : lNum = 0
   Dim lTemp
   Do
       lNum = lNum + 1
       lTemp = UBound(avArray, lNum)
   Loop Until Err.Number <> 0
   Err.Clear
   func_CM_ArrayGetDimensionNumber = lNum - 1
End Function

'***************************************************************************************************
'Function/Sub Name           : sub_CM_Push()
'Overview                    : 配列に要素を追加する
'Detailed Description        : 工事中
'Argument
'     avArray                : 配列
'     avItem                 : 追加する要素
'Return Value
'     配列の次元数
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/11/23         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub sub_CM_Push( _
    byRef avArray _ 
    , byRef avItem _ 
    )
    If func_CM_ArrayIsAvailable(avArray) Then
        Redim Preserve avArray(Ubound(avArray)+1)
    Else
        Redim avArray(0)
    End If
    Call sub_CM_Bind(avArray(Ubound(avArray)), avItem)
    
'    On Error Resume Next
'    If Ubound(avArray)>=0 Then
'        Redim Preserve avArray(Ubound(avArray)+1)
'    End If
'    If Err.Number Then
'        Redim avArray(0)
'        Err.Clear
'    End If
'    Call sub_CM_Bind(avArray(Ubound(avArray)), avItem)
End Sub

'チェック系

'***************************************************************************************************
'Function/Sub Name           : func_CM_ValidationlIsWithinTheRangeOf()
'Overview                    : 数値型の範囲内にあるか検査する
'Detailed Description        : 工事中
'Argument
'     avNumber               : 数値
'     alType                 : 変数の型
'                                1:整数型（Integer）
'                                2:長整数型（Long）
'                                3:バイト型（Byte）
'                                4:単精度浮動小数点型（Single）
'                                5:倍精度浮動小数点型（Double）
'                                6:通貨型（Currency）
'Return Value
'     整形した浮動小数点型
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/08/26         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_ValidationlIsWithinTheRangeOf( _
    byVal avNumber _
    , byVal alType _
    )
    Dim vMin,vMax
    Select Case alType
        Case 1:                   '整数型（Integer）
            vMin = -1 * 2^15
            vMax = 2^15 - 1
        Case 2:                   '長整数型（Long）
            vMin = -1 * 2^31
            vMax = 2^31 - 1
        Case 3:                   'バイト型（Byte）
            vMin = 0
            vMax = 2^8 - 1
        Case 4:                   '単精度浮動小数点型（Single）
            vMin = -3.402823E38
            vMax = 3.402823E38
        Case 5:                   '倍精度浮動小数点型（Double）
            vMin = -1.79769313486231E308
            vMax = 1.79769313486231E308
        Case 6:                   '通貨型（Currency）
            vMin = -1 * 2^59 / 1000
            vMax = ( 2^59 - 1 ) / 1000
    End Select
    
    func_CM_ValidationlIsWithinTheRangeOf = False
    If vMin<=avNumber And avNumber<=vMax Then
        func_CM_ValidationlIsWithinTheRangeOf = True
    End If
End Function


'インスタンス生成系

'***************************************************************************************************
'Function/Sub Name           : new_Dictionary()
'Overview                    : Dictionaryオブジェクト生成関数
'Detailed Description        : 工事中
'Argument
'     なし
'Return Value
'     生成したDictionaryオブジェクトのインスタンス
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/03         Y.Fujii                  First edition
'***************************************************************************************************
Private Function new_Dictionary( _
    )
    Set new_Dictionary = CreateObject("Scripting.Dictionary")
End Function

'***************************************************************************************************
'Function/Sub Name           : new_DictSetValues()
'Overview                    : Dictionaryオブジェクトを生成し初期値を設定する
'Detailed Description        : 工事中
'Argument
'     avParams               : 初期値奇数（1,3,5,...）はKey、偶数（2,4,6,...）はValue
'                              Keyだけの場合は値にEmptyを設定する。
'Return Value
'     生成したDictionaryオブジェクトのインスタンス
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/03         Y.Fujii                  First edition
'***************************************************************************************************
Private Function new_DictSetValues( _
    byVal avParams _
    )
    Dim oDict, vItem, vKey, boIsKey
    
    boIsKey = True
    Set oDict = new_Dictionary()
    
    For Each vItem In avParams
        If boIsKey Then
            Call sub_CM_Bind(vKey, vItem)
            Call sub_CM_BindAt(oDict, vKey, Empty)
        Else
            Call sub_CM_BindAt(oDict, vKey, vItem)
        End If
        boIsKey = Not boIsKey
    Next
    
    Set new_DictSetValues = oDict
    Set oDict = Nothing
End Function

'***************************************************************************************************
'Function/Sub Name           : new_RegExp()
'Overview                    : 正規表現オブジェクト生成関数
'Detailed Description        : 工事中
'Argument
'     asPattern              : 正規表現のパターン
'     asOptions              : この引数内にある文字の有無で正規表現の以下のプロパティをTrueにする
'                                "i":大文字と小文字を区別する（.IgnoreCase = True）
'                                "g"文字列全体を検索する（.Global = True）
'                                "m"文字列を複数行として扱う（.Multiline = True）
'Return Value
'     生成した正規表現オブジェクトのインスタンス
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/03         Y.Fujii                  First edition
'***************************************************************************************************
Private Function new_RegExp( _
    byVal asPattern _
    , byVal asOptions _
    )
    Dim oRe, sOpts
    
    Set oRe = New RegExp
    oRe.Pattern = asPattern
    
    sOpts = LCase(asOptions)
    If InStr(sOpts, "i") > 0 Then oRe.IgnoreCase = True
    If InStr(sOpts, "g") > 0 Then oRe.Global = True
    If InStr(sOpts, "m") > 0 Then oRe.Multiline = True
    
    Set new_RegExp = oRe
    Set oRe = Nothing
End Function



'これ何系かな

'***************************************************************************************************
'Function/Sub Name           : func_CM_GetObjectByIdFromCollection()
'Overview                    : コレクションから指定したIDのメンバーを取得する
'Detailed Description        : エラーは無視する
'Argument
'     aoClloection           : コレクション
'     asId                   : ID
'Return Value
'     該当するメンバー
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/09/27         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_GetObjectByIdFromCollection( _
    byRef aoClloection _
    , byVal asId _
    )
    On Error Resume Next
    Dim oItem
    For Each oItem In aoClloection
        If oItem.Id = asId Then
            Set func_CM_GetObjectByIdFromCollection = oItem
            Exit Function
        End If
    Next
    Set func_CM_GetObjectByIdFromCollection = Nothing
    If Err.Number Then
        Err.Clear
    End If
    Set oItem = Nothing
End Function

'***************************************************************************************************
'Function/Sub Name           : sub_CM_Bind()
'Overview                    : 変数間の項目移送
'Detailed Description        : 移送する値または変数がオブジェクトか否かによるVBS構文の違い（Setの有無）を吸収する
'                              移送先がコレクションのメンバーの場合は動作しない
'                              移送先が変数の場合に使用できる
'Argument
'     avTo                   : 移送先の変数
'     avValue                : 移送する値または変数
'Return Value
'     なし
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/11/06         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub sub_CM_Bind( _
    byRef avTo _
    , byRef avValue _
    )
    If IsObject(avValue) Then Set avTo = avValue Else avTo = avValue
End Sub

'***************************************************************************************************
'Function/Sub Name           : sub_CM_Swap()
'Overview                    : 変数の値を入れ替える
'Detailed Description        : 移送処理はsub_CM_Bind()を使用する
'Argument
'     avA                    : 値を入れ替える変数
'     avB                    : 値を入れ替える変数
'Return Value
'     なし
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/21         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub sub_CM_Swap( _
    byRef avA _
    , byRef avB _
    )
    Dim oTemp
    Call sub_CM_Bind(oTemp, avA)
    Call sub_CM_Bind(avA, avB)
    Call sub_CM_Bind(avB, oTemp)
    Set oTemp = Nothing
End Sub

'***************************************************************************************************
'Function/Sub Name           : sub_CM_BindAt()
'Overview                    : 変数間の項目移送
'Detailed Description        : 移送する値または変数がオブジェクトか否かによるVBS構文の違い（Setの有無）を吸収する
'                              移送先がコレクションの場合は当関数を使用する
'Argument
'     aoCollection           : 移送先のコレクション
'     asKey                  : 移送先のコレクションのキー
'     avValue                : 移送する値または変数
'Return Value
'     なし
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/11/06         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub sub_CM_BindAt( _
    byRef aoCollection _
    , byVal asKey _
    , byRef avValue _
    )
    If IsObject(avValue) Then Set aoCollection.Item(asKey) = avValue Else aoCollection.Item(asKey) = avValue
End Sub

'***************************************************************************************************
'Function/Sub Name           : func_CM_IsSame()
'Overview                    : 変数の値が等しいか
'Detailed Description        : 比較する変数がオブジェクトか否かによるVBS構文の違い（Setの有無）を吸収する
'Argument
'     avA                    : 比較元
'     avB                    : 比較先
'Return Value
'     なし
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/11/06         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_IsSame( _
    byRef avA _
    , byRef avB _
    )
    Dim boReturn : boReturn = False
    If IsObject(avB) Then boReturn = (avA Is avB) Else boReturn = (avA = avB)
    func_CM_IsSame = boReturn
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_FillInTheCharacters()
'Overview                    : 文字を埋める
'Detailed Description        : 対象文字の不足桁を指定したアライメントで指定した文字の1文字目で埋める
'                              対象文字に不足桁がない場合は、指定した文字数で切り取る
'Argument
'     asTarget               : 対象文字列
'     alWordCount            : 文字数
'     asToFillCharacter      : 埋める文字
'     aboIsCutOut            : 文字数で切り取り（True：する/False：しない）
'     aboIsRightAlignment    : アライメント（True：右寄せ/False：左寄せ）
'Return Value
'     埋めた文字列
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/01/04         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FillInTheCharacters( _
    byVal asTarget _
    , byVal alWordCount _
    , byVal asToFillCharacter _
    , byVal aboIsCutOut _
    , byVal aboIsRightAlignment _
    )
    
    '切り取りなしで対象文字列が文字数より大きい場合は処理を抜ける
    Dim lTargetLen : lTargetLen = Len(asTarget)
    If Not(aboIsCutOut) And lTargetLen>=alWordCount Then
        func_CM_FillInTheCharacters = asTarget
        Exit Function
    End If
    
    '埋める文字列の作成
    Dim sFillStrings : sFillStrings = ""
    If alWordCount-lTargetLen > 0 Then
        sFillStrings = String(alWordCount-lTargetLen , asToFillCharacter)
    End If
    
    Dim sResult
    'アライメント指定によって文字列を埋める
    If aboIsRightAlignment Then
        sResult = sFillStrings & asTarget
    Else
        sResult = asTarget & sFillStrings
    End If
    
    '切り取りなしの場合は処理を抜ける
    If Not(aboIsCutOut) Then
        func_CM_FillInTheCharacters = sResult
        Exit Function
    End If
    
    'アライメント指定によって文字列を切り取る
    If aboIsRightAlignment Then
        sResult = Right(sResult, alWordCount)
    Else
        sResult = Left(sResult, alWordCount)
    End If
    func_CM_FillInTheCharacters = sResult
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_FormatDecimalNumber()
'Overview                    : 浮動小数点型を整形する
'Detailed Description        : 工事中
'Argument
'     adbNumber              : 浮動小数点型の数値
'     alDecimalPlaces        : 小数の桁数
'Return Value
'     整形した浮動小数点型
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/08/26         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FormatDecimalNumber( _
    byVal adbNumber _
    , byVal alDecimalPlaces _
    )
    func_CM_FormatDecimalNumber = Fix(adbNumber) & "." _
                             & func_CM_FillInTheCharacters( _
                                                          Abs(Fix( (adbNumber - Fix(adbNumber))*10^alDecimalPlaces )) _
                                                          , alDecimalPlaces _
                                                          , "0" _
                                                          , False _
                                                          , True _
                                                          )
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_ToString()
'Overview                    : 引数の数値・文字列やオブジェクトの中身を可読な表示に変換する
'Detailed Description        : 配列やディクショナリのようなオブジェクトだったら中身を表示し、
'                              そうでない場合はVarTypeでオブジェクトのクラスを表示する
'Argument
'     avTarget               : 対象
'Return Value
'     変換した文字列
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/03         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_ToString( _
    byRef avTarget _
    )
    Dim oEscapingDoubleQuote, sRet
    Set oEscapingDoubleQuote = new_RegExp("""", "g")
    sRet = ""
    
    Err.Clear
    On Error Resume Next
    
    If VarType(avTarget) = vbString Then
        sRet = """" & oEscapingDoubleQuote.Replace(avTarget, """""") & """"
    ElseIf IsArray(avTarget) Then
        sRet = func_CM_ToStringArray(avTarget)
    ElseIf IsObject(avTarget) Then
        sRet = func_CM_ToStringObject(avTarget)
    ElseIf IsEmpty(avTarget) Then
        sRet = "<empty>"
    ElseIf IsNull(avTarget) Then
        sRet = "<null>"
    Else
        sRet = func_CM_ToStringOther(avTarget)
    End If
    
    If Err.Number <> 0 Then
        Err.Clear
        sRet = func_CM_ToStringUnknown(avTarget)
    End If
    
    func_CM_ToString = sRet
    
    Set oEscapingDoubleQuote = Nothing
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_ToStringArray()
'Overview                    : 配列の中身を可読な表示に変換する
'Detailed Description        : 工事中
'Argument
'     avTarget               : 対象
'Return Value
'     変換した文字列
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/03         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_ToStringArray( _
    byRef avTarget _
    )
    Dim oTemp(), vItem
    
    For Each vItem In avTarget
        Call sub_CM_Push(oTemp, func_CM_ToString(vItem))
    Next
    func_CM_ToStringArray = "[" & Join(oTemp, ",") & "]"
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_ToStringDictionary()
'Overview                    : ディクショナリの中身を可読な表示に変換する
'Detailed Description        : 工事中
'Argument
'     avTarget               : 対象
'Return Value
'     変換した文字列
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/03         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_ToStringDictionary( _
    byRef avTarget _
    )
    Dim oTemp(), vKey
    
    For Each vKey In avTarget.Keys
        Call sub_CM_Push(oTemp, func_CM_ToString(vKey) & "=>" & func_CM_ToString(avTarget.Item(vKey)))
    Next
    func_CM_ToStringDictionary = "{" & Join(oTemp, ",") & "}"
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_ToStringObject()
'Overview                    : オブジェクトの中身を可読な表示に変換する
'Detailed Description        : 工事中
'Argument
'     avTarget               : 対象
'Return Value
'     変換した文字列
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/03         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_ToStringObject( _
    byRef avTarget _
    )
    Dim sRet
    
    Err.Clear
    On Error Resume Next
    
    sRet = func_CM_ToStringDictionary(avTarget)
    
    If Err.Number <> 0 Then
        Err.Clear
        sRet = func_CM_ToStringArray(avTarget)
    End If
    
    If Err.Number <> 0 Then
        Err.Clear
        sRet = func_CM_ToStringArray(avTarget.Items)
    End If
    
    If Err.Number <> 0 Then
        Err.Clear
        sRet = "<" & TypeName(avTarget) & ">"
    End If
    
    func_CM_ToStringObject = sRet
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_ToStringOther()
'Overview                    : その他オブジェクトの中身を可読な表示に変換する
'Detailed Description        : 工事中
'Argument
'     avTarget               : 対象
'Return Value
'     変換した文字列
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/03         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_ToStringOther( _
    byRef avTarget _
    )
    Dim sRet
    
    Err.Clear
    On Error Resume Next
    
    sRet = CStr(avTarget)
    
    If Err.Number <> 0 Then
        Err.Clear
        sRet = func_CM_ToStringArray(avTarget)
    End If
    
    If Err.Number <> 0 Then
        Err.Clear
        sRet = func_CM_ToStringDictionary(avTarget)
    End If
    
    If Err.Number <> 0 Then
        Err.Clear
        sRet = func_CM_ToStringUnknown(avTarget)
    End If
    
    func_CM_ToStringOther = sRet
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_ToStringUnknown()
'Overview                    : 引数の型が不明な場合に可読な表示に変換する
'Detailed Description        : 工事中
'Argument
'     avTarget               : 対象
'Return Value
'     変換した文字列
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/03         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_ToStringUnknown( _
    byRef avTarget _
    )
    func_CM_ToStringUnknown = "<unknown:" & VarType(avTarget) & " " & TypeName(avTarget) & ">"
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_ToStringErr()
'Overview                    : Errオブジェクトの内容を可読な表示に変換する
'Detailed Description        : 工事中
'Argument
'     なし
'Return Value
'     変換した文字列
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/25         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_ToStringErr( _
    )
    Dim oRet : Set oRet = new_clsCmArray()
    oRet.Push "Number => " & Err.Number
    oRet.Push "Source => """ & Err.Source & """"
    oRet.Push "Description => """ & Err.Description & """"
    func_CM_ToStringErr = "<Err> {" & oRet.JoinVbs(",") & "}"
    Set oRet = Nothing
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_ToStringArguments()
'Overview                    : Argumentsオブジェクトの内容を可読な表示に変換する
'Detailed Description        : 工事中
'Argument
'     なし
'Return Value
'     変換した文字列
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/26         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_ToStringArguments( _
    )
    func_CM_ToStringArguments = func_CM_ToString(func_CM_UtilStoringArguments())
End Function

'***************************************************************************************************
'Function/Sub Name           : sub_CM_ExcuteSub()
'Overview                    : 関数を実行する
'Detailed Description        : 工事中
'Argument
'     asSubName              : 実行する関数名
'     aoArgument             : 実行する関数に渡す引数
'     aoPubSub               : 出版-購読型（Publish/subscribe）クラスのオブジェクト
'     asTopic                : トピック
'Return Value
'     なし
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/03         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub sub_CM_ExcuteSub( _
    byVal asSubName _
    , byRef aoArgument _
    , byRef aoPubSub _
    , byVal asTopic _
    )
    On Error Resume Next
    
    '出版（Publish） 開始
    If Not aoPubSub Is Nothing Then
        Call aoPubSub.Publish(asTopic, Array(5 ,asSubName ,"Start"))
        Call aoPubSub.Publish(asTopic, Array(9 ,asSubName ,func_CM_ToString(aoArgument)))
    End If
    
    '関数の実行
    Dim oFunc : Set oFunc = GetRef(asSubName)
    If aoArgument Is Nothing Then
        Call oFunc()
    Else
        Call oFunc(aoArgument)
    End If
    
    If Not aoPubSub Is Nothing Then
        If Err.Number <> 0 Then
        'エラー
            Call aoPubSub.Publish(asTopic, Array(1, asSubName, func_CM_ToStringErr()))
        Else
        '正常
            Call aoPubSub.Publish(asTopic, Array(5, asSubName, "End"))
        End If
        Call aoPubSub.Publish(asTopic, Array(9, asSubName, func_CM_ToString(aoArgument)))
    End If
    
    Set oFunc = Nothing
End Sub


'ユーティリティ系

'***************************************************************************************************
'Function/Sub Name           : func_CM_UtilSortBubble()
'Overview                    : バブルソート
'Detailed Description        : 計算回数はO(N^2)
'                              引数の関数の引数は以下のとおり
'                                currentValue :配列の要素
'                                nextValue    :次の配列の要素
'Argument
'     avArray                : 配列
'     aoFunc                 : 関数
'     aboFlg                 : 判定方法
'                                True  :昇順（関数の結果がTrueの場合に入れ替える）
'                                False :降順（関数の結果がFalseの場合に入れ替える）
'Return Value
'     ソート後の配列
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/18         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_UtilSortBubble( _
    byRef avArray _
    , byRef aoFunc _
    , byVal aboFlg _
    )
    If Not func_CM_ArrayIsAvailable(avArray) Then Exit Function
    If Ubound(avArray)=0 Then Exit Function
    
    Dim lEnd, lPos
    lEnd = Ubound(avArray)
    Do While lEnd>0
        For lPos=0 To lEnd-1
            If aoFunc(avArray(lPos), avArray(lPos+1))=aboFlg Then
            'lPos番目の要素と(lPos+1)番目の要素を入れ替える
                Call sub_CM_Swap(avArray(lPos), avArray(lPos+1))
            End If
        Next
        lEnd = lEnd-1
    Loop
    func_CM_UtilSortBubble = avArray
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_UtilSortQuick()
'Overview                    : クイックソート
'Detailed Description        : 計算回数は平均O(N*logN)、最悪はO(N^2)
'                              引数の関数の引数は以下のとおり
'                                currentValue :配列の要素
'                                nextValue    :次の配列の要素
'Argument
'     avArray                : 配列
'     aoFunc                 : 関数
'     aboFlg                 : 判定方法
'                                True  :昇順（関数の結果がTrueの場合に入れ替える）
'                                False :降順（関数の結果がFalseの場合に入れ替える）
'Return Value
'     ソート後の配列
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/18         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_UtilSortQuick( _
    byRef avArray _
    , byRef aoFunc _
    , byVal aboFlg _
    )
    If Not func_CM_ArrayIsAvailable(avArray) Then Exit Function
    
    func_CM_UtilSortQuick = avArray
    If Ubound(avArray)=0 Then Exit Function
    
    '0番目の要素をピボットに決める
    Dim oPivot : Call sub_CM_Bind(oPivot, avArray(0))
    
    'ピボットと要素を関数で判定し判定方法に合致するグループをRight、そうでないグループをLeftとする
    Dim lPos, vRight, vLeft
    For lPos=1 To Ubound(avArray)
        If aoFunc(avArray(lPos), oPivot)=aboFlg Then
            Call sub_CM_Push(vRight, avArray(lPos))
        Else
            Call sub_CM_Push(vLeft, avArray(lPos))
        End If
    Next
    
    '上述で分けたRight、Leftのグループごとに再帰処理する
    vLeft = func_CM_UtilSortQuick(vLeft, aoFunc, aboFlg)
    vRight = func_CM_UtilSortQuick(vRight, aoFunc, aboFlg)
    
    'Leftにピボット＋Rightを結合する
    Call sub_CM_Push(vLeft, oPivot)
    If func_CM_ArrayIsAvailable(vRight) Then
        For lPos=0 To Ubound(vRight)
            Call sub_CM_Push(vLeft, vRight(lPos))
        Next
    End If
    
    func_CM_UtilSortQuick = vLeft
    Set oPivot = Nothing
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_UtilSortMerge()
'Overview                    : マージソート
'Detailed Description        : 計算回数はO(N*logN)
'                              マージ処理はfunc_CM_UtilSortMergeMerge()に委譲する
'Argument
'     avArray                : 配列
'     aoFunc                 : 関数
'     aboFlg                 : 判定方法
'                                True  :昇順（関数の結果がTrueの場合に入れ替える）
'                                False :降順（関数の結果がFalseの場合に入れ替える）
'Return Value
'     ソート後の配列
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/18         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_UtilSortMerge( _
    byRef avArray _
    , byRef aoFunc _
    , byVal aboFlg _
    )
    If Not func_CM_ArrayIsAvailable(avArray) Then Exit Function
    
    func_CM_UtilSortMerge = avArray
    If Ubound(avArray)=0 Then Exit Function
    
    '2つの配列に分解する
    Dim lLength, lMedian
    lLength = Ubound(avArray) - Lbound(avArray) + 1
    lMedian = func_CM_MathRoundup(lLength/2, 1)
    Dim lPos, vFirst, vSecond
    For lPos=Lbound(avArray) To lMedian-1
        Call sub_CM_Push(vFirst, avArray(lPos))
    Next
    For lPos=lMedian To Ubound(avArray)
        Call sub_CM_Push(vSecond, avArray(lPos))
    Next
    
    '再帰処理で配列の要素が1つになるまで分解する
    vFirst = func_CM_UtilSortMerge(vFirst, aoFunc, aboFlg)
    vSecond = func_CM_UtilSortMerge(vSecond, aoFunc, aboFlg)
    
    'マージをしながら上位に戻す
    func_CM_UtilSortMerge = func_CM_UtilSortMergeMerge(vFirst, vSecond, aoFunc, aboFlg)
    
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_UtilSortMergeMerge()
'Overview                    : マージソートのマージ処理
'Detailed Description        : func_CM_UtilSortMerge()から呼び出す
'                              引数の関数の引数は以下のとおり
'                                currentValue :配列の要素
'                                nextValue    :次の配列の要素
'Argument
'     avFirst                : マージするソート済みの配列
'     avSecond               : マージするソート済みの配列
'     aoFunc                 : 関数
'     aboFlg                 : 判定方法
'                                True  :昇順（関数の結果がTrueの場合に入れ替える）
'                                False :降順（関数の結果がFalseの場合に入れ替える）
'Return Value
'     マージ済の配列
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/18         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_UtilSortMergeMerge( _
    byRef avFirst _
    , byRef avSecond _
    , byRef aoFunc _
    , byVal aboFlg _
    )
    Dim lPosF, lPosS, lEndF, lEndS
    lPosF = Lbound(avFirst) : lPosS = Lbound(avSecond)
    lEndF = Ubound(avFirst) : lEndS = Ubound(avSecond)
    
    '双方の配列の先頭の要素同士を関数で判定して戻り値の配列に追加する
    Dim vRet
    Do While lPosF<=lEndF And lPosS<=lEndS
        If aoFunc(avFirst(lPosF), avSecond(lPosS))=aboFlg Then
            Call sub_CM_Push(vRet, avSecond(lPosS))
            lPosS = lPosS + 1
        Else
            Call sub_CM_Push(vRet, avFirst(lPosF))
            lPosF = lPosF + 1
        End If
    Loop
    
    'それぞれ残っている方の配列の要素を追加する
    Dim lPos
    If lPosF<=lEndF Then
        For lPos=lPosF To lEndF
            Call sub_CM_Push(vRet, avFirst(lPos))
        Next
    End If
    If lPosS<=lEndS Then
        For lPos=lPosS To lEndS
            Call sub_CM_Push(vRet, avSecond(lPos))
        Next
    End If
    
    'マージ済の配列を返す
    func_CM_UtilSortMergeMerge = vRet
    
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_UtilSortHeap()
'Overview                    : ヒープソート
'Detailed Description        : 計算回数はO(N*logN)
'Argument
'     avArray                : 配列
'     aoFunc                 : 関数
'     aboFlg                 : 判定方法
'                                True  :昇順（関数の結果がTrueの場合に入れ替える）
'                                False :降順（関数の結果がFalseの場合に入れ替える）
'Return Value
'     ソート後の配列
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/21         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_UtilSortHeap( _
    byRef avArray _
    , byRef aoFunc _
    , byVal aboFlg _
    )
    If Not func_CM_ArrayIsAvailable(avArray) Then Exit Function
    
    'ヒープの作成
    Dim lLb, lUb, lSize, lParent
    lLb = Lbound(avArray) : lUb = Ubound(avArray)
    lSize = lUb - lLb + 1
    '子を持つ最下部のノードから上位に向けて順番にノード単位の処理を行う
    For lParent=lSize\2-1 To lLb Step -1
        Call sub_CM_UtilSortHeapPerNodeProc(avArray, lSize, lParent, aoFunc, aboFlg)
    Next
    
    'ヒープの先頭（最大/最小値）を順番に取り出す
    Do While lSize>0
        'ヒープの先頭と末尾を入れ替える
        Call sub_CM_Swap(avArray(lLb), avArray(lSize-1))
        'ヒープサイズを１つ減らして再作成
        lSize = lSize - 1
        Call sub_CM_UtilSortHeapPerNodeProc(avArray, lSize, 0, aoFunc, aboFlg)
    Loop
    
    'ソート済の配列を返す
    func_CM_UtilSortHeap = avArray
    
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_UtilSortHeapPerNodeProc()
'Overview                    : ヒープソートのノード単位の処理
'Detailed Description        : func_CM_UtilSortHeap()から呼び出す
'                              引数の関数の引数は以下のとおり
'                                currentValue :配列の要素
'                                nextValue    :次の配列の要素
'Argument
'     avArray                : 配列
'     alSize                 : ヒープのサイズ
'     alParent               : ノードの親の配列番号
'     aoFunc                 : 関数
'     aboFlg                 : 判定方法
'                                True  :昇順（関数の結果がTrueの場合に入れ替える）
'                                False :降順（関数の結果がFalseの場合に入れ替える）
'Return Value
'     ソート後の配列
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/21         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub sub_CM_UtilSortHeapPerNodeProc( _
    byRef avArray _
    , byVal alSize _
    , byVal alParent _
    , byRef aoFunc _
    , byVal aboFlg _
    )
    Dim lRight, lLeft, lToSwap
    lLeft = alParent*2 + 1
    lRight = lLeft + 1
    lToSwap = alParent
    
    If lRight<alSize Then
    '右側の子がある場合
        If aoFunc(avArray(lRight), avArray(alParent))=aboFlg Then
        '親と右側の子の要素を関数で判定し判定方法に合致する場合は入れ替える
            lToSwap = lRight
        End If
    End If
    
    If lLeft<alSize Then
    '左側の子がある場合
        If aoFunc(avArray(lLeft), avArray(lToSwap))=aboFlg Then
        '親と右側の子の勝者と左側の子の要素を関数で判定し判定方法に合致する場合は入れ替える
            lToSwap = lLeft
        End If
    End If
    
    If lToSwap<>alParent Then
        '親と子の要素を入れ替える
        Call sub_CM_Swap(avArray(alParent), avArray(lToSwap))
        '入れ替えた子の要素以下のノードを再処理する
        Call sub_CM_UtilSortHeapPerNodeProc(avArray, alSize, lToSwap, aoFunc, aboFlg)
    End If
    
End Sub

'***************************************************************************************************
'Function/Sub Name           : func_CM_UtilGenerateRandomNumber()
'Overview                    : 乱数を生成する
'Detailed Description        : 工事中
'Argument
'     adbMin                 : 生成する乱数の最小値
'     adbMax                 : 生成する乱数の最大値
'     aiPlace                : 切り上げする小数点以下の桁数
'Return Value
'     生成した乱数
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/24         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_UtilGenerateRandomNumber( _
    byVal adbMin _
    , byVal adbMax _
    , byVal aiPlace _
    )
    Randomize
    func_CM_UtilGenerateRandomNumber = func_CM_MathRoundDown( (adbMax - adbMin + 1) * Rnd + adbMin, 1 )
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_UtilGenerateRandomString()
'Overview                    : ランダムな文字列を生成する
'Detailed Description        : 指定した長さ、文字の種類でランダムな文字列を生成する
'Argument
'     alLength               : 文字の長さ
'     alType                 : 文字の種類（複数指定する場合は以下の和を設定する）
'                                 1:半角英字大文字
'                                 2:半角英字小文字
'                                 4:半角数字
'                                 8:半角記号
'     avAdditional           : 配列で指定する文字種、前述の文字の種類と重複する場合は追加しない
'                              指定がない場合はNothingなど配列以外を指定する
'Return Value
'     生成した文字列
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/24         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_UtilGenerateRandomString( _
    byVal alLength _
    , byVal alType _
    , byRef avAdditional _
    )
    
    '文字の種類（alType）で指定した文字のリストを作成する
    Dim vSettings
    vSettings = Array( _
                    Array( Array("A", "Z") ) _
                    , Array( Array("a", "z") ) _
                    , Array( Array("0", "9") ) _
                    , Array( Array("!", "/"), Array(":", "@"), Array("[", "`"), Array("{", "~") ) _
                    )
    Dim lType : lType = alType
    Dim lPowerOf2 : lPowerOf2 = 3
    Dim oChars : Set oChars = new_clsCmArray()
    Dim lQuotient,lDivide, vSetting, vItem, bCode
    Do Until lPowerOf2<0
        lDivide = 2^lPowerOf2
        lQuotient = lType \ lDivide
        lType = lType Mod lDivide
        
        If lQuotient>0 Then
            vSetting = vSettings(lPowerOf2)
            For Each vItem In vSetting
                For bCode = Asc(vItem(0)) To Asc(vItem(1))
                    oChars.Push Chr(bCode)
                Next
            Next
        End If
        
        lPowerOf2 = lPowerOf2 - 1
    Loop
    
    '配列で指定する文字種（avAdditional）を追加する
    If func_CM_ArrayIsAvailable(avAdditional) Then
        Dim sChar
        For Each sChar In avAdditional
            If oChars.IndexOf(sChar)<0 Then
                oChars.Push sChar
            End If
        Next
    End If
    
    '上述で作成した文字のリストを使ってランダムな文字列を生成する
    Dim lPos, oRet
    Set oRet = new_clsCmArray()
    For lPos = 1 To alLength
        oRet.Push oChars( func_CM_UtilGenerateRandomNumber(0, oChars.Length - 1, 1) )
    Next
    func_CM_UtilGenerateRandomString = oRet.JoinVbs("")
    
    Set oRet = Nothing
    Set oChars = Nothing
End Function

'***************************************************************************************************
'Function/Sub Name           : sub_CM_UtilCommonLogger()
'Overview                    : ログ出力する
'Detailed Description        : 工事中
'Argument
'     avParams               : 配列型のパラメータリスト
'     aoWriter               : ファイル出力バッファリング処理クラスのインスタンス
'Return Value
'     なし
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/26         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub sub_CM_UtilCommonLogger( _
    byRef avParams _
    , byRef aoWriter _
    )
    Dim oCont : Set oCont = new_clsCmArray()
    oCont.Push new_clsCalGetNow()
    
    Dim vEle
    For Each vEle In avParams
        oCont.Push vEle
    Next
    
    With aoWriter
        .WriteContents(oCont.JoinVbs(vbTab))
        .newLine()
    End With
    
    Set vEle = Nothing
    Set oCont = Nothing
End Sub

'***************************************************************************************************
'Function/Sub Name           : func_CM_UtilStoringArguments()
'Overview                    : Argumentsオブジェクトの内容をオブジェクトに変換する
'Detailed Description        : 変換したオブジェクトの構成
'                              例は引数が a /X /Hoge:Fuga, b の場合
'                              Key         Value                                        例
'                              ----------  -------------------------------------------  -------------
'                              "All"       WScript.Arguments以下のItemの内容            a /X /Hoge:Fuga, b
'                              "Named"     WScript.Arguments.Named以下のItemの内容      X: Hoge:Fuga
'                              "Unnamed"   WScript.Arguments.Unnamed以下のItemの内容    a b
'Argument
'     なし
'Return Value
'     変換したオブジェクト
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/26         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_UtilStoringArguments( _
    )
    Dim oRet : Set oRet = new_Dictionary()
    Dim oTemp, oEle, oKey
    
    'All
    Set oTemp = new_clsCmArray()
    For Each oEle In WScript.Arguments
        oTemp.Push oEle
    Next
    oRet.Add "All", oTemp
    
    'Named
    Set oTemp = new_Dictionary()
    For Each oKey In WScript.Arguments.Named
        oTemp.Add oKey, WScript.Arguments.Named.Item(oKey)
    Next
    oRet.Add "Named", oTemp
    
    'Unnamed
    Set oTemp = new_clsCmArray()
    For Each oEle In WScript.Arguments.Unnamed
        oTemp.Push oEle
    Next
    oRet.Add "Unnamed", oTemp
    
    Set func_CM_UtilStoringArguments = oRet
    
    Set oKey = Nothing
    Set oEle = Nothing
    Set oTemp = Nothing
    Set oRet = Nothing
End Function

