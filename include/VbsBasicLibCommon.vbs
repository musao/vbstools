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
    On Error Resume Next
    CreateObject("Scripting.FileSystemObject").DeleteFile(asPath)
    func_CM_FsDeleteFile = True
    If Err.Number Then
        Err.Clear
        func_CM_FsDeleteFile = False
    End If
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
'     ファイルオブジェクト
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
'Function/Sub Name           : func_CM_FsCreateFolder()
'Overview                    : ランダムに生成された一時ファイルまたはフォルダーの名前の取得
'Detailed Description        : FileSystemObjectのCreateFolder()と同等
'Argument
'     asPath                 : パス
'Return Value
'     作成したフォルダの絶対パス
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
'Function/Sub Name           : sub_CM_FsWriteFile()
'Overview                    : ファイル出力する
'Detailed Description        : エラーは無視する
'Argument
'     asPath                 : 出力先のフルパス
'     asCont                 : 出力する内容
'     なし
'Return Value
'     作成したフォルダの絶対パス
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
    With CreateObject("Scripting.FileSystemObject").OpenTextFile(asPath, 2, True)
        Call .WriteLine(asCont)
        Call .Close
    End With
    If Err.Number Then
        Err.Clear
    End If
End Sub


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
'Function/Sub Name           : func_CM_GetDateInMilliseconds()
'Overview                    : 日時をミリ秒で取得する
'Detailed Description        : 工事中
'Argument
'     なし
'Return Value
'     なし
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/10/12         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_GetDateInMilliseconds( _
    byVal adtDate _
    , byVal adtTimer _
    )
    Dim dtNowTime        '現在時刻
    Dim lHour            '時
    Dim lngMinute        '分
    Dim lngSecond        '秒
    Dim lngMilliSecond   'ミリ秒

    dtNowTime = adtTimer
    lMilliSecond = dtNowTime - Fix(dtNowTime)
    lMilliSecond = Right("000" & Fix(lMilliSecond * 1000), 3)
    dtNowTime = Fix(dtNowTime)
    lSecond = Right("0" & dtNowTime Mod 60, 2)
    dtNowTime = dtNowTime \ 60
    lMinute = Right("0" & dtNowTime Mod 60, 2)
    dtNowTime = dtNowTime \ 60
    lHour = Right("0" & dtNowTime, 2)

    func_CM_GetDateInMilliseconds = adtDate & " " & lHour & ":" & lMinute & ":" & lSecond & "." & lMilliSecond
End Function
