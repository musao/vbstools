'***************************************************************************************************
'FILENAME                    :VbsBasicLibCommon.vbs
'Generato                    :2022/09/27
'Descrition                  :共通機能
' パラメータ（引数）:
'     PATH         :ファイルのパス
'---------------------------------------------------------------------------------------------------
'Modification Histroy
'
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/09/27         EXA Y.Fujii              Initial Release
'***************************************************************************************************

'オフィス全般

'文書の保護を解除する
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

'エクセルファイルを別名で保存して閉じる
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

'エクセルファイルを開いて（読み取り専用／ダイアログなし）ワークブックオブジェクトを返す
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

'エクセルのオートシェイプのテキストを取り出す
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

'ファイルを削除する
Private Function func_CM_DeleteFile( _
    byVal asPath _
    ) 
    On Error Resume Next
    CreateObject("Scripting.FileSystemObject").DeleteFile(asPath)
    func_CM_DeleteFile = True
    If Err.Number Then
        func_CM_DeleteFile = False
    End If
End Function

'親フォルダパスの取得
Private Function func_CM_GetParentFolderPath( _
    byVal asPath _
    ) 
    func_CM_GetParentFolderPath = CreateObject("Scripting.FileSystemObject").GetParentFolderName(asPath)
End Function

'ファイルパスの結合の作成
Private Function func_CM_BuildPath( _
    byVal asFolderPath _
    , byVal asItemName _
    ) 
    func_CM_BuildPath = CreateObject("Scripting.FileSystemObject").BuildPath(asFolderPath, asItemName)
End Function

'ファイルの存在確認
Private Function func_CM_FileExists( _
    byVal asPath _
    ) 
    func_CM_FileExists = CreateObject("Scripting.FileSystemObject").FileExists(asPath)
End Function

'ファイルオブジェクトの取得
Private Function func_CM_GetFile( _
    byVal asPath _
    ) 
    Set func_CM_GetFile = CreateObject("Scripting.FileSystemObject").GetFile(asPath)
End Function

'一時ファイル名の作成
Private Function func_CM_GetTempFileName()
    func_CM_GetTempFileName = CreateObject("Scripting.FileSystemObject").GetTempName()
End Function


'一般

'Min関数
Private Function func_CM_Min( _
    byVal al1 _ 
    , byVal al2 _
    )
    Dim lReturnValue
    If al1 < al2 Then
        lReturnValue = al1
    Else
        lReturnValue = al2
    End If
    func_CM_Min = lReturnValue
End Function

'Max関数
Private Function func_CM_Max( _
    byVal al1 _ 
    , byVal al2 _
    )
    Dim lReturnValue
    If al1 > al2 Then
        lReturnValue = al1
    Else
        lReturnValue = al2
    End If
    func_CM_Max = lReturnValue
End Function


'これ何系かな

'コレクションから指定した名前のメンバーを取得する
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
