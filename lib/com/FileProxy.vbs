'***************************************************************************************************
'FILENAME                    : FileProxy.vbs
'Overview                    : Fileオブジェクトのプロキシクラス
'Detailed Description        : Fileオブジェクトと同じIFを提供する
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/11/26         Y.Fujii                  First edition
'***************************************************************************************************
Class FileProxy
    'クラス内変数、定数
    Private PoFile,PsPath
    
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
        PsPath = vbNullString
        Set PoFile = Nothing
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
    'Function/Sub Name           : Property Get dateLastModified()
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
    Public Property Get dateLastModified()
        dateLastModified = this_getDateLastModified()
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get name()
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
    Public Property Get name()
        name = this_getName()
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get parentFolder()
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
    Public Property Get parentFolder()
        parentFolder = this_getParentFolder()
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get path()
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
    Public Default Property Get path()
        path = this_getPath()
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get size()
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
    Public Property Get size()
        size = this_getSize()
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
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/03/16         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get toString()
        toString = "<"&TypeName(Me)&">"&this_getPath()
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get type()
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
    Public Property Get [type]()
        [type] = this_getType()
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : of()
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
    Public Function of( _
        byVal asPath _
        )
        this_setData asPath, TypeName(Me)&"+of()"
        Set of = Me
    End Function
    

   
    '***************************************************************************************************
    'Function/Sub Name           : this_getDateLastModified()
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
    '2025/03/15         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_getDateLastModified()
        this_getDateLastModified = Null
        If Not(PoFile Is Nothing) Then this_getDateLastModified = PoFile.ModifyDate
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : this_getName()
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
    '2025/03/15         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_getName()
        this_getName = Null
        If Not(PoFile Is Nothing) Then this_getName = new_Fso().GetFileName(PsPath)
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : this_getParentFolder()
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
    '2025/03/15         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_getParentFolder()
        this_getParentFolder = Null
        If Not(PoFile Is Nothing) Then this_getParentFolder = new_Fso().GetParentFolderName(PsPath)
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : this_getPath()
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
    '2025/03/15         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_getPath()
        this_getPath = Null
        If Not(PoFile Is Nothing) Then this_getPath = PsPath
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : this_getSize()
    'Overview                    : ファイルのサイズを返す
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     ファイルのサイズ（バイト単位）
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/03/15         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_getSize()
        this_getSize = Null
        If Not(PoFile Is Nothing) Then this_getSize = PoFile.Size
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : this_getType()
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
    '2025/03/15         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_getType()
        this_getType = Null
        If Not(PoFile Is Nothing) Then this_getType = PoFile.Type
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : this_setData()
    'Overview                    : データを設定する
    'Detailed Description        : 工事中
    'Argument
    '     asPath                 : ファイルのパス
    'Return Value
    '     なし
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/03/16         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub this_setData( _
        byVal asPath _
        , byVal asSource _
        )
        ast_argNothing PoFile, asSource, "Because it is an immutable variable, its value cannot be changed."

        Dim oFile : Set oFile = Nothing
        On Error Resume Next
        Set oFile = new_FolderItem2Of(asPath)
        ast_argNotNothing PoFile, asSource, "invalid argument. " & cf_toString(asPath)

        If oFile Is Nothing Then Exit Sub

        this_setFile oFile, asSource
        this_setPath asPath, asSource
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : this_setFile()
    'Overview                    : ファイルオブジェクト（FolderItem2）を設定する
    'Detailed Description        : 工事中
    'Argument
    '     aoFile                 : ファイルオブジェクト
    '     asSource               : ソース
    'Return Value
    '     なし
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/03/16         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub this_setFile( _
        byRef aoFile _
        , byVal asSource _
        )
        ast_argsAreSame "FolderItem2", TypeName(aoFile), asSource, "This is not FolderItem2."
        Set PoFile = aoFile
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : this_setPath()
    'Overview                    : ファイルのパスを設定する
    'Detailed Description        : 工事中
    'Argument
    '     asPath                 : ファイルのパス
    'Return Value
    '     なし
    '---------------------------------------------------------------------------------------------------
    'Histroy
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
