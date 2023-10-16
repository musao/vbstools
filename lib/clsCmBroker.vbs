'***************************************************************************************************
'FILENAME                    : clsCmBroker.vbs
'Overview                    : 出版-購読型モデル（Publish/subscribe）のブローカー
'Detailed Description        : 工事中
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/12/02         Y.Fujii                  First edition
'***************************************************************************************************
Class clsCmBroker
    'クラス内変数、定数
    Private PoTopics
    
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
    '2023/09/02         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub Class_Initialize()
        Set PoTopics = new_Dic()
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
    '2023/09/02         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub Class_Terminate()
        Set PoTopics = Nothing
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : publish()
    'Overview                    : 出版
    'Detailed Description        : 工事中
    'Argument
    '     asTopic                : トピック
    '     avArgs                 : コールバック関数に渡す引数
    'Return Value
    '     なし
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2022/12/02         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Sub publish( _
        ByVal asTopic _
        , ByRef avArgs _
        )
        If Not PoTopics.Exists(asTopic) Then Exit Sub
        Call PoTopics.Item(asTopic)(avArgs)
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : subscribe()
    'Overview                    : 購読
    'Detailed Description        : 工事中
    'Argument
    '     asTopic                : トピック
    '     aoCbFunc               : コールバック関数ポインタ
    'Return Value
    '     なし
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2022/12/02         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Sub subscribe( _
        ByVal asTopic _
        , ByRef aoCbFunc _
        )
        If PoTopics.Exists(asTopic) Then PoTopics.Remove asTopic
        PoTopics.Add asTopic, aoCbFunc
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : unsubscribe()
    'Overview                    : 購読解除
    'Detailed Description        : 工事中
    'Argument
    '     asTopic                : トピック
    'Return Value
    '     なし
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2022/12/02         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Sub unsubscribe( _
        ByVal asTopic _
        )
        If PoTopics.Exists(asTopic) Then PoTopics.Remove asTopic
    End Sub
    
End Class
