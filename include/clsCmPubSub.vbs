'***************************************************************************************************
'FILENAME                    : clsCmPubSub.vbs
'Overview                    : 出版-購読型（Publish/subscribe）処理を行うクラス
'Detailed Description        : 工事中
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/12/02         Y.Fujii                  First edition
'***************************************************************************************************
Class clsCmPubSub
    'クラス内変数、定数
    Private PoTopics
    
    'コンストラクタ
    Private Sub Class_Initialize()
        Set PoTopics = CreateObject("Scripting.Dictionary")
    End Sub
    'デストラクタ
    Private Sub Class_Terminate()
        Set PoTopics = Nothing
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : Publish()
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
    Public Sub Publish( _
        ByVal asTopic _
        , ByRef avArgs _
        )
        If Not PoTopics.Exists(asTopic) Then Exit Sub
        Call PoTopics.Item(asTopic)(avArgs)
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : Subscribe()
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
    Public Sub Subscribe( _
        ByVal asTopic _
        , ByRef aoCbFunc _
        )
        If PoTopics.Exists(asTopic) Then PoTopics.Remove asTopic
        PoTopics.Add asTopic, aoCbFunc
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : Subscribe()
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
    Public Sub Unsubscribe( _
        ByVal asTopic _
        )
        If PoTopics.Exists(asTopic) Then PoTopics.Remove asTopic
    End Sub
    
End Class
