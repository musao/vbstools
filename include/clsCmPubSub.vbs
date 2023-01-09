'***************************************************************************************************
'FILENAME                    : clsCmPubSub.vbs
'Overview                    : �o��-�w�ǌ^�iPublish/subscribe�j�������s���N���X
'Detailed Description        : �H����
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/12/02         Y.Fujii                  First edition
'***************************************************************************************************
Class clsCmPubSub
    '�N���X���ϐ��A�萔
    Private PoTopics
    
    '�R���X�g���N�^
    Private Sub Class_Initialize()
        Set PoTopics = CreateObject("Scripting.Dictionary")
    End Sub
    '�f�X�g���N�^
    Private Sub Class_Terminate()
        Set PoTopics = Nothing
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : Publish()
    'Overview                    : �o��
    'Detailed Description        : �H����
    'Argument
    '     asTopic                : �g�s�b�N
    '     avArgs                 : �R�[���o�b�N�֐��ɓn������
    'Return Value
    '     �Ȃ�
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
    'Overview                    : �w��
    'Detailed Description        : �H����
    'Argument
    '     asTopic                : �g�s�b�N
    '     aoCbFunc               : �R�[���o�b�N�֐��|�C���^
    'Return Value
    '     �Ȃ�
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
    'Overview                    : �w�ǉ���
    'Detailed Description        : �H����
    'Argument
    '     asTopic                : �g�s�b�N
    'Return Value
    '     �Ȃ�
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
