'***************************************************************************************************
'FILENAME                    : clsCmBroker.vbs
'Overview                    : �o��-�w�ǌ^���f���iPublish/subscribe�j�̃u���[�J�[
'Detailed Description        : �H����
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/12/02         Y.Fujii                  First edition
'***************************************************************************************************
Class clsCmBroker
    '�N���X���ϐ��A�萔
    Private PoTopics
    
    '***************************************************************************************************
    'Function/Sub Name           : Class_Initialize()
    'Overview                    : �R���X�g���N�^
    'Detailed Description        : �����ϐ��̏�����
    'Argument
    '     �Ȃ�
    'Return Value
    '     �Ȃ�
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
    'Overview                    : �f�X�g���N�^
    'Detailed Description        : �I������
    'Argument
    '     �Ȃ�
    'Return Value
    '     �Ȃ�
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
    Public Sub publish( _
        ByVal asTopic _
        , ByRef avArgs _
        )
        If Not PoTopics.Exists(asTopic) Then Exit Sub
        Call PoTopics.Item(asTopic)(avArgs)
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : subscribe()
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
    Public Sub subscribe( _
        ByVal asTopic _
        , ByRef aoCbFunc _
        )
        If PoTopics.Exists(asTopic) Then PoTopics.Remove asTopic
        PoTopics.Add asTopic, aoCbFunc
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : unsubscribe()
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
    Public Sub unsubscribe( _
        ByVal asTopic _
        )
        If PoTopics.Exists(asTopic) Then PoTopics.Remove asTopic
    End Sub
    
End Class
