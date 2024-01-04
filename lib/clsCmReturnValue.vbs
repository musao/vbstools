'***************************************************************************************************
'FILENAME                    : clsCmReturnValue.vbs
'Overview                    : �߂�l�N���X
'Detailed Description        : �H����
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2024/01/03         Y.Fujii                  First edition
'***************************************************************************************************
Class clsCmReturnValue
    '�N���X���ϐ��A�萔
    Private PvValue
    Private PoErr
    Private PboIsErr
    
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
    '2024/01/03         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub Class_Initialize()
        Set PvValue = Nothing
        Set PoErr = Nothing
        PboIsErr = Empty
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
    '2024/01/03         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub Class_Terminate()
        Set PvValue = Nothing
        Set PoErr = Nothing
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get returnValue()
    'Overview                    : �߂�l��Ԃ�
    'Detailed Description        : �H����
    'Argument
    '     �Ȃ�
    'Return Value
    '     �߂�l
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2024/01/03         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Default Property Get returnValue()
        cf_bind returnValue, PvValue
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Let returnValue()
    'Overview                    : �߂�l��ݒ肷��
    'Detailed Description        : �H����
    'Argument
    '     avRet                  : �߂�l
    'Return Value
    '     �Ȃ�
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2024/01/03         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Let returnValue( _
        byRef avRet _
        )
        cf_bind PvValue, avRet
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Set returnValue()
    'Overview                    : �߂�l��ݒ肷��
    'Detailed Description        : �H����
    'Argument
    '     avRet                  : �߂�l
    'Return Value
    '     �Ȃ�
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2024/01/03         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Set returnValue( _
        byRef avRet _
        )
        cf_bind PvValue, avRet
    End Property

    '***************************************************************************************************
    'Function/Sub Name           : getErr()
    'Overview                    : �G���[����ԋp����
    'Detailed Description        : �H����
    'Argument
    '     �Ȃ�
    'Return Value
    '     �G���[���
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2024/01/03         Y.Fujii                  First edition
    '***************************************************************************************************
    Function getErr( _
        )
        Set getErr = PoErr
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : isErr()
    'Overview                    : �G���[���̗L����ԋp����
    'Detailed Description        : �H����
    'Argument
    '     �Ȃ�
    'Return Value
    '     ���� True:�G���[���� / False:�G���[�Ȃ�
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2024/01/03         Y.Fujii                  First edition
    '***************************************************************************************************
    Function isErr( _
        )
        isErr = PboIsErr
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : setValue()
    'Overview                    : �߂�l�̎擾�ƃG���[�������Err�I�u�W�F�N�g�̏����i�[����
    'Detailed Description        : �H����
    'Argument
    '     avRet                  : �߂�l
    'Return Value
    '     ���g�̃C���X�^���X
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2024/01/03         Y.Fujii                  First edition
    '***************************************************************************************************
    Function setValue( _
        byRef avRet _
        )
        cf_bind PvValue, avRet
        If Err.Number=0 Then
            PboIsErr = False
            Set PoErr = Nothing
        Else
            PboIsErr = True
            Set PoErr = fw_storeErr()
        End If
        Set setValue = Me
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : toString()
    'Overview                    : �I�u�W�F�N�g�̓��e�𕶎���ŕ\������
    'Detailed Description        : cf_toString()����
    'Argument
    '     �Ȃ�
    'Return Value
    '     ������ɕϊ������I�u�W�F�N�g�̓��e
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2024/01/04         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function toString( _
        )
        toString = _
            "<" & TypeName(Me) & ">[" _
            & "returnValue:" & cf_toString(PvValue) _
            & ",isErr:" & cf_toString(PboIsErr) _
            & ",getErr:" & cf_toString(PoErr) _
            & "]"
    End Function

End Class
