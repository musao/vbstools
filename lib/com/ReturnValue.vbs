'***************************************************************************************************
'FILENAME                    : ReturnValue.vbs
'Overview                    : �߂�l�N���X
'Detailed Description        : �H����
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2024/01/03         Y.Fujii                  First edition
'***************************************************************************************************
Class ReturnValue
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
        PvValue = Empty
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
    Public Function getErr( _
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
    Public Function isErr( _
        )
        isErr = PboIsErr
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : setValue()
    'Overview                    : �߂�l�̐ݒ�ƃG���[�������Err�I�u�W�F�N�g�̏����i�[����
    'Detailed Description        : this_setValue()�ɈϏ�����
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
    Public Function setValue( _
        byRef avRet _
        )
        Set setValue = this_setValue(avRet)
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : setValueByState()
    'Overview                    : ��Ԃɂ��߂�l�̐ݒ�ƃG���[�������Err�I�u�W�F�N�g�̏����i�[����
    'Detailed Description        : this_setValueByState()�ɈϏ�����
    'Argument
    '     avNormal               : ����̏ꍇ�̖߂�l
    '     avAbnormal             : �ُ�̏ꍇ�̖߂�l
    'Return Value
    '     ���g�̃C���X�^���X
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2024/04/27         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function setValueByState( _
        byRef avNormal _
        , byRef avAbnormal _
        )
        Set setValueByState = this_setValueByState(avNormal,avAbnormal)
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : toString()
    'Overview                    : �I�u�W�F�N�g�̓��e�𕶎���ŕ\������
    'Detailed Description        : func_CmReturnValueToString()�ɈϏ�����
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
        toString = func_CmReturnValueToString()
    End Function


    '***************************************************************************************************
    'Function/Sub Name           : this_setValue()
    'Overview                    : �߂�l�̐ݒ�ƃG���[�������Err�I�u�W�F�N�g�̏����i�[����
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
    Private Function this_setValue( _
        byRef avRet _
        )
        cf_bind PvValue, avRet
        this_getErrorStatus()
        Set this_setValue = Me
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : this_setValueByState()
    'Overview                    : ��Ԃɂ��߂�l�̐ݒ�ƃG���[�������Err�I�u�W�F�N�g�̏����i�[����
    'Detailed Description        : �H����
    'Argument
    '     avNormal               : ����̏ꍇ�̖߂�l
    '     avAbnormal             : �ُ�̏ꍇ�̖߂�l
    'Return Value
    '     ���g�̃C���X�^���X
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2024/04/27         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_setValueByState( _
        byRef avNormal _
        , byRef avAbnormal _
        )
        If Err.Number=0 Then
            cf_bind PvValue, avNormal
        Else
            cf_bind PvValue, avAbnormal
        End If
        this_getErrorStatus()
        Set this_setValueByState = Me
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : this_getErrorStatus()
    'Overview                    : �G���[��Ԃ��擾���G���[������ꍇ�͏����擾��ɃN���A����
    'Detailed Description        : �H����
    'Argument
    '     �Ȃ�
    'Return Value
    '     �Ȃ�
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2024/04/27         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub this_getErrorStatus( _
        )
        If Err.Number=0 Then
            PboIsErr = False
            Set PoErr = Nothing
        Else
            PboIsErr = True
            Set PoErr = fw_storeErr()
            Err.Clear
        End If
    End Sub
    '***************************************************************************************************
    'Function/Sub Name           : func_CmReturnValueToString()
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
    Private Function func_CmReturnValueToString( _
        )
        func_CmReturnValueToString = _
            "<" & TypeName(Me) & ">[" _
            & "returnValue:" & cf_toString(PvValue) _
            & ",isErr:" & cf_toString(PboIsErr) _
            & ",getErr:" & cf_toString(PoErr) _
            & "]"
    End Function

End Class
