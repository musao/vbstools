'***************************************************************************************************
'FILENAME                    : clsCmEnumElement.vbs
'Overview                    : Enum�̗v�f�N���X
'Detailed Description        : �H����
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2024/05/26         Y.Fujii                  First edition
'***************************************************************************************************
Class clsCmEnumElement
    '�N���X���ϐ��A�萔
    Private PboAlreadySet, PsKind, PvCode, PsName
    
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
    '2024/05/26         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub Class_Initialize()
        PboAlreadySet = False
        PsKind = Empty
        PvCode = Empty
        PsName = Empty
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
    '2024/05/26         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub Class_Terminate()
    End Sub

    '***************************************************************************************************
    'Function/Sub Name           : Property Get code()
    'Overview                    : �R�[�h
    'Detailed Description        : �H����
    'Argument
    '     �Ȃ�
    'Return Value
    '     �R�[�h
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2024/05/26         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Default Property Get code()
        code = func_CmEnumEleGetCode()
    End Property

    '***************************************************************************************************
    'Function/Sub Name           : Property Get kind()
    'Overview                    : ���
    'Detailed Description        : �H����
    'Argument
    '     �Ȃ�
    'Return Value
    '     ���
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2024/05/26         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get kind()
        kind = func_CmEnumEleGetKind()
    End Property

    '***************************************************************************************************
    'Function/Sub Name           : Property Get name()
    'Overview                    : ���O
    'Detailed Description        : �H����
    'Argument
    '     �Ȃ�
    'Return Value
    '     ���O
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2024/05/27         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get name()
        name = func_CmEnumEleGetName()
    End Property

    '***************************************************************************************************
    'Function/Sub Name           : Property Get toString()
    'Overview                    : �C���X�^���X�̓��e�𕶎��o�͂���
    'Detailed Description        : func_CmEnumEleToString()�ɈϏ�����
    'Argument
    '     �Ȃ�
    'Return Value
    '     �C���X�^���X�̓��e
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2024/05/26         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get toString()
        toString = func_CmEnumEleToString()
    End Property

    '***************************************************************************************************
    'Function/Sub Name           : compareTo()
    'Overview                    : ���N���X�̃C���X�^���X��code���r����
    'Detailed Description        : func_CmEnumEleCompareTo()�ɈϏ�����
    'Argument
    '     aoEnumEle              : ��r�Ώ�
    'Return Value
    '     ��r����
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2024/05/26         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function compareTo( _
        ByRef aoEnumEle _
        )
        Dim vRet : vRet = func_CmEnumEleCompareTo(aoEnumEle)
        ast_argNotNull vRet, TypeName(Me)&"+compareTo()", "The type of the argument is different"
        compareTo = vRet
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : equals()
    'Overview                    : �w�肳�ꂽ�I�u�W�F�N�g������enum�萔�Ɠ����ꍇ��true��Ԃ��B
    'Detailed Description        : �H����
    'Argument
    '     aoEnumEle              : ���N���X�̃C���X�^���X
    'Return Value
    '     ���� True:��v / False:�s��v
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2024/05/26         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function equals( _
        ByRef aoEnumEle _
        )
        equals = (func_CmEnumEleCompareTo(aoEnumEle)=0)
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : thisIs()
    'Overview                    : �v�f��ݒ肷��
    'Detailed Description        : �H����
    'Argument
    '     asKind                 : ���
    '     asName                 : ���O
    '     avCode                 : �R�[�h
    'Return Value
    '     ���g�̃C���X�^���X
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2024/05/26         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function thisIs( _
        ByVal asKind _
        , ByVal asName _
        , ByRef avCode _
        )
        thisIs = Empty
        ast_argFalse PboAlreadySet, TypeName(Me)&"+thisIs()", "Value already set"

        sub_CmEnumEleSetKind asKind
        sub_CmEnumEleSetCode avCode
        sub_CmEnumEleSetName asName
        PboAlreadySet = True
        Set thisIs = Me
    End Function
    

    '***************************************************************************************************
    'Function/Sub Name           : func_CmEnumEleCompareTo()
    'Overview                    : ���N���X�̃C���X�^���X��code���r����
    'Detailed Description        : ���L��r���ʂ�Ԃ�
    '                               0  �����Ɠ��l
    '                               -1 ������菬����
    '                               1  �������傫��
    'Argument
    '     aoEnumEle              : ��r�Ώ�
    'Return Value
    '     ��r����
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2024/05/26         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function func_CmEnumEleCompareTo( _
        ByRef aoEnumEle _
        )
        func_CmEnumEleCompareTo = Null
        If Not cf_isSame(TypeName(Me), TypeName(aoEnumEle)) Then Exit Function
        If Not cf_isSame(PsKind, aoEnumEle.kind) Then Exit Function

        Dim lResult : lResult = 0
        If (PvCode < aoEnumEle.code) Then lResult = -1
        If (PvCode > aoEnumEle.code) Then lResult = 1
        func_CmEnumEleCompareTo = lResult
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : func_CmEnumEleGetCode()
    'Overview                    : PvCode�̃Q�b�^�[
    'Detailed Description        : �H����
    'Argument
    'Return Value
    '     �R�[�h
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2024/05/26         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function func_CmEnumEleGetCode( _
        )
        cf_bind func_CmEnumEleGetCode, PvCode
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : func_CmEnumEleGetKind()
    'Overview                    : PvKind�̃Q�b�^�[
    'Detailed Description        : �H����
    'Argument
    'Return Value
    '     ���
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2024/05/26         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function func_CmEnumEleGetKind( _
        )
        func_CmEnumEleGetKind = PsKind
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : func_CmEnumEleGetName()
    'Overview                    : PsName�̃Q�b�^�[
    'Detailed Description        : �H����
    'Argument
    'Return Value
    '     ���O
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2024/05/27         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function func_CmEnumEleGetName( _
        )
        func_CmEnumEleGetName = PsName
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : sub_CmEnumEleSetCode()
    'Overview                    : PvCode�̃Z�b�^�[
    'Detailed Description        : �H����
    'Argument
    '     avCode                 : �R�[�h
    'Return Value
    '     �Ȃ�
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2024/05/26         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub sub_CmEnumEleSetCode( _
        ByVal avCode _
        )
        cf_bind PvCode, avCode
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : sub_CmEnumEleSetKind()
    'Overview                    : PvKind�̃Z�b�^�[
    'Detailed Description        : �H����
    'Argument
    '     asKind                 : ���
    'Return Value
    '     �Ȃ�
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2024/05/26         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub sub_CmEnumEleSetKind( _
        ByVal asKind _
        )
        PsKind = asKind
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : sub_CmEnumEleSetName()
    'Overview                    : PsName�̃Z�b�^�[
    'Detailed Description        : �H����
    'Argument
    '     asName                 : ���O
    'Return Value
    '     �Ȃ�
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2024/05/27         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub sub_CmEnumEleSetName( _
        ByVal asName _
        )
        PsName = asName
    End Sub

    '***************************************************************************************************
    'Function/Sub Name           : func_CmEnumEleToString()
    'Overview                    : �C���X�^���X�̓��e�𕶎��o�͂���
    'Detailed Description        : �H����
    'Argument
    '     �Ȃ�
    'Return Value
    '     �C���X�^���X�̓��e
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2024/05/26         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function func_CmEnumEleToString( _
        )
        func_CmEnumEleToString = "<" & TypeName(Me) & ">(" & cf_toString(PvCode) & ":" & cf_toString(PsName) & " of " & cf_toString(PsKind) & ")"
    End Function
    
End Class
