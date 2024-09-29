'***************************************************************************************************
'FILENAME                    : clsCmReadOnlyObject.vbs
'Overview                    : �ǂݎ���p�I�u�W�F�N�g�̃N���X
'Detailed Description        : �H����
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2024/05/26         Y.Fujii                  First edition
'***************************************************************************************************
Class clsCmReadOnlyObject
    '�N���X���ϐ��A�萔
    Private PboAlreadySet, PoParent, PvValue, PsName
    
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
        Set PoParent = Nothing
        PvValue = Empty
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
        Set PoParent = Nothing
    End Sub

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
        name = this_getName()
    End Property

    '***************************************************************************************************
    'Function/Sub Name           : Property Get parent()
    'Overview                    : �e
    'Detailed Description        : �H����
    'Argument
    '     �Ȃ�
    'Return Value
    '     �e
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2024/05/26         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get parent()
        cf_bind parent, this_getParent()
    End Property

    '***************************************************************************************************
    'Function/Sub Name           : Property Get toString()
    'Overview                    : �C���X�^���X�̓��e�𕶎��o�͂���
    'Detailed Description        : this_toString()�ɈϏ�����
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
        toString = this_toString()
    End Property

    '***************************************************************************************************
    'Function/Sub Name           : Property Get value()
    'Overview                    : �l
    'Detailed Description        : �H����
    'Argument
    '     �Ȃ�
    'Return Value
    '     �l
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2024/05/26         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Default Property Get value()
        cf_bind value, this_getValue()
    End Property

    '***************************************************************************************************
    'Function/Sub Name           : compareTo()
    'Overview                    : ���N���X�̃C���X�^���X��value���r����
    'Detailed Description        : this_compareTo()�ɈϏ�����
    'Argument
    '     aoTarget               : ��r�Ώ�
    'Return Value
    '     ��r����
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2024/05/26         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function compareTo( _
        ByRef aoTarget _
        )
        Dim vRet : vRet = this_compareTo(aoTarget)
        ast_argNotNull vRet, TypeName(Me)&"+compareTo()", "The type of the argument is different."
        compareTo = vRet
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : equals()
    'Overview                    : �w�肳�ꂽ�I�u�W�F�N�g������enum�萔�Ɠ����ꍇ��true��Ԃ��B
    'Detailed Description        : �H����
    'Argument
    '     aoTarget               : ���N���X�̃C���X�^���X
    'Return Value
    '     ���� True:��v / False:�s��v
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2024/05/26         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function equals( _
        ByRef aoTarget _
        )
        equals = (this_compareTo(aoTarget)=0)
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : is()
    'Overview                    : �l��ݒ肷��
    'Detailed Description        : ���ɐݒ�ς݂̏ꍇ�͗�O
    'Argument
    '     aoParent               : �e�̃I�u�W�F�N�g
    '     asName                 : ���O
    '     avValue                : �l
    'Return Value
    '     ���g�̃C���X�^���X
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2024/05/26         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function [is]( _
        ByRef aoParent _
        , ByVal asName _
        , ByRef avValue _
        )
        [is] = Empty
        ast_argFalse PboAlreadySet, TypeName(Me)&"+is()", "Value already set."

        this_setParent aoParent
        this_setName asName
        this_setValue avValue
        PboAlreadySet = True
        Set [is] = Me
    End Function
    

    '***************************************************************************************************
    'Function/Sub Name           : this_compareTo()
    'Overview                    : ���N���X�̃C���X�^���X��value���r����
    'Detailed Description        : ���L��r���ʂ�Ԃ�
    '                               0  �����Ɠ��l
    '                               -1 ������菬����
    '                               1  �������傫��
    'Argument
    '     aoTarget              : ��r�Ώ�
    'Return Value
    '     ��r����
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2024/05/26         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_compareTo( _
        ByRef aoTarget _
        )
        this_compareTo = Null
        If Not cf_isSame(TypeName(Me), TypeName(aoTarget)) Then Exit Function
        If Not cf_isSame(PoParent, aoTarget.parent) Then Exit Function

        Dim lResult : lResult = 0
        If (PvValue < aoTarget.value) Then lResult = -1
        If (PvValue > aoTarget.value) Then lResult = 1
        this_compareTo = lResult
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : this_getValue()
    'Overview                    : PvValue�̃Q�b�^�[
    'Detailed Description        : �H����
    'Argument
    'Return Value
    '     �l
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2024/05/26         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_getValue( _
        )
        cf_bind this_getValue, PvValue
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : this_getName()
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
    Private Function this_getName( _
        )
        this_getName = PsName
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : this_getParent()
    'Overview                    : PvParent�̃Q�b�^�[
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
    Private Function this_getParent( _
        )
        cf_bind this_getParent, PoParent
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : this_setName()
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
    Private Sub this_setName( _
        ByVal asName _
        )
        PsName = asName
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : this_setParent()
    'Overview                    : PvParent�̃Z�b�^�[
    'Detailed Description        : �H����
    'Argument
    '     aoParent               : �e
    'Return Value
    '     �Ȃ�
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2024/05/26         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub this_setParent( _
        ByVal aoParent _
        )
        cf_bind PoParent, aoParent
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : this_setValue()
    'Overview                    : PvValue�̃Z�b�^�[
    'Detailed Description        : �H����
    'Argument
    '     avValue                : �l
    'Return Value
    '     �Ȃ�
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2024/05/26         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub this_setValue( _
        ByRef avValue _
        )
        cf_bind PvValue, avValue
    End Sub

    '***************************************************************************************************
    'Function/Sub Name           : this_toString()
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
    Private Function this_toString( _
        )
        this_toString = "<" & TypeName(Me) & ">{" & cf_toString(PsName) & ":" & cf_toString(PvValue) & "}"
    End Function
    
End Class
