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
        name = func_CmEnumEleGetName()
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
        cf_bind parent, func_CmEnumEleGetParent()
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
        cf_bind value, func_CmEnumEleGetValue()
    End Property

    '***************************************************************************************************
    'Function/Sub Name           : compareTo()
    'Overview                    : ���N���X�̃C���X�^���X��value���r����
    'Detailed Description        : func_CmEnumEleCompareTo()�ɈϏ�����
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
        Dim vRet : vRet = func_CmEnumEleCompareTo(aoTarget)
        ast_argNotNull vRet, TypeName(Me)&"+compareTo()", "The type of the argument is different"
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
        equals = (func_CmEnumEleCompareTo(aoTarget)=0)
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : thisIs()
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
    Public Function thisIs( _
        ByRef aoParent _
        , ByVal asName _
        , ByRef avValue _
        )
        thisIs = Empty
        ast_argFalse PboAlreadySet, TypeName(Me)&"+thisIs()", "Value already set"

        sub_CmEnumEleSetParent aoParent
        sub_CmEnumEleSetName asName
        sub_CmEnumEleSetValue avValue
        PboAlreadySet = True
        Set thisIs = Me
    End Function
    

    '***************************************************************************************************
    'Function/Sub Name           : func_CmEnumEleCompareTo()
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
    Private Function func_CmEnumEleCompareTo( _
        ByRef aoTarget _
        )
        func_CmEnumEleCompareTo = Null
        If Not cf_isSame(TypeName(Me), TypeName(aoTarget)) Then Exit Function
        If Not cf_isSame(PoParent, aoTarget.parent) Then Exit Function

        Dim lResult : lResult = 0
        If (PvValue < aoTarget.value) Then lResult = -1
        If (PvValue > aoTarget.value) Then lResult = 1
        func_CmEnumEleCompareTo = lResult
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : func_CmEnumEleGetValue()
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
    Private Function func_CmEnumEleGetValue( _
        )
        cf_bind func_CmEnumEleGetValue, PvValue
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
    'Function/Sub Name           : func_CmEnumEleGetParent()
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
    Private Function func_CmEnumEleGetParent( _
        )
        cf_bind func_CmEnumEleGetParent, PoParent
    End Function
    
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
    'Function/Sub Name           : sub_CmEnumEleSetParent()
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
    Private Sub sub_CmEnumEleSetParent( _
        ByVal aoParent _
        )
        cf_bind PoParent, aoParent
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : sub_CmEnumEleSetValue()
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
    Private Sub sub_CmEnumEleSetValue( _
        ByRef avValue _
        )
        cf_bind PvValue, avValue
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
        func_CmEnumEleToString = "<" & TypeName(Me) & ">(" & cf_toString(PvValue) & ":" & cf_toString(PsName) & " of " & cf_toString(PoParent) & ")"
    End Function
    
End Class
