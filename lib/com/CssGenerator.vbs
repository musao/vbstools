'***************************************************************************************************
'FILENAME                    : CssGenerator.vbs
'Overview                    : CSS�����N���X
'Detailed Description        : �H����
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/10/25         Y.Fujii                  First edition
'***************************************************************************************************
Class CssGenerator
    '�N���X���ϐ��A�萔
    Private PoTagInfo
    
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
    '2023/10/25         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub Class_Initialize()
        Set PoTagInfo = new_DicOf(Array("selector", Empty, "property", Empty))
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
    '2023/10/25         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub Class_Terminate()
        Set PoTagInfo = Nothing
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get property()
    'Overview                    : �v���p�e�B�i�I�u�W�F�N�g�̔z��j��Ԃ�
    'Detailed Description        : �H����
    'Argument
    '     �Ȃ�
    'Return Value
    '     �v���p�e�B�i�I�u�W�F�N�g�̔z��j��Ԃ�
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/25         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get property()
        property = PoTagInfo.Item("property")
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Let selector()
    'Overview                    : �Z���N�^��ݒ肷��
    'Detailed Description        : �H����
    'Argument
    '     asSelector             : �Z���N�^
    'Return Value
    '     �Ȃ�
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/25         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Let selector( _
        byVal asSelector _
        )
        PoTagInfo.Item("selector") = asSelector
'        If new_Re("^[!-~][ -~]*$", "i").Test(asSelector) Then
'            PoTagInfo.Item("selector") = asSelector
'        Else
'            Err.Raise 1032, "CssGenerator.vbs:CssGenerator+selector()", "�Z���N�^�ɂ͔��p�ȊO�̕������w��ł��܂���B"
'        End If
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get selector()
    'Overview                    : �Z���N�^��Ԃ�
    'Detailed Description        : �H����
    'Argument
    '     �Ȃ�
    'Return Value
    '     �Z���N�^
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/25         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get selector()
        selector = PoTagInfo.Item("selector")
    End Property
        
    '***************************************************************************************************
    'Function/Sub Name           : addProperty()
    'Overview                    : �v���p�e�B��ǉ�����
    'Detailed Description        : �H����
    'Argument
    '     asKey                  : �ǉ�����v���p�e�B�̃L�[
    '     asValue                : �ǉ�����v���p�e�B�̒l
    'Return Value
    '     ���g�̃C���X�^���X
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/25         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function addProperty( _
        byVal asKey _
        , byVal asValue _
        )
        Dim oNewAttr : Set oNewAttr = new_DicOf(Array("key", asKey, "value", asValue))
        Dim vArr : cf_bind vArr, PoTagInfo.Item("property")
        cf_push vArr, oNewAttr
        cf_bindAt PoTagInfo, "property", vArr

        Set addProperty = Me
        Set oNewAttr = Nothing
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : generate()
    'Overview                    : CSS�𐶐�����
    'Detailed Description        : this_generate()�ɈϏ�����
    'Argument
    '     �Ȃ�
    'Return Value
    '     ��������CSS
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/25         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function generate( _
        )
        generate = this_generate(TypeName(Me)&"+generate()")
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
    '2023/12/27         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function toString( _
        )
        toString = this_generate(TypeName(Me)&"+toString()")
    End Function




    '***************************************************************************************************
    'Function/Sub Name           : this_generate()
    'Overview                    : CSS�𐶐�����
    'Detailed Description        : �H����
    'Argument
    '     asSource               : �\�[�X
    'Return Value
    '     ��������CSS
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/25         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_generate( _
        byVal asSource _
        )
        ast_argNotEmpty PoTagInfo.Item("selector"), asSource, "CSS without selectors cannot be generated."

        Dim sRet : sRet = PoTagInfo.Item("selector") & " {" & vbNewLine

        '�v���p�e�B�iproperty�j�̕ҏW
        Dim vArr, vEle
        If Not IsEmpty(PoTagInfo.Item("property")) Then
        'property����łȂ��ꍇ
            For Each vEle In PoTagInfo.Item("property")
                cf_push vArr, "  " & this_editProperty(vEle)
            Next
            sRet = sRet & Join(vArr, vbNewLine) & vbNewLine
        End If
        
        sRet = sRet & "}"

        '��������CSS��ԋp
        this_generate = sRet
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : this_editProperty()
    'Overview                    : �v���p�e�B�iproperty�j�̕ҏW����
    'Detailed Description        : �H����
    'Argument
    '     aoAttr                 : �ҏW����v���p�e�B�iproperty�j
    'Return Value
    '     �ҏW����
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/25         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_editProperty( _
        byRef aoAttr _
        )
        this_editProperty = aoAttr.Item("key") & " : " & aoAttr.Item("value") & " ;"
    End Function

End Class
