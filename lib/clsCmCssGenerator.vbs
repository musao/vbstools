'***************************************************************************************************
'FILENAME                    : clsCmCssGenerator.vbs
'Overview                    : CSS�𐶐�����
'Detailed Description        : �H����
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/10/25         Y.Fujii                  First edition
'***************************************************************************************************
Class clsCmCssGenerator
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
        Set PoTagInfo = new_DicWith(Array("selector", Empty, "property", Empty))
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
        property = PoTagInfo.Item("property").Items()
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
    '     �Ȃ�
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/25         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Sub addProperty( _
        byVal asKey _
        , byVal asValue _
        )
        Dim oNewAttr : Set oNewAttr = new_DicWith(Array("key", asKey, "value", asValue))
        If IsEmpty(PoTagInfo.Item("property")) Then
            Set PoTagInfo.Item("property") = new_ArrWith(oNewAttr)
        Else
            PoTagInfo.Item("property").push oNewAttr
        End If
        Set oNewAttr = Nothing
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : generate()
    'Overview                    : CSS�𐶐�����
    'Detailed Description        : func_CmCssGenGenerate()�ɈϏ�����
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
        generate = func_CmCssGenGenerate()
    End Function




    '***************************************************************************************************
    'Function/Sub Name           : func_CmCssGenGenerate()
    'Overview                    : CSS�𐶐�����
    'Detailed Description        : �H����
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
    Private Function func_CmCssGenGenerate( _
        )
        If IsEmpty(PoTagInfo.Item("selector")) Then
            Err.Raise 17, "clsCmCssGenerator.vbs:clsCmCssGenerator-func_CmCssGenGenerate()", "�Z���N�^���Ȃ�CSS�͐����ł��܂���B"
            Exit Function
        End If

        Dim sRet : sRet = PoTagInfo.Item("selector") & " {" & vbNewLine

        '�v���p�e�B�iproperty�j�̕ҏW
        Dim vNewArr, vArr
        If Not IsEmpty(PoTagInfo.Item("property")) Then
        'property����łȂ��ꍇ
            Set vNewArr = new_Arr()
            Set vArr = PoTagInfo.Item("property").slice(0,vbNullString)
            Do While vArr.length>0
                vNewArr.push func_CmCssGenEditProperty(vArr.shift)
            Loop
            sRet = sRet & vNewArr.join(vbNewLine)
        End If
        
        sRet = sRet & "}"

        '��������CSS��ԋp
        func_CmCssGenGenerate = sRet

        Set vNewArr = Nothing
        Set vArr = Nothing
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : func_CmCssGenEditProperty()
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
    Private Function func_CmCssGenEditProperty( _
        byRef aoAttr _
        )
        func_CmCssGenEditProperty = aoAttr.Item("key") & " : " & aoAttr.Item("value") & " ;"
    End Function

End Class