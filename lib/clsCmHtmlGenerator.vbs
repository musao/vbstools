'***************************************************************************************************
'FILENAME                    : clsCmHtmlGenerator.vbs
'Overview                    : HTML�𐶐�����
'Detailed Description        : �H����
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/10/22         Y.Fujii                  First edition
'***************************************************************************************************
Class clsCmHtmlGenerator
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
    '2023/10/22         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub Class_Initialize()
        Set PoTagInfo = new_DicWith(Array("element", Empty, "attribute", Empty, "content", Empty))
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
    '2023/10/22         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub Class_Terminate()
        Set PoTagInfo = Nothing
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get attribute()
    'Overview                    : �����i�I�u�W�F�N�g�̔z��j��Ԃ�
    'Detailed Description        : �H����
    'Argument
    '     �Ȃ�
    'Return Value
    '     �����i�I�u�W�F�N�g�̔z��j��Ԃ�
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/23         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get attribute()
        attribute = PoTagInfo.Item("attribute").Items()
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get content()
    'Overview                    : ���e�i�I�u�W�F�N�g�̔z��j��Ԃ�
    'Detailed Description        : �H����
    'Argument
    '     �Ȃ�
    'Return Value
    '     ���e�i�I�u�W�F�N�g�̔z��j��Ԃ�
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/23         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get content()
        content = PoTagInfo.Item("content").Items()
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Let element()
    'Overview                    : �v�f��ݒ肷��
    'Detailed Description        : �H����
    'Argument
    '     asElement              : �v�f
    'Return Value
    '     �Ȃ�
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/23         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Let element( _
        byVal asElement _
        )
        PoTagInfo.Item("element") = asElement
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get element()
    'Overview                    : �v�f��Ԃ�
    'Detailed Description        : �H����
    'Argument
    '     �Ȃ�
    'Return Value
    '     �v�f
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/23         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get element()
        element = PoTagInfo.Item("element")
    End Property
        
    '***************************************************************************************************
    'Function/Sub Name           : addContent()
    'Overview                    : ������ǉ�����
    'Detailed Description        : �H����
    'Argument
    '     avCont                 : �ǉ�������e
    'Return Value
    '     �Ȃ�
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/23         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Sub addContent( _
        byRef avCont _
        )
        If IsEmpty(PoTagInfo.Item("content")) Then
            Set PoTagInfo.Item("content") = new_ArrWith(avCont)
        Else
            PoTagInfo.Item("content").push avCont
        End If
        Set oNewAttr = Nothing
    End Sub
        
    '***************************************************************************************************
    'Function/Sub Name           : addAttribute()
    'Overview                    : ������ǉ�����
    'Detailed Description        : �H����
    'Argument
    '     asKey                  : �ǉ����鑮���̃L�[
    '     asValue                : �ǉ����鑮���̒l
    'Return Value
    '     �Ȃ�
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/23         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Sub addAttribute( _
        byVal asKey _
        , byVal asValue _
        )
        Dim oNewAttr : Set oNewAttr = new_DicWith(Array("key", asKey, "value", asValue))
        If IsEmpty(PoTagInfo.Item("attribute")) Then
            Set PoTagInfo.Item("attribute") = new_ArrWith(oNewAttr)
        Else
            PoTagInfo.Item("attribute").push oNewAttr
        End If
        Set oNewAttr = Nothing
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : generate()
    'Overview                    : HTML�𐶐�����
    'Detailed Description        : func_CmHtmlGenGenerate()�ɈϏ�����
    'Argument
    '     �Ȃ�
    'Return Value
    '     ��������HTML
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/22         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function generate( _
        )
        generate = func_CmHtmlGenGenerate()
    End Function




    '***************************************************************************************************
    'Function/Sub Name           : func_CmHtmlGenGenerate()
    'Overview                    : HTML�𐶐�����
    'Detailed Description        : �H����
    'Argument
    '     �Ȃ�
    'Return Value
    '     ��������HTML
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/22         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function func_CmHtmlGenGenerate( _
        )
        If IsEmpty(PoTagInfo.Item("element")) Then
            Err.Raise 17, "clsCmHtmlGenerator.vbs:clsCmHtmlGenerator-func_CmHtmlGenGenerate()", "�v�f���Ȃ�HTML�^�O�͐����ł��܂���B"
            Exit Function
        End If

        Dim sRet : sRet = "<" & PoTagInfo.Item("element")
        Dim vAttr, vArr

        '�����iattribute�j�̕ҏW
        If Not IsEmpty(PoTagInfo.Item("attribute")) Then
        'attribute����łȂ��ꍇ
            Set vAttr = new_Arr()
            Set vArr = PoTagInfo.Item("attribute").slice(0,vbNullString)
            Do While vArr.length>0
                vAttr.push func_CmHtmlEditAttribute(vArr.shift)
            Loop
            sRet = sRet & " " & vAttr.join(" ")
        End If
        
        '���e�icontent�j�̕ҏW
        If Not IsEmpty(PoTagInfo.Item("content")) Then
        'content����łȂ��ꍇ
            sRet = sRet & ">"
            Set vAttr = new_Arr()
            Set vArr = PoTagInfo.Item("content").slice(0,vbNullString)
            Do While vArr.length>0
                vAttr.push func_CmHtmlEditContent(vArr.shift)
            Loop
            sRet = sRet & vAttr.join("")
            sRet = sRet & "</" & PoTagInfo.Item("element") & ">"
        Else
        'content����̏ꍇ
            sRet = sRet & " />"
        End If

        '��������HTML��ԋp
        func_CmHtmlGenGenerate = sRet

        Set vAttr = Nothing
        Set vArr = Nothing
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : func_CmHtmlEditAttribute()
    'Overview                    : �����iattribute�j�̕ҏW����
    'Detailed Description        : �H����
    'Argument
    '     aoAttr                 : �ҏW���鑮���iattribute�j
    'Return Value
    '     �ҏW����
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/22         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function func_CmHtmlEditAttribute( _
        byRef aoAttr _
        )
        Dim sRet
        If IsEmpty(aoAttr.Item("value")) Then
            sRet = aoAttr.Item("key")
        Else
            sRet = aoAttr.Item("key") & "=" & Chr(34) & aoAttr.Item("value") & Chr(34)
        End If
        func_CmHtmlEditAttribute = sRet
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : func_CmHtmlEditContent()
    'Overview                    : ���e�icontent�j�̕ҏW����
    'Detailed Description        : �H����
    'Argument
    '     aoCont                 : �ҏW������e�icontent�j
    'Return Value
    '     �ҏW����
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/22         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function func_CmHtmlEditContent( _
        byRef aoCont _
        )
        Dim sRet
        On Error Resume Next
        sRet = aoCont.generate()
        If Err.Number<>0 Then
            sRet = aoCont
        End If
        On Error GoTo 0
        func_CmHtmlEditContent = sRet
    End Function

End Class
