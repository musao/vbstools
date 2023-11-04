'***************************************************************************************************
'FILENAME                    : clsCmHtmlGenerator.vbs
'Overview                    : HTML�����N���X
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
        If new_Re("^[!-~][ -~]*$", "i").Test(asElement) Then
            PoTagInfo.Item("element") = asElement
        Else
            Err.Raise 1032, "clsCmHtmlGenerator.vbs:clsCmHtmlGenerator+element()", "�v�f�ielement�j�ɂ͔��p�ȊO�̕������w��ł��܂���B"
        End If
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
    'Overview                    : ���e��ǉ�����
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
        Dim vNewArr, vArr

        '�����iattribute�j�̕ҏW
        If Not IsEmpty(PoTagInfo.Item("attribute")) Then
        'attribute����łȂ��ꍇ
            Set vNewArr = new_Arr()
            Set vArr = PoTagInfo.Item("attribute").slice(0,vbNullString)
            Do While vArr.length>0
                vNewArr.push func_CmHtmlGenEditAttribute(vArr.shift)
            Loop
            sRet = sRet & " " & vNewArr.join(" ")
        End If
        
        '���e�icontent�j�̕ҏW
        If Not IsEmpty(PoTagInfo.Item("content")) Then
        'content����łȂ��ꍇ
            sRet = sRet & ">"
            Set vNewArr = new_Arr()
            Set vArr = PoTagInfo.Item("content").slice(0,vbNullString)
            Do While vArr.length>0
                vNewArr.push func_CmHtmlGenEditContent(vArr.shift)
            Loop
            sRet = sRet & vNewArr.join("")
            sRet = sRet & "</" & PoTagInfo.Item("element") & ">"
        Else
        'content����̏ꍇ
            sRet = sRet & " />"
        End If

        '��������HTML��ԋp
        func_CmHtmlGenGenerate = sRet

        Set vNewArr = Nothing
        Set vArr = Nothing
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : func_CmHtmlGenEditAttribute()
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
    Private Function func_CmHtmlGenEditAttribute( _
        byRef aoAttr _
        )
        Dim sRet
        If IsEmpty(aoAttr.Item("value")) Then
            sRet = aoAttr.Item("key")
        Else
            sRet = aoAttr.Item("key") & "=" & Chr(34) & aoAttr.Item("value") & Chr(34)
        End If
        func_CmHtmlGenEditAttribute = sRet
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : func_CmHtmlGenEditContent()
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
    Private Function func_CmHtmlGenEditContent( _
        byRef aoCont _
        )
        Dim sRet
        On Error Resume Next
        sRet = aoCont.generate()
        If Err.Number<>0 Then
            sRet = func_CmHtmlGenHtmlEncoding(aoCont)
        End If
        On Error GoTo 0
        func_CmHtmlGenEditContent = sRet
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : func_CmHtmlGenHtmlEncoding()
    'Overview                    : HTML�G���R�[�h����
    'Detailed Description        : �H����
    'Argument
    '     asTarget               : HTML�G���R�[�h���镶����
    'Return Value
    '     �G���R�[�h����������
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/11/04         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function func_CmHtmlGenHtmlEncoding( _
        byRef asTarget _
        )
        Dim vSettings : vSettings = Array( _
            Array("&", "&amp;") _
            , Array("'", "&#39;") _
            , Array("""", "&quot;") _
            , Array("<", "&lt;") _
            , Array(">", "&gt;") _
            )
        Dim sTarget : sTarget = asTarget
        Dim i
        For Each i In vSettings
            sTarget = Replace(sTarget, i(0), i(1))
        Next
        func_CmHtmlGenHtmlEncoding = sTarget
    End Function

End Class
