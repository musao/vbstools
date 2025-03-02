'***************************************************************************************************
'FILENAME                    : HtmlGenerator.vbs
'Overview                    : HTML�����N���X
'Detailed Description        : �H����
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/10/22         Y.Fujii                  First edition
'***************************************************************************************************
Class HtmlGenerator
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
        Set PoTagInfo = new_DicOf(Array("element", Empty, "attribute", Empty, "content", Empty))
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
        attribute = PoTagInfo.Item("attribute")
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
        content = PoTagInfo.Item("content")
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
'        If new_Re("^[!-~][ -~]*$", "i").Test(asElement) Then
'            PoTagInfo.Item("element") = asElement
'        Else
'            Err.Raise 1032, "HtmlGenerator.vbs:HtmlGenerator+element()", "�v�f�ielement�j�ɂ͔��p�ȊO�̕������w��ł��܂���B"
'        End If
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
    '     ���g�̃C���X�^���X
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/23         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function addContent( _
        byRef avCont _
        )
        Dim vArr : cf_bind vArr, PoTagInfo.Item("content")
        cf_push vArr, avCont
        cf_bindAt PoTagInfo, "content", vArr

        Set addContent = Me
    End Function
        
    '***************************************************************************************************
    'Function/Sub Name           : addAttribute()
    'Overview                    : ������ǉ�����
    'Detailed Description        : �H����
    'Argument
    '     asKey                  : �ǉ����鑮���̃L�[
    '     asValue                : �ǉ����鑮���̒l
    'Return Value
    '     ���g�̃C���X�^���X
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/23         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function addAttribute( _
        byVal asKey _
        , byVal asValue _
        )
        Dim oNewAttr : Set oNewAttr = new_DicOf(Array("key", asKey, "value", asValue))
        Dim vArr : cf_bind vArr, PoTagInfo.Item("attribute")
        cf_push vArr, oNewAttr
        cf_bindAt PoTagInfo, "attribute", vArr

        Set addAttribute = Me
        Set oNewAttr = Nothing
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : generate()
    'Overview                    : HTML�𐶐�����
    'Detailed Description        : this_generate()�ɈϏ�����
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
    'Overview                    : HTML�𐶐�����
    'Detailed Description        : �H����
    'Argument
    '     asSource               : �\�[�X
    'Return Value
    '     ��������HTML
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/22         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_generate( _
        byVal asSource _
        )
        ast_argNotEmpty PoTagInfo.Item("element"), asSource, "HTML tags without elements cannot be generated."
'        If IsEmpty(PoTagInfo.Item("element")) Then
'            Err.Raise 17, "HtmlGenerator.vbs:HtmlGenerator-this_generate()", "�v�f���Ȃ�HTML�^�O�͐����ł��܂���B"
'            Exit Function
'        End If

        '�J�n�^�O�̕ҏW
        Dim sStt : sStt =  "<" & PoTagInfo.Item("element")
        Dim vArr, vEle
        '�����iattribute�j�̕ҏW
        If Not IsEmpty(PoTagInfo.Item("attribute")) Then
        'attribute����łȂ��ꍇ
            For Each vEle In PoTagInfo.Item("attribute")
                cf_push vArr, this_editAttribute(vEle)
            Next
            sStt = sStt & " " & Join(vArr, " ")
        End If
        If Not IsEmpty(PoTagInfo.Item("content")) Then
        'content����łȂ��ꍇ
            sStt = sStt & ">"
        Else
        'content����̏ꍇ
            sStt = sStt & " />"
        End If
        
        '���e�icontent�j�̕ҏW
        Dim sCont : sCont = ""
        If Not IsEmpty(PoTagInfo.Item("content")) Then
        'content����łȂ��ꍇ
            vArr = Array()
            For Each vEle In PoTagInfo.Item("content")
                cf_push vArr, this_editContent(vEle)
            Next
            sCont = new_Re("^([^\n])", "igm").Replace(Join(vArr, vbNewLine),"  $1")
        End If

        '�I���^�O�̕ҏW
        Dim sEnd : sEnd =  ""
        If Not IsEmpty(PoTagInfo.Item("content")) Then
        'content����łȂ��ꍇ
            sEnd =  "</" & PoTagInfo.Item("element") & ">"
        End If

        '��������HTML��ԋp
        sRet = sStt
        If Not IsEmpty(PoTagInfo.Item("content")) Then sRet = sRet & vbNewLine & sCont & vbNewLine & sEnd
        this_generate = sRet

    End Function

    '***************************************************************************************************
    'Function/Sub Name           : this_editAttribute()
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
    Private Function this_editAttribute( _
        byRef aoAttr _
        )
        Dim sRet
        If IsEmpty(aoAttr.Item("value")) Then
            sRet = aoAttr.Item("key")
        Else
            sRet = aoAttr.Item("key") & "=" & Chr(34) & aoAttr.Item("value") & Chr(34)
        End If
        this_editAttribute = sRet
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : this_editContent()
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
    Private Function this_editContent( _
        byRef aoCont _
        )
        Dim sRet
        On Error Resume Next
        sRet = aoCont.generate()
        If Err.Number<>0 Then
            sRet = this_htmlEntityReference(aoCont)
        End If
        On Error GoTo 0
        this_editContent = sRet
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : this_htmlEntityReference()
    'Overview                    : HTML�̓��ꕶ�������̎Q�Ɓientity reference�j��������
    'Detailed Description        : HTML�Ƃ��ē���ȈӖ����������i���ꕶ���܂��̓��^�����j���Ӗ��������Ȃ�
    '                              �ʂ̕�����ɒu������
    'Argument
    '     asTarget               : ���̎Q�Ə������镶����
    'Return Value
    '     ���̎Q�Ə�������������
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/11/04         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_htmlEntityReference( _
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
        this_htmlEntityReference = sTarget
    End Function

End Class
