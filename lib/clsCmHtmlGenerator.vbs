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

        '�����iattribute�j�̕ҏW
        If Not IsEmpty(PoTagInfo.Item("attribute")) Then
        'attribute����łȂ��ꍇ
            Dim oFunc : Set oFunc = new_Func( _
            "function(e,i,a){If IsEmpty(e.Item('value')) Then:return e.Item('key'):Else:return e.Item('key') & '=''' & e.Item('value') & '''':End If}" _
            )
            sRet = sRet & " " & PoTagInfo.Item("attribute").map(oFunc).join(" ")
        End If
        
        '���e�icontent�j�̕ҏW
        If Not IsEmpty(PoTagInfo.Item("content")) Then
        'content����łȂ��ꍇ
            sRet = sRet & ">"
            sRet = sRet & new_ArrWith(PoTagInfo.Item("content")).map(getref("func_CmHtmlEditContents")).join("")
            sRet = sRet & "</" & PoTagInfo.Item("element") & ">"
        Else
        'content����̏ꍇ
            sRet = sRet & " />"
        End If

        '��������HTML��ԋp
        func_CmHtmlGenGenerate = sRet
    End Function

'    '***************************************************************************************************
'    'Function/Sub Name           : func_CmHtmlEditAttributes()
'    'Overview                    : �����iattribute�j�̕ҏW
'    'Detailed Description        : �H����
'    'Argument
'    '     aoEle                  : �z��̗v�f
'    '     alIdx                  : �C���f�b�N�X
'    '     avArr                  : �z��
'    'Return Value
'    '     ��������HTML
'    '---------------------------------------------------------------------------------------------------
'    'Histroy
'    'Date               Name                     Reason for Changes
'    '----------         ----------------------   -------------------------------------------------------
'    '2023/10/23         Y.Fujii                  First edition
'    '***************************************************************************************************
'    Private Function func_CmHtmlEditAttributes( _
'        byRef aoEle _
'        , byVal alIdx _
'        , byRef avArr _
'        )
'        If IsEmpty(aoEle.Item("value")) Then
'            func_CmHtmlEditAttributes = aoEle.Item("key")
'        Else
'            func_CmHtmlEditAttributes = aoEle.Item("key") & "=""" & aoEle.Item("value") & """"
'        End If
'    End Function

    '***************************************************************************************************
    'Function/Sub Name           : func_CmHtmlEditContents()
    'Overview                    : ���e�icontent�j�̕ҏW
    'Detailed Description        : �H����
    'Argument
    '     aoEle                  : �z��̗v�f
    '     alIdx                  : �C���f�b�N�X
    '     avArr                  : �z��
    'Return Value
    '     ��������HTML
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/23         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function func_CmHtmlEditContents( _
        byRef aoEle _
        , byVal alIdx _
        , byRef avArr _
        )
        If func_CM_UtilIsSame(TypeName(aoEle), TypeName(Me)) Then
            func_CmHtmlEditContents = aoEle.generate()
        Else
            func_CmHtmlEditContents = aoEle
        End If
    End Function

End Class
