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
        Set PoTagInfo = new_DicWith("element", Empty, "attribute", Empty, "content", Empty)
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
    'Overview                    : ���̓��t�������擾����
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
        If IsEmpty(PoTagInfo.Item("element")) Then Exit Function

        Dim sRet : sRet = "<" & PoTagInfo.Item("element")
        If Not IsEmpty(PoTagInfo.Item("attribute")) Then
        'attribute����łȂ��ꍇ
            sRet = sRet & " " & newArrWith(PoTagInfo.Item("attribute")).map(new_Func("(e,i,a)=>e.Item(""key"")&""=""&e.Item(""""value"""")")).join(" ")
        End If
        
        If Not IsEmpty(PoTagInfo.Item("content")) Then
        'content����łȂ��ꍇ
            sRet = sRet & ">"
            Dim oEle, sCont
            sCont = ""
            For Each oEle In PoTagInfo.Item("content")
            'content���Ƃɏ�������
                If func_CM_UtilIsSame(TypeName(oEle), TypeName(Me)) Then
                    sCont = sCont & oEle.generate()
                Else
                    sCont = sCont & oEle
                End If
            Next
            sRet = sRet & sCont & "</" & PoTagInfo.Item("element") & ">"
        Else
        'content����̏ꍇ
            sRet = sRet & " />"
        End If

        '��������HTML��ԋp
        func_CmHtmlGenGenerate = sRet
    End Function
End Class
