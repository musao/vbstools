'***************************************************************************************************
'FILENAME                    : clsCmArray.vbs
'Overview                    : �z��N���X
'Detailed Description        : javacsript��Array�I�u�W�F�N�g�����A�v���~�e�B�u�̔z��ł͂Ȃ�
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/08         Y.Fujii                  First edition
'***************************************************************************************************

'***************************************************************************************************
'Function/Sub Name           : new_clsCmArray()
'Overview                    : �C���X�^���X�����֐�
'Detailed Description        : �����������N���X�̃C���X�^���X��Ԃ�
'Argument
'     �Ȃ�
'Return Value
'     ���N���X�̃C���X�^���X
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/08         Y.Fujii                  First edition
'***************************************************************************************************
Private Function new_clsCmArray( _
    )
    Set new_clsCmArray = (New clsCmArray)
End Function

'***************************************************************************************************
'Function/Sub Name           : new_ArraySetData()
'Overview                    : �C���X�^���X�����֐�
'Detailed Description        : �����Ŏw�肵���v�f���܂񂾓��N���X�̃C���X�^���X��Ԃ�
'Argument
'     aoElements             : �z��ɒǉ�����v�f�i�z��j
'Return Value
'     ���N���X�̃C���X�^���X
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/08         Y.Fujii                  First edition
'***************************************************************************************************
Private Function new_ArraySetData( _
    byRef aoElements _
    )
    Dim oArray : Set oArray = new_clsCmArray()
    oArray.PushMulti aoElements
    Set new_ArraySetData = oArray
    Set oArray = Nothing
End Function

'***************************************************************************************************
'Function/Sub Name           : new_ArraySplit()
'Overview                    : �C���X�^���X�����֐�
'Detailed Description        : vbscript��Split�֐��Ɠ����̋@�\�A���N���X�̃C���X�^���X��Ԃ�
'Argument
'     asTarget               : ����������Ƌ�؂蕶�����܂ޕ�����\��
'     asDelimiter            : ��؂蕶��
'     alCompare              : ��r���@
'                                0(vbBinaryCompare):�o�C�i����r�����s���܂�
'                                1(vbTextCompare ):�e�L�X�g��r�����s���܂�
'Return Value
'     ���N���X�̃C���X�^���X
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/08         Y.Fujii                  First edition
'***************************************************************************************************
Private Function new_ArraySplit( _
    byVal asTarget _
    , byVal asDelimiter _
    , byVal alCompare _
    )
    Set new_ArraySplit = new_ArraySetData(Split(asTarget, asDelimiter, -1, alCompare))
End Function

Class clsCmArray
    '�N���X���ϐ��A�萔
    Private PoArray
    
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
    '2023/09/08         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub Class_Initialize()
        Set PoArray = new_Dictionary()
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
    '2023/09/08         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub Class_Terminate()
        Set PoArray = Nothing
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get Item()
    'Overview                    : �z��̎w�肵���C���f�b�N�X�̗v�f��Ԃ�
    'Detailed Description        : func_CmArrayItem()�ɈϏ�����
    'Argument
    '     aIndex                 : �C���f�b�N�X
    'Return Value
    '     �w�肵���C���f�b�N�X�̗v�f
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/08         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Default Property Get Item( _
        byVal aIndex _
        )
        Call sub_CM_Bind(Item, func_CmArrayItem(aIndex))
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Set Item()
    'Overview                    : �z��̎w�肵���C���f�b�N�X�ɗv�f��ݒ肷��
    'Detailed Description        : �H����
    'Argument
    '     aIndex                 : �C���f�b�N�X
    '     aoElement              : �ݒ肷��v�f
    'Return Value
    '     �Ȃ�
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/08         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Set Item( _
        byVal aIndex _
        , byRef aoElement _
        )
        Call sub_CM_BindAt(PoArray, aIndex, aoElement)
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Let Item()
    'Overview                    : �z��̎w�肵���C���f�b�N�X�ɗv�f��ݒ肷��
    'Detailed Description        : �H����
    'Argument
    '     aIndex                 : �C���f�b�N�X
    '     aoElement              : �ݒ肷��v�f
    'Return Value
    '     �Ȃ�
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/08         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Let Item( _
        byVal aIndex _
        , byRef aoElement _
        )
        Call sub_CM_BindAt(PoArray, aIndex, aoElement)
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get Items()
    'Overview                    : �z���Ԃ�
    'Detailed Description        : func_CmArrayConvArray()�ɈϏ�����
    'Argument
    '     �Ȃ�
    'Return Value
    '     �z��
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/08         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get Items( _
        )
        Items = func_CmArrayConvArray()
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get Length()
    'Overview                    : �z����̗v�f����Ԃ�
    'Detailed Description        : �H����
    'Argument
    '     �Ȃ�
    'Return Value
    '     �z����̗v�f��
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/08         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get Length()
        Length = PoArray.Count
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Push()
    'Overview                    : �z��̖����ɗv�f��1�ǉ�����
    'Detailed Description        : func_CmArrayPushMulti()�ɈϏ�����
    'Argument
    '     aoElement              : �z��̖����ɒǉ�����v�f
    'Return Value
    '     �z����̗v�f��
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/08         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function Push( _
        byRef aoElement _
        )
        Push = func_CmArrayPushMulti(Array(aoElement))
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : PushMulti()
    'Overview                    : �z��̖����ɗv�f��1�ǉ�����
    'Detailed Description        : func_CmArrayPushMulti()�ɈϏ�����
    'Argument
    '     aoElements             : �z��̖����ɒǉ�����v�f�i�z��j
    'Return Value
    '     �z����̗v�f��
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/08         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function PushMulti( _
        byRef aoElements _
        )
        PushMulti = func_CmArrayPushMulti(aoElements)
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : Unshift()
    'Overview                    : �z��̐擪�ɗv�f��1�ǉ�����
    'Detailed Description        : func_CmArrayUnshiftMulti()�ɈϏ�����
    'Argument
    '     aoElement              : �z��̐擪�ɒǉ�����v�f
    'Return Value
    '     �z����̗v�f��
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/08         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function Unshift( _
        byRef aoElement _
        )
        Unshift = func_CmArrayUnshiftMulti(Array(aoElement))
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : UnshiftMulti()
    'Overview                    : �z��̐擪�ɗv�f��1�ǉ�����
    'Detailed Description        : func_CmArrayUnshiftMulti()�ɈϏ�����
    'Argument
    '     aoElements             : �z��̐擪�ɒǉ�����v�f�i�z��j
    'Return Value
    '     �z����̗v�f��
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/08         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function UnshiftMulti( _
        byRef aoElements _
        )
        UnshiftMulti = func_CmArrayUnshiftMulti(aoElements)
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : Pop()
    'Overview                    : �z�񂩂疖���̗v�f����菜��
    'Detailed Description        : func_CmArrayPop()�ɈϏ�����
    'Argument
    '     �Ȃ�
    'Return Value
    '     �z�񂩂��菜�����v�f
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/08         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function Pop( _
        )
        Call sub_CM_Bind(Pop, func_CmArrayPop())
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : Shift()
    'Overview                    : �z�񂩂�擪�̗v�f����菜��
    'Detailed Description        : func_CmArrayShift()�ɈϏ�����
    'Argument
    '     �Ȃ�
    'Return Value
    '     �z�񂩂��菜�����v�f
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/08         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function Shift( _
        )
        Call sub_CM_Bind(Shift, func_CmArrayShift())
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : Filter()
    'Overview                    : �����̊֐��Œ��o�����v�f�����̔z����쐬
    'Detailed Description        : func_CmArrayFilter()�ɈϏ�����
    'Argument
    '     aoFunc                 : ���o����֐�
    'Return Value
    '     ���N���X�̃C���X�^���X
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/08         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function Filter( _
        byRef aoFunc _
        )
        Call sub_CM_Bind(Filter, func_CmArrayFilter(aoFunc))
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : FilterVbs()
    'Overview                    : �����Ŏw�肵�������ɍ��v����v�f�����̔z����쐬����
    'Detailed Description        : vbscript��Filter�֐��Ɠ����̋@�\
    'Argument
    '     asTarget               : �������镶����
    '     aobInclude             : �������镶����������ΏۂƂ��邩�ۂ��̋敪
    '                                True :�������镶����������ΏۂƂ���
    '                                False:�������镶����ȊO�������ΏۂƂ���
    '     alCompare              : ��r���@
    '                                0(vbBinaryCompare):�o�C�i����r�����s���܂�
    '                                1(vbTextCompare ):�e�L�X�g��r�����s���܂�
    'Return Value
    '     ���N���X�̃C���X�^���X
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/08         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function FilterVbs( _
        byVal asTarget _
        , byVal aobInclude _
        , byVal alCompare _
        )
        Call sub_CM_Bind(FilterVbs, new_ArraySetData(Filter(func_CmArrayConvArray(), asTarget, aobInclude, alCompare)))
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : JoinVbs()
    'Overview                    : �z��̊e�v�f��A��������������쐬����
    'Detailed Description        : vbscript��Join�֐��Ɠ����̋@�\
    'Argument
    '     asDelimiter            : ��؂蕶��
    'Return Value
    '     �z��̊e�v�f��A������������
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/08         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function JoinVbs( _
        byVal asDelimiter _
        )
        JoinVbs = Join(func_CmArrayConvArray(), asDelimiter)
    End Function
    
    
    
    
    
    '***************************************************************************************************
    'Function/Sub Name           : func_CmArrayItem()
    'Overview                    : �z��̎w�肵���C���f�b�N�X�̗v�f��Ԃ�
    'Detailed Description        : �H����
    'Argument
    '     aIndex                 : �C���f�b�N�X
    'Return Value
    '     �w�肵���C���f�b�N�X�̗v�f
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/08         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function func_CmArrayItem( _
        ByVal aIndex _
        )
        Dim oElement : Set oElement = Nothing
        If PoArray.Count>0 Then
            Call sub_CM_Bind(oElement, PoArray.Item(aIndex))
        End If
        Call sub_CM_Bind(func_CmArrayItem, oElement)
        Set oElement = Nothing
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : func_CmArrayPushMulti()
    'Overview                    : �z��̖����ɗv�f�𕡐��ǉ�����
    'Detailed Description        : �H����
    'Argument
    '     aoElements             : �z��̖����ɒǉ�����v�f�i�z��j
    'Return Value
    '     �z����̗v�f��
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/08         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function func_CmArrayPushMulti( _
        byRef aoElements _
        )
        If IsArray(aoElements) Then
            Dim oItem
            For Each oItem In aoElements
                Call sub_CM_BindAt(PoArray, PoArray.Count, oItem)
            Next
        End If
        func_CmArrayPushMulti = PoArray.Count
        Set oItem = Nothing
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : func_CmArrayUnshiftMulti()
    'Overview                    : �z��̐擪�ɗv�f�𕡐��ǉ�����
    'Detailed Description        : �H����
    'Argument
    '     aoElements             : �z��̐擪�ɒǉ�����v�f�i�z��j
    'Return Value
    '     �z����̗v�f��
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/08         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function func_CmArrayUnshiftMulti( _
        byRef aoElements _
        )
        Dim oArray, oItem
        
        If IsArray(aoElements) Then
            '�����̗v�f��擪�ɒǉ�
            Set oArray = new_Dictionary()
            For Each oItem In aoElements
                Call sub_CM_BindAt(oArray, oArray.Count, oItem)
            Next
        End If
        
        '�����č�����v�f��ǉ�
        For Each oItem In PoArray.Items()
            Call sub_CM_BindAt(oArray, oArray.Count, oItem)
        Next
        
        '�쐬�����z��i�f�B�N�V���i���j��u����
        Set PoArray = oArray
        func_CmArrayUnshiftMulti = PoArray.Count
        
        Set oItem = Nothing
        Set oArray = Nothing
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : func_CmArrayPop()
    'Overview                    : �z�񂩂疖���̗v�f����菜��
    'Detailed Description        : �H����
    'Argument
    '     �Ȃ�
    'Return Value
    '     �z�񂩂��菜�����v�f
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/08         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function func_CmArrayPop( _
        )
        Dim oElement, lCount
        Set oElement = Nothing
        lCount = PoArray.Count
        If lCount>0 Then
            Call sub_CM_Bind(oElement, PoArray.Item(lCount-1))
            PoArray.Remove lCount-1
        End If
        Call sub_CM_Bind(func_CmArrayPop, oElement)
        Set oElement = Nothing
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : func_CmArrayShift()
    'Overview                    : �z�񂩂�擪�̗v�f����菜��
    'Detailed Description        : �H����
    'Argument
    '     �Ȃ�
    'Return Value
    '     �z�񂩂��菜�����v�f
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/08         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function func_CmArrayShift( _
        )
        Dim oElement, oArray, oItem, boFlg
        Set oElement = Nothing
        Set oArray = new_Dictionary()
        
        '�擪�̗v�f���������z����č쐬
        If PoArray.Count>0 Then
            boFlg = False
            For Each oItem In PoArray.Items()
                If boFlg Then
                    Call sub_CM_BindAt(oArray, oArray.Count, oItem)
                Else
                    Call sub_CM_Bind(oElement, PoArray.Item(0))
                    boFlg = True
                End If
            Next
        End If
        
        '�쐬�����z��i�f�B�N�V���i���j��u����
        Set PoArray = oArray
        Call sub_CM_Bind(func_CmArrayShift, oElement)
        
        Set oElement = Nothing
        Set oItem = Nothing
        Set oArray = Nothing
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : func_CmArrayFilter()
    'Overview                    : �����̊֐��Œ��o�����v�f�����̔z����쐬
    'Detailed Description        : �H����
    'Argument
    '     aoFunc                 : ���o����֐�
    'Return Value
    '     ���N���X�̃C���X�^���X
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/08         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function func_CmArrayFilter( _
        byRef aoFunc _
        )
        Dim oItem, oArray
        
        '�����̊֐��Œ��o�����v�f�������o
        If PoArray.Count>0 Then
            For Each oItem In PoArray.Items()
                If aoFunc(oItem) Then
                    Call sub_CM_Push(oArray, oItem)
                End If
            Next
        End If
        
        '�쐬�����z��i�f�B�N�V���i���j�œ��N���X�̃C���X�^���X�𐶐����ĕԋp
        Call sub_CM_Bind(func_CmArrayFilter, new_ArraySetData(oArray))
        
        Set oItem = Nothing
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : func_CmArrayConvArray()
    'Overview                    : �����ŕێ�����z��i�f�B�N�V���i���j���v���~�e�B�u�̔z��ɕϊ�����
    'Detailed Description        : �H����
    'Argument
    '     �Ȃ�
    'Return Value
    '     �z��
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/08         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function func_CmArrayConvArray( _
        )
        Dim oArray
        
        If PoArray.Count>0 Then
            Dim oItem
            For Each oItem In PoArray.Items()
                Call sub_CM_Push(oArray, oItem)
            Next
        End If
        func_CmArrayConvArray = oArray
        
        Set oArray = Nothing
    End Function
    
End Class
