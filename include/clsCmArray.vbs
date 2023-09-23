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
'     aoEles                 : �z��ɒǉ�����v�f�i�z��j
'Return Value
'     ���N���X�̃C���X�^���X
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/08         Y.Fujii                  First edition
'***************************************************************************************************
Private Function new_ArraySetData( _
    byRef aoEles     _
    )
    Dim oArray : Set oArray = new_clsCmArray()
    oArray.PushMulti aoEles    
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
'                                0(vbBinaryCompare):�o�C�i����r
'                                1(vbTextCompare):�e�L�X�g��r
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
    '     alIndex                : �C���f�b�N�X
    'Return Value
    '     �w�肵���C���f�b�N�X�̗v�f
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/08         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Default Property Get Item( _
        byVal alIndex _
        )
        Call sub_CM_Bind(Item, func_CmArrayItem(alIndex))
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Set Item()
    'Overview                    : �z��̎w�肵���C���f�b�N�X�ɗv�f��ݒ肷��
    'Detailed Description        : �H����
    'Argument
    '     alIndex                : �C���f�b�N�X
    '     aoEle                  : �ݒ肷��v�f
    'Return Value
    '     �Ȃ�
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/08         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Set Item( _
        byVal alIndex _
        , byRef aoEle _
        )
        Call sub_CM_BindAt(PoArray, alIndex, aoEle)
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Let Item()
    'Overview                    : �z��̎w�肵���C���f�b�N�X�ɗv�f��ݒ肷��
    'Detailed Description        : �H����
    'Argument
    '     alIndex                : �C���f�b�N�X
    '     aoEle                  : �ݒ肷��v�f
    'Return Value
    '     �Ȃ�
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/08         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Let Item( _
        byVal alIndex _
        , byRef aoEle _
        )
        Call sub_CM_BindAt(PoArray, alIndex, aoEle)
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
        Items = func_CmArrayConvArray(True)
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
    'Function/Sub Name           : Concat()
    'Overview                    : �����Ŏw�肵���v�f��A�������z���Ԃ�
    'Detailed Description        : ���g�̃C���X�^���X�͕ύX���Ȃ�
    'Argument
    '     aoEles                 : �z��ɒǉ�����v�f�i�z��j
    'Return Value
    '     ���N���X�̕ʃC���X�^���X
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/14         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function Concat( _
        byRef aoEles     _
        )
        Dim oArray : Set oArray = new_clsCmArray()
        oArray.PushMulti func_CmArrayConvArray(True)
        oArray.PushMulti aoEles    
        Set Concat = oArray
        
        Set oArray = Nothing
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : Every()
    'Overview                    : �z��̑S�Ă̗v�f�������̊֐��̔���𖞂������m�F����
    'Detailed Description        : func_CmArrayEvery()�ɈϏ�����
    'Argument
    '     aoFunc                 : ���肷��֐�
    'Return Value
    '     ���� True:������ / False:�������Ȃ�
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/14         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function Every( _
        byRef aoFunc _
        )
        Every = func_CmArrayEveryOrSome(aoFunc, True)
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : Filter()
    'Overview                    : �����̊֐��Œ��o�����v�f�����̔z����쐬
    'Detailed Description        : func_CmArrayFilter()�ɈϏ�����
    'Argument
    '     aoFunc                 : ���o����֐�
    'Return Value
    '     ���N���X�̕ʃC���X�^���X
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/08         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function Filter( _
        byRef aoFunc _
        )
        Set Filter = func_CmArrayFilter(aoFunc)
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
    '                                0(vbBinaryCompare):�o�C�i����r
    '                                1(vbTextCompare):�e�L�X�g��r
    'Return Value
    '     ���N���X�̕ʃC���X�^���X
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
        Set FilterVbs = new_ArraySetData( Filter(func_CmArrayConvArray(True), asTarget, aobInclude, alCompare) )
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : Find()
    'Overview                    : �����̊֐��Œ��o�����ŏ��̗v�f��Ԃ�
    'Detailed Description        : func_CmArrayFind()�ɈϏ�����
    'Argument
    '     aoFunc                 : ���o����֐�
    'Return Value
    '     �z�񂩂璊�o�����v�f
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/13         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function Find( _
        byRef aoFunc _
        )
        Call sub_CM_Bind(Find, func_CmArrayFind(aoFunc))
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : ForEach()
    'Overview                    : �z��̑S�Ă̗v�f�ɂ��Ĉ����̊֐��̏������s��
    'Detailed Description        : func_CmArrayForEach()�ɈϏ�����
    'Argument
    '     aoFunc                 : �֐�
    'Return Value
    '     �Ȃ�
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/14         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Sub ForEach( _
        byRef aoFunc _
        )
        Call func_CmArrayForEach(aoFunc)
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : IndexOf()
    'Overview                    : �����ɍ��v����v�f�𐳏��ɒT���ŏ��Ɍ��������C���f�b�N�X�ԍ���Ԃ�
    'Detailed Description        : func_CmArrayIndexOf()�ɈϏ�����
    'Argument
    '     avTarget               : ��v���m�F������e
    'Return Value
    '     �����ɍ��v����v�f�̃C���f�b�N�X�ԍ�
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/14         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function IndexOf( _
        byRef avTarget _
        )
        IndexOf = func_CmArrayIndexOf(avTarget, vbNullString, vbBinaryCompare, True)
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
        JoinVbs = Join(func_CmArrayConvArray(True), asDelimiter)
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : LastIndexOf()
    'Overview                    : �����ɍ��v����v�f���t���ɒT���ŏ��Ɍ��������C���f�b�N�X�ԍ���Ԃ�
    'Detailed Description        : func_CmArrayIndexOf()�ɈϏ�����
    'Argument
    '     avTarget               : ��v���m�F������e
    'Return Value
    '     �����ɍ��v����v�f�̃C���f�b�N�X�ԍ�
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/15         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function LastIndexOf( _
        byRef avTarget _
        )
        LastIndexOf = func_CmArrayIndexOf(avTarget, vbNullString, vbBinaryCompare, False)
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : Map()
    'Overview                    : �z�񂩂�����̊֐��ŐV���Ȕz��𐶐�����
    'Detailed Description        : func_CmArrayMap()�ɈϏ�����
    'Argument
    '     aoFunc                 : �֐�
    'Return Value
    '     ���N���X�̕ʃC���X�^���X
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/16         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function Map( _
        byRef aoFunc _
        )
        Call sub_CM_Bind(Map, func_CmArrayMap(aoFunc))
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
    'Function/Sub Name           : Push()
    'Overview                    : �z��̖����ɗv�f��1�ǉ�����
    'Detailed Description        : func_CmArrayPushMulti()�ɈϏ�����
    'Argument
    '     aoEle                  : �z��̖����ɒǉ�����v�f
    'Return Value
    '     �z����̗v�f��
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/08         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function Push( _
        byRef aoEle _
        )
        Push = func_CmArrayPushMulti(Array(aoEle))
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : PushMulti()
    'Overview                    : �z��̖����ɗv�f��1�ǉ�����
    'Detailed Description        : func_CmArrayPushMulti()�ɈϏ�����
    'Argument
    '     aoEles                 : �z��̖����ɒǉ�����v�f�i�z��j
    'Return Value
    '     �z����̗v�f��
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/08         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function PushMulti( _
        byRef aoEles     _
        )
        PushMulti = func_CmArrayPushMulti(aoEles    )
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : Reduce()
    'Overview                    : �z��̂��ꂼ��̗v�f�ɑ΂��Đ����Ɉ����̊֐��ŎZ�o�������ʂ�Ԃ�
    'Detailed Description        : func_CmArrayReduce()�ɈϏ�����
    'Argument
    '     aoFunc                 : �֐�
    'Return Value
    '     �����̊֐��ŎZ�o��������
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/16         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function Reduce( _
        byRef aoFunc _
        )
        Call sub_CM_Bind(Reduce, func_CmArrayReduce(aoFunc, True))
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : ReduceRight()
    'Overview                    : �z��̂��ꂼ��̗v�f�ɑ΂��ċt���Ɉ����̊֐��ŎZ�o�������ʂ�Ԃ�
    'Detailed Description        : func_CmArrayReduce()�ɈϏ�����
    'Argument
    '     aoFunc                 : �֐�
    'Return Value
    '     �����̊֐��ŎZ�o��������
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/16         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function ReduceRight( _
        byRef aoFunc _
        )
        Call sub_CM_Bind(ReduceRight, func_CmArrayReduce(aoFunc, False))
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : Reverse()
    'Overview                    : �z��̗v�f���t���ɕ��ׂ�
    'Detailed Description        : func_CmArrayReverse()�ɈϏ�����
    'Argument
    '     �Ȃ�
    'Return Value
    '     �Ȃ�
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/14         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Sub Reverse( _
        )
        Call func_CmArrayReverse()
    End Sub
    
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
    'Function/Sub Name           : Slice()
    'Overview                    : �z��̈ꕔ��؂�o�����z��𐶐�����
    'Detailed Description        : func_CmArraySlice()�ɈϏ�����
    'Argument
    '     alStart                : �J�n�ʒu�̃C���f�b�N�X�ԍ��A���l�͍Ō�̗v�f�̂���̈ʒu������
    '                              �Ⴆ��-1�͍Ō�A-2�͍Ōォ��2�ڂ̗v�f�������B
    '     alEnd                  : �I���ʒu�̃C���f�b�N�X�ԍ��A���l��alStart�Ɠ���
    '                              �؂�o�����z��ɏI���ʒu�̗v�f�͊܂܂Ȃ�
    '                              vbNullString���w�肵���ꍇ�͐؂�o�����z��ɍŌ�̗v�f���܂߂�
    'Return Value
    '     ���N���X�̕ʃC���X�^���X
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/16         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function Slice( _
        byVal alStart _
        , byVal alEnd _
        )
        Set Slice = func_CmArraySlice(alStart, alEnd)
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : Some()
    'Overview                    : �z��̂����ꂩ��̗v�f�������̊֐��̔���𖞂������m�F����
    'Detailed Description        : func_CmArrayEvery()�ɈϏ�����
    'Argument
    '     aoFunc                 : ���肷��֐�
    'Return Value
    '     ���� True:������ / False:�������Ȃ�
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/16         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function Some( _
        byRef aoFunc _
        )
        Some = func_CmArrayEveryOrSome(aoFunc, False)
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : Unshift()
    'Overview                    : �z��̐擪�ɗv�f��1�ǉ�����
    'Detailed Description        : func_CmArrayUnshiftMulti()�ɈϏ�����
    'Argument
    '     aoEle                  : �z��̐擪�ɒǉ�����v�f
    'Return Value
    '     �z����̗v�f��
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/08         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function Unshift( _
        byRef aoEle _
        )
        Unshift = func_CmArrayUnshiftMulti(Array(aoElement))
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : UnshiftMulti()
    'Overview                    : �z��̐擪�ɗv�f��1�ǉ�����
    'Detailed Description        : func_CmArrayUnshiftMulti()�ɈϏ�����
    'Argument
    '     aoEles                 : �z��̐擪�ɒǉ�����v�f�i�z��j
    'Return Value
    '     �z����̗v�f��
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/08         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function UnshiftMulti( _
        byRef aoEles     _
        )
        UnshiftMulti = func_CmArrayUnshiftMulti(aoEles    )
    End Function
    
    
    
    
    
    '***************************************************************************************************
    'Function/Sub Name           : func_CmArrayItem()
    'Overview                    : �z��̎w�肵���C���f�b�N�X�̗v�f��Ԃ�
    'Detailed Description        : �H����
    'Argument
    '     alIndex                : �C���f�b�N�X
    'Return Value
    '     �w�肵���C���f�b�N�X�̗v�f
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/08         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function func_CmArrayItem( _
        ByVal alIndex _
        )
        Dim oEle : Set oEle = Nothing
        If PoArray.Count>0 Then
            Call sub_CM_Bind(oEle, PoArray.Item(alIndex))
        End If
        Call sub_CM_Bind(func_CmArrayItem, oEle)
        Set oEle = Nothing
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : func_CmArrayEveryOrSome()
    'Overview                    : �z��̗v�f�������̊֐��̔���𖞂������m�F����
    'Detailed Description        : �����̊֐��̈����͈ȉ��̂Ƃ���
    '                                element :�z��̗v�f
    '                                index   :�C���f�b�N�X
    '                                array   :�z��
    'Argument
    '     aoFunc                 : ���肷��֐�
    '     aboFlg                 : ������@
    '                                True  :�z��̑S�Ă̗v�f�������̊֐��̔���𖞂���
    '                                False :�z��̂����ꂩ��̗v�f�������̊֐��̔���𖞂���
    'Return Value
    '     ���� True:������ / False:�������Ȃ�
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/16         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function func_CmArrayEveryOrSome( _
        byRef aoFunc _
        , byRef aboFlg _
        )
        Dim oItem, lIndex, oArray, boRet
        boRet = aboFlg
        
        '�����̊֐��Ŕ��肷��
        If PoArray.Count>0 Then
            oArray = func_CmArrayConvArray(True)
            For lIndex=0 To PoArray.Count-1
                Call sub_CM_Bind(oItem, oArray(lIndex))
                If Not aoFunc(oItem, lIndex, oArray) = aboFlg Then
                    boRet = Not aboFlg
                    Exit For
                End If
            Next
        End If
        
        '���茋�ʂ�ԋp
        func_CmArrayEveryOrSome = boRet
        
        Set oItem = Nothing
        Set oArray = Nothing
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : func_CmArrayFilter()
    'Overview                    : �����̊֐��Œ��o�����v�f�����̔z����쐬
    'Detailed Description        : ���o�ł��Ȃ��ꍇ�͗v�f���Ȃ����N���X�̃C���X�^���X��Ԃ�
    '                              �����̊֐��̈����͈ȉ��̂Ƃ���
    '                                element :�z��̗v�f
    '                                index   :�C���f�b�N�X
    '                                array   :�z��
    'Argument
    '     aoFunc                 : ���o����֐�
    'Return Value
    '     ���N���X�̕ʃC���X�^���X
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/08         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function func_CmArrayFilter( _
        byRef aoFunc _
        )
        Dim oItem, lIndex, oArray, oRet
        
        '�����̊֐��Œ��o�����v�f�����̔z����쐬
        If PoArray.Count>0 Then
            oArray = func_CmArrayConvArray(True)
            For lIndex=0 To PoArray.Count-1
                Call sub_CM_Bind(oItem, oArray(lIndex))
                If aoFunc(oItem, lIndex, oArray) Then
                    Call sub_CM_Push(oRet, oItem)
                End If
            Next
        End If
        
        '�쐬�����z��i�f�B�N�V���i���j�œ��N���X�̃C���X�^���X�𐶐����ĕԋp
        Set func_CmArrayFilter = new_ArraySetData(oRet)
        
        Set oItem = Nothing
        Set oArray = Nothing
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : func_CmArrayFind()
    'Overview                    : �����̊֐��Œ��o�����ŏ��̗v�f��Ԃ�
    'Detailed Description        : ���o�ł��Ȃ��ꍇ��Nothing��Ԃ�
    '                              �����̊֐��̈����͈ȉ��̂Ƃ���
    '                                element :�z��̗v�f
    '                                index   :�C���f�b�N�X
    '                                array   :�z��
    'Argument
    '     aoFunc                 : ���o����֐�
    'Return Value
    '     �z�񂩂璊�o�����v�f
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/13         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function func_CmArrayFind( _
        byRef aoFunc _
        )
        Dim oItem, lIndex, oArray, oRet
        Set oRet = Nothing
        
        '�����̊֐��Œ��o�ł���ŏ��̗v�f������
        If PoArray.Count>0 Then
            oArray = func_CmArrayConvArray(True)
            For lIndex=0 To PoArray.Count-1
                Call sub_CM_Bind(oItem, oArray(lIndex))
                If aoFunc(oItem, lIndex, oArray) Then
                    Call sub_CM_Bind(oRet, oItem)
                    Exit For
                End If
            Next
        End If
        
        '�z�񂩂璊�o�����v�f��ԋp
        Call sub_CM_Bind(func_CmArrayFind, oRet)
        
        Set oItem = Nothing
        Set oArray = Nothing
        Set oRet = Nothing
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : func_CmArrayForEach()
    'Overview                    : �z��̑S�Ă̗v�f�ɂ��Ĉ����̊֐��̏������s��
    'Detailed Description        : �����̊֐��̈����͈ȉ��̂Ƃ���
    '                                element :�z��̗v�f
    '                                index   :�C���f�b�N�X
    '                                array   :�z��
    'Argument
    '     aoFunc                 : �֐�
    'Return Value
    '     �Ȃ�
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/14         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function func_CmArrayForEach( _
        byRef aoFunc _
        )
        Dim oItem, lIndex, oArray
        
        '�z��̑S�Ă̗v�f�ɂ��Ĉ����̊֐��̏������s��
        If PoArray.Count>0 Then
            oArray = func_CmArrayConvArray(True)
            For lIndex=0 To PoArray.Count-1
                Call sub_CM_Bind(oItem, oArray(lIndex))
                Call aoFunc(oItem, lIndex, oArray)
            Next
        End If
        
        Set oItem = Nothing
        Set oArray = Nothing
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : func_CmArrayIndexOf()
    'Overview                    : �����ɍ��v����v�f��T���ŏ��Ɍ��������C���f�b�N�X�ԍ���Ԃ�
    'Detailed Description        : ���v����v�f���Ȃ��ꍇ��-1��Ԃ�
    'Argument
    '     avTarget               : ��v���m�F������e
    '     alStart                : �����J�n�ʒu�̃C���f�b�N�X�ԍ�
    '                              vbNullString�̏ꍇ��aboOrder�������̏ꍇ��0�A�t���̏ꍇ�͑S�v�f��-1
    '     alCompare              : ��r���@
    '                                0(vbBinaryCompare):�o�C�i����r
    '                                1(vbTextCompare):�e�L�X�g��r
    '     aboOrder               : True�F�����i���Ԃǂ���j / False�F�t��
    'Return Value
    '     �����ɍ��v����v�f�̃C���f�b�N�X�ԍ�
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/14         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function func_CmArrayIndexOf( _
        byRef avTarget _
        , byVal alStart _
        , byVal alCompare _
        , byVal aboOrder _
        )
        func_CmArrayIndexOf = -1
        Dim oItem, lIndex, oArray, boFlg, lStart, lEnd, lStep
        
        '�z��̑S�Ă̗v�f�ɂ��Ĉ����̊֐��̏������s��
        If PoArray.Count>0 Then
            
            If alStart=vbNullString Then
                If aboOrder Then lStart=0 Else lStart=PoArray.Count-1
            Else
                lStart=alStart
            End If
            If aboOrder Then lEnd=PoArray.Count-1 Else lEnd=0
            If aboOrder Then lStep=1 Else lStep=-1
            
            boFlg = False
            oArray = func_CmArrayConvArray(True)
            For lIndex=lStart To lEnd Step lStep
                Call sub_CM_Bind(oItem, oArray(lIndex))
                
                If IsObject(avTarget) And IsObject(oItem) Then
                    If avTarget Is oItem Then boFlg = True
                ElseIf Not IsObject(avTarget) And Not IsObject(oItem) Then
                    If VarType(avTarget) = vbString And VarType(oItem) = vbString Then
                        If Strcomp(avTarget, oItem, alCompare)=0 Then boFlg = True
                    Else
                        If avTarget = oItem Then boFlg = True
                    End If
                End If
                
                If boFlg Then
                    func_CmArrayIndexOf = lIndex
                    Exit For
                End If
                
            Next
        End If
        
        Set oItem = Nothing
        Set oArray = Nothing
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : func_CmArrayMap()
    'Overview                    : �z�񂩂�����̊֐��Ő��������z���Ԃ�
    'Detailed Description        : �����̊֐��̈����͈ȉ��̂Ƃ���
    '                                element :�z��̗v�f
    '                                index   :�C���f�b�N�X
    '                                array   :�z��
    'Argument
    '     aoFunc                 : �֐�
    'Return Value
    '     ���N���X�̕ʃC���X�^���X
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/16         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function func_CmArrayMap( _
        byRef aoFunc _
        )
        Dim oItem, lIndex, oArray, oRet
        
        '�z��̑S�Ă̗v�f�ɂ��Ĉ����̊֐��̏������s��
        If PoArray.Count>0 Then
            oArray = func_CmArrayConvArray(True)
            For lIndex=0 To PoArray.Count-1
                Call sub_CM_Bind(oItem, oArray(lIndex))
                Call sub_CM_Push(oRet, aoFunc(oItem, lIndex, oArray))
            Next
        End If
        
        Call sub_CM_Bind(func_CmArrayMap, new_ArraySetData(oRet))
        
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
        Dim oEle, lCount
        Set oEle = Nothing
        lCount = PoArray.Count
        If lCount>0 Then
            Call sub_CM_Bind(oEle, PoArray.Item(lCount-1))
            PoArray.Remove lCount-1
        End If
        Call sub_CM_Bind(func_CmArrayPop, oEle)
        Set oEle = Nothing
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : func_CmArrayPushMulti()
    'Overview                    : �z��̖����ɗv�f�𕡐��ǉ�����
    'Detailed Description        : �H����
    'Argument
    '     aoEles                 : �z��̖����ɒǉ�����v�f�i�z��j
    'Return Value
    '     �z����̗v�f��
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/08         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function func_CmArrayPushMulti( _
        byRef aoEles     _
        )
        If IsArray(aoEles    ) Then
            Dim oItem
            For Each oItem In aoEles    
                Call sub_CM_BindAt(PoArray, PoArray.Count, oItem)
            Next
        End If
        func_CmArrayPushMulti = PoArray.Count
        Set oItem = Nothing
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : func_CmArrayReduce()
    'Overview                    : �z��̂��ꂼ��̗v�f�ɑ΂��Ĉ����̊֐��ŎZ�o�������ʂ�Ԃ�
    'Detailed Description        : �����̊֐��̈����͈ȉ��̂Ƃ���
    '                                previousValue :1�O�̔z��̗v�f
    '                                currentValue  :�z��̗v�f
    '                                index         :�C���f�b�N�X
    '                                array         :�z��
    'Argument
    '     aoFunc                 : �֐�
    '     aboOrder               : True�F�����i���Ԃǂ���j / False�F�t��
    'Return Value
    '     �����̊֐��ŎZ�o��������
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/16         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function func_CmArrayReduce( _
        byRef aoFunc _
        , byVal aboOrder _
        )
        Dim oItem, lIndex, oArray, oRet
        
        '�z��̑S�Ă̗v�f�ɂ��Ĉ����̊֐��̏������s��
        If PoArray.Count>0 Then
            oArray = func_CmArrayConvArray(aboOrder)
            Call sub_CM_Bind(oRet, oArray(0))
            For lIndex=1 To PoArray.Count-1
                Call sub_CM_Bind(oItem, oArray(lIndex))
                Call sub_CM_Bind(oRet, aoFunc(oRet, oItem, lIndex, oArray))
            Next
        End If
        
        Call sub_CM_Bind(func_CmArrayReduce, oRet)
        
        Set oRet = Nothing
        Set oItem = Nothing
        Set oArray = Nothing
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : func_CmArrayReverse()
    'Overview                    : �z��̗v�f���t���ɕ��ׂ�
    'Detailed Description        : �H����
    'Argument
    '     �Ȃ�
    'Return Value
    '     �Ȃ�
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/14         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function func_CmArrayReverse( _
        )
        If PoArray.Count>0 Then
            Set PoArray = func_CmArrayAddDictionary(func_CmArrayConvArray(False), 0)
        End If
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
        If PoArray.Count>0 Then
            '�z�񂩂��菜�����v�f��Ԃ�
            Call sub_CM_Bind(func_CmArrayShift, PoArray.Item(0))
            '�쐬�����z��i�f�B�N�V���i���j��u����
            Set PoArray = func_CmArrayAddDictionary(func_CmArrayConvArray(True), 1)
        End If
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : func_CmArraySlice()
    'Overview                    : �z��̈ꕔ��؂�o�����z���Ԃ�
    'Detailed Description        : �H����
    'Argument
    '     alStart                : �J�n�ʒu�̃C���f�b�N�X�ԍ��A���l�͍Ō�̗v�f�̂���̈ʒu������
    '                              �Ⴆ��-1�͍Ō�A-2�͍Ōォ��2�ڂ̗v�f�������B
    '     alEnd                  : �I���ʒu�̃C���f�b�N�X�ԍ��A���l��alStart�Ɠ���
    '                              �؂�o�����z��ɏI���ʒu�̗v�f�͊܂܂Ȃ�
    '                              vbNullString���w�肵���ꍇ�͐؂�o�����z��ɍŌ�̗v�f���܂߂�
    'Return Value
    '     ���N���X�̕ʃC���X�^���X
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/16         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function func_CmArraySlice( _
        byVal alStart _
        , byVal alEnd _
        )
        Dim oItem, lIndex, oArray, oRet, lStart, lEnd
        
        '�z��̈ꕔ��؂�o��
        If PoArray.Count>0 Then
            If alStart<0 Then lStart = PoArray.Count + alStart Else lStart = alStart
            If alEnd = vbNullString Then
                lEnd = PoArray.Count - 1
            Else
                If alEnd<0 Then lEnd = PoArray.Count + alEnd - 1 Else lEnd = alEnd - 1
            End If
            oArray = func_CmArrayConvArray(True)
            For lIndex=lStart To lEnd
                Call sub_CM_Bind(oItem, oArray(lIndex))
                Call sub_CM_Push(oRet, oItem)
            Next
        End If
        
        Set func_CmArraySlice = new_ArraySetData(oRet)
        
        Set oItem = Nothing
        Set oArray = Nothing
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : func_CmArrayUnshiftMulti()
    'Overview                    : �z��̐擪�ɗv�f�𕡐��ǉ�����
    'Detailed Description        : �H����
    'Argument
    '     aoEles                 : �z��̐擪�ɒǉ�����v�f�i�z��j
    'Return Value
    '     �z����̗v�f��
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/08         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function func_CmArrayUnshiftMulti( _
        byRef aoEles     _
        )
        Dim oArray, oItem
        Set oArray = new_Dictionary()
        
        If IsArray(aoEles    ) Then
        '�����̗v�f��擪�ɒǉ�
            Set oArray = func_CmArrayAddDictionary(aoEles    , 0)
        End If
        
        '�����č�����v�f��ǉ�
        For Each oItem In func_CmArrayConvArray(True)
            Call sub_CM_BindAt(oArray, oArray.Count, oItem)
        Next
        
        '�쐬�����z��i�f�B�N�V���i���j��u����
        Set PoArray = oArray
        func_CmArrayUnshiftMulti = PoArray.Count
        
        Set oItem = Nothing
        Set oArray = Nothing
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : func_CmArrayConvArray()
    'Overview                    : �����ŕێ�����z��i�f�B�N�V���i���j���v���~�e�B�u�̔z��ɕϊ�����
    'Detailed Description        : �H����
    'Argument
    '     aboOrder               : True�F�����i���Ԃǂ���j / False�F�t��
    'Return Value
    '     �z��
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/08         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function func_CmArrayConvArray( _
        aboOrder _
        )
        Dim oItem, lIndex, oArray, oRet, lStt, lEnd, lStep
        
        '�z��̑S�Ă̗v�f
        If PoArray.Count>0 Then
            oArray = PoArray.Items()
            If aboOrder Then
                lStt = 0 : lEnd = PoArray.Count-1 : lStep = 1
            Else
                lStt = PoArray.Count-1 : lEnd = 0 : lStep = -1
            End If
            
            For lIndex=lStt To lEnd Step lStep
                Call sub_CM_Bind(oItem, oArray(lIndex))
                Call sub_CM_Push(oRet, oItem)
            Next
        End If
        
        func_CmArrayConvArray = oRet
        
        Set oItem = Nothing
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : func_CmArrayAddDictionary()
    'Overview                    : �����̔z��̓��e��z��i�f�B�N�V���i���j�ɒǉ�����
    'Detailed Description        : �H����
    'Argument
    '     aoEles                 : �z��i�f�B�N�V���i���j�ɒǉ�����v�f�i�z��j
    '     alStart                : �J�n�ʒu�̃C���f�b�N�X�ԍ��A���l�͍Ō�̗v�f�̂���̈ʒu������
    '                              �Ⴆ��-1�͍Ō�A-2�͍Ōォ��2�ڂ̗v�f�������B
    'Return Value
    '     �z��i�f�B�N�V���i���j
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/23         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function func_CmArrayAddDictionary( _
        byRef aoEles     _
        , byVal alStart _
        )
        Dim oArray, lStart, lIndex, lUb
        
        lUb = Ubound(aoEles    )
        If alStart<0 Then lStart = lUb + alStart Else lStart = alStart
        Set oArray = new_Dictionary()
        
        For lIndex = alStart To lUb
            Call sub_CM_BindAt(oArray, oArray.Count, aoEles    (lIndex))
        Next
        
        '�쐬�����z��i�f�B�N�V���i���j��Ԃ�
        Set func_CmArrayAddDictionary = oArray
        
        Set oArray = Nothing
    End Function
    
End Class
