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
Class clsCmArray
    '�N���X���ϐ��A�萔
    Private PoArr

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
        Set PoArr = new_Dic()
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
        Set PoArr = Nothing
    End Sub

    '***************************************************************************************************
    'Function/Sub Name           : Property Get count()
    'Overview                    : �z����̗v�f����Ԃ�
    'Detailed Description        : �H����
    'Argument
    '     �Ȃ�
    'Return Value
    '     �z��̗v�f��
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/09         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get count()
        count = PoArr.Count
    End Property

    '***************************************************************************************************
    'Function/Sub Name           : Property Get item()
    'Overview                    : �z��̎w�肵���C���f�b�N�X�̗v�f��Ԃ�
    'Detailed Description        : func_CmArrayItem()�ɈϏ�����
    'Argument
    '     alIdx                  : �C���f�b�N�X
    'Return Value
    '     �w�肵���C���f�b�N�X�̗v�f
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/08         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Default Property Get item( _
        byVal alIdx _
        )
        cf_bind item, func_CmArrayItem(alIdx)
    End Property

    '***************************************************************************************************
    'Function/Sub Name           : Property Set item()
    'Overview                    : �z��̎w�肵���C���f�b�N�X�ɗv�f��ݒ肷��
    'Detailed Description        : �H����
    'Argument
    '     alIdx                  : �C���f�b�N�X
    '     aoEle                  : �ݒ肷��v�f
    'Return Value
    '     �Ȃ�
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/08         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Set item( _
        byVal alIdx _
        , byRef aoEle _
        )
        If func_CmArrayInspectIndex(alIdx) Then
            cf_bindAt PoArr, alIdx, aoEle
        Else
            Err.Raise 9, "clsCmArray.vbs:clsCmArray+item()", "�C���f�b�N�X���L���͈͂ɂ���܂���B"
        End If
    End Property

    '***************************************************************************************************
    'Function/Sub Name           : Property Let item()
    'Overview                    : �z��̎w�肵���C���f�b�N�X�ɗv�f��ݒ肷��
    'Detailed Description        : �H����
    'Argument
    '     alIdx                  : �C���f�b�N�X
    '     aoEle                  : �ݒ肷��v�f
    'Return Value
    '     �Ȃ�
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/08         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Let item( _
        byVal alIdx _
        , byRef aoEle _
        )
        If func_CmArrayInspectIndex(alIdx) Then
            cf_bindAt PoArr, alIdx, aoEle
        Else
            Err.Raise 9, "clsCmArray.vbs:clsCmArray+item()", "�C���f�b�N�X���L���͈͂ɂ���܂���B"
        End If
    End Property

    '***************************************************************************************************
    'Function/Sub Name           : Property Get items()
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
    Public Property Get items( _
        )
        items = func_CmArrayConvArray(True)
    End Property

    '***************************************************************************************************
    'Function/Sub Name           : Property Get length()
    'Overview                    : �z����̗v�f����Ԃ�
    'Detailed Description        : �H����
    'Argument
    '     �Ȃ�
    'Return Value
    '     �z��̗v�f��
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/08         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get length()
        length = PoArr.Count
    End Property

    '***************************************************************************************************
    'Function/Sub Name           : concat()
    'Overview                    : �����Ŏw�肵���v�f��A�������z���Ԃ�
    'Detailed Description        : ���g�̃C���X�^���X�͕ύX���Ȃ�
    'Argument
    '     avArr                  : �z��ɒǉ�����v�f�i�z��j
    'Return Value
    '     ���N���X�̕ʃC���X�^���X
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/14         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function concat( _
        byRef avArr _
        )
        Dim oArr : Set oArr = new_Arr()
        If PoArr.Count>0 Then
            oArr.pushMulti func_CmArrayConvArray(True)
        End If
        oArr.pushMulti avArr
        Set concat = oArr

        Set oArr = Nothing
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : every()
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
    Public Function every( _
        byRef aoFunc _
        )
        every = func_CmArrayEveryOrSome(aoFunc, True)
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : filter()
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
    Public Function filter( _
        byRef aoFunc _
        )
        Set filter = func_CmArrayFilter(aoFunc)
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : find()
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
    Public Function find( _
        byRef aoFunc _
        )
        cf_bind find, func_CmArrayFind(aoFunc)
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : forEach()
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
    Public Sub forEach( _
        byRef aoFunc _
        )
        Call func_CmArrayForEach(aoFunc)
    End Sub

    '***************************************************************************************************
    'Function/Sub Name           : hasElements()
    'Overview                    : �z�񂪗v�f���܂ނ���������
    'Detailed Description        : func_CmArrayHasElement()�ɈϏ�����
    '                              ������Ԃ̔z���False��Ԃ�
    'Argument
    '     acArr                  : �z��
    'Return Value
    '     ���� True:�v�f���܂� / False:�v�f���܂܂Ȃ�
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/15         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function hasElement( _
        byRef acArr _
        )
        hasElement = func_CmArrayHasElement(acArr)
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : indexOf()
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
    Public Function indexOf( _
        byRef avTarget _
        )
        indexOf = func_CmArrayIndexOf(avTarget, vbNullString, vbBinaryCompare, True)
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : join()
    'Overview                    : �z��̊e�v�f��A��������������쐬����
    'Detailed Description        : vbscript��Join�֐��Ɠ����̋@�\
    'Argument
    '     asDel                  : ��؂蕶��
    'Return Value
    '     �z��̊e�v�f��A������������
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/08         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function join( _
        byVal asDel _
        )
        If PoArr.Count>0 Then
            join = func_CM_UtilJoin(func_CmArrayConvArray(True), asDel)
        Else
            join = ""
        End If
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : lastIndexOf()
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
    Public Function lastIndexOf( _
        byRef avTarget _
        )
        lastIndexOf = func_CmArrayIndexOf(avTarget, vbNullString, vbBinaryCompare, False)
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : map()
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
    Public Function map( _
        byRef aoFunc _
        )
        cf_bind map, func_CmArrayMap(aoFunc)
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : pop()
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
    Public Function pop( _
        )
        cf_bind pop, func_CmArrayPop()
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : push()
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
    Public Function push( _
        byRef aoEle _
        )
        push = func_CmArrayPushMulti(Array(aoEle))
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : pushMulti()
    'Overview                    : �z��̖����ɗv�f�𕡐��ǉ�����
    'Detailed Description        : func_CmArrayPushMulti()�ɈϏ�����
    'Argument
    '     avArr                  : �z��̖����ɒǉ�����v�f�i�z��j
    'Return Value
    '     �z����̗v�f��
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/08         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function pushMulti( _
        byRef avArr _
        )
        pushMulti = func_CmArrayPushMulti(avArr)
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : reduce()
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
    Public Function reduce( _
        byRef aoFunc _
        )
        cf_bind reduce, func_CmArrayReduce(aoFunc, True)
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : reduceRight()
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
    Public Function reduceRight( _
        byRef aoFunc _
        )
        cf_bind reduceRight, func_CmArrayReduce(aoFunc, False)
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : reverse()
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
    Public Sub reverse( _
        )
        Call func_CmArrayReverse()
    End Sub

    '***************************************************************************************************
    'Function/Sub Name           : shift()
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
    Public Function shift( _
        )
        cf_bind shift, func_CmArrayShift()
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : slice()
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
    Public Function slice( _
        byVal alStart _
        , byVal alEnd _
        )
        Set slice = func_CmArraySlice(alStart, alEnd)
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : some()
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
    Public Function some( _
        byRef aoFunc _
        )
        some = func_CmArrayEveryOrSome(aoFunc, False)
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : sort()
    'Overview                    : �z��̗v�f���\�[�g����
    'Detailed Description        : func_CmArraySort()�ɈϏ�����
    'Argument
    '     aboOrder               : True:���� / False:�~��
    'Return Value
    '     �\�[�g��̎��g�̃C���X�^���X
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/09         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function sort( _
        byVal aboOrder _
        )
        Set sort = func_CmArraySort(Getref("func_CM_UtilSortDefaultFunc"), aboOrder)
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : sortUsing()
    'Overview                    : �w�肵���֐����g���Ĕz��̗v�f���\�[�g����
    'Detailed Description        : func_CmArraySort()�ɈϏ�����
    'Argument
    '     aoFunc                 : ���肷��֐�
    'Return Value
    '     �\�[�g��̎��g�̃C���X�^���X
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/23         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function sortUsing( _
        byRef aoFunc _
        )
        Set sortUsing = func_CmArraySort(aoFunc, True)
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : splice()
    'Overview                    : �z��̗v�f�̑}���A�폜�A�u�����s��
    'Detailed Description        : func_CmArraySplice()�ɈϏ�����
    'Argument
    '     alStart                : �J�n�ʒu�̃C���f�b�N�X�ԍ��A���l�͍Ō�̗v�f�̂���̈ʒu������
    '                              �Ⴆ��-1�͍Ō�A-2�͍Ōォ��2�ڂ̗v�f�������B
    '     alDelCnt               : �J�n�ʒu����폜����v�f��
    '                              0�̏ꍇ�͍폜���Ȃ�
    '     avArr                  : �J�n�ʒu�ɒǉ�����v�f�i�z��j
    'Return Value
    '     �폜�����z�񂪂���΁A�폜�����z��̓��N���X�̕ʃC���X�^���X
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/23         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function splice( _
        byVal alStart _
        , byVal alDelCnt _
        , byRef avArr _
        )
        Set splice = func_CmArraySplice(alStart, alDelCnt, avArr)
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : unshift()
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
    Public Function unshift( _
        byRef aoEle _
        )
        unshift = func_CmArrayUnshiftMulti(Array(aoEle))
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : unshiftMulti()
    'Overview                    : �z��̐擪�ɗv�f��1�ǉ�����
    'Detailed Description        : func_CmArrayUnshiftMulti()�ɈϏ�����
    'Argument
    '     avArr                  : �z��̐擪�ɒǉ�����v�f�i�z��j
    'Return Value
    '     �z����̗v�f��
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/08         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function unshiftMulti( _
        byRef avArr _
        )
        unshiftMulti = func_CmArrayUnshiftMulti(avArr)
    End Function





    '***************************************************************************************************
    'Function/Sub Name           : func_CmArrayItem()
    'Overview                    : �z��̎w�肵���C���f�b�N�X�̗v�f��Ԃ�
    'Detailed Description        : �H����
    'Argument
    '     alIdx                  : �C���f�b�N�X
    'Return Value
    '     �w�肵���C���f�b�N�X�̗v�f
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/08         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function func_CmArrayItem( _
        byVal alIdx _
        )
        If PoArr.Exists(alIdx) Then
            cf_bind func_CmArrayItem, PoArr.Item(alIdx)
        Else
            Err.Raise 9, "clsCmArray.vbs:clsCmArray-func_CmArrayItem()", "�C���f�b�N�X���L���͈͂ɂ���܂���B"
        End If
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
        Dim lIdx, vArr, lUb, boRet
        boRet = aboFlg

        '�����̊֐��Ŕ��肷��
        If PoArr.Count>0 Then
            vArr = func_CmArrayConvArray(True)
            lUb = Ubound(vArr)
            
            For lIdx=0 To lUb
                If Not aoFunc(vArr(lIdx), lIdx, vArr) = aboFlg Then
                    boRet = Not aboFlg
                    Exit For
                End If
            Next
        End If

        '���茋�ʂ�ԋp
        func_CmArrayEveryOrSome = boRet
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
        Dim lIdx, vArr, lUb, vRet

        '�����̊֐��Œ��o�����v�f�����̔z����쐬
        If PoArr.Count>0 Then
            vArr = func_CmArrayConvArray(True)
            lUb = Ubound(vArr)
            
            For lIdx=0 To lUb
                If aoFunc(vArr(lIdx), lIdx, vArr) Then
                    cf_push vRet, vArr(lIdx)
                End If
            Next
        End If
        
        '�쐬�����z��i�f�B�N�V���i���j�œ��N���X�̃C���X�^���X�𐶐����ĕԋp
        If func_CmArrayHasElement(vRet) Then
            Set func_CmArrayFilter = new_ArrWith(vRet)
        Else
            Set func_CmArrayFilter = new_Arr()
        End If
        
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : func_CmArrayFind()
    'Overview                    : �����̊֐��Œ��o�����ŏ��̗v�f��Ԃ�
    'Detailed Description        : ���o�ł��Ȃ��ꍇ��Empty��Ԃ�
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
        Dim lIdx, vArr, lUb, oRet
        oRet = Empty

        '�����̊֐��Œ��o�ł���ŏ��̗v�f������
        If PoArr.Count>0 Then
            vArr = func_CmArrayConvArray(True)
            lUb = Ubound(vArr)
            
            For lIdx=0 To lUb
                If aoFunc(vArr(lIdx), lIdx, vArr) Then
                    cf_bind oRet, vArr(lIdx)
                    Exit For
                End If
            Next
        End If

        '�z�񂩂璊�o�����v�f��ԋp
        cf_bind func_CmArrayFind, oRet

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
        Dim lIdx, vArr, lUb

        '�z��̑S�Ă̗v�f�ɂ��Ĉ����̊֐��̏������s��
        If PoArr.Count>0 Then
            vArr = func_CmArrayConvArray(True)
            lUb = Ubound(vArr)
            
            For lIdx=0 To lUb
                Call aoFunc(vArr(lIdx), lIdx, vArr)
            Next
            Set PoArr = func_CmArrayAddDictionary(vArr, 0)
        End If
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : func_CmArrayHasElement()
    'Overview                    : �z�񂪗v�f���܂ނ���������
    'Detailed Description        : ������Ԃ̔z���False��Ԃ�
    'Argument
    '     acArr                  : �z��
    'Return Value
    '     ���� True:�v�f���܂� / False:�v�f���܂܂Ȃ�
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/15         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function func_CmArrayHasElement( _
        byRef acArr _
        )
        func_CmArrayHasElement = False
        If IsArray(acArr) Then
            On Error Resume Next
            Ubound(acArr)
            If Err.Number=0 Then func_CmArrayHasElement = True
            On Error Goto 0
        End If

'        func_CmArrayHasElement = False
'        If IsArray(avArray) Then func_CmArrayHasElement = cf_tryCatch(Getref("func_CM_ArrayUbound"), avArray, Empty, Empty).Item("Result")
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
        Dim lIdx, vArr, lUb, lStart, lEnd, lStep

        '�z��̑S�Ă̗v�f�ɂ��Ĉ����̊֐��̏������s��
        If PoArr.Count>0 Then
            vArr = func_CmArrayConvArray(True)
            lUb = Ubound(vArr)
            
            If alStart=vbNullString Then
                If aboOrder Then lStart=0 Else lStart=lUb
            Else
                lStart=alStart
            End If
            If aboOrder Then lEnd=lUb Else lEnd=0
            If aboOrder Then lStep=1 Else lStep=-1

            For lIdx=lStart To lEnd Step lStep
                If func_CM_UtilIsSame(avTarget, vArr(lIdx)) Then
                    func_CmArrayIndexOf = lIdx
                    Exit For
                End If
            Next
        End If
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
        Dim lIdx, vArr, lUb, vRet

        '�z��̑S�Ă̗v�f�ɂ��Ĉ����̊֐��̏������s��
        If PoArr.Count>0 Then
            vArr = func_CmArrayConvArray(True)
            lUb = Ubound(vArr)
            
            For lIdx=0 To lUb
                cf_push vRet, aoFunc(vArr(lIdx), lIdx, vArr)
            Next
        End If
        
        If func_CmArrayHasElement(vRet) Then
            Set func_CmArrayMap = new_ArrWith(vRet)
        Else
            Set func_CmArrayMap = new_Arr()
        End If
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
        lCount = PoArr.Count
        If lCount>0 Then
            cf_bind oEle, PoArr.Item(lCount-1)
            PoArr.Remove lCount-1
        End If
        cf_bind func_CmArrayPop, oEle
        Set oEle = Nothing
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : func_CmArrayPushMulti()
    'Overview                    : �z��̖����ɗv�f�𕡐��ǉ�����
    'Detailed Description        : �H����
    'Argument
    '     avArr                  : �z��̖����ɒǉ�����v�f�i�z��j
    'Return Value
    '     �z����̗v�f��
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/08         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function func_CmArrayPushMulti( _
        byRef avArr _
        )
        If func_CmArrayHasElement(avArr) Then
            Dim oEle
            For Each oEle In avArr
                cf_bindAt PoArr, PoArr.Count, oEle
            Next
            Set oEle = Nothing
        Elseif Not IsArray(avArr) Then
            cf_bindAt PoArr, PoArr.Count, avArr
        End If
        func_CmArrayPushMulti = PoArr.Count
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
        Dim lIdx, vArr, lUb, oRet
        oRet = Empty

        '�z��̑S�Ă̗v�f�ɂ��Ĉ����̊֐��̏������s��
        If PoArr.Count>0 Then
            vArr = func_CmArrayConvArray(aboOrder)
            lUb = Ubound(vArr)
            
            cf_bind oRet, vArr(0)
            For lIdx=1 To lUb
                cf_bind oRet, aoFunc(oRet, vArr(lIdx), lIdx, vArr)
            Next
            
            cf_bind func_CmArrayReduce, oRet
            Set oRet = Nothing
        Else
            Err.Raise 9, "clsCmArray.vbs:clsCmArray-func_CmArrayReduce()", "�z��̏����l������܂���B"
        End If

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
        If PoArr.Count>0 Then
            Set PoArr = func_CmArrayAddDictionary(func_CmArrayConvArray(False), 0)
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
        If PoArr.Count>0 Then
            '�z�񂩂��菜�����v�f��Ԃ�
            cf_bind func_CmArrayShift, PoArr.Item(0)
            '�쐬�����z��i�f�B�N�V���i���j��u����
            Set PoArr = func_CmArrayAddDictionary(func_CmArrayConvArray(True), 1)
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
        Dim lIdx, vArr, lUb, vRet, lStart, lEnd

        '�z��̈ꕔ��؂�o��
        If PoArr.Count>0 Then
            vArr = func_CmArrayConvArray(True)
            lUb = Ubound(vArr)
            
            If alStart<0 Then lStart=lUb+1 Else lStart=0
            lStart = math_max(lStart+alStart,0)
            lStart = math_min(lStart,lUb+1)
            
            if alEnd=vbNullString Then
                lEnd = lUb
            Else
                If alEnd<0 Then lEnd=lUb Else lEnd=-1
                lEnd = math_max(lEnd+alEnd,-1)
                lEnd = math_min(lEnd,lUb)
            End If
            
            For lIdx=lStart To lEnd
                cf_push vRet, vArr(lIdx)
            Next
        End If
        
        If func_CmArrayHasElement(vRet) Then
            Set func_CmArraySlice = new_ArrWith(vRet)
        Else
            Set func_CmArraySlice = new_Arr()
        End If
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : func_CmArraySort()
    'Overview                    : �w�肵���֐����g���Ĕz��̗v�f���\�[�g����
    'Detailed Description        : �\�[�g������func_CM_UtilSortHeap()�ɈϏ�����
    '                              �����̊֐��̈����͈ȉ��̂Ƃ���
    '                                currentValue :�z��̗v�f
    '                                nextValue    :���̔z��̗v�f
    'Argument
    '     aoFunc                 : �֐�
    '     aboOrder               : True:���� / False:�~��
    'Return Value
    '     �\�[�g��̎��g�̃C���X�^���X
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/14         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function func_CmArraySort( _
        byRef aoFunc _
        , byVal aboOrder _
        )
        If PoArr.Count>0 Then
            Set PoArr = func_CmArrayAddDictionary(func_CM_UtilSortHeap(func_CmArrayConvArray(True), aoFunc, aboOrder), 0)
        End If
        Set func_CmArraySort = Me
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : func_CmArraySplice()
    'Overview                    : �z��̗v�f�̑}���A�폜�A�u�����s��
    'Detailed Description        : �H����
    'Argument
    '     alStart                : �J�n�ʒu�̃C���f�b�N�X�ԍ��A���l�͍Ō�̗v�f�̂���̈ʒu������
    '                              �Ⴆ��-1�͍Ō�A-2�͍Ōォ��2�ڂ̗v�f�������B
    '     alDelCnt               : �J�n�ʒu����폜����v�f��
    '                              0�̏ꍇ�͍폜���Ȃ�
    '     avArr                  : �J�n�ʒu�ɒǉ�����v�f�i�z��j
    'Return Value
    '     �폜�����z�񂪂���΁A�폜�����z��̓��N���X�̕ʃC���X�^���X
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/23         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function func_CmArraySplice( _
        byVal alStart _
        , byVal alDelCnt _
        , byRef avArr _
        )
        Dim lIdx, vArr, lUb, vArrayAft(), vRet(), lStart

        If PoArr.Count>0 Then
            vArr = func_CmArrayConvArray(True)
            lUb = Ubound(vArr)
            
            If alStart<0 Then lStart=lUb+1 Else lStart=0
            lStart = math_max(lStart+alStart,0)
            lStart = math_min(lStart,lUb+1)
            
            For lIdx = 0 To lStart - 1
            '�J�n�ʒu�܂ł͍��̔z��̂܂�
                cf_push vArrayAft, vArr(lIdx)
            Next
            
            For lIdx = lStart To math_min(lStart+alDelCnt-1, lUb)
            '�J�n�ʒu����폜����v�f���͖߂�l�̔z��Ɉڂ�
                cf_push vRet, vArr(lIdx)
            Next
        End If
        
        If func_CmArrayHasElement(avArr) Then
        '�ǉ�����v�f������Βǉ�����
            For lIdx = 0 To Ubound(avArr)
                cf_push vArrayAft, avArr(lIdx)
            Next
        End If
        
        If PoArr.Count>0 Then
            For lIdx = lStart+alDelCnt To lUb
            '�폜�����v�f�ȍ~�͍��̔z��Ɏc��
                cf_push vArrayAft, vArr(lIdx) 
            Next
        End If
        
        If func_CmArrayHasElement(vArrayAft) Then
            '�쐬�����z��i�f�B�N�V���i���j��u����
            Set PoArr = func_CmArrayAddDictionary(vArrayAft, 0)
        End If
        
        '�z�񂩂��菜�����v�f��Ԃ�
        If func_CmArrayHasElement(vRet) Then
            Set func_CmArraySplice = new_ArrWith(vRet)
        Else
            Set func_CmArraySplice = new_Arr()
        End If
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : func_CmArrayUnshiftMulti()
    'Overview                    : �z��̐擪�ɗv�f�𕡐��ǉ�����
    'Detailed Description        : �H����
    'Argument
    '     avArr                  : �z��̐擪�ɒǉ�����v�f�i�z��j
    'Return Value
    '     �z����̗v�f��
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/08         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function func_CmArrayUnshiftMulti( _
        byRef avArr _
        )
        Dim oArr, oEle
        Set oArr = new_Dic()

        If func_CmArrayHasElement(avArr) Then
        '�����̗v�f��擪�ɒǉ�
            Set oArr = func_CmArrayAddDictionary(avArr, 0)
        End If

        '�����č�����v�f��ǉ�
        For Each oEle In func_CmArrayConvArray(True)
            cf_bindAt oArr, oArr.Count, oEle
        Next

        '�쐬�����z��i�f�B�N�V���i���j��u����
        Set PoArr = oArr
        func_CmArrayUnshiftMulti = PoArr.Count

        Set oEle = Nothing
        Set oArr = Nothing
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
        Dim lIdx, vArr, vRet(), lStt, lEnd, lStep

        '�z��̑S�Ă̗v�f
        If PoArr.Count>0 Then
            vArr = PoArr.Items()
            
            If aboOrder Then
                lStt = 0 : lEnd = PoArr.Count-1 : lStep = 1
            Else
                lStt = PoArr.Count-1 : lEnd = 0 : lStep = -1
            End If

            For lIdx=lStt To lEnd Step lStep
                cf_push vRet, vArr(lIdx)
            Next
        End If

        func_CmArrayConvArray = vRet
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : func_CmArrayAddDictionary()
    'Overview                    : �����̔z��̓��e��z��i�f�B�N�V���i���j�ɒǉ�����
    'Detailed Description        : �H����
    'Argument
    '     avArr                  : �z��i�f�B�N�V���i���j�ɒǉ�����v�f�i�z��j
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
        byRef avArr _
        , byVal alStart _
        )
        Dim oArr, lStart, lIdx, lUb

        lUb = Ubound(avArr)
        If alStart<0 Then lStart = lUb + alStart Else lStart = alStart
        Set oArr = new_Dic()

        For lIdx = alStart To lUb
            cf_bindAt oArr, oArr.Count, avArr(lIdx)
        Next

        '�쐬�����z��i�f�B�N�V���i���j��Ԃ�
        Set func_CmArrayAddDictionary = oArr

        Set oArr = Nothing
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : func_CmArrayInspectIndex()
    'Overview                    : �C���f�b�N�X���L������������
    'Detailed Description        : �H����
    'Argument
    '     alIdx                  : �C���f�b�N�X
    'Return Value
    '     ���� True:�L���ȃC���f�b�N�X / False:�����ȃC���f�b�N�X
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/24         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function func_CmArrayInspectIndex( _
        byVal alIdx _
        )
        func_CmArrayInspectIndex = False
        If 0 <= alIdx And alIdx < PoArr.Count Then func_CmArrayInspectIndex = True
    End Function

End Class
