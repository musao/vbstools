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
    Private PvArr,PoBroker,PlCnt

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
        Set PoBroker = Nothing
        PlCnt = 0
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
        Set PoBroker = Nothing
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Set broker()
    'Overview                    : �u���[�J�[�N���X�̃I�u�W�F�N�g��ݒ肷��
    'Detailed Description        : �H����
    'Argument
    '     aoBroker               : �u���[�J�[�N���X�̃C���X�^���X
    'Return Value
    '     �Ȃ�
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2024/01/06         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Set broker( _
        byRef aoBroker _
        )
        Set PoBroker = aoBroker
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get broker()
    'Overview                    : �u���[�J�[�N���X�̃I�u�W�F�N�g��Ԃ�
    'Detailed Description        : �H����
    'Argument
    '     �Ȃ�
    'Return Value
    '     �u���[�J�[�N���X�̃C���X�^���X
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2024/01/06         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get broker()
        Set broker = PoBroker
    End Property

    '***************************************************************************************************
    'Function/Sub Name           : Property Get count()
    'Overview                    : �z����̗v�f����Ԃ�
    'Detailed Description        : this_length()�ɈϏ�����
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
        count = this_length()
    End Property

    '***************************************************************************************************
    'Function/Sub Name           : Property Get item()
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
    Public Default Property Get item( _
        byVal alIdx _
        )
        ast_argTrue this_isValidIndex(alIdx), TypeName(Me)&"+item() Get", "Index is out of range."
        cf_bind item, PvArr(alIdx)
    End Property

    '***************************************************************************************************
    'Function/Sub Name           : Property Set item()
    'Overview                    : �z��̎w�肵���C���f�b�N�X�ɗv�f��ݒ肷��
    'Detailed Description        : this_setItem()�ɈϏ�����
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
        this_setItem alIdx, aoEle, TypeName(Me)&"+item() Set"
    End Property

    '***************************************************************************************************
    'Function/Sub Name           : Property Let item()
    'Overview                    : �z��̎w�肵���C���f�b�N�X�ɗv�f��ݒ肷��
    'Detailed Description        : this_setItem()�ɈϏ�����
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
        this_setItem alIdx, aoEle, TypeName(Me)&"+item() Let"
    End Property

    '***************************************************************************************************
    'Function/Sub Name           : Property Get items()
    'Overview                    : �z���Ԃ�
    'Detailed Description        : this_toArray()�ɈϏ�����
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
        items = this_toArray(True)
    End Property

    '***************************************************************************************************
    'Function/Sub Name           : Property Get length()
    'Overview                    : �z����̗v�f����Ԃ�
    'Detailed Description        : this_length()�ɈϏ�����
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
        length = this_length()
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
        If this_length()>0 Then oArr.pushA PvArr
        oArr.pushA avArr
        Set concat = oArr

        Set oArr = Nothing
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : every()
    'Overview                    : �z��̑S�Ă̗v�f�������̊֐��̔���𖞂������m�F����
    'Detailed Description        : this_Every()�ɈϏ�����
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
        every = this_everyOrSome(aoFunc, True)
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : filter()
    'Overview                    : �����̊֐��Œ��o�����v�f�����̔z����쐬
    'Detailed Description        : this_filter()�ɈϏ�����
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
        Set filter = this_filter(aoFunc)
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : find()
    'Overview                    : �����̊֐��Œ��o�����ŏ��̗v�f��Ԃ�
    'Detailed Description        : this_find()�ɈϏ�����
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
        cf_bind find, this_find(aoFunc)
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : forEach()
    'Overview                    : �z��̑S�Ă̗v�f�ɂ��Ĉ����̊֐��̏������s��
    'Detailed Description        : this_forEach()�ɈϏ�����
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
        this_forEach aoFunc
    End Sub

    '***************************************************************************************************
    'Function/Sub Name           : hasElements()
    'Overview                    : �z�񂪗v�f���܂ނ���������
    'Detailed Description        : this_hasElement()�ɈϏ�����
    '                              ������Ԃ̔z���False��Ԃ�
    'Argument
    '     avArr                  : �z��
    'Return Value
    '     ���� True:�v�f���܂� / False:�v�f���܂܂Ȃ�
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/15         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function hasElement( _
        byRef avArr _
        )
        hasElement = this_hasElement(avArr)
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : indexOf()
    'Overview                    : �����ɍ��v����v�f�𐳏��ɒT���ŏ��Ɍ��������C���f�b�N�X�ԍ���Ԃ�
    'Detailed Description        : this_indexOf()�ɈϏ�����
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
        indexOf = this_indexOf(avTarget, vbNullString, vbBinaryCompare, True)
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
        join = ""
        If this_length()>0 Then join = func_CM_UtilJoin(PvArr, asDel)
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : lastIndexOf()
    'Overview                    : �����ɍ��v����v�f���t���ɒT���ŏ��Ɍ��������C���f�b�N�X�ԍ���Ԃ�
    'Detailed Description        : this_indexOf()�ɈϏ�����
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
        lastIndexOf = this_indexOf(avTarget, vbNullString, vbBinaryCompare, False)
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : map()
    'Overview                    : �z�񂩂�����̊֐��ŐV���Ȕz��𐶐�����
    'Detailed Description        : this_map()�ɈϏ�����
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
        cf_bind map, this_map(aoFunc)
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : pop()
    'Overview                    : �z�񂩂疖���̗v�f����菜��
    'Detailed Description        : this_pop()�ɈϏ�����
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
        cf_bind pop, this_pop()
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : push()
    'Overview                    : �z��̖����ɗv�f��1�ǉ�����
    'Detailed Description        : this_pushA()�ɈϏ�����
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
        push = this_pushA(Array(aoEle))
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : pushA()
    'Overview                    : �z��̖����ɗv�f�𕡐��ǉ�����
    'Detailed Description        : this_pushA()�ɈϏ�����
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
    Public Function pushA( _
        byRef avArr _
        )
        pushA = this_pushA(avArr)
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : reduce()
    'Overview                    : �z��̂��ꂼ��̗v�f�ɑ΂��Đ����Ɉ����̊֐��ŎZ�o�������ʂ�Ԃ�
    'Detailed Description        : this_reduce()�ɈϏ�����
    'Argument
    '     aoFunc                 : �֐�
    '     avInitial              : �����l
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
        , byRef avInitial _
        )
        cf_bind reduce, this_reduce(aoFunc, avInitial, True, TypeName(Me)&"+reduce()")
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : reduceRight()
    'Overview                    : �z��̂��ꂼ��̗v�f�ɑ΂��ċt���Ɉ����̊֐��ŎZ�o�������ʂ�Ԃ�
    'Detailed Description        : this_reduce()�ɈϏ�����
    'Argument
    '     aoFunc                 : �֐�
    '     avInitial              : �����l
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
        , byRef avInitial _
        )
        cf_bind reduceRight, this_reduce(aoFunc, avInitial, False, TypeName(Me)&"+reduceRight()")
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : reverse()
    'Overview                    : �z��̗v�f���t���ɕ��ׂ�
    'Detailed Description        : this_Reverse()�ɈϏ�����
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
        PvArr = this_toArray(False)
    End Sub

    '***************************************************************************************************
    'Function/Sub Name           : shift()
    'Overview                    : �z�񂩂�擪�̗v�f����菜��
    'Detailed Description        : this_shift()�ɈϏ�����
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
        cf_bind shift, this_shift()
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : slice()
    'Overview                    : �z��̈ꕔ��؂�o�����z��𐶐�����
    'Detailed Description        : this_slice()�ɈϏ�����
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
        Set slice = this_slice(alStart, alEnd)
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : some()
    'Overview                    : �z��̂����ꂩ��̗v�f�������̊֐��̔���𖞂������m�F����
    'Detailed Description        : this_Every()�ɈϏ�����
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
        some = this_everyOrSome(aoFunc, False)
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : sort()
    'Overview                    : �z��̗v�f���\�[�g����
    'Detailed Description        : this_sort()�ɈϏ�����
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
        Set sort = this_sort(Getref("func_CM_UtilSortDefaultFunc"), aboOrder)
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : sortUsing()
    'Overview                    : �w�肵���֐����g���Ĕz��̗v�f���\�[�g����
    'Detailed Description        : this_sort()�ɈϏ�����
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
        Set sortUsing = this_sort(aoFunc, True)
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : splice()
    'Overview                    : �z��̗v�f�̑}���A�폜�A�u�����s��
    'Detailed Description        : this_splice()�ɈϏ�����
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
        Set splice = this_splice(alStart, alDelCnt, avArr)
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : toArray()
    'Overview                    : �z���Ԃ�
    'Detailed Description        : �H����
    'Argument
    '     �Ȃ�
    'Return Value
    '     �z��
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2024/01/16         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function toArray( _
        )
        toArray = this_toArray(True)
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
    '2023/12/24         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function toString( _
        )
        toString = "<" & TypeName(Me) & ">[]"
        If this_length()=0 Then Exit Function

        Dim vRet, oEle
        For Each oEle In PvArr
            cf_push vRet, cf_toString(oEle)
        Next
        toString = "<" & TypeName(Me) & ">[" & func_CM_UtilJoin(vRet, ",") & "]"
        Set oEle = Nothing
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : uniq()
    'Overview                    : �z��̏d����r������
    'Detailed Description        : this_uniq()�ɈϏ�����
    'Argument
    '     �Ȃ�
    'Return Value
    '     ������̎��g�̃C���X�^���X
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/11/12         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function uniq( _
        )
        Set uniq = this_uniq()
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : unshift()
    'Overview                    : �z��̐擪�ɗv�f��1�ǉ�����
    'Detailed Description        : this_unshiftA()�ɈϏ�����
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
        unshift = this_unshiftA(Array(aoEle))
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : unshiftA()
    'Overview                    : �z��̐擪�ɗv�f��1�ǉ�����
    'Detailed Description        : this_unshiftA()�ɈϏ�����
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
    Public Function unshiftA( _
        byRef avArr _
        )
        unshiftA = this_unshiftA(avArr)
    End Function





    '***************************************************************************************************
    'Function/Sub Name           : this_comparison()
    'Overview                    : ��r����
    'Detailed Description        : �K�v�ɉ����ē��v�����擾����
    'Argument
    '     avEleA                 : ��r����v�fA
    '     avEleB                 : ��r����v�fB
    'Return Value
    '     ��r����
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2024/01/06         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_comparison( _
        byRef aoFunc _
        , byRef avEleA _
        , byRef avEleB _
        )
        If PoBroker Is Nothing Then
        '���v�����擾���Ȃ��ꍇ
            cf_bind this_comparison, aoFunc(avEleA, avEleB)
            Exit Function
        End If

        '���v�����擾����ꍇ
        Dim lCnt : lCnt = this_getCount()
        this_publish "event", Array("Comparison", lCnt, "0Start", avEleA, avEleB)
        Dim vRet
        cf_bind vRet, aoFunc(avEleA, avEleB)
        cf_bind this_comparison, vRet
        this_publish "event", Array("Comparison", lCnt, "1End", vRet)
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : this_everyOrSome()
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
    Private Function this_everyOrSome( _
        byRef aoFunc _
        , byRef aboFlg _
        )
        this_everyOrSome = aboFlg
        If this_length()=0 Then Exit Function
        
        Dim vArr, lUb, boRet
        vArr = PvArr
        lUb = Ubound(vArr)
        boRet = aboFlg
        
        '�����̊֐��Ŕ��肷��
        Dim lIdx
        For lIdx=0 To lUb
            If Not aoFunc(vArr(lIdx), lIdx, vArr) = boRet Then
                boRet = Not boRet
                Exit For
            End If
        Next

        '���茋�ʂ�ԋp
        this_everyOrSome = boRet
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : this_filter()
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
    Private Function this_filter( _
        byRef aoFunc _
        )
        Set this_filter = new_Arr()
        If this_length()=0 Then Exit Function

        Dim vArr, lUb, vRet
        vArr = PvArr
        lUb = Ubound(vArr)
        
        '�����̊֐��Œ��o�����v�f�����̔z����쐬
        Dim lIdx
        For lIdx=0 To lUb
            If aoFunc(vArr(lIdx), lIdx, vArr) Then
                cf_push vRet, vArr(lIdx)
            End If
        Next
        
        '�쐬�����z��œ��N���X�̃C���X�^���X�𐶐����ĕԋp
        If this_hasElement(vRet) Then Set this_filter = new_ArrWith(vRet)
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : this_find()
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
    Private Function this_find( _
        byRef aoFunc _
        )
        this_find = Empty
        If this_length()=0 Then Exit Function

        Dim vArr, lUb, oRet
        vArr = PvArr
        lUb = Ubound(vArr)
        oRet = Empty

        '�����̊֐��Œ��o�ł���ŏ��̗v�f������
        Dim lIdx
        For lIdx=0 To lUb
            If aoFunc(vArr(lIdx), lIdx, vArr) Then
                cf_bind oRet, vArr(lIdx)
                Exit For
            End If
        Next

        '�z�񂩂璊�o�����v�f��ԋp
        cf_bind this_find, oRet

        Set oRet = Nothing
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : this_forEach()
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
    Private Sub this_forEach( _
        byRef aoFunc _
        )
        If this_length()=0 Then Exit Sub

        Dim vArr, lUb
        vArr = PvArr
        lUb = Ubound(vArr)
        
        '�z��̑S�Ă̗v�f�ɂ��Ĉ����̊֐��̏������s��
        Dim lIdx
        For lIdx=0 To lUb
            aoFunc vArr(lIdx), lIdx, vArr
        Next
        PvArr = vArr
    End Sub

    '***************************************************************************************************
    'Function/Sub Name           : this_getCount()
    'Overview                    : �A�Ԏ擾
    'Detailed Description        : �H����
    'Argument
    '     �Ȃ�
    'Return Value
    '     �A��
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2024/01/06         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_getCount( _
        )
        PlCnt = PlCnt + 1 : this_getCount = PlCnt
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : this_hasElement()
    'Overview                    : �z�񂪗v�f���܂ނ���������
    'Detailed Description        : ������Ԃ̔z���False��Ԃ�
    'Argument
    '     avArr                  : �z��
    'Return Value
    '     ���� True:�v�f���܂� / False:�v�f���܂܂Ȃ�
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/15         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_hasElement( _
        byRef avArr _
        )
        this_hasElement = False
        If IsArray(avArr) Then
            On Error Resume Next
            Dim lUb : lUb = Ubound(avArr)
            If Err.Number=0 And lUb>=0 Then this_hasElement = True
            On Error Goto 0
        End If
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : this_indexOf()
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
    Private Function this_indexOf( _
        byRef avTarget _
        , byVal alStart _
        , byVal alCompare _
        , byVal aboOrder _
        )
        this_indexOf = -1
        If this_length()=0 Then Exit Function

        Dim vArr, lUb
        vArr = PvArr
        lUb = Ubound(vArr)
        
        Dim lStart
        If alStart=vbNullString Then
            If aboOrder Then lStart=0 Else lStart=lUb
        Else
            lStart=alStart
        End If

        Dim lEnd, lStep
        If aboOrder Then lEnd=lUb Else lEnd=0
        If aboOrder Then lStep=1 Else lStep=-1

        '�z��̑S�Ă̗v�f�ɂ��Ĉ����̊֐��̏������s��
        Dim lIdx
        For lIdx=lStart To lEnd Step lStep
            If cf_isSame(avTarget, vArr(lIdx)) Then
                this_indexOf = lIdx
                Exit For
            End If
        Next
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : this_map()
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
    Private Function this_map( _
        byRef aoFunc _
        )
        Set this_map = new_Arr()
        If this_length()=0 Then Exit Function

        Dim vArr, lUb, vRet
        vArr = PvArr
        lUb = Ubound(vArr)

        '�z��̑S�Ă̗v�f�ɂ��Ĉ����̊֐��̏������s��
        Dim lIdx
        For lIdx=0 To lUb
            cf_push vRet, aoFunc(vArr(lIdx), lIdx, vArr)
        Next
        
        '���������z��ō쐬�����V�����C���X�^���X��Ԃ�
        If this_hasElement(vRet) Then Set this_map = new_ArrWith(vRet)
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : this_publish()
    'Overview                    : �o�ŁiPublish�j����
    'Detailed Description        : �H����
    'Argument
    '     asTopic                : �g�s�b�N
    '     asCont                 : ���e
    'Return Value
    '     �Ȃ�
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2024/01/06         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub this_publish( _
        byVal asTopic _
        , byRef avCont _
        )
        PoBroker.publish asTopic, avCont
    End Sub

    '***************************************************************************************************
    'Function/Sub Name           : this_pop()
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
    Private Function this_pop( _
        )
        this_pop = Empty
        If this_length()=0 Then Exit Function

        Dim lUb : lUb = Ubound(PvArr)
        cf_bind this_pop, PvArr(lUb)
        Redim Preserve PvArr(lUb-1)
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : this_pushA()
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
    Private Function this_pushA( _
        byRef avArr _
        )
        cf_pushA PvArr, avArr
        this_pushA = this_length()
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : this_reduce()
    'Overview                    : �z��̂��ꂼ��̗v�f�ɑ΂��Ĉ����̊֐��ŎZ�o�������ʂ�Ԃ�
    'Detailed Description        : �����̊֐��̈����͈ȉ��̂Ƃ���
    '                                previousValue :1�O�̔z��̗v�f
    '                                currentValue  :�z��̗v�f
    '                                index         :�C���f�b�N�X
    '                                array         :�z��
    'Argument
    '     aoFunc                 : �֐�
    '     avInitial              : �����l
    '     aboOrder               : True�F�����i���Ԃǂ���j / False�F�t��
    '     asSource               : �\�[�X
    'Return Value
    '     �����̊֐��ŎZ�o��������
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/16         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_reduce( _
        byRef aoFunc _
        , byRef avInitial _
        , byVal aboOrder _
        , byVal asSource _
        )
        ast_argTrue this_length()>0, asSource, "Array has no elements."

        Dim vArr, lUb, oRet
        If aboOrder Then vArr = PvArr Else vArr = this_toArray(aboOrder)
        lUb = Ubound(vArr)
        If IsEmpty(avInitial) Then cf_bind oRet, vArr(0) Else cf_bind oRet, avInitial

        '�z��̑S�Ă̗v�f�ɂ��Ĉ����̊֐��̏������s��
        Dim lIdx
        For lIdx=1 To lUb
            cf_bind oRet, aoFunc(oRet, vArr(lIdx), lIdx, vArr)
        Next
        
        cf_bind this_reduce, oRet
        Set oRet = Nothing
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : this_shift()
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
    Private Function this_shift( _
        )
        If this_length()=0 Then Exit Function

        Dim vArr : vArr = PvArr
        '�z��̐擪�̗v�f��Ԃ�
        cf_bind this_shift, vArr(0)
        
        '�擪�̗v�f����菜��
        Dim lUb : lUb=Ubound(vArr)
        Redim vNewArr(lUb-1)

        Dim lIdx
        For lIdx=1 To lUb
            cf_bind vNewArr(lIdx-1), vArr(lIdx)
        Next
        PvArr = vNewArr
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : this_slice()
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
    Private Function this_slice( _
        byVal alStart _
        , byVal alEnd _
        )
        Set this_slice = new_Arr()
        If this_length()=0 Then Exit Function

        Dim vArr, lUb
        vArr = PvArr
        lUb = Ubound(vArr)

        Dim lStart
        If alStart<0 Then lStart=lUb+1 Else lStart=0
        lStart = math_max(lStart+alStart,0)
        lStart = math_min(lStart,lUb+1)
        
        Dim lEnd
        if alEnd=vbNullString Then
            lEnd = lUb
        Else
            If alEnd<0 Then lEnd=lUb Else lEnd=-1
            lEnd = math_max(lEnd+alEnd,-1)
            lEnd = math_min(lEnd,lUb)
        End If
        
        '�z��̈ꕔ��؂�o��
        Dim lIdx, vRet
        For lIdx=lStart To lEnd
            cf_push vRet, vArr(lIdx)
        Next
        
        '�z��̈ꕔ��؂�o�����z��ō쐬�����V�����C���X�^���X��Ԃ�
        If this_hasElement(vRet) Then Set this_slice = new_ArrWith(vRet)
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : this_sort()
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
    Private Function this_sort( _
        byRef aoFunc _
        , byVal aboOrder _
        )
'        this_sortBubble aoFunc, aboOrder
'        this_sortQuick aoFunc, aboOrder
        this_sortMerge aoFunc, aboOrder
'        this_sortHeap aoFunc, aboOrder
        
        Set this_sort = Me
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : this_sortBubble()
    'Overview                    : �o�u���\�[�g
    'Detailed Description        : �v�Z�񐔂�O(N^2)
    '                              �z��̗v�f���Ȃ��܂���1�̏ꍇ�͉������Ȃ�
    'Argument
    '     aoFunc                 : �֐�
    '     aboFlg                 : ������@
    '                                True  :�����i�֐��̌��ʂ�True�̏ꍇ�ɓ���ւ���j
    '                                False :�~���i�֐��̌��ʂ�False�̏ꍇ�ɓ���ւ���j
    'Return Value
    '     �Ȃ�
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2024/01/06         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub this_sortBubble( _
        byRef aoFunc _
        , byVal aboFlg _
        )
        If this_length()<2 Then Exit Sub
        Dim vArr : vArr = PvArr
        
        Dim lEnd, lPos
        lEnd = Ubound(vArr)
        Do While lEnd>0
            For lPos=0 To lEnd-1
                If this_comparison(aoFunc, vArr(lPos), vArr(lPos+1))=aboFlg Then
                'lPos�Ԗڂ̗v�f��(lPos+1)�Ԗڂ̗v�f�����ւ���
                    cf_swap vArr(lPos), vArr(lPos+1)
                End If
            Next
            lEnd = lEnd-1
        Loop
        PvArr = vArr
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : this_sortQuick()
    'Overview                    : �N�C�b�N�\�[�g
    'Detailed Description        : �v�Z�񐔂͕���O(N*logN)�A�ň���O(N^2)
    '                              �z��̗v�f���Ȃ��܂���1�̏ꍇ�͉������Ȃ�
    '                              �����̊֐��̈����͈ȉ��̂Ƃ���
    '                                currentValue :�z��̗v�f
    '                                nextValue    :���̔z��̗v�f
    'Argument
    '     aoFunc                 : �֐�
    '     aboFlg                 : ������@
    '                                True  :�����i�֐��̌��ʂ�True�̏ꍇ�ɓ���ւ���j
    '                                False :�~���i�֐��̌��ʂ�False�̏ꍇ�ɓ���ւ���j
    'Return Value
    '     �Ȃ�
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/18         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub this_sortQuick( _
        byRef aoFunc _
        , byVal aboFlg _
        )
        If this_length()<2 Then Exit Sub
        PvArr = this_sortQuickRecursion(PvArr, aoFunc, aboFlg)
    End Sub
    '***************************************************************************************************
    'Function/Sub Name           : this_sortQuickRecursion()
    'Overview                    : �N�C�b�N�\�[�g�̍ċA����
    'Detailed Description        : this_sortQuick()�Q��
    'Argument
    '     avArr                  : �z��
    '     aoFunc                 : �֐�
    '     aboFlg                 : ������@
    '                                True  :�����i�֐��̌��ʂ�True�̏ꍇ�ɓ���ւ���j
    '                                False :�~���i�֐��̌��ʂ�False�̏ꍇ�ɓ���ւ���j
    'Return Value
    '     �\�[�g��̔z��
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/18         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_sortQuickRecursion( _
        byRef avArr _
        , byRef aoFunc _
        , byVal aboFlg _
        )
        this_sortQuickRecursion = avArr
        If Not this_hasElement(avArr) Then Exit Function
        If Ubound(avArr)=0 Then Exit Function
        
        '0�Ԗڂ̗v�f���s�{�b�g�Ɍ��߂�
        Dim oPivot : cf_bind oPivot, avArr(0)
        
        '�s�{�b�g�Ɨv�f���֐��Ŕ��肵������@�ɍ��v����O���[�v��Right�A�����łȂ��O���[�v��Left�Ƃ���
        Dim lPos, vRight, vLeft
        For lPos=1 To Ubound(avArr)
            If this_comparison(aoFunc, avArr(lPos), oPivot)=aboFlg Then
                cf_push vRight, avArr(lPos)
            Else
                cf_push vLeft, avArr(lPos)
            End If
        Next
        
        '��q�ŕ�����Right�ALeft�̃O���[�v���ƂɍċA��������
        vLeft = this_sortQuickRecursion(vLeft, aoFunc, aboFlg)
        vRight = this_sortQuickRecursion(vRight, aoFunc, aboFlg)
        
        'Left�Ƀs�{�b�g�{Right����������
        cf_push vLeft, oPivot
        If this_hasElement(vRight) Then cf_pushA vLeft, vRight
        
        this_sortQuickRecursion = vLeft
        Set oPivot = Nothing
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : this_sortMerge()
    'Overview                    : �}�[�W�\�[�g
    'Detailed Description        : �v�Z�񐔂�O(N*logN)
    '                              �z��̗v�f���Ȃ��܂���1�̏ꍇ�͉������Ȃ�
    '                              �����̊֐��̈����͈ȉ��̂Ƃ���
    '                                currentValue :�z��̗v�f
    '                                nextValue    :���̔z��̗v�f
    'Argument
    '     aoFunc                 : �֐�
    '     aboFlg                 : ������@
    '                                True  :�����i�֐��̌��ʂ�True�̏ꍇ�ɓ���ւ���j
    '                                False :�~���i�֐��̌��ʂ�False�̏ꍇ�ɓ���ւ���j
    'Return Value
    '     �Ȃ�
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/18         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub this_sortMerge( _
        byRef aoFunc _
        , byVal aboFlg _
        )
        If this_length()<2 Then Exit Sub
        PvArr = this_sortMergeRecursion(PvArr, aoFunc, aboFlg)
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : this_sortMergeRecursion()
    'Overview                    : �}�[�W�\�[�g�̍ċA����
    'Detailed Description        : this_sortMerge()�Q��
    '                              �}�[�W������this_SortMergeMerge()�ɈϏ�����
    'Argument
    '     avArr                  : �z��
    '     aoFunc                 : �֐�
    '     aboFlg                 : ������@
    '                                True  :�����i�֐��̌��ʂ�True�̏ꍇ�ɓ���ւ���j
    '                                False :�~���i�֐��̌��ʂ�False�̏ꍇ�ɓ���ւ���j
    'Return Value
    '     �\�[�g��̔z��
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/18         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_sortMergeRecursion( _
        byRef avArr _
        , byRef aoFunc _
        , byVal aboFlg _
        )
        this_sortMergeRecursion = avArr
        If Not this_hasElement(avArr) Then Exit Function
        If Ubound(avArr)=0 Then Exit Function
        
        '2�̔z��ɕ�������
        Dim lLength, lMedian
        lLength = Ubound(avArr) - Lbound(avArr) + 1
        lMedian = math_roundUp(lLength/2, 0)
        Dim lPos, vFirst, vSecond
        For lPos=Lbound(avArr) To lMedian-1
            cf_push vFirst, avArr(lPos)
        Next
        For lPos=lMedian To Ubound(avArr)
            cf_push vSecond, avArr(lPos)
        Next
        
        '�ċA�����Ŕz��̗v�f��1�ɂȂ�܂ŕ�������
        vFirst = this_sortMergeRecursion(vFirst, aoFunc, aboFlg)
        vSecond = this_sortMergeRecursion(vSecond, aoFunc, aboFlg)
        
        '�}�[�W�����Ȃ����ʂɖ߂�
        this_sortMergeRecursion = this_sortMergeMerge(vFirst, vSecond, aoFunc, aboFlg)
        
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : this_sortMergeMerge()
    'Overview                    : �}�[�W�\�[�g�̃}�[�W����
    'Detailed Description        : this_sortMerge()����Ăяo��
    '                              �����̊֐��̈����͈ȉ��̂Ƃ���
    '                                currentValue :�z��̗v�f
    '                                nextValue    :���̔z��̗v�f
    'Argument
    '     avFirst                : �}�[�W����\�[�g�ς݂̔z��
    '     avSecond               : �}�[�W����\�[�g�ς݂̔z��
    '     aoFunc                 : �֐�
    '     aboFlg                 : ������@
    '                                True  :�����i�֐��̌��ʂ�True�̏ꍇ�ɓ���ւ���j
    '                                False :�~���i�֐��̌��ʂ�False�̏ꍇ�ɓ���ւ���j
    'Return Value
    '     �}�[�W�ς̔z��
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/18         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_sortMergeMerge( _
        byRef avFirst _
        , byRef avSecond _
        , byRef aoFunc _
        , byVal aboFlg _
        )
        Dim lPosF, lPosS, lEndF, lEndS
        lPosF = Lbound(avFirst) : lPosS = Lbound(avSecond)
        lEndF = Ubound(avFirst) : lEndS = Ubound(avSecond)
        
        '�o���̔z��̐擪�̗v�f���m���֐��Ŕ��肵�Ė߂�l�̔z��ɒǉ�����
        Dim vRet
        Do While lPosF<=lEndF And lPosS<=lEndS
            If this_comparison(aoFunc, avFirst(lPosF), avSecond(lPosS))=aboFlg Then
                cf_push vRet, avSecond(lPosS)
                lPosS = lPosS + 1
            Else
                cf_push vRet, avFirst(lPosF)
                lPosF = lPosF + 1
            End If
        Loop
        
        '���ꂼ��c���Ă�����̔z��̗v�f��ǉ�����
        Dim lPos
        If lPosF<=lEndF Then
            For lPos=lPosF To lEndF
                cf_push vRet, avFirst(lPos)
            Next
        End If
        If lPosS<=lEndS Then
            For lPos=lPosS To lEndS
                cf_push vRet, avSecond(lPos)
            Next
        End If
        
        '�}�[�W�ς̔z���Ԃ�
        this_sortMergeMerge = vRet
        
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : this_sortHeap()
    'Overview                    : �q�[�v�\�[�g
    'Detailed Description        : �v�Z�񐔂�O(N*logN)
    '                              �z��̗v�f���Ȃ��܂���1�̏ꍇ�͉������Ȃ�
    'Argument
    '     aoFunc                 : �֐�
    '     aboFlg                 : ������@
    '                                True  :�����i�֐��̌��ʂ�True�̏ꍇ�ɓ���ւ���j
    '                                False :�~���i�֐��̌��ʂ�False�̏ꍇ�ɓ���ւ���j
    'Return Value
    '     �\�[�g��̔z��
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/21         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub this_sortHeap( _
        byRef aoFunc _
        , byVal aboFlg _
        )
        If this_length()<2 Then Exit Sub
        Dim vArr : vArr = PvArr
        
        '�q�[�v�̍쐬
        Dim lLb, lUb, lSize, lParent
        lLb = Lbound(vArr) : lUb = Ubound(vArr)
        lSize = lUb - lLb + 1
        '�q�����ŉ����̃m�[�h�����ʂɌ����ď��ԂɃm�[�h�P�ʂ̏������s��
        For lParent=lSize\2-1 To lLb Step -1
            this_sortHeapPerNodeProc vArr, lSize, lParent, aoFunc, aboFlg
        Next
        
        '�q�[�v�̐擪�i�ő�/�ŏ��l�j�����ԂɎ��o��
        Do While lSize>0
            '�q�[�v�̐擪�Ɩ��������ւ���
            cf_swap vArr(lLb), vArr(lSize-1)
            '�q�[�v�T�C�Y���P���炵�čč쐬
            lSize = lSize - 1
            this_sortHeapPerNodeProc vArr, lSize, 0, aoFunc, aboFlg
        Loop

        PvArr = vArr
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : this_sortHeapPerNodeProc()
    'Overview                    : �q�[�v�\�[�g�̃m�[�h�P�ʂ̏���
    'Detailed Description        : this_sortHeap()����Ăяo��
    '                              �����̊֐��̈����͈ȉ��̂Ƃ���
    '                                currentValue :�z��̗v�f
    '                                nextValue    :���̔z��̗v�f
    'Argument
    '     avArr                  : �z��
    '     alSize                 : �q�[�v�̃T�C�Y
    '     alParent               : �m�[�h�̐e�̔z��ԍ�
    '     aoFunc                 : �֐�
    '     aboFlg                 : ������@
    '                                True  :�����i�֐��̌��ʂ�True�̏ꍇ�ɓ���ւ���j
    '                                False :�~���i�֐��̌��ʂ�False�̏ꍇ�ɓ���ւ���j
    'Return Value
    '     �\�[�g��̔z��
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/21         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub this_sortHeapPerNodeProc( _
        byRef avArr _
        , byVal alSize _
        , byVal alParent _
        , byRef aoFunc _
        , byVal aboFlg _
        )
        Dim lRight, lLeft, lToSwap
        lLeft = alParent*2 + 1
        lRight = lLeft + 1
        lToSwap = alParent
        
        If lRight<alSize Then
        '�E���̎q������ꍇ
            If this_comparison(aoFunc, avArr(lRight), avArr(alParent))=aboFlg Then
            '�e�ƉE���̎q�̗v�f���֐��Ŕ��肵������@�ɍ��v����ꍇ�͓���ւ���
                lToSwap = lRight
            End If
        End If
        
        If lLeft<alSize Then
        '�����̎q������ꍇ
            If this_comparison(aoFunc, avArr(lLeft), avArr(lToSwap))=aboFlg Then
            '�e�ƉE���̎q�̏��҂ƍ����̎q�̗v�f���֐��Ŕ��肵������@�ɍ��v����ꍇ�͓���ւ���
                lToSwap = lLeft
            End If
        End If
        
        If lToSwap<>alParent Then
            '�e�Ǝq�̗v�f�����ւ���
            cf_swap avArr(alParent), avArr(lToSwap)
            '����ւ����q�̗v�f�ȉ��̃m�[�h���ď�������
            this_sortHeapPerNodeProc avArr, alSize, lToSwap, aoFunc, aboFlg
        End If
        
    End Sub

    '***************************************************************************************************
    'Function/Sub Name           : this_splice()
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
    Private Function this_splice( _
        byVal alStart _
        , byVal alDelCnt _
        , byRef avArr _
        )
        Set this_splice = new_Arr()
        
        Dim lIdx, vArr, lUb, vArrayAft, lStart
        If this_length()>0 Then
            vArr = PvArr
            lUb = Ubound(vArr)
            
            If alStart<0 Then lStart=lUb+1 Else lStart=0
            lStart = math_max(lStart+alStart,0)
            lStart = math_min(lStart,lUb+1)
            
            For lIdx = 0 To lStart - 1
            '�J�n�ʒu�܂ł͍��̔z��̂܂�
                cf_push vArrayAft, vArr(lIdx)
            Next
            
            '�J�n�ʒu����폜����v�f�͕ʂ̔z��Ɉڂ�
            Dim vRet
            For lIdx = lStart To math_min(lStart+alDelCnt-1, lUb)
                cf_push vRet, vArr(lIdx)
            Next

            '�z�񂩂��菜�����v�f�ō쐬�����V�����C���X�^���X��Ԃ�
            If this_hasElement(vRet) Then Set this_splice = new_ArrWith(vRet)
        End If
        
        If this_hasElement(avArr) Then
        '�ǉ�����v�f������Βǉ�����
            For lIdx = 0 To Ubound(avArr)
                cf_push vArrayAft, avArr(lIdx)
            Next
        End If
        
        If this_length()>0 Then
            For lIdx = lStart+alDelCnt To lUb
            '�폜�����v�f�ȍ~�͍��̔z��Ɏc��
                cf_push vArrayAft, vArr(lIdx) 
            Next
        End If
        
        '�쐬�����z��ɒu������
        If this_hasElement(vArrayAft) Then PvArr = vArrayAft
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : this_uniq()
    'Overview                    : �z��̏d����r������
    'Detailed Description        : �H����
    'Argument
    '     �Ȃ�
    'Return Value
    '     ������̎��g�̃C���X�^���X
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/11/12         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_uniq( _
        )
        '�d����r��
        Dim oEle, oDic : Set oDic = new_Dic()
        For Each oEle In PvArr
            If Not oDic.Exists(oEle) Then oDic.Add oEle, Empty
        Next
        If oDic.Count<this_length() Then
        '�d�����������ꍇ�͐V�����z����쐬
            PvArr = oDic.Keys()
        End If
        '���g�̃C���X�^���X��Ԃ�
        Set this_uniq = Me

        Set oEle = Nothing
        Set oDic = Nothing
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : this_unshiftA()
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
    Private Function this_unshiftA( _
        byRef avArr _
        )
        Dim vArr, lUb, lUbAdd
        lUbAdd = 0
        If this_hasElement(avArr) Then
        '�����̗v�f��擪�ɒǉ�
            vArr = avArr
            lUbAdd = Ubound(avArr)
        End If

        '�����č�����v�f��ǉ�
        If this_length()>0 Then
            lUb = Ubound(PvArr)
            Redim Preserve vArr(lUbAdd + this_length())
            For lIdx=0 To lUb
                cf_bind vArr(lUbAdd+lIdx+1), PvArr(lIdx)
            Next
        End If

        '�쐬�����z��ɒu����
        PvArr = vArr
        this_unshiftA = this_length()

    End Function

    '***************************************************************************************************
    'Function/Sub Name           : this_toArray()
    'Overview                    : �����ŕێ�����z��̕������쐬����
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
    Private Function this_toArray( _
        aboOrder _
        )
        this_toArray = Array()
        Dim lLen : lLen = this_length()
        If lLen=0 Then Exit Function

        Dim vRet
        If aboOrder Then
            vRet=PvArr
        Else
            Redim vRet(lLen-1)
            Dim lIdx, lIdxR : lIdxR = 0
            For lIdx=Ubound(PvArr) To 0 Step -1
                cf_bind vRet(lIdxR), PvArr(lIdx)
                lIdxR = lIdxR + 1
            Next
        End If
        this_toArray = vRet
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : this_isValidIndex()
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
    Private Function this_isValidIndex( _
        byVal alIdx _
        )
        this_isValidIndex = False
        If this_length()>0 Then
            If 0<=alIdx And alIdx<=Ubound(PvArr) Then this_isValidIndex=True
        End If
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : this_length()
    'Overview                    : �z��̗v�f����Ԃ�
    'Detailed Description        : �H����
    'Argument
    '     �Ȃ�
    'Return Value
    '     �z��̗v�f��
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/12/21         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_length( _
        )
        this_length = 0
        If this_hasElement(PvArr) Then this_length = Ubound(PvArr)+1
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : this_setItem()
    'Overview                    : �z��̎w�肵���C���f�b�N�X�ɗv�f��ݒ肷��
    'Detailed Description        : �H����
    'Argument
    '     alIdx                  : �C���f�b�N�X
    '     aoEle                  : �ݒ肷��v�f
    '     asSource               : �\�[�X
    'Return Value
    '     �Ȃ�
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2024/09/29         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub this_setItem( _
        byVal alIdx _
        , byRef aoEle _
        , byVal asSource _
        )
        ast_argTrue this_isValidIndex(alIdx), asSource, "Index is out of range."
        cf_bind PvArr(alIdx), aoEle
    End Sub

End Class
