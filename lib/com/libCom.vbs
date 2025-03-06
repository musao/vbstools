'***************************************************************************************************
'FILENAME                    : libCom.vbs
'Overview                    : ���ʊ֐����C�u����
'Detailed Description        : �H����
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/09/27         Y.Fujii                  First edition
'***************************************************************************************************

'###################################################################################################
'�J�X�^���֐�
'###################################################################################################

'***************************************************************************************************
'Function/Sub Name           : cf_bind()
'Overview                    : �ϐ��Ԃ̍��ڈڑ�
'Detailed Description        : �ڑ�����l�܂��͕ϐ����I�u�W�F�N�g���ۂ��ɂ��VBS�\���̈Ⴂ�iSet�̗L���j���z������
'                              �ڑ��悪�R���N�V�����̃����o�[�̏ꍇ�͓��삵�Ȃ�
'                              �ڑ��悪�ϐ��̏ꍇ�Ɏg�p�ł���
'Argument
'     avTo                   : �ڑ���̕ϐ�
'     avValue                : �ڑ�����l�܂��͕ϐ�
'Return Value
'     �Ȃ�
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/11/06         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub cf_bind( _
    byRef avTo _
    , byRef avValue _
    )
    If IsObject(avValue) Then Set avTo = avValue Else avTo = avValue
End Sub

'***************************************************************************************************
'Function/Sub Name           : cf_bindAt()
'Overview                    : �ϐ��Ԃ̍��ڈڑ�
'Detailed Description        : �ڑ�����l�܂��͕ϐ����I�u�W�F�N�g���ۂ��ɂ��VBS�\���̈Ⴂ�iSet�̗L���j���z������
'                              �ڑ��悪�R���N�V�����̏ꍇ�͓��֐����g�p����
'Argument
'     aoCollection           : �ڑ���̃R���N�V����
'     asKey                  : �ڑ���̃R���N�V�����̃L�[
'     avValue                : �ڑ�����l�܂��͕ϐ�
'Return Value
'     �Ȃ�
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/11/06         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub cf_bindAt( _
    byRef aoCollection _
    , byVal asKey _
    , byRef avValue _
    )
    If IsObject(avValue) Then Set aoCollection.Item(asKey) = avValue Else aoCollection.Item(asKey) = avValue
End Sub

'***************************************************************************************************
'Function/Sub Name           : cf_isAvailableObject()
'Overview                    : �I�u�W�F�N�g�����p�\�����肷��
'Detailed Description        : �H����
'Argument
'     aoObj                  : �I�u�W�F�N�g
'Return Value
'     ���� True:���p�\ / False:���p�s��
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/10/15         Y.Fujii                  First edition
'***************************************************************************************************
Private Function cf_isAvailableObject( _
    byRef aoObj _
    )
    Dim boFlg : boFlg = False
    If IsObject(aoObj) Then
        If Not aoObj Is Nothing Then boFlg = True
    End If
    cf_isAvailableObject = boFlg
End Function

'***************************************************************************************************
'Function/Sub Name           : cf_isInteger()
'Overview                    : �������ǂ�����������
'Detailed Description        : �H����
'Argument
'     avValue                : �����Ώ�
'Return Value
'     ���� True:���� / False:�����łȂ�
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2024/09/29         Y.Fujii                  First edition
'***************************************************************************************************
Private Function cf_isInteger( _
    byRef avValue _
    )
    cf_isInteger = False
    If cf_isNumeric(avValue) Then cf_isInteger = (Fix(avValue) = cdbl(avValue))
End Function

'***************************************************************************************************
'Function/Sub Name           : cf_isNonNegativeNumber()
'Overview                    : ���̐��łȂ��i��0�܂��͐��̐��j���Ƃ���������
'Detailed Description        : �H����
'Argument
'     avValue                : �����Ώ�
'Return Value
'     ���� True:���̐��łȂ� / False:���̐�
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2024/09/30         Y.Fujii                  First edition
'***************************************************************************************************
Private Function cf_isNonNegativeNumber( _
    byRef avValue _
    )
    cf_isNonNegativeNumber = False
    If cf_isNumeric(avValue) Then cf_isNonNegativeNumber = Not (0 > avValue)
End Function

'***************************************************************************************************
'Function/Sub Name           : cf_isNumeric()
'Overview                    : ���l����������
'Detailed Description        : �H����
'Argument
'     avValue                : �����Ώ�
'Return Value
'     ���� True:���l / False:���l�łȂ�
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2024/01/03         Y.Fujii                  First edition
'***************************************************************************************************
Private Function cf_isNumeric( _
    byRef avValue _
    )
    If IsEmpty(avValue) Or IsNull(avValue) Or IsObject(avValue) Or IsArray(avValue) Then
    'Empty,Null,Object,Array�̏ꍇ��False
        cf_isNumeric=False
        Exit Function
    End If
    If VarType(avValue)=vbInteger Or VarType(avValue)=vbLong Or VarType(avValue)=vbSingle Or VarType(avValue)=vbDouble Then
    'Integer,Long,Single,Double�̏ꍇ��True
        cf_isNumeric=True
        Exit Function
    End If
    cf_isNumeric=False
    If VarType(avValue)=vbString Then
    'String�̏ꍇ��IsNumeric�֐��̖߂�l��Ԃ�
        cf_isNumeric=IsNumeric(avValue)
    End If
End Function

'***************************************************************************************************
'Function/Sub Name           : cf_isPositiveInteger()
'Overview                    : ���̐������ǂ�����������
'Detailed Description        : �H����
'Argument
'     avValue                : �����Ώ�
'Return Value
'     ���� True:���̐��� / False:���̐����łȂ�
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2024/09/29         Y.Fujii                  First edition
'***************************************************************************************************
Private Function cf_isPositiveInteger( _
    byRef avValue _
    )
    cf_isPositiveInteger = False
    If cf_isInteger(avValue) Then cf_isPositiveInteger = (cdbl(avValue) > 0)
End Function

'***************************************************************************************************
'Function/Sub Name           : cf_isSame()
'Overview                    : ���ꂩ���肷��
'Detailed Description        : �H����
'Argument
'     avA                    : ��r�Ώ�
'     avB                    : ��r�Ώ�
'Return Value
'     ���� True:���� / False:����łȂ�
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/10/15         Y.Fujii                  First edition
'***************************************************************************************************
Private Function cf_isSame( _
    byRef avA _
    , byRef avB _
    )
    Dim boFlg : boFlg = False
    If IsObject(avA) And IsObject(avB) Then
        If avA Is avB Then boFlg = True
    ElseIf Not IsObject(avA) And Not IsObject(avB) Then
        If VarType(avA) = vbString And VarType(avB) = vbString Then
            If Strcomp(avA, avB, vbBinaryCompare)=0 Then boFlg = True
        Else
            If avA = avB Then boFlg = True
        End If
    End If
    cf_isSame = boFlg
End Function

'***************************************************************************************************
'Function/Sub Name           : cf_isValid()
'Overview                    : �L���Ȓl�i�����l�łȂ��j�����肷��
'Detailed Description        : �H����
'Argument
'     avTgt                  : ����Ώ�
'Return Value
'     ���� True:�L���Ȓl������ / False:�L���Ȓl���Ȃ�
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/12/27         Y.Fujii                  First edition
'***************************************************************************************************
Private Function cf_isValid( _
    byRef avTgt _
    )
    Dim boFlg : boFlg = True
    If IsObject(avTgt) Then
    '�I�u�W�F�N�g�̏ꍇ
        If avTgt Is Nothing Then boFlg = False
    ElseIf IsArray(avTgt) Then
    '�z��̏ꍇ
        boFlg = new_Arr().hasElement(avTgt)
    Else
    '��L�ȊO�̏ꍇ
        If IsEmpty(avTgt) Or IsNull(avTgt) Then
            boFlg = False
        ElseIf cf_isSame(avTgt, vbNullString) Then
            boFlg = False
        End If
    End If
    cf_isValid = boFlg
End Function

'***************************************************************************************************
'Function/Sub Name           : cf_push()
'Overview                    : �z��ɗv�f��ǉ�����
'Detailed Description        : sub_CfPush()�ɈϏ�����
'Argument
'     avArr                  : �z��
'     avEle                  : �ǉ�����v�f
'Return Value
'     �Ȃ�
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/11/23         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub cf_push( _
    byRef avArr _ 
    , byRef avEle _ 
    )
    sub_CfPush avArr, avEle
End Sub

'***************************************************************************************************
'Function/Sub Name           : cf_pushA()
'Overview                    : �����̒ǉ�����z��̗v�f��z��ɒǉ�����
'Detailed Description        : sub_CfPushA()�ɈϏ�����
'Argument
'     avArr                  : �z��
'     avAdd                  : �ǉ�����v�f�̔z��
'Return Value
'     �Ȃ�
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/11/03         Y.Fujii                  First edition
'2024/05/01         Y.Fujii                  Rename cf_pushMulti() -> cf_pushA()
'***************************************************************************************************
Private Sub cf_pushA( _
    byRef avArr _ 
    , byRef avAdd _ 
    )
    sub_CfPushA avArr, avAdd
End Sub

'***************************************************************************************************
'Function/Sub Name           : cf_swap()
'Overview                    : �ϐ��̒l�����ւ���
'Detailed Description        : �ڑ�������cf_bind()���g�p����
'Argument
'     avA                    : �l�����ւ���ϐ�
'     avB                    : �l�����ւ���ϐ�
'Return Value
'     �Ȃ�
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/21         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub cf_swap( _
    byRef avA _
    , byRef avB _
    )
    Dim oTmp
    cf_bind oTmp, avA
    cf_bind avA, avB
    cf_bind avB, oTmp
    Set oTmp = Nothing
End Sub

'***************************************************************************************************
'Function/Sub Name           : cf_toString()
'Overview                    : �����̓��e�𕶎���ŕ\������
'Detailed Description        : func_CfToString()�ɈϏ�����
'Argument
'     avTgt                  : �Ώ�
'Return Value
'     ������ɕϊ����������̓��e
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/12/24         Y.Fujii                  First edition
'***************************************************************************************************
Private Function cf_toString( _
    byRef avTgt _
    )
    cf_toString = func_CfToString(avTgt)
End Function


'***************************************************************************************************
'Function/Sub Name           : func_CfToString()
'Overview                    : �����̓��e�𕶎���ŕ\������
'Detailed Description        : �\���^���͈ȉ��̂Ƃ���
'                               �z��ADictionary�͗v�f���Ƃɓ��e��\������A����q�͍ċA�\������
'                               �@�z��F[<Long>0,<String>"a",<Empty>,[value1,...],{key1=>value1,...},...]
'                               �@Dictionary�F{key1=>value1,key2=>[a_value1,...],key3=>{d_key1=>d_value1,...}...}
'                               ��L�ȊO <VarType>Value�`�� ��Value�͂Ȃ��ꍇ����
'Argument
'     avTgt                  : �Ώ�
'Return Value
'     ������ɕϊ����������̓��e
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/12/24         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CfToString( _
    byRef avTgt _
    )
    If IsArray(avTgt) Then
        func_CfToString = func_CfToStringArray(avTgt)
        Exit Function
    End If
    If IsObject(avTgt) Then
        func_CfToString = func_CfToStringObject(avTgt)
        Exit Function
    End If
    Dim sRet : sRet = "<" & TypeName(avTgt) & ">" 
    If cf_isSame(TypeName(avTgt),"String") Then
        sRet = sRet & Chr(34) & Replace(avTgt,Chr(34),Chr(34)&Chr(34)) & Chr(34)
    ElseIf Not (IsEmpty(avTgt) Or IsNull(avTgt)) Then
        sRet = sRet & CStr(avTgt)
    End If
    func_CfToString = sRet
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CfToStringArray()
'Overview                    : �z��̓��e�𕶎���ŕ\������
'Detailed Description        : �H����
'Argument
'     avTgt                  : �Ώ�
'Return Value
'     ������ɕϊ����������̓��e
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/12/24         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CfToStringArray( _
    byRef avTgt _
    )
    If new_Arr().hasElement(avTgt) Then
        Dim vRet, oEle
        For Each oEle In avTgt
            cf_push vRet, func_CfToString(oEle)
        Next
        func_CfToStringArray = "<Array>[" & Join(vRet, ",") & "]"
        Set oEle = Nothing
    Else
        func_CfToStringArray = "<Array>[]"
    End If
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CfToStringObject()
'Overview                    : �I�u�W�F�N�g�̓��e�𕶎���ŕ\������
'Detailed Description        : �H����
'Argument
'     avTgt                  : �Ώ�
'Return Value
'     ������ɕϊ����������̓��e
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/12/24         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CfToStringObject( _
    byRef avTgt _
    )
    If cf_isSame(TypeName(avTgt),"Dictionary") Then
        func_CfToStringObject = func_CfToStringObjectDictionary(avTgt)
        Exit Function
    End If

    On Error Resume Next
    func_CfToStringObject = avTgt.toString()
    If Err.Number=0 Then Exit Function
    On Error Goto 0

    If cf_isSame(VarType(avTgt), vbString) Then
        func_CfToStringObject = "<" & TypeName(avTgt) & ">" & avTgt
        Exit Function
    End If
    func_CfToStringObject = "<" & TypeName(avTgt) & ">"
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CfToStringObjectDictionary()
'Overview                    : �f�B�N�V���i���̓��e�𕶎���ŕ\������
'Detailed Description        : �H����
'Argument
'     avTgt                  : �Ώ�
'Return Value
'     ������ɕϊ����������̓��e
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/03         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CfToStringObjectDictionary( _
    byRef avTgt _
    )
    Const Cs_SPKEY = "__Special__"
    Dim sLabel : sLabel="Dictionary"
    If avTgt.Count>0 Then
        If avTgt.Exists(Cs_SPKEY) Then sLabel=avTgt.Item(Cs_SPKEY)
        Dim vRet, oEle
        For Each oEle In avTgt.Keys
            If Not cf_isSame(oEle,Cs_SPKEY) Then
                cf_push vRet, func_CfToString(oEle) & "=>" & func_CfToString(avTgt.Item(oEle))
            End If
        Next
        func_CfToStringObjectDictionary = "<" & sLabel & ">{" & Join(vRet, ",") & "}"
        Set oEle = Nothing
    Else
        func_CfToStringObjectDictionary = "<" & sLabel & ">{}"
    End If
End Function

'***************************************************************************************************
'Function/Sub Name           : sub_CfPush()
'Overview                    : �z��ɗv�f��ǉ�����
'Detailed Description        : �H����
'Argument
'     avArr                  : �z��
'     avEle                  : �ǉ�����v�f
'Return Value
'     �Ȃ�
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2024/05/01         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub sub_CfPush( _
    byRef avArr _ 
    , byRef avEle _ 
    )
    On Error Resume Next
    Redim Preserve avArr(Ubound(avArr)+1)
    If Err.Number<>0 Then Redim avArr(0)
    On Error Goto 0
    cf_bind avArr(Ubound(avArr)), avEle
End Sub

'***************************************************************************************************
'Function/Sub Name           : sub_CfPushA()
'Overview                    : �����̒ǉ�����z��̗v�f��z��ɒǉ�����
'Detailed Description        : �ǉ�����z��iavAdd�j���z��łȂ��ꍇ��sub_CfPush()�����s����
'Argument
'     avArr                  : �z��
'     avAdd                  : �ǉ�����z��
'Return Value
'     �Ȃ�
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2024/05/01         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub sub_CfPushA( _
    byRef avArr _ 
    , byRef avAdd _ 
    )
    On Error Resume Next
    Dim lUbAdd,lIdx : lUbAdd = Ubound(avAdd)
    If Err.Number=0 Then
    '�ǉ�����z��iavAdd�j���v�f�����ꍇ
        '�z��iavArr�j���g������
        Dim lUb : lUb = Ubound(avArr)
        If Err.Number=0 Then
            Redim Preserve avArr(lUb+lUbAdd+1)
        Else
        '�z��iavArr�j���v�f�������Ȃ��ꍇ��lUb��-1�ɂ���
            lUb = -1
            Redim avArr(lUbAdd)
        End If

        '�z��iavArr�j�ɒǉ�����v�f�̔z��iavAdd�j��ǉ�����
        For lIdx=0 To lUbAdd
            cf_bind avArr(lUb+1+lIdx), avAdd(lIdx)
        Next
    Elseif Not IsArray(avAdd) Then
    '�ǉ�����z��iavAdd�j���v�f���������z��łȂ��ꍇ
        sub_CfPush avArr, avAdd
    End If
    On Error Goto 0
End Sub

'###################################################################################################
'�����n�̊֐�
'###################################################################################################

'***************************************************************************************************
'Function/Sub Name           : ast_argFalse()
'Overview                    : ������False����������
'Detailed Description        : �������ʂ�NG�̏ꍇ�͗�O���o��
'Argument
'     avArgument             : �Ώ�
'     asSource               : �\�[�X
'     asDescription          : ����
'Return Value
'     �Ȃ�
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2024/05/28         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub ast_argFalse( _
    byRef avArgument _
    , byVal asSource _
    , byVal asDescription _
    )
    If Not cf_isSame(False, avArgument) Then fw_throwException 8193, asSource, asDescription
End Sub

'***************************************************************************************************
'Function/Sub Name           : ast_argNotEmpty()
'Overview                    : ������Empty�łȂ�����������
'Detailed Description        : �������ʂ�NG�̏ꍇ�͗�O���o��
'Argument
'     avArgument             : �Ώ�
'     asSource               : �\�[�X
'     asDescription          : ����
'Return Value
'     �Ȃ�
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2024/06/02         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub ast_argNotEmpty( _
    byRef avArgument _
    , byVal asSource _
    , byVal asDescription _
    )
    If IsEmpty(avArgument) Then fw_throwException 8194, asSource, asDescription
End Sub

'***************************************************************************************************
'Function/Sub Name           : ast_argNotNull()
'Overview                    : ������Null�łȂ�����������
'Detailed Description        : �������ʂ�NG�̏ꍇ�͗�O���o��
'Argument
'     avArgument             : �Ώ�
'     asSource               : �\�[�X
'     asDescription          : ����
'Return Value
'     �Ȃ�
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2024/06/03         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub ast_argNotNull( _
    byRef avArgument _
    , byVal asSource _
    , byVal asDescription _
    )
    If IsNull(avArgument) Then fw_throwException 8195, asSource, asDescription
End Sub

'***************************************************************************************************
'Function/Sub Name           : ast_argTrue()
'Overview                    : ������True����������
'Detailed Description        : �������ʂ�NG�̏ꍇ�͗�O���o��
'Argument
'     avArgument             : �Ώ�
'     asSource               : �\�[�X
'     asDescription          : ����
'Return Value
'     �Ȃ�
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2024/06/16         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub ast_argTrue( _
    byRef avArgument _
    , byVal asSource _
    , byVal asDescription _
    )
    If Not cf_isSame(True, avArgument) Then fw_throwException 8196, asSource, asDescription
End Sub

'***************************************************************************************************
'Function/Sub Name           : ast_argsAreSame()
'Overview                    : ���������ꂩ��������
'Detailed Description        : �������ʂ�NG�̏ꍇ�͗�O���o��
'Argument
'     avA                    : ��r�Ώ�
'     avB                    : ��r�Ώ�
'     asSource               : �\�[�X
'     asDescription          : ����
'Return Value
'     �Ȃ�
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2025/02/02         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub ast_argsAreSame( _
    byRef avA _
    , byRef avB _
    , byVal asSource _
    , byVal asDescription _
    )
    If Not cf_isSame(avA, avB) Then fw_throwException 8197, asSource, asDescription
End Sub

'***************************************************************************************************
'Function/Sub Name           : ast_argNull()
'Overview                    : ������Null����������
'Detailed Description        : �������ʂ�NG�̏ꍇ�͗�O���o��
'Argument
'     avArgument             : �Ώ�
'     asSource               : �\�[�X
'     asDescription          : ����
'Return Value
'     �Ȃ�
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2025/02/02         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub ast_argNull( _
    byRef avArgument _
    , byVal asSource _
    , byVal asDescription _
    )
    If Not(IsNull(avArgument)) Then fw_throwException 8198, asSource, asDescription
End Sub

'***************************************************************************************************
'Function/Sub Name           : ast_failure()
'Overview                    : ��O���o��
'Detailed Description        : ����
'Argument
'     asSource               : �\�[�X
'     asDescription          : ����
'Return Value
'     �Ȃ�
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2025/02/02         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub ast_failure( _
    byVal asSource _
    , byVal asDescription _
    )
    fw_throwException 8199, asSource, asDescription
End Sub

'###################################################################################################
'�t���[�����[�N�n�̊֐�
'###################################################################################################

'***************************************************************************************************
'Function/Sub Name           : fw_excuteSub()
'Overview                    : �֐������s����
'Detailed Description        : �u���[�J�[�̎w�肪����Ύ��s�O��ɏo�ŁiPublish�j�������s��
'Argument
'     asSubName              : ���s����֐���
'     aoArg                  : ���s����֐��ɓn������
'     aoBroker               : �u���[�J�[�N���X�̃I�u�W�F�N�g
'Return Value
'     �Ȃ�
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/03         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub fw_excuteSub( _
    byVal asSubName _
    , byRef aoArg _
    , byRef aoBroker _
    )
    Dim sSubNameForPublish : sSubNameForPublish=asSubName&"()"

    '���s�O�̏o�ŁiPublish�j ����
    If cf_isAvailableObject(aoBroker) Then
        aoBroker.publish topic.LOG, Array(logType.INFO ,sSubNameForPublish ,"Start")
        aoBroker.publish topic.LOG, Array(logType.DETAIL ,sSubNameForPublish ,cf_toString(aoArg))
    End If
    
    '�֐��̎��s
    Dim oRet : Set oRet = fw_tryCatch(GetRef(asSubName), aoArg, Empty, Empty)
    
    '���s��̏o�ŁiPublish�j ����
    If cf_isAvailableObject(aoBroker) Then
        If oRet.isErr() Then
        '�G���[
            aoBroker.publish topic.LOG, Array(logType.ERROR, sSubNameForPublish, cf_toString(oRet.getErr()))
        End If
        aoBroker.publish topic.LOG, Array(logType.INFO, sSubNameForPublish, "End")
        aoBroker.publish topic.LOG, Array(logType.DETAIL, sSubNameForPublish, cf_toString(aoArg))
    End If
    
    Set oRet = Nothing
End Sub

'***************************************************************************************************
'Function/Sub Name           : fw_getLogPath()
'Overview                    : ���s���̃X�N���v�g�̃��O�t�@�C���p�X��Ԃ�
'Detailed Description        : ���s���̃X�N���v�g������t�H���_��log�t�H���_�ȉ���
'                              �X�N���v�g�t�@�C�����{".log"�`���̃t�@�C�����ō쐬����
'                              fw_getPrivatePath()�ɈϏ�����
'Argument
'     �Ȃ�
'Return Value
'     �t�@�C���̃t���p�X
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/26         Y.Fujii                  First edition
'***************************************************************************************************
Private Function fw_getLogPath( _
    )
    fw_getLogPath = fw_getPrivatePath("log", new_Fso().GetBaseName(WScript.ScriptName) & ".log" )
End Function

'***************************************************************************************************
'Function/Sub Name           : fw_getTextstreamForLog()
'Overview                    : log�o�͗p�̃e�L�X�g�X�g���[�����쐬����
'Detailed Description        : log�o�͐��fw_getLogPath()�Ŏ擾����
'Argument
'     �Ȃ�
'Return Value
'     �t�@�C���o�̓o�b�t�@�����O�����N���X�̃C���X�^���X
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2025/03/02         Y.Fujii                  First edition
'***************************************************************************************************
Private Function fw_getTextstreamForLog( _
    )
    Set fw_getTextstreamForLog = new_WriterOf(fw_getLogPath, 8, True, -1)
End Function

'***************************************************************************************************
'Function/Sub Name           : fw_getPrivatePath()
'Overview                    : ���s���̃X�N���v�g������t�H���_�ȉ��̃t���p�X��Ԃ�
'Detailed Description        : �e�t�H���_���̎w�肪����΂��̃t�H���_�ȉ��̃p�X��Ԃ�
'                              �e�t�H���_���̎w�肪�Ȃ��ꍇ�͎��s���̃X�N���v�g������t�H���_�����̃p�X��Ԃ�
'Argument
'     asParentFolderName     : �e�t�H���_��
'     asFileName             : �t�@�C����
'Return Value
'     �t���p�X
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/26         Y.Fujii                  First edition
'***************************************************************************************************
Private Function fw_getPrivatePath( _
    byVal asParentFolderName _
    , byVal asFileName _
    )
    '���s���̃X�N���v�g������t�H���_�̃p�X���擾
    Dim sParentFolderPath : sParentFolderPath = new_Fso().GetParentFolderName(WScript.ScriptFullName)
    
    '�t�@�C���̏�ʃf�B���N�g�������߂�
    If Len(asParentFolderName)>0 Then
    '�����Ŏw�肵���f�B���N�g����������ꍇ
        sParentFolderPath = new_Fso().BuildPath(sParentFolderPath, asParentFolderName)
    End If

    '��ʃf�B���N�g�������݂��Ȃ��ꍇ�͍쐬����
    fs_createFolder(sParentFolderPath)
    
    '�p�X��Ԃ�
    fw_getPrivatePath = new_Fso().BuildPath(sParentFolderPath, asFileName)
End Function

'***************************************************************************************************
'Function/Sub Name           : fw_getTempPath()
'Overview                    : �ꎞ�t�@�C���̃p�X��Ԃ�
'Detailed Description        : ���s���̃X�N���v�g������t�H���_��tmp�t�H���_�ȉ��ɍ쐬����
'                              fw_getPrivatePath()�ɈϏ�����
'Argument
'     �Ȃ�
'Return Value
'     �t�@�C���̃t���p�X
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/26         Y.Fujii                  First edition
'***************************************************************************************************
Private Function fw_getTempPath( _
    )
    fw_getTempPath = fw_getPrivatePath("tmp", new_Fso().GetTempName())
End Function

'***************************************************************************************************
'Function/Sub Name           : fw_logger()
'Overview                    : ���O�o�͂���
'Detailed Description        : �����̏��Ƀ^�C���X�^���v��t�����ăt�@�C���o�͂���
'Argument
'     avParams               : �z��^�̃p�����[�^���X�g
'     aoWriter               : �t�@�C���o�̓o�b�t�@�����O�����N���X�̃C���X�^���X
'Return Value
'     �Ȃ�
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/26         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub fw_logger( _
    byRef avParams _
    , byRef aoWriter _
    )
    Dim vIps, oEle
    For Each oEle In util_getIpAddress()
        cf_push vIps, oEle.Item("Ip").Item("V4")
    Next

    Dim oType : cf_bind oType, avParams(0)
    If cf_isSame("ReadOnlyObject", TypeName(oType)) Then avParams(0) = oType.name

    With aoWriter
        .WriteLine(new_ArrOf(Array(new_Now(), Join(vIps,","), new_Network().ComputerName)).Concat(avParams).join(vbTab))
    End With

    Set oType = Nothing
    Set oEle = Nothing
End Sub

'***************************************************************************************************
'Function/Sub Name           : fw_runShellSilently()
'Overview                    : �V�F�����T�C�����g���s����
'Detailed Description        : ���������A�V�F���̎��s������ɐ����߂�
'Argument
'     asCmd                  : ���s����R�}���h
'Return Value
'     ���� True:���� / False:���s
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2024/02/04         Y.Fujii                  First edition
'***************************************************************************************************
Private Function fw_runShellSilently( _
    byVal asCmd _
    )
    fw_runShellSilently = False
    If 0 = new_Shell().Run(asCmd, 0, True) Then fw_runShellSilently = True
End Function

'***************************************************************************************************
'Function/Sub Name           : fw_storeErr()
'Overview                    : Err�I�u�W�F�N�g�̓��e���I�u�W�F�N�g�ɕϊ�����
'Detailed Description        : �ϊ������I�u�W�F�N�g�̍\��
'                              Key             Value                     ��
'                              --------------  ------------------------  ---------------------------
'                              "Number"        Err.Number�̓��e          11
'                              "Description"   Err.Description�̂̓��e   0 �ŏ��Z���܂����B
'                              "Source"        Err.Source�̓��e          Microsoft VBScript ���s���G���[
'Argument
'     �Ȃ�
'Return Value
'     �ϊ������I�u�W�F�N�g
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/10/01         Y.Fujii                  First edition
'***************************************************************************************************
Private Function fw_storeErr( _
    )
    Dim oRet : Set oRet = new_Dic()
    '����L�[��ǉ�
    oRet.Add "__Special__", "Err"

    oRet.Add "Number", Err.Number
    oRet.Add "Description", Err.Description
    oRet.Add "Source", Err.Source
    Set fw_storeErr = oRet
    Set oRet = Nothing
End Function

'***************************************************************************************************
'Function/Sub Name           : fw_storeArguments()
'Overview                    : Arguments�I�u�W�F�N�g�̓��e���I�u�W�F�N�g�ɕϊ�����
'Detailed Description        : �ϊ������I�u�W�F�N�g�̍\��
'                              ��͈����� a /X /Hoge:Fuga, b �̏ꍇ
'                              Key         Value                                        ��
'                              ----------  -------------------------------------------  -------------
'                              "All"       WScript.Arguments�ȉ���Item�̓��e            a /X /Hoge:Fuga, b
'                              "Named"     WScript.Arguments.Named�ȉ���Item�̓��e      X: Hoge:Fuga
'                              "Unnamed"   WScript.Arguments.Unnamed�ȉ���Item�̓��e    a b
'Argument
'     �Ȃ�
'Return Value
'     �ϊ������I�u�W�F�N�g
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/26         Y.Fujii                  First edition
'***************************************************************************************************
Private Function fw_storeArguments( _
    )
    Dim oRet : Set oRet = new_Dic()
    '����L�[��ǉ�
    oRet.Add "__Special__", "Arguments"
    
    Dim vArr, oDic, oEle, oKey
    'All
    vArr = Array()
    For Each oEle In WScript.Arguments
        cf_push vArr, oEle
    Next
    oRet.Add "All", vArr
    
    'Named
    Set oDic = new_Dic()
    For Each oKey In WScript.Arguments.Named
        oDic.Add oKey, WScript.Arguments.Named.Item(oKey)
    Next
    oRet.Add "Named", oDic
    
    'Unnamed
    vArr = Array()
    For Each oEle In WScript.Arguments.Unnamed
        cf_push vArr, oEle
    Next
    oRet.Add "Unnamed", vArr
    
    Set fw_storeArguments = oRet
    
    Set oKey = Nothing
    Set oEle = Nothing
    Set oDic = Nothing
    Set oRet = Nothing
End Function

'***************************************************************************************************
'Function/Sub Name           : fw_throwException()
'Overview                    : ��O�𓊂���
'Detailed Description        : �������ʂ�NG�̏ꍇ�͗�O���o��
'Argument
'     alNumber               : �G���[�ԍ�
'     asSource               : �\�[�X
'     asDescription          : ����
'Return Value
'     �Ȃ�
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2024/05/28         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub fw_throwException( _
    byVal alNumber _
    , byVal asSource _
    , byVal asDescription _
    )
    Err.Raise alNumber, asSource, asDescription
End Sub

'***************************************************************************************************
'Function/Sub Name           : fw_tryCatch()
'Overview                    : �����̎��s�ƃG���[�������̏������s
'Detailed Description        : ���̌����try-chatch���ɏ���
'Argument
'     aoTry                  : ���s���鏈���itry�u���b�N�̏����j
'     aoArgs                 : ���s���鏈���̈���
'     aoCatch                : �G���[�������̏����icatch�u���b�N�̏����j
'     aoFinary               : �G���[�̗L���Ɉ˂炸�Ō�Ɏ��s���鏈���ifinary�u���b�N�̏����j
'Return Value
'     ��������
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/10/01         Y.Fujii                  First edition
'***************************************************************************************************
Private Function fw_tryCatch( _
    byRef aoTry _
    , byRef aoArgs _
    , byRef aoCatch _
    , byRef aoFinary _
    )
    Dim oRet, oRetF, oErr
    
    'try�u���b�N�̏���
    On Error Resume Next
    If cf_isValid(aoArgs) Then
        cf_bind oRetF, aoTry(aoArgs)
    Else
        cf_bind oRetF, aoTry()
    End If
    Set oRet = new_Ret(oRetF)
    On Error GoTo 0

    'catch�u���b�N�̏���
    If oRet.isErr() And cf_isAvailableObject(aoCatch) Then
        If cf_isValid(aoArgs) Then
            cf_bind oRetF, aoCatch(aoArgs)
        Else
            cf_bind oRetF, aoCatch()
        End If
        if IsObject(oRetF) Then Set oRet.returnValue=oRetF Else oRet.returnValue=oRetF
    End If
    
    'finary�u���b�N�̏���
    If cf_isAvailableObject(aoFinary) Then
        cf_bind oRetF, aoFinary(oRetF)
        if IsObject(oRetF) Then Set oRet.returnValue=oRetF Else oRet.returnValue=oRetF
    End If
    
    '���ʂ�ԋp
    Set fw_tryCatch = oRet
    Set oRet = Nothing
    Set oRetF = Nothing
End Function


'###################################################################################################
'�C���X�^���X�����֐�
'###################################################################################################

'***************************************************************************************************
'Function/Sub Name           : new_Adodb()
'Overview                    : ADO�X�g���[���I�u�W�F�N�g�̐����֐�
'Detailed Description        : �H����
'Argument
'     �Ȃ�
'Return Value
'     ADO�X�g���[���I�u�W�F�N�g�̃C���X�^���X
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2024/01/29         Y.Fujii                  First edition
'***************************************************************************************************
Private Function new_Adodb( _
    )
    Set new_Adodb = CreateObject("ADODB.Stream")
End Function

'***************************************************************************************************
'Function/Sub Name           : new_AdptFile()
'Overview                    : File�I�u�W�F�N�g�̃A�_�v�^�[�����֐�
'Detailed Description        : �H����
'Argument
'     aoFile                 : �t�@�C���̃I�u�W�F�N�g
'Return Value
'     ��������File�I�u�W�F�N�g�̃A�_�v�^�[�̃C���X�^���X
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/11/26         Y.Fujii                  First edition
'***************************************************************************************************
Private Function new_AdptFile( _
    byRef aoFile _
    )
    Set new_AdptFile = (New clsAdptFile).setFileObject(aoFile)
End Function

'***************************************************************************************************
'Function/Sub Name           : new_AdptFileOf()
'Overview                    : File�I�u�W�F�N�g�̃A�_�v�^�[�����֐�
'Detailed Description        : �H����
'Argument
'     asPath                 : �t�@�C���̃p�X
'Return Value
'     ��������File�I�u�W�F�N�g�̃A�_�v�^�[�̃C���X�^���X
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/11/26         Y.Fujii                  First edition
'***************************************************************************************************
Private Function new_AdptFileOf( _
    byVal asPath _
    )
    Set new_AdptFileOf = (New clsAdptFile).setFilePath(asPath)
End Function

'***************************************************************************************************
'Function/Sub Name           : new_Arr()
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
Private Function new_Arr( _
    )
    Set new_Arr = (New ArrayList)
End Function

'***************************************************************************************************
'Function/Sub Name           : new_ArrSplit()
'Overview                    : �C���X�^���X�����֐�
'Detailed Description        : vbscript��Split�֐��Ɠ����̋@�\�A���N���X�̃C���X�^���X��Ԃ�
'Argument
'     asTarget               : ����������Ƌ�؂蕶�����܂ޕ�����\��
'     asDelimiter            : ��؂蕶��
'Return Value
'     ���N���X�̃C���X�^���X
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/08         Y.Fujii                  First edition
'***************************************************************************************************
Private Function new_ArrSplit( _
    byVal asTarget _
    , byVal asDelimiter _
    )
    Set new_ArrSplit = new_ArrOf(Split(asTarget, asDelimiter, -1, vbBinaryCompare))
End Function

'***************************************************************************************************
'Function/Sub Name           : new_ArrOf()
'Overview                    : �C���X�^���X�����֐�
'Detailed Description        : �����Ŏw�肵���v�f���܂񂾓��N���X�̃C���X�^���X��Ԃ�
'Argument
'     avArr                  : �z��ɒǉ�����v�f�i�z��j
'Return Value
'     ���N���X�̃C���X�^���X
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/08         Y.Fujii                  First edition
'***************************************************************************************************
Private Function new_ArrOf( _
    byRef avArr _
    )
    Dim oArr : Set oArr = new_Arr()
    oArr.pushA avArr
    Set new_ArrOf = oArr
    Set oArr = Nothing
End Function

'***************************************************************************************************
'Function/Sub Name           : new_Broker()
'Overview                    : �C���X�^���X�����֐�
'Detailed Description        : �o��-�w�ǌ^�iPublish/Subscribe�j�N���X�̃C���X�^���X��Ԃ�
'Argument
'     �Ȃ�
'Return Value
'     �o��-�w�ǌ^�iPublish/Subscribe�j�N���X�̃C���X�^���X
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/03         Y.Fujii                  First edition
'***************************************************************************************************
Private Function new_Broker( _
    )
    Set new_Broker = (New Broker)
End Function

'***************************************************************************************************
'Function/Sub Name           : new_BrokerOf()
'Overview                    : �C���X�^���X�����֐�
'Detailed Description        : �o��-�w�ǌ^�iPublish/Subscribe�j�N���X�Ɏw�肵��topic��subscribe���ĕԂ�
'Argument
'     avParams               : ��i1,3,5,...�j�Ԗڂ�topic�A�����i2,4,6,...�j�Ԗڂ̓R�[���o�b�N�֐��|�C���^
'                              topic�����̏ꍇ�̓R�[���o�b�N�֐��|�C���^��subscribe���Ȃ�
'Return Value
'     �o��-�w�ǌ^�iPublish/Subscribe�j�N���X�̃C���X�^���X
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2025/03/06         Y.Fujii                  First edition
'***************************************************************************************************
Private Function new_BrokerOf( _
    byVal avParams _
    )
    Dim i,vTmp,lCnt,oBroker
    Set oBroker = New Broker
    lCnt = 0
    For Each i In avParams
        lCnt = lCnt + 1
        cf_push vTmp, i
        If lCnt Mod 2 = 0 Then oBroker.subscribe vTmp(lCnt-2), vTmp(lCnt-1)
    Next
    Set new_BrokerOf = oBroker
End Function

'***************************************************************************************************
'Function/Sub Name           : new_CalAt()
'Overview                    : �C���X�^���X�����֐�
'Detailed Description        : �w�肵�����t�����Ő����������t�N���X�̃C���X�^���X��Ԃ�
'Argument
'     avDateTime             : �ݒ肷����t����
'Return Value
'     ���t�N���X�̃C���X�^���X
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/03         Y.Fujii                  First edition
'***************************************************************************************************
Private Function new_CalAt( _
    ByVal avDateTime _
    )
    Set new_CalAt = (New Calendar).of(avDateTime)
End Function

'***************************************************************************************************
'Function/Sub Name           : new_Char()
'Overview                    : �C���X�^���X�����֐�
'Detailed Description        : ������ފǗ��N���X�̃C���X�^���X��Ԃ�
'Argument
'     �Ȃ�
'Return Value
'     ������ފǗ��N���X�̃C���X�^���X
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/10/31         Y.Fujii                  First edition
'***************************************************************************************************
Private Function new_Char( _
    )
    Set new_Char = (New CharacterType)
End Function

'***************************************************************************************************
'Function/Sub Name           : new_CssOf()
'Overview                    : �C���X�^���X�����֐�
'Detailed Description        : CSS�����N���X�̃C���X�^���X��Ԃ�
'Argument
'     asSelector             : �Z���N�^
'Return Value
'     CSS�����N���X�̃C���X�^���X
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/11/03         Y.Fujii                  First edition
'***************************************************************************************************
Private Function new_CssOf( _
    byVal asSelector _
    )
    Dim oCss : Set oCss = New CssGenerator
    oCss.selector = asSelector
    Set new_CssOf = oCss
    Set oCss = Nothing
End Function

'***************************************************************************************************
'Function/Sub Name           : new_Dic()
'Overview                    : Dictionary�I�u�W�F�N�g�����֐�
'Detailed Description        : �H����
'Argument
'     �Ȃ�
'Return Value
'     ��������Dictionary�I�u�W�F�N�g�̃C���X�^���X
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/03         Y.Fujii                  First edition
'***************************************************************************************************
Private Function new_Dic( _
    )
    Set new_Dic = CreateObject("Scripting.Dictionary")
End Function

'***************************************************************************************************
'Function/Sub Name           : new_DicOf()
'Overview                    : Dictionary�I�u�W�F�N�g�𐶐��������l��ݒ肷��
'Detailed Description        : �H����
'Argument
'     avParams               : �����l��i1,3,5,...�j��Key�A�����i2,4,6,...�j��Value
'                              Key�����̏ꍇ�͒l��Empty��ݒ肷��B
'Return Value
'     ��������Dictionary�I�u�W�F�N�g�̃C���X�^���X
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/03         Y.Fujii                  First edition
'***************************************************************************************************
Private Function new_DicOf( _
    byVal avParams _
    )
    Dim oDict, vItem, vKey, boIsKey
    
    boIsKey = True
    Set oDict = new_Dic()
    
    For Each vItem In avParams
        If boIsKey Then
            cf_bind vKey, vItem
            cf_bindAt oDict, vKey, Empty
        Else
            cf_bindAt oDict, vKey, vItem
        End If
        boIsKey = Not boIsKey
    Next
    
    Set new_DicOf = oDict
    Set oDict = Nothing
End Function

'***************************************************************************************************
'Function/Sub Name           : new_DriveOf()
'Overview                    : Drive�I�u�W�F�N�g�����֐�
'Detailed Description        : �H����
'Argument
'     asDriveName            : �h���C�u��
'Return Value
'     ��������Drive�I�u�W�F�N�g�̃C���X�^���X
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/11/11         Y.Fujii                  First edition
'***************************************************************************************************
Private Function new_DriveOf( _
    byVal asDriveName _
    )
    Set new_DriveOf = new_Fso().GetDrive(asDriveName)
End Function

'***************************************************************************************************
'Function/Sub Name           : new_Enum()
'Overview                    : Enum�����֐�
'Detailed Description        : Enum�̃C���X�^���X�𐶐�����
'Argument
'     asName                 : Enum��
'     aoDef                  : Enum�̒�`
'                              ��`��key����`���Avalue���l��Dictionary
'Return Value
'     �Ȃ�
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2024/05/04         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub new_Enum( _
    byVal asName _
    , byRef aoDef _
    )
    '�N���X���i�����j�쐬
    Dim sClassName : sClassName = "clsTmp_" & new_Fso().GetBaseName(new_Fso().GetTempName())
'    With new_Char()
'        Dim vCharList : vCharList = .charList(.typeHalfWidthAlphabetUppercase + .typeHalfWidthNumbers)
'    End With
'    cf_push vCharList, "_"
'    Dim sClassName : sClassName = "clsTmp_" & util_randStr(vCharList, 10)

    '�N���X��`�̃\�[�X�R�[�h�쐬
    Dim sThisName : sThisName = asName
    Dim vCode,i
    cf_push vCode, "Class " & sClassName
    cf_push vCode, "Private " & Join(aoDef.Keys,"_,")&"_"
    cf_push vCode, "Private PoLists"
    cf_push vCode, "Private Sub Class_Initialize()"
    cf_push vCode, "    Set PoLists = CreateObject('Scripting.Dictionary')"
    For Each i in aoDef.Keys
        cf_push vCode, "    Set " & i & "_ = (new ReadOnlyObject).of(Me, '" & i & "', " & aoDef.Item(i) & ")"
        cf_push vCode, "    cf_bindAt PoLists, '" & i & "', " & i
    Next
    cf_push vCode, "End Sub"
    For Each i in aoDef.Keys
        cf_push vCode, "Public Property Get " & i & "()"
        cf_push vCode, "    cf_bind " & i & ", " & i & "_"
        cf_push vCode, "End Property"
    Next
    cf_push vCode, "Public Property Get toString()"
    cf_push vCode, "    Dim i,ar"
    cf_push vCode, "    For Each i In PoLists.Items"
    cf_push vCode, "        cf_push ar, i.toString()"
    cf_push vCode, "    Next"
    cf_push vCode, "    toString = '<'&TypeName(Me)&'>(" & sThisName & "){'&Join(ar,',')&'}'"
    cf_push vCode, "End Property"
    cf_push vCode, "Public Function values()"
    cf_push vCode, "    values = PoLists.Items"
    cf_push vCode, "End Function"
    cf_push vCode, "Public Function valueOf(n)"
    cf_push vCode, "    ast_argTrue PoLists.Exists(n), TypeName(Me)&'(" & sThisName & ")+valueOf()', 'There is no element with the specified name'"
    cf_push vCode, "    Set valueOf = PoLists.Item(n)"
    cf_push vCode, "End Function"
    cf_push vCode, "End Class"
    '�C���X�^���X�����̃\�[�X�R�[�h�쐬
    cf_push vCode, "Private " & sThisName
    cf_push vCode, "Set " & sThisName & " = new " & sClassName
    '���s
    ExecuteGlobal Replace(Join(vCode,":"), "'", """")

End Sub

'***************************************************************************************************
'Function/Sub Name           : new_FileOf()
'Overview                    : File�I�u�W�F�N�g�����֐�
'Detailed Description        : �H����
'Argument
'     asPath                 : �p�X
'Return Value
'     ��������File�I�u�W�F�N�g�̃C���X�^���X
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/09/27         Y.Fujii                  First edition
'***************************************************************************************************
Private Function new_FileOf( _
    byVal asPath _
    )
    Set new_FileOf = new_Fso().GetFile(asPath)
End Function

'***************************************************************************************************
'Function/Sub Name           : new_FolderItem2Of()
'Overview                    : FolderItem2�I�u�W�F�N�g�����֐�
'Detailed Description        : �H����
'Argument
'     asPath                 : �p�X
'Return Value
'     ��������FolderItem2�I�u�W�F�N�g�̃C���X�^���X
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/02/22         Y.Fujii                  First edition
'***************************************************************************************************
Private Function new_FolderItem2Of( _
    byVal asPath _
    )
    With new_Fso()
        Set new_FolderItem2Of = new_ShellApp().Namespace(.GetParentFolderName(asPath)).Items().Item(.GetFileName(asPath))
    End With
End Function

'***************************************************************************************************
'Function/Sub Name           : new_FolderOf()
'Overview                    : Folder�I�u�W�F�N�g�����֐�
'Detailed Description        : �H����
'Argument
'     asPath                 : �p�X
'Return Value
'     ��������Folder�I�u�W�F�N�g�̃C���X�^���X
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/10/23         Y.Fujii                  First edition
'***************************************************************************************************
Private Function new_FolderOf( _
    byVal asPath _
    )
    Set new_FolderOf = new_Fso().GetFolder(asPath)
End Function

'***************************************************************************************************
'Function/Sub Name           : new_Fso()
'Overview                    : FileSystemObject�I�u�W�F�N�g�����֐�
'Detailed Description        : �H����
'Argument
'     �Ȃ�
'Return Value
'     ��������FileSystemObject�I�u�W�F�N�g�̃C���X�^���X
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/10/13         Y.Fujii                  First edition
'***************************************************************************************************
Private Function new_Fso( _
    )
    Set new_Fso = CreateObject("Scripting.FileSystemObject")
End Function

'***************************************************************************************************
'Function/Sub Name           : new_Func()
'Overview                    : �֐��̃C���X�^���X�𐶐�����
'Detailed Description        : javascript�̖����֐��ɏ����ivbscript�̎d�l�㉼�̖��O�͂���j
'Argument
'     asSoruceCode           : ��������֐��̃\�[�X�R�[�h
'                              �ȉ��̂����ꂩ�̗l���Ƃ��Afunction�isub�ł͂Ȃ��j�𐶐�����
'                              1.�ʏ�
'                               function (�@) {�A}
'                                �@�������J���}��؂�Ŏw�肷��
'                                �Avbscript�̍\���ɏ�������A�߂�l��"return hoge"�ƕ\�L����
'                                  "return"�傪�Ȃ��ꍇ�͖߂�l�͂Ȃ��Ƃ���
'                              2.Arrow�֐�
'                               �@ => �A
'                                �@�������J���}��؂�Ŏw�肷��A�����̏ꍇ��()�ň͂�
'                                �A�P��s�̏ꍇ�͂��̂܂ܖ߂�l�Ƃ���A�����s�̏ꍇ��1.�ʏ�̇A�Ɠ���
'Return Value
'     ���������֐��̃C���X�^���X
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/27         Y.Fujii                  First edition
'***************************************************************************************************
Private Function new_Func( _
    byVal asSoruceCode _
    )
    '��������֐��̃\�[�X�R�[�h�̉��s��:�ɕϊ�
    Dim sSoruceCode
    sSoruceCode = Replace(asSoruceCode, vbCrLf, ":")
    sSoruceCode = Replace(sSoruceCode, vbLf, ":")
    sSoruceCode = Replace(sSoruceCode, vbCr, ":")
    '��������֐��̃\�[�X�R�[�h��'�i�V���O���N�H�[�e�[�V�����j��"�i�_�u���N�H�[�e�[�V�����j�ɕϊ�
    sSoruceCode = Replace(sSoruceCode, "'", """")
    
    '�֐����i�����j�����
    Dim sFuncName : sFuncName = "anonymous_" & new_Fso().GetBaseName(new_Fso().GetTempName())
'    With new_Char()
'        Dim vCharList : vCharList = .charList(.typeHalfWidthAlphabetUppercase + .typeHalfWidthNumbers)
'    End With
'    cf_push vCharList, "_"
'    Dim sFuncName : sFuncName = "anonymous_" & util_randStr(vCharList, 10)
    
    Dim sPattern, oRegExp, sArgStr, sProcStr
    '��������֐��̃\�[�X�R�[�h�̗l�����u1.�ʏ�v�̏ꍇ
    sPattern = "function\s*\((.*)\)\s*{(.*)}"
    Set oRegExp = new_Re(sPattern, "igm")
    If oRegExp.Test(sSoruceCode) Then
        sArgStr = oRegExp.Replace(sSoruceCode, "$1")
        sProcStr = oRegExp.Replace(sSoruceCode, "$2")
        
        'return�傪����Ί֐����ŏ���������
        sProcStr = func_NewRewriteReturnPhrase(sFuncName, False, func_NewAnalyze(sProcStr) )
        
        '�֐��̐���
        Set new_Func = func_NewGenerate(sFuncName, sArgStr, sProcStr)
        Set oRegExp = Nothing
        Exit Function
    End If
    
    '��������֐��̃\�[�X�R�[�h�̗l�����u2.Arrow�֐��v�̏ꍇ
    sPattern = "(.*)\s*=>\s*(.*)\s*"
    Set oRegExp = new_Re(sPattern, "igm")
    If oRegExp.Test(sSoruceCode) Then
        sArgStr = oRegExp.Replace(sSoruceCode, "$1")
        sProcStr = oRegExp.Replace(sSoruceCode, "$2")
        
        '���ꂼ��O��̊��ʂ�����Ώ���
        sPattern = "\(\s?(.*)\s?\)"
        sArgStr = new_Re(sPattern, "igm").Replace(sArgStr, "$1")
        sPattern = "{\s?(.*)\s?}"
        sProcStr = new_Re(sPattern, "igm").Replace(sProcStr, "$1")
        
        'return�傪����Ί֐����ŏ���������
        sProcStr = func_NewRewriteReturnPhrase(sFuncName, True, func_NewAnalyze(sProcStr) )
        
        '�֐��̐���
        Set new_Func = func_NewGenerate(sFuncName, sArgStr, sProcStr)
    End If
    Set oRegExp = Nothing
End Function

'***************************************************************************************************
'Function/Sub Name           : new_HtmlOf()
'Overview                    : �C���X�^���X�����֐�
'Detailed Description        : HTML�����N���X�̃C���X�^���X��Ԃ�
'Argument
'     asElement              : �v�f
'Return Value
'     HTML�����N���X�̃C���X�^���X
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/11/03         Y.Fujii                  First edition
'***************************************************************************************************
Private Function new_HtmlOf( _
    byVal asElement _
    )
    Dim oHtml : Set oHtml = New HtmlGenerator
    oHtml.element = asElement
    Set new_HtmlOf = oHtml
    Set oHtml = Nothing
End Function

'***************************************************************************************************
'Function/Sub Name           : new_Network()
'Overview                    : WScript.Network�I�u�W�F�N�g�����֐�
'Detailed Description        : �H����
'Argument
'     �Ȃ�
'Return Value
'     ��������WScript.Network�I�u�W�F�N�g�̃C���X�^���X
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/12/27         Y.Fujii                  First edition
'***************************************************************************************************
Private Function new_Network( _
    )
    Set new_Network = CreateObject("WScript.Network")
End Function

'***************************************************************************************************
'Function/Sub Name           : new_Now()
'Overview                    : �C���X�^���X�����֐�
'Detailed Description        : ���̓��t�����Ő����������t�N���X�̃C���X�^���X��Ԃ�
'Argument
'     �Ȃ�
'Return Value
'     ���t�N���X�̃C���X�^���X
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/01/04         Y.Fujii                  First edition
'***************************************************************************************************
Private Function new_Now( _
    )
    Set new_Now = (New Calendar).ofNow()
End Function

'***************************************************************************************************
'Function/Sub Name           : new_Re()
'Overview                    : ���K�\���I�u�W�F�N�g�����֐�
'Detailed Description        : �H����
'Argument
'     asPattern              : ���K�\���̃p�^�[��
'     asOptions              : ���̈������ɂ��镶���̗L���Ő��K�\���̈ȉ��̃v���p�e�B��True�ɂ���
'                                "i":�啶���Ə���������ʂ���i.IgnoreCase = True�j
'                                "g"������S�̂���������i.Global = True�j
'                                "m"������𕡐��s�Ƃ��Ĉ����i.Multiline = True�j
'Return Value
'     �����������K�\���I�u�W�F�N�g�̃C���X�^���X
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/03         Y.Fujii                  First edition
'***************************************************************************************************
Private Function new_Re( _
    byVal asPattern _
    , byVal asOptions _
    )
    Dim oRe, sOpts
    
    Set oRe = New RegExp
    oRe.Pattern = asPattern
    
    sOpts = LCase(asOptions)
    If InStr(sOpts, "i") > 0 Then oRe.IgnoreCase = True
    If InStr(sOpts, "g") > 0 Then oRe.Global = True
    If InStr(sOpts, "m") > 0 Then oRe.Multiline = True
    
    Set new_Re = oRe
    Set oRe = Nothing
End Function

'***************************************************************************************************
'Function/Sub Name           : new_Reader()
'Overview                    : �t�@�C���Ǎ��o�b�t�@�����O�����N���X�̃C���X�^���X�����֐�
'Detailed Description        : �H����
'Argument
'     aoTextStream           : �e�L�X�g�X�g���[���I�u�W�F�N�g
'Return Value
'     ���������t�@�C���Ǎ��o�b�t�@�����O�����N���X�̃C���X�^���X
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/10/02         Y.Fujii                  First edition
'***************************************************************************************************
Private Function new_Reader( _
    byRef aoTextStream _
    )
    Set new_Reader = (New BufferedReader).setTextStream(aoTextStream)
End Function

'***************************************************************************************************
'Function/Sub Name           : new_ReaderOf()
'Overview                    : �t�@�C���Ǎ��o�b�t�@�����O�����N���X�̃C���X�^���X�����֐�
'Detailed Description        : �H����
'Argument
'     asPath                 : �ǂݍ��ރt�@�C���̃p�X
'Return Value
'     ���������t�@�C���Ǎ��o�b�t�@�����O�����N���X�̃C���X�^���X
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/10/09         Y.Fujii                  First edition
'***************************************************************************************************
Private Function new_ReaderOf( _
    byVal asPath _
    )
    Set new_ReaderOf = (New BufferedReader).setTextStream(new_Ts(asPath, 1, False, -2))
End Function

'***************************************************************************************************
'Function/Sub Name           : new_Ret()
'Overview                    : �߂�l�N���X�I�u�W�F�N�g�̐����֐�
'Detailed Description        : �H����
'Argument
'     avRet                  : �߂�l
'Return Value
'     ���������߂�l�N���X�̃C���X�^���X
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2024/01/03         Y.Fujii                  First edition
'***************************************************************************************************
Private Function new_Ret( _
    byRef avRet _
    )
    Set new_Ret = (New ReturnValue).setValue(avRet)
End Function

'***************************************************************************************************
'Function/Sub Name           : new_RetByState()
'Overview                    : �߂�l�N���X�I�u�W�F�N�g�̐����֐�
'Detailed Description        : �H����
'Argument
'     avNormal               : ����̏ꍇ�̖߂�l
'     avAbnormal             : �ُ�̏ꍇ�̖߂�l
'Return Value
'     ���������߂�l�N���X�̃C���X�^���X
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2024/04/27         Y.Fujii                  First edition
'***************************************************************************************************
Private Function new_RetByState( _
    byRef avNormal _
    , byRef avAbnormal _
    )
    Set new_RetByState = (New ReturnValue).setValueByState(avNormal,avAbnormal)
End Function

'***************************************************************************************************
'Function/Sub Name           : new_Shell()
'Overview                    : Wscript.Shell�I�u�W�F�N�g�����֐�
'Detailed Description        : �H����
'Argument
'     �Ȃ�
'Return Value
'     ��������Wscript.Shell�I�u�W�F�N�g�̃C���X�^���X
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/12/19         Y.Fujii                  First edition
'***************************************************************************************************
Private Function new_Shell( _
    )
    Set new_Shell = CreateObject("Wscript.Shell")
End Function

'***************************************************************************************************
'Function/Sub Name           : new_ShellApp()
'Overview                    : Shell.Application�I�u�W�F�N�g�����֐�
'Detailed Description        : �H����
'Argument
'     �Ȃ�
'Return Value
'     ��������Shell.Application�I�u�W�F�N�g�̃C���X�^���X
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/11/25         Y.Fujii                  First edition
'***************************************************************************************************
Private Function new_ShellApp( _
    )
    Set new_ShellApp = CreateObject("Shell.Application")
End Function

'***************************************************************************************************
'Function/Sub Name           : new_Ts()
'Overview                    : TextStream�I�u�W�F�N�g�����֐�
'Detailed Description        : FileSystemObject��OpenTextFile()�Ɠ���
'Argument
'     asPath                 : �p�X
'     alIomode               : ����/�o�̓��[�h 1:ForReading,2:ForWriting,8:ForAppending
'     aboCreate              : asPath�����݂��Ȃ��ꍇ True:�V�����t�@�C�����쐬����AFalse:�쐬���Ȃ�
'     alFileFormat           : �t�@�C���̌`�� -2:TristateUseDefault,-1:TristateTrue,0:TristateFalse
'Return Value
'     ��������TextStream�I�u�W�F�N�g�̃C���X�^���X
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/01/09         Y.Fujii                  First edition
'***************************************************************************************************
Private Function new_Ts( _
    byVal asPath _
    , byVal alIomode _
    , byVal aboCreate _
    , byVal alFileFormat _
    )
    Set new_Ts = new_Fso().OpenTextFile(asPath, alIomode, aboCreate, alFileFormat)
End Function

'***************************************************************************************************
'Function/Sub Name           : new_Writer()
'Overview                    : �t�@�C���o�̓o�b�t�@�����O�����N���X�̃C���X�^���X�����֐�
'Detailed Description        : �H����
'Argument
'     aoTextStream           : �e�L�X�g�X�g���[���I�u�W�F�N�g
'Return Value
'     ���������t�@�C���o�̓o�b�t�@�����O�����N���X�̃C���X�^���X
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/08/27         Y.Fujii                  First edition
'***************************************************************************************************
Private Function new_Writer( _
    byRef aoTextStream _
    )
    Set new_Writer = (New BufferedWriter).setTextStream(aoTextStream)
End Function

'***************************************************************************************************
'Function/Sub Name           : new_WriterOf()
'Overview                    : �t�@�C���o�̓o�b�t�@�����O�����N���X�̃C���X�^���X�����֐�
'Detailed Description        : �H����
'Argument
'     asPath                 : �������ރt�@�C���̃p�X
'     alIomode               : �o�̓��[�h 2:ForWriting,8:ForAppending
'     aboCreate              : asPath�����݂��Ȃ��ꍇ True:�V�����t�@�C�����쐬����AFalse:�쐬���Ȃ�
'     alFileFormat           : �t�@�C���̌`�� -2:TristateUseDefault,-1:TristateTrue,0:TristateFalse
'Return Value
'     ���������t�@�C���o�̓o�b�t�@�����O�����N���X�̃C���X�^���X
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/10/09         Y.Fujii                  First edition
'***************************************************************************************************
Private Function new_WriterOf( _
    byVal asPath _
    , byVal alIomode _
    , byVal aboCreate _
    , byVal alFileFormat _
    )
    Set new_WriterOf = (New BufferedWriter).setTextStream(new_Ts(asPath, alIomode, aboCreate, alFileFormat))
End Function

'---------------------------------------------------------------------------------------------------

'***************************************************************************************************
'Function/Sub Name           : func_NewAnalyze()
'Overview                    : �\�[�X�R�[�h�����߂���
'Detailed Description        : new_Func()����g�p����
'                              _�i�A���_�[���C���j�͍s����������
'Argument
'     asCode                 : �\�[�X�R�[�h
'Return Value
'     �\�[�X�R�[�h���s���Ƃɕ��������z��
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/10/12         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_NewAnalyze( _
    byVal asCode _
    )
    Dim sRow, sPtn, oCode, sTemp
    Set oCode = new_Dic()
    sTemp= ""
    For Each sRow In Split(asCode, ":", -1, vbBinaryCompare)
        If Len(Trim(sRow))>0 Then
            sPtn = "^(.*)\s_\s*$"
            If new_Re(sPtn, "ig").Test(sRow) Then
                sTemp = sTemp & Trim(new_Re(sPtn, "ig").Replace(sRow, "$1"))
            Else
                sRow = sTemp & " " & Trim(sRow)
                sTemp = ""
                oCode.Add oCode.Count, Trim(sRow)
            End If
        End If
    Next
    
    func_NewAnalyze = oCode.Items()
    Set oCode = Nothing
End Function

'***************************************************************************************************
'Function/Sub Name           : func_NewGenerate()
'Overview                    : �����̏��Ŋ֐��̃C���X�^���X�𐶐�����
'Detailed Description        : new_Func()����g�p����
'Argument
'     asFuncName             : �֐���
'     asArgStr               : �\�[�X�̈��������̃\�[�X�R�[�h
'     asProcStr              : �\�[�X�̏������e�����̃\�[�X�R�[�h
'Return Value
'     ���������֐��̃C���X�^���X
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/27         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_NewGenerate( _
    byVal asFuncName _
    , byVal asArgStr _
    , byVal asProcStr _
    )
    Dim sCode
    '�\�[�X�R�[�h�쐬
    sCode = _
        "Private Function " & asFuncName & "(" & asArgStr & ")" & vbNewLine _
        & asProcStr & vbNewLine _
        & "End Function"
    
'inputbox "","",sCode
    '�֐��̐���
    ExecuteGlobal sCode
    Set func_NewGenerate = Getref(asFuncName)
End Function

'***************************************************************************************************
'Function/Sub Name           : func_NewRewriteReturnPhrase()
'Overview                    : return�������������
'Detailed Description        : new_Func()����g�p����
'                              Arrow�֐���1�s�̏ꍇ�͂��̍s�S�̂�return����
'Argument
'     asFuncName             : �֐���
'     aboArrowFlg            : Arrow�֐����ۂ��̃t���O
'     avCode                 : �\�[�X�R�[�h���s���Ƃɕ��������z��
'Return Value
'     �����������\�[�X�̏������e�����̃\�[�X�R�[�h
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/27         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_NewRewriteReturnPhrase( _
    byVal asFuncName _
    , byVal aboArrowFlg _
    , byRef avCode _
    )
    Dim sPtnRet : sPtnRet = "^(.*\s+)?return\s+(.*)\s{0,}$"
    
    If Ubound(avCode)=0 And aboArrowFlg=True Then
    'Arrow�֐���1�s�̏ꍇ
        Dim sCode : sCode = avCode(0)
        If new_Re(sPtnRet, "ig").Test(sCode) Then
        'return�傪����ꍇ
            func_NewRewriteReturnPhrase = new_Re(sPtnRet, "ig").Replace(sCode, "$1 cf_bind " & asFuncName & ", ($2)")
        Else
        'return�傪�Ȃ��ꍇ
            func_NewRewriteReturnPhrase = "cf_bind " & asFuncName & ", (" & sCode & ")"
        End If
        Exit Function
    End If
    
    Dim lCnt, sPtn, sRow
    For lCnt=0 To Ubound(avCode)
        sRow = avCode(lCnt)
        If new_Re(sPtnRet, "ig").Test(sRow) Then
            avCode(lCnt) = new_Re(sPtnRet, "ig").Replace(sRow, "$1 cf_bind " & asFuncName & ", ($2)")
        End If
    Next
    
    func_NewRewriteReturnPhrase = Join(avCode, ":")
    
End Function

'###################################################################################################
'���w�n�̊֐�
'###################################################################################################

'***************************************************************************************************
'Function/Sub Name           : math_min()
'Overview                    : �ŏ��l�����߂�
'Detailed Description        : �H����
'Argument
'     al1                    : ���l1
'     al2                    : ���l2
'Return Value
'     al1��al2�̒l����������
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/09/27         Y.Fujii                  First edition
'***************************************************************************************************
Private Function math_min( _
    byVal al1 _ 
    , byVal al2 _
    )
    cf_bind math_min, math_minA(Array(al1, al2))
'    Dim lRet
'    If al1 < al2 Then lRet = al1 Else lRet = al2
'    math_min = lRet
End Function

'***************************************************************************************************
'Function/Sub Name           : math_minA()
'Overview                    : �ŏ��l�����߂�
'Detailed Description        : �H����
'Argument
'     avNums                 : ���l
'Return Value
'     avNums�̍ŏ��l
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2025/03/01         Y.Fujii                  First edition
'***************************************************************************************************
Private Function math_minA( _
    byRef avNums _
    )
    cf_bind math_minA, avNums
    If new_Arr().hasElement(avNums) Then cf_bind math_minA, new_ArrOf(avNums).sort(True)(0)
End Function

'***************************************************************************************************
'Function/Sub Name           : math_max()
'Overview                    : �ő�l�����߂�
'Detailed Description        : �H����
'Argument
'     al1                    : ���l1
'     al2                    : ���l2
'Return Value
'     al1��al2�̒l���傫����
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/09/27         Y.Fujii                  First edition
'***************************************************************************************************
Private Function math_max( _
    byVal al1 _ 
    , byVal al2 _
    )
    cf_bind math_max, math_maxA(Array(al1, al2))
'    Dim lRet
'    If al1 > al2 Then lRet = al1 Else lRet = al2
'    math_max = lRet
End Function

'***************************************************************************************************
'Function/Sub Name           : math_maxA()
'Overview                    : �ő�l�����߂�
'Detailed Description        : �H����
'Argument
'     avNums                 : ���l
'Return Value
'     avNums�̍ő�l
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2025/03/01         Y.Fujii                  First edition
'***************************************************************************************************
Private Function math_maxA( _
    byRef avNums _
    )
    cf_bind math_maxA, avNums
    If new_Arr().hasElement(avNums) Then cf_bind math_maxA, new_ArrOf(avNums).sort(False)(0)
End Function

'***************************************************************************************************
'Function/Sub Name           : math_roundUp()
'Overview                    : �؂�グ����
'Detailed Description        : func_MathRound()�ɈϏ�����
'Argument
'     adbNum                 : ���l
'     alPlace                : �����̈ʁA�؂�グ����[���̈ʒu�������̈ʂŕ\��
'Return Value
'     �؂�グ�����l
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/09/24         Y.Fujii                  First edition
'***************************************************************************************************
Private Function math_roundUp( _
    byVal adbNum _ 
    , byVal alPlace _
    )
    math_roundUp = func_MathRound(adbNum, alPlace, 9, True)
End Function

'***************************************************************************************************
'Function/Sub Name           : math_round()
'Overview                    : �l�̌ܓ�����
'Detailed Description        : func_MathRound()�ɈϏ�����
'Argument
'     adbNum                 : ���l
'     alPlace                : �����̈ʁA�l�̌ܓ�����[���̈ʒu�������̈ʂŕ\��
'Return Value
'     �l�̌ܓ������l
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/09/24         Y.Fujii                  First edition
'***************************************************************************************************
Private Function math_round( _
    byVal adbNum _ 
    , byVal alPlace _
    )
    math_round = func_MathRound(adbNum, alPlace, 5, True)
End Function

'***************************************************************************************************
'Function/Sub Name           : math_roundDown()
'Overview                    : �؂�̂Ă���
'Detailed Description        : func_MathRound()�ɈϏ�����
'Argument
'     adbNum                 : ���l
'     alPlace                : �����̈ʁA�؂�̂Ă���[���̈ʒu�������̈ʂŕ\��
'Return Value
'     �؂�̂Ă����l
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/09/24         Y.Fujii                  First edition
'***************************************************************************************************
Private Function math_roundDown( _
    byVal adbNum _ 
    , byVal alPlace _
    )
    math_roundDown = func_MathRound(adbNum, alPlace, 0, True)
End Function

'***************************************************************************************************
'Function/Sub Name           : math_rand()
'Overview                    : �����𐶐�����
'Detailed Description        : �H����
'Argument
'     adbMin                 : �������闐���̍ŏ��l
'     adbMax                 : �������闐���̍ő�l
'     alPlace                : �����̈ʁA�؂�グ����[���̈ʒu�������̈ʂŕ\��
'Return Value
'     ������������
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/24         Y.Fujii                  First edition
'***************************************************************************************************
Private Function math_rand( _
    byVal adbMin _
    , byVal adbMax _
    , byVal alPlace _
    )
    Randomize
    math_rand = adbMin + Fix( ((adbMax-adbMin)*(10^alPlace)+1)*Rnd )*10^(-1*alPlace)
End Function

'***************************************************************************************************
'Function/Sub Name           : math_log2()
'Overview                    : 2����̑ΐ�
'Detailed Description        : �H����
'Argument
'     adbAntilogarithm       : �^��
'Return Value
'     �p�w��
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/11/03         Y.Fujii                  First edition
'***************************************************************************************************
Private Function math_log2( _
    byVal adbAntilogarithm _
    )
    math_log2 = func_MathLog(2, adbAntilogarithm)
End Function

'***************************************************************************************************
'Function/Sub Name           : math_tranc()
'Overview                    : ��������Ԃ�
'Detailed Description        : �[�������Ɋۂ߂���������Ԃ�
'                               10.8  -> 10
'                               -10.8 -> -10
'Argument
'     adbNum                 : ���l
'Return Value
'     ������
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2025/02/11         Y.Fujii                  First edition
'***************************************************************************************************
Private Function math_tranc( _
    byVal adbNum _ 
    )
    math_tranc = func_MathRound(adbNum,0,0,True)
End Function

'***************************************************************************************************
'Function/Sub Name           : math_fractional()
'Overview                    : ��������Ԃ�
'Detailed Description        : �[�������Ɋۂ߂���������Ԃ�
'                               10.8  -> 0.8
'                               -10.8 -> -0.8
'Argument
'     adbNum                 : ���l
'Return Value
'     ������
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2025/02/11         Y.Fujii                  First edition
'***************************************************************************************************
Private Function math_fractional( _
    byVal adbNum _ 
    )
    math_fractional = adbNum-func_MathRound(adbNum,0,0,True)
End Function

'***************************************************************************************************
'Function/Sub Name           : func_MathRound()
'Overview                    : ���l���ۂ߂�
'Detailed Description        : �[�������̕��@�ɏ]���Đ��l���ۂ߂�
'                              ������alPlace�͊ۂ߂��������̈ʂ��w�肷��A�Ⴆ�Α��ʂ��ꍇ��0���w�肷��
'                              �������ʂ��l�̌ܓ�����ꍇ�AalPlace��1�AalThreshold��5���w�肷��
'                              ��̈ʁA�\�̈ʁA�E�E�E�̏ꍇ��-1,-2,�c�̂悤�ɕ��l���w�肷��
'                                ��j�P�W�Q�D�V�R�Q
'                                �@�@���@�@���@���@ ��
'                                   -3  -1�@0�@ 2
'Argument
'     adbNum                 : ���l
'     alPlace                : �����̈ʁA��������[���̈ʒu�������̈ʂŕ\��
'     alThreshold            : 臒l
'                               0�F�؂�̂�
'                               5�F�l�̌ܓ�
'                               9�F�؂�グ
'     aboMode                : �[�������̕��@
'                               True  �F�����𖳎����Đ�Βl���ۂ߂�i�����Ŋۂ߂�������قȂ�j
'                               False �F�����̏ꍇ�Ƒ����𓯂������Ɋۂ߂�
'                              https://ja.wikipedia.org/wiki/%E7%AB%AF%E6%95%B0%E5%87%A6%E7%90%86
'Return Value
'     �ۂ߂��l
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/09/24         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_MathRound( _
    byVal adbNum _ 
    , byVal alPlace _
    , byVal alThreshold _
    , byVal aboMode _
    )
    Dim lThreshold : lThreshold = alThreshold
    If adbNum<0 Then lThreshold = -1*lThreshold

    Dim dbTemp
    dbTemp = Cstr((adbNum+lThreshold*10^(-1*(alPlace+1))) * 10^(alPlace))

    If aboMode Then
        func_MathRound = Cdbl( Cstr( Fix(dbTemp) * 10^(-1*alPlace) ) )
    Else
        func_MathRound = Cdbl( Cstr( Int(dbTemp) * 10^(-1*alPlace) ) )
    End If
End Function

'***************************************************************************************************
'Function/Sub Name           : func_MathLog()
'Overview                    : �������Ƃ���ΐ�
'Detailed Description        : �H����
'Argument
'     adbBase                : ��
'     adbAntilogarithm       : �^��
'Return Value
'     �p�w��
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/11/03         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_MathLog( _
    byVal adbBase _
    , byVal adbAntilogarithm _
    )
    func_MathLog = log(adbAntilogarithm)/log(adbBase)
End Function


'###################################################################################################
'���[�e�B���e�B�n�̊֐�
'###################################################################################################

'***************************************************************************************************
'Function/Sub Name           : util_escapeForPs()
'Overview                    : powershell�p�̓��ꕶ���G�X�P�[�v���s��
'Detailed Description        : �����T�C�g�ɏ����Ă��Ȃ���shell����N������ꍇ�ɂ������Ή����Ȃ��Ɠ��삵�Ȃ����̂�����
'                              https://learn.microsoft.com/ja-jp/powershell/module/microsoft.powershell.core/about/about_special_characters?view=powershell-7.4
'Argument
'     asTgt                  : �Ώ�
'Return Value
'     �ΏۂɃG�X�P�[�v��������������
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2024/02/05         Y.Fujii                  First edition
'***************************************************************************************************
Private Function util_escapeForPs( _
    byVal asTgt _
    )
    Const Cs_BACKQUOTE = "`"
    Dim vLst : vLst = Array("(",")"," ")
    
    Dim i, sRet : sRet = asTgt
    For Each i In vLst
        sRet = Replace(sRet, i, Cs_BACKQUOTE&i)
    Next
    util_escapeForPs = sRet
End Function

'***************************************************************************************************
'Function/Sub Name           : util_getIpAddress()
'Overview                    : ���g��IP�A�h���X���擾����
'Detailed Description        : IP�A�h���X���i�[�����I�u�W�F�N�g��Ԃ�
'Argument
'     �Ȃ�
'Return Value
'     IP�A�h���X���i�[�����I�u�W�F�N�g�̔z��
'                              ���e�͈ȉ��̂Ƃ���
'                               Key             Value                   ��
'                               --------------  ----------------------  ----------------------------
'                               "Caption"       Adapter��               -
'                               "Ip"            �ȉ��I�u�W�F�N�g        -
'                              
'                              IP Address���i�[�����I�u�W�F�N�g
'                               Key             Value                   ��
'                               --------------  ----------------------  ----------------------------
'                               "V4"            IP Address(v4)          192.168.11.52
'                               "V6"            IP Address(v6)          fe80::ba87:1e93:59ab:28f7%18
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/10/10         Y.Fujii                  First edition
'***************************************************************************************************
Private Function util_getIpAddress( _
    )
    Dim sMyComp, oAdapter, oAddress, oRet, oIpv4, oIpv6
    
    For Each oAdapter in CreateObject("WbemScripting.SWbemLocator").ConnectServer().ExecQuery("Select * From Win32_NetworkAdapterConfiguration Where IPEnabled = True")
         For Each oAddress in oAdapter.IPAddress
             If new_ArrSplit(oAddress, ".").length=4 Then
             'IPv4
                 cf_bind oIpv4, oAddress
             Else
             'IPv6
                 cf_bind oIpv6, oAddress
             End If
         Next
         cf_push oRet, new_DicOf(Array("Caption", oAdapter.Caption, "Ip", new_DicOf(Array("V4", oIpv4, "V6", oIpv6))))
    Next
    util_getIpAddress = oRet
    
    Set oAddress = Nothing
    Set oAdapter = Nothing
End Function

'***************************************************************************************************
'Function/Sub Name           : util_isZipWithPassword()
'Overview                    : �p�X���[�h�t��zip�t�@�C�������肷��
'Detailed Description        : https://vbavb.com/zip/
'Argument
'     asPath                 : �p�X
'Return Value
'     ���� True:�p�X���[�h�t��zip�t�@�C�� / False:�p�X���[�h�t��zip�t�@�C���łȂ�
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2024/01/29         Y.Fujii                  First edition
'***************************************************************************************************
Private Function util_isZipWithPassword( _
    byVal asPath _
    )
    util_isZipWithPassword = False
    If Not new_Fso().FileExists(asPath) Then Exit Function
    With new_Adodb()
        .Type = 1       'adTypeBinary
        .Open
        .LoadFromFile asPath
        .Position = 6
        If Hex(AscB(.Read(1)))=1 Then util_isZipWithPassword = True
        .Close
    End With
End Function

'***************************************************************************************************
'Function/Sub Name           : util_randStr()
'Overview                    : �����_���ȕ�����𐶐�����
'Detailed Description        : �w�肵�������i�z��j�A�w�肵���񐔂Ń����_���ȕ�����𐶐�����
'Argument
'     avStrings              : �����̔z��
'     alTimes                : ��
'Return Value
'     ��������������
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/11/03         Y.Fujii                  First edition
'***************************************************************************************************
Private Function util_randStr( _
    byRef avStrings _
    , byVal alTimes _
    )
    Dim lPos, sRet, lUb
    sRet = "" : lUb = Ubound(avStrings)
    For lPos = 1 To alTimes
        sRet = sRet & avStrings( math_rand(0, lUb, 0) )
    Next
    util_randStr = sRet
End Function

'***************************************************************************************************
'Function/Sub Name           : util_unzip()
'Overview                    : zip�t�@�C�����𓀂���
'Detailed Description        : https://excel-vba.work/2021/12/10/%E3%80%90vba%E3%80%91zip%E3%83%95%E3%82%A1%E3%82%A4%E3%83%AB%E3%82%92%E8%A7%A3%E5%87%8D%E5%B1%95%E9%96%8B%E3%81%99%E3%82%8B/#google_vignette
'Argument
'     asPath                 : zip�t�@�C���̃p�X
'     asDestination          : �𓀐�
'Return Value
'     ���� True:���� / False:���s
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2024/02/03         Y.Fujii                  First edition
'***************************************************************************************************
Private Function util_unzip( _
    byVal asPath _
    , byVal asDestination _
    )
    util_unzip=False
    'PowerShell�̃R�}���h���쐬
    Dim sCmd : sCmd = _
        "powershell -NoProfile -ExecutionPolicy Unrestricted Expand-Archive -Force" _
        & " -Path " & fs_wrapInQuotes(util_escapeForPs(asPath)) _
        & " -DestinationPath " & fs_wrapInQuotes(util_escapeForPs(asDestination))
    '�쐬�����R�}���h���T�C�����g���s����
    util_unzip = fw_runShellSilently(sCmd)
End Function

'***************************************************************************************************
'Function/Sub Name           : util_zip()
'Overview                    : zip�t�@�C�����쐬����
'Detailed Description        : https://learn.microsoft.com/ja-jp/powershell/module/microsoft.powershell.archive/compress-archive?view=powershell-7.4
'Argument
'     asPath                 : ���k����t�@�C���̃p�X
'     asDestination          : zip�t�@�C���̃p�X
'Return Value
'     ���� True:���� / False:���s
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2024/02/05         Y.Fujii                  First edition
'***************************************************************************************************
Private Function util_zip( _
    byRef asPath _
    , byVal asDestination _
    )
    util_zip=False
    If new_Fso().FileExists(asDestination) Then Exit Function

    '���k����t�@�C���̃p�X��A��
    Dim sPath : sPath = new_ArrOf(asPath).map(new_Func("(e,i,a)=>fs_wrapInQuotes(util_escapeForPs(e))")).join(",")

    'PowerShell�̃R�}���h���쐬
    Dim sCmd : sCmd = _
        "powershell -NoProfile -ExecutionPolicy Unrestricted Compress-Archive" _
        & " -Path " & sPath _
        & " -DestinationPath " & fs_wrapInQuotes(util_escapeForPs(asDestination))

    '�쐬�����R�}���h���T�C�����g���s����
    util_zip = fw_runShellSilently(sCmd)
End Function

'###################################################################################################
'�t�@�C������n
'###################################################################################################

'***************************************************************************************************
'Function/Sub Name           : fs_copyFile()
'Overview                    : �t�@�C�����R�s�[����
'Detailed Description        : FileSystemObject��CopyFile()�Ɠ���
'                              func_FsGeneralExecutor()�ɈϏ�����
'Argument
'     asFrom                 : �R�s�[���t�@�C���̃t���p�X
'     asTo                   : �R�s�[��t�@�C���̃t���p�X
'Return Value
'     ���� True:���� / False:���s
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/10/24         Y.Fujii                  First edition
'***************************************************************************************************
Private Function fs_copyFile( _
    byVal asFrom _
    , byVal asTo _
    )
    Set fs_copyFile = func_FsGeneralExecutor(Array(asFrom, asTo), "CopyFile")
End Function

'***************************************************************************************************
'Function/Sub Name           : fs_copyFolder()
'Overview                    : �t�H���_���R�s�[����
'Detailed Description        : FileSystemObject��CopyFolder()�Ɠ���
'                              func_FsGeneralExecutor()�ɈϏ�����
'Argument
'     asFrom                 : �R�s�[���t�H���_�̃t���p�X
'     asTo                   : �R�s�[��t�H���_�̃t���p�X
'Return Value
'     ���� True:���� / False:���s
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/08/19         Y.Fujii                  First edition
'***************************************************************************************************
Private Function fs_copyFolder( _
    byVal asFrom _
    , byVal asTo _
    )
    Set fs_copyFolder = func_FsGeneralExecutor(Array(asFrom, asTo), "CopyFolder")
End Function

'***************************************************************************************************
'Function/Sub Name           : fs_copyHere()
'Overview                    : �t�H���_�܂��̓t�@�C�����t�H���_�[�ɃR�s�[����
'Detailed Description        : Windows�V�F���֐���folder.Copyhere()����
'                              https://learn.microsoft.com/ja-jp/windows/win32/shell/folder-copyhere
'                              �R�s�[��t�H���_�����݂��Ȃ��ꍇ�͍쐬����
'Argument
'     asPath                 : �R�s�[����t�H���_�܂��̓t�@�C���̃p�X
'     asDestination          : �R�s�[��t�H���_�̃p�X
'Return Value
'     ���� True:���� / False:���s
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2024/02/18         Y.Fujii                  First edition
'***************************************************************************************************
Private Function fs_copyHere( _
    byVal asPath _
    , byVal asDestination _
    )

    On Error Resume Next
    '�R�s�[����t�H���_�܂��̓t�@�C���̑��݊m�F
    new_FolderItem2Of(asPath)
    If Err.Number<>0 Then
        Set fs_copyHere = new_Ret(False)
        Err.Clear
        Exit Function
    End If

    '�R�s�[
    new_ShellApp().Namespace(asDestination).CopyHere asPath
    If Err.Number<>0 Then
        Set fs_copyHere = new_Ret(False)
        Err.Clear
        Exit Function
    End If
    On Error Goto 0

    '�R�s�[��̑��݊m�F
    With new_Fso()
        Dim sPath : sPath = .BuildPath(asDestination, .GetFileName(asPath))
    End With
    On Error Resume Next
    new_FolderItem2Of(sPath)
    If Err.Number<>0 Then
        Set fs_copyHere = new_Ret(False)
        Err.Clear
        Exit Function
    End If
    On Error Goto 0

    Set fs_copyHere = new_Ret(True)

End Function

'***************************************************************************************************
'Function/Sub Name           : fs_createFolder()
'Overview                    : �t�H���_���쐬����
'Detailed Description        : FileSystemObject��CreateFolder()�Ɠ���
'                              func_FsGeneralExecutor()�ɈϏ�����
'Argument
'     asPath                 : �쐬����t�H���_�̃t���p�X
'Return Value
'     ���� True:���� / False:���s
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/12/30         Y.Fujii                  First edition
'***************************************************************************************************
Private Function fs_createFolder( _
    byVal asPath _
    )
    Set fs_createFolder = func_FsGeneralExecutor(Array(asPath), "CreateFolder")
End Function

'***************************************************************************************************
'Function/Sub Name           : fs_deleteFile()
'Overview                    : �t�@�C�����폜����
'Detailed Description        : FileSystemObject��DeleteFile()�Ɠ���
'                              func_FsGeneralExecutor()�ɈϏ�����
'Argument
'     asPath                 : �폜����t�@�C���̃t���p�X
'Return Value
'     ���� True:���� / False:���s
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/09/27         Y.Fujii                  First edition
'***************************************************************************************************
Private Function fs_deleteFile( _
    byVal asPath _
    )
    Set fs_deleteFile = func_FsGeneralExecutor(Array(asPath), "DeleteFile")
End Function

'***************************************************************************************************
'Function/Sub Name           : fs_deleteFolder()
'Overview                    : �t�H���_���폜����
'Detailed Description        : FileSystemObject��DeleteFolder()�Ɠ���
'                              func_FsGeneralExecutor()�ɈϏ�����
'Argument
'     asPath                 : �폜����t�H���_�̃t���p�X
'Return Value
'     ���� True:���� / False:���s
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/11/12         Y.Fujii                  First edition
'***************************************************************************************************
Private Function fs_deleteFolder( _
    byVal asPath _
    )
    Set fs_deleteFolder = func_FsGeneralExecutor(Array(asPath), "DeleteFolder")
End Function

'***************************************************************************************************
'Function/Sub Name           : fs_moveFile()
'Overview                    : �t�@�C�����ړ�����
'Detailed Description        : FileSystemObject��MoveFile()�Ɠ���
'                              func_FsGeneralExecutor()�ɈϏ�����
'Argument
'     asFrom                 : �ړ����t�@�C���̃t���p�X
'     asTo                   : �ړ���t�@�C���̃t���p�X
'Return Value
'     ���� True:���� / False:���s
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/08/19         Y.Fujii                  First edition
'***************************************************************************************************
Private Function fs_moveFile( _
    byVal asFrom _
    , byVal asTo _
    )
    Set fs_moveFile = func_FsGeneralExecutor(Array(asFrom, asTo), "MoveFile")
End Function

'***************************************************************************************************
'Function/Sub Name           : fs_moveFolder()
'Overview                    : �t�H���_���ړ�����
'Detailed Description        : FileSystemObject��MoveFolder()�Ɠ���
'                              func_FsGeneralExecutor()�ɈϏ�����
'Argument
'     asFrom                 : �ړ����t�H���_�̃t���p�X
'     asTo                   : �ړ���t�H���_�̃t���p�X
'Return Value
'     ���� True:���� / False:���s
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/08/19         Y.Fujii                  First edition
'***************************************************************************************************
Private Function fs_moveFolder( _
    byVal asFrom _
    , byVal asTo _
    )
    Set fs_moveFolder = func_FsGeneralExecutor(Array(asFrom, asTo), "MoveFolder")
End Function

'***************************************************************************************************
'Function/Sub Name           : fs_readFile()
'Overview                    : Unicode�`���̃t�@�C����ǂ�Œ��g���擾����
'Detailed Description        : func_FsReadFile()�ɈϏ����ȉ��̐ݒ�œǍ���
'                               �t�@�C���̌`��         �FUnicode�`��
'Argument
'     asPath                 : ���͐�̃t���p�X
'Return Value
'     �t�@�C���̓��e
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/12/19         Y.Fujii                  First edition
'***************************************************************************************************
Private Function fs_readFile( _
    byVal asPath _
    )
    Set fs_readFile = func_FsReadFile(asPath, -1)
End Function

'***************************************************************************************************
'Function/Sub Name           : fs_wrapInQuotes()
'Overview                    : ���p���i"�F�_�u���N�H�[�e�[�V�����j�ň͂�
'Detailed Description        : �ΏۂɈ��p���i"�F�_�u���N�H�[�e�[�V�����j���܂ޏꍇ�̓G�X�P�[�v����
'Argument
'     asTgt                  : �Ώ�
'Return Value
'     �Ώۂ����p���i"�F�_�u���N�H�[�e�[�V�����j�ň͂񂾕�����
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2024/02/04         Y.Fujii                  First edition
'***************************************************************************************************
Private Function fs_wrapInQuotes( _
    byVal asTgt _
    )
    fs_wrapInQuotes = Chr(34) & Replace(asTgt, Chr(34), Chr(34)&Chr(34)) & Chr(34)
End Function

'***************************************************************************************************
'Function/Sub Name           : fs_writeFile()
'Overview                    : Unicode�`���Ńt�@�C���o�͂���
'Detailed Description        : func_FsWriteFile()�ɈϏ����ȉ��̐ݒ�ŏo�͂���
'                               �o�̓��[�h            �F�����̃t�@�C����V�����f�[�^�Œu��������
'                               �t�@�C�������݂��Ȃ��ꍇ�F�V�����t�@�C�����쐬����
'                               �t�@�C���̌`��         �FUnicode�`��
'Argument
'     asPath                 : �o�͐�̃t���p�X
'     asCont                 : �o�͂�����e
'Return Value
'     ���� True:���� / False:���s
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/12/19         Y.Fujii                  First edition
'***************************************************************************************************
Private Function fs_writeFile( _
    byVal asPath _
    , byVal asCont _
    )
    Set fs_writeFile = func_FsWriteFile(asPath, 2, True, -1, asCont)
End Function

'***************************************************************************************************
'Function/Sub Name           : fs_writeFileDefault()
'Overview                    : �V�X�e���̊���̌`���Ńt�@�C���o�͂���
'Detailed Description        : func_FsWriteFile()�ɈϏ����ȉ��̐ݒ�ŏo�͂���
'                               �o�̓��[�h            �F�����̃t�@�C����V�����f�[�^�Œu��������
'                               �t�@�C�������݂��Ȃ��ꍇ�F�V�����t�@�C�����쐬����
'                               �t�@�C���̌`��         �F�V�X�e���̊���
'Argument
'     asPath                 : �o�͐�̃t���p�X
'     asCont                 : �o�͂�����e
'Return Value
'     ���� True:���� / False:���s
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/12/27         Y.Fujii                  First edition
'***************************************************************************************************
Private Function fs_writeFileDefault( _
    byVal asPath _
    , byVal asCont _
    )
    Set fs_writeFileDefault = func_FsWriteFile(asPath, 2, True, -2, asCont)
End Function


'***************************************************************************************************
'Function/Sub Name           : fs_getAllFiles()
'Overview                    : �t�H���_�z���̃t�@�C���I�u�W�F�N�g���擾����
'Detailed Description        : �H����
'Argument
'     asPath                 : �t�@�C��/�t�H���_�̃p�X
'Return Value
'     File�I�u�W�F�N�g�����i�A�_�v�^�[�Ń��b�v�����j�̃I�u�W�F�N�g�̔z��
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/11/12         Y.Fujii                  First edition
'***************************************************************************************************
Private Function fs_getAllFiles( _
    byVal asPath _
    )
    fs_getAllFiles = func_FsGetAllFilesByFso(asPath)
'    fs_getAllFiles = func_FsGetAllFilesByShell(asPath)
'    fs_getAllFiles = func_FsGetAllFilesByDir(asPath)
End Function

'***************************************************************************************************
'Function/Sub Name           : func_FsGetAllFilesByFso()
'Overview                    : �t�H���_�z���̃t�@�C���I�u�W�F�N�g���擾����iFSO�Łj
'Detailed Description        : zip�t�@�C�����̌�����func_FsGetAllFilesByShell()�ɈϏ�����
'Argument
'     asPath                 : �t�@�C��/�t�H���_�̃p�X
'Return Value
'     File�I�u�W�F�N�g�����i�A�_�v�^�[�Ń��b�v�����j�̃I�u�W�F�N�g�̔z��
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/11/12         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_FsGetAllFilesByFso( _
    byVal asPath _
    )
    If new_Fso().FolderExists(asPath) Then
    '�t�H���_�̏ꍇ
        Dim oFolder : Set oFolder = new_FolderOf(asPath)
        Dim oEle, vRet()
        '�t�@�C���̎擾
        For Each oEle In oFolder.Files
            If StrComp(new_Fso().GetExtensionName(oEle.Path), "zip", vbTextCompare)=0 Then
            'zip�t�@�C���̏ꍇ�Afunc_FsGetAllFilesByShell()��zip���̃t�@�C�����X�g���擾����
                cf_pushA vRet, func_FsGetAllFilesByShell(oEle.Path)
            Else
            'zip�t�@�C���ȊO�̏ꍇ�A�t�@�C�������擾����
                cf_push vRet, new_AdptFileOf(oEle.Path)
            End If
        Next
        '�t�H���_�̎擾
        For Each oEle In oFolder.SubFolders
            cf_pushA vRet, func_FsGetAllFilesByFso(oEle)
        Next
        func_FsGetAllFilesByFso = vRet
    Else
    '�t�@�C���̏ꍇ
        func_FsGetAllFilesByFso = Array(new_AdptFileOf(asPath))
    End If

    Set oFolder = Nothing
End Function

'***************************************************************************************************
'Function/Sub Name           : func_FsGetAllFilesByShell()
'Overview                    : �t�H���_�z���̃t�@�C���I�u�W�F�N�g���擾����iShellApp�Łj
'Detailed Description        : zip�t�@�C�����̃t�@�C�����X�g���擾�ł���
'Argument
'     asPath                 : �t�@�C��/�t�H���_�̃p�X
'Return Value
'     File�I�u�W�F�N�g�����i�A�_�v�^�[�Ń��b�v�����j�̃I�u�W�F�N�g�̔z��
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/11/25         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_FsGetAllFilesByShell( _
    byVal asPath _
    )
    '�����^�C�v����
    Dim boFlg : boFlg = True 'AsFolder
    If new_Fso().FileExists(asPath) Then
        If StrComp(new_Fso().GetExtensionName(asPath), "zip", vbTextCompare)<>0 Then boFlg=False 'AsFile
    End If
    
    If boFlg Then
    '�t�H���_��zip�t�@�C���̏ꍇ
        Dim oFolder : Set oFolder = new_ShellApp().Namespace(asPath)
        Dim oItem, vRet()
        For Each oItem In oFolder.Items
        '�t�H���_���S�ẴA�C�e���ɂ���
            If oItem.IsFolder Then
            '�t�H���_�̏ꍇ
                cf_pushA vRet, func_FsGetAllFilesByShell(oItem.Path)
            Else
            '�t�@�C���̏ꍇ
                cf_push vRet, new_AdptFile(oItem)
            End If
        Next
        func_FsGetAllFilesByShell = vRet
        Set oItem = Nothing
    Else
    '��L�ȊO�̏ꍇ
        func_FsGetAllFilesByShell = Array(new_AdptFileOf(asPath))
    End If
End Function

'***************************************************************************************************
'Function/Sub Name           : func_FsGetAllFilesByDir()
'Overview                    : �t�H���_�z���̃t�@�C���I�u�W�F�N�g���擾����iDir�Łj
'Detailed Description        : zip�t�@�C�����̌�����func_FsGetAllFilesByShell()�ɈϏ�����
'Argument
'     asPath                 : �t�@�C��/�t�H���_�̃p�X
'Return Value
'     File�I�u�W�F�N�g�����i�A�_�v�^�[�Ń��b�v�����j�̃I�u�W�F�N�g�̔z��
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/11/25         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_FsGetAllFilesByDir( _
    byVal asPath _
    )
    Dim sDir : sDir = "dir /S /B /A-D " & fs_wrapInQuotes(asPath)
    Dim sTmpPath : sTmpPath = fw_getTempPath()
    fw_runShellSilently "cmd /U /C " & sDir & " > " & fs_wrapInQuotes(sTmpPath)
    Dim sLists : sLists = fs_readFile(sTmpPath)
    fs_deleteFile sTmpPath
    
    Dim vArrList : vArrList = Split(sLists, vbNewLine)
    Redim Preserve vArrList(Ubound(vArrList)-1)
    Dim sList, vRet()
    For Each sList In vArrList
        If StrComp(new_Fso().GetExtensionName(sList), "zip", vbTextCompare)=0 Then
        'zip�t�@�C���̏ꍇ�Afunc_FsGetAllFilesByShell()��zip���̃t�@�C�����X�g���擾����
            cf_pushA vRet, func_FsGetAllFilesByShell(sList)
        Else
        'zip�t�@�C���ȊO�̏ꍇ�A�t�@�C�������擾����
            cf_push vRet, new_AdptFileOf(sList)
        End If
    Next
    func_FsGetAllFilesByDir = vRet
End Function

'***************************************************************************************************
'Function/Sub Name           : func_FsGeneralExecutor()
'Overview                    : Fso�R�}���h�ėp���s�֐�
'Detailed Description        : �H����
'Argument
'     asPath                 : �p�X
'     asCmd                  : ���s�R�}���h
'Return Value
'     ���� True:���� / False:���s
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/12/30         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_FsGeneralExecutor( _
    byRef asPath _
    , byVal asCmd _
    )
    With new_Fso()
        On Error Resume Next
        Select Case asCmd
            Case "CopyFile":
                .CopyFile asPath(0), asPath(1)
            Case "CopyFolder":
                .CopyFolder asPath(0), asPath(1)
            Case "CreateFolder":
                .CreateFolder asPath(0)
            Case "DeleteFile":
                .DeleteFile asPath(0)
            Case "DeleteFolder":
                .DeleteFolder asPath(0)
            Case "MoveFile":
                .MoveFile asPath(0), asPath(1)
            Case "MoveFolder":
                .MoveFolder asPath(0), asPath(1)
            Case Else
                Err.Raise 9999, "libCom.vbs:func_FsGeneralExecutor()", "�s���Ȏ��s�R�}���h�F"&asCmd
        End Select
        Set func_FsGeneralExecutor = new_RetByState(True,False)
        On Error Goto 0
    End With
End Function

'***************************************************************************************************
'Function/Sub Name           : func_FsReadFile()
'Overview                    : �t�@�C����ǂ�Œ��g���擾����
'Detailed Description        : �H����
'Argument
'     asPath                 : ���͐�̃t���p�X
'     alFormat               : �t�@�C���̌`��
'                               -2�F�V�X�e���̊��� / -1�FUnicode / 0�FAscii
'Return Value
'     �t�@�C���̓��e
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/12/19         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_FsReadFile( _
    byVal asPath _
    , byVal alFormat _
    )
    Dim sRet : sRet = Empty
    On Error Resume Next
    With new_Ts(asPath, 1, False, alFormat)
        sRet = .ReadAll
        .Close
    End With
    Set func_FsReadFile = new_Ret(sRet)
    On Error Goto 0
End Function

'***************************************************************************************************
'Function/Sub Name           : func_FsWriteFile()
'Overview                    : �t�@�C���ɏo�͂���
'Detailed Description        : �H����
'Argument
'     asPath                 : �o�͐�̃t���p�X
'     alMode                 : �o�̓��[�h
'                               2�F�����̃t�@�C����V�����f�[�^�Œu�������� / 8�F�t�@�C���̍Ō�ɏ������݁j
'     aboCreate              : �t�@�C�������݂��Ȃ��ꍇ�ɐV�����t�@�C�����쐬�ł��邩�ǂ���������
'                               True�F�V�����t�@�C�����쐬���� / False�F�쐬���Ȃ�
'     alFormat               : �t�@�C���̌`��
'                               -2�F�V�X�e���̊��� / -1�FUnicode / 0�FAscii
'     asCont                 : �o�͂�����e
'Return Value
'     ���� True:���� / False:���s
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/12/19         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_FsWriteFile( _
    byVal asPath _
    , byVal alMode _
    , byVal aboCreate _
    , byVal alFormat _
    , byVal asCont _
    )
    On Error Resume Next
    With new_Ts(asPath, alMode, aboCreate, alFormat)
        .Write asCont
        .Close
    End With
    Set func_FsWriteFile=new_RetByState(True,False)
    On Error Goto 0
End Function





'###################################################################################################
'�G�N�Z���n
'###################################################################################################

'***************************************************************************************************
'Function/Sub Name           : sub_CM_ExcelSaveAs()
'Overview                    : �G�N�Z���t�@�C����ʖ��ŕۑ����ĕ���
'Detailed Description        : �H����
'Argument
'     aoWorkBook             : �G�N�Z���̃��[�N�u�b�N
'     asPath                 : �ۑ�����t�@�C���̃t���p�X
'     alFileformat           : XlFileFormat �񋓑́i�f�t�H���g��xlOpenXMLWorkbook 51 Excel�u�b�N�j
'Return Value
'     �Ȃ�
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/09/27         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub sub_CM_ExcelSaveAs( _
    byRef aoWorkBook _
    , byVal asPath _
    , byVal alFileformat _
    )
    If Not(IsNumeric(alFileformat)) Then
        alFileformat = 51                  'xlOpenXMLWorkbook 51 Excel�u�b�N
    End If
    Call aoWorkBook.SaveAs( _
                            asPath _
                            , alFileformat _
                            , , _
                            , False _
                            , False _
                            )
    Call aoWorkBook.Close(False)
End Sub

'***************************************************************************************************
'Function/Sub Name           : func_CM_ExcelOpenFile()
'Overview                    : �G�N�Z���t�@�C����ǂݎ���p�^�_�C�A���O�Ȃ��ŊJ��
'Detailed Description        : �H����
'Argument
'     aoExcel                : �G�N�Z��
'     asPath                 : �G�N�Z���t�@�C���̃t���p�X
'Return Value
'     �J�����G�N�Z���̃��[�N�u�b�N
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/09/27         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_ExcelOpenFile( _
    byRef aoExcel _
    , byVal asPath _
    )    
    Set func_CM_ExcelOpenFile = aoExcel.Workbooks.Open( _
                                                        asPath _
                                                        , 0 _
                                                        , True _
                                                        , , , _
                                                        , True _
                                                        )
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_ExcelGetTextFromAutoshape()
'Overview                    : �G�N�Z���̃I�[�g�V�F�C�v�̃e�L�X�g�����o��
'Detailed Description        : �G���[�͖�������
'Argument
'     aoAutoshape            : �I�[�g�V�F�C�v
'Return Value
'     �I�[�g�V�F�C�v�̃e�L�X�g
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/09/27         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_ExcelGetTextFromAutoshape( _
    byRef aoAutoshape _
    )
    func_CM_ExcelGetTextFromAutoshape = aoAutoshape.TextFrame.Characters.Text
End Function


'###################################################################################################
'�����񑀍�n
'###################################################################################################

'***************************************************************************************************
'Function/Sub Name           : func_CM_StrConvOnlyAlphabet()
'Overview                    : �p�������啶���^�������ɕϊ�����
'Detailed Description        : �H����
'Argument
'     asTarget               : �ϊ����镶����
'     alConversion           : ���s����ϊ��̎�� 1:UpperCase,2:LowerCase
'Return Value
'     �ϊ�����������
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/08/19         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_StrConvOnlyAlphabet( _
    byVal asTarget _
    , byVal alConversion _
    )
    Dim sChar, asTargetTmp
    
    '1���������肷��
    Dim asTargetNew : asTargetNew = asTarget
    Dim lPos : lPos = 1
    Do While Len(asTargetNew) >= lPos
        '�ϊ��Ώۂ�1�������擾
        sChar = Mid(asTargetNew, lPos, 1)
        
        If new_Char().whatType(sChar)<3 Then
        '�ϊ��Ώۂ��p���̏ꍇ�̂ݕϊ�����
            asTargetTmp = ""
            
            '�ϊ��Ώۂ̕����܂ł̕�����
            If lPos > 1 Then
                asTargetTmp = Mid(asTargetNew, 1, lPos-1)
            End If
            
            '�ϊ�����1����������
            sChar = func_CM_StrConv(sChar, alConversion)
            asTargetTmp = asTargetTmp & sChar
            
            '�ϊ��Ώۂ̕����ڍs�̕����������
            If lPos < Len(asTargetNew) Then
                asTargetTmp = asTargetTmp & Mid(asTargetNew, lPos+1, Len(asTargetNew)-lPos)
            End If
            
            '�ϊ���̕�������i�[
            asTargetNew = asTargetTmp
        End If
        
        '�J�E���g�A�b�v
        lPos = lPos+1
    Loop
    
    func_CM_StrConvOnlyAlphabet = asTargetNew
    
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_StrConv()
'Overview                    : ��������w��̂Ƃ���ϊ�����
'Detailed Description        : �H����
'Argument
'     asTarget               : �ϊ����镶����
'     alConversion           : ���s����ϊ��̎�ށi�����_��1,2�̂ݎ����j
'                                 1:�������啶���ɕϊ�
'                                 2:��������������ɕϊ�
'                                 3:��������̂��ׂĂ̒P��̍ŏ��̕�����啶���ɕϊ�
'                                 4:��������̋��� (1 �o�C�g) ���������C�h (2 �o�C�g) �����ɕϊ�
'                                 8:��������̃��C�h (2 �o�C�g) ���������� (1 �o�C�g) �����ɕϊ�
'                                16:��������̂Ђ炪�ȕ������J�^�J�i�����ɕϊ�
'                                32:��������̃J�^�J�i�������Ђ炪�ȕ����ɕϊ�
'Return Value
'     �ϊ�����������
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/08/19         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_StrConv( _
    byVal asTarget _
    , byVal alalConversion _
    )
    Dim sChar, asTargetTmp
    func_CM_StrConv = asTarget
    Select Case alalConversion
        Case 1:
            func_CM_StrConv = UCase(asTarget)
        Case 2:
            func_CM_StrConv = LCase(asTarget)
    End Select
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_StrLen()
'Overview                    : �S�p��2�����A���p��1�����Ƃ��ĕ�������Ԃ�
'Detailed Description        : �H����
'Argument
'     asTarget               : ������
'Return Value
'     ������
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/08/19         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_StrLen( _
    byVal asTarget _
    )
    '1���������肷��
    Dim sChar
    Dim lLength : lLength = 0
    Dim lPos : lPos = 1
    Do While Len(asTarget) >= lPos
        '1�������擾
        sChar = Mid(asTarget, lPos, 1)
        
        If (Asc(sChar) And &HFF00) <> 0 Then
            lLength = lLength+2
        Else
            lLength = lLength+1
        End If
        
        '�J�E���g�A�b�v
        lPos = lPos+1
    Loop
    
    func_CM_StrLen = lLength
End Function


'###################################################################################################
'�z��n
'###################################################################################################

'***************************************************************************************************
'Function/Sub Name           : func_CM_ArrayGetDimensionNumber()
'Overview                    : �z��̎����������߂�
'Detailed Description        : �H����
'Argument
'     avArray                : �z��
'Return Value
'     �z��̎�����
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/11/19         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_ArrayGetDimensionNumber( _
    byRef avArray _ 
    )
   If Not IsArray(avArray) Then Exit Function
   On Error Resume Next
   Dim lNum : lNum = 0
   Dim lTemp
   Do
       lNum = lNum + 1
       lTemp = UBound(avArray, lNum)
   Loop Until Err.Number <> 0
   Err.Clear
   func_CM_ArrayGetDimensionNumber = lNum - 1
End Function


'###################################################################################################
'���̑�
'###################################################################################################

'***************************************************************************************************
'Function/Sub Name           : func_CM_FillInTheCharacters()
'Overview                    : �����𖄂߂�
'Detailed Description        : �Ώە����̕s�������w�肵���A���C�����g�Ŏw�肵��������1�����ڂŖ��߂�
'                              �Ώە����ɕs�������Ȃ��ꍇ�́A�w�肵���������Ő؂���
'Argument
'     asTarget               : �Ώە�����
'     alWordCount            : ������
'     asToFillCharacter      : ���߂镶��
'     aboIsCutOut            : �������Ő؂���iTrue�F����/False�F���Ȃ��j
'     aboIsRightAlignment    : �A���C�����g�iTrue�F�E��/False�F���񂹁j
'Return Value
'     ���߂�������
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/01/04         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FillInTheCharacters( _
    byVal asTarget _
    , byVal alWordCount _
    , byVal asToFillCharacter _
    , byVal aboIsCutOut _
    , byVal aboIsRightAlignment _
    )
    
    '�؂���Ȃ��őΏە����񂪕��������傫���ꍇ�͏����𔲂���
    Dim lTargetLen : lTargetLen = Len(asTarget)
    If Not(aboIsCutOut) And lTargetLen>=alWordCount Then
        func_CM_FillInTheCharacters = asTarget
        Exit Function
    End If
    
    '���߂镶����̍쐬
    Dim sFillStrings : sFillStrings = ""
    If alWordCount-lTargetLen > 0 Then
        sFillStrings = String(alWordCount-lTargetLen , asToFillCharacter)
    End If
    
    Dim sResult
    '�A���C�����g�w��ɂ���ĕ�����𖄂߂�
    If aboIsRightAlignment Then
        sResult = sFillStrings & asTarget
    Else
        sResult = asTarget & sFillStrings
    End If
    
    '�؂���Ȃ��̏ꍇ�͏����𔲂���
    If Not(aboIsCutOut) Then
        func_CM_FillInTheCharacters = sResult
        Exit Function
    End If
    
    '�A���C�����g�w��ɂ���ĕ������؂���
    If aboIsRightAlignment Then
        sResult = Right(sResult, alWordCount)
    Else
        sResult = Left(sResult, alWordCount)
    End If
    func_CM_FillInTheCharacters = sResult
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_FormatDecimalNumber()
'Overview                    : ���������_�^�𐮌`����
'Detailed Description        : �H����
'Argument
'     adbNumber              : ���������_�^�̐��l
'     alDecimalPlaces        : �����̌���
'Return Value
'     ���`�������������_�^
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/08/26         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FormatDecimalNumber( _
    byVal adbNumber _
    , byVal alDecimalPlaces _
    )
    func_CM_FormatDecimalNumber = Fix(adbNumber) & "." _
                             & func_CM_FillInTheCharacters( _
                                                          Abs(Fix( (adbNumber - Fix(adbNumber))*10^alDecimalPlaces )) _
                                                          , alDecimalPlaces _
                                                          , "0" _
                                                          , False _
                                                          , True _
                                                          )
End Function


'###################################################################################################
'���[�e�B���e�B�n
'###################################################################################################

'***************************************************************************************************
'Function/Sub Name           : func_CM_UtilSortDefaultFunc()
'Overview                    : �v�f�̔�r���ʂ�Ԃ�
'Detailed Description        : �\�[�g�֐��Q�Ŏg���f�t�H���g�̊֐�
'Argument
'     aoCurrentValue         : �z��̗v�f
'     aoNextValue            : ���̔z��̗v�f
'Return Value
'     �v�f�̔�r����
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/10/15         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_UtilSortDefaultFunc( _
    byRef aoCurrentValue _
    , byRef aoNextValue _
    )
    func_CM_UtilSortDefaultFunc = aoCurrentValue>aoNextValue
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_UtilJoin()
'Overview                    : Join�֐�
'Detailed Description        : vbscript��Join�֐��Ɠ����̋@�\
'Argument
'     avArr                  : �z��
'     asDel                  : ��؂蕶��
'Return Value
'     �z��̊e�v�f��A������������
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/10/17         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_UtilJoin( _
    byRef avArr _
    , byVal asDel _
    )
    func_CM_UtilJoin = Join(avArr, asDel)
End Function
