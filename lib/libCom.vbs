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
'Function/Sub Name           : cf_isNumeric()
'Overview                    : ���l�����肷��
'Detailed Description        : �H����
'Argument
'     avTgt                  : �Ώ�
'Return Value
'     ���� True:���l / False:���l�łȂ�
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2024/01/03         Y.Fujii                  First edition
'***************************************************************************************************
Private Function cf_isNumeric( _
    byRef avTgt _
    )
    If IsEmpty(avTgt) Or IsNull(avTgt) Or IsObject(avTgt) Or IsArray(avTgt) Then
    'Empty,Null,Object,Array�̏ꍇ��False
        cf_isNumeric=False
        Exit Function
    End If
    If VarType(avTgt)=vbInteger Or VarType(avTgt)=vbLong Or VarType(avTgt)=vbSingle Or VarType(avTgt)=vbDouble Then
    'Integer,Long,Single,Double�̏ꍇ��True
        cf_isNumeric=True
        Exit Function
    End If
    cf_isNumeric=False
    If VarType(avTgt)=vbString Then
    'String�̏ꍇ��IsNumeric�֐��̖߂�l��Ԃ�
        cf_isNumeric=IsNumeric(avTgt)
    End If
End Function

'***************************************************************************************************
'Function/Sub Name           : cf_isSame()
'Overview                    : ���ꂩ���肷��
'Detailed Description        : �H����
'Argument
'     aoA                    : ��r�Ώ�
'     aoB                    : ��r�Ώ�
'Return Value
'     ���� True:���� / False:����łȂ�
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/10/15         Y.Fujii                  First edition
'***************************************************************************************************
Private Function cf_isSame( _
    byRef aoA _
    , byRef aoB _
    )
    Dim boFlg : boFlg = False
    If IsObject(aoA) And IsObject(aoB) Then
        If aoA Is aoB Then boFlg = True
    ElseIf Not IsObject(aoA) And Not IsObject(aoB) Then
        If VarType(aoA) = vbString And VarType(aoB) = vbString Then
            If Strcomp(aoA, aoB, vbBinaryCompare)=0 Then boFlg = True
        Else
            If aoA = aoB Then boFlg = True
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
'2022/11/23         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub cf_push( _
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
'Function/Sub Name           : cf_pushMulti()
'Overview                    : �z��ɕ����̗v�f��ǉ�����
'Detailed Description        : �H����
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
'***************************************************************************************************
Private Sub cf_pushMulti( _
    byRef avArr _ 
    , byRef avAdd _ 
    )
    On Error Resume Next
    Dim lUbAdd,lIdx : lUbAdd = Ubound(avAdd)
    If Err.Number=0 Then
    '�ǉ�����z��iavAdd�j���v�f�����ꍇ
        Dim lUb : lUb = Ubound(avArr)
        If Err.Number=0 Then 
        '�z��iavArr�j���v�f�����ꍇ
            Redim Preserve avArr(lUb+lUbAdd+1)
            For lIdx=0 To lUbAdd
                cf_bind avArr(lUb+1+lIdx), avAdd(lIdx)
            Next
        Else
        '�z��iavArr�j���v�f�������Ȃ��ꍇ
            Redim avArr(Ubound(avAdd))
            For lIdx=0 To Ubound(avArr)
                cf_bind avArr(lIdx), avAdd(lIdx)
            Next
        End If
    Elseif Not IsArray(avAdd) Then
    '�ǉ�����z��iavAdd�j���v�f���������z��łȂ��ꍇ
        cf_push avArr, avAdd
    End If
    On Error Goto 0
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
    Const Cs_TOPIC = "log"
    
    '���s�O�̏o�ŁiPublish�j ����
    If cf_isAvailableObject(aoBroker) Then
        aoBroker.publish Cs_TOPIC, Array(5 ,asSubName ,"Start")
        aoBroker.publish Cs_TOPIC, Array(9 ,asSubName ,cf_toString(aoArg))
    End If
    
    '�֐��̎��s
    Dim oRet : Set oRet = fw_tryCatch(GetRef(asSubName), aoArg, Empty, Empty)
    
    '���s��̏o�ŁiPublish�j ����
    If cf_isAvailableObject(aoBroker) Then
        If oRet.isErr() Then
        '�G���[
            aoBroker.publish Cs_TOPIC, Array(1, asSubName, cf_toString(oRet.getErr()))
        End If
        aoBroker.publish Cs_TOPIC, Array(5, asSubName, "End")
        aoBroker.publish Cs_TOPIC, Array(9, asSubName, cf_toString(aoArg))
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
        sParentFolderPath = new_Fso().BuildPath(sParentFolderPath ,asParentFolderName)
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

    With aoWriter
        .WriteLine(new_ArrWith(Array(new_Now(), Join(vIps,","), new_Network().ComputerName)).Concat(avParams).join(vbTab))
    End With

    Set oEle = Nothing
End Sub

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
'Function/Sub Name           : new_DicWith()
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
Private Function new_DicWith( _
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
    
    Set new_DicWith = oDict
    Set oDict = Nothing
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
    Set new_Ret = (New clsCmReturnValue).setValue(avRet)
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
    Set new_Reader = (New clsCmBufferedReader).setTextStream(aoTextStream)
End Function

'***************************************************************************************************
'Function/Sub Name           : new_ReaderFrom()
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
Private Function new_ReaderFrom( _
    byVal asPath _
    )
    Set new_ReaderFrom = (New clsCmBufferedReader).setTextStream(new_Ts(asPath, 1, False, -2))
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
    Set new_Writer = (New clsCmBufferedWriter).setTextStream(aoTextStream)
End Function

'***************************************************************************************************
'Function/Sub Name           : new_WriterTo()
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
Private Function new_WriterTo( _
    byVal asPath _
    , byVal alIomode _
    , byVal aboCreate _
    , byVal alFileFormat _
    )
    Set new_WriterTo = (New clsCmBufferedWriter).setTextStream(new_Ts(asPath, alIomode, aboCreate, alFileFormat))
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
    Set new_Now = (New clsCmCalendar).getNow()
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
    Set new_CalAt = (New clsCmCalendar).setDateTime(avDateTime)
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
    Set new_Broker = (New clsCmBroker)
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
    Set new_Arr = (New clsCmArray)
End Function

'***************************************************************************************************
'Function/Sub Name           : new_ArrWith()
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
Private Function new_ArrWith( _
    byRef avArr _
    )
    Dim oArr : Set oArr = new_Arr()
    oArr.PushMulti avArr
    Set new_ArrWith = oArr
    Set oArr = Nothing
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
    Set new_ArrSplit = new_ArrWith(Split(asTarget, asDelimiter, -1, vbBinaryCompare))
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
    Dim oHtml : Set oHtml = New clsCmHtmlGenerator
    oHtml.element = asElement
    Set new_HtmlOf = oHtml
    Set oHtml = Nothing
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
    Dim oCss : Set oCss = New clsCmCssGenerator
    oCss.selector = asSelector
    Set new_CssOf = oCss
    Set oCss = Nothing
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
    Set new_Char = (New clsCmCharacterType)
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
    With new_Char()
        Dim vCharList : vCharList = .getCharList(.typeHalfWidthAlphabetUppercase + .typeHalfWidthNumbers)
    End With
    cf_push vCharList, "_"
    Dim sFuncName : sFuncName = "anonymous_" & util_randStr(vCharList, 10)
    
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
    Dim lRet
    If al1 < al2 Then lRet = al1 Else lRet = al2
    math_min = lRet
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
    Dim lRet
    If al1 > al2 Then lRet = al1 Else lRet = al2
    math_max = lRet
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
'Function/Sub Name           : func_MathRound()
'Overview                    : ���l���ۂ߂�
'Detailed Description        : �����𖳎����Đ�Βl���ۂ߂�
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
         cf_push oRet, new_DicWith(Array("Caption", oAdapter.Caption, "Ip", new_DicWith(Array("V4", oIpv4, "V6", oIpv6))))
    Next
    util_getIpAddress = oRet
    
    Set oAddress = Nothing
    Set oAdapter = Nothing
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
    Set fs_copyFile = func_FsGeneralExecutor(False, False, Array(asFrom, asTo), "CopyFile")
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
    Set fs_copyFolder = func_FsGeneralExecutor(True, False, Array(asFrom, asTo), "CopyFolder")
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
    Set fs_createFolder = func_FsGeneralExecutor(True, True, Array(asPath), "CreateFolder")
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
    Set fs_deleteFile = func_FsGeneralExecutor(False, False, Array(asPath), "DeleteFile")
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
    Set fs_deleteFolder = func_FsGeneralExecutor(True, False, Array(asPath), "DeleteFolder")
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
    Set fs_moveFile = func_FsGeneralExecutor(False, False, Array(asFrom, asTo), "MoveFile")
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
    Set fs_moveFolder = func_FsGeneralExecutor(True, False, Array(asFrom, asTo), "MoveFolder")
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
    fs_readFile = func_FsReadFile(asPath, -1)
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
    fs_writeFile = func_FsWriteFile(asPath, 2, True, -1, asCont)
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
    fs_writeFileDefault = func_FsWriteFile(asPath, 2, True, -2, asCont)
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
                cf_pushMulti vRet, func_FsGetAllFilesByShell(oEle.Path)
            Else
            'zip�t�@�C���ȊO�̏ꍇ�A�t�@�C�������擾����
                cf_push vRet, new_AdptFileOf(oEle.Path)
            End If
        Next
        '�t�H���_�̎擾
        For Each oEle In oFolder.SubFolders
            cf_pushMulti vRet, func_FsGetAllFilesByFso(oEle)
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
                cf_pushMulti vRet, func_FsGetAllFilesByShell(oItem.Path)
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
    Dim sDir : sDir = "dir /S /B /A-D " & Chr(34) & asPath & Chr(34)
    Dim sTmpPath : sTmpPath = fw_getTempPath()
    new_Shell().run "cmd /U /C " & sDir & " > " & Chr(34) & sTmpPath & Chr(34), 0, True
    Dim sLists : sLists = fs_readFile(sTmpPath)
    fs_deleteFile sTmpPath
    
    Dim vArrList : vArrList = Split(sLists, vbNewLine)
    Redim Preserve vArrList(Ubound(vArrList)-1)
    Dim sList, vRet()
    For Each sList In vArrList
        If StrComp(new_Fso().GetExtensionName(sList), "zip", vbTextCompare)=0 Then
        'zip�t�@�C���̏ꍇ�Afunc_FsGetAllFilesByShell()��zip���̃t�@�C�����X�g���擾����
            cf_pushMulti vRet, func_FsGetAllFilesByShell(sList)
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
'     aboIsFolder            : True:�t�H���_�L���̔��� / False:�t�@�C���L���̔���
'     aboFlg                 : ����Ɏg�p����t���O
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
    byVal asIsFolder _
    , byVal aboFlg _
    , byRef asPath _
    , byVal asCmd _
    )
    Set func_FsGeneralExecutor=new_Ret(False)
    With new_Fso()
        If asIsFolder Then
            If .FolderExists(asPath(0))=aboFlg Then Exit Function
        Else
            If .FileExists(asPath(0))=aboFlg Then Exit Function
        End If
    
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
'        Eval("new_Fso()." & asCmd & "(" & Chr(34) & asPath & Chr(34) & ")")
'        If Err.Number=0 Then func_FsGeneralExecutor=True
        Dim boRet : If Err.Number=0 Then boRet=True Else boRet=False
        Set func_FsGeneralExecutor = new_Ret(boRet)
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
    func_FsReadFile = Empty
    On Error Resume Next
    With new_Ts(asPath, 1, False, alFormat)
        func_FsReadFile = .ReadAll
        .Close
    End With
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
    func_FsWriteFile = True
    On Error Resume Next
    With new_Ts(asPath, alMode, aboCreate, alFormat)
        .Write asCont
        .Close
    End With
    If Err.Number Then func_FsWriteFile = False
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
'�t�@�C������n
'###################################################################################################

'***************************************************************************************************
'Function/Sub Name           : func_CM_FsGetFiles()
'Overview                    : �w�肵���t�H���_�ȉ���Files�R���N�V�������擾����
'Detailed Description        : FileSystemObject��Folder�I�u�W�F�N�g��Files�R���N�V�����Ɠ���
'Argument
'     asPath                 : �p�X
'Return Value
'     Files�R���N�V����
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/10/23         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FsGetFiles( _
    byVal asPath _
    ) 
    Set func_CM_FsGetFiles = new_Fso().GetFolder(asPath).Files
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_FsGetFolders()
'Overview                    : �w�肵���t�H���_�ȉ���Folders�R���N�V�������擾����
'Detailed Description        : FileSystemObject��Folder�I�u�W�F�N�g��SubFolders�R���N�V�����Ɠ���
'Argument
'     asPath                 : �p�X
'Return Value
'     Folders�R���N�V����
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/10/23         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FsGetFolders( _
    byVal asPath _
    ) 
    Set func_CM_FsGetFolders = new_Fso().GetFolder(asPath).SubFolders
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_FsIsSame()
'Overview                    : �w�肵���p�X�������t�@�C��/�t�H���_����������
'Detailed Description        : �H����
'Argument
'     asPathA                : �t�@�C��/�t�H���_�̃t���p�X
'     asPathB                : �t�@�C��/�t�H���_�̃t���p�X
'Return Value
'     ���� True:���� / False:����łȂ�
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/08/19         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FsIsSame( _
    byVal asPathA _
    , byVal asPathB _
    )
    func_CM_FsIsSame = (func_CM_FsGetFsObject(asPathA) Is func_CM_FsGetFsObject(asPathB))
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
'�`�F�b�N�n
'###################################################################################################

''***************************************************************************************************
''Function/Sub Name           : func_CM_ValidationlIsWithinTheRangeOf()
''Overview                    : ���l�^�͈͓̔��ɂ��邩��������
''Detailed Description        : �H����
''Argument
''     avNumber               : ���l
''     alType                 : �ϐ��̌^
''                                1:�����^�iInteger�j
''                                2:�������^�iLong�j
''                                3:�o�C�g�^�iByte�j
''                                4:�P���x���������_�^�iSingle�j
''                                5:�{���x���������_�^�iDouble�j
''                                6:�ʉ݌^�iCurrency�j
''Return Value
''     ���`�������������_�^
''---------------------------------------------------------------------------------------------------
''Histroy
''Date               Name                     Reason for Changes
''----------         ----------------------   -------------------------------------------------------
''2023/08/26         Y.Fujii                  First edition
''***************************************************************************************************
'Private Function func_CM_ValidationlIsWithinTheRangeOf( _
'    byVal avNumber _
'    , byVal alType _
'    )
'    Dim vMin,vMax
'    Select Case alType
'        Case 1:                   '�����^�iInteger�j
'            vMin = -1 * 2^15
'            vMax = 2^15 - 1
'        Case 2:                   '�������^�iLong�j
'            vMin = -1 * 2^31
'            vMax = 2^31 - 1
'        Case 3:                   '�o�C�g�^�iByte�j
'            vMin = 0
'            vMax = 2^8 - 1
'        Case 4:                   '�P���x���������_�^�iSingle�j
'            vMin = -3.402823E38
'            vMax = 3.402823E38
'        Case 5:                   '�{���x���������_�^�iDouble�j
'            vMin = -1.79769313486231E308
'            vMax = 1.79769313486231E308
'        Case 6:                   '�ʉ݌^�iCurrency�j
'            vMin = -1 * 2^59 / 1000
'            vMax = ( 2^59 - 1 ) / 1000
'    End Select
'    
'    func_CM_ValidationlIsWithinTheRangeOf = False
'    If vMin<=avNumber And avNumber<=vMax Then
'        func_CM_ValidationlIsWithinTheRangeOf = True
'    End If
'End Function


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
'Function/Sub Name           : func_CM_UtilSortBubble()
'Overview                    : �o�u���\�[�g
'Detailed Description        : �v�Z�񐔂�O(N^2)
'                              �z��iavArr�j�������Ȕz��̏ꍇ�͔z��iavArr�j�����̂܂ܕԂ�
'                              �����̊֐��̈����͈ȉ��̂Ƃ���
'                                currentValue :�z��̗v�f
'                                nextValue    :���̔z��̗v�f
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
Private Function func_CM_UtilSortBubble( _
    byRef avArr _
    , byRef aoFunc _
    , byVal aboFlg _
    )
    func_CM_UtilSortBubble = avArr
    If Not new_Arr().hasElement(avArr) Then Exit Function
    If Ubound(avArr)=0 Then Exit Function
    
    Dim lEnd, lPos
    lEnd = Ubound(avArr)
    Do While lEnd>0
        For lPos=0 To lEnd-1
            If aoFunc(avArr(lPos), avArr(lPos+1))=aboFlg Then
            'lPos�Ԗڂ̗v�f��(lPos+1)�Ԗڂ̗v�f�����ւ���
                cf_swap avArr(lPos), avArr(lPos+1)
            End If
        Next
        lEnd = lEnd-1
    Loop
    func_CM_UtilSortBubble = avArr
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_UtilSortQuick()
'Overview                    : �N�C�b�N�\�[�g
'Detailed Description        : �v�Z�񐔂͕���O(N*logN)�A�ň���O(N^2)
'                              �z��iavArr�j�������Ȕz��̏ꍇ�͔z��iavArr�j�����̂܂ܕԂ�
'                              �����̊֐��̈����͈ȉ��̂Ƃ���
'                                currentValue :�z��̗v�f
'                                nextValue    :���̔z��̗v�f
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
Private Function func_CM_UtilSortQuick( _
    byRef avArr _
    , byRef aoFunc _
    , byVal aboFlg _
    )
    func_CM_UtilSortQuick = avArr
    If Not new_Arr().hasElement(avArr) Then Exit Function
    If Ubound(avArr)=0 Then Exit Function
    
    '0�Ԗڂ̗v�f���s�{�b�g�Ɍ��߂�
    Dim oPivot : Call cf_bind(oPivot, avArr(0))
    
    '�s�{�b�g�Ɨv�f���֐��Ŕ��肵������@�ɍ��v����O���[�v��Right�A�����łȂ��O���[�v��Left�Ƃ���
    Dim lPos, vRight, vLeft
    For lPos=1 To Ubound(avArr)
        If aoFunc(avArr(lPos), oPivot)=aboFlg Then
            cf_push vRight, avArr(lPos)
        Else
            cf_push vLeft, avArr(lPos)
        End If
    Next
    
    '��q�ŕ�����Right�ALeft�̃O���[�v���ƂɍċA��������
    vLeft = func_CM_UtilSortQuick(vLeft, aoFunc, aboFlg)
    vRight = func_CM_UtilSortQuick(vRight, aoFunc, aboFlg)
    
    'Left�Ƀs�{�b�g�{Right����������
    cf_push vLeft, oPivot
    If new_Arr().hasElement(vRight) Then
        For lPos=0 To Ubound(vRight)
            cf_push vLeft, vRight(lPos)
        Next
    End If
    
    func_CM_UtilSortQuick = vLeft
    Set oPivot = Nothing
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_UtilSortMerge()
'Overview                    : �}�[�W�\�[�g
'Detailed Description        : �v�Z�񐔂�O(N*logN)
'                              �z��iavArr�j�������Ȕz��̏ꍇ�͔z��iavArr�j�����̂܂ܕԂ�
'                              �}�[�W������func_CM_UtilSortMergeMerge()�ɈϏ�����
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
Private Function func_CM_UtilSortMerge( _
    byRef avArr _
    , byRef aoFunc _
    , byVal aboFlg _
    )
    func_CM_UtilSortMerge = avArr
    If Not new_Arr().hasElement(avArr) Then Exit Function
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
    vFirst = func_CM_UtilSortMerge(vFirst, aoFunc, aboFlg)
    vSecond = func_CM_UtilSortMerge(vSecond, aoFunc, aboFlg)
    
    '�}�[�W�����Ȃ����ʂɖ߂�
    func_CM_UtilSortMerge = func_CM_UtilSortMergeMerge(vFirst, vSecond, aoFunc, aboFlg)
    
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_UtilSortMergeMerge()
'Overview                    : �}�[�W�\�[�g�̃}�[�W����
'Detailed Description        : func_CM_UtilSortMerge()����Ăяo��
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
Private Function func_CM_UtilSortMergeMerge( _
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
        If aoFunc(avFirst(lPosF), avSecond(lPosS))=aboFlg Then
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
    func_CM_UtilSortMergeMerge = vRet
    
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_UtilSortHeap()
'Overview                    : �q�[�v�\�[�g
'Detailed Description        : �v�Z�񐔂�O(N*logN)
'                              �z��iavArr�j�������Ȕz��̏ꍇ�͔z��iavArr�j�����̂܂ܕԂ�
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
'2023/09/21         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_UtilSortHeap( _
    byRef avArr _
    , byRef aoFunc _
    , byVal aboFlg _
    )
    func_CM_UtilSortHeap = avArr
    If Not new_Arr().hasElement(avArr) Then Exit Function
    If Ubound(avArr)=0 Then Exit Function
    
    '�q�[�v�̍쐬
    Dim lLb, lUb, lSize, lParent
    lLb = Lbound(avArr) : lUb = Ubound(avArr)
    lSize = lUb - lLb + 1
    '�q�����ŉ����̃m�[�h�����ʂɌ����ď��ԂɃm�[�h�P�ʂ̏������s��
    For lParent=lSize\2-1 To lLb Step -1
        sub_CM_UtilSortHeapPerNodeProc avArr, lSize, lParent, aoFunc, aboFlg
    Next
    
    '�q�[�v�̐擪�i�ő�/�ŏ��l�j�����ԂɎ��o��
    Do While lSize>0
        '�q�[�v�̐擪�Ɩ��������ւ���
        cf_swap avArr(lLb), avArr(lSize-1)
        '�q�[�v�T�C�Y���P���炵�čč쐬
        lSize = lSize - 1
        sub_CM_UtilSortHeapPerNodeProc avArr, lSize, 0, aoFunc, aboFlg
    Loop
    
    '�\�[�g�ς̔z���Ԃ�
    func_CM_UtilSortHeap = avArr
    
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_UtilSortHeapPerNodeProc()
'Overview                    : �q�[�v�\�[�g�̃m�[�h�P�ʂ̏���
'Detailed Description        : func_CM_UtilSortHeap()����Ăяo��
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
Private Sub sub_CM_UtilSortHeapPerNodeProc( _
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
        If aoFunc(avArr(lRight), avArr(alParent))=aboFlg Then
        '�e�ƉE���̎q�̗v�f���֐��Ŕ��肵������@�ɍ��v����ꍇ�͓���ւ���
            lToSwap = lRight
        End If
    End If
    
    If lLeft<alSize Then
    '�����̎q������ꍇ
        If aoFunc(avArr(lLeft), avArr(lToSwap))=aboFlg Then
        '�e�ƉE���̎q�̏��҂ƍ����̎q�̗v�f���֐��Ŕ��肵������@�ɍ��v����ꍇ�͓���ւ���
            lToSwap = lLeft
        End If
    End If
    
    If lToSwap<>alParent Then
        '�e�Ǝq�̗v�f�����ւ���
        cf_swap avArr(alParent), avArr(lToSwap)
        '����ւ����q�̗v�f�ȉ��̃m�[�h���ď�������
        sub_CM_UtilSortHeapPerNodeProc avArr, alSize, lToSwap, aoFunc, aboFlg
    End If
    
End Sub

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
'Function/Sub Name           : func_CM_UtilIsTextStream()
'Overview                    : �I�u�W�F�N�g��TextStream�����肷��
'Detailed Description        : �H����
'Argument
'     aoObj                  : �I�u�W�F�N�g
'Return Value
'     ���� True:TextStream�ł��� / False:TextStream�łȂ�
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/10/15         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_UtilIsTextStream( _
    byRef aoObj _
    )
    Dim boFlg : boFlg = False
    If cf_isSame(Vartype(aoObj),vbObject) And cf_isSame(Typename(aoObj),"TextStream") Then boFlg = True
    func_CM_UtilIsTextStream = boFlg
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
