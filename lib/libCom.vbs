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
'Function/Sub Name           : cf_push()
'Overview                    : �z��ɗv�f��ǉ�����
'Detailed Description        : �H����
'Argument
'     avArr                  : �z��
'     aoEle                  : �ǉ�����v�f
'Return Value
'     �z��̎�����
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/11/23         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub cf_push( _
    byRef avArr _ 
    , byRef aoEle _ 
    )
    If new_Arr().hasElement(avArr) Then
'    If func_CM_ArrayIsAvailable(avArr) Then
        Redim Preserve avArr(Ubound(avArr)+1)
    Else
        Redim avArr(0)
    End If
    cf_bind avArr(Ubound(avArr)), aoEle

'    cf_tryCatch Getref("func_CM_ArrayAddElement"), avArr, Getref("func_CM_ArrayInitialize"), Empty
'    cf_bind avArr(Ubound(avArr)), aoEle

End Sub

'***************************************************************************************************
'Function/Sub Name           : cf_tryCatch()
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
Private Function cf_tryCatch( _
    byRef aoTry _
    , byRef aoArgs _
    , byRef aoCatch _
    , byRef aoFinary _
    )
    Dim oRet, oErr, boFlg
    Set oErr = Nothing : boFlg = True
    
    'try�u���b�N�̏���
    On Error Resume Next
    cf_bind oRet, aoTry(aoArgs)
    If Err.Number<>0 Then
        boFlg = False
        Set oErr = func_CM_UtilStoringErr()
    End If
    On Error GoTo 0

    'catch�u���b�N�̏���
    If Not boFlg And func_CM_UtilIsAvailableObject(aoCatch) Then
        cf_bind oRet, aoCatch(aoArgs, oErr)
    End If
    
    'finary�u���b�N�̏���
    If func_CM_UtilIsAvailableObject(aoFinary) Then
        cf_bind oRet, aoFinary(aoArgs, oRet, oErr)
    End If
    
    '���ʂ�ԋp
    Set cf_tryCatch = new_DicWith(Array("Result", boFlg, "Return", oRet, "Err", oErr))
    Set oRet = Nothing
    Set oErr = Nothing
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
    Set new_ReaderFrom = (New clsCmBufferedReader).setTextStream(func_CM_FsOpenTextFile(asPath, 1, False, -2))
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
    Set new_WriterTo = (New clsCmBufferedWriter).setTextStream(func_CM_FsOpenTextFile(asPath, alIomode, aboCreate, alFileFormat))
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
    Dim sFuncName : sFuncName = "anonymous_" & func_CM_UtilGenerateRandomString(10, 5, Array("_"))
    
    Dim sPattern, oRegExp, sArgStr, sProcStr
    '��������֐��̃\�[�X�R�[�h�̗l�����u1.�ʏ�v�̏ꍇ
    sPattern = "function\s*\((.*)\)\s*{(.*)}"
    Set oRegExp = new_Re(sPattern, "igm")
    If oRegExp.Test(sSoruceCode) Then
        sArgStr = oRegExp.Replace(sSoruceCode, "$1")
        sProcStr = oRegExp.Replace(sSoruceCode, "$2")
        
        'return�傪����Ί֐����ŏ���������
        sProcStr = func_FuncRewriteReturnPhrase(sFuncName, False, func_FuncAnalyze(sProcStr) )
        
        '�֐��̐���
        Set new_Func = func_FuncGenerate(sFuncName, sArgStr, sProcStr)
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
        sProcStr = func_FuncRewriteReturnPhrase(sFuncName, True, func_FuncAnalyze(sProcStr) )
        
        '�֐��̐���
        Set new_Func = func_FuncGenerate(sFuncName, sArgStr, sProcStr)
    End If
    Set oRegExp = Nothing
End Function

'***************************************************************************************************
'Function/Sub Name           : func_FuncAnalyze()
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
Private Function func_FuncAnalyze( _
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
    
    func_FuncAnalyze = oCode.Items()
    Set oCode = Nothing
End Function

'***************************************************************************************************
'Function/Sub Name           : func_FuncRewriteReturnPhrase()
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
Private Function func_FuncRewriteReturnPhrase( _
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
            func_FuncRewriteReturnPhrase = new_Re(sPtnRet, "ig").Replace(sCode, "$1 cf_bind " & asFuncName & ", ($2)")
        Else
        'return�傪�Ȃ��ꍇ
            func_FuncRewriteReturnPhrase = "cf_bind " & asFuncName & ", (" & sCode & ")"
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
    
    func_FuncRewriteReturnPhrase = Join(avCode, ":")
    
End Function

'***************************************************************************************************
'Function/Sub Name           : func_FuncGenerate()
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
Private Function func_FuncGenerate( _
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
    Set func_FuncGenerate = Getref(asFuncName)
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
'Function/Sub Name           : func_CM_FsDeleteFile()
'Overview                    : �t�@�C�����폜����
'Detailed Description        : FileSystemObject��DeleteFile()�Ɠ���
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
Private Function func_CM_FsDeleteFile( _
    byVal asPath _
    ) 
    If Not func_CM_FsFileExists(asPath) Then func_CM_FsDeleteFile = False
    func_CM_FsDeleteFile = cf_tryCatch(new_Func("a=>a(0).DeleteFile(a(1))"), Array(new_Fso(), asPath), Empty, Empty).Item("Result")
    
'    On Error Resume Next
'    new_Fso().DeleteFile(asPath)
'    func_CM_FsDeleteFile = True
'    If Err.Number Then
'        Err.Clear
'        func_CM_FsDeleteFile = False
'    End If
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_FsDeleteFolder()
'Overview                    : �t�H���_���폜����
'Detailed Description        : FileSystemObject��DeleteFolder()�Ɠ���
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
Private Function func_CM_FsDeleteFolder( _
    byVal asPath _
    ) 
    If Not func_CM_FsFolderExists(asPath) Then func_CM_FsDeleteFolder = False
    On Error Resume Next
    new_Fso().DeleteFolder(asPath)
    func_CM_FsDeleteFolder = True
    If Err.Number Then
        Err.Clear
        func_CM_FsDeleteFolder = False
    End If
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_FsDeleteFsObject()
'Overview                    : �t�@�C�����t�H���_���폜����
'Detailed Description        : func_CM_FsDeleteFile()��func_CM_FsDeleteFolder()�ɈϏ�
'Argument
'     asPath                 : �p�X
'Return Value
'     ���� True:���� / False:���s
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/08/19         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FsDeleteFsObject( _
    byVal asPath _
    )
    func_CM_FsDeleteFsObject = False
    If func_CM_FsFileExists(asPath) Then func_CM_FsDeleteFsObject = func_CM_FsDeleteFile(asPath)
    If func_CM_FsFolderExists(asPath) Then func_CM_FsDeleteFsObject = func_CM_FsDeleteFolder(asPath)
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_FsCopyFile ()
'Overview                    : �t�@�C�����R�s�[����
'Detailed Description        : FileSystemObject��CopyFile()�Ɠ���
'Argument
'     asPathFrom             : �R�s�[���t�@�C���̃t���p�X
'     asPathTo               : �R�s�[��t�@�C���̃t���p�X
'Return Value
'     ���� True:���� / False:���s
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/10/24         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FsCopyFile( _
    byVal asPathFrom _
    , byVal asPathTo _
    ) 
    If Not func_CM_FsFileExists(asPathFrom) Then func_CM_FsCopyFile = False
    On Error Resume Next
    Call new_Fso().CopyFile(asPathFrom, asPathTo)
    func_CM_FsCopyFile = True
    If Err.Number Then
        Err.Clear
        func_CM_FsCopyFile = False
    End If
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_FsCopyFolder ()
'Overview                    : �t�H���_���R�s�[����
'Detailed Description        : FileSystemObject��CopyFolder()�Ɠ���
'Argument
'     asPathFrom             : �R�s�[���t�H���_�̃t���p�X
'     asPathTo               : �R�s�[��t�H���_�̃t���p�X
'Return Value
'     ���� True:���� / False:���s
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/08/19         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FsCopyFolder( _
    byVal asPathFrom _
    , byVal asPathTo _
    ) 
    If Not func_CM_FsFolderExists(asPathFrom) Then func_CM_FsCopyFolder = False
    On Error Resume Next
    Call new_Fso().CopyFolder(asPathFrom, asPathTo)
    func_CM_FsCopyFolder = True
    If Err.Number Then
        Err.Clear
        func_CM_FsCopyFolder = False
    End If
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_FsCopyFsObject()
'Overview                    : �t�@�C�����t�H���_���R�s�[����
'Detailed Description        : func_CM_FsCopyFile()��func_CM_FsCopyFolder()�ɈϏ�
'Argument
'     asPathFrom             : �R�s�[���t�@�C��/�t�H���_�̃t���p�X
'     asPathTo               : �R�s�[��̃t���p�X
'Return Value
'     ���� True:���� / False:���s
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/08/19         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FsCopyFsObject( _
    byVal asPathFrom _
    , byVal asPathTo _
    )
    func_CM_FsCopyFsObject = False
    If func_CM_FsFileExists(asPathFrom) Then func_CM_FsCopyFsObject = func_CM_FsCopyFile(asPathFrom, asPathTo)
    If func_CM_FsFolderExists(asPathFrom) Then func_CM_FsCopyFsObject = func_CM_FsCopyFolder(asPathFrom, asPathTo)
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_FsMoveFile ()
'Overview                    : �t�@�C�����ړ�����
'Detailed Description        : FileSystemObject��MoveFile()�Ɠ���
'Argument
'     asPathFrom             : �ړ����t�@�C���̃t���p�X
'     asPathTo               : �ړ���t�@�C���̃t���p�X
'Return Value
'     ���� True:���� / False:���s
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/08/19         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FsMoveFile( _
    byVal asPathFrom _
    , byVal asPathTo _
    ) 
    If Not func_CM_FsFileExists(asPathFrom) Then func_CM_FsMoveFile = False
    On Error Resume Next
    Call new_Fso().MoveFile(asPathFrom, asPathTo)
    func_CM_FsMoveFile = True
    If Err.Number Then
        Err.Clear
        func_CM_FsMoveFile = False
    End If
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_FsMoveFolder ()
'Overview                    : �t�H���_���ړ�����
'Detailed Description        : FileSystemObject��MoveFolder()�Ɠ���
'Argument
'     asPathFrom             : �ړ����t�H���_�̃t���p�X
'     asPathTo               : �ړ���t�H���_�̃t���p�X
'Return Value
'     ���� True:���� / False:���s
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/08/19         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FsMoveFolder( _
    byVal asPathFrom _
    , byVal asPathTo _
    ) 
    If Not func_CM_FsFolderExists(asPathFrom) Then func_CM_FsMoveFolder = False
    On Error Resume Next
    Call new_Fso().MoveFolder(asPathFrom, asPathTo)
    func_CM_FsMoveFolder = True
    If Err.Number Then
        Err.Clear
        func_CM_FsMoveFolder = False
    End If
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_FsMoveFsObject()
'Overview                    : �t�@�C�����t�H���_���ړ�����
'Detailed Description        : func_CM_FsMoveFile()��func_CM_FsMoveFolder()�ɈϏ�
'Argument
'     asPathFrom             : �ړ����t�@�C��/�t�H���_�̃t���p�X
'     asPathTo               : �ړ���̃t���p�X
'Return Value
'     ���� True:���� / False:���s
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/08/19         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FsMoveFsObject( _
    byVal asPathFrom _
    , byVal asPathTo _
    )
    func_CM_FsMoveFsObject = False
    If func_CM_FsFileExists(asPathFrom) Then func_CM_FsMoveFsObject = func_CM_FsMoveFile(asPathFrom, asPathTo)
    If func_CM_FsFolderExists(asPathFrom) Then func_CM_FsMoveFsObject = func_CM_FsMoveFolder(asPathFrom, asPathTo)
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_FsGetParentFolderPath()
'Overview                    : �e�t�H���_�p�X�̎擾
'Detailed Description        : FileSystemObject��GetParentFolderName()�Ɠ���
'Argument
'     asPath                 : �t�@�C���̃p�X
'Return Value
'     �e�t�H���_�p�X
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/09/27         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FsGetParentFolderPath( _
    byVal asPath _
    ) 
    func_CM_FsGetParentFolderPath = new_Fso().GetParentFolderName(asPath)
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_FsGetGetBaseName()
'Overview                    : �t�@�C�����i�g���q�������j�̎擾
'Detailed Description        : FileSystemObject��GetBaseName()�Ɠ���
'Argument
'     asPath                 : �t�@�C���̃p�X
'Return Value
'     �t�@�C�����i�g���q�������j
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/10/23         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FsGetGetBaseName( _
    byVal asPath _
    ) 
    func_CM_FsGetGetBaseName = new_Fso().GetBaseName(asPath)
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_FsGetGetExtensionName()
'Overview                    : �t�@�C���̊g���q�̎擾
'Detailed Description        : FileSystemObject��GetExtensionName()�Ɠ���
'Argument
'     asPath                 : �t�@�C���̃p�X
'Return Value
'     �t�@�C���̊g���q
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/10/23         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FsGetGetExtensionName( _
    byVal asPath _
    ) 
    func_CM_FsGetGetExtensionName = new_Fso().GetExtensionName(asPath)
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_FsBuildPath()
'Overview                    : �t�@�C���p�X�̘A��
'Detailed Description        : FileSystemObject��BuildPath()�Ɠ���
'Argument
'     asFolderPath           : �p�X
'     asItemName             : asFolderPath�ɘA������t�H���_���܂��̓t�@�C����
'Return Value
'     �A�������t�@�C���p�X
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/09/27         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FsBuildPath( _
    byVal asFolderPath _
    , byVal asItemName _
    ) 
    func_CM_FsBuildPath = new_Fso().BuildPath(asFolderPath, asItemName)
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_FsFileExists()
'Overview                    : �t�@�C���̑��݊m�F
'Detailed Description        : FileSystemObject��FileExists()�Ɠ���
'Argument
'     asPath                 : �p�X
'Return Value
'     ���� True:���݂��� / False:���݂��Ȃ�
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/09/27         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FsFileExists( _
    byVal asPath _
    ) 
    func_CM_FsFileExists = new_Fso().FileExists(asPath)
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_FsFolderExists()
'Overview                    : �t�H���_�̑��݊m�F
'Detailed Description        : FileSystemObject��FolderExists()�Ɠ���
'Argument
'     asPath                 : �p�X
'Return Value
'     ���� True:���݂��� / False:���݂��Ȃ�
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/10/16         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FsFolderExists( _
    byVal asPath _
    ) 
    func_CM_FsFolderExists = new_Fso().FolderExists(asPath)
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_FsGetFile()
'Overview                    : �t�@�C���I�u�W�F�N�g�̎擾
'Detailed Description        : FileSystemObject��GetFile()�Ɠ���
'Argument
'     asPath                 : �p�X
'Return Value
'     File�I�u�W�F�N�g
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/09/27         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FsGetFile( _
    byVal asPath _
    ) 
    Set func_CM_FsGetFile = new_Fso().GetFile(asPath)
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_FsGetFolder()
'Overview                    : �t�H���_�I�u�W�F�N�g�̎擾
'Detailed Description        : FileSystemObject��GetFolder()�Ɠ���
'Argument
'     asPath                 : �p�X
'Return Value
'     Folder�I�u�W�F�N�g
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/10/23         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FsGetFolder( _
    byVal asPath _
    ) 
    Set func_CM_FsGetFolder = new_Fso().GetFolder(asPath)
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_FsGetFsObject()
'Overview                    : �t�@�C�����t�H���_�I�u�W�F�N�g�̎擾
'Detailed Description        : func_CM_FsGetFile()��func_CM_FsGetFolder()�ɈϏ�
'Argument
'     asPath                 : �p�X
'Return Value
'     File/Folder�I�u�W�F�N�g
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/08/19         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FsGetFsObject( _
    byVal asPath _
    )
    Set func_CM_FsGetFsObject = Nothing
    If func_CM_FsFileExists(asPath) Then Set func_CM_FsGetFsObject = func_CM_FsGetFile(asPath)
    If func_CM_FsFolderExists(asPath) Then Set func_CM_FsGetFsObject = func_CM_FsGetFolder(asPath)
End Function

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
'Function/Sub Name           : func_CM_FsGetFsObjects()
'Overview                    : �w�肵���t�H���_�ȉ���Files�R���N�V������Folders�R���N�V�������擾����
'Detailed Description        : func_CM_FsGetFiles()��func_CM_FsGetFolders()�ɈϏ�
'Argument
'     asPath                 : �p�X
'Return Value
'     Files�R���N�V������Folders�R���N�V����
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/08/19         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FsGetFsObjects( _
    byVal asPath _
    )
    Set func_CM_FsGetFsObjects = Nothing
    If Not func_CM_FsFolderExists(asPath) Then Exit Function
    Dim oTemp : Set oTemp = new_Dic()
    With oTemp
        .Add "Filse", func_CM_FsGetFiles(asPath)
        .Add "Folders", func_CM_FsGetFolders(asPath)
    End With
    Set func_CM_FsGetFsObjects = oTemp
    Set oTemp = Nothing
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_FsGetTempFileName()
'Overview                    : �����_���ɐ������ꂽ�ꎞ�t�@�C���܂��̓t�H���_�[�̖��O�̎擾
'Detailed Description        : FileSystemObject��GetTempName()�Ɠ���
'Argument
'     asPath                 : �p�X
'Return Value
'     �ꎞ�t�@�C���܂��̓t�H���_�[�̖��O
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/09/27         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FsGetTempFileName()
    func_CM_FsGetTempFileName = new_Fso().GetTempName()
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_FsGetPrivateFilePath()
'Overview                    : ���s���̃X�N���v�g������t�H���_����̃p�X��Ԃ�
'Detailed Description        : ��ʃt�H���_�����݂��Ȃ��ꍇ�͍쐬����
'Argument
'     asParentFolderName     : �e�t�H���_��
'     asFileName             : �t�@�C����
'Return Value
'     �t�@�C���̃t���p�X
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/26         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FsGetPrivateFilePath( _
    byVal asParentFolderName _
    , byVal asFileName _
    )
    Dim sParentFolderPath : sParentFolderPath = func_CM_FsGetParentFolderPath(WScript.ScriptFullName)
    If Len(asParentFolderName)>0 Then
    '�����Ŏw�肵���f�B���N�g����������ꍇ
        sParentFolderPath = func_CM_FsBuildPath(sParentFolderPath ,asParentFolderName)
    End If
    func_CM_FsGetPrivateFilePath = func_CM_FsGetFilePathWithCreateParentFolder(sParentFolderPath, asFileName)
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_FsGetTempFilePath()
'Overview                    : �ꎞ�t�@�C���̃p�X��Ԃ�
'Detailed Description        : ���s���̃X�N���v�g������t�H���_��tmp�t�H���_�ȉ��ɍ쐬����
'                              ��ʃt�H���_�����݂��Ȃ��ꍇ�͍쐬����
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
Private Function func_CM_FsGetTempFilePath( _
    )
    func_CM_FsGetTempFilePath = func_CM_FsGetPrivateFilePath("tmp", func_CM_FsGetTempFileName())
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_FsGetPrivateLogFilePath()
'Overview                    : ���s���̃X�N���v�g�̃��O�t�@�C���p�X��Ԃ�
'Detailed Description        : ���s���̃X�N���v�g������t�H���_��log�t�H���_�ȉ���
'                              �X�N���v�g�t�@�C�����{".log"�`���̃t�@�C�����ō쐬����
'                              ��ʃt�H���_�����݂��Ȃ��ꍇ�͍쐬����
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
Private Function func_CM_FsGetPrivateLogFilePath( _
    )
    func_CM_FsGetPrivateLogFilePath = func_CM_FsGetPrivateFilePath("log", func_CM_FsGetGetBaseName(WScript.ScriptName) & ".log" )
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_FsGetFilePathWithCreateParentFolder()
'Overview                    : �t�@�C���̃p�X���擾
'Detailed Description        : ��ʃt�H���_�����݂��Ȃ��ꍇ�͍쐬����
'Argument
'     asParentFolderPath     : �e�t�H���_�̃p�X
'     asFileName             : �t�@�C����
'Return Value
'     �t�@�C���̃t���p�X
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/26         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FsGetFilePathWithCreateParentFolder( _
    byVal asParentFolderPath _
    , byVal asFileName _
    )
    If Not(func_CM_FsFolderExists(asParentFolderPath)) Then func_CM_FsCreateFolder(asParentFolderPath)
    func_CM_FsGetFilePathWithCreateParentFolder = func_CM_FsBuildPath(asParentFolderPath, asFileName)
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_FsCreateFolder()
'Overview                    : �t�H���_���쐬����
'Detailed Description        : FileSystemObject��CreateFolder()�Ɠ���
'Argument
'     asPath                 : �p�X
'Return Value
'     �쐬�����t�H���_�̃t���p�X
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/09/27         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FsCreateFolder( _
    byVal asPath _
    )
    func_CM_FsCreateFolder = new_Fso().CreateFolder(asPath)
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_FsOpenTextFile()
'Overview                    : �t�@�C�����J��TextStream�I�u�W�F�N�g��Ԃ�
'Detailed Description        : FileSystemObject��OpenTextFile()�Ɠ���
'Argument
'     asPath                 : �p�X
'     alIomode               : ����/�o�̓��[�h 1:ForReading,2:ForWriting,8:ForAppending
'     aboCreate              : asPath�����݂��Ȃ��ꍇ True:�V�����t�@�C�����쐬����AFalse:�쐬���Ȃ�
'     alFileFormat           : �t�@�C���̌`�� -2:TristateUseDefault,-1:TristateTrue,0:TristateFalse
'Return Value
'     TextStream�I�u�W�F�N�g
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/01/09         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FsOpenTextFile( _
    byVal asPath _
    , byVal alIomode _
    , byVal aboCreate _
    , byVal alFileFormat _
    )
    '�t�@�C�����J��
    Set func_CM_FsOpenTextFile = new_Fso().OpenTextFile( _
                                                              asPath _
                                                              , alIomode _
                                                              , aboCreate _
                                                              , alFileFormat _
                                                              )
End Function

'***************************************************************************************************
'Function/Sub Name           : sub_CM_FsWriteFile()
'Overview                    : �t�@�C���o�͂���
'Detailed Description        : �G���[�͖�������
'Argument
'     asPath                 : �o�͐�̃t���p�X
'     asCont                 : �o�͂�����e
'     �Ȃ�
'Return Value
'     �쐬�����t�H���_�̃t���p�X
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/10/16         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub sub_CM_FsWriteFile( _
    byVal asPath _
    , byVal asCont _
    )
    On Error Resume Next
    '�t�@�C�����J���i���݂��Ȃ��ꍇ�͍쐬����j
    With func_CM_FsOpenTextFile(asPath, 2, True, -2)
        Call .WriteLine(asCont)
        Call .Close
    End With
    If Err.Number Then
        Err.Clear
    End If
End Sub

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

'***************************************************************************************************
'Function/Sub Name           : func_CM_ArrayAddElement()
'Overview                    : �z��̗v�f��ǉ�����
'Detailed Description        : �H����
'Argument
'     avArr                  : �z��
'Return Value
'     �Ȃ�
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/10/15         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_ArrayAddElement( _
    byRef avArr _
    )
    Redim Preserve avArr(Ubound(avArr)+1)
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_ArrayInitialize()
'Overview                    : �z�������������
'Detailed Description        : �H����
'Argument
'     avArr                  : �z��
'     avErr                  : �G���[���
'Return Value
'     �Ȃ�
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/10/15         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_ArrayInitialize( _
    byRef avArr _
    , byRef avErr _
    )
    Redim avArr(0)
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_ArrayUbound()
'Overview                    : �z��̃C���f�b�N�X�̍ő吔��Ԃ�
'Detailed Description        : �H����
'Argument
'     avArr                  : �z��
'Return Value
'     �Ȃ�
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/10/15         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_ArrayUbound( _
    byRef avArr _
    )
    func_CM_ArrayUbound = Ubound(avArr)
End Function

'###################################################################################################
'�`�F�b�N�n
'###################################################################################################

'***************************************************************************************************
'Function/Sub Name           : func_CM_ValidationlIsWithinTheRangeOf()
'Overview                    : ���l�^�͈͓̔��ɂ��邩��������
'Detailed Description        : �H����
'Argument
'     avNumber               : ���l
'     alType                 : �ϐ��̌^
'                                1:�����^�iInteger�j
'                                2:�������^�iLong�j
'                                3:�o�C�g�^�iByte�j
'                                4:�P���x���������_�^�iSingle�j
'                                5:�{���x���������_�^�iDouble�j
'                                6:�ʉ݌^�iCurrency�j
'Return Value
'     ���`�������������_�^
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/08/26         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_ValidationlIsWithinTheRangeOf( _
    byVal avNumber _
    , byVal alType _
    )
    Dim vMin,vMax
    Select Case alType
        Case 1:                   '�����^�iInteger�j
            vMin = -1 * 2^15
            vMax = 2^15 - 1
        Case 2:                   '�������^�iLong�j
            vMin = -1 * 2^31
            vMax = 2^31 - 1
        Case 3:                   '�o�C�g�^�iByte�j
            vMin = 0
            vMax = 2^8 - 1
        Case 4:                   '�P���x���������_�^�iSingle�j
            vMin = -3.402823E38
            vMax = 3.402823E38
        Case 5:                   '�{���x���������_�^�iDouble�j
            vMin = -1.79769313486231E308
            vMax = 1.79769313486231E308
        Case 6:                   '�ʉ݌^�iCurrency�j
            vMin = -1 * 2^59 / 1000
            vMax = ( 2^59 - 1 ) / 1000
    End Select
    
    func_CM_ValidationlIsWithinTheRangeOf = False
    If vMin<=avNumber And avNumber<=vMax Then
        func_CM_ValidationlIsWithinTheRangeOf = True
    End If
End Function


'###################################################################################################
'���̑�
'###################################################################################################

'***************************************************************************************************
'Function/Sub Name           : func_CM_GetObjectByNameFromCollection()
'Overview                    : �R���N�V��������w�肵��name�̃����o�[���擾����
'Detailed Description        : �G���[�����͍s��Ȃ�
'Argument
'     aoArr                  : 0�Ԗځ@�R���N�V�����A1�Ԗځ@name
'Return Value
'     �Y�����郁���o�[
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/09/27         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_GetObjectByNameFromCollection( _
    byRef aoArr _
    )
    cf_bind func_CM_GetObjectByNameFromCollection, aoArr(0).Item(aoArr(1))
End Function

'***************************************************************************************************
'Function/Sub Name           : sub_CM_Swap()
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
Private Sub sub_CM_Swap( _
    byRef avA _
    , byRef avB _
    )
    Dim oTemp
    Call cf_bind(oTemp, avA)
    Call cf_bind(avA, avB)
    Call cf_bind(avB, oTemp)
    Set oTemp = Nothing
End Sub

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

'***************************************************************************************************
'Function/Sub Name           : func_CM_ToString()
'Overview                    : �����̐��l�E�������I�u�W�F�N�g�̒��g���ǂȕ\���ɕϊ�����
'Detailed Description        : �z���f�B�N�V���i���̂悤�ȃI�u�W�F�N�g�������璆�g��\�����A
'                              �����łȂ��ꍇ��VarType�ŃI�u�W�F�N�g�̃N���X��\������
'Argument
'     avTarget               : �Ώ�
'Return Value
'     �ϊ�����������
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/03         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_ToString( _
    byRef avTarget _
    )
    Dim oEscapingDoubleQuote, sRet
    Set oEscapingDoubleQuote = new_Re("""", "g")
    sRet = ""
    
    Err.Clear
    On Error Resume Next
    
    If VarType(avTarget) = vbString Then
        sRet = """" & oEscapingDoubleQuote.Replace(avTarget, """""") & """"
    ElseIf IsArray(avTarget) Then
        sRet = func_CM_ToStringArray(avTarget)
    ElseIf IsObject(avTarget) Then
        sRet = func_CM_ToStringObject(avTarget)
    ElseIf IsEmpty(avTarget) Then
        sRet = "<empty>"
    ElseIf IsNull(avTarget) Then
        sRet = "<null>"
    Else
        sRet = func_CM_ToStringOther(avTarget)
    End If
    
    If Err.Number <> 0 Then
        Err.Clear
        sRet = func_CM_ToStringUnknown(avTarget)
    End If
    
    func_CM_ToString = sRet
    
    Set oEscapingDoubleQuote = Nothing
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_ToStringArray()
'Overview                    : �z��̒��g���ǂȕ\���ɕϊ�����
'Detailed Description        : �H����
'Argument
'     avTarget               : �Ώ�
'Return Value
'     �ϊ�����������
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/03         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_ToStringArray( _
    byRef avTarget _
    )
    Dim oTemp(), vItem
    
    For Each vItem In avTarget
        Call cf_push(oTemp, func_CM_ToString(vItem))
    Next
    func_CM_ToStringArray = "[" & Join(oTemp, ",") & "]"
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_ToStringDictionary()
'Overview                    : �f�B�N�V���i���̒��g���ǂȕ\���ɕϊ�����
'Detailed Description        : �H����
'Argument
'     avTarget               : �Ώ�
'Return Value
'     �ϊ�����������
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/03         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_ToStringDictionary( _
    byRef avTarget _
    )
    Dim oTemp(), vKey
    
    For Each vKey In avTarget.Keys
        Call cf_push(oTemp, func_CM_ToString(vKey) & "=>" & func_CM_ToString(avTarget.Item(vKey)))
    Next
    func_CM_ToStringDictionary = "{" & Join(oTemp, ",") & "}"
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_ToStringObject()
'Overview                    : �I�u�W�F�N�g�̒��g���ǂȕ\���ɕϊ�����
'Detailed Description        : �H����
'Argument
'     avTarget               : �Ώ�
'Return Value
'     �ϊ�����������
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/03         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_ToStringObject( _
    byRef avTarget _
    )
    Dim sRet
    
    Err.Clear
    On Error Resume Next
    
    sRet = func_CM_ToStringDictionary(avTarget)
    
    If Err.Number <> 0 Then
        Err.Clear
        sRet = func_CM_ToStringArray(avTarget)
    End If
    
    If Err.Number <> 0 Then
        Err.Clear
        sRet = func_CM_ToStringArray(avTarget.Items)
    End If
    
    If Err.Number <> 0 Then
        Err.Clear
        sRet = "<" & TypeName(avTarget) & ">"
    End If
    
    func_CM_ToStringObject = sRet
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_ToStringOther()
'Overview                    : ���̑��I�u�W�F�N�g�̒��g���ǂȕ\���ɕϊ�����
'Detailed Description        : �H����
'Argument
'     avTarget               : �Ώ�
'Return Value
'     �ϊ�����������
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/03         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_ToStringOther( _
    byRef avTarget _
    )
    Dim sRet
    
    Err.Clear
    On Error Resume Next
    
    sRet = CStr(avTarget)
    
    If Err.Number <> 0 Then
        Err.Clear
        sRet = func_CM_ToStringArray(avTarget)
    End If
    
    If Err.Number <> 0 Then
        Err.Clear
        sRet = func_CM_ToStringDictionary(avTarget)
    End If
    
    If Err.Number <> 0 Then
        Err.Clear
        sRet = func_CM_ToStringUnknown(avTarget)
    End If
    
    func_CM_ToStringOther = sRet
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_ToStringUnknown()
'Overview                    : �����̌^���s���ȏꍇ�ɉǂȕ\���ɕϊ�����
'Detailed Description        : �H����
'Argument
'     avTarget               : �Ώ�
'Return Value
'     �ϊ�����������
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/03         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_ToStringUnknown( _
    byRef avTarget _
    )
    func_CM_ToStringUnknown = "<unknown:" & VarType(avTarget) & " " & TypeName(avTarget) & ">"
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_ToStringErr()
'Overview                    : Err�I�u�W�F�N�g�̓��e���ǂȕ\���ɕϊ�����
'Detailed Description        : �H����
'Argument
'     �Ȃ�
'Return Value
'     �ϊ�����������
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/25         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_ToStringErr( _
    )
    func_CM_ToStringErr = "<Err> " & func_CM_ToString(func_CM_UtilStoringErr())
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_ToStringArguments()
'Overview                    : Arguments�I�u�W�F�N�g�̓��e���ǂȕ\���ɕϊ�����
'Detailed Description        : �H����
'Argument
'     �Ȃ�
'Return Value
'     �ϊ�����������
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/26         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_ToStringArguments( _
    )
    func_CM_ToStringArguments = "<Arguments> " & func_CM_ToString(func_CM_UtilStoringArguments())
End Function

'***************************************************************************************************
'Function/Sub Name           : sub_CM_ExcuteSub()
'Overview                    : �֐������s����
'Detailed Description        : �H����
'Argument
'     asSubName              : ���s����֐���
'     aoArgument             : ���s����֐��ɓn������
'     aoBroker               : �o��-�w�ǌ^�iPublish/subscribe�j�N���X�̃I�u�W�F�N�g
'Return Value
'     �Ȃ�
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/03         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub sub_CM_ExcuteSub( _
    byVal asSubName _
    , byRef aoArgument _
    , byRef aoBroker _
    )
    Const Cs_TOPIC = "log"
    
    '�o�ŁiPublish�j �J�n
    If Not aoBroker Is Nothing Then
        aoBroker.Publish Cs_TOPIC, Array(5 ,asSubName ,"Start")
        aoBroker.Publish Cs_TOPIC, Array(9 ,asSubName ,func_CM_ToString(aoArgument))
    End If
    
    '�֐��̎��s
    Dim oFunc, oRet
    Set oFunc = GetRef(asSubName)
    If aoArgument Is Nothing Then
        Set oRet = cf_tryCatch( new_Func("function(a){a()}"), oFunc, Empty, Empty )
    Else
        Set oRet = cf_tryCatch( oFunc, aoArgument, Empty, Empty )
    End If
    
    '�o�ŁiPublish�j �I��
    If Not aoBroker Is Nothing Then
        If oRet.Item("Result")=False Then
        '�G���[
            aoBroker.Publish Cs_TOPIC, Array(1, asSubName, func_CM_ToString(oRet.Item("Err")))
        Else
        '����
            aoBroker.Publish Cs_TOPIC, Array(5, asSubName, "End")
        End If
        aoBroker.Publish Cs_TOPIC, Array(9, asSubName, func_CM_ToString(aoArgument))
    End If
    
    Set oRet = Nothing
    Set oFunc = Nothing
End Sub

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
'    If Not func_CM_ArrayIsAvailable(avArr) Then Exit Function
    If Ubound(avArr)=0 Then Exit Function
    
    Dim lEnd, lPos
    lEnd = Ubound(avArr)
    Do While lEnd>0
        For lPos=0 To lEnd-1
            If aoFunc(avArr(lPos), avArr(lPos+1))=aboFlg Then
            'lPos�Ԗڂ̗v�f��(lPos+1)�Ԗڂ̗v�f�����ւ���
                Call sub_CM_Swap(avArr(lPos), avArr(lPos+1))
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
'    If Not func_CM_ArrayIsAvailable(avArr) Then Exit Function
    If Ubound(avArr)=0 Then Exit Function
    
    '0�Ԗڂ̗v�f���s�{�b�g�Ɍ��߂�
    Dim oPivot : Call cf_bind(oPivot, avArr(0))
    
    '�s�{�b�g�Ɨv�f���֐��Ŕ��肵������@�ɍ��v����O���[�v��Right�A�����łȂ��O���[�v��Left�Ƃ���
    Dim lPos, vRight, vLeft
    For lPos=1 To Ubound(avArr)
        If aoFunc(avArr(lPos), oPivot)=aboFlg Then
            Call cf_push(vRight, avArr(lPos))
        Else
            Call cf_push(vLeft, avArr(lPos))
        End If
    Next
    
    '��q�ŕ�����Right�ALeft�̃O���[�v���ƂɍċA��������
    vLeft = func_CM_UtilSortQuick(vLeft, aoFunc, aboFlg)
    vRight = func_CM_UtilSortQuick(vRight, aoFunc, aboFlg)
    
    'Left�Ƀs�{�b�g�{Right����������
    Call cf_push(vLeft, oPivot)
    If new_Arr().hasElement(vRight) Then
'    If func_CM_ArrayIsAvailable(vRight) Then
        For lPos=0 To Ubound(vRight)
            Call cf_push(vLeft, vRight(lPos))
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
'    If Not func_CM_ArrayIsAvailable(avArr) Then Exit Function
    If Ubound(avArr)=0 Then Exit Function
    
    '2�̔z��ɕ�������
    Dim lLength, lMedian
    lLength = Ubound(avArr) - Lbound(avArr) + 1
    lMedian = math_roundUp(lLength/2, 0)
'    lMedian = math_roundUp(lLength/2, 1)
    Dim lPos, vFirst, vSecond
    For lPos=Lbound(avArr) To lMedian-1
        Call cf_push(vFirst, avArr(lPos))
    Next
    For lPos=lMedian To Ubound(avArr)
        Call cf_push(vSecond, avArr(lPos))
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
            Call cf_push(vRet, avSecond(lPosS))
            lPosS = lPosS + 1
        Else
            Call cf_push(vRet, avFirst(lPosF))
            lPosF = lPosF + 1
        End If
    Loop
    
    '���ꂼ��c���Ă�����̔z��̗v�f��ǉ�����
    Dim lPos
    If lPosF<=lEndF Then
        For lPos=lPosF To lEndF
            Call cf_push(vRet, avFirst(lPos))
        Next
    End If
    If lPosS<=lEndS Then
        For lPos=lPosS To lEndS
            Call cf_push(vRet, avSecond(lPos))
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
'    If Not func_CM_ArrayIsAvailable(avArr) Then Exit Function
    If Ubound(avArr)=0 Then Exit Function
    
    '�q�[�v�̍쐬
    Dim lLb, lUb, lSize, lParent
    lLb = Lbound(avArr) : lUb = Ubound(avArr)
    lSize = lUb - lLb + 1
    '�q�����ŉ����̃m�[�h�����ʂɌ����ď��ԂɃm�[�h�P�ʂ̏������s��
    For lParent=lSize\2-1 To lLb Step -1
        Call sub_CM_UtilSortHeapPerNodeProc(avArr, lSize, lParent, aoFunc, aboFlg)
    Next
    
    '�q�[�v�̐擪�i�ő�/�ŏ��l�j�����ԂɎ��o��
    Do While lSize>0
        '�q�[�v�̐擪�Ɩ��������ւ���
        Call sub_CM_Swap(avArr(lLb), avArr(lSize-1))
        '�q�[�v�T�C�Y���P���炵�čč쐬
        lSize = lSize - 1
        Call sub_CM_UtilSortHeapPerNodeProc(avArr, lSize, 0, aoFunc, aboFlg)
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
        Call sub_CM_Swap(avArr(alParent), avArr(lToSwap))
        '����ւ����q�̗v�f�ȉ��̃m�[�h���ď�������
        Call sub_CM_UtilSortHeapPerNodeProc(avArr, alSize, lToSwap, aoFunc, aboFlg)
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
'     �\�[�g��̔z��
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
'Function/Sub Name           : func_CM_UtilGenerateRandomString()
'Overview                    : �����_���ȕ�����𐶐�����
'Detailed Description        : �w�肵�������A�����̎�ނŃ����_���ȕ�����𐶐�����
'Argument
'     alLength               : �����̒���
'     alType                 : �����̎�ށi�����w�肷��ꍇ�͈ȉ��̘a��ݒ肷��j
'                                    1:���p�p���啶��
'                                    2:���p�p��������
'                                    4:���p����
'                                    8:���p�L��
'                                   16:���p�J�^�J�i
'                                   32:���p�J�^�J�i�L��
'                                   64:�S�p�p���啶��
'                                  128:�S�p�p��������
'                                  256:�S�p����
'                                  512:�S�p�L��
'                                 1024:�S�p�Ђ炪��
'                                 2048:�S�p�J�^�J�i
'                                 4096:�S�p�M���V���A�L���������̑啶��
'                                 8192:�S�p�M���V���A�L���������̏�����
'                                16384:�S�p���g
'                                32768:�S�p���� ��1����(16��`47��)
'                                65536:�S�p���� ��2����(48��`84��)
'     avAdditional           : �z��Ŏw�肷�镶����A�O�q�̕����̎�ނƏd������ꍇ�͒ǉ����Ȃ�
'                              �w�肪�Ȃ��ꍇ��Nothing�Ȃǔz��ȊO���w�肷��
'Return Value
'     ��������������
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/24         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_UtilGenerateRandomString( _
    byVal alLength _
    , byVal alType _
    , byRef avAdditional _
    )
    
    '�����̎�ށialType�j�Ŏw�肵�������̃��X�g���쐬����
    Dim vSettings : vSettings = Array( _
          Array( Array("A", "Z") ) _
          , Array( Array("a", "z") ) _
          , Array( Array("0", "9") ) _
          , Array( Array("!", "/"), Array(":", "@"), Array("[", "`"), Array("{", "~") ) _
          , Array( Array("�", "�"), Array("�", "�") ) _
          , Array( Array("�", "�"), Array("�", "�") ) _
          , Array( Array("�`", "�y") ) _
          , Array( Array("��", "��") ) _
          , Array( Array("�O", "�X") ) _
          , Array( Array("�A", "��"), Array("��", "��"), Array("��", "��"), Array("��", "��"), Array("��", "��"), Array("��", "��") ) _
          , Array( Array("��", "��") ) _
          , Array( Array("�@", "��") ) _
          , Array( Array("��", "��"), Array("�@", "�`") ) _
          , Array( Array("��", "��"), Array("�p", "��") ) _
          , Array( Array("��", "��") ) _
          , Array( Array("��", "�r") ) _
          , Array( Array("��", "��"), Array("�@", "�") ) _
          )
    
    Dim lType : lType = alType
    Dim lPowerOf2 : lPowerOf2 = 16          '2^16 = 65536 <= alType�̍ő�l
    Dim oChars : Set oChars = new_Dic()
    Dim lQuotient,lDivide, vSetting, vItem, bCode
'    Dim lQuotient,lDivide, vSetting, vItem, bCode, sCodeHex
    Do Until lPowerOf2<0
        lDivide = 2^lPowerOf2
        lQuotient = lType \ lDivide
        lType = lType Mod lDivide
        
        If lQuotient>0 Then
            vSetting = vSettings(lPowerOf2)
            For Each vItem In vSetting
                For bCode = Asc(vItem(0)) To Asc(vItem(1))
                    oChars.Add bCode, Chr(bCode)
'                    sCodeHex = Right(Hex(bCode),2)
'                    If bCode>0 Or (sCodeHex<>"7F" And ("3F"<sCodeHex And sCodeHex<"FD")) Then
'                        oChars.Add bCode, Chr(bCode)
'                    End If
                Next
            Next
        End If
        
        lPowerOf2 = lPowerOf2 - 1
    Loop
    
    'sjis�g�p�͈͊O�̃R�[�h�����O����
    Dim sCodeHex
    For Each bCode In oChars.Keys()
        If bCode<0 Then
            sCodeHex = Right(Hex(bCode),2)
            If sCodeHex="7F" Or sCodeHex<="3F" Or "FD"<=sCodeHex Then
                oChars.Remove bCode
            End If
        End If
    Next
    
    '�z��Ŏw�肷�镶����iavAdditional�j��ǉ�����
    If Not IsObject(avAdditional) Then
        If IsArray(avAdditional) And (Not IsEmpty(avAdditional)) Then
            Dim sChar
            For Each sChar In avAdditional
                If Not oChars.Exists(Asc(sChar)) Then
                    oChars.Add Asc(sChar), sChar
                End If
            Next
        End If
    End If

    '��q�ō쐬���������̃��X�g���g���ă����_���ȕ�����𐶐�����
    Dim lPos, sRet
    sRet = ""
    For lPos = 1 To alLength
        sRet = sRet & oChars.Items()( math_rand(0, oChars.Count - 1, 0) )
    Next
    func_CM_UtilGenerateRandomString = sRet
    
    Set oChars = Nothing
End Function

'***************************************************************************************************
'Function/Sub Name           : sub_CM_UtilLogger()
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
Private Sub sub_CM_UtilLogger( _
    byRef avParams _
    , byRef aoWriter _
    )
    Dim oCont, sIp
    sIp = new_ArrWith(func_CM_UtilGetIpaddress()).filter(new_Func("(e,i,a)=>left(e.item(""Ip"").item(""V4""),3)<>""172"""))(0).Item("Ip").Item("V4")
    Set oCont = new_ArrWith(Array(new_Now(), sIp, func_CM_UtilGetComputerName()))
    
    With aoWriter
        .Write(oCont.Concat(avParams).join(vbTab))
        .newLine()
    End With
    Set oCont = Nothing
End Sub

'***************************************************************************************************
'Function/Sub Name           : func_CM_UtilStoringErr()
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
Private Function func_CM_UtilStoringErr( _
    )
    Dim oRet : Set oRet = new_Dic()
    oRet.Add "Number", Err.Number
    oRet.Add "Description", Err.Description
    oRet.Add "Source", Err.Source
    Set func_CM_UtilStoringErr = oRet
    Set oRet = Nothing
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_UtilStoringArguments()
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
Private Function func_CM_UtilStoringArguments( _
    )
    Dim oRet : Set oRet = new_Dic()
    Dim oTemp, oEle, oKey
    
    'All
    Set oTemp = new_Arr()
    For Each oEle In WScript.Arguments
        oTemp.Push oEle
    Next
    oRet.Add "All", oTemp
    
    'Named
    Set oTemp = new_Dic()
    For Each oKey In WScript.Arguments.Named
        oTemp.Add oKey, WScript.Arguments.Named.Item(oKey)
    Next
    oRet.Add "Named", oTemp
    
    'Unnamed
    Set oTemp = new_Arr()
    For Each oEle In WScript.Arguments.Unnamed
        oTemp.Push oEle
    Next
    oRet.Add "Unnamed", oTemp
    
    Set func_CM_UtilStoringArguments = oRet
    
    Set oKey = Nothing
    Set oEle = Nothing
    Set oTemp = Nothing
    Set oRet = Nothing
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_UtilGetIpaddress()
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
Private Function func_CM_UtilGetIpaddress( _
    )
    Dim sMyComp, oAdapter, oAddress, oRet, oIpv4, oIpv6
    
    sMyComp = "."
    Set oRet = new_Arr()
    For Each oAdapter in GetObject("winmgmts:\\"&sMyComp&"\root\cimv2").ExecQuery("Select * From Win32_NetworkAdapterConfiguration Where IPEnabled = True")
         For Each oAddress in oAdapter.IPAddress
             If new_ArrSplit(oAddress, ".").length=4 Then
             'IPv4
                 cf_bind oIpv4, oAddress
             Else
             'IPv6
                 cf_bind oIpv6, oAddress
             End If
         Next
         oRet.push new_DicWith(Array("Caption", oAdapter.Caption, "Ip", new_DicWith(Array("V4", oIpv4, "V6", oIpv6))))
    Next
    func_CM_UtilGetIpaddress = oRet.items
    
    Set oAddress = Nothing
    Set oAdapter = Nothing
    Set oRet = Nothing
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_UtilGetComputerName()
'Overview                    : ���g�̃R���s���[�^�����擾����
'Detailed Description        : �H����
'Argument
'     �Ȃ�
'Return Value
'     ���g�̃R���s���[�^��
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/10/10         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_UtilGetComputerName( _
    )
    func_CM_UtilGetComputerName = CreateObject("WScript.Network").ComputerName
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_UtilIsSame()
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
Private Function func_CM_UtilIsSame( _
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
    func_CM_UtilIsSame = boFlg
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_UtilIsAvailableObject()
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
Private Function func_CM_UtilIsAvailableObject( _
    byRef aoObj _
    )
    Dim boFlg : boFlg = False
    If IsObject(aoObj) Then
        If Not aoObj Is Nothing Then boFlg = True
    End If
    func_CM_UtilIsAvailableObject = boFlg
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
    If Vartype(aoObj)=9 And Strcomp(Typename(aoObj),"TextStream",vbBinaryCompare)=0 Then boFlg = True
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
