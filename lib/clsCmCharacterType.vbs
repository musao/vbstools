'***************************************************************************************************
'FILENAME                    : clsCmCharacterType.vbs
'Overview                    : ������ފǗ��N���X
'Detailed Description        : �H����
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/10/28         Y.Fujii                  First edition
'***************************************************************************************************
Class clsCmCharacterType
    '�N���X���ϐ��A�萔
    Private Cl_MAX_POWER_OF_2
    Private PvSettings
    Private PoChar2Type
    Private PoType2Chars
    
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
    '2023/10/28         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub Class_Initialize()
        Cl_MAX_POWER_OF_2 = 16          '2^16 = 65536 <= Type�̍ő�l
        PvSettings = Array( _
              Array( Array("A", "Z") ) _
              , Array( Array("a", "z") ) _
              , Array( Array("0", "9") ) _
              , Array( Array(" ", "/"), Array(":", "@"), Array("[", "`"), Array("{", "~") ) _
              , Array( Array("�", "�"), Array("�", "�") ) _
              , Array( Array("�", "�"), Array("�", "�") ) _
              , Array( Array("�`", "�y") ) _
              , Array( Array("��", "��") ) _
              , Array( Array("�O", "�X") ) _
              , Array( Array("�@", "��"), Array("��", "��"), Array("��", "��"), Array("��", "��"), Array("��", "��"), Array("��", "��") ) _
              , Array( Array("��", "��") ) _
              , Array( Array("�@", "��") ) _
              , Array( Array("��", "��"), Array("�@", "�`") ) _
              , Array( Array("��", "��"), Array("�p", "��") ) _
              , Array( Array("��", "��") ) _
              , Array( Array("��", "�r") ) _
              , Array( Array("��", "��"), Array("�@", "�") ) _
              )
        Set PoChar2Type = new_Dic()
        Set PoType2Chars = new_Dic()
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
    '2023/10/28         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub Class_Terminate()
        Set PoChar2Type = Nothing
        Set PoType2Chars = Nothing
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : whatType()
    'Overview                    : �����̎�ނ�Ԃ�
    'Detailed Description        : �H����
    'Argument
    '     asChar                 : ����
    'Return Value
    '     �����̎�ށi���e��getCharList()�̈����ialType�j�Ɠ����j
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/28         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function whatType( _
        byVal asChar _
        )
        Dim bCode : bCode = Asc(asChar)
        If PoChar2Type.Exists(bCode) Then
            whatType = PoChar2Type.Item(bCode)
            Exit Function
        End If

        Dim lPowerOf2 : lPowerOf2 = 0
        Do While lPowerOf2 <= Cl_MAX_POWER_OF_2
            If Not PoType2Chars.Exists(2^lPowerOf2) Then
                sub_CmCharTypeCreateDefinitionsByCharacterType lPowerOf2
                If PoChar2Type.Exists(bCode) Then
                    whatType = PoChar2Type.Item(bCode)
                    Exit Function
                End If
            End If
            lPowerOf2 = lPowerOf2+1
        Loop
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : getCharList()
    'Overview                    : �w�肵�������̎�ނ̔z���Ԃ�
    'Detailed Description        : http://charset.7jp.net/sjis.html
    'Argument
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
    'Return Value
    '     �����̎�ށi�z��j
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/28         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function getCharList( _
        byVal alType _
        )
        Dim lType : lType = alType
        Dim lPowerOf2 : lPowerOf2 = Cl_MAX_POWER_OF_2
        Dim vRet : Set vRet = new_Arr()
        Dim lQuotient,lDivide
        Do Until lPowerOf2<0
            lDivide = 2^lPowerOf2
            lQuotient = lType \ lDivide
            lType = lType Mod lDivide
            If lQuotient>0 Then
                If Not PoType2Chars.Exists(lDivide) Then
                    sub_CmCharTypeCreateDefinitionsByCharacterType lPowerOf2
                End If
                vRet.pushMulti PoType2Chars.Item(lDivide)
            End If
            lPowerOf2 = lPowerOf2 - 1
        Loop
        getCharList = vRet.items()
    End Function
    

    '***************************************************************************************************
    'Function/Sub Name           : sub_CmCharTypeCreateDefinitionsByCharacterType
    'Overview                    : �w�肵��������ނ̒�`���쐬����
    'Detailed Description        : �H����
    'Argument
    '     alPowerOf2             : �����̎�ށi���e��getCharList()�̈����ialType�j�Ɠ����j��2^n�Ƃ����ꍇ��n
    'Return Value
    '     �Ȃ�
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/30         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub sub_CmCharTypeCreateDefinitionsByCharacterType( _
        byVal alPowerOf2 _
        )
        Dim lType : lType = 2^alPowerOf2
        If PoType2Chars.Exists(lType) Then Exit Sub

        Dim vArr : vArr = Array()
        Dim vSetting : vSetting = PvSettings(alPowerOf2)
        Dim vEle, bCode, sCodeHex
        For Each vEle In vSetting
            For bCode = Asc(vEle(0)) To Asc(vEle(1))
                sCodeHex = "" : If bCode<0 Then sCodeHex = Right(Hex(bCode),2)
                If bCode>=0 Or (sCodeHex<>"7F" And "3F"<sCodeHex And sCodeHex<"FD" ) Then
                    PoChar2Type.Add bCode, lType
                    cf_push vArr, Chr(bCode)
                End If
            Next
        Next
        PoType2Chars.Add lType, vArr
    End Sub
    
End Class
