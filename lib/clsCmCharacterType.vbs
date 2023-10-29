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
        sub_CmCharTypeCreateCharData
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
        End If
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
        Dim lPowerOf2 : lPowerOf2 = 16          '2^16 = 65536 <= alType�̍ő�l
        Dim vRet : Set vRet = new_Arr()
        Dim lQuotient,lDivide
        Do Until lPowerOf2<0
            lDivide = 2^lPowerOf2
            lQuotient = lType \ lDivide
            lType = lType Mod lDivide
            If lQuotient>0 Then
                vRet.pushMulti PoType2Chars.Item(lDivide)
            End If
            lPowerOf2 = lPowerOf2 - 1
        Loop
        getCharList = vRet.items()
    End Function
    


    '***************************************************************************************************
    'Function/Sub Name           : sub_CmCharTypeCreateCharData()
    'Overview                    : �����̎�ނ̒�`���쐬����
    'Detailed Description        : �H����
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
    Private Sub sub_CmCharTypeCreateCharData( _
        )
        '��`
        Dim vSettings : vSettings = Array( _
              Array( Array("A", "Z") ) _
              , Array( Array("a", "z") ) _
              , Array( Array("0", "9") ) _
              , Array( Array(" ", "/"), Array(":", "@"), Array("[", "`"), Array("{", "~") ) _
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
        
        Dim oChar2Type : Set oChar2Type = new_Dic()
        Dim oType2Chars : Set oType2Chars = new_Dic()

        Const Cl_MAX_POWER_OF_2 = 16          '2^16 = 65536 <= Type�̍ő�l
        Dim lPowerOf2 : lPowerOf2 = 0
        Dim lType, vSetting, vEle, bCode, vArr, sCodeHex
        Do While lPowerOf2 <= Cl_MAX_POWER_OF_2
            lType = 2^lPowerOf2
            vSetting = vSettings(lPowerOf2)
            vArr = Array()
            For Each vEle In vSetting
                For bCode = Asc(vEle(0)) To Asc(vEle(1))
                    sCodeHex = "" : If bCode<0 Then sCodeHex = Right(Hex(bCode),2)
                    If bCode>=0 Or (sCodeHex<>"7F" And "3F"<sCodeHex And sCodeHex<"FD" ) Then
                        oChar2Type.Add bCode, lType
                        cf_push vArr, bCode
                    End If
                Next
            Next
            oType2Chars.Add lType, vArr
            lPowerOf2 = lPowerOf2+1
        Loop

        Set PoChar2Type = oChar2Type
        Set PoType2Chars = oType2Chars
    End Sub
    
End Class
