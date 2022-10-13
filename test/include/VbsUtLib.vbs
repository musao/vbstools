'***************************************************************************************************
'FILENAME                    : VbsUrLib.vbs
'Overview                    : �P�̃e�X�g�p���C�u����
'Detailed Description        : �H����
'Argument
'     �Ȃ�
'Return Value
'     �Ȃ�
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/10/12         Y.Fujii                  First edition
'***************************************************************************************************

'***************************************************************************************************
'Function/Sub Name           : sub_UtResultOutput()
'Overview                    : UT���ʂ��o�͂���
'Detailed Description        : �H����
'Argument
'     aoUtAssistant
'Return Value
'     �Ȃ�
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/10/13         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub sub_UtResultOutput(_
    byRef aoUtAssistant _
    )
    
    With aoUtAssistant
        '���O�t�@�C���o��
        Call sub_UtWriteFile(func_UtGetThisLogFilePath(), .OutputReportInTsvFormat())
        
        '���ʂ����b�Z�[�W�ŏo��
        Dim sMsg : sMsg = "NG������܂��A���O���m�F��������"
        If .isAllOk Then sMsg = "�S�P�[�XOK�ł��I"
        Call Msgbox(sMsg)
    End With
    
End sub

'***************************************************************************************************
'Function/Sub Name           : func_UtGetThisWorkFolderPath()
'Overview                    : UT�Ώۂ̃\�[�X�t�@�C���p�̃��[�N�f�B���N�g���̃t���p�X���擾
'Detailed Description        : �f�B���N�g�����Ȃ��ꍇ�͍쐬����
'Argument
'     �Ȃ�
'Return Value
'     UT�Ώۂ̃\�[�X�t�@�C���p�̃��[�N�f�B���N�g���̃t���p�X
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/10/12         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_UtGetThisWorkFolderPath()
    With CreateObject("Scripting.FileSystemObject")
        Dim sThisWorkFolderPath
        sThisWorkFolderPath = .BuildPath( _
                                        .GetParentFolderName(WScript.ScriptFullName) _
                                        , .GetBaseName(WScript.ScriptFullName) _
                                        )
        If Not(.FolderExists(sThisWorkFolderPath)) Then .CreateFolder(sThisWorkFolderPath)
        func_UtGetThisWorkFolderPath = sThisWorkFolderPath
    End With
End Function

'***************************************************************************************************
'Function/Sub Name           : func_UtGetThisTempFilePath()
'Overview                    : �ꎞ�t�@�C���̃t���p�X���擾
'Detailed Description        : �H����
'Argument
'     �Ȃ�
'Return Value
'     �ꎞ�t�@�C���̃t���p�X
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/10/12         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_UtGetThisTempFilePath()
    With CreateObject("Scripting.FileSystemObject")
        func_UtGetThisTempFilePath = .BuildPath( _
                                        func_UtGetThisWorkFolderPath() _
                                        , .GetTempName() _
                                        )
    End With
End Function

'***************************************************************************************************
'Function/Sub Name           : func_UtGetThisLogFilePath()
'Overview                    : ���O�t�@�C���̃t���p�X���擾
'Detailed Description        : �H����
'Argument
'     �Ȃ�
'Return Value
'     ���O�t�@�C���̃t���p�X
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/10/12         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_UtGetThisLogFilePath()
    With CreateObject("Scripting.FileSystemObject")
        func_UtGetThisLogFilePath = .BuildPath( _
                                        func_UtGetThisWorkFolderPath() _
                                        , .GetBaseName(WScript.ScriptFullName) _
                                            & "_" & func_UtGetGetDateInYyyymmddhhmmssFormat(Now()) _
                                            & ".log" _
                                        )
    End With
End Function

'***************************************************************************************************
'Function/Sub Name           : func_UtGetGetDateInYyyymmddhhmmssFormat()
'Overview                    : ������YYYYMMDD_HHMMSS�`���Ŏ擾����
'Detailed Description        : �H����
'Argument
'     adtDate                : ����
'Return Value
'     YYYYMMDD_HHMMSS�`���̕�����
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/10/12         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_UtGetGetDateInYyyymmddhhmmssFormat(_
    byVal adtDate _
    )
    Dim dtNow : dtNow = adtDate
    Dim sCont : sCont = Year(dtNow)
    sCont = sCont & Right("0" & Month(dtNow) , 2)
    sCont = sCont & Right("0" & Day(dtNow) , 2)
    sCont = sCont & "_"
    sCont = sCont & Right("0" & Hour(dtNow) , 2)
    sCont = sCont & Right("0" & Minute(dtNow) , 2)
    sCont = sCont & Right("0" & Second(dtNow) , 2)
    func_UtGetGetDateInYyyymmddhhmmssFormat = sCont
End Function

'***************************************************************************************************
'Function/Sub Name           : sub_UtWriteFile()
'Overview                    : �t�@�C���o�͂���
'Detailed Description        : �G���[�͖�������
'Argument
'     asPath                 : �o�͐�̃t���p�X
'     asCont                 : �o�͂�����e
'     �Ȃ�
'Return Value
'     �Ȃ�
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/10/12         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub sub_UtWriteFile(_
    byVal asPath _
    , byVal asCont _
    )
    On Error Resume Next
    With CreateObject("Scripting.FileSystemObject")
        Call .OpenTextFile(asPath, 8, True).WriteLine(asCont)
    End With
    If Err.Number Then
        Err.Clear
    End If
End sub
