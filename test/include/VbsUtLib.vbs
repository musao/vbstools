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
                                            & "_" & func_UtGetGetDateInYyyymmddhhmmssFormat() _
                                            & ".log" _
                                        )
    End With
End Function

'***************************************************************************************************
'Function/Sub Name           : func_UtGetGetDateInYyyymmddhhmmssFormat()
'Overview                    : ���t��YYYYMMDD_HHMMSS�`���Ŏ擾����
'Detailed Description        : �H����
'Argument
'     �Ȃ�
'Return Value
'     YYYYMMDD_HHMMSS�`���̕�����
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/10/12         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_UtGetGetDateInYyyymmddhhmmssFormat()
    Dim dtNow : dtNow = Now()
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
