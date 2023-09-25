'***************************************************************************************************
'FILENAME                    : GeneratePassword.vbs
'Overview                    : �p�X���[�h�𐶐�����
'Detailed Description        : �H����
'Argument
'     �Ȃ�
'Return Value
'     �Ȃ�
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/24         Y.Fujii                  First edition
'***************************************************************************************************
Option Explicit

'�萔
Private Const Cs_FOLDER_INCLUDE = "include"
Private PoWriter
Private PoPubSub

'Include�p�֐���`
Sub sub_Include( _
    byVal asIncludeFileName _
    )
    With CreateObject("Scripting.FileSystemObject")
        Dim sParentFolderName : sParentFolderName = .GetParentFolderName(WScript.ScriptFullName)
        Dim sIncludeFilePath
        sIncludeFilePath = .BuildPath(sParentFolderName, Cs_FOLDER_INCLUDE)
        sIncludeFilePath = .BuildPath(sIncludeFilePath, asIncludeFileName)
        ExecuteGlobal .OpenTextfile(sIncludeFilePath).ReadAll
    End With
End Sub
'Include
Call sub_Include("clsCmArray.vbs")
Call sub_Include("clsCmBufferedWriter.vbs")
Call sub_Include("clsCmCalendar.vbs")
Call sub_Include("clsCmPubSub.vbs")
Call sub_Include("clsCompareExcel.vbs")
Call sub_Include("VbsBasicLibCommon.vbs")

'���C���֐����s
Call Main()
Wscript.Quit


'***************************************************************************************************
'Processing Order            : First
'Function/Sub Name           : Main()
'Overview                    : ���C���֐�
'Detailed Description        : �H����
'Argument
'     �Ȃ�
'Return Value
'     �Ȃ�
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/24         Y.Fujii                  First edition
'***************************************************************************************************
Sub Main()
    Set PoPubSub = Nothing
    Dim oParams : Set oParams = new_Dictionary()
    
    '������
    Call sub_CM_ExcuteSub("sub_GnrtPwInitialize", oParams, PoPubSub, "log")
    
    '���X�N���v�g�̈����擾�i�����Ȃ��j
    Call sub_CM_ExcuteSub("sub_GnrtPwGetParameters", oParams, PoPubSub, "log")
    
    '��r�Ώۃt�@�C�����͉�ʂ̕\���Ǝ擾
    Call sub_CM_ExcuteSub("sub_GnrtPwGenerate", oParams, PoPubSub, "log")
    
    '�I������
    Call sub_CM_ExcuteSub("sub_GnrtPwTerminate", oParams, PoPubSub, "log")
    
    '�t�@�C���ڑ����N���[�Y����
    Call PoWriter.FileClose()
    
    '�I�u�W�F�N�g���J��
    Set oParams = Nothing
    Set PoPubSub = Nothing
    Set PoWriter = Nothing
End Sub

'***************************************************************************************************
'Processing Order            : 1
'Function/Sub Name           : sub_GnrtPwInitialize()
'Overview                    : ������
'Detailed Description        : �H����
'Argument
'     aoParams               : �p�����[�^�i�[�p�ėp�n�b�V���}�b�v
'Return Value
'     �Ȃ�
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/24         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub sub_GnrtPwInitialize( _
    byRef aoParams _
    )
    '���O�o�͂̐ݒ�
    Dim sPath : sPath = func_CM_FsBuildPath( _
                    func_CM_FsGetPrivateFolder("log") _
                    , func_CM_FsGetGetBaseName(WScript.ScriptName) & new_clsCalGetNow().DisplayFormatAs("_YYMMDD_hhmmss.000.log") _
                    )
    Set PoWriter = new_clsCmBufferedWriter(func_CM_FsOpenTextFile(sPath, 8, True, -2))
    
    '�o��-�w�ǌ^�iPublish/subscribe�j�C���X�^���X�̐ݒ�
    Set PoPubSub = new_clsCmPubSub()
    Call PoPubSub.Subscribe("log", GetRef("sub_GnrtPwLogger"))
    
End Sub

'***************************************************************************************************
'Processing Order            : 2
'Function/Sub Name           : sub_GnrtPwGetParameters()
'Overview                    : ���X�N���v�g�̈����擾�i�����Ȃ��j
'Detailed Description        : �H����
'Argument
'     aoParams               : �p�����[�^�i�[�p�ėp�n�b�V���}�b�v
'Return Value
'     �Ȃ�
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/24         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub sub_GnrtPwGetParameters( _
    byRef aoParams _
    )
End Sub

'***************************************************************************************************
'Processing Order            : 3
'Function/Sub Name           : sub_GnrtPwDispInputFiles()
'Overview                    : ��r�Ώۃt�@�C�����͉�ʂ̕\���Ǝ擾
'Detailed Description        : �p�����[�^�i�[�p�ėp�n�b�V���}�b�v��Key="Parameter"�Ŋi�[����
'                              �p�����[�^�i�[�p�n�b�V���}�b�v�̍\��
'                              Key                      Value
'                              -------------------      --------------------------------------------
'                              Seq(1,2)                 ��r����G�N�Z���t�@�C���̃p�X
'                              ��r����G�N�Z���t�@�C���̃p�X��2�����̏ꍇ�ɕs�������擾�i�[����
'Argument
'     aoParams               : �p�����[�^�i�[�p�ėp�n�b�V���}�b�v
'Return Value
'     �Ȃ�
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/24         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub sub_GnrtPwGenerate( _
    byRef aoParams _
    )
    Dim sPw : sPw = func_CM_UtilGenerateRandomString(16, 15, Nothing)
    aoParams.Add "GeneratedPassword", sPw
'    Call CreateObject("Wscript.Shell").Run("cmd /c clip <""" & sPw & """", 0, True)
    
    Dim sMsg, sTitle
    sMsg = "�p�X���[�h�𐶐����܂���" & vbNewLine & "OK�{�^������������ƃN���b�v�{�[�h�ɃR�s�[���܂�"
    sTitle = new_clsCalGetNow() & " �ɍ쐬"
    Call Inputbox(sMsg, sTitle, sPw)
    
End Sub

'***************************************************************************************************
'Processing Order            : 4
'Function/Sub Name           : sub_GnrtPwTerminate()
'Overview                    : �I������
'Detailed Description        : �H����
'Argument
'     aoParams               : �p�����[�^�i�[�p�ėp�n�b�V���}�b�v
'Return Value
'     �Ȃ�
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/03         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub sub_GnrtPwTerminate( _
    byRef aoParams _
    )
End Sub

'***************************************************************************************************
'Processing Order            : -
'Function/Sub Name           : sub_GnrtPwLogger()
'Overview                    : ���O�o�͂���
'Detailed Description        : �H����
'Argument
'     avParams               : �z��^�̃p�����[�^���X�g
'Return Value
'     �Ȃ�
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/03         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub sub_GnrtPwLogger( _
    byRef avParams _
    )
    With PoWriter
        Dim sCont : sCont = new_clsCalGetNow()
        sCont = sCont & vbTab & avParams(0)
        sCont = sCont & vbTab & avParams(1)
        sCont = sCont & vbTab & avParams(2)
        .WriteContents(sCont)
        .newLine()
    End With
End Sub
