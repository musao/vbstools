'***************************************************************************************************
'FILENAME                    : CompExcel.vbs
'Overview                    : �G�N�Z���t�@�C�����r����
'Detailed Description        : �����Ŏw�肳�ꂽ�G�N�Z���t�@�C�����r�ΏۂƂ���
'                              �w�肪�Ȃ��܂���1�����̏ꍇ�́A�_�C�A���O�Ŕ�r�Ώۂ̓��͂����߂�
'Argument                    : ���O�Ȃ������i/Key:Value �`���łȂ��j�̂�
'                                1,2�Ԗ�   : ��r����G�N�Z���t�@�C���̃p�X�i�Ƃ��ɏȗ��\�j
'                                3�Ԗڈȍ~ : ��������
'Return Value
'     �Ȃ�
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2017/04/26         Y.Fujii                  First edition
'***************************************************************************************************
Option Explicit

'�萔
Private Const Cs_FOLDER_LIB = "lib"
Private PoWriter, PoBroker

'import��`
Sub sub_import( _
    byVal asIncludeFileName _
    )
    With CreateObject("Scripting.FileSystemObject")
        Dim sParentFolderName : sParentFolderName = .GetParentFolderName(WScript.ScriptFullName)
        Dim sIncludeFilePath
        sIncludeFilePath = .BuildPath(sParentFolderName, Cs_FOLDER_LIB)
        sIncludeFilePath = .BuildPath(sIncludeFilePath, asIncludeFileName)
        ExecuteGlobal .OpenTextfile(sIncludeFilePath).ReadAll
    End With
End Sub
'import
sub_import "clsCmArray.vbs"
sub_import "clsCmBroker.vbs"
sub_import "clsCmBufferedReader.vbs"
sub_import "clsCmBufferedWriter.vbs"
sub_import "clsCmCalendar.vbs"
sub_import "clsCmCharacterType.vbs"
sub_import "clsCmCssGenerator.vbs"
sub_import "clsCmHtmlGenerator.vbs"
sub_import "libCom.vbs"
sub_import "clsCompareExcel.vbs"

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
'2017/04/26         Y.Fujii                  First edition
'***************************************************************************************************
Sub Main()
    '���O�o�͂̐ݒ�
    Set PoWriter = new_WriterTo(func_CM_FsGetPrivateLogFilePath(), 8, True, -1)
    '�u���[�J�[�N���X�̃C���X�^���X�̐ݒ�
    Set PoBroker = new_Broker()
    PoBroker.subscribe "log", GetRef("sub_CmpExcelLogger")
    '�p�����[�^�i�[�p�ėp�I�u�W�F�N�g�錾
    Dim oParams : Set oParams = new_Dic()
    
    '���X�N���v�g�̈����擾
    fw_excuteSub "sub_CmpExcelGetParameters", oParams, PoBroker
    
    '��r�Ώۃt�@�C�����͉�ʂ̕\���Ǝ擾
    fw_excuteSub "sub_CmpExcelDispInputFiles", oParams, PoBroker
    
    '�G�N�Z���t�@�C�����r����
    fw_excuteSub "sub_CmpExcelCompareFiles", oParams, PoBroker
    
    '���O�o�͂��N���[�Y
    PoWriter.close
    
    '�I�u�W�F�N�g���J��
    Set oParams = Nothing
    Set PoBroker = Nothing
    Set PoWriter = Nothing
End Sub

'***************************************************************************************************
'Processing Order            : 1
'Function/Sub Name           : sub_CmpExcelGetParameters()
'Overview                    : ���X�N���v�g�̈����擾
'Detailed Description        : �p�����[�^�i�[�p�ėp�I�u�W�F�N�g��Key="Param"�Ŋi�[����
'                              �z��iclsCmArray�^�j�ɖ��O�Ȃ������i/Key:Value �`���łȂ��j��������
'                              2�Ԗڂ܂Ŏ擾����
'                              ���O�Ȃ�������3�Ԗڈȍ~���邢�͖��O�t�������i/Key:Value �`���j�͖�������
'                              Index   Contents
'                              -----   -------------------------------------------------------------
'                              0       ���O�Ȃ�������1�Ԗ�
'                              1       ���O�Ȃ�������2�Ԗ�
'Argument
'     aoParams               : �p�����[�^�i�[�p�ėp�I�u�W�F�N�g
'Return Value
'     �Ȃ�
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2017/04/26         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub sub_CmpExcelGetParameters( _
    byRef aoParams _
    )
    '�I���W�i���̈������擾
    Dim oArg : Set oArg = fw_storeArguments()
    '�����O�o��
    sub_CmpExcelLogger Array(9, "sub_CmpExcelGetParameters", cf_toString(oArg))
    
    '�p�����[�^�i�[�p�I�u�W�F�N�g�ɐݒ�
    cf_bindAt aoParams, "Param", new_ArrWith(oArg.Item("Unnamed")).slice(0,2)
    
    Set oArg = Nothing
End Sub

'***************************************************************************************************
'Processing Order            : 2
'Function/Sub Name           : sub_CmpExcelDispInputFiles()
'Overview                    : ��r�Ώۃt�@�C�����͉�ʂ̕\���Ǝ擾
'Detailed Description        : �����Ŕ�r����G�N�Z���t�@�C���̎w�肪�Ȃ��ꍇ�AExcel.Application��
'                              �_�C�A���O��\�����ă��[�U�Ƀt�@�C����I��������
'                              Index   Contents
'                              -----   -------------------------------------------------------------
'                              0       Excel.Application�̃_�C�A���O�őI�������t�@�C���p�X��ݒ肷��
'                              1       ����
'Argument
'     aoParams               : �p�����[�^�i�[�p�ėp�I�u�W�F�N�g
'Return Value
'     �Ȃ�
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2017/04/26         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub sub_CmpExcelDispInputFiles( _
    byRef aoParams _
    )
    Dim oParam : Set oParam = aoParams.Item("Param")
    If oParam.length > 1 Then
    '�p�����[�^��2�ȏゾ������֐��𔲂���
        '�����O�o��
        sub_CmpExcelLogger Array(3, "sub_CmpExcelDispInputFiles", "No dialog required.")
        Exit Sub
    End If
    
    '�p�����[�^�i�[�p�ėp�I�u�W�F�N�g
    Const Cs_TITLE_EXCEL = "��r�Ώۃt�@�C�����J��"
    With CreateObject("Excel.Application")
        Dim sPath
        Do Until oParam.Length > 1
            
            sPath = .GetOpenFilename( , , Cs_TITLE_EXCEL, , False)
            If sPath = False Then
            '�t�@�C���I���L�����Z���̏ꍇ�͓��X�N���v�g���I������
                '�����O�o��
                sub_CmpExcelLogger Array(3, "sub_CmpExcelDispInputFiles", "Dialog input canceled.")
                
                PoWriter.close
                Set oParam = Nothing
                Wscript.Quit
            End If
            '�I�������t�@�C���̃p�X���擾
            oParam.push sPath
        Loop
        
        .Quit
    End With
    
    '�I�u�W�F�N�g���J��
    Set oParam = Nothing
End Sub

'***************************************************************************************************
'Processing Order            : 3
'Function/Sub Name           : sub_CmpExcelCompareFiles()
'Overview                    : �G�N�Z���t�@�C�����r����
'Detailed Description        : �G���[�͖�������
'Argument
'     aoParams               : �p�����[�^�i�[�p�ėp�I�u�W�F�N�g
'Return Value
'     �Ȃ�
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2017/04/26         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub sub_CmpExcelCompareFiles( _
    byRef aoParams _
    )
    '�p�����[�^�i�[�p�ėp�I�u�W�F�N�g
    Dim oParam : Set oParam = aoParams.Item("Param")
    
    '�t�@�C���̍ŏI�X�V�������ɕ��בւ���
    oParam.sortUsing new_Func("(c,n)=>new_CalAt(new_FileOf(c).DateLastModified).compareTo(new_CalAt(new_FileOf(n).DateLastModified))>0")
    '�����O�o��
    sub_CmpExcelLogger Array(3, "sub_CmpExcelCompareFiles", "aoParams sorted.")
    sub_CmpExcelLogger Array(9, "sub_CmpExcelCompareFiles", "aoParams is " & cf_toString(aoParams))
    
    '��r
    With New clsCompareExcel
        Set .broker = PoBroker
        .pathFrom = oParam(0)
        .pathTo = oParam(1)
        .compare()
    End With
    
    '�I�u�W�F�N�g���J��
    Set oParam = Nothing
End Sub

'***************************************************************************************************
'Processing Order            : -
'Function/Sub Name           : sub_CmpExcelLogger()
'Overview                    : ���O�o�͂���
'Detailed Description        : fw_logger()�ɈϏ�����
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
Private Sub sub_CmpExcelLogger( _
    byRef avParams _
    )
    fw_logger avParams, PoWriter
End Sub
