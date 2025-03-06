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

'lib\com import
Dim sRelativeFolderName : sRelativeFolderName = "lib\com"
With CreateObject("Scripting.FileSystemObject")
    Dim sParentFolderPath : sParentFolderPath = .GetParentFolderName(WScript.ScriptFullName)
    Dim sLibFolderPath : sLibFolderPath = .BuildPath(sParentFolderPath, sRelativeFolderName)
    Dim oLibFile
    For Each oLibFile In CreateObject("Shell.Application").Namespace(sLibFolderPath).Items
        If Not oLibFile.IsFolder Then
            If StrComp(.GetExtensionName(oLibFile.Path), "vbs", vbTextCompare)=0 Then ExecuteGlobal .OpenTextfile(oLibFile.Path).ReadAll
        End If
    Next
End With
Set oLibFile = Nothing
'lib import
sRelativeFolderName = "lib"
With new_FSO()
    sLibFolderPath = .BuildPath(sParentFolderPath, sRelativeFolderName)
    ExecuteGlobal .OpenTextfile(.BuildPath(sLibFolderPath,"libEnum.vbs")).ReadAll
    ExecuteGlobal .OpenTextfile(.BuildPath(sLibFolderPath,"clsCompareExcel.vbs")).ReadAll
End With


'���O�o�͐�A�u���[�J�[�N���X�̃C���X�^���X�̐ݒ�
Private PoTs4Log, PoBroker
Set PoTs4Log = fw_getTextstreamForLog()
Set PoBroker = new_BrokerOf(Array(topic.LOG, GetRef("this_logger")))

'Main�֐����s
Call Main()

'�I������
PoTs4Log.close()
Set PoBroker = Nothing : Set PoTs4Log = Nothing
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
    '�p�����[�^�i�[�p�ėp�I�u�W�F�N�g�錾
    Dim oParams : Set oParams = new_Dic()
    
    '���X�N���v�g�̈����擾
    fw_excuteSub "this_getParameters", oParams, PoBroker
    
    '��r�Ώۃt�@�C�����͉�ʂ̕\���Ǝ擾
    fw_excuteSub "this_dispInputFiles", oParams, PoBroker
    
    '�G�N�Z���t�@�C�����r����
    fw_excuteSub "this_compareFiles", oParams, PoBroker
    
    '�I�u�W�F�N�g���J��
    Set oParams = Nothing
End Sub

'***************************************************************************************************
'Processing Order            : 1
'Function/Sub Name           : this_getParameters()
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
Private Sub this_getParameters( _
    byRef aoParams _
    )
    '�I���W�i���̈������擾
    Dim oArg : Set oArg = fw_storeArguments()
    '�����O�o��
    this_logger Array(logType.DETAIL, "this_getParameters()", cf_toString(oArg))
    
    '�p�����[�^�i�[�p�I�u�W�F�N�g�ɐݒ�
    cf_bindAt aoParams, "Param", new_ArrOf(oArg.Item("Unnamed")).slice(0,2)
    
    Set oArg = Nothing
End Sub

'***************************************************************************************************
'Processing Order            : 2
'Function/Sub Name           : this_dispInputFiles()
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
Private Sub this_dispInputFiles( _
    byRef aoParams _
    )
    Dim oParam : Set oParam = aoParams.Item("Param")
    If oParam.length > 1 Then
    '�p�����[�^��2�ȏゾ������֐��𔲂���
        '�����O�o��
        this_logger Array(logType.INFO, "this_dispInputFiles()", "No dialog required.")
        Exit Sub
    End If
    
    '�p�����[�^�i�[�p�ėp�I�u�W�F�N�g
    Const Cs_TITLE_EXCEL = "��r�Ώۃt�@�C�����J��"
    With CreateObject("Excel.Application")
        Dim sPath
        Do Until oParam.length > 1
            
            sPath = .GetOpenFilename( , , Cs_TITLE_EXCEL, , False)
            If sPath = False Then
            '�t�@�C���I���L�����Z���̏ꍇ�͓��X�N���v�g���I������
                .Quit

                '�����O�o��
                this_logger Array(logType.WARNING, "this_dispInputFiles()", "Dialog input canceled.")
                PoTs4Log.close
                
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
'Function/Sub Name           : this_compareFiles()
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
Private Sub this_compareFiles( _
    byRef aoParams _
    )
    '�p�����[�^�i�[�p�ėp�I�u�W�F�N�g
    Dim oParam : Set oParam = aoParams.Item("Param")
    
    '�t�@�C���̍ŏI�X�V�������ɕ��בւ���
    oParam.sortUsing new_Func("(c,n)=>new_CalAt(new_FileOf(c).DateLastModified).compareTo(new_CalAt(new_FileOf(n).DateLastModified))>0")
    '�����O�o��
    this_logger Array(logType.INFO, "this_compareFiles()", "aoParams sorted.")
    this_logger Array(logType.DETAIL, "this_compareFiles()", "aoParams is " & cf_toString(aoParams))
    
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
'Function/Sub Name           : this_logger()
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
Private Sub this_logger( _
    byRef avParams _
    )
    fw_logger avParams, PoTs4Log
End Sub
