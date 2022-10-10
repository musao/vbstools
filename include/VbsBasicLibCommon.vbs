'***************************************************************************************************
'FILENAME                    :VbsBasicLibCommon.vbs
'Generato                    :2022/09/27
'Descrition                  :���ʋ@�\
' �p�����[�^�i�����j:
'     PATH         :�t�@�C���̃p�X
'---------------------------------------------------------------------------------------------------
'Modification Histroy
'
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/09/27         EXA Y.Fujii              Initial Release
'***************************************************************************************************

'�I�t�B�X�S��

'�����̕ی����������
Private Sub sub_CM_OfficeUnprotect( _
    byRef aoOffice _
    , byVal asPassword _
    )
    On Error Resume Next
    aoOffice.Unprotect(asPassword)
    If Err.Number Then
        Err.Clear
    End If
End Sub

'�G�N�Z���n

'�G�N�Z���t�@�C����ʖ��ŕۑ����ĕ���
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

'�G�N�Z���t�@�C�����J���āi�ǂݎ���p�^�_�C�A���O�Ȃ��j���[�N�u�b�N�I�u�W�F�N�g��Ԃ�
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

'�G�N�Z���̃I�[�g�V�F�C�v�̃e�L�X�g�����o��
Private Function func_CM_ExcelGetTextFromAutoshape( _
    byRef aoAutoshape _
    )
    On Error Resume Next
    func_CM_ExcelGetTextFromAutoshape = aoAutoshape.TextFrame.Characters.Text
    If Err.Number Then
        Err.Clear
    End If
End Function

'�t�@�C������n

'�t�@�C�����폜����
Private Function func_CM_DeleteFile( _
    byVal asPath _
    ) 
    On Error Resume Next
    CreateObject("Scripting.FileSystemObject").DeleteFile(asPath)
    func_CM_DeleteFile = True
    If Err.Number Then
        func_CM_DeleteFile = False
    End If
End Function

'�e�t�H���_�p�X�̎擾
Private Function func_CM_GetParentFolderPath( _
    byVal asPath _
    ) 
    func_CM_GetParentFolderPath = CreateObject("Scripting.FileSystemObject").GetParentFolderName(asPath)
End Function

'�t�@�C���p�X�̌����̍쐬
Private Function func_CM_BuildPath( _
    byVal asFolderPath _
    , byVal asItemName _
    ) 
    func_CM_BuildPath = CreateObject("Scripting.FileSystemObject").BuildPath(asFolderPath, asItemName)
End Function

'�t�@�C���̑��݊m�F
Private Function func_CM_FileExists( _
    byVal asPath _
    ) 
    func_CM_FileExists = CreateObject("Scripting.FileSystemObject").FileExists(asPath)
End Function

'�t�@�C���I�u�W�F�N�g�̎擾
Private Function func_CM_GetFile( _
    byVal asPath _
    ) 
    Set func_CM_GetFile = CreateObject("Scripting.FileSystemObject").GetFile(asPath)
End Function

'�ꎞ�t�@�C�����̍쐬
Private Function func_CM_GetTempFileName()
    func_CM_GetTempFileName = CreateObject("Scripting.FileSystemObject").GetTempName()
End Function


'���

'Min�֐�
Private Function func_CM_Min( _
    byVal al1 _ 
    , byVal al2 _
    )
    Dim lReturnValue
    If al1 < al2 Then
        lReturnValue = al1
    Else
        lReturnValue = al2
    End If
    func_CM_Min = lReturnValue
End Function

'Max�֐�
Private Function func_CM_Max( _
    byVal al1 _ 
    , byVal al2 _
    )
    Dim lReturnValue
    If al1 > al2 Then
        lReturnValue = al1
    Else
        lReturnValue = al2
    End If
    func_CM_Max = lReturnValue
End Function


'���ꉽ�n����

'�R���N�V��������w�肵�����O�̃����o�[���擾����
Private Function func_CM_GetObjectByIdFromCollection( _
    byRef aoClloection _
    , byVal asId _
    )
    On Error Resume Next
    Dim oItem
    For Each oItem In aoClloection
        If oItem.Id = asId Then
            Set func_CM_GetObjectByIdFromCollection = oItem
            Exit Function
        End If
    Next
    Set func_CM_GetObjectByIdFromCollection = Nothing
    If Err.Number Then
        Err.Clear
    End If
    Set oItem = Nothing
End Function
