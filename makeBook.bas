Attribute VB_Name = "makeBook"

Const ���� = 2
Const sht�ݒ� = "Sheet1"
Const �v���� = 7
Dim �v������(2, �v����) As String

'
' �@�\      : �����e�`�w�p�ɂe�`�w�m�n�D���H�ƕۑ��t�H���_���w�肵�ۑ�������
'
' ������    :
' �Ԃ�l    :
' �@�\����  :
Sub mf()
    Dim fn As String
    Dim strFile As String
    Dim I As Integer
    Dim sht�o�� As String
    Dim Fax As c_Fax

    Set Fax = New c_Fax

    AppOnOff = False
    Call �v���ݒ�

    With Sheets(sht�ݒ�)
        ''<--�R�[�h����t�@�b�N�X�p�����Z����
        Fax.strCode = .Cells(2, 2)
        Fax.Fax�� = .Cells(11, 2)
        Fax.�O���� = .Cells(13, 2)
        Fax.�㕶�� = .Cells(14, 2)
        Fax.���̗� = .Cells(12, 2)
        .Cells(4, 2) = Fax.Get_f_m
        .Cells(6, 2) = Fax.Get_����
        .Cells(7, 2) = Fax.Get_NumStr
        Set Fax = Nothing
        ''-->

        sht�o�� = .Cells(10, 2)        '�Z���̓��e��ǂݍ���
        fn = .Cells(6, 2) & "_" & Format(Now, "yyyymmddhhnnss") & ".xlsx"

        '�����ň��於�Ƃ����ݒ� ����
        fn = .Cells(4, 2) & "_" & fn
        strFile = .Cells(3, 2) & fn
        Set Newbook = Workbooks.Add                  '�V�u�b�N�쐬
        Newbook.SaveAs Filename:=strFile             '���O��t���ĕۑ�
        ThisWorkbook.Activate
        Sheets(sht�o��).Cells(1, 1).Value = .Cells(7, 2) '�t�@�b�N�X�m���D������
    End With
    With Workbooks(fn)
        cnt = .Sheets.Count                     '�V�����u�b�N�쐬���I�v�V�����ɂ���ăV�[�g�����قȂ�׃V�[�g�𐔂���
        Sheets(sht�o��).Copy After:=.Sheets(cnt) '�ΏۃV�[�g��V�t�@�C���ɃR�s�[
        '�R�s�[�����V�[�g�ȊO���폜�@�����V�[�g������ꍇ�̎w�肪������Ȃ��̂ŗ]�v�ȕ���r��
        For I = 1 To cnt
            .Sheets(I).Delete
        Next
        '�u�b�N�̃v���p�e�B��ݒ肷��
        For I = 1 To �v����
            .BuiltinDocumentProperties(�v������(1, I)).Value = �v������(2, I)
        Next
        .Save                           '�㏑���ۑ�
        .Close                          '�u�b�N�����
    End With

    '�x���\���E��ʕ`��E�����Čv�Z��߂�
    AppOnOff = True
    
End Sub

'2018/11/29
'�u�b�N�̃v���p�e�B�Ή�
Private Sub �v���ݒ�()
    Dim I As Integer
    ThisWorkbook.Activate
    With Sheets(sht�ݒ�)
        For I = 1 To �v����
            �v������(1, I) = .Cells(I + 15, 5)
            �v������(2, I) = .Cells(I + 15, 2)
        Next
    End With
End Sub

'2018/11/30
' �@�\      :�x���\���E��ʕ`��E�����Čv�Z�̃I���I�t
' ������    : blnOO - True ���� False
' �@�\����  : True�Ȃ�f�t�H���gFalse�Ȃ�S��~
Property Let AppOnOff(blnOO As Boolean)
    With Application
        .DisplayAlerts = blnOO
        .ScreenUpdating = blnOO
        .EnableEvents = blnOO '2018/12/6 add
        .Calculation = IIf(blnOO, xlCalculationAutomatic, xlCalculationManual)
        '�V�[�g�Ɋ֐��������t�@�C���̎��������x���S�R�Ⴄ
    End With
End Property
