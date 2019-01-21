VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_Fax"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public �O���� As String
Public �㕶�� As String
Public strCode As String
Public ���̗� As Long
Public Fax�� As Long

'
' �@�\      : �j�d�x���o�͂���FAXno.�܂ޕ������Ԃ�
'
' ������    :
' �Ԃ�l    : �����e�`�w�p������
' �@�\����  : �������e�`�w�m���D�����������蕶����ŋ��ށi�n�C�t�����������j
'�@�@�@�@�@ :
' ���l      : 2018/11/21
Public Function Get_NumStr() As String
    Dim Fax_Num As String

    Fax_Num = ""
    If get����(Fax��) <> "#N/a" Then
        Fax_Num = get����(Fax��)  'mf�ɓ����Ă���e�`�w�m�n�D
        Fax_Num = StrConv(Fax_Num, vbNarrow)     '���p��
        Fax_Num = ReplaceX(Fax_Num, " -", "")    '�󔒂ƃn�C�t�����폜
        Get_NumStr = �O���� & Fax_Num & �㕶��
    Else
        Get_NumStr = ""
    End If
End Function

' �@�\      : �t�@�C�����Ɏg�p���闪�̖����擾
'
' ������    :
' �Ԃ�l    : ���̖�
' �@�\����  :
'
' ���l      : �t�@�C�����Ɏg�p�ł��Ȃ���������菜��
'
Public Function Get_����() As String
    Get_���� = get����(���̗�)
    Get_���� = ReplaceX(Get_����, "\/:*?<>|" & Chr(34), "")
End Function

'
' �@�\      : �j�d�x���o�͂���t�H���_����Ԃ�
'
' ������    :
' �Ԃ�l    : �r�o�p�t�H���_���@���́@�G���[("#N/a")
' �@�\����  : �t�H���_���𔻒f�o�^�O�ł����other��Ԃ�
'�@�@�@�@�@ :
' ���l      : 2018/11/21
Public Function Get_f_m() As String

    Get_f_m = ""
    
    If get����(Fax��) <> "#N/a" Then '�݂�������
        Get_f_m = Fol_Name(get����(Fax��))
    Else  '������Ȃ�
        Get_f_m = "#N/a"
    End If
    If Get_f_m = "" Then  '�Ή��t�H���_�����w��
        Get_f_m = "other"
    End If
End Function

'
' �@�\      : �w��L�[���
'
' ������    : lng�� - �擾�����
' �Ԃ�l    : ���������s��lng��̃f�[�^,�@������Ȃ��ꍇ- #N/a
' �@�\����  :
'
' ���l      :
'
Private Function get����(lng�� As Long) As String
    Dim Rng As Range
    Set Rng = Sheets("mf").Columns(1).find(what:=strCode)
    If Rng Is Nothing Then
        get���� = "#N/a"
    Else
        get���� = Sheets("mf").Cells(Rng.Row, lng��)
    End If
    Set Rng = Nothing
End Function

' �@�\      : �Ή��t�H���_����Ԃ�
'
' ������    : strFaxNum - FAXNO.
' �Ԃ�l    : �r�o�p�t�H���_,������Ȃ��ꍇ�͋�
' �@�\����  : �e�`�wNO.����n�C�t�����������A�U�蕪���p�V�[�g
'�@�@�@�@�@ : ����Ή�����t�H���_����������
' ���l      : 2018/11/21
'
Private Function Fol_Name(strFaxNum As String) As String
    Dim strHead As String

    Fol_Name = ""
    If strFaxNum <> "#N/a" Then
        strFaxNum = StrConv(strFaxNum, vbNarrow)     '���p��
        strFaxNum = ReplaceX(strFaxNum, " -", "")    '�󔒂ƃn�C�t�����폜
        I = 3
        Do Until Sheets("�U�蕪���p").Cells(I, 2) = Empty
            strHead = Sheets("�U�蕪���p").Cells(I, 2)
            If strHead = Left(strFaxNum, Len(strHead)) Then
                Fol_Name = Sheets("�U�蕪���p").Cells(I, 3)
                Exit Do
            End If
            I = I + 1
        Loop
    Else
        Fol_Name = "#N/a"
    End If
End Function

' �@�\      :������u��
'
' ������    : expression - �Ώە�����
'           : find - �u�����镶���� /������P�ʂł͂Ȃ��܂܂��P�����Â�
'           : rep - �u����̕���
' �Ԃ�l    : �u����̕�����
' �@�\����  :
'
' ���l      : 2018/11/21
'
Private Function ReplaceX(expression As String, find As String, rep As String) As String
    Dim I As Integer

    I = 1
    ReplaceX = expression
    Do Until I > Len(find)
        ReplaceX = Replace(ReplaceX, Mid(find, I, 1), rep)
        I = I + 1
    Loop
End Function