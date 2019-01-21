Attribute VB_Name = "makeDW"
'��xdwapi.dll�̕ۑ��ꏊ����
'API�̏I���������s���B
Public Declare Function XDW_Finalize Lib "G:\docuworks_dev\dwsdk730jpn\XDWAPI\DLL\xdwapi.dll" (ByVal reserved As String) As Long
''DocuWorks�����ɕϊ�
Public Declare Function XDW_BeginCreationFromAppFile Lib "G:\docuworks_dev\dwsdk730jpn\XDWAPI\DLL\xdwapi.dll" (ByVal lpszInputPath As String, ByVal lpszOutputPath As String, ByVal bWithOrg As Boolean, ByRef pHandle As Long, ByVal reserved As String) As Long
Public Declare Function XDW_EndCreationFromAppFile Lib "G:\docuworks_dev\dwsdk730jpn\XDWAPI\DLL\xdwapi.dll" (ByVal handle As Long, ByVal reserved As String) As Long

'2018/12/13
Sub Excel2DW()
    Dim fol As String  '�G�N�Z���������Ă���t�H���_
    Dim sfol As String  'xdw���U�蕪������t�H���_
    Dim f As String
    Dim e_fol As String
    Dim strFinFile As String '�ϊ��ς�ۑ������
    Dim strInFile As String '���̓t�@�C����:�G�N�Z��
    Dim strOutFile As String '�o�̓t�@�C����:DocuWorks
    Dim lngHandle As Long '�n���h��:�s��

    AppOnOff = False
    fol = Sheets("Sheet1").Cells(3, 2)  '''�G�N�Z���������Ă���t�H���_
    e_fol = Sheets("Sheet1").Cells(25, 2)
    f = Dir(fol & "*.xlsx")
    Do While f <> ""
        sfol = Mid(f, 1, InStr(f, "_") - 1) & "\" '�G�N�Z���̃t�@�C�����̍ŏ���_�܂ł��t�H���_���ɂȂ��Ă���
        strInFile = fol & f
        strFinFile = e_fol & f
        f = Replace(f, Mid(f, 1, InStr(f, "_")), "")
        f = Replace(f, "xlsx", "xdw")
        strOutFile = fol & sfol & f
        If XDW_BeginCreationFromAppFile(strInFile, strOutFile, False, lngHandle, vbNullString) = 0 Then
            Debug.Print ima & strInFile & " : " & strOutFile
        Else
            Debug.Print ima & "Err : BeginCreationFromAppFile (" & strInFile & " : " & strOutFile & ")"
        End If
        If XDW_EndCreationFromAppFile(lngHandle, vbNullString) <> 0 Then
            Debug.Print ima & "Err : EndCreationFromAppFile"
        End If
        f = Dir()
        Name strInFile As strFinFile
    Loop
    Call XDW_Finalize(vbNullString)
    Debug.Print ima & "end"
    
    AppOnOff = True
End Sub

Private Function ima() As String
    ima = " [" & Now & "] "
End Function

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
        .Visible = blnOO
        '�V�[�g�Ɋ֐��������t�@�C���̎��������x���S�R�Ⴄ
    End With
End Property

