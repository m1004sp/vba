Attribute VB_Name = "makeDW"
'■xdwapi.dllの保存場所注意
'APIの終了処理を行う。
Public Declare Function XDW_Finalize Lib "G:\docuworks_dev\dwsdk730jpn\XDWAPI\DLL\xdwapi.dll" (ByVal reserved As String) As Long
''DocuWorks文書に変換
Public Declare Function XDW_BeginCreationFromAppFile Lib "G:\docuworks_dev\dwsdk730jpn\XDWAPI\DLL\xdwapi.dll" (ByVal lpszInputPath As String, ByVal lpszOutputPath As String, ByVal bWithOrg As Boolean, ByRef pHandle As Long, ByVal reserved As String) As Long
Public Declare Function XDW_EndCreationFromAppFile Lib "G:\docuworks_dev\dwsdk730jpn\XDWAPI\DLL\xdwapi.dll" (ByVal handle As Long, ByVal reserved As String) As Long

'2018/12/13
Sub Excel2DW()
    Dim fol As String  'エクセルが入っているフォルダ
    Dim sfol As String  'xdwが振り分けられるフォルダ
    Dim f As String
    Dim e_fol As String
    Dim strFinFile As String '変換済を保存する為
    Dim strInFile As String '入力ファイル名:エクセル
    Dim strOutFile As String '出力ファイル名:DocuWorks
    Dim lngHandle As Long 'ハンドル:不明

    AppOnOff = False
    fol = Sheets("Sheet1").Cells(3, 2)  '''エクセルが入っているフォルダ
    e_fol = Sheets("Sheet1").Cells(25, 2)
    f = Dir(fol & "*.xlsx")
    Do While f <> ""
        sfol = Mid(f, 1, InStr(f, "_") - 1) & "\" 'エクセルのファイル名の最初の_までがフォルダ名になっている
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
' 機能      :警告表示・画面描画・自動再計算のオンオフ
' 引き数    : blnOO - True 又は False
' 機能説明  : TrueならデフォルトFalseなら全停止
Property Let AppOnOff(blnOO As Boolean)
    With Application
        .DisplayAlerts = blnOO
        .ScreenUpdating = blnOO
        .EnableEvents = blnOO '2018/12/6 add
        .Calculation = IIf(blnOO, xlCalculationAutomatic, xlCalculationManual)
        .Visible = blnOO
        'シートに関数が多いファイルの時処理速度が全然違う
    End With
End Property

