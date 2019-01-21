Attribute VB_Name = "makeBook"

Const 略称 = 2
Const sht設定 = "Sheet1"
Const プロ数 = 7
Dim プロ項目(2, プロ数) As String

'
' 機能      : 自動ＦＡＸ用にＦＡＸＮＯ．加工と保存フォルダを指定し保存させる
'
' 引き数    :
' 返り値    :
' 機能説明  :
Sub mf()
    Dim fn As String
    Dim strFile As String
    Dim I As Integer
    Dim sht出力 As String
    Dim Fax As c_Fax

    Set Fax = New c_Fax

    AppOnOff = False
    Call プロ設定

    With Sheets(sht設定)
        ''<--コードからファックス用情報をセルへ
        Fax.strCode = .Cells(2, 2)
        Fax.Fax列 = .Cells(11, 2)
        Fax.前文字 = .Cells(13, 2)
        Fax.後文字 = .Cells(14, 2)
        Fax.略称列 = .Cells(12, 2)
        .Cells(4, 2) = Fax.Get_f_m
        .Cells(6, 2) = Fax.Get_略称
        .Cells(7, 2) = Fax.Get_NumStr
        Set Fax = Nothing
        ''-->

        sht出力 = .Cells(10, 2)        'セルの内容を読み込む
        fn = .Cells(6, 2) & "_" & Format(Now, "yyyymmddhhnnss") & ".xlsx"

        'ここで宛先名とかも設定 割愛
        fn = .Cells(4, 2) & "_" & fn
        strFile = .Cells(3, 2) & fn
        Set Newbook = Workbooks.Add                  '新ブック作成
        Newbook.SaveAs Filename:=strFile             '名前を付けて保存
        ThisWorkbook.Activate
        Sheets(sht出力).Cells(1, 1).Value = .Cells(7, 2) 'ファックスＮｏ．を入れる
    End With
    With Workbooks(fn)
        cnt = .Sheets.Count                     '新しいブック作成時オプションによってシート数が異なる為シートを数える
        Sheets(sht出力).Copy After:=.Sheets(cnt) '対象シートを新ファイルにコピー
        'コピーしたシート以外を削除　複数シートがある場合の指定が分からないので余計な物を排除
        For I = 1 To cnt
            .Sheets(I).Delete
        Next
        'ブックのプロパティを設定する
        For I = 1 To プロ数
            .BuiltinDocumentProperties(プロ項目(1, I)).Value = プロ項目(2, I)
        Next
        .Save                           '上書き保存
        .Close                          'ブックを閉じる
    End With

    '警告表示・画面描画・自動再計算を戻す
    AppOnOff = True
    
End Sub

'2018/11/29
'ブックのプロパティ対応
Private Sub プロ設定()
    Dim I As Integer
    ThisWorkbook.Activate
    With Sheets(sht設定)
        For I = 1 To プロ数
            プロ項目(1, I) = .Cells(I + 15, 5)
            プロ項目(2, I) = .Cells(I + 15, 2)
        Next
    End With
End Sub

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
        'シートに関数が多いファイルの時処理速度が全然違う
    End With
End Property
