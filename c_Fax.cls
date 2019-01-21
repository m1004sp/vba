VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_Fax"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public 前文字 As String
Public 後文字 As String
Public strCode As String
Public 略称列 As Long
Public Fax列 As Long

'
' 機能      : ＫＥＹより出力するFAXno.含む文字列を返す
'
' 引き数    :
' 返り値    : 自動ＦＡＸ用文字列
' 機能説明  : ｍｆよりＦＡＸＮｏ．を検索し特定文字列で挟む（ハイフン等を除く）
'　　　　　 :
' 備考      : 2018/11/21
Public Function Get_NumStr() As String
    Dim Fax_Num As String

    Fax_Num = ""
    If get項目(Fax列) <> "#N/a" Then
        Fax_Num = get項目(Fax列)  'mfに入っているＦＡＸＮＯ．
        Fax_Num = StrConv(Fax_Num, vbNarrow)     '半角へ
        Fax_Num = ReplaceX(Fax_Num, " -", "")    '空白とハイフンを削除
        Get_NumStr = 前文字 & Fax_Num & 後文字
    Else
        Get_NumStr = ""
    End If
End Function

' 機能      : ファイル名に使用する略称名を取得
'
' 引き数    :
' 返り値    : 略称名
' 機能説明  :
'
' 備考      : ファイル名に使用できない文字を取り除く
'
Public Function Get_略称() As String
    Get_略称 = get項目(略称列)
    Get_略称 = ReplaceX(Get_略称, "\/:*?<>|" & Chr(34), "")
End Function

'
' 機能      : ＫＥＹより出力するフォルダ名を返す
'
' 引き数    :
' 返り値    : 排出用フォルダ名　又は　エラー("#N/a")
' 機能説明  : フォルダ名を判断登録外であればotherを返す
'　　　　　 :
' 備考      : 2018/11/21
Public Function Get_f_m() As String

    Get_f_m = ""
    
    If get項目(Fax列) <> "#N/a" Then 'みつかったら
        Get_f_m = Fol_Name(get項目(Fax列))
    Else  '見つからない
        Get_f_m = "#N/a"
    End If
    If Get_f_m = "" Then  '対応フォルダが無指定
        Get_f_m = "other"
    End If
End Function

'
' 機能      : 指定キーより
'
' 引き数    : lng列 - 取得する列
' 返り値    : 見つかった行のlng列のデータ,　見つからない場合- #N/a
' 機能説明  :
'
' 備考      :
'
Private Function get項目(lng列 As Long) As String
    Dim Rng As Range
    Set Rng = Sheets("mf").Columns(1).find(what:=strCode)
    If Rng Is Nothing Then
        get項目 = "#N/a"
    Else
        get項目 = Sheets("mf").Cells(Rng.Row, lng列)
    End If
    Set Rng = Nothing
End Function

' 機能      : 対応フォルダ名を返す
'
' 引き数    : strFaxNum - FAXNO.
' 返り値    : 排出用フォルダ,見つからない場合は空白
' 機能説明  : ＦＡＸNO.からハイフン等を除き、振り分け用シート
'　　　　　 : から対応するフォルダを検索する
' 備考      : 2018/11/21
'
Private Function Fol_Name(strFaxNum As String) As String
    Dim strHead As String

    Fol_Name = ""
    If strFaxNum <> "#N/a" Then
        strFaxNum = StrConv(strFaxNum, vbNarrow)     '半角へ
        strFaxNum = ReplaceX(strFaxNum, " -", "")    '空白とハイフンを削除
        I = 3
        Do Until Sheets("振り分け用").Cells(I, 2) = Empty
            strHead = Sheets("振り分け用").Cells(I, 2)
            If strHead = Left(strFaxNum, Len(strHead)) Then
                Fol_Name = Sheets("振り分け用").Cells(I, 3)
                Exit Do
            End If
            I = I + 1
        Loop
    Else
        Fol_Name = "#N/a"
    End If
End Function

' 機能      :文字列置換
'
' 引き数    : expression - 対象文字列
'           : find - 置換する文字列 /文字列単位ではなく含まれる１文字づつ
'           : rep - 置換後の文字
' 返り値    : 置換後の文字列
' 機能説明  :
'
' 備考      : 2018/11/21
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
