Attribute VB_Name = "test"
'Sub フォルダ作成()
'    Dim I As Integer
'    Dim FolderName As String    '作成したいフォルダパスを格納'
'
'    Sheets("フォルダ・機器").Select
'    I = 2
'
'    Do Until Cells(I, 2) = Empty
'        FolderName = "C:\ttt\" & Cells(I, 1) & "_" & Cells(I, 2)
'        If Dir(FolderName, vbDirectory) = "" Then   '同名のフォルダがない場合フォルダを作成'
'            MkDir FolderName
'        End If
'        I = I + 1
'    Loop
'
'End Sub


Const sht設定 = "Sheet1"

Sub test()
    Dim Fax As c_Fax
    Set Fax = New c_Fax
    Fax.strCode = "250003"
    Fax.Fax列 = Sheets(sht設定).Cells(11, 2)
    Fax.前文字 = "FAX<"
    Fax.後文字 = ">"
    Fax.略称列 = 2
    Debug.Print Fax.Get_NumStr() & "::" & Fax.Get_f_m
    Debug.Print Fax.Get_略称
    Set Fax = Nothing
End Sub

Sub test2()
    Dim Ste As SteelC
    Set Ste = New SteelC

    Ste.cJuryo = 1750
    Ste.cNaik = 24
    Ste.Saizu = "11.5X1219.5x2438"
    Ste.Hiju = 7.85
    Ste.Mai = 10
    Ste.Hiju = 0
    Debug.Print Ste.sizeCut(2)

    Set Ste = Nothing
End Sub

Function FeWeight(Sai As String, Su As Integer, Optional Met As Double, Optional Hij As Double) As Double
    Dim Ste As SteelC
    Set Ste = New SteelC

    Ste.Saizu = Sai
    Ste.Mai = Su
    Ste.Hiju = 0
    Ste.Metsuke = Met
    Ste.Hiju = Hij
    FeWeight = Ste.Juryo

    Set Ste = Nothing
End Function

