VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

'' コマンドボタンを登録するクラスの配列
'' 今回は３個固定
Private MyBtnArray(1 To 3) As New Class1

Private hoge As label

Private Sub btnInitialize_Click()

    ' ラベルをクリアしておく
    Label1.Caption = ""

    ' MyBtnArray配列（クラス）にボタンを登録する
    ' UserFormならforで回してCall MyBtnArray(i).RegistButton(UserForm1.Controls("CommandButton" & i), i)を使えるんだけど、
    ' シート上の場合は無理っぽい
    Call MyBtnArray(1).RegistButton(CommandButton1, 1, OLEObjects("Label1"))
    Call MyBtnArray(2).RegistButton(CommandButton2, 2, OLEObjects("Label1"))
    Call MyBtnArray(3).RegistButton(CommandButton3, 3, OLEObjects("Label1"))
End Sub
