VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   2430
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6045
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'' コマンドボタンを登録するクラスの配列
'' 今回は３個固定
Private MyBtnArray(1 To 3) As New Class1


''**************************************************************
'' UserFormのInitializeイベント
''**************************************************************
Private Sub UserForm_Initialize()
    Dim i As Integer
    
    ' ボタンの個数分ループして、MyBtnArray配列（クラス）にボタンを登録する
    For i = LBound(MyBtnArray) To UBound(MyBtnArray)
        Call MyBtnArray(i).RegistButton(UserForm1.Controls("CommandButton" & i), i)
    Next i

End Sub
