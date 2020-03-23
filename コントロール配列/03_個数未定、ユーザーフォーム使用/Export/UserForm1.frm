VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   10125
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13185
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
'' ボタンは動的に追加していく
Dim indexArrayButton As Integer
Dim arrayButton() As Class1


''**************************************************************
'' UserFormのクリックイベント
'' ※クリックするとボタンが1個動的に追加される
''**************************************************************
Private Sub UserForm_Click()
    ' 新たに追加するボタン
    Dim newButton As MSForms.CommandButton
    
    ' ボタン配列（arrayButton）の現在のインデックスを次へ進める
    indexArrayButton = indexArrayButton + 1
    
    '---------------------------------------------------------------
    ' ユーザーフォームにコマンドボタンを追加する
    '---------------------------------------------------------------
    ' まずは追加するボタンを作成
    Set newButton = UserForm1.Controls.Add("Forms.CommandButton.1", , True)
    ' キャプションとかユーザーフォームでの追加位置とか指定
    With newButton
        .Caption = "追加したボタン" & indexArrayButton
        .Top = 10 + (indexArrayButton - 1) * 25
        .Left = 10
        .Height = 20
        .Width = 150
    End With
    
    ' ボタン配列（arrayButton）のサイズを作り直し
    ' ※Preserve指定してるので既存のデータは保持される
    ReDim Preserve arrayButton(1 To indexArrayButton)
    
    ' 配列数が1個増えただけになる
    ' その増えた1個は空っぽなので新たなClass1のインスタンスを入れておく
    Set arrayButton(indexArrayButton) = New Class1
    
    ' その増えた新たなClass1のインスタンスにコントロール配列とするボタンを追加
    Call arrayButton(indexArrayButton).RegistButton(newButton, indexArrayButton)
    
End Sub
