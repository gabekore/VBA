VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''**************************************************************
''
'' コントロール配列を実現するクラス（ComandButtonのみ対応）
''
''   登録したコマンドボタンのイベントハンドラがここにある
''
''**************************************************************
Option Explicit

''**************************************************************
'' このクラスのプロパティ
''**************************************************************
' このクラスに登録したボタン
Private WithEvents TargetBtn As MSForms.CommandButton
Attribute TargetBtn.VB_VarHelpID = -1
' このクラスに登録したボタンの番号（どのボタンをクリックしたのか知るために必要）
Private TargetIndex As Integer


''**************************************************************
'' ボタンの登録
''**************************************************************
Public Sub RegistButton(regButton As MSForms.CommandButton, Index As Integer)
    Set TargetBtn = regButton
    TargetIndex = Index
End Sub

''**************************************************************
'' クリックイベントのハンドラ
'' 登録したボタン全てがこのイベントハンドラを使う
''**************************************************************
Private Sub TargetBtn_Click()
    Call MsgBox("ボタン=" & TargetBtn.Caption & "：index=" & TargetIndex)
End Sub
