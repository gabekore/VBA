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
Private WithEvents TargetBtn As CommandButton
Attribute TargetBtn.VB_VarHelpID = -1
' このクラスに登録したボタンの番号（どのボタンをクリックしたのか知るために必要）
Private TargetIndex As Integer
' ↓このオブジェクトにはラベルを期待している
Private objOLE As OLEObject

''**************************************************************
'' ボタンの登録
''   regButton：コントロール配列対象のボタン
''   Index：コントロール配列対象のボタンを識別するための番号
''   ole：クリックイベントで使うオブジェクト（今回はラベルを期待している）
''**************************************************************
Public Sub RegistButton(regButton As CommandButton, Index As Integer, ole As OLEObject)
    Set TargetBtn = regButton
    TargetIndex = Index
''    Set lblDisp = Label1
    Set objOLE = ole
End Sub

''**************************************************************
'' クリックイベントのハンドラ
'' 登録したボタン全てがこのイベントハンドラを使う
''**************************************************************
Private Sub TargetBtn_Click()
    ' ラベルにクリックしたバタンのキャプションを表示する
    objOLE.Object.Caption = "ボタン=" & TargetBtn.Caption & "：index=" & TargetIndex
End Sub

