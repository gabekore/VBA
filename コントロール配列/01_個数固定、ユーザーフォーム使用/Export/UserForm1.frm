VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   2430
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6045
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'' �R�}���h�{�^����o�^����N���X�̔z��
'' ����͂R�Œ�
Private MyBtnArray(1 To 3) As New Class1


''**************************************************************
'' UserForm��Initialize�C�x���g
''**************************************************************
Private Sub UserForm_Initialize()
    Dim i As Integer
    
    ' �{�^���̌������[�v���āAMyBtnArray�z��i�N���X�j�Ƀ{�^����o�^����
    For i = LBound(MyBtnArray) To UBound(MyBtnArray)
        Call MyBtnArray(i).RegistButton(UserForm1.Controls("CommandButton" & i), i)
    Next i

End Sub
