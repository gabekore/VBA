VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   10125
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13185
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
'' �{�^���͓��I�ɒǉ����Ă���
Dim indexArrayButton As Integer
Dim arrayButton() As Class1


''**************************************************************
'' UserForm�̃N���b�N�C�x���g
'' ���N���b�N����ƃ{�^����1���I�ɒǉ������
''**************************************************************
Private Sub UserForm_Click()
    ' �V���ɒǉ�����{�^��
    Dim newButton As MSForms.CommandButton
    
    ' �{�^���z��iarrayButton�j�̌��݂̃C���f�b�N�X�����֐i�߂�
    indexArrayButton = indexArrayButton + 1
    
    '---------------------------------------------------------------
    ' ���[�U�[�t�H�[���ɃR�}���h�{�^����ǉ�����
    '---------------------------------------------------------------
    ' �܂��͒ǉ�����{�^�����쐬
    Set newButton = UserForm1.Controls.Add("Forms.CommandButton.1", , True)
    ' �L���v�V�����Ƃ����[�U�[�t�H�[���ł̒ǉ��ʒu�Ƃ��w��
    With newButton
        .Caption = "�ǉ������{�^��" & indexArrayButton
        .Top = 10 + (indexArrayButton - 1) * 25
        .Left = 10
        .Height = 20
        .Width = 150
    End With
    
    ' �{�^���z��iarrayButton�j�̃T�C�Y����蒼��
    ' ��Preserve�w�肵�Ă�̂Ŋ����̃f�[�^�͕ێ������
    ReDim Preserve arrayButton(1 To indexArrayButton)
    
    ' �z�񐔂�1�����������ɂȂ�
    ' ���̑�����1�͋���ۂȂ̂ŐV����Class1�̃C���X�^���X�����Ă���
    Set arrayButton(indexArrayButton) = New Class1
    
    ' ���̑������V����Class1�̃C���X�^���X�ɃR���g���[���z��Ƃ���{�^����ǉ�
    Call arrayButton(indexArrayButton).RegistButton(newButton, indexArrayButton)
    
End Sub
