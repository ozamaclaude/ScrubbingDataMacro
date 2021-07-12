VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "�f�[�^�N�����W���O"
   ClientHeight    =   4770
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9465.001
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
' �Ƃ���{�^���������̃C�x���g����
'******************************************************
Private Sub CloseButton_Click()
    Unload UserForm1
End Sub

'******************************************************
' ���s�{�^���������̃C�x���g����
'******************************************************
Private Sub ExecBtn_Click()
    ' �I���W�i���t�@�C���̓Ǎ�
    Call AddSheet(True)
    ' �ϊ���̓Ǎ�
    Call AddSheet(False)
    ' Excel��\��
    ActivateApp (True)
End Sub

'******************************************************
' ���s�{�^���̊����^�񊈐���ݒ�
'******************************************************
Sub ActivateExecuteButton(sts As Boolean)
    If sts <> "False" Then
        UserForm1.Controls("ExecBtn").Enabled = True
        Exit Sub
    End If
    UserForm1.Controls("ExecBtn").Enabled = False
End Sub

'******************************************************
' ���̓t�@�C���̃`�F�b�N�{�b�N�X(�J���})�I�����̃C�x���g
'******************************************************
Private Sub InputCheckBoxComma_Click()
    If IsControlChecked("InputCheckBoxComma") = True Then
        Call DoCheckControl("InputCheckBoxPipe", False)
        Call DoCheckControl("InputCheckBoxSpace", False)
        Exit Sub
    End If
End Sub

'******************************************************
' ���̓t�@�C���̃`�F�b�N�{�b�N�X(�p�C�v)�I�����̃C�x���g
'******************************************************
Private Sub InputCheckBoxPipe_Click()
    If IsControlChecked("InputCheckBoxPipe") = True Then
        Call DoCheckControl("InputCheckBoxComma", False)
        Call DoCheckControl("InputCheckBoxSpace", False)
        Exit Sub
    End If
End Sub

'******************************************************
' ���̓t�@�C���̃`�F�b�N�{�b�N�X(�X�y�[�X)�I�����̃C�x���g
'******************************************************
Private Sub InputCheckBoxSpace_Click()
    If IsControlChecked("InputCheckBoxSpace") = True Then
        Call DoCheckControl("InputCheckBoxComma", False)
        Call DoCheckControl("InputCheckBoxPipe", False)
        Exit Sub
    End If
End Sub

'******************************************************
' �o�̓t�@�C���̃`�F�b�N�{�b�N�X(�J���})�I�����̃C�x���g
'******************************************************
Private Sub OutputCheckBoxComma_Click()
    If IsControlChecked("OutputCheckBoxComma") = True Then
        Call DoCheckControl("OutputCheckBoxSpace", False)
        Call DoCheckControl("OutputCheckBoxPipe", False)
        Exit Sub
    End If
End Sub

'******************************************************
' �o�̓t�@�C���̃`�F�b�N�{�b�N�X(�p�C�v)�I�����̃C�x���g
'******************************************************
Private Sub OutputCheckBoxPipe_Click()
    If IsControlChecked("OutputCheckBoxPipe") = True Then
        Call DoCheckControl("OutputCheckBoxComma", False)
        Call DoCheckControl("OutputCheckBoxSpace", False)
        Exit Sub
    End If
End Sub

'******************************************************
' �o�̓t�@�C���̃`�F�b�N�{�b�N�X(�X�y�[�X)�I�����̃C�x���g
'******************************************************
Private Sub OutputCheckBoxSpace_Click()
    If IsControlChecked("OutputCheckBoxSpace") = True Then
        Call DoCheckControl("OutputCheckBoxComma", False)
        Call DoCheckControl("OutputCheckBoxPipe", False)
        Exit Sub
    End If
End Sub

'******************************************************
' �Q�ƃ{�^���������̃C�x���g����
'******************************************************
Private Sub RefferenceBtn1_Click()
    ' �J�����g�t�H���_�ݒ�
    CreateObject("WScript.Shell").CurrentDirectory = ReferenceFolderPath
    OpenFileName = Application.GetOpenFilename()
    If OpenFileName <> "False" Then
        TargetFilePath.Value = Dir(OpenFileName)
        ReadFilePath = OpenFileName
        ActivateExecuteButton (True)
    End If
End Sub
