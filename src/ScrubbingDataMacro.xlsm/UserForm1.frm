VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "データクレンジング"
   ClientHeight    =   4770
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9465.001
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
' とじるボタン押下時のイベント操作
'******************************************************
Private Sub CloseButton_Click()
    Unload UserForm1
End Sub

'******************************************************
' 実行ボタン押下時のイベント操作
'******************************************************
Private Sub ExecBtn_Click()
    ' オリジナルファイルの読込
    Call AddSheet(True)
    ' 変換後の読込
    Call AddSheet(False)
    ' Excelを表示
    ActivateApp (True)
End Sub

'******************************************************
' 実行ボタンの活性／非活性を設定
'******************************************************
Sub ActivateExecuteButton(sts As Boolean)
    If sts <> "False" Then
        UserForm1.Controls("ExecBtn").Enabled = True
        Exit Sub
    End If
    UserForm1.Controls("ExecBtn").Enabled = False
End Sub

'******************************************************
' 入力ファイルのチェックボックス(カンマ)選択時のイベント
'******************************************************
Private Sub InputCheckBoxComma_Click()
    If IsControlChecked("InputCheckBoxComma") = True Then
        Call DoCheckControl("InputCheckBoxPipe", False)
        Call DoCheckControl("InputCheckBoxSpace", False)
        Exit Sub
    End If
End Sub

'******************************************************
' 入力ファイルのチェックボックス(パイプ)選択時のイベント
'******************************************************
Private Sub InputCheckBoxPipe_Click()
    If IsControlChecked("InputCheckBoxPipe") = True Then
        Call DoCheckControl("InputCheckBoxComma", False)
        Call DoCheckControl("InputCheckBoxSpace", False)
        Exit Sub
    End If
End Sub

'******************************************************
' 入力ファイルのチェックボックス(スペース)選択時のイベント
'******************************************************
Private Sub InputCheckBoxSpace_Click()
    If IsControlChecked("InputCheckBoxSpace") = True Then
        Call DoCheckControl("InputCheckBoxComma", False)
        Call DoCheckControl("InputCheckBoxPipe", False)
        Exit Sub
    End If
End Sub

'******************************************************
' 出力ファイルのチェックボックス(カンマ)選択時のイベント
'******************************************************
Private Sub OutputCheckBoxComma_Click()
    If IsControlChecked("OutputCheckBoxComma") = True Then
        Call DoCheckControl("OutputCheckBoxSpace", False)
        Call DoCheckControl("OutputCheckBoxPipe", False)
        Exit Sub
    End If
End Sub

'******************************************************
' 出力ファイルのチェックボックス(パイプ)選択時のイベント
'******************************************************
Private Sub OutputCheckBoxPipe_Click()
    If IsControlChecked("OutputCheckBoxPipe") = True Then
        Call DoCheckControl("OutputCheckBoxComma", False)
        Call DoCheckControl("OutputCheckBoxSpace", False)
        Exit Sub
    End If
End Sub

'******************************************************
' 出力ファイルのチェックボックス(スペース)選択時のイベント
'******************************************************
Private Sub OutputCheckBoxSpace_Click()
    If IsControlChecked("OutputCheckBoxSpace") = True Then
        Call DoCheckControl("OutputCheckBoxComma", False)
        Call DoCheckControl("OutputCheckBoxPipe", False)
        Exit Sub
    End If
End Sub

'******************************************************
' 参照ボタン押下時のイベント操作
'******************************************************
Private Sub RefferenceBtn1_Click()
    ' カレントフォルダ設定
    CreateObject("WScript.Shell").CurrentDirectory = ReferenceFolderPath
    OpenFileName = Application.GetOpenFilename()
    If OpenFileName <> "False" Then
        TargetFilePath.Value = Dir(OpenFileName)
        ReadFilePath = OpenFileName
        ActivateExecuteButton (True)
    End If
End Sub
