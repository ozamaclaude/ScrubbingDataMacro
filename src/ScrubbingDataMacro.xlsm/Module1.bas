Attribute VB_Name = "Module1"
'******************************************************
' �ϐ��錾��
'******************************************************
Public ReadFilePath As String
Public ReferenceFolderPath As String
Public Const SheetLabelOriginal = "_original"
Public Const SheetLabelScrubbed = "_scrubbed"
Public Const DebugFlag = "True"

'******************************************************
' ������
'******************************************************
Public Sub Initialize()
    ' �f�t�H���g�ł̓t�H�[���݂̂�\��
    ActivateApp (False)
    
    ' �`�F�b�N�{�b�N�X�̏����l�ݒ�
    UserForm1.Controls("InputCheckBoxComma").Value = True
    UserForm1.Controls("InputCheckBoxSpace").Value = False
    UserForm1.Controls("InputCheckBoxPipe").Value = False
    UserForm1.Controls("OutputCheckBoxComma").Value = True
    UserForm1.Controls("OutputCheckBoxSpace").Value = False
    UserForm1.Controls("OutputCheckBoxPipe").Value = False
        
    ' �Q�ƃt�H���_�̐ݒ�
    ReferenceFolderPath = Application.ThisWorkbook.Path
    ' ���s�{�^���̓f�t�H���g�񊈐��Ƃ���
    UserForm1.Controls("ExecBtn").Enabled = False
    ' ���[�h���X�\��
    UserForm1.Show vbModeless
End Sub

'******************************************************
' �R���g���[�����C�l�[�u�����ۂ�
'******************************************************
Public Function IsControlEnabled(cbName As String)
    IsCheckBoxEnabled = UserForm1.Controls(cbName).Enabled
End Function

'******************************************************
' Control���`�F�b�N����Ă��邩�ۂ�
'******************************************************
Public Function IsControlChecked(cbName As String)
    IsControlChecked = UserForm1.Controls(cbName).Value
End Function

'******************************************************
' Conrol���C�l�[�u����
'******************************************************
Public Sub EnableControl(ctlName As String, sts As Boolean)
    UserForm1.Controls(ctlName).Enabled = sts
End Sub

'******************************************************
' Conrol�Ƀ`�F�b�N������
'******************************************************
Public Sub DoCheckControl(ctlName As String, sts As Boolean)
    UserForm1.Controls(ctlName).Value = sts
End Sub

'******************************************************
' Excel�A�v���P�[�V�����̃A�N�e�B�x�[�V����
'******************************************************
Public Sub ActivateApp(sts As Boolean)
    If IsDebug = True Then
        ' Debug���[�h����Excel��\������
        Application.Visible = True
        Exit Sub
    End If
    Application.Visible = sts
End Sub

'******************************************************
' �f�o�b�O���[�h���ۂ�
'******************************************************
Public Function IsDebug()
    If DebugFlag = "True" Then
        IsDebug = True
        Exit Function
    End If
    IsDebug = False
End Function

'******************************************************
' �V�[�g���̐���
'******************************************************
Public Function GenerateSheetName(ptn As Boolean)
    Dim tempSheetName As String
    tempSheetName = Format(Now, "yyyymmdd-hhmmss")
    If ptn = True Then
        GenerateSheetName = tempSheetName & SheetLabelOriginal
        Exit Function
    End If
    GenerateSheetName = tempSheetName & SheetLabelScrubbed
End Function

'******************************************************
' �f���~�^�̌���
'******************************************************
Public Function FindDelimeter(isInput As Boolean)
    If isInput = True Then
        If IsControlChecked("InputCheckBoxSpace") = True Then
            FindDelimeter = " "
            Exit Function
        End If
        If IsControlChecked("InputCheckBoxComma") = True Then
            FindDelimeter = ","
            Exit Function
        End If
        If IsControlChecked("InputCheckBoxPipe") = True Then
            FindDelimeter = "|"
            Exit Function
        End If
    End If
    If IsControlChecked("InputCheckBoxSpace") = True Then
        FindDelimeter = " "
        Exit Function
    End If
    If IsControlChecked("InputCheckBoxComma") = True Then
        FindDelimeter = ","
        Exit Function
    End If
    If IsControlChecked("InputCheckBoxPipe") = True Then
        FindDelimeter = "|"
        Exit Function
    End If
End Function

'******************************************************
' �V�[�g�̒ǉ�
' [Parameters]
'     Ptn : True(Original�̏ꍇ), False(Scrubbing��̏ꍇ)
'******************************************************
Public Function AddSheet(ptn As Boolean)
    Dim sheetName As String
    Dim fso As New Scripting.FileSystemObject
    Dim csvFile As Object
    Dim csvData As String
    Dim splitcsvData As Variant
    Dim i As Integer
    Dim j As Integer
    sheetName = GenerateSheetName(ptn)
    Worksheets.Add(After:=Worksheets(Worksheets.Count)).Name = sheetName
    ' �ǂݎ���p�Ńt�@�C�����J��
    Set csvFile = fso.OpenTextFile(ReadFilePath, 1)
    i = 1
    Do While csvFile.AtEndOfStream = False
        csvData = csvFile.ReadLine
        ' TODO CSV��O��Ƃ����ϐ����ɂȂ��Ă���̂Ō���������
        'splitcsvData = Split(csvData, ",")
        splitcsvData = Split(csvData, FindDelimeter(True))
        If ptn = False Then
            Call EditRecord(splitcsvData)
        End If
        
        ' �z��̗v�f���擾(0�X�^�[�g�̂���1�𑫂�)
        j = UBound(splitcsvData) + 1
        Worksheets(sheetName).Range(Worksheets(sheetName).Cells(i, 1), Worksheets(sheetName).Cells(i, j)).Value _
        = splitcsvData
        i = i + 1
    Loop
    csvFile.Close
    Set csvFile = Nothing
    Set fso = Nothing
End Function

'******************************************************
' ��s���R�[�h�̕ҏW
' [Parameters]
'    rec : �f���~�^�ɂ�蕪�����ꂽ�z��f�[�^
'******************************************************
Sub EditRecord(ByRef rec As Variant)
    Debug.Print (rec(0))
    Debug.Print (rec(1))
    Debug.Print (rec(2))
    Debug.Print (rec(3))
    Debug.Print (rec(4))
    Dim BirthDate
    Dim birthDay As String
    
    On Error GoTo ErrLabel
    
    ' �J����0�FID�Ƃ݂Ȃ�
    rec(0) = 99999999
    ' �J����1�F���N�����@�a�����ϊ�
    tempBirthDate = DateValue(rec(1))
    BirthDate = Format(tempBirthDate, "yyyy/mm/dd")
    rec(1) = BirthDate
    
ErrLabel:
    If tempBirthDate = Empty Then
        MsgBox "���N�����f�[�^�Ɉُ킪����܂�"
    End If
End Sub

'******************************************************
' �t�@�C����������
' memo : https://loosecarrot.com/2020/02/15/3951/
'******************************************************
Public Sub WriteFile(outputSheetName As String)
    Dim startRow As Long
    Dim endRow As Long
    Dim startCol As Long
    Dim endCol As Long
    Dim ws As Worksheet
    Dim outputArray As Variant
    Dim outputText As String
    Dim i As Integer
    Dim j As Long
    Dim filePath As String

    ' �o�͌��V�[�g�̐ݒ�
    Set ws = ThisWorkbook.Worksheets(outputSheetName)
    ws.Activate
    
    ' �o�͑Ώ۔͈�
    startCol = 1
    startRow = 1
    endCol = Range("A1").End(xlToRight).Column
    endRow = Range("A3").End(xlDown).Row - 1
    
    Set outputArray = ActiveSheet.Range(Cells(startRow, startCol), Cells(endRow, endCol))
    
    '��~�s�������[�v
    For i = 1 To endRow
        For j = 1 To endCol
        
            '�ŏI�񂩂ŏI�s�ɗ�������s
            If i = endRow And j = endCol Then
                outputText = outputText & outputArray(i, j)
                GoTo NextLoop
            End If
            
            '�ŏI��ɗ�������s
            If j = endCol Then
                outputText = outputText & outputArray(i, j) & vbCrLf
                GoTo NextLoop
            End If
            
            '�^�u��؂��1�Z�����ϐ��֊i�[
            outputText = outputText & outputArray(i, j) & vbTab
            
NextLoop:
            
        Next j
    Next i
    
    filePath = ReadFilePath & "aaaaaaaaa"
    Open filePath For Output As #1
    Print #1, outputText
    Close #1
    
    Set ws = ThisWorkbook.Worksheets(outputSheetName)
    ws.Activate
    
    MsgBox "�o�͊���"
End Sub
