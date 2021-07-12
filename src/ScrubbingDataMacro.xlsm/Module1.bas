Attribute VB_Name = "Module1"
'******************************************************
' 変数宣言部
'******************************************************
Public ReadFilePath As String
Public ReferenceFolderPath As String
Public Const SheetLabelOriginal = "_original"
Public Const SheetLabelScrubbed = "_scrubbed"
Public Const DebugFlag = "True"

'******************************************************
' 初期化
'******************************************************
Public Sub Initialize()
    ' デフォルトではフォームのみを表示
    ActivateApp (False)
    
    ' チェックボックスの初期値設定
    UserForm1.Controls("InputCheckBoxComma").Value = True
    UserForm1.Controls("InputCheckBoxSpace").Value = False
    UserForm1.Controls("InputCheckBoxPipe").Value = False
    UserForm1.Controls("OutputCheckBoxComma").Value = True
    UserForm1.Controls("OutputCheckBoxSpace").Value = False
    UserForm1.Controls("OutputCheckBoxPipe").Value = False
        
    ' 参照フォルダの設定
    ReferenceFolderPath = Application.ThisWorkbook.Path
    ' 実行ボタンはデフォルト非活性とする
    UserForm1.Controls("ExecBtn").Enabled = False
    ' モードレス表示
    UserForm1.Show vbModeless
End Sub

'******************************************************
' コントロールがイネーブルか否か
'******************************************************
Public Function IsControlEnabled(cbName As String)
    IsCheckBoxEnabled = UserForm1.Controls(cbName).Enabled
End Function

'******************************************************
' Controlがチェックされているか否か
'******************************************************
Public Function IsControlChecked(cbName As String)
    IsControlChecked = UserForm1.Controls(cbName).Value
End Function

'******************************************************
' Conrolをイネーブル化
'******************************************************
Public Sub EnableControl(ctlName As String, sts As Boolean)
    UserForm1.Controls(ctlName).Enabled = sts
End Sub

'******************************************************
' Conrolにチェックを入れる
'******************************************************
Public Sub DoCheckControl(ctlName As String, sts As Boolean)
    UserForm1.Controls(ctlName).Value = sts
End Sub

'******************************************************
' Excelアプリケーションのアクティベーション
'******************************************************
Public Sub ActivateApp(sts As Boolean)
    If IsDebug = True Then
        ' Debugモード時はExcelを表示する
        Application.Visible = True
        Exit Sub
    End If
    Application.Visible = sts
End Sub

'******************************************************
' デバッグモードか否か
'******************************************************
Public Function IsDebug()
    If DebugFlag = "True" Then
        IsDebug = True
        Exit Function
    End If
    IsDebug = False
End Function

'******************************************************
' シート名の生成
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
' デリミタの検索
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
' シートの追加
' [Parameters]
'     Ptn : True(Originalの場合), False(Scrubbing後の場合)
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
    ' 読み取り専用でファイルを開く
    Set csvFile = fso.OpenTextFile(ReadFilePath, 1)
    i = 1
    Do While csvFile.AtEndOfStream = False
        csvData = csvFile.ReadLine
        ' TODO CSVを前提とした変数名になっているので見直しする
        'splitcsvData = Split(csvData, ",")
        splitcsvData = Split(csvData, FindDelimeter(True))
        If ptn = False Then
            Call EditRecord(splitcsvData)
        End If
        
        ' 配列の要素数取得(0スタートのため1を足す)
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
' 一行レコードの編集
' [Parameters]
'    rec : デリミタにより分割された配列データ
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
    
    ' カラム0：IDとみなす
    rec(0) = 99999999
    ' カラム1：生年月日　和暦→西暦変換
    tempBirthDate = DateValue(rec(1))
    BirthDate = Format(tempBirthDate, "yyyy/mm/dd")
    rec(1) = BirthDate
    
ErrLabel:
    If tempBirthDate = Empty Then
        MsgBox "生年月日データに異常があります"
    End If
End Sub

'******************************************************
' ファイル書き込み
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

    ' 出力元シートの設定
    Set ws = ThisWorkbook.Worksheets(outputSheetName)
    ws.Activate
    
    ' 出力対象範囲
    startCol = 1
    startRow = 1
    endCol = Range("A1").End(xlToRight).Column
    endRow = Range("A3").End(xlDown).Row - 1
    
    Set outputArray = ActiveSheet.Range(Cells(startRow, startCol), Cells(endRow, endCol))
    
    '列×行数分ループ
    For i = 1 To endRow
        For j = 1 To endCol
        
            '最終列かつ最終行に来たら改行
            If i = endRow And j = endCol Then
                outputText = outputText & outputArray(i, j)
                GoTo NextLoop
            End If
            
            '最終列に来たら改行
            If j = endCol Then
                outputText = outputText & outputArray(i, j) & vbCrLf
                GoTo NextLoop
            End If
            
            'タブ区切りで1セルずつ変数へ格納
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
    
    MsgBox "出力完了"
End Sub
