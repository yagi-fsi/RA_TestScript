Attribute VB_Name = "C_PRSSelectModule"
' 手順C　PRSファイル選択時の挙動
Sub PRS_Select()
    Call ShowFileDialog
End Sub

Sub ShowFileDialog()
    Dim fd As FileDialog
    Dim filePath As String
    Dim initialPath As String
    
    ' マクロファイルのフォルダを初期パスとして設定
    initialPath = ThisWorkbook.path
    
    ' ファイルダイアログを作成
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    ' ダイアログのタイトルを設定
    fd.Title = "ファイルを選択してください"
    
    ' 初期パスを設定
    fd.InitialFileName = initialPath & "\"

    ' ファイルフィルターを設定
    fd.Filters.Clear
    fd.Filters.Add "Excelファイル", "*.xls; *.xlsx"
    
    ' ダイアログを表示して、ユーザーがファイルを選択する
    If fd.Show = -1 Then
        ' 選択したファイルのフルパスを取得
        filePath = fd.SelectedItems(1)
        
        ' 指定されたセルにフルパスを入力
        ThisWorkbook.Sheets(SHEET_CREATE_TEST).Range(CELL_SOURCE_FILE).value = filePath
    End If
End Sub
