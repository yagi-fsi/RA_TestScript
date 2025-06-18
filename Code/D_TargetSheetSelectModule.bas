Attribute VB_Name = "D_TargetSheetSelectModule"
' 手順D　作成対象シート選択時の挙動

' データ開始行
Private Const ROW_NUMBER_DATA_START As Long = 4

Sub TargetSheet_Select_Click()
    ' 前回情報をクリア
    ThisWorkbook.Sheets(SHEET_CREATE_TEST).Range(CELL_SOURCE_SHEET).value = ""
    
    ' シート選択
    Call SelectPRSSheet
    
    ' テスト仕様書作成開始
    If ThisWorkbook.Sheets(SHEET_CREATE_TEST).Range(CELL_SOURCE_SHEET).value <> "" Then
        Call LoadPRS
        DebugOPInformationList g_OPInformationList
    End If
End Sub

Sub SelectPRSSheet()
    Dim form As New SelectSheetDialog
    Dim SourceFilePath As String
    
    ' ファイルパスを取得
    SourceFilePath = ThisWorkbook.Sheets(SHEET_CREATE_TEST).Range(CELL_SOURCE_FILE).value
    
    ' ファイルパスをユーザーフォームに設定
    form.SourceFilePath = SourceFilePath
    form.Show

    ' フォームが閉じた後に選択されたシート名をセルへ書き込む
    Dim sheetName As String
    sheetName = form.SelectedSheet
    
    If sheetName <> "" Then
        ThisWorkbook.Sheets(SHEET_CREATE_TEST).Range(CELL_SOURCE_SHEET).value = sheetName
    End If
    
    ' 終わったら明示的に解放
    Set form = Nothing
End Sub


Sub LoadPRS()
    Dim wb As Workbook
    Dim ws As Worksheet
    
    ' シートの初期化
    Application.ScreenUpdating = False
    Set wb = Workbooks.Open(filename:=ThisWorkbook.Sheets(SHEET_CREATE_TEST).Range(CELL_SOURCE_FILE).value, ReadOnly:=True, Notify:=False)
    Application.ScreenUpdating = True
    Set ws = wb.Sheets(ThisWorkbook.Sheets(SHEET_CREATE_TEST).Range(CELL_SOURCE_SHEET).value)
    
    ' ヘッダー情報保存
    Set g_prsHeader = New PRSHeaderInformation
    g_prsHeader.Initialize ws
    If Not g_prsHeader.IsValid Then
        MsgBox "PRSファイルの列ヘッダが正しくありません"
        Exit Sub
    End If
    
    ' 情報設定
    SetOPInformations ws
    
    wb.Close
End Sub

' OP情報設定
Sub SetOPInformations(ByRef ws As Worksheet)
    Dim opInfo As OPInformation
    Dim phaseInfo As PhaseInformation
    Dim row As Long
    Dim lastrow As Long
    Dim phaseList As Collection
    
    Set g_OPInformationList = New Collection
    

    ' 最終行を計算
    lastrow = ws.Cells(ws.Rows.Count, g_prsHeader.GetColumnNumberID).End(xlUp).row
    
    ' データの読み込み
    row = ROW_NUMBER_DATA_START
    Do While row <= lastrow
        If Trim(ws.Cells(row, g_prsHeader.GetColumnNumberOP).value) <> "" Then
            ' OP情報（１つ分）の作成
            Set opInfo = New OPInformation
            opInfo.SetOPName ws.Cells(row, g_prsHeader.GetColumnNumberOP).value
            opInfo.SetCBBName ws.Cells(row, g_prsHeader.GetColumnNumberOP).Offset(1, 0).value
            opInfo.SetID ws.Cells(row, g_prsHeader.GetColumnNumberID).value
            opInfo.SetOPName ws.Cells(row, g_prsHeader.GetColumnNumberOP).value
            opInfo.SetCapability ws.Cells(row, g_prsHeader.GetColumnNumberComment).value    ' TBD
            
            row = row + 1
            
            Set phaseList = New Collection
            
            ' Phase情報設定
            SetPhaseInformations ws, phaseList, row, lastrow
            For Each phaseInfo In phaseList
                opInfo.AddPhaseInformation phaseInfo
            Next phaseInfo
            
            ' OP情報リストへ追加
            g_OPInformationList.Add opInfo
        Else
            row = row + 1
        End If
    Loop
End Sub

' Phase情報設定
Sub SetPhaseInformations(ByRef ws As Worksheet, ByRef phaseList As Collection, ByRef row As Long, ByVal lastrow As Long)

    Do While row <= lastrow And Trim(ws.Cells(row, g_prsHeader.GetColumnNumberOP).value) = ""
        Set phaseInfo = New PhaseInformation
        phaseInfo.Initialize ws.Cells(row, g_prsHeader.GetColumnNumberID).value, _
                             ws.Cells(row, g_prsHeader.GetColumnNumberPhaseIntroduction).value, _
                            "", _
                            ws.Cells(row, g_prsHeader.GetColumnNumberComment).value, _
                            ws.Cells(row, g_prsHeader.GetColumnNumberRecipeParameter).value, _
                            "", _
                            ws.Cells(row, g_prsHeader.GetColumnNumberMaterial).value, _
                            ws.Cells(row, g_prsHeader.GetColumnNumberEquipment).value, _
                            ws.Cells(row, g_prsHeader.GetColumnNumberPlace).value, _
                            ws.Cells(row, g_prsHeader.GetColumnNumberGMP).value
        phaseList.Add phaseInfo
        row = row + 1
    Loop
End Sub

'--------------------------------------'--------------------------------------'--------------------------------------

