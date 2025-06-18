Attribute VB_Name = "E_CreateTestScriptModule"
' 手順E　テスト仕様書作成の挙動
Sub TestScript_Create_Click()
    Call CreatePhaseDefineData
    Call PopulateFormatDefs
    Call LoadPRS
    Call CreateTestScriptSheet
End Sub

' テスト仕様書作成
Sub CreateTestScriptSheet()
    Dim wsTestScript As Worksheet
    
    ' データチェック
    If Not BaseDataCheck(g_prsHeader, g_phaseDefs, g_phaseDefs, g_OPInformationList) Then
        Exit Sub
    End If
    
    ' シートコピー
    Set wsTestScript = CopySheet("Template", "テストスクリプト")

    ' 情報書き込み
    WritePRSInformation wsTestScript
    
    ' STEP番号設定
    SetStepNumber wsTestScript
    
    MsgBox "作成完了しました", vbInformation
End Sub

' PRS情報書き込み
Sub WritePRSInformation(ByRef wsTestScript As Worksheet)
    Dim opInfo As OPInformation
    Dim rowNumber As Long

    ' 最初は2行目から書き込み開始
    rowNumber = 2

    ' PRSファイルの情報（g_OPInformationList）を書込む
    For Each opInfo In g_OPInformationList
        WriteOpInformation wsTestScript, opInfo, rowNumber
    Next opInfo
End Sub

' OP情報書き込み
Sub WriteOpInformation(ByRef wsTestScript As Worksheet, ByVal opInfo As OPInformation, ByRef rowNumber As Long)
    Dim phaseInfo As PhaseInformation
    
    ' OPInformationの情報を書き込み (1行)
    With wsTestScript
        .Cells(rowNumber, 2).value = opInfo.GetID() & vbCrLf & opInfo.GetOPName() & vbCrLf & opInfo.GetCBBName()
        ' A列からJ列までの背景色を水色に設定
        .Range(.Cells(rowNumber, 1), .Cells(rowNumber, 11)).Interior.Color = RGB(173, 216, 230)
    End With
    rowNumber = rowNumber + 1
    
    ' PhaseInformationの情報を書き込み（複数行）
    For Each phaseInfo In opInfo.GetPhaseInformationList()
        WritePhaseInformation wsTestScript, phaseInfo, rowNumber
    Next phaseInfo
End Sub

' Phase情報書き込み
Sub WritePhaseInformation(ByRef wsTestScript As Worksheet, ByVal phaseInfo As PhaseInformation, ByRef rowNumber As Long)
    Dim phaseDef As PhaseDefine
    Dim riskIDs As Variant
    Dim riskIDEnum As riskIDEnum
    Dim phaseDefFound As Boolean
    Dim formatDefFound As Boolean
    
    ' Phase名に該当するリスクID群を取得
    riskIDs = GetRiskIDsByMatrix(phaseInfo.phaseName)
    
    ' リスクID数分定義情報を書き込む
    Dim i As Long
    For i = LBound(riskIDs) To UBound(riskIDs)
        riskIDEnum = riskIDs(i)
        
        ' データ（リスクIDごとの）書き込み
        WriteData wsTestScript, phaseInfo, riskIDEnum, rowNumber
    Next i

    ' PhaseDefineが見つからなかった場合のエラーハンドリングはここで
    If Not phaseDefFound Then
        ' 必要に応じて処理を追加
    End If
End Sub

' データ（リスクIDごと）書き込み
Sub WriteData(ByRef wsTestScript As Worksheet, ByVal phaseInfo As PhaseInformation, ByVal riskIDEnum As riskIDEnum, ByRef rowNumber As Long)
    Dim formatDef As FormatDefine
    
    For Each formatDef In g_formatDefs
        If formatDef.riskID = riskIDEnum Then
            If IsValidRiskIDs(phaseInfo, formatDef) Then
                ' フォーマット書き込み
                WriteFormat wsTestScript, formatDef, rowNumber
                
                ' フォーマットの置換文字列をPhase情報へ置き換え
                ReplaceData wsTestScript, phaseInfo, formatDef, rowNumber
    
                rowNumber = rowNumber + 1
            Else
                ' 書き込まない
            End If
        End If
    Next formatDef
End Sub

'
Function IsValidRiskIDs(ByVal phaseInfo As PhaseInformation, ByVal formatDef As FormatDefine) As Boolean
    Select Case formatDef.riskID
        Case SOPLINK_TYPE
            IsValidRiskIDs = (InStr(phaseInfo.RecipeParameter, "リンク：") > 0)
        Case Else
            IsValidRiskIDs = True
    End Select
End Function

' フォーマット（リスクIDごと）書き込み
Sub WriteFormat(ByRef wsTestScript As Worksheet, ByVal formatDef As FormatDefine, ByRef rowNumber As Long)
    ' 一行分の情報を書き込み
    With wsTestScript
        ' PRS参照
        '.Cells(rowNumber, 2).value = formatDef.PRSReference.baseString
        ' データ/前提条件
        '.Cells(rowNumber, 3).value = formatDef.Data_Prerequisites.baseString
        ' 試験指図書
        '.Cells(rowNumber, 4).value = formatDef.TestInstruction.baseString
        ' 期待される結果
        '.Cells(rowNumber, 5).value = formatDef.ExpectedResult.baseString
        ' リスクID
        .Cells(rowNumber, 6).value = GetRiskIDString(formatDef.riskID)
        ' 検査結果
        .Cells(rowNumber, 7).value = formatDef.TestResult
        ' エビデンス
        .Cells(rowNumber, 8).value = formatDef.Evidence
    End With
End Sub

Sub ReplaceData(ByRef wsTestScript As Worksheet, ByVal phaseInfo As PhaseInformation, ByVal formatDef As FormatDefine, ByRef rowNumber As Long)
    With wsTestScript
        ' PRS参照
        .Cells(rowNumber, 2).value = ReplaceFormatToPhaseData(wsTestScript, phaseInfo, formatDef.PRSReference)
        ' データ/前提条件
        .Cells(rowNumber, 3).value = ReplaceFormatToPhaseData(wsTestScript, phaseInfo, formatDef.Data_Prerequisites)
        ' 試験指図書
        .Cells(rowNumber, 4).value = ReplaceFormatToPhaseData(wsTestScript, phaseInfo, formatDef.TestInstruction)
        ' 期待される結果
        .Cells(rowNumber, 5).value = ReplaceFormatToPhaseData(wsTestScript, phaseInfo, formatDef.ExpectedResult)
    End With
End Sub

' フォーマットの置換文字列をPhase情報へ置き換え
Function ReplaceFormatToPhaseData(ByRef wsTestScript As Worksheet, ByVal phaseInfo As PhaseInformation, ByVal formatValue As FormatSettingValue) As String
    Dim replaceString As String
    Dim StringList As Collection
    Dim i As Long
    
    Set StringList = New Collection
    
    For Each header In formatValue.ReplaceTargetList
        StringList.Add (phaseInfo.GetMemberValueByHeader(CStr(header)))
    Next header
    
    ReplaceFormatToPhaseData = formatValue.ReplaceStrings(StringList)
End Function

' フェーズ名に該当するリスクID群を取得
Function GetRiskIDsByMatrix(ByVal phaseName As String) As Variant
    phaseDefFound = False
    
    For Each phaseDef In g_phaseDefs
        If phaseName = phaseDef.phaseName Then
            GetRiskIDsByMatrix = phaseDef.riskIDs
            Exit Function
        End If
    Next phaseDef
    
    ' エラー時または該当するphaseNameが見つからない場合には空の配列を返す
    GetRiskIDsByMatrix = Array()
End Function


