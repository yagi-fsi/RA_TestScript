Attribute VB_Name = "A_MatrixLoadModule"
' 手順A マトリクス読込時の挙動

' 列番号定義
Private Const COLUMN_NUMBER_PHASE_NAME As Long = 1
Private Const COLUMN_NUMBER_BUNDLE_TYPE As Long = 2
Private Const COLUMN_NUMBER_VERSION As Long = 3

' データ開始行
Private Const ROW_NUMBER_DATA_START As Long = 4


Sub Matrix_Load_Click()
    Call CreatePhaseDefineData
End Sub

' リスク列とリスクIDのマッピング初期化
Sub InitializeRiskMappings(ByRef riskColumns As Variant, ByRef riskIDMapping As Variant)
    ' リスクIDに対応する列番号を指定
    ' この順番がそのままリスクIDのフォーマット出力順（「Format」シートの内容）となる
    ' 1番目はA2, 2番目はA1, 3番目はB1, ・・・
    riskColumns = Array(5, 4, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36)

    ' 各リスクIDに対応するリスクIDEnum値を指定
    riskIDMapping = Array(RiskID_A2, RiskID_A1, RiskID_B1, RiskID_B2, RiskID_B3, RiskID_B4, RiskID_B5, RiskID_B6, _
                          RiskID_C1, RiskID_C2, RiskID_C3, RiskID_C4, RiskID_C5, RiskID_D1, RiskID_D2, RiskID_D3, RiskID_D4, RiskID_D5, _
                          RiskID_E1, RiskID_E2, RiskID_E3, RiskID_F1, RiskID_F2, RiskID_F3, RiskID_F4, _
                          RiskID_G1, RiskID_H1, RiskID_I1, RiskID_I2, RiskID_I3, RiskID_J1, RiskID_K1)
End Sub

' マトリクスシートからフェーズ定義データを作成する
Sub CreatePhaseDefineData()
    Dim wsMatrix As Worksheet
    Dim lastrow As Long
    Dim i As Long
    Dim pd As PhaseDefine
    Dim riskIDArray() As riskIDEnum

    ' マトリクスシートを設定し、最終行を取得
    Set wsMatrix = ThisWorkbook.Sheets(SHEET_MATRIX)
    lastrow = wsMatrix.Cells(wsMatrix.Rows.Count, COLUMN_NUMBER_PHASE_NAME).End(xlUp).row
    Set g_phaseDefs = New Collection

    ' リスク列とリスクIDのマッピングを初期化
    Dim riskColumns As Variant
    Dim riskIDMapping As Variant
    Call InitializeRiskMappings(riskColumns, riskIDMapping)

    Dim beforePhaseName As String ' 直前のフェーズ名を記憶しておく

    ' 各行のデータをフェーズ定義として処理
    For rowIndex = ROW_NUMBER_DATA_START To lastrow
        Set pd = New PhaseDefine

        ' フェーズ名を取得
        Dim phaseName As String
        phaseName = GetPhaseName(wsMatrix.Cells(rowIndex, COLUMN_NUMBER_PHASE_NAME).value, beforePhaseName)
        beforePhaseName = phaseName

        ' バンドルタイプを取得
        Dim bundleTypeStr As String
        bundleTypeStr = wsMatrix.Cells(rowIndex, COLUMN_NUMBER_BUNDLE_TYPE).value
        Dim bandleType As BandleTypeEnum
        bandleType = GetBandleType(bundleTypeStr)

        ' バージョンを取得
        Dim version As Double
        version = GetVersion(wsMatrix.Cells(rowIndex, COLUMN_NUMBER_VERSION).value)

        ' リスクIDを取得
        riskIDArray = GetRiskIDs(wsMatrix, rowIndex, riskColumns, riskIDMapping)

        ' フェーズ定義データを初期化
        pd.Initialize phaseName, version, "", riskIDArray, bandleType

        ' コレクションに追加
        g_phaseDefs.Add pd
    Next rowIndex
End Sub

' セル値からフェーズ名を取得し、必要であれば前のフェーズ名を使用
Function GetPhaseName(fullCellValue As String, ByVal beforePhaseName As String) As String
    Dim phaseName As String

    ' フルセル値にスペースが含まれる場合は、フェーズ名としてスペース前の文字列を取得
    If InStr(fullCellValue, " ") > 0 Then
        phaseName = Split(fullCellValue, " ")(0)
    Else
        phaseName = fullCellValue
    End If

    ' フェーズ名が空の場合は前のフェーズ名を使用
    If IsEmpty(phaseName) Then
        phaseName = beforePhaseName
    End If
    
    GetPhaseName = phaseName
End Function

' 文字列からバンドルタイプを判別
Function GetBandleType(bundleTypeStr As String) As BandleTypeEnum
    Select Case bundleTypeStr
        Case "Boolean": GetBandleType = BandleType_Boolean
        Case "DateTime": GetBandleType = BandleType_DateTime
        Case "DateTimeWithCulculate": GetBandleType = BandleType_DateTimeWithCulculate
        Case "Duration": GetBandleType = BandleType_Duration
        Case "DurationWithCulculate": GetBandleType = BandleType_DurationWithCulculate
        Case "MeasuredValue": GetBandleType = BandleType_MeasuredValue
        Case "MeasuredValueWithCulculate": GetBandleType = BandleType_MeasuredValueWithCulculate
        Case "ProcessValue": GetBandleType = BandleType_ProcessValue
        Case "ProcessValueWithCulculate": GetBandleType = BandleType_ProcessValueWithCulculate
        Case "String": GetBandleType = BandleType_String
        Case "Timestamp": GetBandleType = BandleType_Timestamp
        Case "TimestampWithCulculate": GetBandleType = BandleType_TimestampWithCulculate
        Case Else: GetBandleType = BandleType_String
    End Select
End Function

' セルから取得した値が数値である場合、その値をバージョンとして取得。そうでない場合は0を返す。
Function GetVersion(versionInput As Variant) As Double
    If IsNumeric(versionInput) And Not IsEmpty(versionInput) Then
        GetVersion = CDbl(versionInput)
    Else
        GetVersion = 0
    End If
End Function

' 指定された行のリスクIDを取得
Function GetRiskIDs(wsMatrix As Worksheet, ByVal rowIndex As Long, _
                    ByVal riskColumns As Variant, ByVal riskIDMapping As Variant) As riskIDEnum()
    Dim riskIDs As Collection
    Set riskIDs = New Collection

    Dim riskIDIndex As Long
    ' 各リスク列について"X"または"△"があれば対応するリスクIDを収集
    For riskIDIndex = LBound(riskColumns) To UBound(riskColumns)
        If wsMatrix.Cells(rowIndex, riskColumns(riskIDIndex)).value = "X" Or wsMatrix.Cells(rowIndex, riskColumns(riskIDIndex)).value = "△" Then
            riskIDs.Add riskIDMapping(riskIDIndex)
        End If
    Next riskIDIndex
    
    Dim riskIDArray() As riskIDEnum
    ' 収集したリスクIDを配列に変換
    If riskIDs.Count > 0 Then
        ReDim riskIDArray(riskIDs.Count - 1)
        Dim j As Long
        For j = 1 To riskIDs.Count
            riskIDArray(j - 1) = riskIDs.Item(j)
        Next j
    Else
        ReDim riskIDArray(0)
    End If
    
    GetRiskIDs = riskIDArray
End Function

