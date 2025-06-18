Attribute VB_Name = "A_MatrixLoadModule"
' �菇A �}�g���N�X�Ǎ����̋���

' ��ԍ���`
Private Const COLUMN_NUMBER_PHASE_NAME As Long = 1
Private Const COLUMN_NUMBER_BUNDLE_TYPE As Long = 2
Private Const COLUMN_NUMBER_VERSION As Long = 3

' �f�[�^�J�n�s
Private Const ROW_NUMBER_DATA_START As Long = 4


Sub Matrix_Load_Click()
    Call CreatePhaseDefineData
End Sub

' ���X�N��ƃ��X�NID�̃}�b�s���O������
Sub InitializeRiskMappings(ByRef riskColumns As Variant, ByRef riskIDMapping As Variant)
    ' ���X�NID�ɑΉ������ԍ����w��
    ' ���̏��Ԃ����̂܂܃��X�NID�̃t�H�[�}�b�g�o�͏��i�uFormat�v�V�[�g�̓��e�j�ƂȂ�
    ' 1�Ԗڂ�A2, 2�Ԗڂ�A1, 3�Ԗڂ�B1, �E�E�E
    riskColumns = Array(5, 4, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36)

    ' �e���X�NID�ɑΉ����郊�X�NIDEnum�l���w��
    riskIDMapping = Array(RiskID_A2, RiskID_A1, RiskID_B1, RiskID_B2, RiskID_B3, RiskID_B4, RiskID_B5, RiskID_B6, _
                          RiskID_C1, RiskID_C2, RiskID_C3, RiskID_C4, RiskID_C5, RiskID_D1, RiskID_D2, RiskID_D3, RiskID_D4, RiskID_D5, _
                          RiskID_E1, RiskID_E2, RiskID_E3, RiskID_F1, RiskID_F2, RiskID_F3, RiskID_F4, _
                          RiskID_G1, RiskID_H1, RiskID_I1, RiskID_I2, RiskID_I3, RiskID_J1, RiskID_K1)
End Sub

' �}�g���N�X�V�[�g����t�F�[�Y��`�f�[�^���쐬����
Sub CreatePhaseDefineData()
    Dim wsMatrix As Worksheet
    Dim lastrow As Long
    Dim i As Long
    Dim pd As PhaseDefine
    Dim riskIDArray() As riskIDEnum

    ' �}�g���N�X�V�[�g��ݒ肵�A�ŏI�s���擾
    Set wsMatrix = ThisWorkbook.Sheets(SHEET_MATRIX)
    lastrow = wsMatrix.Cells(wsMatrix.Rows.Count, COLUMN_NUMBER_PHASE_NAME).End(xlUp).row
    Set g_phaseDefs = New Collection

    ' ���X�N��ƃ��X�NID�̃}�b�s���O��������
    Dim riskColumns As Variant
    Dim riskIDMapping As Variant
    Call InitializeRiskMappings(riskColumns, riskIDMapping)

    Dim beforePhaseName As String ' ���O�̃t�F�[�Y�����L�����Ă���

    ' �e�s�̃f�[�^���t�F�[�Y��`�Ƃ��ď���
    For rowIndex = ROW_NUMBER_DATA_START To lastrow
        Set pd = New PhaseDefine

        ' �t�F�[�Y�����擾
        Dim phaseName As String
        phaseName = GetPhaseName(wsMatrix.Cells(rowIndex, COLUMN_NUMBER_PHASE_NAME).value, beforePhaseName)
        beforePhaseName = phaseName

        ' �o���h���^�C�v���擾
        Dim bundleTypeStr As String
        bundleTypeStr = wsMatrix.Cells(rowIndex, COLUMN_NUMBER_BUNDLE_TYPE).value
        Dim bandleType As BandleTypeEnum
        bandleType = GetBandleType(bundleTypeStr)

        ' �o�[�W�������擾
        Dim version As Double
        version = GetVersion(wsMatrix.Cells(rowIndex, COLUMN_NUMBER_VERSION).value)

        ' ���X�NID���擾
        riskIDArray = GetRiskIDs(wsMatrix, rowIndex, riskColumns, riskIDMapping)

        ' �t�F�[�Y��`�f�[�^��������
        pd.Initialize phaseName, version, "", riskIDArray, bandleType

        ' �R���N�V�����ɒǉ�
        g_phaseDefs.Add pd
    Next rowIndex
End Sub

' �Z���l����t�F�[�Y�����擾���A�K�v�ł���ΑO�̃t�F�[�Y�����g�p
Function GetPhaseName(fullCellValue As String, ByVal beforePhaseName As String) As String
    Dim phaseName As String

    ' �t���Z���l�ɃX�y�[�X���܂܂��ꍇ�́A�t�F�[�Y���Ƃ��ăX�y�[�X�O�̕�������擾
    If InStr(fullCellValue, " ") > 0 Then
        phaseName = Split(fullCellValue, " ")(0)
    Else
        phaseName = fullCellValue
    End If

    ' �t�F�[�Y������̏ꍇ�͑O�̃t�F�[�Y�����g�p
    If IsEmpty(phaseName) Then
        phaseName = beforePhaseName
    End If
    
    GetPhaseName = phaseName
End Function

' �����񂩂�o���h���^�C�v�𔻕�
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

' �Z������擾�����l�����l�ł���ꍇ�A���̒l���o�[�W�����Ƃ��Ď擾�B�����łȂ��ꍇ��0��Ԃ��B
Function GetVersion(versionInput As Variant) As Double
    If IsNumeric(versionInput) And Not IsEmpty(versionInput) Then
        GetVersion = CDbl(versionInput)
    Else
        GetVersion = 0
    End If
End Function

' �w�肳�ꂽ�s�̃��X�NID���擾
Function GetRiskIDs(wsMatrix As Worksheet, ByVal rowIndex As Long, _
                    ByVal riskColumns As Variant, ByVal riskIDMapping As Variant) As riskIDEnum()
    Dim riskIDs As Collection
    Set riskIDs = New Collection

    Dim riskIDIndex As Long
    ' �e���X�N��ɂ���"X"�܂���"��"������ΑΉ����郊�X�NID�����W
    For riskIDIndex = LBound(riskColumns) To UBound(riskColumns)
        If wsMatrix.Cells(rowIndex, riskColumns(riskIDIndex)).value = "X" Or wsMatrix.Cells(rowIndex, riskColumns(riskIDIndex)).value = "��" Then
            riskIDs.Add riskIDMapping(riskIDIndex)
        End If
    Next riskIDIndex
    
    Dim riskIDArray() As riskIDEnum
    ' ���W�������X�NID��z��ɕϊ�
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

