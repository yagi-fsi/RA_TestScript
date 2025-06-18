Attribute VB_Name = "E_CreateTestScriptModule"
' �菇E�@�e�X�g�d�l���쐬�̋���
Sub TestScript_Create_Click()
    Call CreatePhaseDefineData
    Call PopulateFormatDefs
    Call LoadPRS
    Call CreateTestScriptSheet
End Sub

' �e�X�g�d�l���쐬
Sub CreateTestScriptSheet()
    Dim wsTestScript As Worksheet
    
    ' �f�[�^�`�F�b�N
    If Not BaseDataCheck(g_prsHeader, g_phaseDefs, g_phaseDefs, g_OPInformationList) Then
        Exit Sub
    End If
    
    ' �V�[�g�R�s�[
    Set wsTestScript = CopySheet("Template", "�e�X�g�X�N���v�g")

    ' ��񏑂�����
    WritePRSInformation wsTestScript
    
    ' STEP�ԍ��ݒ�
    SetStepNumber wsTestScript
    
    MsgBox "�쐬�������܂���", vbInformation
End Sub

' PRS��񏑂�����
Sub WritePRSInformation(ByRef wsTestScript As Worksheet)
    Dim opInfo As OPInformation
    Dim rowNumber As Long

    ' �ŏ���2�s�ڂ��珑�����݊J�n
    rowNumber = 2

    ' PRS�t�@�C���̏��ig_OPInformationList�j��������
    For Each opInfo In g_OPInformationList
        WriteOpInformation wsTestScript, opInfo, rowNumber
    Next opInfo
End Sub

' OP��񏑂�����
Sub WriteOpInformation(ByRef wsTestScript As Worksheet, ByVal opInfo As OPInformation, ByRef rowNumber As Long)
    Dim phaseInfo As PhaseInformation
    
    ' OPInformation�̏����������� (1�s)
    With wsTestScript
        .Cells(rowNumber, 2).value = opInfo.GetID() & vbCrLf & opInfo.GetOPName() & vbCrLf & opInfo.GetCBBName()
        ' A�񂩂�J��܂ł̔w�i�F�𐅐F�ɐݒ�
        .Range(.Cells(rowNumber, 1), .Cells(rowNumber, 11)).Interior.Color = RGB(173, 216, 230)
    End With
    rowNumber = rowNumber + 1
    
    ' PhaseInformation�̏����������݁i�����s�j
    For Each phaseInfo In opInfo.GetPhaseInformationList()
        WritePhaseInformation wsTestScript, phaseInfo, rowNumber
    Next phaseInfo
End Sub

' Phase��񏑂�����
Sub WritePhaseInformation(ByRef wsTestScript As Worksheet, ByVal phaseInfo As PhaseInformation, ByRef rowNumber As Long)
    Dim phaseDef As PhaseDefine
    Dim riskIDs As Variant
    Dim riskIDEnum As riskIDEnum
    Dim phaseDefFound As Boolean
    Dim formatDefFound As Boolean
    
    ' Phase���ɊY�����郊�X�NID�Q���擾
    riskIDs = GetRiskIDsByMatrix(phaseInfo.phaseName)
    
    ' ���X�NID������`������������
    Dim i As Long
    For i = LBound(riskIDs) To UBound(riskIDs)
        riskIDEnum = riskIDs(i)
        
        ' �f�[�^�i���X�NID���Ƃ́j��������
        WriteData wsTestScript, phaseInfo, riskIDEnum, rowNumber
    Next i

    ' PhaseDefine��������Ȃ������ꍇ�̃G���[�n���h�����O�͂�����
    If Not phaseDefFound Then
        ' �K�v�ɉ����ď�����ǉ�
    End If
End Sub

' �f�[�^�i���X�NID���Ɓj��������
Sub WriteData(ByRef wsTestScript As Worksheet, ByVal phaseInfo As PhaseInformation, ByVal riskIDEnum As riskIDEnum, ByRef rowNumber As Long)
    Dim formatDef As FormatDefine
    
    For Each formatDef In g_formatDefs
        If formatDef.riskID = riskIDEnum Then
            If IsValidRiskIDs(phaseInfo, formatDef) Then
                ' �t�H�[�}�b�g��������
                WriteFormat wsTestScript, formatDef, rowNumber
                
                ' �t�H�[�}�b�g�̒u���������Phase���֒u������
                ReplaceData wsTestScript, phaseInfo, formatDef, rowNumber
    
                rowNumber = rowNumber + 1
            Else
                ' �������܂Ȃ�
            End If
        End If
    Next formatDef
End Sub

'
Function IsValidRiskIDs(ByVal phaseInfo As PhaseInformation, ByVal formatDef As FormatDefine) As Boolean
    Select Case formatDef.riskID
        Case SOPLINK_TYPE
            IsValidRiskIDs = (InStr(phaseInfo.RecipeParameter, "�����N�F") > 0)
        Case Else
            IsValidRiskIDs = True
    End Select
End Function

' �t�H�[�}�b�g�i���X�NID���Ɓj��������
Sub WriteFormat(ByRef wsTestScript As Worksheet, ByVal formatDef As FormatDefine, ByRef rowNumber As Long)
    ' ��s���̏�����������
    With wsTestScript
        ' PRS�Q��
        '.Cells(rowNumber, 2).value = formatDef.PRSReference.baseString
        ' �f�[�^/�O�����
        '.Cells(rowNumber, 3).value = formatDef.Data_Prerequisites.baseString
        ' �����w�}��
        '.Cells(rowNumber, 4).value = formatDef.TestInstruction.baseString
        ' ���҂���錋��
        '.Cells(rowNumber, 5).value = formatDef.ExpectedResult.baseString
        ' ���X�NID
        .Cells(rowNumber, 6).value = GetRiskIDString(formatDef.riskID)
        ' ��������
        .Cells(rowNumber, 7).value = formatDef.TestResult
        ' �G�r�f���X
        .Cells(rowNumber, 8).value = formatDef.Evidence
    End With
End Sub

Sub ReplaceData(ByRef wsTestScript As Worksheet, ByVal phaseInfo As PhaseInformation, ByVal formatDef As FormatDefine, ByRef rowNumber As Long)
    With wsTestScript
        ' PRS�Q��
        .Cells(rowNumber, 2).value = ReplaceFormatToPhaseData(wsTestScript, phaseInfo, formatDef.PRSReference)
        ' �f�[�^/�O�����
        .Cells(rowNumber, 3).value = ReplaceFormatToPhaseData(wsTestScript, phaseInfo, formatDef.Data_Prerequisites)
        ' �����w�}��
        .Cells(rowNumber, 4).value = ReplaceFormatToPhaseData(wsTestScript, phaseInfo, formatDef.TestInstruction)
        ' ���҂���錋��
        .Cells(rowNumber, 5).value = ReplaceFormatToPhaseData(wsTestScript, phaseInfo, formatDef.ExpectedResult)
    End With
End Sub

' �t�H�[�}�b�g�̒u���������Phase���֒u������
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

' �t�F�[�Y���ɊY�����郊�X�NID�Q���擾
Function GetRiskIDsByMatrix(ByVal phaseName As String) As Variant
    phaseDefFound = False
    
    For Each phaseDef In g_phaseDefs
        If phaseName = phaseDef.phaseName Then
            GetRiskIDsByMatrix = phaseDef.riskIDs
            Exit Function
        End If
    Next phaseDef
    
    ' �G���[���܂��͊Y������phaseName��������Ȃ��ꍇ�ɂ͋�̔z���Ԃ�
    GetRiskIDsByMatrix = Array()
End Function


