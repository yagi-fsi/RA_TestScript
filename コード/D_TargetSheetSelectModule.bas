Attribute VB_Name = "D_TargetSheetSelectModule"
' �菇D�@�쐬�ΏۃV�[�g�I�����̋���

' �f�[�^�J�n�s
Private Const ROW_NUMBER_DATA_START As Long = 4

Sub TargetSheet_Select_Click()
    ' �O������N���A
    ThisWorkbook.Sheets(SHEET_CREATE_TEST).Range(CELL_SOURCE_SHEET).value = ""
    
    ' �V�[�g�I��
    Call SelectPRSSheet
    
    ' �e�X�g�d�l���쐬�J�n
    If ThisWorkbook.Sheets(SHEET_CREATE_TEST).Range(CELL_SOURCE_SHEET).value <> "" Then
        Call LoadPRS
        DebugOPInformationList g_OPInformationList
    End If
End Sub

Sub SelectPRSSheet()
    Dim form As New SelectSheetDialog
    Dim SourceFilePath As String
    
    ' �t�@�C���p�X���擾
    SourceFilePath = ThisWorkbook.Sheets(SHEET_CREATE_TEST).Range(CELL_SOURCE_FILE).value
    
    ' �t�@�C���p�X�����[�U�[�t�H�[���ɐݒ�
    form.SourceFilePath = SourceFilePath
    form.Show

    ' �t�H�[����������ɑI�����ꂽ�V�[�g�����Z���֏�������
    Dim sheetName As String
    sheetName = form.SelectedSheet
    
    If sheetName <> "" Then
        ThisWorkbook.Sheets(SHEET_CREATE_TEST).Range(CELL_SOURCE_SHEET).value = sheetName
    End If
    
    ' �I������疾���I�ɉ��
    Set form = Nothing
End Sub


Sub LoadPRS()
    Dim wb As Workbook
    Dim ws As Worksheet
    
    ' �V�[�g�̏�����
    Application.ScreenUpdating = False
    Set wb = Workbooks.Open(filename:=ThisWorkbook.Sheets(SHEET_CREATE_TEST).Range(CELL_SOURCE_FILE).value, ReadOnly:=True, Notify:=False)
    Application.ScreenUpdating = True
    Set ws = wb.Sheets(ThisWorkbook.Sheets(SHEET_CREATE_TEST).Range(CELL_SOURCE_SHEET).value)
    
    ' �w�b�_�[���ۑ�
    Set g_prsHeader = New PRSHeaderInformation
    g_prsHeader.Initialize ws
    If Not g_prsHeader.IsValid Then
        MsgBox "PRS�t�@�C���̗�w�b�_������������܂���"
        Exit Sub
    End If
    
    ' ���ݒ�
    SetOPInformations ws
    
    wb.Close
End Sub

' OP���ݒ�
Sub SetOPInformations(ByRef ws As Worksheet)
    Dim opInfo As OPInformation
    Dim phaseInfo As PhaseInformation
    Dim row As Long
    Dim lastrow As Long
    Dim phaseList As Collection
    
    Set g_OPInformationList = New Collection
    

    ' �ŏI�s���v�Z
    lastrow = ws.Cells(ws.Rows.Count, g_prsHeader.GetColumnNumberID).End(xlUp).row
    
    ' �f�[�^�̓ǂݍ���
    row = ROW_NUMBER_DATA_START
    Do While row <= lastrow
        If Trim(ws.Cells(row, g_prsHeader.GetColumnNumberOP).value) <> "" Then
            ' OP���i�P���j�̍쐬
            Set opInfo = New OPInformation
            opInfo.SetOPName ws.Cells(row, g_prsHeader.GetColumnNumberOP).value
            opInfo.SetCBBName ws.Cells(row, g_prsHeader.GetColumnNumberOP).Offset(1, 0).value
            opInfo.SetID ws.Cells(row, g_prsHeader.GetColumnNumberID).value
            opInfo.SetOPName ws.Cells(row, g_prsHeader.GetColumnNumberOP).value
            opInfo.SetCapability ws.Cells(row, g_prsHeader.GetColumnNumberComment).value    ' TBD
            
            row = row + 1
            
            Set phaseList = New Collection
            
            ' Phase���ݒ�
            SetPhaseInformations ws, phaseList, row, lastrow
            For Each phaseInfo In phaseList
                opInfo.AddPhaseInformation phaseInfo
            Next phaseInfo
            
            ' OP��񃊃X�g�֒ǉ�
            g_OPInformationList.Add opInfo
        Else
            row = row + 1
        End If
    Loop
End Sub

' Phase���ݒ�
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

