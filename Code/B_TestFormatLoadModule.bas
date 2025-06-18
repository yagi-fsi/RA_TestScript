Attribute VB_Name = "B_TestFormatLoadModule"
' �菇B�@�e�X�g�t�H�[�}�b�g��ݒ莞�̋���

' ��ԍ���`
Private Const COLUMN_NUMBER_PRSREFERENCE As Long = 2
Private Const COLUMN_NUMBER_PREREQUISISITES As Long = 3
Private Const COLUMN_NUMBER_TESTINSTRUCTION As Long = 4
Private Const COLUMN_NUMBER_EXPECTEDRESULT As Long = 5
Private Const COLUMN_NUMBER_RISKID As Long = 6
Private Const COLUMN_NUMBER_TESTRESULT As Long = 7
Private Const COLUMN_NUMBER_EVIDENCE As Long = 8

' �f�[�^�J�n�s
Private Const ROW_NUMBER_DATA_START As Long = 2

Sub Format_Load_Click()
    Call PopulateFormatDefs
    'Call DebugFormatDefs
End Sub

Sub PopulateFormatDefs()
    Dim ws As Worksheet
    Dim lastrow As Long
    Dim currentRow As Long
    Dim formatDef As FormatDefine

    ' "Format"�V�[�g���J��
    Set ws = ThisWorkbook.Sheets(SHEET_FORMAT)

    ' FormatDefine�̃��X�g�쐬
    Set g_formatDefs = New Collection

    ' �f�[�^�ݒ�
    lastrow = ws.Cells(ws.Rows.Count, 2).End(xlUp).row
    For currentRow = ROW_NUMBER_DATA_START To lastrow
        Set formatDef = New FormatDefine
        
        ' �u��������̂���Z��
        formatDef.PRSReference.Initialize ws.Cells(currentRow, COLUMN_NUMBER_PRSREFERENCE).value
        formatDef.Data_Prerequisites.Initialize ws.Cells(currentRow, COLUMN_NUMBER_PREREQUISISITES).value
        formatDef.TestInstruction.Initialize ws.Cells(currentRow, COLUMN_NUMBER_TESTINSTRUCTION).value
        formatDef.ExpectedResult.Initialize ws.Cells(currentRow, COLUMN_NUMBER_EXPECTEDRESULT).value
        
        ' �u��������̂Ȃ��Z��
        formatDef.riskID = GetRiskID(ws.Cells(currentRow, COLUMN_NUMBER_RISKID).value)
        formatDef.TestResult = ws.Cells(currentRow, COLUMN_NUMBER_TESTRESULT).value
        formatDef.Evidence = ws.Cells(currentRow, COLUMN_NUMBER_EVIDENCE).value

        g_formatDefs.Add formatDef
    Next currentRow
End Sub


