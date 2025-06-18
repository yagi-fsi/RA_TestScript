Attribute VB_Name = "A_ConstantsModule"
' ���ʒ萔/���ʊ֐����`����ꏊ
Option Explicit

' -------------Enum��`--------------------------
' ���X�NID
Public Enum riskIDEnum
    RiskID_None = -1
    RiskID_A1 = 0
    RiskID_A2 = 1
    RiskID_B1 = 2
    RiskID_B2 = 3
    RiskID_B3 = 4
    RiskID_B4 = 5
    RiskID_B5 = 6
    RiskID_B6 = 7
    RiskID_C1 = 8
    RiskID_C2 = 9
    RiskID_C3 = 10
    RiskID_C4 = 11
    RiskID_C5 = 12
    RiskID_D1 = 13
    RiskID_D2 = 14
    RiskID_D3 = 15
    RiskID_D4 = 16
    RiskID_D5 = 17
    RiskID_E1 = 18
    RiskID_E2 = 19
    RiskID_E3 = 20
    RiskID_F1 = 21
    RiskID_F2 = 22
    RiskID_F3 = 23
    RiskID_F4 = 24
    RiskID_G1 = 25
    RiskID_H1 = 26
    RiskID_I1 = 27
    RiskID_I2 = 28
    RiskID_I3 = 29
    RiskID_J1 = 30
    RiskID_K1 = 31
    RiskID_L1 = 32
End Enum

' �o���h���p�����[�^�^�C�v�j
Public Enum BandleTypeEnum
    BandleType_Boolean
    BandleType_DateTime
    BandleType_DateTimeWithCulculate
    BandleType_Duration
    BandleType_DurationWithCulculate
    BandleType_MeasuredValue
    BandleType_MeasuredValueWithCulculate
    BandleType_ProcessValue
    BandleType_ProcessValueWithCulculate
    BandleType_String
    BandleType_Timestamp
    BandleType_TimestampWithCulculate
End Enum

' -------------�萔��`--------------------------
' �V�[�g��
Public Const SHEET_CREATE_TEST As String = "�e�X�g�쐬"
Public Const SHEET_MATRIX As String = "Matrix"
Public Const SHEET_FORMAT As String = "Format"


' �ΏۃZ��
Public Const CELL_SOURCE_FILE As String = "G14"
Public Const CELL_SOURCE_SHEET As String = "G19"


' �萔��`
Public Const HEADER_UP As String = "�H��: UP"
Public Const HEADER_OP As String = "�T�u�H��: OP"
Public Const HEADER_ID As String = "ID"
Public Const HEADER_PHASE_INTRODUCTION As String = "�H��: PH"
Public Const HEADER_COMMENT As String = "Comment"
Public Const HEADER_RECIPE_PARAMETER As String = "���V�s�p�����[�^"
Public Const HEADER_MATERIAL As String = "�}�e���A��"
Public Const HEADER_EQUIPMENT As String = "�@��"
Public Const HEADER_PLACE As String = "�ꏊ"
Public Const HEADER_GMP As String = "GMP����"


' -------------���ʊ֐�-------------------------
' �����Ƀt�H�[����z�u����֐�
Public Sub CenterForm(frm As Object)
    With Application
        frm.Top = .Top + ((.Height - frm.Height) / 2)
        frm.Left = .Left + ((.Width - frm.Width) / 2)
    End With
End Sub


' �f�[�^�`�F�b�N�p
Function BaseDataCheck(ParamArray args() As Variant) As Boolean
    Dim i As Integer
    Dim errorList As String
    Dim hasError As Boolean
    
    ' �����l�Ƃ���True��ݒ�
    BaseDataCheck = True
    hasError = False
    errorList = "�ȉ��̈������s���ł�:" & vbCrLf
    
    ' �e������Nothing���ǂ������`�F�b�N
    For i = LBound(args) To UBound(args)
        If IsNothing(args(i)) Then
            hasError = True
            errorList = errorList & "���� " & (i + 1) & " ��Nothing" & vbCrLf
        End If
    Next i
    
    ' �����G���[������ꍇ�A�܂Ƃ߂ăG���[���b�Z�[�W��\��
    If hasError Then
        MsgBox errorList, vbExclamation
        BaseDataCheck = False
    End If
End Function

Function IsNothing(var As Variant) As Boolean
    ' Nothing�̏ꍇ��True��Ԃ�
    IsNothing = (VarType(var) = vbObject) And (var Is Nothing)
End Function


' �V�[�g�R�s�[�p
Function CopySheet(ByVal sourceSheetName As String, ByVal destinationSheetBaseName As String) As Worksheet
    Dim srcSheet As Worksheet
    Dim newSheetName As String
    Dim counter As Integer
    Dim ws As Worksheet
    
    On Error Resume Next
    Set srcSheet = ThisWorkbook.Sheets(sourceSheetName)
    On Error GoTo 0
    
    ' �R�s�[���V�[�g���Ȃ��ꍇ��Nothing��Ԃ�
    If srcSheet Is Nothing Then
        Set CopySheet = Nothing
        Exit Function
    End If

    ' �R�s�[��V�[�g��������
    newSheetName = destinationSheetBaseName
    counter = 1
    Do While SheetExists(newSheetName)
        newSheetName = destinationSheetBaseName & "_" & Format(counter, "00")
        counter = counter + 1
    Loop

    ' �V�[�g�R�s�[�i�擪�ɑ}���j
    srcSheet.Copy Before:=ThisWorkbook.Sheets(1)
    Set ws = ThisWorkbook.Sheets(1)
    ws.Name = newSheetName

    ' �R�s�[���ꂽ�V�[�g�I�u�W�F�N�g��Ԃ�
    Set CopySheet = ws
End Function

Function SheetExists(sheetName As String) As Boolean
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    
    SheetExists = Not ws Is Nothing
End Function


' �X�e�b�v�ԍ��ݒ�
Sub SetStepNumber(ByRef wsTestScript As Worksheet)
    Dim rowNumber As Long
    Dim lastrow As Long
    Dim stepCounter As Long
    
    ' �ŏ���2�s�ڂ��珑�����݊J�n
    rowNumber = 2
    
    ' �K���ȏ����l��ݒ�A�����ł�1����J�n
    stepCounter = 1
    
    ' �Ō�̍s���擾
    lastrow = wsTestScript.Cells(wsTestScript.Rows.Count, "G").End(xlUp).row

    ' �w�肳�ꂽ�͈͂����[�v
    For rowNumber = 2 To lastrow
        ' G�񂪋󕶎��܂���"N/A"�łȂ���Ώ������s��
        If wsTestScript.Cells(rowNumber, "G").value <> "" And _
           wsTestScript.Cells(rowNumber, "G").value <> "N/A" Then
           
            ' A���"StepXXXX"��XXXX��4���̐���������
            wsTestScript.Cells(rowNumber, "A").value = "Step" & Format(stepCounter, "0000")
            
            ' �X�e�b�v�J�E���^�[���C���N�������g
            stepCounter = stepCounter + 1
        End If
    Next rowNumber
End Sub


' ���X�NID������@���@���X�gIDEnum�l
Public Function GetRiskID(ByVal riskIDStr As String) As riskIDEnum
    ' �n�C�t������菜��
    Dim cleanStr As String
    cleanStr = Replace(riskIDStr, "-", "")

    Select Case cleanStr
        Case "A1": GetRiskID = RiskID_A1
        Case "A2": GetRiskID = RiskID_A2
        Case "B1": GetRiskID = RiskID_B1
        Case "B2": GetRiskID = RiskID_B2
        Case "B3": GetRiskID = RiskID_B3
        Case "B4": GetRiskID = RiskID_B4
        Case "B5": GetRiskID = RiskID_B5
        Case "B6": GetRiskID = RiskID_B6
        Case "C1": GetRiskID = RiskID_C1
        Case "C2": GetRiskID = RiskID_C2
        Case "C3": GetRiskID = RiskID_C3
        Case "C4": GetRiskID = RiskID_C4
        Case "C5": GetRiskID = RiskID_C5
        Case "D1": GetRiskID = RiskID_D1
        Case "D2": GetRiskID = RiskID_D2
        Case "D3": GetRiskID = RiskID_D3
        Case "D4": GetRiskID = RiskID_D4
        Case "D5": GetRiskID = RiskID_D5
        Case "E1": GetRiskID = RiskID_E1
        Case "E2": GetRiskID = RiskID_E2
        Case "E3": GetRiskID = RiskID_E3
        Case "F1": GetRiskID = RiskID_F1
        Case "F2": GetRiskID = RiskID_F2
        Case "F3": GetRiskID = RiskID_F3
        Case "F4": GetRiskID = RiskID_F4
        Case "G1": GetRiskID = RiskID_G1
        Case "H1": GetRiskID = RiskID_H1
        Case "I1": GetRiskID = RiskID_I1
        Case "I2": GetRiskID = RiskID_I2
        Case "I3": GetRiskID = RiskID_I3
        Case "J1": GetRiskID = RiskID_J1
        Case "K1": GetRiskID = RiskID_K1
        Case "L1": GetRiskID = RiskID_L1
        Case Else: GetRiskID = RiskID_None
    End Select
End Function


' ���X�NIDEnum�l�@���@���X�NID������i�n�C�t���t���j
Public Function GetRiskIDString(ByVal riskIDEnum As riskIDEnum) As String
    Select Case riskIDEnum
        Case RiskID_A1: GetRiskIDString = "A-1"
        Case RiskID_A2: GetRiskIDString = "A-2"
        Case RiskID_B1: GetRiskIDString = "B-1"
        Case RiskID_B2: GetRiskIDString = "B-2"
        Case RiskID_B3: GetRiskIDString = "B-3"
        Case RiskID_B4: GetRiskIDString = "B-4"
        Case RiskID_B5: GetRiskIDString = "B-5"
        Case RiskID_B6: GetRiskIDString = "B-6"
        Case RiskID_C1: GetRiskIDString = "C-1"
        Case RiskID_C2: GetRiskIDString = "C-2"
        Case RiskID_C3: GetRiskIDString = "C-3"
        Case RiskID_C4: GetRiskIDString = "C-4"
        Case RiskID_C5: GetRiskIDString = "C-5"
        Case RiskID_D1: GetRiskIDString = "D-1"
        Case RiskID_D2: GetRiskIDString = "D-2"
        Case RiskID_D3: GetRiskIDString = "D-3"
        Case RiskID_D4: GetRiskIDString = "D-4"
        Case RiskID_D5: GetRiskIDString = "D-5"
        Case RiskID_E1: GetRiskIDString = "E-1"
        Case RiskID_E2: GetRiskIDString = "E-2"
        Case RiskID_E3: GetRiskIDString = "E-3"
        Case RiskID_F1: GetRiskIDString = "F-1"
        Case RiskID_F2: GetRiskIDString = "F-2"
        Case RiskID_F3: GetRiskIDString = "F-3"
        Case RiskID_F4: GetRiskIDString = "F-4"
        Case RiskID_G1: GetRiskIDString = "G-1"
        Case RiskID_H1: GetRiskIDString = "H-1"
        Case RiskID_I1: GetRiskIDString = "I-1"
        Case RiskID_I2: GetRiskIDString = "I-2"
        Case RiskID_I3: GetRiskIDString = "I-3"
        Case RiskID_J1: GetRiskIDString = "J-1"
        Case RiskID_K1: GetRiskIDString = "K-1"
        Case RiskID_L1: GetRiskIDString = "L-1"
        Case Else: GetRiskIDString = "None"
    End Select
End Function

