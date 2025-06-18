VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SelectSheetDialog 
   Caption         =   "�V�[�g��I�����Ă�������"
   ClientHeight    =   3810
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3450
   OleObjectBlob   =   "SelectSheetDialog.frx":0000
End
Attribute VB_Name = "SelectSheetDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' SelectSheetDialog�̃t�H�[��
' �V�[�g�I���_�C�A���O�p
Option Explicit

' �����o�ϐ�
Private SelectedSheetName As String
Private filePath As String
Private Initialized As Boolean


Private Sub UserForm_Activate()
    ' ����̂ݎ��s���邽�߂̐���
    If Not Initialized Then
        If InitializeForm Then
            Initialized = True
        Else
            Me.Hide
        End If
    End If
End Sub


Private Function InitializeForm() As Boolean
    Dim wb As Workbook
    Dim i As Integer
    
    ' �t�H�[���̈ʒu���A�N�e�B�u�E�B���h�E�̒����ɂ���
    CenterForm Me

    If filePath = "" Then
        MsgBox "�t�@�C���p�X���w�肳��Ă��܂���B", vbExclamation
        InitializeForm = False
        Exit Function
    End If

    On Error Resume Next
    Application.ScreenUpdating = False
    Set wb = Workbooks.Open(filename:=filePath, ReadOnly:=True, Notify:=False)
    Application.ScreenUpdating = True
    On Error GoTo 0
    
    If wb Is Nothing Then
        MsgBox "�t�@�C����������܂���B�p�X���m�F���Ă��������B", vbExclamation
        InitializeForm = False
        Exit Function
    End If
    
    ' �V�[�g�����擾�i��\���������j
    For i = 1 To wb.Sheets.Count
        If wb.Sheets(i).Visible = xlSheetVisible Then
            Me.ListBox1.AddItem wb.Sheets(i).Name
        End If
    Next i
    
    ' �J�������[�N�u�b�N�����i���e��ێ����Ȃ��j
    wb.Close False
    
    InitializeForm = True
End Function

Private Sub OKButton_Click()
    ' �I�����ꂽ�V�[�g����ێ�����
    If Me.ListBox1.ListIndex <> -1 Then
        SelectedSheetName = Me.ListBox1.value
        Me.Hide
    Else
        MsgBox "�V�[�g��I�����Ă��������B", vbExclamation
    End If
End Sub


Public Property Let SourceFilePath(ByVal path As String)
    filePath = path
End Property

Public Property Get SelectedSheet() As String
    SelectedSheet = SelectedSheetName
End Property
