Attribute VB_Name = "C_PRSSelectModule"
' �菇C�@PRS�t�@�C���I�����̋���
Sub PRS_Select()
    Call ShowFileDialog
End Sub

Sub ShowFileDialog()
    Dim fd As FileDialog
    Dim filePath As String
    Dim initialPath As String
    
    ' �}�N���t�@�C���̃t�H���_�������p�X�Ƃ��Đݒ�
    initialPath = ThisWorkbook.path
    
    ' �t�@�C���_�C�A���O���쐬
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    ' �_�C�A���O�̃^�C�g����ݒ�
    fd.Title = "�t�@�C����I�����Ă�������"
    
    ' �����p�X��ݒ�
    fd.InitialFileName = initialPath & "\"

    ' �t�@�C���t�B���^�[��ݒ�
    fd.Filters.Clear
    fd.Filters.Add "Excel�t�@�C��", "*.xls; *.xlsx"
    
    ' �_�C�A���O��\�����āA���[�U�[���t�@�C����I������
    If fd.Show = -1 Then
        ' �I�������t�@�C���̃t���p�X���擾
        filePath = fd.SelectedItems(1)
        
        ' �w�肳�ꂽ�Z���Ƀt���p�X�����
        ThisWorkbook.Sheets(SHEET_CREATE_TEST).Range(CELL_SOURCE_FILE).value = filePath
    End If
End Sub
