VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SelectSheetDialog 
   Caption         =   "シートを選択してください"
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
' SelectSheetDialogのフォーム
' シート選択ダイアログ用
Option Explicit

' メンバ変数
Private SelectedSheetName As String
Private filePath As String
Private Initialized As Boolean


Private Sub UserForm_Activate()
    ' 初回のみ実行するための制御
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
    
    ' フォームの位置をアクティブウィンドウの中央にする
    CenterForm Me

    If filePath = "" Then
        MsgBox "ファイルパスが指定されていません。", vbExclamation
        InitializeForm = False
        Exit Function
    End If

    On Error Resume Next
    Application.ScreenUpdating = False
    Set wb = Workbooks.Open(filename:=filePath, ReadOnly:=True, Notify:=False)
    Application.ScreenUpdating = True
    On Error GoTo 0
    
    If wb Is Nothing Then
        MsgBox "ファイルが見つかりません。パスを確認してください。", vbExclamation
        InitializeForm = False
        Exit Function
    End If
    
    ' シート名を取得（非表示を除く）
    For i = 1 To wb.Sheets.Count
        If wb.Sheets(i).Visible = xlSheetVisible Then
            Me.ListBox1.AddItem wb.Sheets(i).Name
        End If
    Next i
    
    ' 開いたワークブックを閉じる（内容を保持しない）
    wb.Close False
    
    InitializeForm = True
End Function

Private Sub OKButton_Click()
    ' 選択されたシート名を保持する
    If Me.ListBox1.ListIndex <> -1 Then
        SelectedSheetName = Me.ListBox1.value
        Me.Hide
    Else
        MsgBox "シートを選択してください。", vbExclamation
    End If
End Sub


Public Property Let SourceFilePath(ByVal path As String)
    filePath = path
End Property

Public Property Get SelectedSheet() As String
    SelectedSheet = SelectedSheetName
End Property
