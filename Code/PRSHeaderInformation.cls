VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PRSHeaderInformation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' PRSHeaderInformation クラスモジュール
' PRSファイルのヘッダ情報（列番号）を保持する

Option Explicit

' メンバ変数
Private m_Column_UP As Long
Private m_Column_OP As Long
Private m_Column_ID As Long
Private m_Column_PhaseIntroduction As Long
Private m_Column_Comment As Long
Private m_Column_RecipeParameter As Long
Private m_Column_Material As Long
Private m_Column_Equipment As Long
Private m_Column_Place As Long
Private m_Column_GMP As Long

' コンストラクタ（設定は生成時のみ）
Private Sub Class_Initialize()
    ' 何もしない
End Sub

' PRSファイルの該当シートから列番号を検索/保持
Public Sub Initialize(ws As Worksheet)
    m_Column_UP = FindColumn(ws, HEADER_UP)
    m_Column_OP = FindColumn(ws, HEADER_OP)
    m_Column_ID = FindColumn(ws, HEADER_ID)
    m_Column_PhaseIntroduction = FindColumn(ws, HEADER_PHASE_INTRODUCTION)
    m_Column_Comment = FindColumn(ws, HEADER_COMMENT)
    m_Column_RecipeParameter = FindColumn(ws, HEADER_RECIPE_PARAMETER)
    m_Column_Material = FindColumn(ws, HEADER_MATERIAL)
    m_Column_Equipment = FindColumn(ws, HEADER_EQUIPMENT)
    m_Column_Place = FindColumn(ws, HEADER_PLACE)
    m_Column_GMP = FindColumn(ws, HEADER_GMP)
End Sub

' 列ヘッダ名の列番号を返す関数
Private Function FindColumn(ws As Worksheet, headerName As String) As Long
    Dim c As Range
    For Each c In ws.Rows(1).Cells
        ' 長さだけ切り出して前方一致を確認
        If Left(c.value, Len(headerName)) = headerName Then
            FindColumn = c.Column
            Exit Function
        End If
    Next c
    FindColumn = 0 ' 見つからない場合
End Function

' 不正チェック
Public Function IsValid() As Boolean
    IsValid = (m_Column_UP <> 0) And _
              (m_Column_OP <> 0) And _
              (m_Column_ID <> 0) And _
              (m_Column_PhaseIntroduction <> 0) And _
              (m_Column_Comment <> 0) And _
              (m_Column_RecipeParameter <> 0) And _
              (m_Column_Material <> 0) And _
              (m_Column_Equipment <> 0) And _
              (m_Column_Place <> 0) And _
              (m_Column_GMP <> 0)
End Function

' 指定された文字列に該当する列番号を返す
Public Function GetColumnNumberByString(headerString As String) As Long
    Select Case headerString
        Case HEADER_UP
            GetColumnNumberByString = m_Column_UP
        Case HEADER_OP
            GetColumnNumberByString = m_Column_OP
        Case HEADER_ID
            GetColumnNumberByString = m_Column_ID
        Case HEADER_PHASE_INTRODUCTION
            GetColumnNumberByString = m_Column_PhaseIntroduction
        Case HEADER_COMMENT
            GetColumnNumberByString = m_Column_Comment
        Case HEADER_RECIPE_PARAMETER
            GetColumnNumberByString = m_Column_RecipeParameter
        Case HEADER_MATERIAL
            GetColumnNumberByString = m_Column_Material
        Case HEADER_EQUIPMENT
            GetColumnNumberByString = m_Column_Equipment
        Case HEADER_PLACE
            GetColumnNumberByString = m_Column_Place
        Case HEADER_GMP
            GetColumnNumberByString = m_Column_GMP
        Case Else
            GetColumnNumberByString = 0 ' 見つからない場合
    End Select
End Function


' ---------------------メンバアクセス用---------------------
Public Function GetColumnNumberUP() As Long
    GetColumnNumberUP = m_Column_UP
End Function

Public Function GetColumnNumberOP() As Long
    GetColumnNumberOP = m_Column_OP
End Function

Public Function GetColumnNumberID() As Long
    GetColumnNumberID = m_Column_ID
End Function

Public Function GetColumnNumberPhaseIntroduction() As Long
    GetColumnNumberPhaseIntroduction = m_Column_PhaseIntroduction
End Function

Public Function GetColumnNumberComment() As Long
    GetColumnNumberComment = m_Column_Comment
End Function

Public Function GetColumnNumberRecipeParameter() As Long
    GetColumnNumberRecipeParameter = m_Column_RecipeParameter
End Function

Public Function GetColumnNumberMaterial() As Long
    GetColumnNumberMaterial = m_Column_Material
End Function

Public Function GetColumnNumberEquipment() As Long
    GetColumnNumberEquipment = m_Column_Equipment
End Function

Public Function GetColumnNumberPlace() As Long
    GetColumnNumberPlace = m_Column_Place
End Function

Public Function GetColumnNumberGMP() As Long
    GetColumnNumberGMP = m_Column_GMP
End Function
