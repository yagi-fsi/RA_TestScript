VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PhaseInformation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' PhaseInformation クラスモジュール
' PRSファイル内の1Phaseの情報を保持する
Option Explicit

' フロー種別
Public Enum FlowKindEnum
    LoopStart_Type
    LoopEnd_Type
    Branch_Type
    Normal_Type
End Enum

' メンバ変数
Private m_ID As String
Private m_PhaseIntroduction As String
Private m_PhaseName As String
Private m_Comment As String
Private m_FlowKind As FlowKindEnum
Private m_FlowInformation As String
Private m_RecipeParameter As String
Private m_SOPLink As String
Private m_Material As String
Private m_Equipment As String
Private m_Place As String
Private m_GMP As String

' コンストラクタ（設定は生成時のみ）
Private Sub Class_Initialize()
    ' 初期化を必要に応じて行う
End Sub

' 初期化用
Public Sub Initialize(ID As String, PhaseIntroduction As String, Comment As String, InfoString As String, _
                     RecipeParameter As String, SOPLink As String, Material As String, _
                     Equipment As String, Place As String, GMP As String)
    
    m_ID = ID
    m_PhaseIntroduction = PhaseIntroduction
    m_PhaseName = GetPhaseName(PhaseIntroduction)
    m_Comment = Comment
    SetFlowKindAndInformation InfoString    'm_FlowKind/m_FlowInformation設定
    m_RecipeParameter = RecipeParameter
    m_SOPLink = SOPLink
    m_Material = Material
    m_Equipment = Equipment
    m_Place = Place
    m_GMP = GMP
End Sub

' フロー情報文字列を解析し、対応するフロー種別とフロー情報を設定します。
Private Sub SetFlowKindAndInformation(InfoString As String)
    ' 定数: 各フロー種別の条件を識別するためのプレフィックス
    Dim LoopConditionPrefix As String
    Dim LoopEndConditionPrefix As String
    Dim BranchConditionPrefix As String
    
    ' プレフィックスの初期化
    LoopConditionPrefix = "ループ条件："
    LoopEndConditionPrefix = "ループ終了条件："
    BranchConditionPrefix = "分岐条件："
    
    ' InfoStringの先頭に基づいてフロー種別と情報を決定
    If Left(InfoString, Len(LoopConditionPrefix)) = LoopConditionPrefix Then
        ' ループ開始
        m_FlowKind = LoopStart_Type
        ' プレフィックスを除いたフロー情報を設定（以降も同様）
        m_FlowInformation = Mid(InfoString, Len(LoopConditionPrefix) + 1)
        
    ElseIf Left(InfoString, Len(LoopEndConditionPrefix)) = LoopEndConditionPrefix Then
        ' ループ終了
        m_FlowKind = LoopEnd_Type
        m_FlowInformation = Mid(InfoString, Len(LoopEndConditionPrefix) + 1)
        
    ElseIf Left(InfoString, Len(BranchConditionPrefix)) = BranchConditionPrefix Then
        ' 分岐
        m_FlowKind = Branch_Type
        m_FlowInformation = Mid(InfoString, Len(BranchConditionPrefix) + 1)
        
    Else
        ' 特定のプレフィックスがない場合、通常フローと見なす
        m_FlowKind = Normal_Type
        ' 文字列全体をフロー情報とする
        m_FlowInformation = InfoString
    End If
End Sub


' フェーズ名取得
Function GetPhaseName(inputString As String) As String
    Dim splitArray() As String
    Dim separator As String
    Dim firstWord As String
    
    ' 「Skip」から始まる場合はそのまま返す
    If Left(inputString, 4) = "Skip" Then
        GetPhaseName = "Skip"
        Exit Function
    End If
    
    ' 半角スペースと全角スペースを同じセパレータとして扱う
    inputString = Replace(inputString, "　", " ") ' 全角スペースを半角に変換
    
    ' 文字列を半角スペースで分割
    splitArray = Split(inputString, " ")
    
    ' 半角スペースや全角スペースがない場合は空文字を返す
    If UBound(splitArray) < 0 Then
        GetPhaseName = ""
    Else
        firstWord = splitArray(0)
        ' 最初の文字列を返す
        GetPhaseName = firstWord
    End If
End Function


' メンバ変数を返す関数
Public Function GetMemberValueByHeader(header As String) As String
    Select Case header
        Case HEADER_ID
            GetMemberValueByHeader = m_ID
        Case HEADER_PHASE_INTRODUCTION
            GetMemberValueByHeader = m_PhaseIntroduction
        Case HEADER_COMMENT
            GetMemberValueByHeader = m_Comment
        Case HEADER_RECIPE_PARAMETER
            GetMemberValueByHeader = m_RecipeParameter
        Case HEADER_MATERIAL
            GetMemberValueByHeader = m_Material
        Case HEADER_EQUIPMENT
            GetMemberValueByHeader = m_Equipment
        Case HEADER_PLACE
            GetMemberValueByHeader = m_Place
        Case HEADER_GMP
            GetMemberValueByHeader = m_GMP
        Case Else
            ' 不正なヘッダー文字列の場合は適切な処理を行う（エラーメッセージなど）
            GetMemberValueByHeader = "Invalid header"
    End Select
End Function

' ---------------------メンバアクセス用---------------------
Public Property Get ID() As String
    ID = m_ID
End Property
 
Public Property Let ID(value As String)
    m_ID = value
End Property

Public Property Get PhaseIntroduction() As String
    PhaseIntroduction = m_PhaseIntroduction
End Property

Public Property Let PhaseIntroduction(value As String)
    m_PhaseIntroduction = value
End Property

Public Property Get phaseName() As String
    phaseName = m_PhaseName
End Property

Public Property Let phaseName(value As String)
    m_PhaseName = value
End Property

Public Property Get Comment() As String
    Comment = m_Comment
End Property

Public Property Let Comment(value As String)
    m_Comment = value
End Property

Public Property Get FlowKind() As FlowKindEnum
    FlowKind = m_FlowKind
End Property

Public Property Let FlowKind(value As FlowKindEnum)
    m_FlowKind = value
End Property

Public Property Get FlowInformation() As String
    FlowInformation = m_FlowInformation
End Property

Public Property Let FlowInformation(value As String)
    SetFlowKindAndInformation value
End Property

Public Property Get RecipeParameter() As String
    RecipeParameter = m_RecipeParameter
End Property

Public Property Let RecipeParameter(value As String)
    m_RecipeParameter = value
End Property

Public Property Get SOPLink() As String
    SOPLink = m_SOPLink
End Property

Public Property Let SOPLink(value As String)
    m_SOPLink = value
End Property

Public Property Get Material() As String
    Material = m_Material
End Property

Public Property Let Material(value As String)
    m_Material = value
End Property

Public Property Get Equipment() As String
    Equipment = m_Equipment
End Property

Public Property Let Equipment(value As String)
    m_Equipment = value
End Property

Public Property Get Place() As String
    Place = m_Place
End Property

Public Property Let Place(value As String)
    m_Place = value
End Property

Public Property Get GMP() As String
    GMP = m_GMP
End Property

Public Property Let GMP(value As String)
    m_GMP = value
End Property
