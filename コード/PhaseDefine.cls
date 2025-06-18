VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PhaseDefine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' PhaseDefine クラスモジュール
Option Explicit

' メンバ変数
Private m_PhaseName As String
Private m_Version As Double
Private m_ShortDescription As String
Private m_RiskIDs() As riskIDEnum
Private m_BandleType As BandleTypeEnum

' コンストラクタ（設定は生成時のみ）
Private Sub Class_Initialize()
    ' 初期化を必要に応じて行う
End Sub

' 初期化用
Public Sub Initialize(ByVal phaseName As String, ByVal version As Double, ByVal ShortDescription As String, ByRef riskIDs() As riskIDEnum, ByVal bandleType As BandleTypeEnum)
    m_PhaseName = phaseName
    m_Version = version
    m_ShortDescription = ShortDescription
    m_RiskIDs = riskIDs
    m_BandleType = bandleType
End Sub

' ---------------------メンバアクセス用---------------------
Public Property Get phaseName() As String
    phaseName = m_PhaseName
End Property

Public Property Get version() As Double
    version = m_Version
End Property

Public Property Get ShortDescription() As String
    ShortDescription = m_ShortDescription
End Property

Public Property Get riskIDs() As Variant
    riskIDs = m_RiskIDs
End Property

Public Property Get bandleType() As BandleTypeEnum
    bandleType = m_BandleType
End Property
