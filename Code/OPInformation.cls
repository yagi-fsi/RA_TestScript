VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OPInformation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' OPInformation クラスモジュール
' PRSファイル内の1OPの情報を保持する
Option Explicit

' メンバ変数
Private m_OPName As String
Private m_CBBName As String
Private m_ID As String
Private m_PhaseInformationList As Collection
Private m_Capability As String

' コンストラクタ
Private Sub Class_Initialize()
    Set m_PhaseInformationList = New Collection
End Sub


' ---------------------メンバアクセス用---------------------
Public Function GetOPName() As String
    GetOPName = m_OPName
End Function

Public Sub SetOPName(OPName As String)
    m_OPName = OPName
End Sub

Public Function GetCBBName() As String
    GetCBBName = m_CBBName
End Function

Public Sub SetCBBName(CBBName As String)
    m_CBBName = CBBName
End Sub

Public Function GetID() As String
    GetID = m_ID
End Function

Public Sub SetID(ID As String)
    m_ID = ID
End Sub

Public Function GetCapability() As String
    GetCapability = m_Capability
End Function

Public Sub SetCapability(Capability As String)
    m_Capability = Capability
End Sub

Public Function GetPhaseInformationList() As Collection
    Set GetPhaseInformationList = m_PhaseInformationList
End Function

Public Sub AddPhaseInformation(phaseInfo As PhaseInformation)
    m_PhaseInformationList.Add phaseInfo
End Sub

Public Sub RemovePhaseInformation(index As Long)
    If index > 0 And index <= m_PhaseInformationList.Count Then
        m_PhaseInformationList.Remove index
    End If
End Sub
