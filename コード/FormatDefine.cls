VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FormatDefine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' FormatDefine クラスモジュール
' 「Format」シートの各行の情報（リスクIDごとのフォーマット情報）を保持する
Option Explicit

' メンバ変数の定義（各列に対応）
Private m_PRSReference As FormatSettingValue
Private m_Data_Prerequisites As FormatSettingValue
Private m_TestInstruction As FormatSettingValue
Private m_ExpectedResult As FormatSettingValue
Private m_TestResult As String
Private m_Evidence As String
Private m_RiskID As riskIDEnum

' コンストラクタ（何も行わない）
Private Sub Class_Initialize()
    Set m_PRSReference = New FormatSettingValue
    Set m_Data_Prerequisites = New FormatSettingValue
    Set m_TestInstruction = New FormatSettingValue
    Set m_ExpectedResult = New FormatSettingValue
End Sub


' ---------------------メンバアクセス用---------------------
Public Property Get PRSReference() As FormatSettingValue
    Set PRSReference = m_PRSReference
End Property

Public Property Set PRSReference(value As FormatSettingValue)
    Set m_PRSReference = value
End Property

Public Property Get Data_Prerequisites() As FormatSettingValue
    Set Data_Prerequisites = m_Data_Prerequisites
End Property

Public Property Set Data_Prerequisites(value As FormatSettingValue)
    Set m_Data_Prerequisites = value
End Property

Public Property Get TestInstruction() As FormatSettingValue
    Set TestInstruction = m_TestInstruction
End Property

Public Property Set TestInstruction(value As FormatSettingValue)
    Set m_TestInstruction = value
End Property

Public Property Get ExpectedResult() As FormatSettingValue
    Set ExpectedResult = m_ExpectedResult
End Property

Public Property Set ExpectedResult(value As FormatSettingValue)
    Set m_ExpectedResult = value
End Property

Public Property Get TestResult() As String
    TestResult = m_TestResult
End Property

Public Property Let TestResult(value As String)
    m_TestResult = value
End Property

Public Property Get Evidence() As String
    Evidence = m_Evidence
End Property

Public Property Let Evidence(value As String)
    m_Evidence = value
End Property

Public Property Get riskID() As riskIDEnum
    riskID = m_RiskID
End Property

Public Property Let riskID(value As riskIDEnum)
    m_RiskID = value
End Property
