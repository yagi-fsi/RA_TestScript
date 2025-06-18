VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FormatSettingValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' FormatSettingValue クラスモジュール
' 「Format」シートの各セル情報を保持する
Option Explicit

' メンバ変数の定義
' セル文字列（無加工）
Private m_BaseString As String

' セル文字列を置換対象/非置換対象で分けたリスト
' 例：m_BaseString　　　　：ああああ%Rep1%ID%ALL%%いいいい%Rep2%工程:PH%LINE%1%うううう
' 　　m_SplitStringList(0)：ああああ
' 　　m_SplitStringList(1)：%Rep1%ID%ALL%%
' 　　m_SplitStringList(2)：いいいい
' 　　m_SplitStringList(3)：%Rep2%工程:PH%LINE%1%
' 　　m_SplitStringList(4)：うううう
Private m_SplitStringList As Collection

' 置換対象数　※上記例では2が入る
Private m_ReplaceCount As Integer

' 置換対象リスト　※上記例では"ID","工程:PH"が入る
Private m_ReplaceTargetList As Collection

' 置換タイプリスト　※上記例ではALL,LINEが入る
Private m_ReplaceTypeList As Collection

' 変換方法リスト　※上記例では"","1"が入る
Private m_ConversionMethodList As Collection


' Enum型の定義
Public Enum replaceType
    LINE_TYPE
    SOPLINK_TYPE
    ALL
    UNDEFINE
End Enum


' コンストラクタ（何も行わない）
Private Sub Class_Initialize()
    Set m_SplitStringList = New Collection
    Set m_ReplaceTargetList = New Collection
    Set m_ReplaceTypeList = New Collection
    Set m_ConversionMethodList = New Collection
End Sub


' 初期化関数
' 引数の文字列（「Format」シートのセル文字列）を分解し、各メンバへ格納する
Public Sub Initialize(baseString As String)
    Dim percentDelimitedArray() As String

    ' ベース文字列の格納
    m_BaseString = baseString

    ' パーセント記号(%)で文字列を分解
    percentDelimitedArray = Split(baseString, "%")
    
    ' パーセント記号(%)が含まれていない場合は処理完了
    If UBound(percentDelimitedArray) = 0 Then
        m_SplitStringList.Add percentDelimitedArray(0)
        Exit Sub
    End If

    ' 各メンバへ設定
    PopulateMembersWithDelimitedStrings percentDelimitedArray
End Sub


' メンバ変数の定義で記載した通りにメンバが設定されるように処理を行う
Private Sub PopulateMembersWithDelimitedStrings(ByRef percentDelimitedArray() As String)
    Dim i As Integer
    m_ReplaceCount = 0

    ' 文字列を解析してリストを作成
    ' 下記のフォーマットで文字列が設定されているので、インデックスは5飛ばしでループを回す
    ' 　1番目：非置換対象文字
    ' 　2番目：識別子（"RepXX"）
    ' 　3番目：置換参照先（"ID","工程:PH"など）
    ' 　4番目：置換種別（"ALL","LINE"など）
    '　 5番目：変換方法（空白、"1"など）
    For i = 0 To UBound(percentDelimitedArray) - 1 Step 5
        Dim fixedString As String           ' 非置換対象文字
        Dim replaceTargetString As String   ' 置換対象文字
        
        ' 非置換対象文字をリストへ追加
        fixedString = percentDelimitedArray(i)
        m_SplitStringList.Add fixedString
        If i + 1 > UBound(percentDelimitedArray) Then Exit For

        ' 置換対象文字は上記2～5番目までの文字列＋"%"となる
        replaceTargetString = ""
        
        ' 識別子確認
        If i + 1 <= UBound(percentDelimitedArray) Then
            replaceTargetString = replaceTargetString + "%" + percentDelimitedArray(i + 1)
            If Left(percentDelimitedArray(i + 1), 3) = "Rep" Then
                m_ReplaceCount = m_ReplaceCount + 1
            End If
        End If
        
        ' 置換参照先確認／メンバへ格納
        If i + 2 <= UBound(percentDelimitedArray) Then
            replaceTargetString = replaceTargetString + "%" + percentDelimitedArray(i + 2)
            m_ReplaceTargetList.Add percentDelimitedArray(i + 2)
        End If

        ' 置換種別確認／メンバへ格納
        If i + 3 <= UBound(percentDelimitedArray) Then
            replaceTargetString = replaceTargetString + "%" + percentDelimitedArray(i + 3)
            Select Case percentDelimitedArray(i + 3)
                Case "LINE"
                    m_ReplaceTypeList.Add LINE_TYPE
                Case "SOPLINK"
                    m_ReplaceTypeList.Add SOPLINK_TYPE
                Case "ALL"
                    m_ReplaceTypeList.Add ALL
                Case Else
                    m_ReplaceTypeList.Add UNDEFINE
            End Select
        Else
            m_ReplaceTypeList.Add UNDEFINE
        End If

        ' 変換方法確認／メンバへ格納
        If i + 4 <= UBound(percentDelimitedArray) Then
            replaceTargetString = replaceTargetString + "%" + percentDelimitedArray(i + 4) + "%"
            m_ConversionMethodList.Add percentDelimitedArray(i + 4)
        End If
        
        m_SplitStringList.Add replaceTargetString
    Next i
    
    ' 最後の非置換対象文字を設定
    If UBound(percentDelimitedArray) = i Then
        m_SplitStringList.Add percentDelimitedArray(i)
    End If
End Sub


' 置換関数の定義
Public Function ReplaceStrings(replaceList As Collection) As String
    Dim resultString As String
    Dim replaceIndex As Integer
    Dim strValue As Variant

    ' 要素数が正しいか
    If replaceList.Count <> m_ReplaceCount Then
        MsgBox "置換リストの要素数が一致しません。"
    End If
    
    ' 結果用文字列を初期化
    resultString = ""
    replaceIndex = 1

    For Each strValue In m_SplitStringList
        Dim str As String
        str = CStr(strValue)
        
        If InStr(1, str, "%") > 0 Then
            Dim replaceType As replaceType
            replaceType = m_ReplaceTypeList(replaceIndex)
            
            Select Case replaceType
                Case LINE_TYPE
                    resultString = resultString + ReplaceLineType(replaceList(replaceIndex), m_ConversionMethodList(replaceIndex))
                Case SOPLINK_TYPE
                    resultString = resultString + ReplaceSopLinkType(replaceList(replaceIndex))
                Case ALL
                    resultString = resultString + ReplaceAll(replaceList(replaceIndex))
                Case Else
                    ' UNDEFINE等の未定義タイプの場合は何もしない
            End Select

            replaceIndex = replaceIndex + 1
        Else
            resultString = resultString + str
        End If
    Next strValue
    
    ' 結果の文字列を返す
    ReplaceStrings = resultString
End Function

Private Function ReplaceLineType(replaceString As String, conversionMethod As Variant) As String
    ' 行番号を取得し、指定された行番号の文字列を返す
    Dim lineNumber As Integer
    Dim splitLines() As String

    lineNumber = CInt(conversionMethod)
    splitLines = Split(replaceString, vbLf)
    
    If lineNumber <= UBound(splitLines) + 1 Then
        ReplaceLineType = splitLines(lineNumber - 1)
    Else
        ReplaceLineType = ""
    End If
End Function

Private Function ReplaceSopLinkType(ByVal replaceString As String) As String
    Dim lines() As String
    Dim result As String
    Dim i As Integer
    
    ' 文字列全体に "リンク：" が存在するか確認
    If InStr(replaceString, "リンク：") = 0 Then
        ReplaceSopLinkType = ""
        Exit Function
    End If
    
    ' 入力文字列を vbLf で分割して、個々の行に分ける
    lines = Split(replaceString, vbLf)
    
    ' 結果の文字列を初期化
    result = ""
    
    ' 各行を繰り返し処理
    For i = LBound(lines) To UBound(lines)
        ' 行に "リンク：" が含まれているか確認
        If InStr(lines(i), "リンク：") > 0 Then
            ' 行から "リンク：" を除去して結果に追加
            result = result & Replace(lines(i), "リンク：", "") & vbLf
        End If
    Next i
    
    ' もし結果に余分な vbLf があれば削除
    If Len(result) > 0 Then
        result = Left(result, Len(result) - Len(vbLf))
    End If
    
    ' 最終的な文字列を返す
    ReplaceSopLinkType = result
End Function

Private Function ReplaceAll(replaceString As String) As String
    ' すべての文字列を置換
    ReplaceAll = replaceString
End Function


' ---------------------メンバアクセス用---------------------
Public Property Get baseString() As String
    baseString = m_BaseString
End Property
Public Property Get ReplaceCount() As Integer
    ReplaceCount = m_ReplaceCount
End Property
Public Property Get SplitStringList() As Collection
    SplitStringList = m_SplitStringList
End Property
Public Property Get ReplaceTargetList() As Collection
    Set ReplaceTargetList = m_ReplaceTargetList
End Property
Public Property Get ReplaceTypeList() As Collection
    Set ReplaceTypeList = m_ReplaceTypeList
End Property
Public Property Get ConversionMethodList() As Collection
    Set ConversionMethodList = m_ConversionMethodList
End Property
