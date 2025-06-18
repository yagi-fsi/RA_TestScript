Attribute VB_Name = "Z_DebugModule"
' デバッグ用

Sub DebugFormatDefs()
    Dim formatDef As FormatDefine
    Dim i As Integer

    ' コレクションの内容をデバッグ出力
    For i = 1 To g_formatDefs.Count
        Set formatDef = g_formatDefs(i)
        
        Debug.Print "Record " & i & ": PRSReference=" & formatDef.PRSReference.baseString & _
                    ", Data_Prerequisites=" & formatDef.Data_Prerequisites.baseString & _
                    ", TestInstruction=" & formatDef.TestInstruction.baseString & _
                    ", ExpectedResult=" & formatDef.ExpectedResult.baseString & _
                    ", RiskID=" & formatDef.riskID & _
                    ", TestResult=" & formatDef.TestResult & _
                    ", Evidence=" & formatDef.Evidence

        ' 各 FormatSettingValue の追加テスト結果出力
        TestFormatSettingValue2 formatDef.PRSReference
        TestFormatSettingValue2 formatDef.Data_Prerequisites
        TestFormatSettingValue2 formatDef.TestInstruction
        TestFormatSettingValue2 formatDef.ExpectedResult
    Next i
End Sub


Sub DebugOPInformationList(opInfoList As Collection)
    Dim opInfo As OPInformation
    Dim phaseInfo As PhaseInformation
    Dim index As Long
    Dim filePath As String
    Dim fileNum As Integer
    
    ' ファイルに出力する
    filePath = ThisWorkbook.path & "\Debug_" & Format(Now, "MMddhhmmss") & ".txt"
    fileNum = FreeFile
    Open filePath For Output As #fileNum
    Print #fileNum, "Total OP Information count: " & opInfoList.Count
    
    For index = 1 To opInfoList.Count
        Set opInfo = opInfoList.Item(index)
        Print #fileNum, FormatOPInformation(opInfo)
        
        For Each phaseInfo In opInfo.GetPhaseInformationList()
            Print #fileNum, FormatPhaseInformation(phaseInfo)
        Next phaseInfo
        
        Print #fileNum, "-----------------------------"
    Next index

    Close #fileNum
End Sub

Function FormatOPInformation(opInfo As OPInformation) As String
    Dim result As String
    result = "OP Name: " & opInfo.GetOPName() & ", "
    result = result & "CBB Name: " & opInfo.GetCBBName() & ", "
    result = result & "ID: " & opInfo.GetID() & ", "
    result = result & "Capability: " & opInfo.GetCapability()
    FormatOPInformation = result
End Function

Function FormatPhaseInformation(phaseInfo As PhaseInformation) As String
    Dim result As String
    result = "  Phase ID: " & phaseInfo.ID & ", "
    result = result & "Introduction: " & phaseInfo.PhaseIntroduction & ", "
    result = result & "Comment: " & phaseInfo.Comment & ", "
    result = result & "Flow Kind: " & phaseInfo.FlowKind & ", "
    result = result & "Material: " & phaseInfo.Material & ", "
    result = result & "Equipment: " & phaseInfo.Equipment & ", "
    result = result & "Place: " & phaseInfo.Place & ", "
    result = result & "GMP: " & phaseInfo.GMP
    FormatPhaseInformation = result
End Function



Sub TestFormatSettingValue(teststr As String)
    Dim formatSetting As FormatSettingValue
    Set formatSetting = New FormatSettingValue

    formatSetting.Initialize teststr

    Debug.Print "Base String:", formatSetting.baseString
    Debug.Print "Replace Count:", formatSetting.ReplaceCount

    Dim Item As Variant
    Debug.Print "Replace Targets:"
    For Each Item In formatSetting.ReplaceTargetList
        Debug.Print Item
    Next Item

    Debug.Print "Replace Types:"
    For Each Item In formatSetting.ReplaceTypeList
        Debug.Print Item
    Next Item

    Debug.Print "Conversion Methods:"
    For Each Item In formatSetting.ConversionMethodList
        Debug.Print Item
    Next Item
End Sub



Sub TestFormatSettingValue2(testValue As FormatSettingValue)
    Debug.Print "Base String:", testValue.baseString
    Debug.Print "Replace Count:", testValue.ReplaceCount

    Dim Item As Variant
    Debug.Print "Replace Targets:"
    For Each Item In testValue.ReplaceTargetList
        Debug.Print Item
    Next Item

    Debug.Print "Replace Types:"
    For Each Item In testValue.ReplaceTypeList
        Debug.Print Item
    Next Item

    Debug.Print "Conversion Methods:"
    For Each Item In testValue.ConversionMethodList
        Debug.Print Item
    Next Item
End Sub
