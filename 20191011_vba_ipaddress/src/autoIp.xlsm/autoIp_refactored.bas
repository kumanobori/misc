Attribute VB_Name = "autoIp_refactored"
Option Explicit

Sub AutoIP()


    Dim CIDR_Dic As Object
    Set CIDR_Dic = makeCIDR_Dic
    
    Dim Keys() As Variant
    Keys = CIDR_Dic.Keys

    Dim MaxRow_B As Integer
    MaxRow_B = Range("B3").End(xlDown).Row

 '///B列(IPアドレス)と、C列(サブネット)の入力を元に、D列(IP+CIDR)、E列(ネットワークアドレス)を記入するツール///

    Dim EachRow As Integer
    For EachRow = 3 To MaxRow_B
        Dim Cidr As Integer
        Cidr = 0
        Dim Subnet As Variant
        Subnet = Split(Worksheets("Sheet1").Cells(EachRow, 3), ".")
        Dim IP As Variant
        IP = Split(Worksheets("Sheet1").Cells(EachRow, 2), ".")

        Dim IPCount As Integer
        For IPCount = 0 To 3

            Dim Num As Integer
            Dim KeyCount As Integer
            For KeyCount = 0 To 8
                If Int(Subnet(IPCount)) = Keys(KeyCount) Then
                    Num = CIDR_Dic.Item(Keys(KeyCount))
                ElseIf CIDR_Dic.Exists(Int(Subnet(IPCount))) = False Then
                MsgBox (EachRow & "行目のサブネットにエラーがあります")
                End
                End If
            Next KeyCount

            Cidr = Cidr + Num
        Next IPCount

        Worksheets("Sheet1").Cells(EachRow, 4) = (Worksheets("Sheet1").Cells(EachRow, 2)) + "/" + LTrim(str(Cidr))
        
        Dim wr As Range: Set wr = Worksheets("Sheet1").Cells(EachRow, 5)
        If Cidr >= 24 And Cidr < 32 Then
            Call writeNetworkAddress(3, Subnet, IP, wr)
        ElseIf Cidr >= 16 And Cidr < 24 Then
            Call writeNetworkAddress(2, Subnet, IP, wr)
        ElseIf Cidr >= 8 And Cidr < 16 Then
            Call writeNetworkAddress(1, Subnet, IP, wr)
        ElseIf Cidr >= 0 And Cidr < 8 Then
            Call writeNetworkAddress(0, Subnet, IP, wr)
        Else
            wr.Value = Worksheets("Sheet1").Cells(EachRow, 2)
        End If
    Next EachRow


End Sub

Function makeCIDR_Dic()
    Dim dic As Object
    Set dic = CreateObject("Scripting.Dictionary")

    dic.Add 0, 0
    dic.Add 128, 1
    dic.Add 192, 2
    dic.Add 224, 3
    dic.Add 240, 4
    dic.Add 248, 5
    dic.Add 252, 6
    dic.Add 254, 7
    dic.Add 255, 8
    
    Set makeCIDR_Dic = dic
End Function
Function writeNetworkAddress(octetSequence As Integer, Subnet As Variant, IP As Variant, wr As Range)
    Dim oc As Variant
    oc = IP(octetSequence) / (256 - Int(Subnet(octetSequence)))
    oc = (Application.WorksheetFunction.RoundUp(oc, 0) - 1) * (256 - Int(Subnet(octetSequence)))
    
    Dim strAddress As String: strAddress = ""
    
    strAddress = strAddress & IIf(octetSequence = 0, LTrim(str(oc)), IP(0))
    strAddress = strAddress & "." & IIf(octetSequence = 1, LTrim(str(oc)), IP(1))
    strAddress = strAddress & "." & IIf(octetSequence = 2, LTrim(str(oc)), IP(2))
    strAddress = strAddress & "." & IIf(octetSequence = 3, LTrim(str(oc)), IP(3))
    
    wr.Value = strAddress
End Function

