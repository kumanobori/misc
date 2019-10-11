Attribute VB_Name = "autoIp_original"
Option Explicit

Const ROW_DATA_START = 3
Const COL_IP_ADDRESS = 2
Const COL_SUBNET_MASK = 3
Const COL_CIDR = 4
Const COL_NETWORK_ADDRESS = 5

Sub AutoIP()
    
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim wkRow As Long: wkRow = ROW_DATA_START
    
    Do Until ws.Cells(wkRow, COL_IP_ADDRESS) = ""
        Dim IP As New IpAddress
        Call IP.setIpAddress(ws.Cells(wkRow, COL_IP_ADDRESS))
        Call IP.setSubnetMask(ws.Cells(wkRow, COL_SUBNET_MASK))
        Call IP.deriveNetworkAddress
        ws.Cells(wkRow, COL_CIDR).Value = IP.getCidrFormat
        ws.Cells(wkRow, COL_NETWORK_ADDRESS).Value = IP.getNetworkAddress
        wkRow = wkRow + 1
    Loop
    
    MsgBox "done"
End Sub

