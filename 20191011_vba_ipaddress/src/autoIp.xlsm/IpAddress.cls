VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IpAddress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private IpAddress As String ' アドレスの文字列
Private ipAddressOctets() As Integer ' オクテット数値の配列

Private subnetMask As String ' アドレスの文字列
Private subnetMaskOctets() As Integer ' オクテット数値の配列
Private subnetMaskLength As Integer ' サブネットマスクのビット長。CIDR表記に用いる。

Private networkAddress As String ' アドレスの文字列
Private networkAddressOctets() As Integer ' オクテット数値の配列

Public Function setIpAddress(str As String)
    IpAddress = str
    ipAddressOctets = addressToOctets(str)
End Function

' IPアドレス文字列を、オクテットの数値配列に変換する
Private Function addressToOctets(address As String) As Integer()
    
    Dim strOctets() As String: strOctets = Split(address, ".")
    Dim intOctets(0 To 3) As Integer
    Dim i As Long
    For i = 0 To 3
        intOctets(i) = Val(strOctets(i))
    Next i
    addressToOctets = intOctets
End Function

Public Function setSubnetMask(str As String)
    subnetMask = str
    subnetMaskOctets = addressToOctets(str)
    
    Dim wkLength As Integer: wkLength = 0
    Dim i As Long
    For i = 0 To 3
        ' まずオクテット値と128とのAND演算を行う。
        ' 結果が真ならサブネットマスク長を1加算し、次は64とのAND演算・・・と半々にしていき、1まで実施する。
        ' 偽の場合はその時点で終了。
        Dim bitDecimal As Integer: bitDecimal = 128
        Do While ((bitDecimal And subnetMaskOctets(i)) = bitDecimal) And (bitDecimal >= 1)
            wkLength = wkLength + 1
            bitDecimal = bitDecimal / 2
        Loop
    Next i
    subnetMaskLength = wkLength
End Function

' IPアドレスとサブネットマスクから、ネットワークアドレスを導出する
Public Function deriveNetworkAddress()

    Dim strOctets(0 To 3) As String
    Dim i As Long
    For i = 0 To 3
        strOctets(i) = Trim(str(ipAddressOctets(i) And subnetMaskOctets(i)))
    Next i
    networkAddress = Join(strOctets, ".")
    networkAddressOctets = addressToOctets(networkAddress)
    
End Function


Public Function getCidrFormat()
    getCidrFormat = IpAddress & "/" & subnetMaskLength
End Function

Public Function getNetworkAddress()
    getNetworkAddress = networkAddress
End Function

' 以下、テスト用

Public Function getIpAddressOctets() As Integer()
    getIpAddressOctets = ipAddressOctets
End Function
Public Function getSubnetMaskOctets() As Integer()
    getSubnetMaskOctets = subnetMaskOctets
End Function
Public Function getSubnetMaskLength() As Integer
    getSubnetMaskLength = subnetMaskLength
End Function
