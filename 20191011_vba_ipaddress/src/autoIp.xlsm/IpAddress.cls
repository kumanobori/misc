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

Private IpAddress As String ' �A�h���X�̕�����
Private ipAddressOctets() As Integer ' �I�N�e�b�g���l�̔z��

Private subnetMask As String ' �A�h���X�̕�����
Private subnetMaskOctets() As Integer ' �I�N�e�b�g���l�̔z��
Private subnetMaskLength As Integer ' �T�u�l�b�g�}�X�N�̃r�b�g���BCIDR�\�L�ɗp����B

Private networkAddress As String ' �A�h���X�̕�����
Private networkAddressOctets() As Integer ' �I�N�e�b�g���l�̔z��

Public Function setIpAddress(str As String)
    IpAddress = str
    ipAddressOctets = addressToOctets(str)
End Function

' IP�A�h���X��������A�I�N�e�b�g�̐��l�z��ɕϊ�����
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
        ' �܂��I�N�e�b�g�l��128�Ƃ�AND���Z���s���B
        ' ���ʂ��^�Ȃ�T�u�l�b�g�}�X�N����1���Z���A����64�Ƃ�AND���Z�E�E�E�Ɣ��X�ɂ��Ă����A1�܂Ŏ��{����B
        ' �U�̏ꍇ�͂��̎��_�ŏI���B
        Dim bitDecimal As Integer: bitDecimal = 128
        Do While ((bitDecimal And subnetMaskOctets(i)) = bitDecimal) And (bitDecimal >= 1)
            wkLength = wkLength + 1
            bitDecimal = bitDecimal / 2
        Loop
    Next i
    subnetMaskLength = wkLength
End Function

' IP�A�h���X�ƃT�u�l�b�g�}�X�N����A�l�b�g���[�N�A�h���X�𓱏o����
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

' �ȉ��A�e�X�g�p

Public Function getIpAddressOctets() As Integer()
    getIpAddressOctets = ipAddressOctets
End Function
Public Function getSubnetMaskOctets() As Integer()
    getSubnetMaskOctets = subnetMaskOctets
End Function
Public Function getSubnetMaskLength() As Integer
    getSubnetMaskLength = subnetMaskLength
End Function