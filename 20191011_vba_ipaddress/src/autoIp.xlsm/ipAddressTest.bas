Attribute VB_Name = "ipAddressTest"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")

Private wb As Workbook
Private ws As Worksheet

Private Assert As Object
Private Fakes As Object

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
    
    Application.ScreenUpdating = False
    
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    'this method runs before every test in the module.
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub
'@TestMethod("Uncategorized")
Private Sub testSetIpAddressAndSubnetMask()
    ' テストケース：引数で渡した値がそのままセットされること
    On Error GoTo TestFail
    
    'Arrange:

    'Act:
    Dim obj As New IpAddress
    obj.setIpAddress ("192.168.55.3")
    obj.setSubnetMask ("1.2.3.4")
    Dim ipAddressOctets() As Integer: ipAddressOctets = obj.getIpAddressOctets
    Dim subnetMaskOctets() As Integer: subnetMaskOctets = obj.getSubnetMaskOctets
    'Assert:
    
    Call Assert.AreEqual(192, ipAddressOctets(0))
    Call Assert.AreEqual(168, ipAddressOctets(1))
    Call Assert.AreEqual(55, ipAddressOctets(2))
    Call Assert.AreEqual(3, ipAddressOctets(3))
    Call Assert.AreEqual(1, subnetMaskOctets(0))
    Call Assert.AreEqual(2, subnetMaskOctets(1))
    Call Assert.AreEqual(3, subnetMaskOctets(2))
    Call Assert.AreEqual(4, subnetMaskOctets(3))

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Uncategorized")
Private Sub testSubnetMaskLength()
    ' テストケース：サブネットマスク長が正しいこと
    On Error GoTo TestFail
    
    'Arrange:

    'Act:
    Dim obj As New IpAddress
    
    obj.setSubnetMask ("128.0.0.0")
    Call Assert.AreEqual(1, obj.getSubnetMaskLength)
    
    obj.setSubnetMask ("255.255.255.0")
    Call Assert.AreEqual(24, obj.getSubnetMaskLength)
    
    obj.setSubnetMask ("255.255.255.192")
    Call Assert.AreEqual(26, obj.getSubnetMaskLength)
    

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Uncategorized")
Private Sub testNetworkAddress()
    ' テストケース：ネットワークアドレスが正しいこと
    On Error GoTo TestFail
    
    'Arrange:

    'Act:
    Dim obj As New IpAddress
    
    obj.setIpAddress ("192.168.123.170")
    obj.setSubnetMask ("255.255.255.0")
    obj.deriveNetworkAddress
    Call Assert.AreEqual("192.168.123.0", obj.getNetworkAddress)
    
    obj.setSubnetMask ("255.255.255.240")
    obj.deriveNetworkAddress
    Call Assert.AreEqual("192.168.123.160", obj.getNetworkAddress)
    

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub



