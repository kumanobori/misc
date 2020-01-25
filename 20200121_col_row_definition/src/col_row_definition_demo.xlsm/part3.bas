Attribute VB_Name = "part3"
Option Explicit

' --- �V�[�g��`
Const INPUT_SHEET_NAME = "INPUT"
Const OUTPUT_SHEET_NAME = "OUTPUT"

' --- ���͗��`
Enum COL_IN
    Index = 1
    familyName ' 2 (�l���w�肵�Ȃ���΁A�O�̍��ڂ���C���N�������g���ꂽ�l�ɂȂ�)
    lastName ' 3
    sex ' 4
    zip1 ' (�ȉ���)
    zip2
    prefecture
    city
    town
    building
    addresslistProhibited
End Enum

' --- ���͍s��`
Const ROW_IN_DATA_START = 2

' --- �o�͗��`
Enum COL_OUT
    Name = 1
    zip
    address
End Enum

' --- �o�͍s��`
Const ROW_OUT_DATA_START = 3


' �Z���e�[�u������A�Z���^���ۂ����̂�����B
' �Ώۂ̃u�b�N���A�N�e�B�u�ɂ�����ԂŋN������B
Sub tableToAddlessList()
    
    ' ���[�N�u�b�N�E���[�N�V�[�g�̒�`
    Dim wb As Workbook: Set wb = ActiveWorkbook
    Dim wsIn As Worksheet: Set wsIn = wb.Worksheets(INPUT_SHEET_NAME)
    Dim wsOut As Worksheet: Set wsOut = createOutputSheet(wb)
    
    ' ���͊J�n�s�E�o�͊J�n�s��ݒ�
    Dim rowNoOut As Long: rowNoOut = ROW_OUT_DATA_START
    Dim rowNoIn As Long: rowNoIn = ROW_IN_DATA_START
    
    ' ���͂�No.�񂪋�ɂȂ�܂ŁA1�s���Ƃɏ�������
    Do Until wsIn.Cells(rowNoIn, COL_IN.Index).Value = ""
        
        ' �o�͋֎~�łȂ���Ώo�͂���
        If wsIn.Cells(rowNoIn, COL_IN.addresslistProhibited).Value <> "Y" Then
            Call writeAddress(wsIn, wsOut, rowNoIn, rowNoOut)
            rowNoOut = rowNoOut + 1
        End If
        
        rowNoIn = rowNoIn + 1
    Loop
    
MsgBox "done"

End Sub


' �o�̓V�[�g���쐬����B���ɑ��݂���ꍇ�́A�폜���Ă���쐬���Ȃ����B
Function createOutputSheet(wb As Workbook)
    
    Dim wsDel As Worksheet
    For Each wsDel In wb.Worksheets
        If wsDel.Name = OUTPUT_SHEET_NAME Then
            Application.DisplayAlerts = False
            wsDel.Delete
            Application.DisplayAlerts = True
        End If
    Next wsDel
    
    wb.Worksheets.Add
    Dim wsOut As Worksheet: Set wsOut = ActiveSheet
    wsOut.Name = OUTPUT_SHEET_NAME
    
    Set createOutputSheet = wsOut
    
End Function

' �A�h���X1������������
Function writeAddress(wsIn As Worksheet, wsOut As Worksheet, rowNoIn As Long, rowNoOut As Long)
        
    ' ���O�F�c���{�󔒁{���O
    wsOut.Cells(rowNoOut, COL_OUT.Name).Value = wsIn.Cells(rowNoIn, COL_IN.familyName) _
                                              & " " & wsIn.Cells(rowNoIn, COL_IN.lastName)
    
    ' �X�֔ԍ��F��4�����Ȃ��ꍇ�ɍl��
    Dim zip1 As String: zip1 = wsIn.Cells(rowNoIn, COL_IN.zip1).Value
    Dim zip2 As String: zip2 = wsIn.Cells(rowNoIn, COL_IN.zip2).Value
    wsOut.Cells(rowNoOut, COL_OUT.zip).Value = IIf(zip2 = "", zip1, zip1 & "-" & zip2)
        
    ' �Z��
    Dim prefecture As String: prefecture = wsIn.Cells(rowNoIn, COL_IN.prefecture).Value
    Dim city As String: city = wsIn.Cells(rowNoIn, COL_IN.city).Value
    Dim town As String: town = wsIn.Cells(rowNoIn, COL_IN.town).Value
    Dim building As String: building = wsIn.Cells(rowNoIn, COL_IN.building).Value
    Dim address As String
    address = prefecture & city & town
    If building <> "" Then
        address = address & " " & building
    End If
    wsOut.Cells(rowNoOut, COL_OUT.address).Value = address
    
End Function




