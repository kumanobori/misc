Attribute VB_Name = "part2"
Option Explicit

' --- �V�[�g��`
Const INPUT_SHEET_NAME = "INPUT"
Const OUTPUT_SHEET_NAME = "OUTPUT"

' --- ���͗��`
Const COL_IN_ITEM_START = 1
Dim COL_IN_INDEX As Long
Dim COL_IN_FAMILY_NAME As Long
Dim COL_IN_LAST_NAME As Long
Dim COL_IN_SEX As Long
Dim COL_IN_ZIP1 As Long
Dim COL_IN_ZIP2 As Long
Dim COL_IN_PREFECTURE As Long
Dim COL_IN_CITY As Long
Dim COL_IN_TOWN As Long
Dim COL_IN_BUILDING As Long
Dim COL_IN_ADDRESSLIST_PROHIBITED As Long

' --- ���͍s��`
Const ROW_IN_DATA_START = 2

' --- �o�͗��`
Const COL_OUT_ITEM_START = 1
Dim COL_OUT_NAME As Long
Dim COL_OUT_ZIP As Long
Dim COL_OUT_ADDRESS As Long

' --- �o�͍s��`
Const ROW_OUT_DATA_START = 3


' �Z���e�[�u������A�Z���^���ۂ����̂�����B
' �Ώۂ̃u�b�N���A�N�e�B�u�ɂ�����ԂŋN������B
Sub tableToAddlessList()
    
    ' ���[�N�u�b�N�E���[�N�V�[�g�̒�`
    Dim wb As Workbook: Set wb = ActiveWorkbook
    Dim wsIn As Worksheet: Set wsIn = wb.Worksheets(INPUT_SHEET_NAME)
    Dim wsOut As Worksheet: Set wsOut = createOutputSheet(wb)
    
    Call defineColNumber(wsIn, wsOut)
    
    ' ���͊J�n�s�E�o�͊J�n�s��ݒ�
    Dim rowNoOut As Long: rowNoOut = ROW_OUT_DATA_START
    Dim rowNoIn As Long: rowNoIn = ROW_IN_DATA_START
    
    ' ���͂�No.�񂪋�ɂȂ�܂ŁA1�s���Ƃɏ�������
    Do Until wsIn.Cells(rowNoIn, COL_IN_INDEX).Value = ""
        
        ' �o�͋֎~�łȂ���Ώo�͂���
        If wsIn.Cells(rowNoIn, COL_IN_ADDRESSLIST_PROHIBITED).Value <> "Y" Then
            Call writeAddress(wsIn, wsOut, rowNoIn, rowNoOut)
            rowNoOut = rowNoOut + 1
        End If
        
        rowNoIn = rowNoIn + 1
    Loop
    
MsgBox "done"

End Sub

' ��ԍ����`����
Function defineColNumber(wsIn As Worksheet, wsOut As Worksheet)
    
    Dim i As Long: i = COL_IN_ITEM_START - 1
    i = i + 1: COL_IN_INDEX = i
    i = i + 1: COL_IN_FAMILY_NAME = i
    i = i + 1: COL_IN_LAST_NAME = i
    i = i + 1: COL_IN_SEX = i
    i = i + 1: COL_IN_ZIP1 = i
    i = i + 1: COL_IN_ZIP2 = i
    i = i + 1: COL_IN_PREFECTURE = i
    i = i + 1: COL_IN_CITY = i
    i = i + 1: COL_IN_TOWN = i
    i = i + 1: COL_IN_BUILDING = i
    i = i + 1: COL_IN_ADDRESSLIST_PROHIBITED = i
    
    Dim j As Long: j = COL_OUT_ITEM_START
    j = j + 1: COL_OUT_NAME = j
    j = j + 1: COL_OUT_ZIP = j
    j = j + 1: COL_OUT_ADDRESS = j
    
End Function


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
    wsOut.Cells(rowNoOut, COL_OUT_NAME).Value = wsIn.Cells(rowNoIn, COL_IN_FAMILY_NAME) _
                                              & " " & wsIn.Cells(rowNoIn, COL_IN_LAST_NAME)
    
    ' �X�֔ԍ��F��4�����Ȃ��ꍇ�ɍl��
    Dim zip1 As String: zip1 = wsIn.Cells(rowNoIn, COL_IN_ZIP1).Value
    Dim zip2 As String: zip2 = wsIn.Cells(rowNoIn, COL_IN_ZIP2).Value
    wsOut.Cells(rowNoOut, COL_OUT_ZIP).Value = IIf(zip2 = "", zip1, zip1 & "-" & zip2)
        
    ' �Z��
    Dim prefecture As String: prefecture = wsIn.Cells(rowNoIn, COL_IN_PREFECTURE).Value
    Dim city As String: city = wsIn.Cells(rowNoIn, COL_IN_CITY).Value
    Dim town As String: town = wsIn.Cells(rowNoIn, COL_IN_TOWN).Value
    Dim building As String: building = wsIn.Cells(rowNoIn, COL_IN_BUILDING).Value
    Dim address As String
    address = prefecture & city & town
    If building <> "" Then
        address = address & " " & building
    End If
    wsOut.Cells(rowNoOut, COL_OUT_ADDRESS).Value = address
    
End Function


