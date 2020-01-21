Attribute VB_Name = "part1"
Option Explicit

' --- シート定義
Const INPUT_SHEET_NAME = "INPUT"
Const OUTPUT_SHEET_NAME = "OUTPUT"

' --- 入力列定義
Const COL_IN_INDEX = 1
Const COL_IN_FAMILY_NAME = 2
Const COL_IN_LAST_NAME = 3
Const COL_IN_SEX = 4
Const COL_IN_ZIP1 = 5
Const COL_IN_ZIP2 = 6
Const COL_IN_PREFECTURE = 7
Const COL_IN_CITY = 8
Const COL_IN_TOWN = 9
Const COL_IN_BUILDING = 10
Const COL_IN_ADDRESSLIST_FORHIBITED = 11

' --- 入力行定義
Const ROW_IN_DATA_START = 2

' --- 出力列定義
Const COL_OUT_NAME = 1
Const COL_OUT_ZIP = 2
Const COL_OUT_ADDRESS = 3

' --- 出力行定義
Const ROW_OUT_DATA_START = 3


' 住所テーブルから、住所録っぽいものをつくる。
' 対象のブックをアクティブにした状態で起動する。
Sub tableToAddlessList()
    
    ' ワークブック・ワークシートの定義
    Dim wb As Workbook: Set wb = ActiveWorkbook
    Dim wsIn As Worksheet: Set wsIn = wb.Worksheets(INPUT_SHEET_NAME)
    Dim wsOut As Worksheet: Set wsOut = createOutputSheet(wb)
    
    ' 入力開始行・出力開始行を設定
    Dim rowNoOut As Long: rowNoOut = ROW_OUT_DATA_START
    Dim rowNoIn As Long: rowNoIn = ROW_IN_DATA_START
    
    ' 入力のNo.列が空になるまで、1行ごとに処理する
    Do Until wsIn.Cells(rowNoIn, COL_IN_INDEX).Value = ""
        
        ' 出力禁止でなければ出力する
        If wsIn.Cells(rowNoIn, COL_IN_ADDRESSLIST_FORHIBITED).Value <> "Y" Then
            Call writeAddress(wsIn, wsOut, rowNoIn, rowNoOut)
            rowNoOut = rowNoOut + 1
        End If
        
        rowNoIn = rowNoIn + 1
    Loop
    
MsgBox "done"

End Sub


' 出力シートを作成する。既に存在する場合は、削除してから作成しなおす。
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

' アドレス1件を書き込む
Function writeAddress(wsIn As Worksheet, wsOut As Worksheet, rowNoIn As Long, rowNoOut As Long)
        
    ' 名前：苗字＋空白＋名前
    wsOut.Cells(rowNoOut, COL_OUT_NAME).Value = wsIn.Cells(rowNoIn, COL_IN_FAMILY_NAME) _
                                              & " " & wsIn.Cells(rowNoIn, COL_IN_LAST_NAME)
    
    ' 郵便番号：下4桁がない場合に考慮
    Dim zip1 As String: zip1 = wsIn.Cells(rowNoIn, COL_IN_ZIP1).Value
    Dim zip2 As String: zip2 = wsIn.Cells(rowNoIn, COL_IN_ZIP2).Value
    wsOut.Cells(rowNoOut, COL_OUT_ZIP).Value = IIf(zip2 = "", zip1, zip1 & "-" & zip2)
        
    ' 住所
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
