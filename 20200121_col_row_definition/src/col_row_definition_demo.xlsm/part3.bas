Attribute VB_Name = "part3"
Option Explicit

' --- シート定義
Const INPUT_SHEET_NAME = "INPUT"
Const OUTPUT_SHEET_NAME = "OUTPUT"

' --- 入力列定義
Enum COL_IN
    Index = 1
    familyName ' 2 (値を指定しなければ、前の項目からインクリメントされた値になる)
    lastName ' 3
    sex ' 4
    zip1 ' (以下略)
    zip2
    prefecture
    city
    town
    building
    addresslistProhibited
End Enum

' --- 入力行定義
Const ROW_IN_DATA_START = 2

' --- 出力列定義
Enum COL_OUT
    Name = 1
    zip
    address
End Enum

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
    Do Until wsIn.Cells(rowNoIn, COL_IN.Index).Value = ""
        
        ' 出力禁止でなければ出力する
        If wsIn.Cells(rowNoIn, COL_IN.addresslistProhibited).Value <> "Y" Then
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
    wsOut.Cells(rowNoOut, COL_OUT.Name).Value = wsIn.Cells(rowNoIn, COL_IN.familyName) _
                                              & " " & wsIn.Cells(rowNoIn, COL_IN.lastName)
    
    ' 郵便番号：下4桁がない場合に考慮
    Dim zip1 As String: zip1 = wsIn.Cells(rowNoIn, COL_IN.zip1).Value
    Dim zip2 As String: zip2 = wsIn.Cells(rowNoIn, COL_IN.zip2).Value
    wsOut.Cells(rowNoOut, COL_OUT.zip).Value = IIf(zip2 = "", zip1, zip1 & "-" & zip2)
        
    ' 住所
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




