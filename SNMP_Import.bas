' ============================================================
' SNMP Data Import & Chart Generator
' ============================================================
' 使い方:
'   1. Alt+F11 でVBEを開く
'   2. 挿入 > 標準モジュール でこのコードを貼り付け
'   3. Alt+F8 で「ImportSNMPData」を実行
'   4. テキストファイルが入ったフォルダを選択
' ============================================================

Option Explicit

' --- 定数 ---
Private Const DATA_SHEET_PREFIX  As String = "Data_"
Private Const CHART_SHEET_PREFIX As String = "Chart_"
Private Const MAX_SHEET_NAME     As Integer = 31   ' Excelシート名上限

' ============================================================
' メインエントリポイント
' ============================================================
Public Sub ImportSNMPData()
    Dim folderPath As String
    Dim txtFiles   As Collection
    Dim fileName   As Variant

    ' フォルダ選択ダイアログ
    folderPath = SelectFolder()
    If folderPath = "" Then
        MsgBox "フォルダが選択されませんでした。処理を中止します。", vbExclamation
        Exit Sub
    End If

    ' *.txt ファイル一覧取得
    Set txtFiles = GetTxtFiles(folderPath)
    If txtFiles.Count = 0 Then
        MsgBox "指定フォルダに .txt ファイルが見つかりませんでした。", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    On Error GoTo ErrHandler

    ' 各ファイルを処理
    For Each fileName In txtFiles
        Dim dataSheet  As Worksheet
        Dim chartSheet As Worksheet
        Dim baseName   As String

        baseName = GetBaseName(CStr(fileName))   ' 拡張子なしファイル名

        ' --- データシート作成 ---
        Set dataSheet = CreateOrClearSheet(DATA_SHEET_PREFIX & Left(baseName, MAX_SHEET_NAME - Len(DATA_SHEET_PREFIX)))
        ImportTabDelimitedFile folderPath & "\" & CStr(fileName), dataSheet

        ' --- グラフシート作成 ---
        Set chartSheet = CreateOrClearSheet(CHART_SHEET_PREFIX & Left(baseName, MAX_SHEET_NAME - Len(CHART_SHEET_PREFIX)))
        CreateChart dataSheet, chartSheet, baseName
    Next fileName

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

    MsgBox "インポート完了！" & vbCrLf & txtFiles.Count & " ファイルを処理しました。", vbInformation
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    MsgBox "エラーが発生しました。" & vbCrLf & Err.Number & ": " & Err.Description, vbCritical
End Sub

' ============================================================
' フォルダ選択ダイアログ
' ============================================================
Private Function SelectFolder() As String
    Dim shell  As Object
    Dim folder As Object

    Set shell = CreateObject("Shell.Application")
    Set folder = shell.BrowseForFolder(0, "テキストファイルが入ったフォルダを選択してください", 0)

    If Not folder Is Nothing Then
        SelectFolder = folder.Self.Path
    Else
        SelectFolder = ""
    End If
End Function

' ============================================================
' 指定フォルダの *.txt ファイル一覧を返す
' ============================================================
Private Function GetTxtFiles(folderPath As String) As Collection
    Dim col      As New Collection
    Dim fileName As String

    fileName = Dir(folderPath & "\*.txt")
    Do While fileName <> ""
        col.Add fileName
        fileName = Dir()
    Loop

    Set GetTxtFiles = col
End Function

' ============================================================
' 拡張子なしファイル名を返す
' ============================================================
Private Function GetBaseName(fileName As String) As String
    Dim dotPos As Integer
    dotPos = InStrRev(fileName, ".")
    If dotPos > 1 Then
        GetBaseName = Left(fileName, dotPos - 1)
    Else
        GetBaseName = fileName
    End If
End Function

' ============================================================
' シートを新規作成 or 既存シートをクリア
' ============================================================
Private Function CreateOrClearSheet(sheetName As String) As Worksheet
    Dim ws As Worksheet

    ' 既存シートを検索
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        ' 新規作成（末尾に追加）
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = sheetName
    Else
        ' 既存シートをクリア
        ws.Cells.Clear
        ' グラフオブジェクトも削除
        Dim obj As OLEObject
        For Each obj In ws.OLEObjects
            obj.Delete
        Next obj
        Dim cht As ChartObject
        For Each cht In ws.ChartObjects
            cht.Delete
        Next cht
    End If

    Set CreateOrClearSheet = ws
End Function

' ============================================================
' タブ区切りテキストファイルをシートに読み込む
' ============================================================
Private Sub ImportTabDelimitedFile(filePath As String, ws As Worksheet)
    Dim fileNum  As Integer
    Dim lineText As String
    Dim fields() As String
    Dim rowNum   As Long
    Dim colNum   As Integer

    fileNum = FreeFile()
    Open filePath For Input As #fileNum

    rowNum = 1
    Do While Not EOF(fileNum)
        Line Input #fileNum, lineText
        If Len(Trim(lineText)) > 0 Then
            fields = Split(lineText, vbTab)
            For colNum = 0 To UBound(fields)
                Dim cellVal As String
                cellVal = Trim(fields(colNum))

                ' 数値変換を試みる
                If IsNumeric(cellVal) Then
                    ws.Cells(rowNum, colNum + 1).Value = CDbl(cellVal)
                Else
                    ws.Cells(rowNum, colNum + 1).Value = cellVal
                End If
            Next colNum
            rowNum = rowNum + 1
        End If
    Loop

    Close #fileNum

    ' --- 書式設定 ---
    FormatDataSheet ws, rowNum - 1
End Sub

' ============================================================
' データシートの書式を整える
' ============================================================
Private Sub FormatDataSheet(ws As Worksheet, lastRow As Long)
    If lastRow < 1 Then Exit Sub

    Dim lastCol As Integer
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    ' ヘッダー行
    With ws.Rows(1)
        .Font.Bold = True
        .Interior.Color = RGB(31, 73, 125)   ' 濃い青
        .Font.Color = RGB(255, 255, 255)      ' 白文字
    End With

    ' データ範囲に罫線
    With ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.Color = RGB(180, 180, 180)
    End With

    ' 1列目（タイムスタンプ）の書式
    ws.Columns(1).NumberFormat = "yyyy/mm/dd hh:mm:ss"
    ws.Columns(1).AutoFit

    ' 数値列の書式
    Dim c As Integer
    For c = 2 To lastCol
        ws.Columns(c).NumberFormat = "#,##0.00"
        ws.Columns(c).AutoFit
    Next c

    ' ゼブラストライプ（偶数行）
    Dim r As Long
    For r = 2 To lastRow
        If r Mod 2 = 0 Then
            ws.Rows(r).Interior.Color = RGB(235, 241, 250)
        End If
    Next r

    ' ウィンドウ枠の固定（ヘッダー行）
    ws.Activate
    ws.Range("A2").Select
    ActiveWindow.FreezePanes = True
End Sub

' ============================================================
' グラフシートにグラフを作成する
' ============================================================
Private Sub CreateChart(dataWs As Worksheet, chartWs As Worksheet, baseName As String)
    Dim lastRow  As Long
    Dim lastCol  As Integer
    Dim chartObj As ChartObject
    Dim cht      As Chart
    Dim ser      As Series

    lastRow = dataWs.Cells(dataWs.Rows.Count, 1).End(xlUp).Row
    lastCol = dataWs.Cells(1, dataWs.Columns.Count).End(xlToLeft).Column

    If lastRow < 2 Or lastCol < 2 Then Exit Sub

    ' --- グラフ種別をファイル名で切り替え ---
    Dim chartTitle As String
    Dim chartType  As XlChartType

    Select Case LCase(baseName)
        Case "cpu"
            chartTitle = "CPU 使用率"
            chartType  = xlLine
        Case "memory"
            chartTitle = "メモリ使用量"
            chartType  = xlLine
        Case "snmptrap"
            chartTitle = "SNMP Trap 発生件数"
            chartType  = xlColumnClustered
        Case Else
            chartTitle = baseName & " データ"
            chartType  = xlLine
    End Select

    ' --- ChartObject を chartWs に配置 ---
    Set chartObj = chartWs.ChartObjects.Add( _
        Left:=20, Top:=20, Width:=900, Height:=400)
    Set cht = chartObj.Chart

    cht.ChartType = chartType

    ' --- データ系列を追加（2列目以降を数値列として追加）---
    cht.SeriesCollection.NewSeries   ' ダミーを消す
    Do While cht.SeriesCollection.Count > 0
        cht.SeriesCollection(1).Delete
    Loop

    Dim colIdx As Integer
    For colIdx = 2 To lastCol
        Set ser = cht.SeriesCollection.NewSeries()

        ' X軸（タイムスタンプ列）
        ser.XValues = dataWs.Range( _
            dataWs.Cells(2, 1), dataWs.Cells(lastRow, 1))

        ' Y値（各数値列）
        ser.Values = dataWs.Range( _
            dataWs.Cells(2, colIdx), dataWs.Cells(lastRow, colIdx))

        ' 系列名（ヘッダー）
        ser.Name = dataWs.Cells(1, colIdx).Value
    Next colIdx

    ' --- グラフ装飾 ---
    With cht
        .HasTitle = True
        .ChartTitle.Text = chartTitle
        .ChartTitle.Font.Size = 14
        .ChartTitle.Font.Bold = True

        ' X軸
        With .Axes(xlCategory)
            .HasTitle = True
            .AxisTitle.Text = "時刻"
            .TickLabels.NumberFormat = "hh:mm"
            .TickLabels.Orientation = 45
        End With

        ' Y軸
        With .Axes(xlValue)
            .HasTitle = True
            .AxisTitle.Text = GetYAxisLabel(baseName)
        End With

        .HasLegend = True
        .Legend.Position = xlLegendPositionBottom

        ' プロットエリア背景
        .PlotArea.Interior.Color = RGB(245, 245, 245)

        ' 折れ線グラフはスムージング & マーカー
        If chartType = xlLine Then
            Dim s As Series
            For Each s In .SeriesCollection
                s.Smooth = True
                s.MarkerStyle = xlMarkerStyleCircle
                s.MarkerSize = 5
            Next s
        End If
    End With

    ' --- snmptrap 専用：重大度ごとに色分け ---
    If LCase(baseName) = "snmptrap" Then
        AddTrapSummaryTable dataWs, chartWs, lastRow
    End If
End Sub

' ============================================================
' Y軸ラベルをファイル名から返す
' ============================================================
Private Function GetYAxisLabel(baseName As String) As String
    Select Case LCase(baseName)
        Case "cpu"    : GetYAxisLabel = "使用率 (%)"
        Case "memory" : GetYAxisLabel = "使用量 (MB)"
        Case Else     : GetYAxisLabel = "値"
    End Select
End Function

' ============================================================
' snmptrap 用：重大度別集計テーブルをグラフシートに追加
' ============================================================
Private Sub AddTrapSummaryTable(dataWs As Worksheet, chartWs As Worksheet, lastRow As Long)
    ' Severity 列を探す
    Dim sevCol As Integer
    sevCol = 0
    Dim c As Integer
    For c = 1 To dataWs.Cells(1, dataWs.Columns.Count).End(xlToLeft).Column
        If InStr(LCase(dataWs.Cells(1, c).Value), "severity") > 0 Then
            sevCol = c
            Exit For
        End If
    Next c
    If sevCol = 0 Then Exit Sub

    ' 集計
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    Dim r As Long
    For r = 2 To lastRow
        Dim key As String
        key = Trim(CStr(dataWs.Cells(r, sevCol).Value))
        If key <> "" Then
            If dict.Exists(key) Then
                dict(key) = dict(key) + 1
            Else
                dict.Add key, 1
            End If
        End If
    Next r

    ' テーブル出力（グラフの下）
    Dim startRow As Long
    startRow = 26

    chartWs.Cells(startRow, 1).Value = "重大度"
    chartWs.Cells(startRow, 2).Value = "件数"
    With chartWs.Range(chartWs.Cells(startRow, 1), chartWs.Cells(startRow, 2))
        .Font.Bold = True
        .Interior.Color = RGB(31, 73, 125)
        .Font.Color = RGB(255, 255, 255)
    End With

    Dim rowOffset As Long
    rowOffset = 1
    Dim k As Variant
    For Each k In dict.Keys
        chartWs.Cells(startRow + rowOffset, 1).Value = k
        chartWs.Cells(startRow + rowOffset, 2).Value = dict(k)
        ' 重大度に応じた色
        Select Case LCase(CStr(k))
            Case "critical" : chartWs.Cells(startRow + rowOffset, 1).Interior.Color = RGB(255, 199, 199)
            Case "warning"  : chartWs.Cells(startRow + rowOffset, 1).Interior.Color = RGB(255, 242, 199)
            Case "info"     : chartWs.Cells(startRow + rowOffset, 1).Interior.Color = RGB(199, 235, 255)
        End Select
        rowOffset = rowOffset + 1
    Next k

    chartWs.Columns(1).AutoFit
    chartWs.Columns(2).AutoFit
End Sub
