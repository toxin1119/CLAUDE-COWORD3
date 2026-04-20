Sub 建立出貨單模板()
' ============================================
' 建立出貨單空白模板
' 使用方式：在 Excel 中按 Alt+F11 開啟 VBA 編輯器
'           貼入後按 F5 執行
' 產出檔案：D:\課程規劃\115Q2-3\data\出貨單.xlsx
' ============================================

    Dim ws As Worksheet
    Dim wb As Workbook
    Dim savePath As String
    Dim r As Long, c As Long, i As Long, edge As Long, infoRow As Long
    
    savePath = "D:\課程規劃\115Q2-3\data\出貨單.xlsx"
    
    Application.ScreenUpdating = False
    
    Set wb = Workbooks.Add
    Set ws = wb.Worksheets(1)
    ws.Name = "出貨單"
    
    ' ============================================
    ' 頁面設定
    ' ============================================
    With ws.PageSetup
        .PaperSize = xlPaperA4
        .Orientation = xlPortrait
        .TopMargin = Application.CentimetersToPoints(1.5)
        .BottomMargin = Application.CentimetersToPoints(1.5)
        .LeftMargin = Application.CentimetersToPoints(1.5)
        .RightMargin = Application.CentimetersToPoints(1.5)
    End With
    
    ' ============================================
    ' 欄寬設定
    ' ============================================
    ws.Columns("A").ColumnWidth = 6
    ws.Columns("B").ColumnWidth = 22
    ws.Columns("C").ColumnWidth = 12
    ws.Columns("D").ColumnWidth = 8
    ws.Columns("E").ColumnWidth = 6
    ws.Columns("F").ColumnWidth = 14
    ws.Columns("G").ColumnWidth = 14
    ws.Columns("H").ColumnWidth = 18
    
    ' 全表預設字體
    ws.Cells.Font.Name = "微軟正黑體"
    ws.Cells.Font.Size = 10
    
    ' ============================================
    ' 第 1-2 列：公司名稱 & 文件標題
    ' ============================================
    ws.Range("A1:H1").Merge
    With ws.Range("A1")
        .Value = "○○○股份有限公司"
        .Font.Size = 18
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    ws.Rows(1).RowHeight = 30
    
    ws.Range("A2:H2").Merge
    With ws.Range("A2")
        .Value = "出  貨  單"
        .Font.Size = 16
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Font.Color = RGB(0, 126, 57)
    End With
    ws.Rows(2).RowHeight = 28
    
    ' ============================================
    ' 第 3 列：空白分隔
    ' ============================================
    ws.Rows(3).RowHeight = 6
    
    ' ============================================
    ' 第 4-9 列：客戶資訊區
    ' ============================================
    Dim infoLabels As Variant
    infoLabels = Array("客戶名稱：", "聯 絡 人：", "送貨地址：", "客戶訂單編號：", "出貨日期：", "交貨條件：")
    
    For i = 0 To 5
        infoRow = 4 + i
        ws.Range("A" & infoRow & ":B" & infoRow).Merge
        With ws.Range("A" & infoRow)
            .Value = infoLabels(i)
            .Font.Bold = True
            .Font.Size = 10
        End With
        ws.Range("C" & infoRow & ":H" & infoRow).Merge
        ws.Range("C" & infoRow).Interior.Color = RGB(255, 255, 230)
        ws.Rows(infoRow).RowHeight = 22
        With ws.Range("A" & infoRow & ":H" & infoRow).Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Color = RGB(200, 200, 200)
        End With
    Next i
    
    ' ============================================
    ' 第 10 列：空白分隔
    ' ============================================
    ws.Rows(10).RowHeight = 8
    
    ' ============================================
    ' 第 11 列：明細表頭
    ' ============================================
    Dim headers As Variant
    headers = Array("項次", "品名", "型號", "數量", "單位", "單價(NTD)", "小計(NTD)", "備註")
    
    For i = 0 To 7
        With ws.Cells(11, i + 1)
            .Value = headers(i)
            .Font.Bold = True
            .Font.Size = 10
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(0, 126, 57)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
    Next i
    ws.Rows(11).RowHeight = 26
    
    ' 表頭框線
    For c = 1 To 8
        For edge = xlEdgeLeft To xlEdgeBottom
            With ws.Cells(11, c).Borders(edge)
                .LineStyle = xlContinuous
                .Color = RGB(0, 100, 50)
                .Weight = xlMedium
            End With
        Next edge
    Next c
    
    ' ============================================
    ' 第 12-21 列：10 行空白明細（含公式）
    ' ============================================
    For r = 12 To 21
        ws.Cells(r, 1).HorizontalAlignment = xlCenter
        ws.Cells(r, 4).HorizontalAlignment = xlCenter
        ws.Cells(r, 4).NumberFormat = "#,##0"
        ws.Cells(r, 5).HorizontalAlignment = xlCenter
        ws.Cells(r, 6).NumberFormat = "#,##0"
        ws.Cells(r, 6).HorizontalAlignment = xlRight
        ws.Cells(r, 7).Formula = "=IF(D" & r & "*F" & r & "=0,"""",D" & r & "*F" & r & ")"
        ws.Cells(r, 7).NumberFormat = "#,##0"
        ws.Cells(r, 7).HorizontalAlignment = xlRight
        
        ' 偶數列淺綠底色
        If r Mod 2 = 0 Then
            ws.Range("A" & r & ":H" & r).Interior.Color = RGB(245, 252, 245)
        End If
        
        ws.Rows(r).RowHeight = 22
        
        ' 明細區框線
        For c = 1 To 8
            For edge = xlEdgeLeft To xlEdgeBottom
                With ws.Cells(r, c).Borders(edge)
                    .LineStyle = xlContinuous
                    .Color = RGB(180, 180, 180)
                    .Weight = xlThin
                End With
            Next edge
        Next c
    Next r
    
    ' ============================================
    ' 第 22 列：空白分隔
    ' ============================================
    ws.Rows(22).RowHeight = 4
    
    ' ============================================
    ' 第 23-25 列：合計區
    ' ============================================
    ws.Range("A23:E23").Merge
    ws.Range("F23").Value = "商品合計"
    ws.Range("F23").Font.Bold = True
    ws.Range("F23").HorizontalAlignment = xlRight
    ws.Range("G23").Formula = "=SUM(G12:G21)"
    ws.Range("G23").NumberFormat = "#,##0"
    ws.Range("G23").Font.Bold = True
    ws.Range("G23").HorizontalAlignment = xlRight
    ws.Rows(23).RowHeight = 24
    
    ws.Range("F24").Value = "稅額(5%)"
    ws.Range("F24").Font.Bold = True
    ws.Range("F24").HorizontalAlignment = xlRight
    ws.Range("G24").Formula = "=G23*0.05"
    ws.Range("G24").NumberFormat = "#,##0"
    ws.Range("G24").Font.Bold = True
    ws.Range("G24").HorizontalAlignment = xlRight
    ws.Rows(24).RowHeight = 24
    
    ws.Range("F25").Value = "總　　計"
    ws.Range("F25").Font.Bold = True
    ws.Range("F25").Font.Size = 12
    ws.Range("F25").HorizontalAlignment = xlRight
    With ws.Range("G25")
        .Formula = "=G23+G24"
        .NumberFormat = "#,##0"
        .Font.Bold = True
        .Font.Size = 12
        .Font.Color = RGB(220, 38, 38)
        .HorizontalAlignment = xlRight
    End With
    ws.Rows(25).RowHeight = 28
    
    ' 合計區框線
    For r = 23 To 25
        For c = 6 To 7
            For edge = xlEdgeLeft To xlEdgeBottom
                With ws.Cells(r, c).Borders(edge)
                    .LineStyle = xlContinuous
                    .Color = RGB(100, 100, 100)
                    .Weight = xlMedium
                End With
            Next edge
        Next c
    Next r
    
    ' 總計列底色
    ws.Range("F25:G25").Interior.Color = RGB(254, 249, 195)
    
    ' ============================================
    ' 第 27-29 列：備註區
    ' ============================================
    ws.Rows(26).RowHeight = 8
    
    ws.Range("A27").Value = "備　註："
    ws.Range("A27").Font.Bold = True
    
    For r = 27 To 29
        ws.Range("B" & r & ":H" & r).Merge
        ws.Range("B" & r).Interior.Color = RGB(255, 255, 230)
        ws.Rows(r).RowHeight = 22
        With ws.Range("A" & r & ":H" & r).Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Color = RGB(200, 200, 200)
        End With
    Next r
    
    ' ============================================
    ' 第 31-32 列：簽核區
    ' ============================================
    ws.Rows(30).RowHeight = 15
    
    ws.Range("A31").Value = "製表人："
    ws.Range("A31").Font.Bold = True
    ws.Range("B31:C31").Merge
    ws.Range("B31").Borders(xlEdgeBottom).LineStyle = xlContinuous
    
    ws.Range("E31").Value = "審核主管："
    ws.Range("E31").Font.Bold = True
    ws.Range("F31:G31").Merge
    ws.Range("F31").Borders(xlEdgeBottom).LineStyle = xlContinuous
    ws.Rows(31).RowHeight = 28
    
    ws.Range("A32:H32").Merge
    With ws.Range("A32")
        .Value = "※ 本出貨單經主管簽核後生效，請妥善保存。"
        .Font.Size = 8
        .Font.Color = RGB(150, 150, 150)
        .Font.Italic = True
        .HorizontalAlignment = xlCenter
    End With
    
    ' ============================================
    ' 列印範圍 & 儲存
    ' ============================================
    ws.PageSetup.PrintArea = "A1:H32"
    
    Application.DisplayAlerts = False
    wb.SaveAs savePath, xlOpenXMLWorkbook
    Application.DisplayAlerts = True
    
    Application.ScreenUpdating = True
    
    MsgBox "✅ 出貨單模板已建立完成！" & vbCrLf & vbCrLf & "儲存位置：" & savePath, vbInformation, "完成"

End Sub
