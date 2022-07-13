Attribute VB_Name = "Analysis_incidents"
Sub Analyse()
Dim ex As Excel.Application
Dim bk As Excel.Workbook
Dim sht As Excel.Worksheet
Dim ppt As PowerPoint.Application
Dim pf As PowerPoint.Presentation
Dim sld As PowerPoint.Slide

Dim sum_row, i, ix, y, marker_bot, marker_dif, marker_head As Integer
Dim col_array, col_desc_array, col As Variant
col_array = Array(4, 5, 9, 10, 11, 12, 13)
col_desc_array = Array(6, 7, 8)

On Error Resume Next

Set ex = GetObject(, "Excel.Application")

If Err <> 0 Then
   Set ex = New Excel.Application
End If

Set bk = ex.ActiveWorkbook
If MsgBox("Работаем с файлом " & bk.Name, vbOKCancel, "Выбор файла") <> vbOK Then Exit Sub
Set sht = bk.ActiveSheet
sum_row = sht.Cells(Rows.Count, 1).End(xlUp).Row

If ppt Is Nothing Then
    Set ppt = New PowerPoint.Application
    Set pf = ppt.Presentations.Add
    End If
Set sld = pf.Slides.Add(1, ppLayoutBlank)
With sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 30, 300, 900, 20).TextFrame.TextRange
    .Text = "Обзор постмортемов"
    .Font.Size = 32
    .Font.Bold = True
End With
Set pptLayout = ActivePresentation.Slides(1).CustomLayout


For i = 1 To sum_row
    ix = i + 1
    Set sld = pf.Slides.Add(ix, ppLayoutBlank)
    marker_bot = 30
    ' название инцидента
    With sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 30, 30, 900, 20).TextFrame.TextRange
        .Text = sht.Cells(ix, 1) & " - " & sht.Cells(ix, 3)
        .Font.Size = 24
        marker_bot = marker_bot + .BoundHeight + 10
        marker_head = marker_bot
    End With

    For col = 0 To UBound(col_array)
        With sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 30, marker_bot, 100, 20).TextFrame.TextRange
            .Text = sht.Cells(1, col_array(col))
            .Font.Size = 12
            marker_dif = .BoundHeight
        End With
        With sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 140, marker_bot, 220, 20).TextFrame.TextRange
            .Text = sht.Cells(ix, col_array(col))
            .Font.Size = 12
            If marker_dif > .BoundHeight Then
                marker_bot = marker_bot + marker_dif + 10
            Else
                marker_bot = marker_bot + .BoundHeight + 10
            End If
            marker_dif = 0
        End With
    Next
    marker_bot = marker_head
    For col = 0 To UBound(col_desc_array)
        With sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 390, marker_bot, 500, 20).TextFrame.TextRange
            .Text = sht.Cells(1, col_desc_array(col))
            .Font.Size = 12
            .Font.Bold = True
            marker_bot = marker_bot + .BoundHeight + 10
        End With

        With sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 390, marker_bot, 500, 20).TextFrame.TextRange
            .Text = sht.Cells(ix, col_desc_array(col))
            .Font.Size = 12
            marker_bot = marker_bot + .BoundHeight + 10
        End With
    Next

Next

ppt.Visible = True

End Sub
