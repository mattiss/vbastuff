Private Const CALENDAR_COLUMN_WIDTH = 6.14
Private Const CALENDAR_ROW_HEIGTH = 36
Public Sub CreateCalendar(dt As Date, rng As Range, Optional NumberOfMonths As Integer = 12, Optional NumberOfColumns As Integer = 3)
    For i = 0 To NumberOfMonths - 1
        r = i \ NumberOfColumns
        c = i Mod NumberOfColumns
        Call CreateMonthCalendar(DateAdd("m", i, dt), rng.Offset(r * 10, c * 10))
    Next i
End Sub
Public Sub CreateMonthCalendar(dt As Date, rng As Range)
    Application.ScreenUpdating = False
    Dim rngTemp As Range
    Set rngTemp = CreateTitle(dt, rng)
    Set rngTemp = CreateGridHeader(dt, rngTemp)
    Set rngTemp = CreateGrid(dt, rngTemp)
    Application.ScreenUpdating = True
End Sub
Private Function CreateTitle(dt As Date, rng As Range) As Range
    With rng.Offset(0, 1)
        .NumberFormat = "@"
        .FormulaR1C1 = CStr(Format(dt, "mmm YYYY"))
        .Font.Size = 20
        .Font.Bold = True
    End With
    Set CreateTitle = rng.Offset(1, 0)
End Function
Private Function CreateGridHeader(dt As Date, rng As Range) As Range
    For i = 1 To 7
        With rng.Offset(0, i)
            .FormulaR1C1 = Left(WeekdayName(i, True, vbMonday), 2)
            .Font.Size = 12
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
        End With
    Next i
    Set CreateGridHeader = rng.Offset(1, 0)
End Function
Private Function CreateGrid(dt As Date, rng As Range) As Range
    wkDay = Weekday(dt, vbMonday) - 1
    dtTemp = dt
    i = 0
    Dim rngTemp As Range
    While (Month(dtTemp) = Month(dt))
        r = (wkDay + i) \ 7
        c = (wkDay + i) Mod 7
        Set rngTemp = rng.Offset(r, c + 1)
        rngTemp.Value = dtTemp
        DrawBorders rngTemp
        dtTemp = dtTemp + 1
        i = i + 1
    Wend
    Set CreateGrid = rng.Offset(6, 0)
End Function
Private Sub DrawBorders(rng As Range)
    With rng
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlMedium
        .Font.ThemeColor = xlThemeColorDark1
        .ColumnWidth = CALENDAR_COLUMN_WIDTH
        .RowHeight = CALENDAR_ROW_HEIGTH
    End With
End Sub
