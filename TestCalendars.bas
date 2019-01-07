Public Sub TestMonthCalendar()
    Dim rng As Range, StartDate As Date
    Set rng = ActiveSheet.Range("A1")
    StartDate = CDate("01/01/2017")
    Call CreateMonthCalendar(StartDate, rng)
End Sub
Public Sub TestCalendar()
    Dim rng As Range, rngCalendar As Range, StartDate As Date
    Set rng = ActiveSheet.Range("A1")
    StartDate = CDate("01/01/2019")
    Call CreateCalendar(StartDate, rng, 12, 3)
End Sub

