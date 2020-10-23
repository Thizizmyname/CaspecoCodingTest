 '------------------------------------------------------------------------------------
 '
 '      1      Ge några förslag på hur man skulle kunna förbättra denna kod. Rangordna dem i viktighetsordning.
 '
 '------------------------------------------------------------------------------------
 

Public Function SumPeriods(ByVal fromDate As Date, ByVal toDate As Date, ByVal summering As String, ByVal const_ As String, ByVal type As Integer)
    Dim level As Integer
    Dim periods As Integer
    Dim startRow As Integer

    Dim progress As frmProgress
    If type = 1 Then
        startRow = 13
    Else
        startRow = rowForReport + 5
    End If

    ws.ResetRowsFrom(startRow)
    tmpRapport.snittnotaSum = 0
    tmpRapport.blgGuestSum = 0
    tmpRapport.blgSeatSum = 0
    tmpRapport.guestsSum = 0
    tmpRapport.löneSaleSum = 0
    tmpRapport.löneKostSum = 0
    tmpRapport.arbtimSalesSum = 0
    tmpRapport.arbtimSum = 0
    tmpRapport.arbtimSumLP = 0
    tmpRapport.dateCol.Clear()
    tmpRapport.snittnotaCol1.Clear()
    tmpRapport.snitttnotaCol2.Clear()
    tmpRapport.blgCol.Clear()
    tmpRapport.blgColGuests.Clear()
    tmpRapport.blgColSeats.Clear()
    tmpRapport.löneprocCol1.Clear()
    tmpRapport.löneprocCol2.Clear()
    tmpRapport.löneprocCol3.Clear()
    tmpRapport.arbtimCol1.Clear()
    tmpRapport.arbtimCol2.Clear()
    tmpRapport.arbtimColLP.Clear()
    tmpRapport.weatherCol.Clear()
    tmpRapport.snittnotaCol1PrevYear.Clear()
    tmpRapport.snitttnotaCol2PrevYear.Clear()
    tmpRapport.blgColPrevYear.Clear()
    tmpRapport.blgColGuestsPrevYear.Clear()
    tmpRapport.blgColSeatsPrevYear.Clear()
    tmpRapport.löneprocCol1PrevYear.Clear()
    tmpRapport.löneprocCol2PrevYear.Clear()
    tmpRapport.löneprocCol3PrevYear.Clear()
    tmpRapport.arbtimCol1PrevYear.Clear()
    tmpRapport.arbtimCol2PrevYear.Clear()
    tmpRapport.arbtimColLPPrevYear.Clear()
    tmpRapport.datesForGraph.Clear()

    Dim rowsBetweenReports As Integer = 0
    If tmpRapport.target > 0 Then
        rowsBetweenReports += 2
    End If
    If CBSShowPrevYear.Checked Then
        rowsBetweenReports += 2
    End If


    'Dag
    If summering = "Dag" Then
        ws.Range("b1").Offset(startRow - 1, 0).Value = tmpRapport.summering
        periods = DateDiff(DateInterval.Day, fromDate, toDate)
        level = 1 - 1
        Dim tmpPeriod As Integer = 0
        Dim tmpDay As Integer

        progress = New frmProgress("Beräknar...", periods)
        If periods > 10 Then
            progress.ShowProgress()
        End If

        If rbSelectDays.Checked Then
            For Each datum As Date In dateList
                If tmpRapport.weekDaysSelected.Contains(DatePart(DateInterval.Weekday, datum, FirstDayOfWeek.Monday,FirstWeekOfYear.FirstFourDays)) Then
                    addToLists(datum, datum, const_)
                    tmpPeriod += 1
                End If
            Next
        Else
            For day As Integer = 0 To periods
                If tmpRapport.weekDaysSelected.Contains(DatePart(DateInterval.Weekday, fromDate.AddDays(day), FirstDayOfWeek.Monday,FirstWeekOfYear.FirstFourDays)) Then
                    tmpDay += 1
                    Dim startDate As Date = fromDate.AddDays(day)
                    Dim endDate As Date = fromDate.AddDays(day + level)
                    tmpRapport.datesForGraph.Add(startDate.Day & "/" & startDate.Month & "-" & startDate.Year - 2000)
                    addToLists(startDate, endDate, const_)
                    tmpPeriod += 1
                End If
                progress.Tick()
            Next
        End If

        Dim extrarows As Integer = renderTable(tmpPeriod - 1, startRow, summering)
        progress.Close()
        

        rowForReport = startRow + tmpPeriod + rowsBetweenReports + extrarows
    End If

    'Vecka
    If summering = "Vecka" Then
        ws.Range("b1").Offset(startRow - 1, 0).Value = tmpRapport.summering
        level = 7 - 1
        Dim tmpPeriod As Integer = 0
        Dim currDay = DatePart(DateInterval.Weekday, fromDate, FirstDayOfWeek.Monday, FirstWeekOfYear.FirstFourDays)
        Dim weekadjust As Integer = 1 - currDay
        Dim firstDayOfW As Date = fromDate.AddDays(weekadjust)
        progress = New frmProgress("Beräknar...", periods)
        progress.ShowProgress()
        periods = DateDiff(DateInterval.WeekOfYear, fromDate, toDate, FirstDayOfWeek.Monday, FirstWeekOfYear.FirstFourDays)
        If rbSelectDays.Checked Then
            With ws.Range("b1").Offset(startRow, 0)
                .SetColumnSpan(11)
                .Value = "VECKOSUMMERING FUNGERAR INTE MED ENSTAKA DAGAR"
                .SetBGColor(Color.White)
                .SetFontColor(Color.Black)
                .SetFont(boldFont)
            End With
        Else
            For week As Integer = 0 To periods
                Dim startDate As Date = firstDayOfW.AddDays(week * 7)
                Dim endDate As Date = firstDayOfW.AddDays(week * 7 + level)

                With ws.Range("b1").Offset(startRow + week, 0)
                    .SetBorderLeft(2, frmUI.currentTheme.BorderColor, TableViewCellStyle.BorderStyle.Continuous)
                    .SetBorderRight(2, frmUI.currentTheme.BorderColor, TableViewCellStyle.BorderStyle.Continuous)
                    .SetBorderBottom(1, frmUI.currentTheme.BorderColor, TableViewCellStyle.BorderStyle.Continuous)
                    .SetFont(boldFont)
                    .Cell.tag = New Object() {startDate, endDate}
                    .Value = DatePart(DateInterval.WeekOfYear, startDate, FirstDayOfWeek.Monday, FirstWeekOfYear.FirstFourDays) & "-" & startDate.Year
                    .SetBGColor(frmUI.currentTheme.LineHeaderColor)
                    .SetFontColor(frmUI.currentTheme.LineHeaderTextColor)
                End With
                tmpRapport.datesForGraph.Add(DatePart(DateInterval.WeekOfYear, startDate, FirstDayOfWeek.Monday, FirstWeekOfYear.FirstFourDays) & "-" & startDate.Year)
                addToLists(startDate, endDate, const_)
                tmpPeriod += 1
                progress.Tick()

            Next
        End If
        renderTable(tmpPeriod - 1, startRow, summering)
        rowForReport = startRow + tmpPeriod + rowsBetweenReports
        progress.Close()
    End If

    'Månad
    If summering = "Månad" Then
        ws.Range("b1").Offset(startRow - 1, 0).Value = tmpRapport.summering
        Dim firstDayOfM As New Date(fromDate.Year, fromDate.Month, 1)
        periods = DateDiff(DateInterval.Month, fromDate, toDate, FirstDayOfWeek.Monday, FirstWeekOfYear.FirstFourDays)
        Dim tmpPeriod As Integer = 0
        progress = New frmProgress("Beräknar...", periods)
        progress.ShowProgress()
        If rbSelectDays.Checked Then
            With ws.Range("b1").Offset(startRow, 0)
                .SetColumnSpan(11)
                .Value = "MÅNADSSUMMERING FUNGERAR INTE MED ENSTAKA DAGAR"
                .SetBGColor(Color.White)
                .SetFontColor(Color.Black)
                .SetFont(boldFont)
            End With
        Else
            For month As Integer = 0 To periods
                Dim startDate As Date = firstDayOfM.AddMonths(month)
                Dim year As Integer = startDate.Year
                Dim tmpendDate As Date = startDate.AddMonths(1)
                Dim endDate As Date = tmpendDate.AddDays(-1)
                With ws.Range("b1").Offset(startRow + month, 0)
                    .SetBorderLeft(2, frmUI.currentTheme.BorderColor, TableViewCellStyle.BorderStyle.Continuous)
                    .SetBorderRight(2, frmUI.currentTheme.BorderColor, TableViewCellStyle.BorderStyle.Continuous)
                    .SetBorderBottom(1, frmUI.currentTheme.BorderColor, TableViewCellStyle.BorderStyle.Continuous)
                    .SetFont(boldFont)
                    .Cell.tag = New Object() {startDate, endDate}
                    .Value = MonthName(startDate.Month) & "-" & startDate.Year
                    .SetBGColor(frmUI.currentTheme.LineHeaderColor)
                    .SetFontColor(frmUI.currentTheme.LineHeaderTextColor)
                End With
                tmpRapport.datesForGraph.Add(MonthName(startDate.Month) & "-" & startDate.Year)
                addToLists(startDate, endDate, const_)
                tmpPeriod += 1
                progress.Tick()
            Next
        End If
        progress.Close()
        renderTable(tmpPeriod - 1, startRow, summering)
        rowForReport = startRow + tmpPeriod + rowsBetweenReports
    End If

    'år
    If summering = "År" Then
        ws.Range("b1").Offset(startRow - 1, 0).Value = tmpRapport.summering
        Dim firstYear As Integer = DatePart(DateInterval.Year, fromDate, FirstDayOfWeek.Monday, FirstWeekOfYear.FirstFourDays)
        Dim firstDayOfY As New Date(firstYear, 1, 1)
        periods = DateDiff(DateInterval.Year, fromDate, toDate, FirstDayOfWeek.Monday, FirstWeekOfYear.FirstFourDays)
        Dim tmpPeriod As Integer = 0
        progress = New frmProgress("Beräknar...", periods)
        progress.ShowProgress()
        If rbSelectDays.Checked Then
            With ws.Range("b1").Offset(startRow, 0)
                .SetColumnSpan(11)
                .Value = "ÅRSSUMMERING FUNGERAR INTE MED ENSTAKA DAGAR"
                .SetBGColor(Color.White)
                .SetFontColor(Color.Black)
                .SetFont(boldFont)
            End With
        Else
            For halfYear As Integer = 0 To periods
                Dim startDate As Date = firstDayOfY.AddMonths(halfYear * 12)
                Dim year As Integer = startDate.Year
                Dim tmpendDate As Date = firstDayOfY.AddMonths((halfYear + 1) * 12)
                Dim endDate As Date = tmpendDate.AddDays(-1)
                With ws.Range("b1").Offset(startRow + halfYear, 0)
                    .SetBorderLeft(2, frmUI.currentTheme.BorderColor, TableViewCellStyle.BorderStyle.Continuous)
                    .SetBorderRight(2, frmUI.currentTheme.BorderColor, TableViewCellStyle.BorderStyle.Continuous)
                    .SetBorderBottom(1, frmUI.currentTheme.BorderColor, TableViewCellStyle.BorderStyle.Continuous)
                    .SetFont(boldFont)
                    .Cell.tag = New Object() {startDate, endDate}
                    .Value = startDate.Year
                    .SetBGColor(frmUI.currentTheme.LineHeaderColor)
                    .SetFontColor(frmUI.currentTheme.LineHeaderTextColor)
                End With
                tmpRapport.datesForGraph.Add(startDate.Year)
                addToLists(startDate, endDate, const_)
                tmpPeriod += 1
                progress.Tick()
            Next
        End If
        renderTable(tmpPeriod - 1, startRow, summering)
        rowForReport = startRow + tmpPeriod + rowsBetweenReports
        progress.Close()
    End If
    ws.WorkingTable.Refresh()
End Function