 '------------------------------------------------------------------------------------
 '
 '      1      Ge några förslag på hur man skulle kunna förbättra denna kod. Rangordna dem i viktighetsordning.
 '
 '------------------------------------------------------------------------------------

 ' As I apparently can not make a code review without a PR, and not be able to comment on code which haven't been changed.
 ' All suggestions are added as comments.
 '
 ' Disclaimer: I have never programmed in vb before, so there might be some erroneous assumptions made.


 ' Importance levels:
 ' The levels are defined 1-9 where 1 is most important. here are a somewhat general relevance of their meaning and categorization.
 ' With this, I might have set a somewhat small fix or a non-issue as very important, but without knowing the whole system some of
 ' these changes are hard to classify the importance of.
 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 ' 1 - Bug or could lead to bug
 ' 2 - Implementation that should be handled elsewhere
 ' 3 - risk of undefined behavior
 ' 4 - Undefined or requires clear documentation
 ' 5 - Refactor for better code
 ' 6 - Mitigate code duplication
 ' 7 - Not critical but could lead to performance deficits
 ' 8 - Could lead to misinterpretation and future bugs
 ' 9 - Nitpicking
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 

Public Function SumPeriods(ByVal fromDate As Date, ByVal toDate As Date, ByVal summering As String, ByVal const_ As String, ByVal type As Integer)
 ' Importance level: 9
 ' level is hardcoded in day and week to remove the last day of a period, but is not clearly understood by either the usage, or the variable name.
 ' month and year generalizes this to .AddDays(-1)
 ' Consider using the same logic
    Dim level As Integer
    Dim periods As Integer
    Dim startRow As Integer

    Dim progress As frmProgress
    If type = 1 Then
        startRow = 13
    Else
        startRow = rowForReport + 5
    End If

 ' Importance level: 2
 ' There are only two members in this section that are directly affected by SumPeriods: ws and tmpRapport.datesForGraph
 ' All changes to tmpRapport should be extracted to a separate method, regardless if the are related to each other or not
 ' (If they are not, then the clearing of all other data should not be done inside SumPeriods)
 ' Consider extracting to class method, e.g. tmpRapport.Clear() 
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


 ' See note: Consolidate Duplicate Code



 ' Importance level: 7
 ' String comparisons on string literal.
 ' While this comparison is only done once, future changes might change this and require multiple comparisons.
 ' This could be bad in three ways:
 ' 1, String comparisons are not efficient.
 ' 2, String literals are harder for the compiler to optimise away supplemental comparisons (compared to references to "static" data)
 ' 3, Could lead to Shotgun Surgery (or Divergent Change), if a change is made, e.g. changed to english, all literals have to be changed.
 ' Consider changing to a single parsing to enum or polymorphic class (enumeration class).
    'Dag
    If summering = "Dag" Then
        ws.Range("b1").Offset(startRow - 1, 0).Value = tmpRapport.summering
        periods = DateDiff(DateInterval.Day, fromDate, toDate)
        level = 1 - 1
        Dim tmpPeriod As Integer = 0
 ' Importance level: 9
 ' tmpDay is modified but never used.
 ' Consider removing.
        Dim tmpDay As Integer

        progress = New frmProgress("Beräknar...", periods)
        If periods > 10 Then    ' See note: Polymorph
            progress.ShowProgress()
        End If

        If rbSelectDays.Checked Then
            For Each datum As Date In dateList    ' See note: Polymorph
                If tmpRapport.weekDaysSelected.Contains(DatePart(DateInterval.Weekday, datum, FirstDayOfWeek.Monday,FirstWeekOfYear.FirstFourDays)) Then
                    addToLists(datum, datum, const_)
                    tmpPeriod += 1
                End If
            Next
        Else
            For day As Integer = 0 To periods
                ' See note: Polymorph
                If tmpRapport.weekDaysSelected.Contains(DatePart(DateInterval.Weekday, fromDate.AddDays(day), FirstDayOfWeek.Monday,FirstWeekOfYear.FirstFourDays)) Then
                    tmpDay += 1
                    Dim startDate As Date = fromDate.AddDays(day)
                    Dim endDate As Date = fromDate.AddDays(day + level)
                    ' See note: Visual inconsistency


 ' Importance level: 9
 ' Day assumes year to be 2000-2099 and cuts the first two numbers, inconsistent with other cases.
 ' Consider keeping the entire year
                    tmpRapport.datesForGraph.Add(startDate.Day & "/" & startDate.Month & "-" & startDate.Year - 2000)
                    addToLists(startDate, endDate, const_)
                    tmpPeriod += 1
                End If
                progress.Tick()
            Next
        End If

        ' See note: Visual inconsistency
        Dim extrarows As Integer = renderTable(tmpPeriod - 1, startRow, summering)
        progress.Close()
        

        rowForReport = startRow + tmpPeriod + rowsBetweenReports + extrarows
    End If


 ' See note If-ElseIf-Else
    'Vecka
    If summering = "Vecka" Then
        ws.Range("b1").Offset(startRow - 1, 0).Value = tmpRapport.summering
        level = 7 - 1
        Dim tmpPeriod As Integer = 0
        Dim currDay = DatePart(DateInterval.Weekday, fromDate, FirstDayOfWeek.Monday, FirstWeekOfYear.FirstFourDays)
        Dim weekadjust As Integer = 1 - currDay
        Dim firstDayOfW As Date = fromDate.AddDays(weekadjust)
 ' Importance level: 1
 ' Bug, reads value from uninitialized data.
 ' VB initializes all numerical data types to 0, but its quite clear that was not the intention here.
 ' Regardless, I would recommend initializing the value to 0 in source if the intention was to read the value before modification.
 ' probably faulty ordering of lines.
 ' Consider moving the assignment-expression of periods up
        progress = New frmProgress("Beräknar...", periods)
        progress.ShowProgress()
        periods = DateDiff(DateInterval.WeekOfYear, fromDate, toDate, FirstDayOfWeek.Monday, FirstWeekOfYear.FirstFourDays)
        If rbSelectDays.Checked Then
        ' See note Extract Method
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

                ' See note: Visual inconsistency
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

 ' See note If-ElseIf-Else
    'Månad
    If summering = "Månad" Then
        ws.Range("b1").Offset(startRow - 1, 0).Value = tmpRapport.summering
        Dim firstDayOfM As New Date(fromDate.Year, fromDate.Month, 1)
        periods = DateDiff(DateInterval.Month, fromDate, toDate, FirstDayOfWeek.Monday, FirstWeekOfYear.FirstFourDays)
        Dim tmpPeriod As Integer = 0
        progress = New frmProgress("Beräknar...", periods)
        progress.ShowProgress()
        If rbSelectDays.Checked Then
            ' See note Extract Method
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
 ' Importance level: 9
 ' year is an unused variable.
 ' Consider removing
                Dim year As Integer = startDate.Year
                Dim tmpendDate As Date = startDate.AddMonths(1)
                Dim endDate As Date = tmpendDate.AddDays(-1)
                ' See note: Visual inconsistency
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

    ' See note If-ElseIf-Else
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
            ' See note Extract Method
            With ws.Range("b1").Offset(startRow, 0)
                .SetColumnSpan(11)
                .Value = "ÅRSSUMMERING FUNGERAR INTE MED ENSTAKA DAGAR"
                .SetBGColor(Color.White)
                .SetFontColor(Color.Black)
                .SetFont(boldFont)
            End With
        Else
 ' Importance level: 8
 ' Variable name of each period is "halfYear", but is operated as if it signifies a whole year.
 ' If a future developer misinterpret its usage, it could lead to undefined behavior
 ' Consider renaming it
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



 '------------------------------------------------------------------------------------
 '
 '      Notes
 '
 '------------------------------------------------------------------------------------


 ' Consolidate Duplicate Code
 ' ------------------------
 ' Importance level: 5
 ' The main part or this function is that the implementation of Day, Week, Month, and Year are (almost*) identical.
 ' If changing one condition, the other cases should probably be updated as well.
 ' If not it could lead to inconsistencies, or undefined behavior.
 ' This could be solved in multiple ways:
 ' 1, Consolidating the conditional fragments, perhaps storing the value as a variable to be used later,
 '    and have a single implementation that is generalized for all cases.
 ' 2, Extracting the common implementation and logic into external methods to be called by each case with their unique difference as parameters
 ' 3, Utilize polymorphic classes to modify the evaluation of a single function body.
 '
 ' "Dag" have the largest differentiation of logic which might make some solutions more tricky
 ' (Conditional on ShowProgress() and if date should be added to report)
 ' which is why I would suggest polymorphic classes because I would been it best suited for changes in evaluation and not just in data values.


 ' Polymorph
 ' ------------------------
 ' Importance level: 5
 ' "Dag" have some conditionals that is not present in the other cases.
 ' These evaluations could be extracted to a property/method in an abstract superclass,
 ' DayClass can then override it to perform its evaluation.
 ' This would ease the consolidation into a single generalized algorithm


 ' Extract Method
 ' ------------------------
 ' Importance level: 6
 ' Week, Month and year cannot operate on a select amount of days, but have the same implementation except for a slight change
 ' of string literal.
 ' Consider extracting the implementation to a method with part or the entire string as parameter.


 ' Visual inconsistency
 ' ------------------------
 ' Importance level: 4
 ' Week, Month and year operates on a table by changing style, borders, font and color.
 ' This is not done on Days, while days modify rowForReport.
 ' Without having more understanding of the project I can not state if this is intentional or mistakes due to Divergent Change.
 ' These inconsistencies would be clarified more if the code wasn't duplicated and specialized for each case (see: Consolidate Duplicate Code)
 ' If such a refactor is inconvenient for the moment, consider either rectify these inconsistencies, or clearly document the differences as intentional


 ' If-ElseIf-Else
 ' ------------------------
 ' Importance level: 3
 ' Each case is a standalone If-case. This entails two consequences:
 ' 1, If a case have already evaluated as True, then the other cases will have to do a string comparison even though we know they will evaluate to False.
 ' 2, By a combination of no Else-case, and using comparison on a string, there are no guarantees that the if-cases have exhausted every possible alternative.
 ' Without knowing more of the system I can not determine if the function would produce an undefined behavior or not.
 '
 ' Consider first using a control-logic that utilizes short-circuiting (If-ElseIf or Select Case), but mainly using a more defensive methodology in the algorithm,
 ' either having a clear default case (Else) and/or more clearly define the possibility-set exhaustive (e.g. using enums, polymorphic classes and a generalized implementation)
      



