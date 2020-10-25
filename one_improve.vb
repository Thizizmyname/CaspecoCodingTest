 '------------------------------------------------------------------------------------
 '
 '      1      Ge några förslag på hur man skulle kunna förbättra denna kod. Rangordna dem i viktighetsordning.
 '
 '------------------------------------------------------------------------------------

' Note:
' As SumPeriods uses values not included in the parameterlist I assume these are included as variables or properties of the class.
' In this example I have encapsulated everything as a SumClass to illustrate the separation of a Rapport (Class of tmpRapport)
' and the rest of implementation.

' This implementation is using enums to separate the different unique properties between the summation types.
' Although, since there quite a few, this should probably be refactored once again into polymorphic classes intsead of enum and multiple switchcases.
' All Extracted methods are below SumPeriods

Public Class SumClass


    Public Function ParseToEnum(Byval summering As String) As SumType
        Select Case summering
            Case "Dag"
                Return SumType.SumDay
            Case "Vecka"
                Return SumType.SumWeek
            Case "Månad"
                Return SumType.SumMonth
            Case "År"
                Return SumType.SumYear
            Case Else
                Return SumType.Invalid
        End Select
    End Function



    Public Function SumPeriods(ByVal fromDate As Date, ByVal toDate As Date, ByVal summering As String, ByVal const_ As String, ByVal type As Integer)
        Dim periods As Integer
        Dim startRow As Integer

        Dim sumType As SumType = ParseToEnum(summering)
    
        Dim progress As frmProgress
        If type = 1 Then
            startRow = 13
        Else
            startRow = rowForReport + 5
        End If
    
        ws.ResetRowsFrom(startRow)
        tmpRapport.Clear()

        Dim rowsBetweenReports As Integer = 0
        If tmpRapport.target > 0 Then
            rowsBetweenReports += 2
        End If
        If CBSShowPrevYear.Checked Then
            rowsBetweenReports += 2
        End If


        If sumType <> SumType.Invalid then
            ws.Range("b1").Offset(startRow - 1, 0).Value = tmpRapport.summering
            Dim tmpPeriod As Integer = 0
            periods = DateDiff(sumType.DateInterval, fromDate, toDate, FirstDayOfWeek.Monday, FirstWeekOfYear.FirstFourDays)
            progress = New frmProgress("Beräknar...", periods)

            Dim firstDayOfRange = sumType.CalcFirstDay(fromDate)

            If sumType is SumType.Day Then                
                If periods > 10 Then
                    progress.ShowProgress()
                End If

                If rbSelectDays.Checked Then
                    For Each datum As Date In dateList
                        If IgnoreDate(datum) Then
                            addToLists(datum, datum, const_)
                            tmpPeriod += 1
                        End If
                    Next
                Else
                    For day As Integer = 0 To periods
                        If IgnoreDate(fromDate.AddDays(day)) Then
                            Dim startDate As Date = sumType.DateofPeriod(firstDayOfRange, day)
                            ' While the DateofPeriod call is generalized enough to work for Days aswell, it seemed silly to calculate it when known
                            Dim endDate As Date = startDate 'sumType.DateofPeriod(firstDayOfRange, day+1).AddDays(-1)
                            tmpRapport.datesForGraph.Add(startDate.Day & "/" & startDate.Month & "-" & startDate.Year)
                            addToLists(startDate, endDate, const_)
                            tmpPeriod += 1
                        End If
                        progress.Tick()
                    Next
                End If


            ' sumType is not SumType.Day
            ElseIf sumType is Not Invalid Then
                progress.ShowProgress()

                If rbSelectDays.Checked Then
                    sumType.ErrorRangedSummationOnSelectDays()

                Else    
                    For period As Integer = 0 To periods
                        If Not IgnoreDate(fromDate, period) Then
                            Dim startDate As Date = sumType.DateofPeriod(firstDayOfRange, period)
                            Dim endDate As Date = sumType.DateofPeriod(firstDayOfRange, period+1).AddDays(-1)

                            With ws.Range("b1").Offset(startRow + period, 0)
                                .SetBorderLeft(2, frmUI.currentTheme.BorderColor, TableViewCellStyle.BorderStyle.Continuous)
                                .SetBorderRight(2, frmUI.currentTheme.BorderColor, TableViewCellStyle.BorderStyle.Continuous)
                                .SetBorderBottom(1, frmUI.currentTheme.BorderColor, TableViewCellStyle.BorderStyle.Continuous)
                                .SetFont(boldFont)
                                .Cell.tag = New Object() {startDate, endDate}
                                .Value = sumType.DateName(startdate)
                                .SetBGColor(frmUI.currentTheme.LineHeaderColor)
                                .SetFontColor(frmUI.currentTheme.LineHeaderTextColor)
                            End With
                            tmpRapport.datesForGraph.Add(sumType.DateName(date))
                            addToLists(startDate, endDate, const_)
                            tmpPeriod += 1
                        End If
                        progress.Tick()

                    Next
                End If
            End If
            Dim extrarows As Integer = renderTable(tmpPeriod - 1, startRow, summering)
            rowForReport = startRow + tmpPeriod + rowsBetweenReports

            If sumType is SumType.Day Then
                rowForReport += extrarows
            End If
            progress.Close()
        End If

        ws.WorkingTable.Refresh()
    End Function

End Class




'-------------------------------
'
'   Polymorphic class
'
'-------------------------------

Public MustInherit Class SumType

    Private rangeError As String
    Private dateInterval As DateInterval

    Public Sub New(error As String, interval As DateInterval)
      rangeError = error
      dateInterval = interval
    End Sub

    Public ReadOnly Property DateInterval As DateInterval
        Get
            Return dateInterval
        End Get
    End Property

    Public MustOverride Function CalcFirstDay(ByVal fromDate As Date) As Date
    Public MustOverride Function DateInterval() As DateInterval
    Public MustOverride Function DateofPeriod(ByVal FirstDate As Date, ByVal period As Integer) As Date
    Public MustOverride Function DateName(ByVal date As Date) As String


    Public Overridable Function IgnoreDate(Byval fromDate As Date, ByVal period As Integer) As Boolean
        Return True
    End Function
    Public Overridable Function IgnoreDate(Byval datum As Date) As Boolean
        Return True
    End Function

    Public NotOverridable Sub ErrorRangedSummationOnSelectDays(ws As Object)
        With ws.Range("b1").Offset(startRow, 0)
            .SetColumnSpan(11)
            .Value = rangeError & " FUNGERAR INTE MED ENSTAKA DAGAR"
            .SetBGColor(Color.White)
            .SetFontColor(Color.Black)
            .SetFont(boldFont)
        End With
    End Sub



    Public Class SumDay
        Inherits SumType
        Public Sub New()
            MyBase.New("", DateInterval.Day)
        End Sub

        Overrides Function CalcFirstDay(ByVal fromDate As Date) As Date
            Return fromDate
        End Function

        Overrides Function DateofPeriod(ByVal firstDate As Date, ByVal period As Integer) As Date
            Return firstDate.AddDays(period)
        End Function

        Overrides Function DateName(ByVal date As Date) As String
            Return startDate.Day & "/" & date.Month & "-" & startDate.Year
        End Function

        Overrides Function IgnoreDate(Byval fromDate As Date, ByVal period As Integer) As Boolean
            Return IgnoreDate(fromDate.AddDays(day))
        End Function
        Overrides Function IgnoreDate(Byval datum As Date) As Boolean
            Return tmpRapport.weekDaysSelected.Contains(DatePart(DateInterval.Weekday, datum, FirstDayOfWeek.Monday,FirstWeekOfYear.FirstFourDays))
        End Function
    End Class

    Public Class SumWeek
        Inherits SumType
        Public Sub New()
            MyBase.New("VECKOSUMMERING", DateInterval.WeekOfYear)
        End Sub

        Overrides Function CalcFirstDay(ByVal fromDate As Date) As Date
            Dim currDay = DatePart(DateInterval.Weekday, fromDate, FirstDayOfWeek.Monday, FirstWeekOfYear.FirstFourDays)
            Dim weekadjust As Integer = 1 - currDay
            Return fromDate.AddDays(weekadjust)
        End Function

        Overrides Function DateofPeriod(ByVal firstDate As Date, ByVal period As Integer) As Date
            Return firstDate.AddDays(period * 7)
        End Function

        Overrides Function DateName(ByVal date As Date) As String
            Return DatePart(DateInterval.WeekOfYear, date, FirstDayOfWeek.Monday, FirstWeekOfYear.FirstFourDays) & "-" & startDate.Year
        End Function
    End Class

    Public Class SumMonth
        Inherits SumType
        Public Sub New()
            MyBase.New("MÅNADSSUMMERING", DateInterval.Month)
        End Sub

        Overrides Function CalcFirstDay(ByVal fromDate As Date) As Date
            Return New Date(fromDate.Year, fromDate.Month, 1)
        End Function

        Overrides Function DateofPeriod(ByVal firstDate As Date, ByVal period As Integer) As Date
            Return firstDate.AddMonths(period)
        End Function

        Overrides Function DateName(ByVal date As Date) As String
            Return MonthName(date.Month) & "-" & startDate.Year
        End Function
    End Class

    Public Class SumYear
        Inherits SumType
        Public Sub New()
            MyBase.New("ÅRSSUMMERING", DateInterval.Year)
        End Sub

        Overrides Function CalcFirstDay(ByVal fromDate As Date) As Date
            Dim firstYear As Integer = DatePart(DateInterval.Year, fromDate, FirstDayOfWeek.Monday, FirstWeekOfYear.FirstFourDays)
            Return New Date(firstYear, 1, 1)
        End Function

        Overrides Function DateofPeriod(ByVal firstDate As Date, ByVal period As Integer) As Date
            Return firstDate.AddMonths(period * 12)
        End Function

        Overrides Function DateName(ByVal date As Date) As String
            Return date.Year
        End Function
    End Class

    Public Class Invalid
        Inherits SumType
        Public Sub New()
            MyBase.New("", DateInterval.Day)
        End Sub

        Overrides Function CalcFirstDay(ByVal fromDate As Date) As Date
            Return New Date(2001, 1, 1)
        End Function

        Overrides Function DateofPeriod(ByVal firstDate As Date, ByVal period As Integer) As Date
            Return New Date(2001, 1, 1)
        End Function

        Overrides Function DateName(ByVal date As Date) As String
            Return "Invalid SumType"
        End Function
    End Class
End Class





Public Class Rapport
    'Data members

    Public Sub Clear()
    ' Note: none of these except datesForGraph are seemingly accessed or modified in this function,
    ' It should not be this functions responsility to modify perhaps indirectly related data members.
        snittnotaSum = 0
        blgGuestSum = 0
        blgSeatSum = 0
        guestsSum = 0
        löneSaleSum = 0
        löneKostSum = 0
        arbtimSalesSum = 0
        arbtimSum = 0
        arbtimSumLP = 0
        dateCol.Clear()
        snittnotaCol1.Clear()
        snitttnotaCol2.Clear()
        blgCol.Clear()
        blgColGuests.Clear()
        blgColSeats.Clear()
        löneprocCol1.Clear()
        löneprocCol2.Clear()
        löneprocCol3.Clear()
        arbtimCol1.Clear()
        arbtimCol2.Clear()
        arbtimColLP.Clear()
        weatherCol.Clear()
        snittnotaCol1PrevYear.Clear()
        snitttnotaCol2PrevYear.Clear()
        blgColPrevYear.Clear()
        blgColGuestsPrevYear.Clear()
        blgColSeatsPrevYear.Clear()
        löneprocCol1PrevYear.Clear()
        löneprocCol2PrevYear.Clear()
        löneprocCol3PrevYear.Clear()
        arbtimCol1PrevYear.Clear()
        arbtimCol2PrevYear.Clear()
        arbtimColLPPrevYear.Clear()
        datesForGraph.Clear()
    End Sub
End Class