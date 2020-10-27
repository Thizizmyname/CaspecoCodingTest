 '------------------------------------------------------------------------------------
 '
 '      1      Ge några förslag på hur man skulle kunna förbättra denna kod. Rangordna dem i viktighetsordning.
 '
 '------------------------------------------------------------------------------------

' Note:
' As SumPeriods uses values not included in the parameterlist I assume these are included as variables or properties of the class.
' In this example I have encapsulated everything as a SumClass to illustrate the separation of a Rapport (Class of tmpRapport)
' and the rest of implementation.

' This implementation is using a polymorhpic class to separate the different unique properties between the summation types.
' All Extracted methods for sunType are in a separate file.

' Disclaimer:
' I have never worked in VB before and as this is just a code snippet of a larger system, I have had no way of
' checking for syntactical or logical errors. Expect faults in the code or my reasoning.

imports SumType

Public Class SumClass


    Public Function ParseToEnum(Byval summering As String) As SumType
        Select Case summering
            Case "Dag"
                Return New SumType.SumDay
            Case "Vecka"
                Return New SumType.SumWeek
            Case "Månad"
                Return New SumType.SumMonth
            Case "År"
                Return New SumType.SumYear
            Case Else
                Return New SumType.Invalid
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

            sumType.ShowProgress()

            If sumType is Not Invalid Then

                If rbSelectDays.Checked Then
                    sumType.OnSelectedDays(dateList, tmpPeriod)

                Else    
                    For period As Integer = 0 To periods
                        If Not IgnoreDate(fromDate, period) Then
                            Dim startDate As Date = sumType.DateofPeriod(firstDayOfRange, period)
                            Dim endDate As Date = sumType.DateofPeriod(firstDayOfRange, period+1).AddDays(-1)

                            if sumType.Ranged Then
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
                            End If
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

            If Not sumType.Ranged Then
                rowForReport += extrarows
            End If
            progress.Close()
        End If

        ws.WorkingTable.Refresh()
    End Function

End Class




Public Class Rapport
    'Class members

    Public Sub Clear()
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