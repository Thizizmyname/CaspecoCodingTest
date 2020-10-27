
'-------------------------------
'
'   Polymorphic class
'
'-------------------------------

Public MustInherit Class SumType

    Private rangeError As String
    Private dateInterval As DateInterval
    Private ranged As Boolean

    Public Sub New(error As String, interval As DateInterval, ranged As Boolean)
      rangeError = error
      dateInterval = interval
      ranged = ranged

    End Sub

    Public ReadOnly Property DateInterval As DateInterval
        Get
            Return dateInterval
        End Get
    End Property

    Public ReadOnly Property Ranged As Boolean
        Get
            Return ranged
        End Get
    End Property

    Public MustOverride Function CalcFirstDay(ByVal fromDate As Date) As Date
    Public MustOverride Function DateInterval() As DateInterval
    Public MustOverride Function DateofPeriod(ByVal FirstDate As Date, ByVal period As Integer) As Date
    Public MustOverride Function DateName(ByVal date As Date) As String

    Public Overridable Sub ShowProgress(progress As frmProgress, ByVal periods As Integer)
        progress.ShowProgress()
    End Sub

    Public Overridable Sub OnSelectedDays(ByRef dateList As List(Of Date), ByRef tmpPeriod As Integer)
        ErrorRangedSummationOnSelectDays()
    End Sub

    Public Overridable Function IgnoreDate(ByVal fromDate As Date, ByVal period As Integer) As Boolean
        Return True
    End Function
    Public Overridable Function IgnoreDate(ByVal datum As Date) As Boolean
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



'-------------------------------
'
'   Derived Classes
'
'-------------------------------

'   SumType.SumDay
'''''''''''''''''''''''

    Public Class SumDay
        Inherits SumType
        Public Sub New()
            MyBase.New("", DateInterval.Day, False)
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

        Overrides Sub ShowProgress(progress As frmProgress, Byval periods As Integer)
            If periods > 10 Then
                progress.ShowProgress()
            End If
        End Sub

        Overrides Sub OnSelectedDays(ByRef dateList As List(Of Date), ByRef tmpPeriod As Integer)
            For Each datum As Date In dateList
                If IgnoreDate(datum) Then
                    addToLists(datum, datum, const_)
                    tmpPeriod += 1
                End If
            Next
        End Sub

        Overrides Function IgnoreDate(Byval fromDate As Date, ByVal period As Integer) As Boolean
            Return IgnoreDate(fromDate.AddDays(day))
        End Function
        Overrides Function IgnoreDate(Byval datum As Date) As Boolean
            Return tmpRapport.weekDaysSelected.Contains(DatePart(DateInterval.Weekday, datum, FirstDayOfWeek.Monday,FirstWeekOfYear.FirstFourDays))
        End Function
    End Class


'   SumType.SumWeek
'''''''''''''''''''''''

    Public Class SumWeek
        Inherits SumType
        Public Sub New()
            MyBase.New("VECKOSUMMERING", DateInterval.WeekOfYear, True)
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


'   SumType.SumMonth
'''''''''''''''''''''''

    Public Class SumMonth
        Inherits SumType
        Public Sub New()
            MyBase.New("MÅNADSSUMMERING", DateInterval.Month, True)
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


'   SumType.SumYear
'''''''''''''''''''''''

    Public Class SumYear
        Inherits SumType
        Public Sub New()
            MyBase.New("ÅRSSUMMERING", DateInterval.Year, True)
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


'   SumType.Invalid
'''''''''''''''''''''''

' Not properly implemented but I just wanted something to signify an otherwise unhandled edge-case
    Public Class Invalid
        Inherits SumType
        
    End Class

End Class