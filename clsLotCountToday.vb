Public Class clsLotCountToday
    ' Variables to hold property values
    Private mvarnCount As Integer ' The number of lots created today
    Private mvardtCurrentDay As Date ' Stores the current date with a time of 00:00:00

    ' Constructor (Replaces Class_Initialize)
    Public Sub New()
        mvardtCurrentDay = DateTime.Today
    End Sub

    ' Property to set the lot count
    Public Property nLotCount As Integer
        Get
            Return mvarnCount
        End Get
        Set(value As Integer)
            mvarnCount = value
        End Set
    End Property

    ''' <summary>
    ''' Adds a new lot and updates the count if the date is still the same.
    ''' Resets the count if a new day is detected.
    ''' </summary>
    ''' <param name="dtBirth">The birthday of the new lot.</param>
    ''' <returns>The updated lot count.</returns>
    Friend Function AddLot(ByVal dtBirth As Date) As Integer
        ' Check if the lot belongs to the current day
        If DateBetween(dtBirth) Then
            mvarnCount += 1
        ElseIf dtBirth >= mvardtCurrentDay.AddDays(1) Then
            ' Reset count for the new day
            mvarnCount = 1
            mvardtCurrentDay = DateTime.Today
        End If

        ' Return updated lot count
        Return mvarnCount
    End Function

    ''' <summary>
    ''' Determines what the lot count would be if a new lot were added.
    ''' </summary>
    ''' <param name="dtBirth">The birthday of the new lot.</param>
    ''' <returns>The next lot count based on date.</returns>
    Public Function NextLotCount(ByVal dtBirth As Date) As Integer
        If DateBetween(dtBirth) Then
            Return mvarnCount + 1
        ElseIf dtBirth >= mvardtCurrentDay.AddDays(1) Then
            Return 1
        End If

        Return 0 ' Default return value if none of the conditions match
    End Function

    ''' <summary>
    ''' Checks if a given date falls within the current day's range.
    ''' </summary>
    ''' <param name="dtCompare">The date to compare.</param>
    ''' <returns>True if the date is within the current day's range.</returns>
    Private Function DateBetween(ByVal dtCompare As Date) As Boolean
        Return dtCompare >= mvardtCurrentDay AndAlso dtCompare < mvardtCurrentDay.AddDays(1)
    End Function
End Class
