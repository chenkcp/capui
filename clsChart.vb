Public Class clsChart
    Public Property StrIndex As String
    Public Property NLimitFillPattern As Integer
    Public Property LngLimitHighValue As Long
    Public Property StrLimitHighLabel As String
    Public Property NLimitLinePattern As Integer
    Public Property LngLimitLowValue As Long
    Public Property StrLimitLowLabel As String
    Public Property StrUnitType As String
    Public Property StrLegendGood As String
    Public Property StrLegendBad As String
    Public Property NColorGood As Integer
    Public Property NColorBad As Integer
    Public Property ColChartDefects As colChartDefects

    ' Returns the 1-based index of the defect, or 0 if not found (for "Good")
    Public Function GetDefectIndex(defectName As String) As Integer
        If ColChartDefects Is Nothing Then Return 0
        For i As Integer = 0 To ColChartDefects.Count - 1
            If String.Equals(ColChartDefects(i).StrDesc, defectName, StringComparison.OrdinalIgnoreCase) Then
                Return i
            End If
        Next
        Return 0
    End Function
End Class
