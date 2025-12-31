Public Class colCharts
    Implements IEnumerable(Of clsChart)

    Private mCol As New List(Of clsChart)()
    Private mKeyDict As New Dictionary(Of String, ClsChart)(StringComparer.OrdinalIgnoreCase)

    ' Add a new chart to the collection
    Public Function Add(
        strIndex As String,
        nLimitFillPattern As Integer,
        lngLimitHighValue As Long,
        strLimitHighLabel As String,
        nLimitLinePattern As Integer,
        lngLimitLowValue As Long,
        strLimitLowLabel As String,
        strUnitType As String,
        strLegendGood As String,
        strLegendBad As String,
        nColorGood As Integer,
        nColorBad As Integer,
        colChartDefects As ColChartDefects,
        Optional sKey As String = Nothing
    ) As ClsChart

        Dim objNewMember As New ClsChart With {
            .strIndex = strIndex,
            .nLimitFillPattern = nLimitFillPattern,
            .lngLimitHighValue = lngLimitHighValue,
            .strLimitHighLabel = strLimitHighLabel,
            .nLimitLinePattern = nLimitLinePattern,
            .lngLimitLowValue = lngLimitLowValue,
            .strLimitLowLabel = strLimitLowLabel,
            .strUnitType = strUnitType,
            .strLegendGood = strLegendGood,
            .strLegendBad = strLegendBad,
            .nColorGood = nColorGood,
            .nColorBad = nColorBad,
            .colChartDefects = colChartDefects
        }

        ' Use strIndex as key if provided
        If Not String.IsNullOrEmpty(strIndex) Then sKey = strIndex

        mCol.Add(objNewMember)
        If Not String.IsNullOrEmpty(sKey) Then
            mKeyDict(sKey) = objNewMember
        End If

        Return objNewMember
    End Function
    Default Public ReadOnly Property Item(indexKey As Object) As clsChart
        Get
            ' indexKey can be Integer (1-based index) or String (key)
            If TypeOf indexKey Is Integer Then
                Return CType(mCol(indexKey), clsChart)
            ElseIf TypeOf indexKey Is String Then
                Return CType(mCol(indexKey), clsChart)
            Else
                Throw New ArgumentException("Index must be Integer or String")
            End If
        End Get
    End Property

    ' Count property
    Public ReadOnly Property Count As Integer
        Get
            Return mCol.Count
        End Get
    End Property

    ' Remove by index or key
    Public Sub Remove(indexOrKey As Object)
        If TypeOf indexOrKey Is Integer Then
            Dim idx As Integer = CInt(indexOrKey) - 1
            If idx >= 0 AndAlso idx < mCol.Count Then
                Dim obj = mCol(idx)
                mCol.RemoveAt(idx)
                ' Remove from key dictionary if present
                Dim keyToRemove = mKeyDict.FirstOrDefault(Function(kvp) kvp.Value Is obj).Key
                If keyToRemove IsNot Nothing Then mKeyDict.Remove(keyToRemove)
            End If
        ElseIf TypeOf indexOrKey Is String Then
            Dim key As String = CStr(indexOrKey)
            If mKeyDict.ContainsKey(key) Then
                Dim obj = mKeyDict(key)
                mCol.Remove(obj)
                mKeyDict.Remove(key)
            End If
        End If
    End Sub

    ' Enumerator support
    Public Function GetEnumerator() As IEnumerator(Of ClsChart) Implements IEnumerable(Of ClsChart).GetEnumerator
        Return mCol.GetEnumerator()
    End Function

    Private Function IEnumerable_GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
        Return mCol.GetEnumerator()
    End Function
End Class
