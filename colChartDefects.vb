Public Class colChartDefects
    Implements IEnumerable(Of ClsChartDefect)

    Private mCol As New List(Of ClsChartDefect)()
    Private mKeyDict As New Dictionary(Of String, ClsChartDefect)(StringComparer.OrdinalIgnoreCase)

    ' Add a new defect to the collection
    Public Function Add(
        strDesc As String,
        nColor As Integer,
        Optional sKey As String = Nothing
    ) As ClsChartDefect

        Dim objNewMember As New ClsChartDefect With {
            .strDesc = strDesc,
            .nColor = nColor
        }

        mCol.Add(objNewMember)
        If Not String.IsNullOrEmpty(sKey) Then
            mKeyDict(sKey) = objNewMember
        End If

        Return objNewMember
    End Function
    Default Public ReadOnly Property Item(indexKey As Object) As clsChartDefect
        Get
            ' indexKey can be Integer (1-based index) or String (key)
            If TypeOf indexKey Is Integer Then
                ' VB.NET Collection is 1-based like VB6
                Return CType(mCol(indexKey), clsChartDefect)
            ElseIf TypeOf indexKey Is String Then
                Return CType(mCol(indexKey), clsChartDefect)
            Else
                Throw New ArgumentException("Index must be Integer or String")
            End If
        End Get
    End Property
    '' Indexer by integer (1-based, for VB6 compatibility)
    'Default Public ReadOnly Property Item(index As Integer) As ClsChartDefect
    '    Get
    '        Return mCol(index - 1)
    '    End Get
    'End Property

    '' Indexer by key
    'Public ReadOnly Property Item(key As String) As ClsChartDefect
    '    Get
    '        Return mKeyDict(key)
    '    End Get
    'End Property

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
    Public Function GetEnumerator() As IEnumerator(Of ClsChartDefect) Implements IEnumerable(Of ClsChartDefect).GetEnumerator
        Return mCol.GetEnumerator()
    End Function

    Private Function IEnumerable_GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
        Return mCol.GetEnumerator()
    End Function
End Class
