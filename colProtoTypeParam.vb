Public Class colProtoTypeParam
    Implements IEnumerable(Of clsProtoTypeParam)

    Private ReadOnly _params As New Dictionary(Of String, clsProtoTypeParam)(StringComparer.OrdinalIgnoreCase)
    Private ReadOnly _order As New List(Of clsProtoTypeParam)

    ' Count of items
    Public ReadOnly Property Count As Integer
        Get
            Return _order.Count
        End Get
    End Property

    ' Indexer by sName
    Default Public ReadOnly Property Item(sName As String) As clsProtoTypeParam
        Get
            If _params.ContainsKey(sName) Then
                Return _params(sName)
            End If
            Throw New KeyNotFoundException($"Parameter '{sName}' not found.")
        End Get
    End Property
    ' Indexer by position (0-based)
    Default Public ReadOnly Property Item(index As Integer) As clsProtoTypeParam
        Get
            If index >= 0 AndAlso index < _order.Count Then
                Return _order(index)
            End If
            Throw New IndexOutOfRangeException($"Parameter index {index} is out of range.")
        End Get
    End Property
    ' Add by instance
    Public Sub Add(param As clsProtoTypeParam)
        If _params.ContainsKey(param.sName) Then
            Throw New ArgumentException($"Parameter with name '{param.sName}' already exists.")
        End If
        _params.Add(param.sName, param)
        _order.Add(param) ' preserve FIFO order
    End Sub

    ' Add by name and type
    Public Sub Add(sName As String, sType As String)
        Add(New clsProtoTypeParam(sName, sType))
    End Sub

    ' Update value by name
    Public Sub UpdateValue(sName As String, value As Object)
        If Not _params.ContainsKey(sName) Then
            Throw New KeyNotFoundException($"Parameter '{sName}' not found.")
        End If
        _params(sName).Value = value
    End Sub

    ' TryGet pattern
    Public Function TryGetValue(sName As String, ByRef result As clsProtoTypeParam) As Boolean
        Return _params.TryGetValue(sName, result)
    End Function

    ' FIFO enumerator
    Public Function GetEnumerator() As IEnumerator(Of clsProtoTypeParam) Implements IEnumerable(Of clsProtoTypeParam).GetEnumerator
        Return _order.GetEnumerator()
    End Function

    Private Function IEnumerable_GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
        Return _order.GetEnumerator()
    End Function
    Public Function Contains(sName As String) As Boolean
        Return _params.ContainsKey(sName)
    End Function
    ' Remove by name
    Public Function Remove(sName As String) As Boolean
        If Not _params.ContainsKey(sName) Then
            Return False
        End If

        ' Remove from dictionary
        Dim param As clsProtoTypeParam = _params(sName)
        _params.Remove(sName)

        ' Remove from ordered list
        _order.Remove(param)

        Return True
    End Function
End Class