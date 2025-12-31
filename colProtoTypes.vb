Public Class colProtoTypes
    Implements IEnumerable(Of clsProtoType)
    ' collection of any function or sub that can be used as a prototype for a script handler.

    Private mCol As New List(Of clsProtoType)
    Private mKeyedCol As New Dictionary(Of String, clsProtoType)(StringComparer.OrdinalIgnoreCase)

    ' Add method (optionally with a key, e.g., function name)
    Public Function Add(oProto As clsProtoType, Optional key As String = Nothing) As Integer
        mCol.Add(oProto)
        If Not String.IsNullOrEmpty(key) Then
            If Not mKeyedCol.ContainsKey(key) Then
                mKeyedCol.Add(key, oProto)
            End If
        End If
        Return mCol.Count - 1
    End Function

    ' Item property (by 0-based index or key)
    Default Public ReadOnly Property Item(indexOrKey As Object) As clsProtoType
        Get
            If TypeOf indexOrKey Is Integer Then
                Dim idx As Integer = CInt(indexOrKey)
                If idx >= 0 AndAlso idx < mCol.Count Then
                    Return mCol(idx)
                Else
                    Throw New ArgumentOutOfRangeException("Index out of range.")
                End If
            ElseIf TypeOf indexOrKey Is String Then
                Dim key As String = CStr(indexOrKey)
                If mKeyedCol.ContainsKey(key) Then
                    Return mKeyedCol(key)
                Else
                    Throw New KeyNotFoundException("Key not found: " & key)
                End If
            Else
                Throw New ArgumentException("Index must be Integer (0-based) or String key.")
            End If
        End Get
    End Property

    ' Count property
    Public ReadOnly Property Count As Integer
        Get
            Return mCol.Count
        End Get
    End Property

    ' Remove method (by 0-based index or key)
    Public Sub Remove(indexOrKey As Object)
        If TypeOf indexOrKey Is Integer Then
            Dim idx As Integer = CInt(indexOrKey)
            If idx >= 0 AndAlso idx < mCol.Count Then
                Dim obj = mCol(idx)
                mCol.RemoveAt(idx)
                ' Remove from keyed collection if present
                Dim keyToRemove = mKeyedCol.FirstOrDefault(Function(kvp) kvp.Value Is obj).Key
                If keyToRemove IsNot Nothing Then
                    mKeyedCol.Remove(keyToRemove)
                End If
            End If
        ElseIf TypeOf indexOrKey Is String Then
            Dim key As String = CStr(indexOrKey)
            If mKeyedCol.ContainsKey(key) Then
                Dim obj = mKeyedCol(key)
                mCol.Remove(obj)
                mKeyedCol.Remove(key)
            End If
        End If
    End Sub

    ' Enumerator for For Each support
    Public Function GetEnumerator() As IEnumerator(Of clsProtoType) Implements IEnumerable(Of clsProtoType).GetEnumerator
        Return mCol.GetEnumerator()
    End Function

    Private Function IEnumerable_GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
        Return mCol.GetEnumerator()
    End Function
End Class
