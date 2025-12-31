

Public Class colQualityStatuses
    Implements IEnumerable(Of KeyValuePair(Of String, clsQualityStatus)) 'FFFF
    ' 使用泛型字典来存储 clsQualityStatus 对象
    Private mCol As New Dictionary(Of String, clsQualityStatus)
    Private _statuses As New Dictionary(Of String, clsQualityStatus)(StringComparer.OrdinalIgnoreCase)

    Public Function Add(ByVal strName As String, ByVal strImagePath As String, ByVal strMessage As String, Optional ByVal sKey As String = "") As clsQualityStatus
        Dim objNewMember As New clsQualityStatus()

        ' 设置传入方法的属性
        objNewMember.strName = strName
        objNewMember.strImagePath = strImagePath
        objNewMember.strMessage = strMessage

        If String.IsNullOrEmpty(sKey) Then
            ' 生成一个唯一的键，这里简单使用名称作为键，你可以根据实际需求调整
            sKey = strName
        End If

        mCol.Add(sKey, objNewMember)

        ' 返回创建的对象
        Return objNewMember
    End Function

    Default Public ReadOnly Property Item(ByVal vntIndexKey As Object) As clsQualityStatus
        Get
            If TypeOf vntIndexKey Is String Then
                Dim key As String = CStr(vntIndexKey)
                If mCol.ContainsKey(key) Then
                    Return mCol(key)
                End If
            ElseIf TypeOf vntIndexKey Is Integer Then
                Dim index As Integer = CInt(vntIndexKey)
                If index >= 0 AndAlso index < mCol.Count Then
                    Return mCol.ElementAt(index).Value
                End If
            End If
            Return Nothing
        End Get
    End Property

    Public ReadOnly Property Count() As Integer
        Get
            Return mCol.Count
        End Get
    End Property

    Public Sub Remove(ByVal vntIndexKey As Object)
        If TypeOf vntIndexKey Is String Then
            Dim key As String = CStr(vntIndexKey)
            If mCol.ContainsKey(key) Then
                mCol.Remove(key)
            End If
        ElseIf TypeOf vntIndexKey Is Integer Then
            Dim index As Integer = CInt(vntIndexKey)
            If index >= 0 AndAlso index < mCol.Count Then
                Dim keyToRemove = mCol.ElementAt(index).Key
                mCol.Remove(keyToRemove)
            End If
        End If
    End Sub

    Public Function GetEnumerator() As IEnumerator(Of KeyValuePair(Of String, clsQualityStatus)) Implements IEnumerable(Of KeyValuePair(Of String, clsQualityStatus)).GetEnumerator
        Return mCol.GetEnumerator()
    End Function

    Private Function IEnumerable_GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
        Return mCol.GetEnumerator()
    End Function

    Public Function IsUsed(ByVal strName As String) As Boolean
        For Each kvp In mCol
            If kvp.Value.strName = strName Then
                Return True
            End If
        Next
        Return False
    End Function
End Class
