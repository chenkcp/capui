Imports System.Collections
Imports System.Collections.Generic

Public Class colHotFeatures
    Implements IEnumerable(Of clsHotFeature)

    ' Dictionary to store the collection with key access
    Private mCol As Dictionary(Of String, clsHotFeature)

    ' Constructor
    Public Sub New()
        mCol = New Dictionary(Of String, clsHotFeature)()
    End Sub

    ' Add method
    Public Function Add(lngFeature As Long, lngSubFeature As Long, Optional sKey As String = Nothing) As clsHotFeature
        Dim objNewMember As New clsHotFeature With {
            .lngFeature = lngFeature,
            .lngSubFeature = lngSubFeature
        }

        ' Generate a unique key if not provided
        If String.IsNullOrEmpty(sKey) Then
            sKey = Guid.NewGuid().ToString() ' Auto-generate unique key if none provided
        End If

        ' Add to dictionary
        mCol(sKey) = objNewMember

        ' Return the newly added object
        Return objNewMember
    End Function

    ' Default property to access elements by key or index
    Default Public ReadOnly Property Item(ByVal index As Object) As clsHotFeature
        Get
            If TypeOf index Is String AndAlso mCol.ContainsKey(CStr(index)) Then
                Return mCol(CStr(index))
            ElseIf TypeOf index Is Integer AndAlso CInt(index) >= 0 AndAlso CInt(index) < mCol.Count Then
                Return mCol.Values.ElementAtOrDefault(CInt(index))
            Else
                Throw New KeyNotFoundException("Key or index not found.")
            End If
        End Get
    End Property
    ' Add this method inside colHotFeatures class
    Public Sub Clear()
        mCol.Clear()
    End Sub
    ' Count property
    Public ReadOnly Property Count() As Integer
        Get
            Return mCol.Count
        End Get
    End Property

    ' Remove method
    Public Sub Remove(ByVal vntIndexKey As String)
        ' delete the first element
        If mCol.Count > 0 Then
            Dim firstKey As String = mCol.Keys.First()
            mCol.Remove(firstKey)
        Else
            Throw New KeyNotFoundException("Key not found in the collection.")
        End If
    End Sub

    ' Enumerator for For Each support
    Public Function GetEnumerator() As IEnumerator(Of clsHotFeature) Implements IEnumerable(Of clsHotFeature).GetEnumerator
        Return mCol.Values.GetEnumerator()
    End Function

    Private Function IEnumerable_GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
        Return GetEnumerator()
    End Function
End Class
