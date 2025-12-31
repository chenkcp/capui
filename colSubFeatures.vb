Imports System.Collections
Imports System.Collections.Generic

Public Class colSubFeatures
    Implements IEnumerable(Of clsSubFeature)

    ' Local variable to hold collection
    Private mCol As List(Of clsSubFeature)

    Public Sub New()
        ' Creates the collection when this class is created
        mCol = New List(Of clsSubFeature)()
    End Sub

    Public Function Add(strDesc As String, strCode As String, strURL As String, Optional sKey As String = "") As clsSubFeature
        ' Create a new object
        Dim objNewMember As New clsSubFeature()

        ' Set the properties
        objNewMember.strDesc = strDesc
        objNewMember.strCode = strCode
        objNewMember.strURL = strURL

        ' Add to collection
        mCol.Add(objNewMember)

        ' Return the object created
        Return objNewMember
    End Function

    Default Public ReadOnly Property Item(index As Integer) As clsSubFeature
        Get
            ' Used when referencing an element in the collection
            Return mCol(index)
        End Get
    End Property

    Public ReadOnly Property Count() As Integer
        Get
            ' Retrieve the number of elements in the collection
            Return mCol.Count
        End Get
    End Property

    Public Sub Remove(index As Integer)
        ' Remove an element from the collection
        mCol.RemoveAt(index)
    End Sub

    Public Function GetEnumerator() As IEnumerator(Of clsSubFeature) Implements IEnumerable(Of clsSubFeature).GetEnumerator
        ' Allows iteration with For Each
        Return mCol.GetEnumerator()
    End Function

    Private Function IEnumerable_GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
        ' Non-generic version for compatibility
        Return GetEnumerator()
    End Function

End Class
