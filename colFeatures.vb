Imports System.Collections
Imports System.Collections.Generic
Imports System.Windows.Forms
Public Class colFeatures
    Implements IEnumerable(Of clsFeature)

    ' Local variable to hold collection
    Private mCol As List(Of clsFeature)

    Public Sub New()
        ' Creates the collection when this class is created
        mCol = New List(Of clsFeature)()
    End Sub

    Public Function Add(strDesc As String, strCode As String, strURL As String, colSub As colSubFeatures, Optional sKey As String = "") As clsFeature
        ' Create a new object
        Dim objNewMember As New clsFeature With {
            .strDesc = strDesc,
            .strCode = strCode,
            .strURL = strURL,
            .colSub = colSub
        }

        ' Add object to the collection
        mCol.Add(objNewMember)

        ' Return the object created
        Return objNewMember
    End Function

    Default Public ReadOnly Property Item(index As Integer) As clsFeature
        Get
            If index >= 0 AndAlso index < mCol.Count Then
                Return mCol(index)
            Else
                Throw New ArgumentOutOfRangeException("Index out of range")
            End If
        End Get
    End Property

    Public ReadOnly Property Count() As Integer
        Get
            Return mCol.Count
        End Get
    End Property

    Public Sub Remove(index As Integer)
        If index >= 0 AndAlso index < mCol.Count Then
            mCol.RemoveAt(index)
        Else
            Throw New ArgumentOutOfRangeException("Index out of range")
        End If
    End Sub
    Public Function GetEnumerator() As IEnumerator(Of clsFeature) Implements IEnumerable(Of clsFeature).GetEnumerator
        Return mCol.GetEnumerator()
    End Function

    Private Function IEnumerable_GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
        Return mCol.GetEnumerator()
    End Function
End Class
