Imports System
Imports System.Collections.Generic
Imports System.Security.Cryptography

Public Class colParameters
    Private mCol As Dictionary(Of String, clsParameter)


    Public Sub New()
        ' Initializes the collection when this class is created
        mCol = New Dictionary(Of String, clsParameter)()
    End Sub
    Public Sub Add(strTitle As String, strValue As String)
        If Not mCol.ContainsKey(strTitle) Then
            Dim objNewMember As New clsParameter With {
                .Title = strTitle,
                .Value = strValue
            }
            mCol.Add(strTitle, objNewMember)
        End If
    End Sub

    Public Function Contains(value As String) As Boolean
        Return mCol.Values.Any(Function(p) p.Value = value)
    End Function
    Public ReadOnly Property Item(ByVal key As String) As clsParameter
        Get
            If mCol.ContainsKey(key) Then
                Return mCol(key)
            Else
                Return Nothing
            End If
        End Get
    End Property

    Public ReadOnly Property Count As Integer
        Get
            Return mCol.Count
        End Get
    End Property

    Public Sub Remove(ByVal key As String)
        If mCol.ContainsKey(key) Then
            mCol.Remove(key)
        End If
    End Sub

    Public Function GetEnumerator() As IEnumerator(Of KeyValuePair(Of String, clsParameter))
        Return mCol.GetEnumerator()
    End Function
End Class

