Imports System.IO

Public Class colFeatureClasses
    Private mCol As New List(Of clsFeatureClass)

    Public Function Add(strTitle As String, strClassType As String, lngSeverity As Integer, colClassFeatures As colFeatures, colHotFeature As colHotFeatures, lngHotMax As Integer, strUnitType As String, nColor As Integer, Optional sKey As String = "") As clsFeatureClass
        Dim objNewMember As New clsFeatureClass With {
            .strTitle = strTitle,
            .strClassType = strClassType,
            .lngSeverity = lngSeverity,
            .strUnitType = strUnitType,
            .colClassFeatures = colClassFeatures,
            .colHotFeature = colHotFeature,
            .lngHotMax = lngHotMax,
            .nColor = nColor,
            .strIndex = sKey
        }

        If String.IsNullOrEmpty(sKey) Then
            mCol.Add(objNewMember)
        Else
            mCol.Add(objNewMember) ' Keys are not directly supported in List
        End If

        Return objNewMember
    End Function
    Public Function ToList() As List(Of clsFeatureClass)
        Return mCol
    End Function

    Default Public ReadOnly Property Item(vntIndexKey As Integer) As clsFeatureClass
        Get
            Return mCol(vntIndexKey)
        End Get
    End Property

    Public ReadOnly Property Count() As Integer
        Get
            Return mCol.Count
        End Get
    End Property

    Public Sub Remove(vntIndexKey As Integer)
        If vntIndexKey >= 0 AndAlso vntIndexKey < mCol.Count Then
            mCol.RemoveAt(vntIndexKey)
        End If
    End Sub

    Public Sub ResetHotlist()
        For Each oFeatureClass As clsFeatureClass In mCol
            oFeatureClass.colHotFeature.Clear()
        Next
    End Sub

    Public Function IsIndexUsed(strIndex As String) As Boolean ' Changed from ByRef to ByVal
        Return mCol.Any(Function(c) c.strIndex = strIndex)
    End Function

End Class

