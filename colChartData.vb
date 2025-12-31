Imports System.Collections.Generic

Public Class colChartData
    Implements IEnumerable(Of clsChartData)

    ' Local variable to hold collection
    Private mCol As List(Of clsChartData)
    ' Indicates whether this is the Globally active Chart or not
    Private mvarbActive As Boolean

    ' Constructor (Equivalent to Class_Initialize)
    Public Sub New()
        mCol = New List(Of clsChartData)()
        mvarbActive = False
    End Sub

    ' Destructor (Equivalent to Class_Terminate)
    Protected Overrides Sub Finalize()
        mCol = Nothing
    End Sub

    ' Add method
    Public Function Add(strGroupId As String, dtBirth As Date, nGoodCount As Integer, nBadCount As Integer,
                        nMaxDefectNum As Integer, Optional strIconPath As String = "", Optional strIconName As String = "",
                        Optional sKey As String = "") As clsChartData
        ' Create a new object
        Dim objNewMember As New clsChartData()

        ' Set properties
        objNewMember.strGroupId = strGroupId
        objNewMember.dtBirth = dtBirth
        objNewMember.nGoodCount = nGoodCount
        objNewMember.nBadCount = nBadCount
        ' Partition the Defect Array
        objNewMember.colIndivBad(nMaxDefectNum) = 0
        objNewMember.strIconPath = strIconPath
        objNewMember.strIconName = strIconName

        ' Add to collection
        If String.IsNullOrEmpty(sKey) Then
            mCol.Add(objNewMember)
        Else
            ' Assuming unique keys, if required use a Dictionary
            mCol.Add(objNewMember) ' No direct key-based add in List(Of T)
        End If

        ' Return created object
        Return objNewMember
    End Function

    ' Retrieve an item by index
    Default Public ReadOnly Property Item(vntIndexKey As Integer) As clsChartData
        Get
            Return mCol(vntIndexKey) ' FFFF Adjust for 0-based index in VB.NET
        End Get
    End Property

    ' Get the count of items in collection
    Public ReadOnly Property Count() As Integer
        Get
            Return mCol.Count
        End Get
    End Property

    ' Remove an item from the collection
    Public Sub Remove(vntIndexKey As Integer)
        If vntIndexKey > 0 AndAlso vntIndexKey < mCol.Count Then
            mCol.RemoveAt(vntIndexKey) 'FFFF Adjust for 0-based index
        End If
    End Sub

    ' Enable iteration using For Each
    Public Function GetEnumerator() As IEnumerator(Of clsChartData) Implements IEnumerable(Of clsChartData).GetEnumerator
        Return mCol.GetEnumerator()
    End Function

    Private Function IEnumerable_GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
        Return GetEnumerator()
    End Function

    ' Property to check active state
    Public Property bActive() As Boolean
        Get
            Return mvarbActive
        End Get
        Set(value As Boolean)
            mvarbActive = value
        End Set
    End Property

    ' Checks if GroupId exists in the collection
    Public Function GroupIdExist(ByRef strGroupId As String, ByRef dtBirth As Date) As Integer
        For nLoop As Integer = mCol.Count - 1 To 0 Step -1
            If mCol(nLoop).strGroupId = strGroupId AndAlso mCol(nLoop).dtBirth = dtBirth Then
                Return nLoop + 1 'FFFF  Adjust for 1-based index
            End If
        Next
        Return 0 ' Not found
    End Function

    ' Checks the quality status of a GroupId
    Public Function GroupQualityStatus(ByRef strGroupId As String, ByRef dtBirth As Date) As Integer
        For nLoop As Integer = mCol.Count - 1 To 0 Step -1
            If mCol(nLoop).strGroupId = strGroupId AndAlso mCol(nLoop).dtBirth = dtBirth Then
                Return nLoop  'FFFF Adjust for 1-based index
            End If
        Next
        Return -1 ' Not found
    End Function

End Class

