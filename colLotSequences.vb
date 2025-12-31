Imports System.Collections
Imports System.Collections.Generic
Imports NextcapServer = ncapsrv

''' <summary>
''' Collection of LotSequence which acts as a stack for controlling the sequencing of the EndOfLot, LotStatus and NewLot forms.
''' </summary>
Public Class colLotSequences
    Implements IEnumerable(Of clsLotSequence)

    ' Internal storage: key to LotSequence, and ordered list for index access
    Private ReadOnly _dict As New Dictionary(Of String, clsLotSequence)(StringComparer.OrdinalIgnoreCase)
    Private ReadOnly _list As New List(Of clsLotSequence)()

    ' Validate if a key exists in the collection
    Public Function ValidKey(key As String) As Boolean
        Return _dict.ContainsKey(key)
    End Function

    ' Set the Quality Status of a Group based on the Id and birth specified
    Public Sub SetQualityStatus(groupId As String, birth As Date, qualityStatus As String)
        If String.IsNullOrEmpty(qualityStatus) Then qualityStatus = "UNKNOWN"
        For Each seq In _list
            If seq.GroupId = groupId AndAlso seq.Birth = birth Then
                seq.QualityResult = qualityStatus
                seq.RequestQuality = 1
                Exit Sub
            End If
        Next
    End Sub

    ' Request creation of a new LotSequence (does not add to collection)
    Public Function RequestCreateLot(requester As frmLotManagerSink) As clsLotSequence
        Dim seq As New clsLotSequence With {
            .CreateLotDone = "PENDING",
            .RequestingManager = requester
        }
        ' NOTE: In .NET, you should trigger the timer in the relevant form/controller, not here.
        Return seq
    End Function

    ' Add a new LotSequence to the collection
    Public Function Add(groupId As String, birth As Date, lotManager As NextcapServer.clsLotManager, chartData As colChartData, requestingManager As frmLotManagerSink, Optional key As String = Nothing) As clsLotSequence
        Dim seq As New clsLotSequence With {
            .GroupId = groupId,
            .Birth = birth,
            .LotManager = lotManager,
            .ChartData = chartData,
            .RequestingManager = requestingManager
        }
        If String.IsNullOrEmpty(key) Then key = CreateUniqueKey()
        seq.Index = key

        _dict(key) = seq
        _list.Add(seq)

        ' NOTE: In .NET, you should trigger the timer in the relevant form/controller, not here.
        Return seq
    End Function

    ' Remove a LotSequence by key or index
    Public Sub Remove(keyOrIndex As Object)
        If TypeOf keyOrIndex Is String Then
            Dim key As String = CStr(keyOrIndex)
            If _dict.ContainsKey(key) Then
                Dim seq = _dict(key)
                _list.Remove(seq)
                _dict.Remove(key)
            End If
        ElseIf TypeOf keyOrIndex Is Integer Then
            Dim idx As Integer = CInt(keyOrIndex)
            If idx >= 0 AndAlso idx < _list.Count Then
                Dim seq = _list(idx)
                _dict.Remove(seq.Index)
                _list.RemoveAt(idx)
            End If
        End If
    End Sub

    ' Get a LotSequence by key or index
    Default Public ReadOnly Property Item(keyOrIndex As Object) As clsLotSequence
        Get
            If TypeOf keyOrIndex Is String Then
                Dim key As String = CStr(keyOrIndex)
                If _dict.ContainsKey(key) Then
                    Return _dict(key)
                End If
            ElseIf TypeOf keyOrIndex Is Integer Then
                Dim idx As Integer = CInt(keyOrIndex)
                If idx >= 0 AndAlso idx < _list.Count Then
                    Return _list(idx)
                End If
            End If
            Throw New KeyNotFoundException("Invalid key or index for LotSequences collection.")
        End Get
    End Property

    ' Number of items in the collection
    Public ReadOnly Property Count As Integer
        Get
            Return _list.Count
        End Get
    End Property

    ' Enumerator for For Each support
    Public Function GetEnumerator() As IEnumerator(Of clsLotSequence) Implements IEnumerable(Of clsLotSequence).GetEnumerator
        Return _list.GetEnumerator()
    End Function

    Private Function IEnumerable_GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
        Return _list.GetEnumerator()
    End Function

    ' Helper: Generate a unique key (replace with your own logic if needed)
    Private Shared Function CreateUniqueKey() As String
        Return DateTime.Now.ToString("yyyyMMddHHmmssfff")
    End Function
End Class
