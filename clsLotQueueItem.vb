Public Class clsLotQueueItem
    Public Property LotId As String
    Public Property FromState As Integer
    Public Property ToState As Integer
    Public Property BirthDate As Date

    Public Sub New(lotId As String, fromState As Integer, toState As Integer, birthDate As Date)
        Me.LotId = lotId
        Me.FromState = fromState
        Me.ToState = toState
        Me.BirthDate = birthDate
    End Sub
End Class
