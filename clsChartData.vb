Public Class clsChartData
    ' Local variables to hold property values
    Private mvarstrGroupId As String ' Group ID
    Private mvarnGoodCount As Integer ' Good Count
    Private mvarnBadCount As Integer ' Bad Count
    Private mvarcolIndivBad As List(Of Integer) ' List of individual bad counts
    Private mvarstrIconName As String ' Icon Name
    Private mvarstrIconPath As String ' Icon Path
    Private mvardtBirth As Date ' Birth Date

    ' Constructor (Equivalent to Class_Initialize)
    Public Sub New()
        mvarcolIndivBad = New List(Of Integer)()
    End Sub

    ' Property for dtBirth
    Public Property dtBirth As Date
        Get
            Return mvardtBirth
        End Get
        Set(value As Date)
            mvardtBirth = value
        End Set
    End Property

    ' Property for strIconPath
    Public Property strIconPath As String
        Get
            Return mvarstrIconPath
        End Get
        Set(value As String)
            mvarstrIconPath = value
        End Set
    End Property

    ' Property for strGroupId
    Public Property strGroupId As String
        Get
            Return mvarstrGroupId
        End Get
        Set(value As String)
            mvarstrGroupId = value
        End Set
    End Property

    ' Property for strIconName
    Public Property strIconName As String
        Get
            Return mvarstrIconName
        End Get
        Set(value As String)
            mvarstrIconName = value
        End Set
    End Property

    ' Property for nGoodCount
    Public Property nGoodCount As Integer
        Get
            Return mvarnGoodCount
        End Get
        Set(value As Integer)
            mvarnGoodCount = value
        End Set
    End Property

    ' Property for nBadCount
    Public Property nBadCount As Integer
        Get
            Return mvarnBadCount
        End Get
        Set(value As Integer)
            mvarnBadCount = value
        End Set
    End Property

    ' Property for colIndivBad (handles resizing dynamically)
    Public Property colIndivBad(ByVal nIndex As Integer) As Integer
        Get
            If nIndex >= 0 AndAlso nIndex < mvarcolIndivBad.Count Then
                Return mvarcolIndivBad(nIndex)
            Else
                Return 0 ' Default value if out of bounds
            End If
        End Get
        Set(value As Integer)
            If nIndex >= 0 Then
                If nIndex >= mvarcolIndivBad.Count Then
                    While mvarcolIndivBad.Count <= nIndex
                        mvarcolIndivBad.Add(0) ' Expand list with default values
                    End While
                End If
                mvarcolIndivBad(nIndex) = value
            End If
        End Set
    End Property

    ''' <summary>
    ''' Updates defect count and returns the new total count.
    ''' </summary>
    ''' <param name="nDefectIndex">Index of the defect type.</param>
    ''' <param name="nCountIn">New defect count.</param>
    ''' <returns>Updated total count.</returns>
    Public Function ChangeDataSeries(ByRef nDefectIndex As Integer, ByRef nCountIn As Integer) As Integer
        Dim nNewCount As Integer

        ' If the defect index is greater than 0, update bad counts
        If nDefectIndex > 0 Then

            ' Ensure the list is properly sized
            While nDefectIndex >= mvarcolIndivBad.Count
                mvarcolIndivBad.Add(0)
            End While

            ' Set the new defect count
            mvarcolIndivBad(nDefectIndex) = nCountIn

            ' Recalculate total bad count;  Sums all elements except the first (index 0 is for Good only)
            mvarnBadCount = mvarcolIndivBad.Skip(1).Sum()

            ' Set return value
            nNewCount = mvarnBadCount
        Else
            ' Update good count
            mvarnGoodCount = nCountIn
            mvarcolIndivBad(0) = nCountIn
            nNewCount = mvarnGoodCount
        End If

        ' Return the new count
        Return nNewCount
    End Function
End Class


