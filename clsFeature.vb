Public Class clsFeature
    Private _strDesc As String
    Private _strCode As String
    Private _strURL As String
    Private _colSub As colSubFeatures ' Change from List(Of colSubFeatures) to colSubFeatures

    ' Property for colSub
    Public Property colSub As colSubFeatures
        Get
            Return _colSub
        End Get
        Set(ByVal value As colSubFeatures)
            _colSub = value
        End Set
    End Property

    ' Property for strURL
    Public Property strURL As String
        Get
            Return _strURL
        End Get
        Set(ByVal value As String)
            _strURL = value
        End Set
    End Property

    ' Property for strCode
    Public Property strCode As String
        Get
            Return _strCode
        End Get
        Set(ByVal value As String)
            _strCode = value
        End Set
    End Property

    ' Property for strDesc
    Public Property strDesc As String
        Get
            Return _strDesc
        End Get
        Set(ByVal value As String)
            _strDesc = value
        End Set
    End Property

    ' Constructor
    Public Sub New()
        _colSub = New colSubFeatures() ' Initialize colSubFeatures correctly
    End Sub
End Class
