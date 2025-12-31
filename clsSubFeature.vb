Public Class clsSubFeature
    Private _strDesc As String
    Private _strCode As String
    Private _strURL As String

    ' Property for strDesc
    Public Property strDesc As String
        Get
            Return _strDesc
        End Get
        Set(ByVal value As String)
            _strDesc = value
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

    ' Property for strURL
    Public Property strURL As String
        Get
            Return _strURL
        End Get
        Set(ByVal value As String)
            _strURL = value
        End Set
    End Property

    ' Constructor
    Public Sub New(Optional ByVal desc As String = "", Optional ByVal code As String = "", Optional ByVal url As String = "")
        _strDesc = desc
        _strCode = code
        _strURL = url
    End Sub

End Class

