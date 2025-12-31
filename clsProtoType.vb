Public Class clsProtoType

    'local variable(s) to hold property value(s)
    Private mvarsClassName As String 'local copy
    Private mvarsFunctionName As String 'local copy
    Private mvarsReturnType As String 'local copy
    Private mvaroParams As colProtoTypeParam 'local copy
    'local variable(s) to hold property value(s)
    Private mvarsCodeBody As String 'local copy
    'local variable(s) to hold property value(s)
    Private mvarsMode As String 'local copy
    Private mvarsFullProtoType As String
    Private mvarsProtoReturnType As String

    Public Property sClassName As String
        Get
            Return mvarsClassName
        End Get
        Set(ByVal vData As String)
            mvarsClassName = vData
        End Set
    End Property

    Public Property sFunctionName As String
        Get
            Return mvarsFunctionName
        End Get
        Set(ByVal vData As String)
            mvarsFunctionName = vData
        End Set
    End Property

    Public Property sReturnType As String
        Get
            Return mvarsReturnType
        End Get
        Set(ByVal vData As String)
            mvarsReturnType = vData
        End Set
    End Property

    Public Property oParams As colProtoTypeParam
        Get
            Return mvaroParams
        End Get
        Set(ByVal vData As colProtoTypeParam)
            mvaroParams = vData
        End Set
    End Property

    Public Property sCodeBody As String
        Get
            Return mvarsCodeBody
        End Get
        Set(ByVal vData As String)
            mvarsCodeBody = vData
        End Set
    End Property

    Public Property sMode As String
        Get
            Return mvarsMode
        End Get
        Set(ByVal vData As String)
            mvarsMode = vData
        End Set
    End Property

    Public Property sProtoReturnType As String
        Get
            Return mvarsProtoReturnType
        End Get
        Set(ByVal vData As String)
            mvarsProtoReturnType = vData
        End Set
    End Property

    Public Property sFullProtoType As String
        Get
            Return mvarsFullProtoType
        End Get
        Set(ByVal vData As String)
            mvarsFullProtoType = vData
        End Set
    End Property
End Class

