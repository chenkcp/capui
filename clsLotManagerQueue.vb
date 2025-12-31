Public Class clsLotManagerQueue
    ' 私有字段保存属性值
    Private mvarsLotId As String
    Private mvarnFrom As Integer
    Private mvarnTo As Integer
    Private mvardtBirth As Date

    ' 公共属性
    Public Property sLotId() As String
        Get
            Return mvarsLotId
        End Get
        Set(value As String)
            mvarsLotId = value
        End Set
    End Property

    Public Property nFrom() As Integer
        Get
            Return mvarnFrom
        End Get
        Set(value As Integer)
            mvarnFrom = value
        End Set
    End Property

    Public Property nTo() As Integer
        Get
            Return mvarnTo
        End Get
        Set(value As Integer)
            mvarnTo = value
        End Set
    End Property

    Public Property dtBirth() As Date
        Get
            Return mvardtBirth
        End Get
        Set(value As Date)
            mvardtBirth = value
        End Set
    End Property
End Class
