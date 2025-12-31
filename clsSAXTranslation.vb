Public Class clsSAXTranslation
    ' 私有字段存储属性值
    Private mvarstrProtoType As String
    Private mvarstrCSBIndex As String
    Private mvarbSelfRegister As Boolean

    ' 公共属性 - 原型类型
    Public Property strProtoType As String
        Get
            Return mvarstrProtoType
        End Get
        Set(value As String)
            mvarstrProtoType = value
        End Set
    End Property

    ' 公共属性 - CSB 索引
    Public Property strCSBIndex As String
        Get
            Return mvarstrCSBIndex
        End Get
        Set(value As String)
            mvarstrCSBIndex = value
        End Set
    End Property

    ' 公共属性 - 自注册标志
    Public Property bSelfRegister As Boolean
        Get
            Return mvarbSelfRegister
        End Get
        Set(value As Boolean)
            mvarbSelfRegister = value
        End Set
    End Property
End Class
