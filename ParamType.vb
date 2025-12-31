Public Class ParamType
    Private mvarsType As String  ' 私有字段存储属性值

    ' 公共属性 - 访问模式（公有/私有）
    Public Property sType As String
        Get
            Return mvarsType
        End Get
        Set(value As String)
            mvarsType = value
        End Set
    End Property
End Class
