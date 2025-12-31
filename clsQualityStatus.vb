Public Class clsQualityStatus '这里需要替换为实际的类名
    ' 私有字段用于存储属性值
    Private mvarstrName As String
    Private mvarstrImagePath As String
    Private mvarstrMessage As String

    ' strMessage 属性
    Public Property strMessage As String
        Get
            Return mvarstrMessage
        End Get
        Set(ByVal value As String)
            mvarstrMessage = value
        End Set
    End Property

    ' strImagePath 属性
    Public Property strImagePath As String
        Get
            Return mvarstrImagePath
        End Get
        Set(ByVal value As String)
            mvarstrImagePath = value
        End Set
    End Property

    ' strName 属性
    Public Property strName As String
        Get
            Return mvarstrName
        End Get
        Set(ByVal value As String)
            mvarstrName = value
        End Set
    End Property
End Class
