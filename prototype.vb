Imports System.ComponentModel

Public Class prototype
    ' 私有字段存储属性值
    Private mvarsClassName As String
    Private mvarsFunctionName As String
    Private mvarsReturnType As String
    Private mvaroParams As ParamTypes
    Private mvarsCodeBody As String
    Private mvarsMode As String
    Private mvarsFullProtoType As String

    ' 公共属性 - 完整原型
    Public Property sFullProtoType As String
        Get
            Return mvarsFullProtoType
        End Get
        Set(value As String)
            mvarsFullProtoType = value
        End Set
    End Property

    ' 公共属性 - 访问模式（公有/私有）
    <Description("List if public or private")>
    Public Property sMode As String
        Get
            Return mvarsMode
        End Get
        Set(value As String)
            mvarsMode = value
        End Set
    End Property

    ' 公共属性 - 代码体（适合加载到 SAX 的代码）
    <Description("The code suitable for loading into SAX")>
    Public Property sCodeBody As String
        Get
            Return mvarsCodeBody
        End Get
        Set(value As String)
            mvarsCodeBody = value
        End Set
    End Property

    ' 公共属性 - 参数对象
    Public Property oParams As ParamTypes
        Get
            Return mvaroParams
        End Get
        Set(value As ParamTypes)
            mvaroParams = value
        End Set
    End Property

    ' 公共属性 - 返回类型
    Public Property sReturnType As String
        Get
            Return mvarsReturnType
        End Get
        Set(value As String)
            mvarsReturnType = value
        End Set
    End Property

    ' 公共属性 - 函数名
    Public Property sFunctionName As String
        Get
            Return mvarsFunctionName
        End Get
        Set(value As String)
            mvarsFunctionName = value
        End Set
    End Property

    ' 公共属性 - 类名
    Public Property sClassName As String
        Get
            Return mvarsClassName
        End Get
        Set(value As String)
            mvarsClassName = value
        End Set
    End Property
End Class
