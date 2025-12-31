Public Class Security
    ' 私有字段，用于存储用户权限级别字符串
    Private mvarstrUserAuthority As String
    ' 私有字段，用于存储当前系统权限级别
    Private mvarnAuthority As Integer
    ' 私有字段，用于存储客户端应用程序的权限级别
    Private mvarnClientAuthority As Integer
    ' 私有字段，用于存储是否需要登录的标志
    Private mvarbLoginRequired As Boolean

    ' 定义 SecurityChange 事件
    Public Event SecurityChange(ByVal nNewAuthority As Integer)

    ' bLoginRequired 属性的设置器
    Public Property bLoginRequired As Boolean
        Set(ByVal value As Boolean)
            mvarbLoginRequired = value
        End Set
        Get
            Return mvarbLoginRequired
        End Get
    End Property

    ' nClientAuthority 属性的设置器
    Public Property nClientAuthority As Integer
        Set(ByVal value As Integer)
            mvarnClientAuthority = value
            '/*Set the authority level of the client
            '/*if -1:Not Set Yet or Less than the Current User value
            If mvarnAuthority = -1 Or value < vtoi(mvarstrUserAuthority) Then
                mvarnAuthority = value
            End If
        End Set
        Get
            Return mvarnClientAuthority
        End Get
    End Property

    ' nAuthority 属性的设置器（私有）
    Public Property nAuthority As Integer
        Set(ByVal value As Integer)
            mvarnAuthority = value
            '/*Inform any connected processes
            RaiseEvent SecurityChange(value)
        End Set
        Get
            Return mvarnAuthority
        End Get
    End Property

    ' strUserAuthority 属性的设置器
    Public Property strUserAuthority As String
        Set(ByVal value As String)
            mvarstrUserAuthority = value
            '/*Test to see if this Authority level is less
            '/*than the Clients level
            'If IsNumeric(vData) And mvarbLoginRequired Then
            If IsNumeric(value) Then
                '/*Compare the two levels
                If vtoi(value) <= mvarnClientAuthority Then
                    nAuthority = vtoi(value)
                End If
            End If
        End Set
        Get
            Return mvarstrUserAuthority
        End Get
    End Property

    ' 类的初始化方法
    Public Sub New()
        '/*Preset the value of the authority so that we know if it has been touched
        mvarnAuthority = -1
        mvarstrUserAuthority = "-1"
    End Sub

    ' 辅助函数，将字符串转换为整数（模拟 VB6 的 vtoi 函数）
    Private Function vtoi(ByVal value As String) As Integer
        Dim result As Integer
        If Integer.TryParse(value, result) Then
            Return result
        End If
        Return 0
    End Function
End Class
