Public Class clsLocalApp
    ' 私有字段
    Private _strDesc As String         ' 应用程序唯一标题
    Private _strType As String         ' 本地应用类型（枚举值：CSB,SHELL,VS）
    Private _strIcon As String         ' 工具栏图标路径
    Private _strApp As String          ' 原始 SAX 代码或可执行文件路径
    Private _bToolBar As Boolean       ' 是否显示在浮动工具栏
    Private _strCSBIndex As String     ' CSB 集合的命名指针（索引）
    Private _strCSBStatus As String    ' CSB 的枚举状态

    ' 公共属性（带 Get/Set 访问器）
    Public Property strCSBStatus As String
        Get
            Return _strCSBStatus
        End Get
        Set(ByVal value As String)
            _strCSBStatus = value
        End Set
    End Property

    Public Property bToolBar As Boolean
        Get
            Return _bToolBar
        End Get
        Set(ByVal value As Boolean)
            _bToolBar = value
        End Set
    End Property

    Public Property strApp As String
        Get
            Return _strApp
        End Get
        Set(ByVal value As String)
            _strApp = value
        End Set
    End Property

    Public Property strCSBIndex As String
        Get
            Return _strCSBIndex
        End Get
        Set(ByVal value As String)
            _strCSBIndex = value
        End Set
    End Property

    Public Property strIcon As String
        Get
            Return _strIcon
        End Get
        Set(ByVal value As String)
            _strIcon = value
        End Set
    End Property

    Public Property strType As String
        Get
            Return _strType
        End Get
        Set(ByVal value As String)
            _strType = value
        End Set
    End Property

    Public Property strDesc As String
        Get
            Return _strDesc
        End Get
        Set(ByVal value As String)
            _strDesc = value
        End Set
    End Property

    ' 可选：添加默认构造函数
    Public Sub New()
        ' 初始化字段（可选）
        _strDesc = String.Empty
        _strType = String.Empty
        _strIcon = String.Empty
        _strApp = String.Empty
        _bToolBar = False
        _strCSBIndex = String.Empty
        _strCSBStatus = String.Empty
    End Sub
End Class
