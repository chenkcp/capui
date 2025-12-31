Public Class clsWindow
    'local variable(s) to hold property value(s)
    Private mvarfrmIn As Form 'local copy
    Private mvarbEnabled As Boolean 'local copy
    Private mvarbVisible As Boolean 'local copy
    Private mvarbFloat As Boolean 'local copy

    'bFloat 属性的 Get 方法
    Public Property bFloat As Boolean
        Get
            Return mvarbFloat
        End Get
        Set(ByVal value As Boolean)
            mvarbFloat = value
        End Set
    End Property

    'bVisible 属性的 Get 方法
    Public Property bVisible As Boolean
        Get
            Return mvarbVisible
        End Get
        Set(ByVal value As Boolean)
            mvarbVisible = value
        End Set
    End Property

    'bEnabled 属性的 Get 方法
    Public Property bEnabled As Boolean
        Get
            Return mvarbEnabled
        End Get
        Set(ByVal value As Boolean)
            mvarbEnabled = value
        End Set
    End Property

    'frmIn 属性的 Get 方法
    Public Property frmIn As Form
        Get
            Return mvarfrmIn
        End Get
        Set(ByVal value As Form)
            mvarfrmIn = value
        End Set
    End Property
End Class
