Public Class clsDefect
    ' Local variables to hold property values
    Private mvarlngSeverity As Long
    Private mvarbPrimary As Boolean
    Private mvarnFeatureClass As Integer
    Private mvarnFeatureIndex As Integer
    Private mvarnSubFeatureIndex As Integer
    Private mvarstrFeatureCd As String
    Private mvarstrFeatureDesc As String
    Private mvarstrSubFeatureCd As String
    Private mvarstrSubFeatureDesc As String
    Private mvarstrCauseCd As String
    Private mvarstrCauseDesc As String
    Private mvarstrSubCauseCd As String
    Private mvarstrSubCauseDesc As String
    Private mvarstrFeatureClassCd As String
    Private mvarstrFeatureClassDesc As String
    Private mvarstrComment As String
    Private mvarstrNumericInput As String
    Private mvarnEnteredOrder As Integer

    ' Properties
    Public Property nEnteredOrder As Integer
        Get
            Return mvarnEnteredOrder
        End Get
        Set(value As Integer)
            mvarnEnteredOrder = value
        End Set
    End Property

    Public Property strNumericInput As String
        Get
            Return mvarstrNumericInput
        End Get
        Set(value As String)
            mvarstrNumericInput = value
        End Set
    End Property

    Public Property strComment As String
        Get
            Return mvarstrComment
        End Get
        Set(value As String)
            mvarstrComment = value
        End Set
    End Property

    Public Property strFeatureClassDesc As String
        Get
            Return mvarstrFeatureClassDesc
        End Get
        Set(value As String)
            mvarstrFeatureClassDesc = value
        End Set
    End Property

    Public Property strFeatureClassCd As String
        Get
            Return mvarstrFeatureClassCd
        End Get
        Set(value As String)
            mvarstrFeatureClassCd = value
        End Set
    End Property

    Public Property strSubCauseDesc As String
        Get
            Return mvarstrSubCauseDesc
        End Get
        Set(value As String)
            mvarstrSubCauseDesc = value
        End Set
    End Property

    Public Property strSubCauseCd As String
        Get
            Return mvarstrSubCauseCd
        End Get
        Set(value As String)
            mvarstrSubCauseCd = value
        End Set
    End Property

    Public Property strCauseDesc As String
        Get
            Return mvarstrCauseDesc
        End Get
        Set(value As String)
            mvarstrCauseDesc = value
        End Set
    End Property

    Public Property strCauseCd As String
        Get
            Return mvarstrCauseCd
        End Get
        Set(value As String)
            mvarstrCauseCd = value
        End Set
    End Property

    Public Property strSubFeatureDesc As String
        Get
            Return mvarstrSubFeatureDesc
        End Get
        Set(value As String)
            mvarstrSubFeatureDesc = value
        End Set
    End Property

    Public Property strSubFeatureCd As String
        Get
            Return mvarstrSubFeatureCd
        End Get
        Set(value As String)
            mvarstrSubFeatureCd = value
        End Set
    End Property

    Public Property strFeatureDesc As String
        Get
            Return mvarstrFeatureDesc
        End Get
        Set(value As String)
            mvarstrFeatureDesc = value
        End Set
    End Property

    Public Property strFeatureCd As String
        Get
            Return mvarstrFeatureCd
        End Get
        Set(value As String)
            mvarstrFeatureCd = value
        End Set
    End Property

    Public Property nSubFeatureIndex As Integer
        Get
            Return mvarnSubFeatureIndex
        End Get
        Set(value As Integer)
            mvarnSubFeatureIndex = value
        End Set
    End Property

    Public Property nFeatureIndex As Integer
        Get
            Return mvarnFeatureIndex
        End Get
        Set(value As Integer)
            mvarnFeatureIndex = value
        End Set
    End Property

    Public Property nFeatureClass As Integer
        Get
            Return mvarnFeatureClass
        End Get
        Set(value As Integer)
            mvarnFeatureClass = value
        End Set
    End Property

    Public Property bPrimary As Boolean
        Get
            Return mvarbPrimary
        End Get
        Set(value As Boolean)
            mvarbPrimary = value
        End Set
    End Property

    Public Property lngSeverity As Long
        Get
            Return mvarlngSeverity
        End Get
        Set(value As Long)
            mvarlngSeverity = value
        End Set
    End Property
End Class

