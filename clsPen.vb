Public Class clsPen

    ' Local variables to hold property values
    Private mvarstrPenId As String
    Private mvardtInspectionDate As Date
    Private mvarstrOperator As Object ' Variant replaced with Object
    Private mvarstrShift As String
    Private mvarstrDisposition As String
    Private mvarstrTestBed As String
    Private mvarbPenNotShipped As Boolean
    Private mvarstrRecoveryStep As String
    Private mvarnCount As Integer
    Private mvarstrRunType As String
    Private mvarstrExperimentId As String
    Private mvarstrPartNumber As String
    Private mvarstrPartName As String
    Private mvarstrThinFilmLotId As String

    Private mvarcolPenDefects As New colDefect()
    Private mvarstrSource As String
    Private mvarstrLineType As String
    Private mvarstrLineId As String
    Private mvarstrAccumulator As String
    Private mvardtBirthDay As Date
    Private mvarstrGroupId As String
    Private mvarstrUnitType As String

    ' Properties
    Public Property strUnitType As String
        Get
            Return mvarstrUnitType
        End Get
        Set(value As String)
            mvarstrUnitType = value
        End Set
    End Property

    Public Property strGroupId As String
        Get
            Return mvarstrGroupId
        End Get
        Set(value As String)
            mvarstrGroupId = value
        End Set
    End Property

    Public Property dtBirthDay As Date
        Get
            Return mvardtBirthDay
        End Get
        Set(value As Date)
            mvardtBirthDay = value
        End Set
    End Property

    Public Property strAccumulator As String
        Get
            Return mvarstrAccumulator
        End Get
        Set(value As String)
            mvarstrAccumulator = value
        End Set
    End Property

    Public Property strLineId As String
        Get
            Return mvarstrLineId
        End Get
        Set(value As String)
            mvarstrLineId = value
        End Set
    End Property

    Public Property strLineType As String
        Get
            Return mvarstrLineType
        End Get
        Set(value As String)
            mvarstrLineType = value
        End Set
    End Property

    Public Property strSource As String
        Get
            Return mvarstrSource
        End Get
        Set(value As String)
            mvarstrSource = value
        End Set
    End Property

    Public Property colPenDefects As colDefect
        Get
            Return mvarcolPenDefects
        End Get
        Set(value As colDefect)
            mvarcolPenDefects = value
        End Set
    End Property

    Public Property strThinFilmLotId As String
        Get
            Return mvarstrThinFilmLotId
        End Get
        Set(value As String)
            mvarstrThinFilmLotId = value
        End Set
    End Property

    Public Property strPartName As String
        Get
            Return mvarstrPartName
        End Get
        Set(value As String)
            mvarstrPartName = value
        End Set
    End Property

    Public Property strPartNumber As String
        Get
            Return mvarstrPartNumber
        End Get
        Set(value As String)
            mvarstrPartNumber = value
        End Set
    End Property

    Public Property strExperimentId As String
        Get
            Return mvarstrExperimentId
        End Get
        Set(value As String)
            mvarstrExperimentId = value
        End Set
    End Property

    Public Property strRunType As String
        Get
            Return mvarstrRunType
        End Get
        Set(value As String)
            mvarstrRunType = value
        End Set
    End Property

    Public Property nCount As Integer
        Get
            Return mvarnCount
        End Get
        Set(value As Integer)
            mvarnCount = value
        End Set
    End Property

    Public Property strRecoveryStep As String
        Get
            Return mvarstrRecoveryStep
        End Get
        Set(value As String)
            mvarstrRecoveryStep = value
        End Set
    End Property

    Public Property bPenNotShipped As Boolean
        Get
            Return mvarbPenNotShipped
        End Get
        Set(value As Boolean)
            mvarbPenNotShipped = value
        End Set
    End Property

    Public Property strTestBed As String
        Get
            Return mvarstrTestBed
        End Get
        Set(value As String)
            mvarstrTestBed = value
        End Set
    End Property

    Public Property strDisposition As String
        Get
            Return mvarstrDisposition
        End Get
        Set(value As String)
            mvarstrDisposition = value
        End Set
    End Property

    Public Property strShift As String
        Get
            Return mvarstrShift
        End Get
        Set(value As String)
            mvarstrShift = value
        End Set
    End Property

    Public Property strOperator As Object
        Get
            Return mvarstrOperator
        End Get
        Set(value As Object)
            mvarstrOperator = value
        End Set
    End Property

    Public Property dtInspectionDate As Date
        Get
            Return mvardtInspectionDate
        End Get
        Set(value As Date)
            mvardtInspectionDate = value
        End Set
    End Property

    Public Property strPenId As String
        Get
            Return mvarstrPenId
        End Get
        Set(value As String)
            mvarstrPenId = value
        End Set
    End Property

    ' Constructor
    Public Sub New()
        ' Initialize new colPenDefects
        mvarcolPenDefects = New colDefect()
    End Sub

End Class

