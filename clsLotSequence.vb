Imports NextcapServer = ncapsrv

Public Class clsLotSequence
    ' String for Quality Status form completed
    Public Property QualityDone As String
    ' String for NewLot form completed
    Public Property CreateLotDone As String
    ' Time at which to expire this request
    Public Property ExpirationDttm As DateTime
    ' Quality status from server
    Public Property QualityResult As String
    ' Result of EndOfLot form: "" (pending), "CLOSE", "CONTINUE"
    Public Property EndOfLotResult As String
    ' The Group's ID
    Public Property GroupId As String
    ' The Group's Birthday
    Public Property Birth As DateTime
    ' Reference to the requestor LotManager
    Public Property LotManager As NextCapServer.clsLotManager
    ' Reference to the ChartData collection on the Sink
    Public Property ChartData As colChartData
    ' The LotManager requesting this call
    Public Property RequestingManager As frmLotManagerSink
    ' Unique index for this sequence
    Public Property Index As String
    ' Value indicating the phase of this request (TimeToClose=4, Quality=2, TimeToCreate=1)
    Public Property Phase As Integer
    ' 0/1 need Close
    Public Property RequestClose As Integer
    ' 0/1 need quality window
    Public Property RequestQuality As Integer
    ' 0/1 need create window
    Public Property RequestCreate As Integer

    ' Constructor to set default expiration (10 minutes from now)
    'Public Sub New()
    '    ExpirationDttm = DateTime.Now.AddMinutes(10)
    'End Sub

    ' Fields
    Private m_strQualityDone As String
    Private m_strCreateLotDone As String
    Private m_dtExpirationDttm As Date
    Private m_strQualityResult As String
    Private m_strEndOfLotResult As String
    Private m_strGroupId As String
    Private m_dtBirth As Date
    Private m_oLotManager As NextcapServer.clsLotManager
    Private m_ocolChartData As colChartData
    Private m_frmRequestingManager As frmLotManagerSink
    Private m_strIndex As String
    Private m_nPhase As Integer
    Private m_nRequestClose As Integer
    Private m_nRequestQuality As Integer
    Private m_nRequestCreate As Integer

    ' Properties
    Public Property nRequestCreate As Integer
        Get
            Return m_nRequestCreate
        End Get
        Set(value As Integer)
            m_nRequestCreate = value
        End Set
    End Property

    Public Property nRequestQuality As Integer
        Get
            Return m_nRequestQuality
        End Get
        Set(value As Integer)
            m_nRequestQuality = value
        End Set
    End Property

    Public Property nRequestClose As Integer
        Get
            Return m_nRequestClose
        End Get
        Set(value As Integer)
            m_nRequestClose = value
        End Set
    End Property

    Public Property nPhase As Integer
        Get
            Return m_nPhase
        End Get
        Set(value As Integer)
            m_nPhase = value
        End Set
    End Property

    Public Property strIndex As String
        Get
            Return m_strIndex
        End Get
        Set(value As String)
            m_strIndex = value
        End Set
    End Property

    Public Property frmRequestingManager As frmLotManagerSink
        Get
            Return m_frmRequestingManager
        End Get
        Set(value As frmLotManagerSink)
            m_frmRequestingManager = value
        End Set
    End Property

    Public Property ocolChartData As colChartData
        Get
            Return m_ocolChartData
        End Get
        Set(value As colChartData)
            m_ocolChartData = value
        End Set
    End Property

    Public Property oLotManager As NextcapServer.clsLotManager
        Get
            Return m_oLotManager
        End Get
        Set(value As NextcapServer.clsLotManager)
            m_oLotManager = value
        End Set
    End Property

    Public Property dtBirth As Date
        Get
            Return m_dtBirth
        End Get
        Set(value As Date)
            m_dtBirth = value
        End Set
    End Property

    Public Property strGroupId As String
        Get
            Return m_strGroupId
        End Get
        Set(value As String)
            m_strGroupId = value
        End Set
    End Property

    Public Property strEndOfLotResult As String
        Get
            Return m_strEndOfLotResult
        End Get
        Set(value As String)
            m_strEndOfLotResult = value
        End Set
    End Property

    Public Property strQualityResult As String
        Get
            Return m_strQualityResult
        End Get
        Set(value As String)
            m_strQualityResult = value
        End Set
    End Property

    Public Property dtExpirationDttm As Date
        Get
            Return m_dtExpirationDttm
        End Get
        Set(value As Date)
            m_dtExpirationDttm = value
        End Set
    End Property

    Public Property strCreateLotDone As String
        Get
            Return m_strCreateLotDone
        End Get
        Set(value As String)
            m_strCreateLotDone = value
        End Set
    End Property

    Public Property strQualityDone As String
        Get
            Return m_strQualityDone
        End Get
        Set(value As String)
            m_strQualityDone = value
        End Set
    End Property

    ' Constructor
    Public Sub New()
        ' Preset values
        m_dtExpirationDttm = Now.AddMinutes(10)
    End Sub

    ' Destructor (Finalizer)
    Protected Overrides Sub Finalize()
        Try
            m_frmRequestingManager = Nothing
        Finally
            MyBase.Finalize()
        End Try
    End Sub
End Class
