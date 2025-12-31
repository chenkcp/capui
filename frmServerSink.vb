Imports System.Windows.Forms
Imports NextCapServer = ncapsrv


Public Class frmServerSink
    Inherits Form

    Public WithEvents tmrWaitShutdown As New Timer() ' 2 seconds
    Public WithEvents tmrHeartBeat As New Timer()    ' 10 seconds
    Public WithEvents tmrDBmaint As New Timer()      ' 5 ms


    Private WithEvents oServerMessage As NextCapServer.clsBoot
    Private WithEvents m_oSupervisor As NextCapServer.clsSupervisor

    Private m_strHeartBeatMinutes As String
    Private m_dtNextHeartBeat As DateTime
    Private m_nLockType As Integer
    Private m_bShutdown As Boolean
    Private m_nCountDown As Integer
    Private m_nClientAuthority As Integer
    Private m_dtWatchDogTime As DateTime

    Public ReadOnly Property ClientAuthority As Integer
        Get
            Return m_nClientAuthority
        End Get
    End Property

    Public Sub New()
        oServerMessage = go_BusinessServer '88888
        InitializeComponent()
    End Sub

    Private Sub InitializeComponent()
        Me.tmrDBmaint = New Timer() With {.Interval = 5}
        Me.tmrHeartBeat = New Timer() With {.Interval = 10000}
        Me.tmrWaitShutdown = New Timer() With {.Interval = 2000}

        Me.ClientSize = New Drawing.Size(462, 313)
        Me.Text = "Business Server Message Sink"
        Me.StartPosition = FormStartPosition.WindowsDefaultLocation

        AddHandler tmrDBmaint.Tick, AddressOf tmrDBmaint_Tick
        AddHandler tmrHeartBeat.Tick, AddressOf tmrHeartBeat_Tick
        AddHandler tmrWaitShutdown.Tick, AddressOf tmrWaitShutdown_Tick
    End Sub

    Private Sub frmServerSink_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' tmrWaitShutdown
        tmrWaitShutdown.Enabled = False
        tmrWaitShutdown.Interval = 2000 ' 2 seconds

        ' tmrHeartBeat
        tmrHeartBeat.Enabled = False
        tmrHeartBeat.Interval = 10000 ' 10 seconds

        ' tmrDBmaint
        tmrDBmaint.Enabled = False
        tmrDBmaint.Interval = 5 ' 5 milliseconds
        ' bind the clsBoot instance whose ReadyToGo we handle
        oServerMessage = go_BusinessServer
    End Sub

    Public WriteOnly Property ServerTimeOut As Integer
        Set(value As Integer)
            Dim nHeartBeatMinutes As Integer = If(value < 8, 1, value \ 8)
            Dim nHours As Integer = 0
            While nHeartBeatMinutes > 59
                nHours += 1
                nHeartBeatMinutes -= 60
            End While
            m_strHeartBeatMinutes = $"{nHours:D2}:{nHeartBeatMinutes:D2}:00"
            m_dtNextHeartBeat = DateTime.Now
            tmrHeartBeat.Enabled = True
        End Set
    End Property

    Private Sub NextHeartBeat()
        m_dtNextHeartBeat = DateTime.Now.Add(TimeSpan.Parse(m_strHeartBeatMinutes))
    End Sub

    Private Sub frmServerSink_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        Try
            m_oSupervisor?.ClientGone(PackContext())
        Catch ex As Exception
            LogToFile("Error.txt", $"frmServerSink.Form_Unload: {ex.Message}")
        End Try
    End Sub

    Private Sub oServerMessage_ReadyToGo(oSupervisor As NextCapServer.clsSupervisor, strClientType As String) Handles oServerMessage.ReadyToGo
        go_Supervisor = oSupervisor
        m_oSupervisor = oSupervisor

        Select Case strClientType.ToUpper()
            Case "MONITOR" : m_nClientAuthority = 0
            Case "EDITOR" : m_nClientAuthority = 2
            Case "AUDITOR" : m_nClientAuthority = 6
            Case "NORESTRICTION" : m_nClientAuthority = 8
        End Select

        go_Security.nClientAuthority = m_nClientAuthority
    End Sub

    Private Sub oServerMessage_Problem(strRef As String) Handles oServerMessage.Problem
        frmMessage.GenerateMessage("Business Server Error: " & strRef)
    End Sub

    Private Sub m_oSupervisor_ShuttingDown() Handles m_oSupervisor.ShuttingDown
        tmrWaitShutdown.Enabled = True
    End Sub

    Private Sub m_oSupervisor_LocalDBOffLine() Handles m_oSupervisor.LocalDBOffLine
        m_dtWatchDogTime = Now.AddMinutes(10)
        tmrDBmaint.Enabled = True
        frmMessage.LockedMessage("Supervisor #103 Client is locked for database maintenance", "Maintenance")
    End Sub

    Private Sub m_oSupervisor_LocalDBOnLine() Handles m_oSupervisor.LocalDBOnLine
        Try
            tmrDBmaint.Enabled = False
            frmMessage.UnlockMessage()
            ' gb_DBLockRequest = False
        Catch ex As Exception
            'gb_DBLockRequest = False
            frmMessage.Hide()
            Gb_InputBusyFlag = False
            mdlMain.frmNextCapInstance.Enabled = True
        End Try
    End Sub

    Private Sub tmrHeartBeat_Tick(sender As Object, e As EventArgs)
        If m_dtNextHeartBeat < Now AndAlso Not gb_ServerCallActive Then
            Try
                gb_ServerCallActive = True
                go_Supervisor.StayAwake()
                NextHeartBeat()
                Debug.WriteLine("frmServerSink heartbeat up")
            Finally
                gb_ServerCallActive = False
            End Try
        End If
    End Sub

    Private Sub tmrDBmaint_Tick(sender As Object, e As EventArgs)
        If Now > m_dtWatchDogTime Then
            m_oSupervisor_LocalDBOnLine()
        End If
    End Sub

    Private Sub tmrWaitShutdown_Tick(sender As Object, e As EventArgs)
        If Not m_bShutdown Then
            m_bShutdown = True
            m_nCountDown = 30
            frmMessage.GenerateMessage("Supervisor #100 Immediate shutdown taking place in 30 seconds.", "Shutdown", True)
        ElseIf m_nCountDown > 0 Then
            m_nCountDown -= 2
            If frmMessage.Visible Then
                frmMessage.lblMessage.Text = $"Supervisor #100 Immediate shutdown taking place in {m_nCountDown} seconds."
            Else
                frmMessage.GenerateMessage($"Supervisor #100 Immediate shutdown taking place in {m_nCountDown} seconds.", "Shutdown", True)
            End If
        Else
            tmrHeartBeat.Enabled = False
            tmrWaitShutdown.Enabled = False
            mdlWindow.RemoveForm(frmMessage.strFrmId)
            mdlMain.ShutDown(AppWinStyle.NormalFocus)
        End If
    End Sub

    Public Sub ShutdownServer()
        Try
            m_oSupervisor.ShutDown()
            frmMessage.GenerateMessage("Shutdown requested.", "Shutdown", True)
        Catch ex As Exception
            frmMessage.GenerateMessage("Shutdown Failed Error: " & ex.Message, "Shutdown", True)
        End Try
    End Sub

    Public Sub InitiateUpload()
        Try
            m_oSupervisor?.UploadDatabase()
        Catch ex As Exception
            frmMessage.GenerateMessage("Upload Failed Error: " & ex.Message, , True)
        End Try
    End Sub

    Public Sub UpdateServerContext()
        Try
            m_oSupervisor?.ContextChange(PackContext())
        Catch ex As Exception
            LogToFile("Error.txt", $"frmServerSink.UpdateServerContext: {ex.Message}")
        End Try
    End Sub

    Public Function Regression(ByVal bRegression As Boolean) As Boolean
        Return True
    End Function

    Public Function PackContext() As Object
        Dim vrtContext(12) As Object
        Try
            vrtContext(0) = go_clsSystemSettings.strUNC_Name
            With go_Context
                vrtContext(1) = .Operator
                vrtContext(2) = .Shift
                vrtContext(3) = .RunType
                vrtContext(4) = .ExperimentId
                vrtContext(5) = .PartNumber
                vrtContext(6) = .PartName
                vrtContext(7) = .ThinFilmLot
                vrtContext(8) = .LineType
                vrtContext(9) = .LineNumber
                vrtContext(10) = .Source
                vrtContext(11) = .Accumulator
                vrtContext(12) = .ProductionDate
            End With
        Catch ex As Exception
            LogToFile("Error.txt", $"frmServerSink.PackContext: {ex.Message}")
        End Try
        Return vrtContext
    End Function
End Class

