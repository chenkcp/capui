
'Imports System.Runtime.Remoting.Contexts
Imports ncapsrv
Imports NextcapServer = ncapsrv
Module mdlGlobal
    '/*Global word to inentify the Continuous mode of operation
    Public Const gcstrContinuous As String = "Continuous"
    Public Const gcstrLot As String = "Lot"

    Private mb_InputBusyFlag As Boolean
    Public go_clsSystemSettings As clsSystemSettings
    Public gcol_ChartData As colChartData
    '/*Business Server Objects
    Public WithEvents go_BusinessServer As NextcapServer.clsBoot
    Public go_Supervisor As NextcapServer.clsSupervisor
    Public WithEvents go_ActiveLotManager As NextcapServer.clsLotManager
    '/*Collection of LotManagers
    Public gcol_LotManagers As List(Of frmLotManagerSink)
    '/*Context object
    Public go_Context As Context
    '/*Global suystem copy of the current Security level
    Public go_Security As Security
    Public gcol_QualityStatus As colQualityStatuses
    Public gcol_StationKeys As colStationKeys
    Public go_ActiveLotSink As frmLotManagerSink
    Public gcol_LocalApps As colLocalApps
    Public gn_ImageCheckCount As Integer
    Public gcol_SAXhandlers As colSAXhandlers
    Public colProtos As New colProtoTypes()
    Public go_ChartConfigs As clsChartConfig
    Public go_SampleCount As clsSampleCount
    Public gcol_LotSequence As colLotSequences
    Public gb_InputRequestFlag As Boolean
    Public go_CAPmain As CAPmain
    Public Const AppWindows As Integer = 2
    Public gb_ServerCallActive As Boolean
    Public go_SharedIdInput As clsUnitId
    Public go_AutoCreateLot As Boolean

    Private Sub OnReadyToGo(ByVal supervisor As clsSupervisor, ByVal strClientType As String) Handles go_BusinessServer.ReadyToGo
        go_Supervisor = supervisor
    End Sub

    ' Converts an object to a string, handling null values safely
    Public Function vtoa(ByVal value As Object) As String
        If value Is Nothing OrElse IsDBNull(value) Then
            Return String.Empty
        End If
        Return value.ToString().Trim()
    End Function

    ' Converts an object to an integer, handling null or invalid values
    Public Function vtoi(ByVal value As Object) As Integer
        If value Is Nothing OrElse IsDBNull(value) OrElse Not IsNumeric(value) Then
            Return 0 ' Default value if conversion fails
        End If
        Return Convert.ToInt32(value)
    End Function

    ' Converts an object to a Boolean, handling null and invalid values
    Public Function vtob(ByVal value As Object) As Boolean
        If value Is Nothing OrElse IsDBNull(value) Then
            Return False ' Default to False if NULL
        End If
        Dim strValue As String = value.ToString().Trim().ToLower()
        Return (strValue = "true" Or strValue = "1" Or strValue = "yes")
    End Function

    ' Converts an object to a DateTime, handling null values safely
    Public Function vtoDate(ByVal value As Object) As DateTime
        If value Is Nothing OrElse IsDBNull(value) OrElse Not IsDate(value) Then
            Return DateTime.MinValue ' Default value if invalid
        End If
        Return Convert.ToDateTime(value)
    End Function

    Public Property gb_InputBusyFlag As Boolean
        Get
            Return mb_InputBusyFlag
        End Get
        Set(value As Boolean)
            mb_InputBusyFlag = value
            ' Pass this on to the busy indicator on the Client
            'Form1.Busy(value)
        End Set
    End Property

    'Private Sub go_ActiveLotManager_UpdateDefectCount(ByVal strLotId As String, ByVal dateBirthday As DateTime, ByVal strClassName As String, ByVal intNewCount As Integer) Handles go_ActiveLotManager.UpdateDefectCount
    '    ' Handle the event
    '    Debug.WriteLine($"Backend Lot Manager events - LotId: {strLotId}, Birthday: {dateBirthday:d}, Class: {strClassName}, New Count: {intNewCount}")
    'End Sub

End Module
