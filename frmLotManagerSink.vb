Imports System
Imports System.Collections.Generic
Imports System.DirectoryServices.ActiveDirectory
Imports System.Globalization
Imports System.Windows.Forms
Imports NextcapServer = ncapsrv

Public Class frmLotManagerSink
    Inherits System.Windows.Forms.Form

    ''COM sink for LotManager update events
    Private WithEvents m_oLotManager As NextcapServer.clsLotManager
    Private WithEvents m_oNextCap As frmNextCap
    ''Storage for the ChartData associated with this LotManager instance
    Private m_colChartData As colChartData
    ''Storage for the current state
    'Public m_clsWindow As clsWindowState
    'Private m_frmParent As Form
    'Storage for Lot BirthDates. This is synchronized with the visible list names at an offset of +1
    Private m_colClosed As New List(Of String)
    Private m_colClosedDate As New List(Of Date)
    Private m_strOpenLot As String
    Private m_dtOpen As Date
    Private m_colSuspended As New List(Of String)
    Private m_colSuspendedDate As New List(Of Date)
    ''Constant to define server error command
    'Private Const cServerError As String = "ERROR"
    'Counter object to track the number of Groups assigned today.
    'This is utilized by default counter for generating ids
    Public oLotsToday As clsLotCountToday
    Private m_nAuthority As Integer = 6

    Public Sub New()
        Me.Text = "frmLotManagersink"
    End Sub
    ''Storage for Lot to close and its associated date
    'Private m_strAsyncLotId As String
    'Private m_dtAsyncBirth As Date
    ''Collection to track requests for Close Lot Events
    'Private m_colTrackClosedId As New List(Of String)
    'Private m_colTrackClosedDate As New List(Of Date)
    'Private Const MAX_TRACK As Integer = 20
    ''Reference to the global security object
    'Private WithEvents m_oSecurity As Security
    'Private m_nAuthority As Integer
    ''Constant for the wild card used in CycleEventTask
    'Private Const m_cstrWildCard As String = "ALL"

    ''Property to allow for reading the open lot in this manager
    'Public Property OpenLot As String
    '    Get
    '        Return m_strOpenLot
    '    End Get
    '    Set(ByVal value As String)
    '        m_strOpenLot = value
    '    End Set
    'End Property

    'Routine: frmLotManager.CreateLotManager()
    'Purpose: This creates the LotManager from the business server if possible and then runs the required setup routines such as RecvLots.
    Public Sub SetLotManager(ByVal oLotManager As NextcapServer.clsLotManager)
        'Connect the COM sink for the B:server
        If oLotManager IsNot Nothing Then
            m_oLotManager = oLotManager
            If Not (m_colChartData Is Nothing) Then m_colChartData = Nothing 'PPPP

            'Generate a new ChartData collection
            m_colChartData = New colChartData
            'Generate the LotToday() object for this Sink
            oLotsToday = New clsLotCountToday
            If m_oLotManager.LotIdCount > 0 Then 'PPPP
                'Query the Business server for the Lot Manager information
                RecvOpenLot(m_oLotManager)
                RecvClosedLot(m_oLotManager)
                RecvSuspendedLot(m_oLotManager)
            End If
            ' Ensure m_colChartData is initialized
            If m_colChartData Is Nothing OrElse m_colChartData.Count = 0 Then
                Debug.WriteLine("m_colChartData is empty or uninitialized.")
                Exit Sub
            End If

            ' Iterate through the collection and print each item
            For Each chartData As clsChartData In m_colChartData
                Debug.WriteLine($"GroupId: {chartData.strGroupId}, BirthDate: {chartData.dtBirth}, GoodCount: {chartData.nGoodCount}, BadCount: {chartData.nBadCount}, IconPath: {chartData.strIconPath}, IconName: {chartData.strIconName}")
            Next
            'Initialize the data set of this LotManager
            'oLotManager.InitializeClient() 'PPPP
        End If
    End Sub

    ''Routine: frmLotManagerSink.InitializeData()
    ''Purpose: This is an external interface for generating the initial set of data belonging to this LotManager.
    'Public Sub InitializeData()
    '    Dim vrtArrayCount As Object()
    '    Dim vrtArrayStatus As Object()
    '    Dim strData As String
    '    Dim strMsg As String
    '    Dim lngItem As Long
    '    Dim lngVrtItem As Long
    '    Dim strCurrentLot As String

    '    'On Error Resume Next
    '    'm_oLotManager.InitializeClient
    '    vrtArrayCount = GetAllData()
    '    vrtArrayStatus = GetAllQualityStatus()

    '    'Make sure we did not receive an error message
    '    If vrtArrayCount(0).ToString() <> cServerError Then
    '        'Load up the lot data
    '        For lngVrtItem = 0 To UBound(vrtArrayCount, 1)
    '            'Put in the counts
    '            m_oLotManager_UpdateDefectCount(Convert.ToString(vrtArrayCount(0, lngVrtItem)), Convert.ToDate(vrtArrayCount(1, lngVrtItem)), Convert.ToString(vrtArrayCount(0, lngVrtItem)), Convert.ToInt32(vrtArrayCount(3, lngVrtItem)))

    '            'Watch for changes in the name of the current Lot counts
    '            If Convert.ToString(vrtArrayCount(0, lngVrtItem)) <> strCurrentLot Then
    '                'Set our trigger to the new lot
    '                strCurrentLot = Convert.ToString(vrtArrayCount(0, lngVrtItem))
    '                'linear search for the lot id and load its status
    '                For lngItem = 0 To UBound(vrtArrayStatus, 1)
    '                    If Convert.ToString(vrtArrayStatus(0, lngItem)) = strCurrentLot Then
    '                        'Put in the status
    '                        m_oLotManager_UpdateMaterialStatus(Convert.ToString(vrtArrayStatus(0, lngItem)), Convert.ToDate(vrtArrayStatus(1, lngItem)), Convert.ToString(vrtArrayStatus(0, lngItem)))
    '                    End If
    '                End If
    '            End If
    '    Next lngVrtItem
    '    End If
    'End Sub

    ''Routine: frmLotManagerSink.GetAllData()
    ''Purpose: This is an external interface for generating the initial set of data belonging to this LotManager that contains the Defect count and type portions.
    'Public Function GetAllData() As Object()
    '    Dim vrtArray As Object()
    '    Dim strData As String
    '    Dim strMsg As String
    '    Dim lngItem As Long
    '    Dim lngVrtItem As Long

    '    Try
    '        'Insure this is set
    '        If m_oLotManager IsNot Nothing Then
    '            'Loop through the station keys
    '            vrtArray = m_oLotManager.GetAllDefectCounts()

    '            'An array of dimension (0,0) indicates an error occured.
    '            If vrtArray(0, 0).ToString() = cServerError Then
    '                'Route the error through the centralized handler
    '                mdlMain.ReportServerError("frmLotManagerSink.GetAllData", vrtArray)
    '            Else
    '                'Loop through each item
    '                'For lngVrtItem = 0 To UBound(vrtArray, 2)
    '                '    m_oLotManager_UpdateDefectCount vtoa(vrtArray(0, lngVrtItem)), _
    '                '                                            vtoDate(vrtArray(1, lngVrtItem)), _
    '                '                                            vtoa(vrtArray(2, lngVrtItem)), _
    '                '                                            vtoi(vrtArray(3, lngVrtItem))
    '                'Next lngVrtItem
    '            End If
    '        End If
    '    Catch ex As Exception
    '        MainErrorHandler("frmLotManagerSink.GetAllData", ex.Message & "-" & ex.HResult)
    '        LogToFile("Error.txt", "frmLotManagerSink.GetAllData:" & ex.HResult & " - " & ex.Message)
    '        Err.Clear()
    '    End Try

    '    Return vrtArray
    'End Function

    ''Routine: frmLotManagerSink.GetAllQualityStatus()
    ''Purpose: This is an external interface for generating the initial set of data belonging to this LotManager that contains the Quality Status.
    'Public Function GetAllQualityStatus() As Object()
    '    Dim vrtArray As Object()
    '    Dim strData As String
    '    Dim strMsg As String
    '    Dim lngItem As Long
    '    Dim lngVrtItem As Long

    '    Try
    '        'Insure this is set
    '        If m_oLotManager IsNot Nothing Then
    '            'Loop through the station keys
    '            vrtArray = m_oLotManager.GetAllQualityStatuses()

    '            'An array of dimension (0,0) indicates an error occured.
    '            If vrtArray(0, 0).ToString() = cServerError Then
    '                'Route the error through the centralized handler
    '                mdlMain.ReportServerError("frmLotManagerSink.GetAllQualityStatus", vrtArray)
    '            Else
    '                'Loop through each item
    '                'For lngVrtItem = 0 To UBound(vrtArray, 2)
    '                '    m_oLotManager_UpdateMaterialStatus vtoa(vrtArray(0, lngVrtItem)), _
    '                '                                            vtoDate(vrtArray(1, lngVrtItem)), _
    '                '                                            vtoa(vrtArray(2, lngVrtItem))
    '                'Next lngVrtItem
    '            End If
    '        End If
    '    Catch ex As Exception
    '        MainErrorHandler("frmLotManagerSink.GetAllQualityStatus", ex.Message & "-" & ex.HResult)
    '        LogToFile("Error.txt", "frmLotManagerSink.GetAllQualityStatus:" & ex.HResult & " - " & ex.Message)
    '        Err.Clear()
    '    End Try

    '    Return vrtArray
    'End Function

    ''Standard form Load event
    'Private Sub Form_Load()
    '    'Disable the QueryUnload X
    '    mdlTools.DisableCloseX(Me)
    '    'Initialize the private tracking collections for the first use
    '    m_colSuspended = New List(Of String)
    '    m_colSuspendedDate = New List(Of Date)
    '    m_colClosed = New List(Of String)
    '    m_colClosedDate = New List(Of Date)
    '    'Generate the tracking collection
    '    m_colTrackClosedId = New List(Of String)
    '    m_colTrackClosedDate = New List(Of Date)
    '    'Attach to the global security object; set the current Authority
    '    m_oSecurity = go_Security
    '    m_nAuthority = m_oSecurity.nAuthority
    'End Sub

    ''Routine: frmLotManager.LotIsEmpty(n)
    ''Purpose: Use the empty lot collection to look up if the questioned lot is empty.
    'Private Function LotIsEmpty(ByRef strGroupId As String, ByRef dtBirth As Date) As Boolean
    '    'Test if the Lot is available for Delete
    '    If m_oLotManager IsNot Nothing Then
    '        'LotIsEmpty = m_oLotManager.islotempty(strGroupId, dtBirth)
    '    End If
    '    Return False
    'End Function

    ''Centralized server error
    'Private Sub ReportServerError(ByRef strCaller As String, ByRef strMsg As String)
    '    MessageBox.Show(strCaller & "-" & strMsg, "Error")
    'End Sub

    'Routine: frmLotManager.RecvOpenLot(o)
    'Purpose: This queries the Lot Manager to get a list of the current Open lots.
    Private Sub RecvOpenLot(ByVal oLotManager As NextcapServer.clsLotManager)
        Dim vrtArray As DataTable = Nothing

        If mdlLotManager.RecvOpenLot(oLotManager, vrtArray) Then

            m_strOpenLot = vtoa(vrtArray.Rows(0)(0)) 'tttt
            m_dtOpen = vrtArray.Rows(0)(1) 'open lot is local not from database

            'If Not String.IsNullOrEmpty(m_strOpenLot) Then 'tttt
            If Not String.IsNullOrEmpty(m_strOpenLot) AndAlso Not String.Equals(m_strOpenLot, "There is currently no open lot") Then
                oLotsToday.AddLot(m_dtOpen)
            End If
        End If

        'Dim vrtArray As Object(,)  ' 声明为二维数组  'PPPP

        ''request the data
        'mdlLotManager.RecvOpenLot(oLotManager, vrtArray) 'PPPP

        ''Set the new items into the list
        'If vrtArray IsNot Nothing AndAlso vrtArray.Rank = 2 Then 'PPPP
        '    m_strOpenLot = If(vrtArray(0, 0) IsNot Nothing, vrtArray(0, 0).ToString(), "") 'PPPP

        '    'Now get the birthday
        '    If vrtArray.GetLength(0) > 1 AndAlso vrtArray(1, 0) IsNot Nothing Then'PPPP
        '        ' 使用 TryParse 避免转换异常
        '        If DateTime.TryParse(vrtArray(1, 0).ToString(), m_dtOpen) Then'PPPP
        '            'Add this to the Lot Counter object'PPPP
        '            If Not String.IsNullOrEmpty(m_strOpenLot) Then'PPPP
        '                oLotsToday.AddLot(m_dtOpen)'PPPP
        '            End If'PPPP
        '        End If'PPPP
        '    End If'PPPP
        'End If'PPPP
    End Sub

    'Routine: frmLotManager.RecvClosedLot(o)
    'Purpose: This queries the Lot Manager to get a list of the current Closed lots.
    Private Sub RecvClosedLot(oLotManager As NextcapServer.clsLotManager)
        Dim vrtArray As DataTable = Nothing

        If mdlLotManager.RecvClosedLot(oLotManager, vrtArray) Then
            m_colClosed = New List(Of String)
            m_colClosedDate = New List(Of Date)

            If vrtArray IsNot Nothing AndAlso vrtArray.Rows.Count > 0 Then
                For lngVrtItem As Integer = 0 To vrtArray.Rows.Count - 1
                    Dim closedLot As String = vtoa(vrtArray.Rows(lngVrtItem)(0))

                    ' Normalize the DateTime value to avoid ambiguity
                    Dim rawDate As DateTime = CType(vrtArray.Rows(lngVrtItem)(1), DateTime)
                    Dim normalizedDate As DateTime = DateTime.ParseExact(
                    rawDate.ToString("MM/dd/yyyy HH:mm:ss", CultureInfo.InvariantCulture),
                    "MM/dd/yyyy HH:mm:ss",
                    CultureInfo.InvariantCulture
                )

                    m_colClosed.Add(closedLot)
                    m_colClosedDate.Add(normalizedDate)
                    oLotsToday.AddLot(normalizedDate)
                Next
            End If
        End If
    End Sub

    'Routine: frmLotManager.RecvSuspendedLot(o)
    'Purpose: This queries the Lot Manager to get a list of the current Suspended lots.
    Private Sub RecvSuspendedLot(oLotManager As NextcapServer.clsLotManager)
        Dim vrtArray As DataTable = Nothing

        If mdlLotManager.RecvSuspendedLot(oLotManager, vrtArray) Then
            m_colSuspended = New List(Of String)
            m_colSuspendedDate = New List(Of Date)

            If vrtArray IsNot Nothing AndAlso vrtArray.Rows.Count > 0 Then
                For i As Integer = 0 To vrtArray.Rows.Count - 1
                    ' Read column 0 (string) and column 1 (date) from each row
                    Dim suspendedLot As String = vtoa(vrtArray.Rows(i)(0))

                    ' Normalize the DateTime value to avoid ambiguity
                    Dim rawDate As DateTime = CType(vrtArray.Rows(i)(1), DateTime)
                    Dim normalizedDate As DateTime = DateTime.ParseExact(
                    rawDate.ToString("MM/dd/yyyy HH:mm:ss", CultureInfo.InvariantCulture),
                    "MM/dd/yyyy HH:mm:ss",
                    CultureInfo.InvariantCulture
                )

                    m_colSuspended.Add(suspendedLot)
                    m_colSuspendedDate.Add(normalizedDate)
                    oLotsToday.AddLot(normalizedDate)
                Next
            End If
        End If
    End Sub


    Private Sub m_oLotManager_LotOpened(ByVal strGroupId As String, ByVal dtBirth As DateTime)
        Dim nListIndex As Integer

        Try
            ' Normalize the DateTime value to avoid ambiguity
            Dim normalizedDate As DateTime = DateTime.ParseExact(
            dtBirth.ToString("MM/dd/yyyy HH:mm:ss", CultureInfo.InvariantCulture),
            "MM/dd/yyyy HH:mm:ss",
            CultureInfo.InvariantCulture
        )

            ' Scan the closed list
            For nListIndex = 0 To m_colClosed.Count - 1
                ' Find a matching name first
                If m_colClosed(nListIndex) = strGroupId Then
                    ' Now test the Date collection
                    If m_colClosedDate(nListIndex) = normalizedDate Then
                        ' Synchronize to match the business server
                        m_colClosed.RemoveAt(nListIndex)
                        m_colClosedDate.RemoveAt(nListIndex)
                        ' Exit the loop
                        Exit For
                    End If
                End If
            Next nListIndex

            ' Scan the suspended list
            For nListIndex = 0 To m_colSuspended.Count - 1
                ' Find a matching name first
                If m_colSuspended(nListIndex) = strGroupId Then
                    ' Now test the Date collection
                    If m_colSuspendedDate(nListIndex) = normalizedDate Then
                        ' Synchronize to match the business server
                        m_colSuspended.RemoveAt(nListIndex)
                        m_colSuspendedDate.RemoveAt(nListIndex)
                        ' Exit the loop
                        Exit For
                    End If
                End If
            Next nListIndex

            ' If the Lot was not found in either of the special case selections
            ' then it must be a new Lot and we need to add it to the Lot counter
            oLotsToday.AddLot(normalizedDate)

            ' Add the item to the open lot
            m_strOpenLot = strGroupId
            m_dtOpen = normalizedDate

            ' -- Execute the LotPointer routine
            ' mdlChart.SetLotPointerIcon()
        Catch ex As Exception
            LogToFile("Error.txt", $"frmLotManagerSink.LotOpened: {ex.HResult} {ex.Message}")
            Err.Clear()
        End Try
    End Sub



    'Routine: frmLotManagerSink.RecvClosedLot()
    'Purpose: This passes back this objects closed lots when another form requests them.
    Public Sub RecvClosedLots(ByRef colLots As List(Of String), ByRef colLotDates As List(Of Date))
        ' Ensure the collections are initialized
        If m_colClosed IsNot Nothing AndAlso m_colClosedDate IsNot Nothing Then
            ' Normalize the DateTime values before returning
            colLots = New List(Of String)(m_colClosed)
            colLotDates = m_colClosedDate.Select(Function(d) DateTime.ParseExact(
            d.ToString("MM/dd/yyyy HH:mm:ss", CultureInfo.InvariantCulture),
            "MM/dd/yyyy HH:mm:ss",
            CultureInfo.InvariantCulture
        )).ToList()
        Else
            colLots = New List(Of String)()
            colLotDates = New List(Of Date)()
        End If
    End Sub

    ' Routine: frmLotManagerSink.RecvSuspendedLots()
    ' Purpose: This passes back this object's suspended lots when another form requests them.
    Public Sub RecvSuspendedLots(ByRef colLots As List(Of String), ByRef colLotDates As List(Of Date))
        ' Ensure the collections are initialized
        If m_colSuspended IsNot Nothing AndAlso m_colSuspendedDate IsNot Nothing Then
            ' Normalize the DateTime values before returning
            colLots = New List(Of String)(m_colSuspended)
            colLotDates = m_colSuspendedDate.Select(Function(d) DateTime.ParseExact(
            d.ToString("MM/dd/yyyy HH:mm:ss", CultureInfo.InvariantCulture),
            "MM/dd/yyyy HH:mm:ss",
            CultureInfo.InvariantCulture
        )).ToList()
        Else
            colLots = New List(Of String)()
            colLotDates = New List(Of Date)()
        End If
    End Sub

    ' Routine: frmLotManagerSink.RecvOpenLots()
    ' Purpose: This passes back this object's open lot when another form requests it.
    Public Sub RecvOpenLots(ByRef strOpenLot As String, ByRef dtOpen As Date)
        ' Ensure the open lot and date are returned with normalized DateTime
        strOpenLot = If(m_strOpenLot, String.Empty)
        dtOpen = DateTime.ParseExact(
        m_dtOpen.ToString("MM/dd/yyyy HH:mm:ss", CultureInfo.InvariantCulture),
        "MM/dd/yyyy HH:mm:ss",
        CultureInfo.InvariantCulture
    )
    End Sub

    'Routine: frmLotManagerSink.Property Get LotManager()
    'Purpose: This passes back a reference to this objects LotManager instance.
    Public Property LotManager As NextcapServer.clsLotManager
        Get
            Return m_oLotManager
        End Get
        Set(value As NextcapServer.clsLotManager)
            m_oLotManager = value
        End Set
    End Property

    'Routine: frmLotManagerSink.Property Get ChartData()
    'Purpose: This passes back a reference to the this objects ChartData collection instance.
    Public ReadOnly Property ChartDataCol As colChartData
        Get
            Return m_colChartData
        End Get
    End Property

    'Routine: frmLotManagerSink.IsStation(str,n,str)
    'Purpose: This returns a True/False for this station matching a requested station type.
    Public Function IsStation(ByRef strLine As String, ByRef nLineNumber As Integer, ByRef strSource As String) As Boolean
        'See if this matches our LotManager objects Station-Key
        If m_oLotManager IsNot Nothing Then
            If m_oLotManager.LineType = strLine Then 'Match the line type
                If m_oLotManager.LineNumber = nLineNumber Then 'Match the line number
                    If m_oLotManager.Source = strSource Then 'Match the source type
                        Return True 'Return true if all match
                    End If
                End If
            End If
        End If
        Return False
    End Function

    ''Routine: frmLotManagerSink.CloseLotUnique(str,dt)
    ''Purpose: This tracks to see if we have already serviced a request to closed the specififed Lot.
    'Private Function CloseLotUnique(ByRef strGroupId As String, ByRef dtBirth As Date) As Boolean
    '    Dim lngItem As Integer

    '    'Cycle through the collection from oldest to newest
    '    For lngItem = m_colTrackClosedId.Count - 1 To 0 Step -1
    '        If m_colTrackClosedId(lngItem) = strGroupId AndAlso m_colTrackClosedDate(lngItem) = dtBirth Then
    '            'Exit out of the function returning FALSE
    '            Return False
    '        End If
    '    Next lngItem

    '    'Trim the collection
    '    While m_colTrackClosedId.Count > MAX_TRACK
    '        m_colTrackClosedId.RemoveAt(0)
    '        m_colTrackClosedDate.RemoveAt(0)
    '    End While

    '    'Add the newest Group Id
    '    m_colTrackClosedId.Add(strGroupId)
    '    m_colTrackClosedDate.Add(dtBirth)
    '    Return True
    'End Function

    ''Event raised when security changes
    'Private Sub m_oSecurity_SecurityChange(ByVal nNewAuthority As Integer)
    '    'This just sets the local copy of the Authority level
    '    m_nAuthority = nNewAuthority
    'End Sub

    Public Sub InitializeData()
        Dim dtAllDefectCount As DataTable = GetAllData() ' LotId, Birthday, ClassName(Cosmetic/Functional/Good/Risk), [Count] From LotDefectCounts
        Dim dtStatus As DataTable = GetAllQualityStatus() ' LotId, Birthday, QualityStatus From Lots

        If dtAllDefectCount IsNot Nothing AndAlso dtAllDefectCount.Rows.Count > 0 Then
            For Each row As DataRow In dtAllDefectCount.Rows
                Dim strGroupId As String = row("LotId").ToString()
                Dim dtBirth As DateTime?

                ' Parse the Birthday column using mdlTools.ParseDateTime
                If dtAllDefectCount.Columns("Birthday").DataType = GetType(DateTime) Then
                    dtBirth = row("Birthday")
                Else
                    dtBirth = Nothing
                    Debug.WriteLine($"Invalid Birthday format: {row("Birthday")}")
                End If



                ' Process defect counts and quality status
                For Each dtStatus_row As DataRow In dtStatus.Rows
                    Dim dtstatus_lotid As String = dtStatus_row("LotId").ToString()

                    If dtstatus_lotid = strGroupId Then
                        Dim QualityStatus As String = dtStatus_row("QualityStatus").ToString()

                        ' Iterate through defect class columns
                        For Each col As DataColumn In dtAllDefectCount.Columns
                            If Not (col.ColumnName.Equals("LotId", StringComparison.OrdinalIgnoreCase) _
                                Or col.ColumnName.Equals("Birthday", StringComparison.OrdinalIgnoreCase)) Then
                                ' Safely get the integer value; if NULL then set to -1
                                Dim value As Integer = -1
                                If Not IsDBNull(row(col.ColumnName)) Then
                                    value = Convert.ToInt32(row(col.ColumnName))
                                End If
                                Debug.WriteLine($"{strGroupId} {col.ColumnName}: {value}")
                                m_oLotManager_UpdateDefectCount(strGroupId, dtBirth.Value, col.ColumnName, value, QualityStatus)
                            End If
                        Next
                    End If
                Next
            Next
        End If
    End Sub

    Public Function GetAllData() As DataTable
        If m_oLotManager Is Nothing Then Return Nothing

        Dim dt As DataTable = m_oLotManager.GetAllDefectCounts()
        If dt.Rows.Count = 0 OrElse dt.Columns.Count = 0 Then Return Nothing
        If dt.Rows(0)(0).ToString() = "ERROR" Then
            mdlMain.ReportServerError("LotManagerSink.GetAllData", dt)
            Return Nothing
        End If

        ' Get all unique class names
        Dim classNames = dt.AsEnumerable().
        Select(Function(r) r.Field(Of String)("ClassName")).
        Distinct().
        ToList()

        ' Prepare pivot structure
        Dim pivotTable As New DataTable()
        pivotTable.Columns.Add("LotId", GetType(String))
        pivotTable.Columns.Add("Birthday", GetType(DateTime))
        For Each className In classNames
            pivotTable.Columns.Add(className, GetType(Integer))
        Next

        ' Normalize birthday values using invariant datetime formatting
        Dim groups = dt.AsEnumerable().
        GroupBy(Function(r)
                    Dim rawDate As DateTime = r.Field(Of DateTime)("Birthday")
                    ' Format to Access-safe string and re-parse to ensure consistency
                    Dim normalized As DateTime = DateTime.ParseExact(
                        rawDate.ToString("MM/dd/yyyy HH:mm:ss", CultureInfo.InvariantCulture),
                        "MM/dd/yyyy HH:mm:ss",
                        CultureInfo.InvariantCulture
                    )
                    Return New With {
                        Key .LotId = r.Field(Of String)("LotId"),
                        Key .Birthday = normalized
                    }
                End Function)

        ' Build pivot rows
        For Each g In groups
            Dim newRow = pivotTable.NewRow()
            newRow("LotId") = g.Key.LotId
            newRow("Birthday") = g.Key.Birthday

            For Each className In classNames
                Dim found = g.FirstOrDefault(Function(r) r.Field(Of String)("ClassName") = className)
                newRow(className) = If(found IsNot Nothing, found.Field(Of Integer)("Count"), DBNull.Value)
            Next

            pivotTable.Rows.Add(newRow)
        Next

        Return pivotTable
    End Function

    ' Get all quality status
    Public Function GetAllQualityStatus() As DataTable
        If m_oLotManager Is Nothing Then Return Nothing

        Dim dt As DataTable = m_oLotManager.GetAllQualityStatuses()

        If dt.Rows.Count = 0 OrElse dt.Columns.Count = 0 Then Return dt

        If dt.Rows(0)(0).ToString() = "ERROR" Then
            mdlMain.ReportServerError("LotManagerSink.GetAllQualityStatus", dt)
            Return dt
        End If

        ' Normalize DateTime columns explicitly to avoid ambiguity
        For Each row As DataRow In dt.Rows
            If dt.Columns.Contains("Birthday") AndAlso Not IsDBNull(row("Birthday")) Then
                Dim rawDate As DateTime = CType(row("Birthday"), DateTime)

                ' Reformat and re-parse to eliminate hidden millisecond/ticks inconsistency
                row("Birthday") = DateTime.ParseExact(
                rawDate.ToString("MM/dd/yyyy HH:mm:ss", CultureInfo.InvariantCulture),
                "MM/dd/yyyy HH:mm:ss",
                CultureInfo.InvariantCulture
            )
            End If
        Next

        Return dt
    End Function

    Private Sub m_oLotManager_UpdateDefectCount(strLotId As String, dateBirthday As DateTime, strClassName As String, intNewCount As Integer, strQualityStatus As String) Handles m_oLotManager.UpdateDefectCount

        Try
            ' Ensure valid DateTime before passing
            If dateBirthday = DateTime.MinValue Then
                LogToFile("Error.txt", $"Invalid birthday for LotId: {strLotId}")
                Exit Sub
            End If

            mdlChart.AddData(strLotId, dateBirthday, strClassName, intNewCount, m_colChartData, strQualityStatus)
        Catch ex As Exception
            LogToFile("Error.txt", $"Error in m_oLotManager_UpdateDefectCount for LotId {strLotId}: {ex.Message}")
        End Try
    End Sub

    Private Sub m_oLotManager_LotClosed(ByVal strGroupId As String, ByVal dtBirth As Date) Handles m_oLotManager.LotClosed
        Debug.WriteLine("this is frmLotManagerSink m_oLotManager_LotClosed()")
    End Sub

    Public Sub m_oLotManager_TimeToCreateLot() Handles m_oLotManager.TimeToCreateLot
        Dim oSequence As clsLotSequence

        If m_nAuthority > 3 AndAlso Not mdlLotManager.ActiveLotManagerForm.AutoCreateLot AndAlso
           (go_clsSystemSettings.strMaterialMode <> mdlGlobal.gcstrContinuous) Then
            'Console.WriteLine({mdlWindow.GetTopForm().Name})
            If mdlWindow.GetTopForm() IsNot Nothing AndAlso mdlWindow.GetTopForm().Name <> "frmLotManager" Then
                oSequence = gcol_LotSequence.Add("", New Date(1999, 1, 1), m_oLotManager, m_colChartData, Me)
                oSequence.nRequestCreate = 1
            End If
        End If
    End Sub

    Public Sub m_oLotManager_TimeToCloseLot(ByVal strLotId As String, ByVal dateBirthday As Date) Handles m_oLotManager.TimeToCloseLot
        Dim oSequence As clsLotSequence

        If m_nAuthority > 3 AndAlso Not mdlLotManager.ActiveLotManagerForm.AutoCreateLot AndAlso
       (go_clsSystemSettings.strMaterialMode <> mdlGlobal.gcstrContinuous) Then
            If mdlWindow.GetTopForm() IsNot Nothing AndAlso mdlWindow.GetTopForm().Name <> "frmLotManager" Then
                '/*Place a request on the stack to show this Group ending
                oSequence = gcol_LotSequence.Add(strLotId, dateBirthday, m_oLotManager, m_colChartData, Me)
                oSequence.nRequestClose = 1
            End If
        End If
    End Sub

    Private Sub m_oLotManager_UpdateMaterialStatus(strLotId As String, dateBirthday As Date, strQualityStatus As String) Handles m_oLotManager.UpdateMaterialStatus
        Try
            gn_ImageCheckCount += 1  ' 简化自增语法

            ' 调用 mdlChart.SetStatusIcon，注意 dateBirthday 已为 Date 类型，无需 vtoDate 转换
            mdlChart.SetStatusIcon(strLotId, dateBirthday, strQualityStatus, m_colChartData)

            gn_ImageCheckCount -= 1  ' 简化自减语法
            If gn_ImageCheckCount <= 0 AndAlso m_colChartData.bActive Then
                mdlChart.ExecuteImageCheck()
                gn_ImageCheckCount = 0
            End If

        Catch ex As Exception
            ' 异常发生时执行的逻辑（对应原 On Error 跳转后的处理）
            gn_ImageCheckCount -= 1
            If gn_ImageCheckCount <= 0 AndAlso m_colChartData.bActive Then
                mdlChart.ExecuteImageCheck()
                gn_ImageCheckCount = 0
            End If

            ' 记录错误日志，使用 ex 替代原 Err 对象
            LogToFile("Error.txt", $"frmLotManagerSink.m_oLotManager_UpdateMaterialStatus: {ex.HResult} {ex.Message}")
            ' VB.NET 无需显式 Clear 异常，Catch 块执行后自动处理

        End Try
    End Sub
    Private Sub m_oLotManager_ClosedLotMaterialStatus(ByVal strLotId As String, ByVal dateBirthday As Date, ByVal strQualityStatus As String) Handles m_oLotManager.ClosedLotMaterialStatus
        ' 检查 ImageCheck 入口计数
        gn_ImageCheckCount += 1  ' 简化自增语法

        ' 检查交易权限
        If m_nAuthority > 3 AndAlso Not frmLotManager.AutoCreateLot Then
            Dim topForm As Form = GetTopForm()
            If topForm IsNot Nothing AndAlso topForm.Name <> "frmLotManager" Then
                gcol_LotSequence.SetQualityStatus(strLotId, dateBirthday, strQualityStatus)
            End If
        End If


        ' 更新集合中的状态
        mdlChart.SetStatusIcon(strLotId, dateBirthday, strQualityStatus, m_colChartData)  ' 移除 Call 关键字，VB.NET 中可省略


        ' 检查 ImageCheck 出口计数
        gn_ImageCheckCount -= 1  ' 简化自减语法
        If gn_ImageCheckCount <= 0 AndAlso m_colChartData.bActive Then
            mdlChart.ExecuteImageCheck()  ' 补充方法调用括号，符合 VB.NET 规范
            gn_ImageCheckCount = 0
        End If
    End Sub
    Private Function GetTopForm() As Form
        ' 遍历所有已打开的表单，找到当前获得焦点的表单（最上方）
        For Each form As Form In Application.OpenForms
            ' 检查表单是否可见且获得焦点（通常就是最上方的表单）
            If form.Visible AndAlso form.Focused Then
                Return form
            End If
        Next

        ' 如果没有找到焦点表单，返回第一个打开的表单（备选逻辑）
        If Application.OpenForms.Count > 0 Then
            Return Application.OpenForms(0)
        End If

        Return Nothing
    End Function


End Class
