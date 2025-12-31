Imports System.Globalization
Imports System.Threading
Imports LiveCharts
Imports LiveCharts.Wpf
Imports ncapsrv
Imports NextCapServer = ncapsrv
Module mdlMain
    '/*Define the app name used in the registry
    Public Const cstrRegName As String = "NextCap"
    '/*Constant to define server error command
    Public Const cServerError As String = "ERROR"
    '/*command line argument to indicate what are mode is
    Public strMode As String
    Public Const cstrStandAlone As String = "STANDALONE"
    Public Const m_lngHotMax As Integer = 3
    '/*Flag to indicate that we are in teh process of shutting down
    Public ShutDownFlag As Boolean
    Private serverSinkForm As frmServerSink

    '/*Constants to define state machine
    Private Const cnINPUT_READY As Integer = 0
    Private Const cnINPUT_GOOD As Integer = 1
    Private Const cnINPUT_BAD As Integer = 2
    Private Const cnINPUT_DELETE As Integer = 3
    Private Const cnINPUT_EDIT As Integer = 4
    Private Const cnINPUT_UNDO As Integer = 5
    Private Const cnINPUT_EXECUTE As Integer = 6
    'Private Const cServerError As String = "ERROR"
    Dim oFeatureClass As clsFeatureClass
    Dim oFeature As clsFeature

    '/*Current input state
    Private m_nState As Integer
    Private m_strLastButtonKey As String
    '/*Setting for input mode
    Private m_strInputMode As String
    Private Const cstrScanMode As String = "SCANNER"
    '/*Flag to indicate if we are in continuous bad pen mode
    Private m_bContinuousBadMode As Boolean


    ' Declare the cartesianChart1 (from LiveCharts.WinForms)
    Private WithEvents cartesianChart1 As New LiveCharts.WinForms.CartesianChart()
    Private WithEvents cartesianChart2 As New LiveCharts.WinForms.CartesianChart()

    ' Create a ContextMenuStrip for right-click options
    Private contextMenu As New ContextMenuStrip()

    'Public WithEvents go_BusinessServer As NextCapServer.clsBoot

    Private WithEvents m_oLotManager As NextCapServer.clsLotManager
    'Public go_Supervisor As NextCapServer.clsSupervisor
    Private selectedChartPoint As ChartPoint ' Store the clicked bar's data

    ' Series for Stacked Bars
    Private correctSeries As StackedColumnSeries
    Private errorSeries As StackedColumnSeries

    Private m_colChartData As colChartData
    Public oLotsToday As clsLotCountToday
    Public frmNextCapInstance As frmNextCap

    Public Sub Main()
        Try
            Console.WriteLine("Starting Nextcap...")
            CreateLogEvent()


            LogEvent("Instantiate the Security object.")
            '/*Instantiate the Security object
            go_Security = New Security()

            LogEvent("Build a new system settings object.")
            '/*Build a new clsSystemSettings object
            GenerateSystemObject()

            LogEvent("Initilialize the golbal context object.")
            '/*Initilialize the golbal context object
            go_Context = New Context()

            LogEvent("Create the LotManagerSink server request stack.")
            '/*Create the LotManagerSink server request stack
            gcol_LotSequence = New colLotSequences

            LogEvent("Create the Lot Manager global collection.")
            '/*Create the Lot Manager global collection
            gcol_LotManagers = New List(Of frmLotManagerSink)

            LogEvent("Initialize the gcol_LocalApps object.")
            '/*Initialize the gcol_LocalApps object
            GenerateLocalAppsObject()

            LogEvent("Generate the global collection of Quality Statuses.")
            '/*Generate the global collection of Quality Statuses
            gcol_QualityStatus = New colQualityStatuses

            LogEvent("Generate a new chart data object")
            '/*Generate a new chart data object
            gcol_ChartData = New colChartData ' populate by mdlChart.InitializeChart()-> sink.InitializeData()

            LogEvent("Display the Splash screen so the user knows we are starting.")
            '/*Display the Splash screen so the user knows we are starting
            Dim splash As New frmSplash()
            splash.Show()

            Application.DoEvents()
            LogEvent("Connect to the business server and establish any COM sinks through frmServerSink")
            If EstablishConnection() Then
                LogEvent("Get the runtime parameters from the server")
                GenerateRunTimeParameters()
                LogEvent("Generate the station key collection")
                gcol_StationKeys = New colStationKeys()
                ' Execute required steps for a login condition


                LogEvent("Retrieve the station keys (LineType,LineNumber,Source)")
                '/*Retrieve the station keys (LineType,LineNumber,Source)
                If RecvStations(go_Supervisor) Then
                    RecvAccumulators(go_Supervisor)
                    RecvItemTypes(go_Supervisor)
                    RecvRunTypes(go_Supervisor)
                    RecvTestBeds(go_Supervisor)
                    RecvShifts(go_Supervisor)
                    RecvLotComments(go_Supervisor)
                    RecvUsers(go_Supervisor)
                    RecvProducts(go_Supervisor)
                    RecvTasks(go_Supervisor)
                    RecvQualityStates(go_Supervisor)
                    RecvRecoverySteps(go_Supervisor)
                End If
                GenerateLotManager() 'go_ActiveLotManager for chart data; populate lotdata only

                'frmSAX.AddActiveLotSink()
            Else
                MessageBox.Show("Could not connect to the NextCap server for " & CurrentMachineName() & ", contact the application support, existing...", "Server Connect Error")
                Application.Exit()
                Return ' Ensure no further code runs in Main()
            End If
            LogEvent("Initialize the CAPmain object used locally")
            '/*Initialize the CAPmain object used locally
            ' GenerateCAPMainObject

            'LogEvent("create colSAXhandlers object")
            '/*Create colSAXhandlers
            GenerateSAXObject()

            'LogEvent("Generate colSAXTranslations for the current Collection of SAX Translations")
            '/*Generate the current Collection of SAX Translations
            CreateProtoTypes()

            'LogEvent("load the Predefined set of CSBs")
            '/*Generate the Predefined set of CSBs
            GenerateCSBs()

            LogEvent("Initialize go_SampleCount")
            '/*Initialize go_SampleCount, taking from registry
            GenerateSampleCount()

            LogEvent("Generate a new Chart Configuration Object")
            '/*Generate a new Chart Configuration Object
            go_ChartConfigs = New clsChartConfig

            LogEvent("Configure the GoodPenEnter screen")
            '/*Configure the GoodPenEnter screen
            'ConfigurePenEntry

            LogEvent("Restore the settings for the Defect Editor")
            '/*Restore the settings for the Defect Editor
            'frmDefectEditor.Refresh()
            'RestoreWindowPos(frmDefectEditor)


            LogEvent("toggle where the source for the configurations is coming from based on the start-up mode")
            '/*toggle where the source for the configurations
            '/*is coming from based on the start-up mode
            '    If strMode = cstrStandAlone Then
            '        LogEvent("Build up the Local Apps screen if there is any data"
            ''/*Build up the Local Apps screen if there is any data
            '        If mdlLocalApp.Test_LoadLocalApps() Then
            '            mdlLocalApp.FillLocalAppList
            '        Else
            '            MsgBox("No demo file found for app.path\Testfiles\LocalApps.txt")
            '        End If
            '    Else
            '        mdlLocalApp.FillLocalAppList
            '    End If

            LogEvent("Fetch the Feature class headers (this is actually just thier descriptions)")
            '/*Fetch the Feature class headers (this is actually just thier descriptions)
            mdlMain.RecvDefectClass(go_Supervisor)

            LogEvent("Populate the lists")
            '/*Populate the lists
            If Not (gcol_FeatureClasses Is Nothing) Then
                mdlDefectEditor.ConfigureDefectEditor()
            End If

            '/*Build up the Local Apps screen if there is any data
            LogEvent("Build up the Local Apps screen if there is any data")
            'mdlLocalApp.FillLocalAppList

            '/*This forces the Form to be Loaded but not viewed
            ' refresh does not trigger custom methods like ShowMe, nor does it re-run Form_Load or any other initialization code. The Refresh method only repaints the form and its controls
            ' Dim frmNextCapInstance As New frmNextCap()
            frmNextCapInstance = New frmNextCap()

            ' Start a background task for non-modal operations

            LogEvent("Populate the Chart Configuration Object")
            '/*Populate the Chart Configuration Object
            mdlChart.LoadChartConfig() ' this will new frmNextcap -> mdlChart.SetChartStaticProperties

            'If frmNextCap.panelStatusIcons IsNot Nothing Then
            '    Debug.WriteLine("panelStatusIcons is created.")
            'Else
            '    Debug.WriteLine("panelStatusIcons is NOT created.")
            'End If

            'If frmNextCap.imgStatus IsNot Nothing AndAlso frmNextCap.imgStatus.Count > 0 Then
            '    Debug.WriteLine("imgStatus is created and has PictureBoxes.")
            'ElseIf frmNextCap.imgStatus IsNot Nothing Then
            '    Debug.WriteLine("imgStatus is created but empty.")
            'Else
            '    Debug.WriteLine("imgStatus is NOT created.")
            'End If

            LogEvent("Set the default on the Production Date input.Removed to allow the object to drive this property.")
            '/*Set the default on the Production Date input.
            '/*Removed to allow the object to drive this property.
            frmAllContext.GetInstance().EnableProductionDate = go_clsSystemSettings.bProductionDateAdjust




            LogEvent("Load the data for the chart")
            '/*Load the data for the chart; populate colChartData

            mdlChart.InitializeChart()
            Debug.WriteLine("Checking inmdlMain after mdlChart.InitializeChart()")

            If frmLotManager.ActiveLotManager IsNot Nothing Then
                Dim chartDataCollection = frmLotManager.ActiveLotManager.ChartDataCol
                If chartDataCollection IsNot Nothing Then
                    For Each chartDataItem In chartDataCollection
                        Console.WriteLine($"Group ID: {chartDataItem.strGroupId}, Birth Date: {chartDataItem.dtBirth}")
                    Next
                Else
                    Console.WriteLine("after InitializeChart gcol_ChartData is empty or uninitialized.")
                End If
            Else
                Console.WriteLine("after InitializeChart gcol_ChartData is not initialized.")
            End If

            LogEvent("Set the flag indicating whether the Status Bar CSB is in use")
            '/*Set the flag indicating whether the Status Bar CSB is in use
            ' mdlSAX.SetStatusBarFlag

            LogEvent("Bring the Reports module on-line and build the default pen header")
            '/*Bring the Reports module on-line and build the default pen header
            ' mdlReports.BuildDefaultPenHeader

            LogEvent("Unload the Splash screen since we are almost done")
            '/*Unload the Splash screen since we are almost done
            splash.Close()
            splash.Dispose()
            LogEvent("This forces the Form to be Loaded but not viewed")


            LogEvent("Restore its position")
            '/*Restore its position
            'RestoreWindowPos(frmNextCap)

            LogEvent("Insure that the message window is down if it is only a warning")

            '/*Insure that the message window is down if
            '/*it is only a warning
            If frmMessage.Visible Then
                If frmMessage.Text = "Warning" Then
                    frmMessage.Hide()
                End If
            End If

            LogEvent("Increment the ImageCheck concurency counter")
            '/*Increment the ImageCheck concurency counter
            gn_ImageCheckCount = gn_ImageCheckCount + 1



            LogEvent("Decrement after section")
            '/*Decrement after section
            'gn_ImageCheckCount = gn_ImageCheckCount - 1
            LogEvent("Update the server's copy of the context")
            '/*Update the server's copy of the context
            'frmServerSink.UpdateServerContext()

            LogEvent("Execute user script from ClientSystemStartUp")
            '/*Run the user's CSB initialize script
            mdlSAX.ExecuteCSB_ClientSystemStartUp()

            ' -- Set the CSB IDE menu option
            'frmNextCap.frmNextCap_CSB_IDE.Visible = go_clsSystemSettings.bEnableCSBIDE
            LogEvent("If a login is required lock down the client and display the login screen Test - go_Security.bLoginRequired = True")
            '/*If a login is required lock down the client
            '/*and display the login screen
            ' Test - go_Security.bLoginRequired = True
            If go_Security.bLoginRequired Then
                frmPassword.StartUpLogin()
            Else
                LogEvent(" the NextCap procedure to initialize and show the window")
                mdlMain.frmNextCapInstance.ShowMe()
                mdlMain.frmNextCapInstance.PerformLayout()
                mdlMain.frmNextCapInstance.Refresh()

                Application.Run(mdlMain.frmNextCapInstance)
            End If


        Catch ex As Exception
            MainErrorHandler("MainModule.Main", ex.Message)
        End Try
    End Sub
    Private Sub PerformBackgroundTasks()
        Try
            ' Example of non-modal operations
            While True
                Console.WriteLine("Performing background tasks...")
                Thread.Sleep(1000) ' Simulate work
            End While
        Catch ex As Exception
            Console.WriteLine($"Error in background tasks: {ex.Message}")
        End Try
    End Sub
    '=======================================================
    'Routine: GenerateCSBs()
    'Purpose: This performs the orderly call sequence to
    'the server to get each of the CSB blob's that it
    'is holding.
    '
    'Globals:None
    '
    'Input:None
    '
    'Return:None
    '
    'Tested:
    '
    'Modifications:
    '   02-24-1999 As written for Pass1.7
    '
    '   07-08-1999 switched from detecting @ in the RT
    '   value to detecting <> "undefined"
    '
    '   08-09-2001 Adding in v1.1.2 csb hooks for lot
    '   open changes.
    '
    '   08-16-2001 v1.1.3 csb hook for CSBVetoLotStatusChange
    '
    '   08-24-2001 v1.1.5 csb for LotManagerChange
    '=======================================================
    Public Sub GenerateCSBs()
        Const cstrNone As String = "Undefined"
        Dim oRT As NextCapServer.clsRuntime
        Dim strTemp As String

        ' 获取运行时对象引用
        oRT = go_BusinessServer.GetRTObject()

        ' 定义需要处理的 CSB 类型列表
        Dim csbTypes As New List(Of String) From {
            "CSBAugmentPenId", "CSBCreateLot", "CSBEndOfLot", "CSBPostBadPen",
            "CSBPostGoodPen", "CSBPostShiftChange", "CSBPreCreateLot",
            "CSBPreShiftChange", "CSBStatusBar", "CSBVerifyPenid",
            "CSBContextLock", "CSBLotWizard", "CSBClientSystemStartUp",
            "CSBLotOpened", "CSBClosedLotOpened", "CSBSuspendedLotOpened",
            "CSBLotSuspended", "CSBVetoLotStatusChange", "CSBLotManagerChange",
            "CSBVerifyLotId"
        }

        ' 遍历处理所有 CSB 类型
        For Each csbType In csbTypes
            strTemp = Convert.ToString(oRT(csbType))  ' 替代 vtoa 函数
            'strTemp = csbType
            If strTemp <> cstrNone Then
                PassCSBCodeToRecv(strTemp)
            End If
        Next
    End Sub
    '=======================================================
    'Routine: PassCSBCodeToRecv(str)
    'Purpose: This parses the CSB code received from the
    'business server and passes it on to be analyzed and
    'loaded into the SAX engine.
    '
    'Globals:None
    '
    'Input:None
    '
    'Return:None
    '
    'Tested:
    '
    'Modifications:
    '   02-24-1999 As written for Pass1.7
    '
    '   07-08-1999 Modified to reflect that there is no
    '   '@' in the name anymore.
    '=======================================================
    Public Sub PassCSBCodeToRecv(ByRef strCSB As String)
        If strCSB.Contains("@") Then
            ' 使用 IndexOf 替代 InStr，Substring 替代 Left$ 和 Right$
            Dim atIndex As Integer = strCSB.IndexOf("@")
            mdlMain.RecvCSBs(go_Supervisor,
                             strCSB.Substring(0, atIndex),
                             strCSB.Substring(atIndex + 1))
        Else
            Dim spaceIndex As Integer = strCSB.IndexOf(" ")
            If spaceIndex > 0 Then ' 防止空格在开头导致异常
                mdlMain.RecvCSBs(go_Supervisor,
                                 strCSB.Substring(0, spaceIndex),
                                 strCSB.Substring(spaceIndex + 1))
            End If
        End If
    End Sub

    Public Function RecvCSBs(ByRef oSupervisor As clsSupervisor, ByRef strCSBname As String, ByRef dtVersion As String) As Boolean
        Dim vrtArray As DataTable  ' 显式声明为 DataTable 类型

        Try
            ' 确保对象不为空
            If oSupervisor IsNot Nothing Then
                ' Try to parse the version date string
                Dim versionDate As DateTime
                Dim isValid As Boolean = DateTime.TryParseExact(
                    dtVersion,
                    "M/d/yyyy h:mm:ss tt",
                    CultureInfo.InvariantCulture,
                    DateTimeStyles.None,
                    versionDate
                )

                If Not isValid Then
                    Console.WriteLine("mdlMain.RecvCSBs()", "Invalid date format for dtVersion: " & dtVersion)
                    Return False
                End If
                ' 从服务器请求 CSB（假设返回 DataTable）
                ' CSB will return vrtArray , it contains CSBType, Program , CSBStatus
                vrtArray = oSupervisor.GetCSB(strCSBname, versionDate)

                ' 检查服务器错误（假设第一行第一列为错误码）
                If vrtArray IsNot Nothing AndAlso vrtArray.Rows.Count > 0 AndAlso vrtArray.Columns.Count > 0 AndAlso vrtArray.Rows(0)(0).Equals(cServerError) Then

                    ' 通过集中式处理器路由错误
                    ReportServerError("mdlMain.RecvCSBs()", vrtArray)
                Else
                    ' 记录 CSB 代码（假设第二行第一列为代码）
                    If vrtArray IsNot Nothing AndAlso vrtArray.Rows.Count > 0 AndAlso vrtArray.Columns.Count > 0 Then
                        Dim cScript As String = Convert.ToString(vrtArray.Rows(0)(1))
                        mdlSAX.AddSAXmodule(cScript)
                    End If
                End If
            End If

            Return True  ' 默认返回成功

        Catch ex As Exception
            ' 处理异常
            MainErrorHandler("mdlMain.RecvCSBs()", $"{ex.Message}-{ex.HResult}")
            Return False  ' 发生异常时返回失败

        Finally
            ' 释放 DataTable 资源
            If vrtArray IsNot Nothing Then
                vrtArray.Dispose()
                vrtArray = Nothing
            End If
        End Try
    End Function

    '=======================================================
    'Routine: mdlMain.ReportServerError()
    'Purpose: This displays any error encountered from the
    'server. It also decodes the variant array error message
    'returned by the ServerSink/LotManager.
    '
    'Globals:None
    '
    'Input:None
    '
    'Return:None
    '
    'Modifications:
    '   11-18-1998 As written for Pass1.4
    '
    '
    '=======================================================
    Public Sub ReportServerError(ByRef strRoutine As String, ByRef vrtError As DataTable)
        Dim strError As String = "Unknown Error"

        ' 检查 DataTable 是否包含错误信息
        If vrtError IsNot Nothing AndAlso
           vrtError.Rows.Count > 0 AndAlso
           vrtError.Columns.Count > 1 Then

            ' 获取错误信息（假设第二列包含错误描述）
            strError = Convert.ToString(vrtError.Rows(0)(1))
        End If

        ' 构建完整的错误消息
        strError = "Recv Business Server Error" & vbCrLf & "Server Message: " & strError

        ' 调用主错误处理程序
        MainErrorHandler(strRoutine, strError, "Warning")
    End Sub

    '=======================================================
    'Routine: mdlMain.MainErrorHandler(str)
    'Purpose: This is a central point for controling how
    '         errors are reported to the user and if
    '         the error is logged.
    '
    'Globals:None
    '
    'Input: strCaller - The name of the module and routine.
    '
    '       strMsg - The error description and number.
    '
    'Return:None
    '
    'Modifications:
    '   11-24-1998 As written for Pass1.5
    '
    '   12-16-1998 Switched to using Form to prevent
    '   code lock down from a Modal Msgbox.
    '=======================================================
    Public Sub MainErrorHandler(ByVal strCaller As String, ByVal strMsg As String, Optional ByVal strWindowHeader As String = "Error")
        Dim strMsgOut As String
        '/*Cat the output message
        strMsgOut = strCaller & vbCrLf & strMsg
        '/*Raise the form
        'frmMessage.GenerateMessage(strMsgOut, strWindowHeader)
    End Sub
    '=======================================================
    'Routine: mdlMain.GenerateSAXObject()
    'Purpose: This generates and configures the Collection
    '         that holds the SAX engine handlers and code.
    '
    'Globals:None
    '
    'Input:None
    '
    'Return:None
    '
    '
    'Modifications:
    '   10-22-1998 As written for Pass1.2
    '
    '
    '=======================================================
    Public Sub GenerateSAXObject()
        gcol_SAXhandlers = New colSAXhandlers
    End Sub

    Public Function RecvDefectClass(oSupervisor As NextCapServer.clsSupervisor) As Boolean
        Dim dtDefectClass As DataTable
        Dim stationKey As clsStationKey
        Try
            If oSupervisor Is Nothing Then Return False

            ' Initialize the feature‐classes collection
            gcol_FeatureClasses = New colFeatureClasses()
            ' 遍历所有站点键
            For lngItem As Integer = 0 To gcol_StationKeys.Count - 1

                stationKey = gcol_StationKeys(lngItem)

                dtDefectClass = oSupervisor.GetDefectClasses(
                        stationKey.strLineType,
                        stationKey.nLineNumber,
                        stationKey.strSource,
                        "PEN"
                        )

                ' If the first cell signals a server‐error code, report and skip
                If dtDefectClass.Rows.Count > 0 AndAlso dtDefectClass.Rows(0)(0).ToString() = cServerError Then
                    ReportServerError("mdlMain.RecvDefectClass", dtDefectClass)
                Else
                    ' Now retrieve Level-1 descriptions as DataTable
                    RecvLevel1Desc(oSupervisor,
                                   stationKey.strLineType,
                                   stationKey.nLineNumber,
                                   stationKey.strSource,
                                   "PEN", dtDefectClass)

                End If
            Next
            Return True

        Catch ex As Exception
            MainErrorHandler("mdlMain.RecvDefectClass", $"{ex.Message} – HResult: {ex.HResult}")
            Return False
        End Try
    End Function
    Public Sub RecvLevel1Desc(
        oSupervisor As NextCapServer.clsSupervisor,
        strLineType As String,
        nLineNumber As Integer,
        strSource As String,
        strItemType As String,
        dtDefectClass As DataTable 'ClassName,Severity,ColorIndex
    )
        Dim oFeature As clsFeature
        Dim oFeatureClass As clsFeatureClass

        If oSupervisor Is Nothing Then Exit Sub

        '"ClassName,Severity,ColorIndex"
        Const m_lngHotMax As Integer = 3

        For Each row As DataRow In dtDefectClass.Rows
            Dim strTitle As String = row("ClassName").ToString()

            If strTitle <> "N/A" Then
                Dim strKey As String = strItemType & strTitle

                Dim severity As Integer = Convert.ToInt32(row("Severity"))
                Dim colorIndex As Integer = Convert.ToInt32(row("ColorIndex"))

                If Not gcol_FeatureClasses.IsIndexUsed(strKey) Then
                    ' Add new feature class to collection
                    oFeatureClass = gcol_FeatureClasses.Add(
                                strTitle,
                                strTitle,
                                severity,
                                mdlDefectEditor.AddFeatureCol(),
                                mdlDefectEditor.AddHotFeatureCol(),
                                m_lngHotMax,
                                strItemType,
                                colorIndex,
                                strKey
                            )
                Else
                    ' Retrieve existing class

                    oFeatureClass = gcol_FeatureClasses(strKey)
                End If
                Dim vrtArrayL1 As DataTable = go_Supervisor.GetL1Descriptions(strLineType, nLineNumber, strSource, strItemType, strTitle)

                If vrtArrayL1.Rows.Count = 0 OrElse vrtArrayL1.Rows(0)(0).ToString() = cServerError Then
                    ' Route the error through the centralized handler
                    'Debug.WriteLine("mdlMain.RecvLevel1Desc", vrtArrayL1)
                Else
                    ' Debug.WriteLine("RecvLevel1Desc adding clsFeature to colClassFeatures of clsFeatureClass ")
                    ' Loop through the returned level-1 descriptions
                    For Each rowL1 As DataRow In vrtArrayL1.Rows
                        ' Add the Level-1 description

                        oFeature = oFeatureClass.colClassFeatures.Add(
                            rowL1("Description1").ToString(),   ' Adjust to match column name
                            rowL1("Code1").ToString(),
                            rowL1("Url1").ToString(),
                            mdlDefectEditor.AddSubFeatureCol()
                        )

                        ' 5) Recurse into Level-2
                        RecvLevel2Desc(
                            oSupervisor, oFeature,
                            strLineType, nLineNumber,
                            strSource, "PEN"
                        )
                    Next
                End If
            End If
        Next


    End Sub
    Public Sub RecvLevel2Desc(
        oSupervisor As NextCapServer.clsSupervisor,
        oFeature As clsFeature,
        strLineType As String,
        nLineNumber As Integer,
        strSource As String,
        strItemType As String
    )
        ' Call the service and get a DataTable
        Dim dt As DataTable = oSupervisor.GetL2Descriptions(
            strLineType, nLineNumber, strSource, "PEN", oFeature.strCode)

        If dt.Rows.Count = 0 OrElse dt.Rows(0)(0).ToString() = cServerError Then
            ' Route the error through the centralized handler
            'Debug.WriteLine("mdlMain.RecvLevel2Desc", dt)
        Else
            'Debug.WriteLine("RecvLevel2Desc adding colSub to clsFeature of colClassFeatures in clsFeatureClass ")
            ' Loop through the returned level-1 descriptions
            For Each rowL2 As DataRow In dt.Rows
                ' Add the Level-2 description

                oFeature.colSub.Add(
                                        rowL2("Description2").ToString(),   ' Adjust to match column name
                                        rowL2("Code2").ToString(),
                                        rowL2("Url2").ToString()
                                    )

            Next
        End If
    End Sub

    Private Sub GenerateSampleCount()

        '/*Generate the global counter object

        go_SampleCount = New clsSampleCount

        '/* set a pointer for the CreatePen module

        '/*to enable it to broadcast to the counter

        mdlCreatePen.PenCounter = go_SampleCount

    End Sub

    '=======================================================
    'Routine: mdlMain.RecvAccumulators(o)
    'Purpose: This requests the stations assinged to this
    '         NextCap PC.
    '
    'Globals:None
    '
    'Input: oSupervisor - A reference to the Supevisor object.
    '
    'Return: Boolean - true if the retrieval resulted in no
    '        errors.
    '
    'Modifications:
    '   11-18-1998 As written for Pass1.4
    '
    '
    '=======================================================
    Public Function RecvAccumulators(ByRef oSupervisor As clsSupervisor) As Boolean
        Dim dtAccumulators As DataTable
        Dim stationKey As clsStationKey

        Try
            If oSupervisor IsNot Nothing Then
                ' 遍历所有站点键
                For lngItem As Integer = 0 To gcol_StationKeys.Count - 1
                    stationKey = gcol_StationKeys(lngItem)

                    ' 获取累加器数据
                    dtAccumulators = oSupervisor.GetAccumulators(
                    stationKey.strLineType,
                    stationKey.nLineNumber,
                    stationKey.strSource
                )

                    ' 验证DataTable有效性
                    If dtAccumulators IsNot Nothing AndAlso dtAccumulators.Rows.Count > 0 Then
                        ' 检查是否为错误状态
                        If Convert.ToString(dtAccumulators.Rows(0)(0)) = cServerError Then
                            MainErrorHandler("mdlMain.RecvAccumulators", Convert.ToString(dtAccumulators.Rows(0)(1)))
                        Else
                            ' 遍历DataTable的行
                            For Each row As DataRow In dtAccumulators.Rows
                                ' 获取累加器数据
                                Dim strData As String = Convert.ToString(row(0))

                                ' 更新站点键的累加器（使用带参数的属性）
                                stationKey.Accumulators(0) = strData

                                ' 添加到界面显示
                                frmAllContext.GetInstance().GetInstance().AddAccumulator(strData)
                            Next
                        End If
                    End If
                Next
            End If

            Return True

        Catch ex As Exception
            MainErrorHandler("mdlMain.RecvAccumulators", $"{ex.Message}-{ex.HResult}")
            Return False
        End Try
    End Function
    '=======================================================
    'Routine: mdlMain.RecvItemTypes(o)
    'Purpose: This requests the stations assinged to this
    '         NextCap PC.
    '
    'Globals:None
    '
    'Input: oSupervisor - A reference to the Supevisor object.
    '
    'Return: Boolean - true if the retrieval resulted in no
    '        errors.
    '
    'Modifications:
    '   11-18-1998 As written for Pass1.4
    '
    '
    '=======================================================
    Public Function RecvItemTypes(ByRef oSupervisor As clsSupervisor) As Boolean
        Dim dtItemTypes As DataTable
        Dim stationKey As clsStationKey

        Try
            If oSupervisor IsNot Nothing Then
                ' 遍历所有站点键
                For lngItem As Integer = 0 To gcol_StationKeys.Count - 1
                    stationKey = gcol_StationKeys(lngItem)

                    ' 获取ItemTypes数据
                    dtItemTypes = oSupervisor.GetItemTypes(
                        stationKey.strLineType,
                        stationKey.nLineNumber,
                        stationKey.strSource
                    )

                    ' 验证DataTable有效性
                    If dtItemTypes IsNot Nothing AndAlso dtItemTypes.Rows.Count > 0 Then
                        ' 检查是否为错误状态
                        If Convert.ToString(dtItemTypes.Rows(0)(0)) = cServerError Then
                            MainErrorHandler("mdlMain.RecvItemTypes", Convert.ToString(dtItemTypes.Rows(0)(1)))
                        Else
                            ' 遍历DataTable的行
                            For Each row As DataRow In dtItemTypes.Rows
                                ' 获取ItemType数据
                                Dim strData As String = Convert.ToString(row(0))

                                ' 更新站点键的ItemTypes（使用带参数的属性）
                                stationKey.ItemTypes(0) = strData

                                ' 添加到界面显示
                                'frmDefectEditor.AddUnitType(strData)
                            Next
                        End If
                    End If
                Next
            End If

            Return True

        Catch ex As Exception
            MainErrorHandler("mdlMain.RecvItemTypes", $"{ex.Message}-{ex.HResult}")
            Return False
        End Try
    End Function
    '=======================================================
    'Routine: mdlMain.RecvTestBeds(o)
    'Purpose: This requests the stations assinged to this
    '         NextCap PC.
    '
    'Globals:None
    '
    'Input: oSupervisor - A reference to the Supevisor object.
    '
    'Return: Boolean - true if the retrieval resulted in no
    '        errors.
    '
    'Modifications:
    '   11-18-1998 As written for Pass1.4
    '
    '   04-26-1999 Added in portion to track in the TestBed
    '   Type.
    '=======================================================
    Public Function RecvTestBeds(ByRef oSupervisor As clsSupervisor) As Boolean
        Dim dtTestBeds As DataTable
        Dim stationKey As clsStationKey

        Try
            If oSupervisor IsNot Nothing Then
                ' 遍历所有站点键
                For lngItem As Integer = 0 To gcol_StationKeys.Count - 1
                    stationKey = gcol_StationKeys(lngItem)

                    ' 获取测试台数据
                    dtTestBeds = oSupervisor.GetTestbeds(
                        stationKey.strLineType,
                        stationKey.nLineNumber,
                        stationKey.strSource
                    )

                    ' 验证DataTable有效性
                    If dtTestBeds IsNot Nothing AndAlso dtTestBeds.Rows.Count > 0 Then
                        ' 检查是否为错误状态
                        If Convert.ToString(dtTestBeds.Rows(0)(0)) = cServerError Then
                            MainErrorHandler("mdlMain.RecvTestBeds", Convert.ToString(dtTestBeds.Rows(0)(1)))
                        Else
                            ' 遍历DataTable的行
                            For Each row As DataRow In dtTestBeds.Rows
                                ' 获取测试台编号和类型（假设第一列是编号，第二列是类型）
                                Dim strTestBedNumber As String = Convert.ToString(row(0))
                                Dim strTestBedType As String = Convert.ToString(row(1))

                                ' 更新站点键的测试台信息（使用带参数的属性）
                                stationKey.TestBeds(0) = strTestBedNumber
                                stationKey.TestBedTypes(0) = strTestBedType

                                ' 添加到界面显示
                                frmDefectEditor.AddTestBed(strTestBedNumber, strTestBedType)
                            Next
                        End If
                    End If
                Next
            End If

            Return True

        Catch ex As Exception
            MainErrorHandler("mdlMain.RecvTestBeds", $"{ex.Message}-{ex.HResult}")
            Return False
        End Try
    End Function
    '=======================================================
    'Routine: mdlMain.RecvRecoverySteps(o)
    'Purpose: This requests the recovery steps assinged to this
    '         NextCap PC.
    '
    'Globals:None
    '
    'Input: oSupervisor - A reference to the Supevisor object.
    '
    'Return: Boolean - true if the retrieval resulted in no
    '        errors.
    '
    'Modifications:
    '   11-23-1998 As written for Pass1.5
    '
    '
    '=======================================================
    Public Function RecvRecoverySteps(ByRef oSupervisor As clsSupervisor) As Boolean
        Dim dtRecoverySteps As DataTable
        Dim stationKey As clsStationKey

        Try
            If oSupervisor IsNot Nothing Then
                ' 遍历所有站点键
                For lngItem As Integer = 0 To gcol_StationKeys.Count - 1
                    stationKey = gcol_StationKeys(lngItem)

                    ' 获取恢复步骤数据
                    dtRecoverySteps = oSupervisor.GetRecoverySteps(
                        stationKey.strLineType,
                        stationKey.nLineNumber,
                        stationKey.strSource
                    )

                    ' 验证DataTable有效性
                    If dtRecoverySteps IsNot Nothing AndAlso dtRecoverySteps.Rows.Count > 0 Then
                        ' 检查是否为错误状态
                        If Convert.ToString(dtRecoverySteps.Rows(0)(0)) = cServerError Then
                            MainErrorHandler("mdlMain.RecvRecoverySteps", Convert.ToString(dtRecoverySteps.Rows(0)(1)))
                        Else
                            '' 清空当前恢复步骤列表
                            'frmEnterGoodPen.ClearRecoverySteps()

                            '' 遍历DataTable的行
                            'For Each row As DataRow In dtRecoverySteps.Rows
                            '    ' 获取恢复步骤数据（假设第一列包含数据）
                            '    Dim strData As String = Convert.ToString(row(0))

                            '    ' 添加到界面
                            '    frmEnterGoodPen.AddRecoveryStep(strData)
                            'Next
                        End If
                    End If
                Next
            End If

            Return True

        Catch ex As Exception
            MainErrorHandler("mdlMain.RecvRecoverySteps", $"{ex.Message}-{ex.HResult}")
            Return False
        End Try
    End Function
    '=======================================================
    'Routine: mdlMain.RecvUsers(o)
    'Purpose: This requests the stations assinged to this
    '         NextCap PC.
    '
    'Globals:None
    '
    'Input: oSupervisor - A reference to the Supevisor object.
    '
    'Return: Boolean - true if the retrieval resulted in no
    '        errors.
    '
    'Modifications:
    '   11-18-1998 As written for Pass1.4
    '
    '
    '=======================================================
    Public Function RecvUsers(ByRef oSupervisor As clsSupervisor) As Boolean
        Dim dtUsers As DataTable
        Dim stationKey As clsStationKey

        Try
            If oSupervisor IsNot Nothing Then
                ' 遍历所有站点键
                For lngItem As Integer = 0 To gcol_StationKeys.Count - 1
                    stationKey = gcol_StationKeys(lngItem)

                    ' 获取用户数据（假设返回DataTable）
                    dtUsers = oSupervisor.GetUsers(
                        stationKey.strLineType,
                        stationKey.nLineNumber,
                        stationKey.strSource
                    )

                    ' 验证DataTable有效性
                    If dtUsers IsNot Nothing AndAlso dtUsers.Rows.Count > 0 Then
                        ' 检查是否为错误状态
                        If Convert.ToString(dtUsers.Rows(0)(0)) = cServerError Then
                            MainErrorHandler("mdlMain.RecvUsers", Convert.ToString(dtUsers.Rows(0)(1)))
                        Else
                            ' 遍历DataTable的行
                            For Each row As DataRow In dtUsers.Rows
                                ' 获取用户数据（假设第一列是用户名，第二列是权限）
                                Dim strData As String = Convert.ToString(row(0))
                                Dim strAuthority As String = Convert.ToString(row(1))

                                ' 添加到全局上下文
                                frmAllContext.GetInstance().AddOperator(strData, strAuthority)
                            Next
                        End If
                    End If
                Next
            End If

            Return True  ' 若无错误，返回True

        Catch ex As Exception
            MainErrorHandler("mdlMain.RecvUsers", $"{ex.Message}-{ex.HResult}")
            Return False
        End Try
    End Function
    Public Function RecvProducts(ByRef oSupervisor As clsSupervisor) As Boolean
        Dim dtProducts As DataTable
        Dim stationKey As clsStationKey

        Try
            If oSupervisor IsNot Nothing Then
                ' 遍历所有站点键
                For lngItem As Integer = 0 To gcol_StationKeys.Count - 1
                    stationKey = gcol_StationKeys(lngItem)

                    ' 获取产品数据（假设返回DataTable）
                    dtProducts = oSupervisor.GetParts(
                        stationKey.strLineType,
                        stationKey.nLineNumber,
                        stationKey.strSource
                    )

                    ' 验证DataTable有效性
                    If dtProducts IsNot Nothing AndAlso dtProducts.Rows.Count > 0 Then
                        ' 检查是否为错误状态
                        If Convert.ToString(dtProducts.Rows(0)(0)) = cServerError Then
                            MainErrorHandler("mdlMain.RecvProducts", Convert.ToString(dtProducts.Rows(0)(1)))
                        Else
                            ' 遍历DataTable的行
                            For Each row As DataRow In dtProducts.Rows
                                ' 获取产品信息（假设列顺序：名称、编号、系列）
                                Dim strName As String = Convert.ToString(row(0))
                                Dim strNumber As String = Convert.ToString(row(1))
                                Dim strFamily As String = Convert.ToString(row(2))

                                ' 调用站点键对象的方法添加产品（假设LetParts是Sub方法）
                                stationKey.LetParts(strName, strNumber, strFamily)

                                ' 添加到全局上下文和界面
                                go_Context.AddParts(strName, strNumber, strFamily)
                                frmAllContext.GetInstance().AddPart(strName, strNumber, strFamily)
                            Next
                        End If
                    End If
                Next
            End If

            Return True

        Catch ex As Exception
            MainErrorHandler("mdlMain.RecvProducts", $"{ex.Message}-{ex.HResult}")
            Return False
        End Try
    End Function
    '=======================================================
    'Routine: mdlMain.RecvRunTypes(o)
    'Purpose: This requests the run types assinged to this
    '         NextCap PC.
    '
    'Globals:None
    '
    'Input: oSupervisor - A reference to the Supevisor object.
    '
    'Return: Boolean - true if the retrieval resulted in no
    '        errors.
    '
    'Modifications:
    '   11-23-1998 As written for Pass1.5
    '
    '
    '=======================================================
    Public Function RecvRunTypes(ByRef oSupervisor As clsSupervisor) As Boolean
        Dim dtRunTypes As DataTable
        Dim stationKey As clsStationKey

        Try
            If oSupervisor IsNot Nothing Then
                ' 遍历所有站点键
                For lngItem As Integer = 0 To gcol_StationKeys.Count - 1
                    stationKey = gcol_StationKeys(lngItem)

                    ' 获取运行类型数据（假设返回DataTable）
                    dtRunTypes = oSupervisor.GetRunTypes(
                        stationKey.strLineType,
                        stationKey.nLineNumber,
                        stationKey.strSource
                    )

                    ' 验证DataTable有效性
                    If dtRunTypes IsNot Nothing AndAlso dtRunTypes.Rows.Count > 0 Then
                        ' 检查是否为错误状态
                        If Convert.ToString(dtRunTypes.Rows(0)(0)) = cServerError Then
                            MainErrorHandler("mdlMain.RecvRunTypes", Convert.ToString(dtRunTypes.Rows(0)(1)))
                        Else
                            ' 遍历DataTable的行
                            For Each row As DataRow In dtRunTypes.Rows
                                ' 获取运行类型数据（假设第一列包含数据）
                                Dim strData As String = Convert.ToString(row(0))

                                ' 更新站点键的运行类型（使用带参数的属性）
                                stationKey.RunTypes(0) = strData

                                ' 添加到界面显示
                                frmAllContext.GetInstance().AddRunType(strData)
                            Next
                        End If
                    End If
                Next
            End If

            Return True  ' 若无错误，返回True

        Catch ex As Exception
            MainErrorHandler("mdlMain.RecvRunTypes", $"{ex.Message}-{ex.HResult}")
            Return False
        End Try
    End Function
    '=======================================================
    'Routine: mdlMain.RecvShifts(o)
    'Purpose: This requests the shifts assinged to this
    '         NextCap PC.
    '
    'Globals:None
    '
    'Input: oSupervisor - A reference to the Supevisor object.
    '
    'Return: Boolean - true if the retrieval resulted in no
    '        errors.
    '
    'Modifications:
    '   11-18-1998 As written for Pass1.4
    '
    '
    '=======================================================
    Public Function RecvShifts(ByRef oSupervisor As clsSupervisor) As Boolean
        Dim dtShifts As DataTable
        Dim stationKey As clsStationKey

        Try
            If oSupervisor IsNot Nothing Then
                ' 遍历所有站点键
                For lngItem As Integer = 0 To gcol_StationKeys.Count - 1
                    stationKey = gcol_StationKeys(lngItem)

                    ' 获取班次数据（假设返回 DataTable）
                    dtShifts = oSupervisor.GetShifts(
                        stationKey.strLineType,
                        stationKey.nLineNumber,
                        stationKey.strSource
                    )

                    ' 验证 DataTable 有效性
                    If dtShifts IsNot Nothing AndAlso dtShifts.Rows.Count > 0 Then
                        ' 检查是否为错误状态
                        If Convert.ToString(dtShifts.Rows(0)(0)) = cServerError Then
                            MainErrorHandler("mdlMain.RecvShifts", Convert.ToString(dtShifts.Rows(0)(1)))
                        Else
                            ' 遍历 DataTable 的行
                            For Each row As DataRow In dtShifts.Rows
                                ' 获取班次数据（假设第一列包含数据）
                                Dim strData As String = Convert.ToString(row(0))

                                ' 添加到界面显示
                                frmAllContext.GetInstance().AddShifts(strData)
                            Next
                        End If
                    End If
                Next
            End If

            Return True  ' 若无错误，返回 True

        Catch ex As Exception
            MainErrorHandler("mdlMain.RecvShifts", $"{ex.Message}-{ex.HResult}")
            Return False
        End Try
    End Function
    '=======================================================
    'Routine: mdlMain.RecvLotComments(o)
    'Purpose: This requests the Lot Comments assinged to this
    '         NextCap PC.
    '
    'Globals:None
    '
    'Input: oSupervisor - A reference to the Supevisor object.
    '
    'Return: Boolean - true if the retrieval resulted in no
    '        errors.
    '
    'Modifications:
    '   11-18-1998 As written for Pass1.4
    '
    '
    '=======================================================
    Public Function RecvLotComments(ByRef oSupervisor As clsSupervisor) As Boolean
        Dim dtLotComments As DataTable
        Dim stationKey As clsStationKey

        Try
            If oSupervisor IsNot Nothing Then
                ' 遍历所有站点键（索引从 0 开始）
                For lngItem As Integer = 0 To gcol_StationKeys.Count - 1
                    stationKey = gcol_StationKeys(lngItem)

                    ' 获取批次评论数据（假设返回 DataTable）
                    dtLotComments = oSupervisor.GetLotComments(
                        stationKey.strLineType,
                        stationKey.nLineNumber,
                        stationKey.strSource
                    )

                    ' 验证 DataTable 有效性
                    If dtLotComments IsNot Nothing AndAlso dtLotComments.Rows.Count > 0 Then
                        ' 检查错误状态（假设第一列存储状态码）
                        If Convert.ToString(dtLotComments.Rows(0)(0)) = cServerError Then
                            MainErrorHandler("mdlMain.RecvLotComments", Convert.ToString(dtLotComments.Rows(0)(1)))
                        Else
                            ' 遍历数据行
                            For Each row As DataRow In dtLotComments.Rows
                                ' 获取评论数据（假设第一列是评论内容）
                                Dim strData As String = Convert.ToString(row(0))

                                ' 更新站点键的批次评论（假设 LotComments 是带参数的属性）
                                stationKey.LotComments(0) = strData
                            Next
                        End If
                    End If
                Next
            End If

            Return True

        Catch ex As Exception
            MainErrorHandler("mdlMain.RecvLotComments", $"{ex.Message}-{ex.HResult}")
            Return False
        End Try
    End Function
    '=======================================================
    'Routine: mdlMain.RecvTasks(o)
    'Purpose: This requests the tasks assinged to this
    '         NextCap PC. These 'tasks' are the items that
    '         populate the User task functions.
    '
    'Globals:None
    '
    'Input: oSupervisor - A reference to the Supevisor object.
    '
    'Return: Boolean - true if the retrieval resulted in no
    '        errors.
    'Tested:
    '
    'Modifications:
    '   01-14-1999 As written for Pass1.6
    '
    '
    '=======================================================
    Public Function RecvTasks(ByRef oSupervisor As clsSupervisor) As Boolean
        Dim dtTasks As DataTable
        Dim stationKey As clsStationKey

        Try
            If oSupervisor IsNot Nothing Then
                ' 遍历所有站点键
                For lngItem As Integer = 0 To gcol_StationKeys.Count - 1
                    stationKey = gcol_StationKeys(lngItem)

                    ' 获取任务数据（返回 DataTable）
                    dtTasks = oSupervisor.GetTasks(
                    stationKey.strLineType,
                    stationKey.nLineNumber,
                    stationKey.strSource
                )

                    ' 验证 DataTable 有效性
                    If dtTasks IsNot Nothing AndAlso dtTasks.Rows.Count > 0 Then
                        ' 检查错误状态（假设第一列是状态码）
                        If dtTasks.Rows(0).Item(0).ToString() = cServerError Then
                            MainErrorHandler("mdlMain.RecvTasks", dtTasks.Rows(0).Item(1).ToString())
                        Else
                            ' 遍历 DataTable 的行（每行对应一个任务）
                            For Each row As DataRow In dtTasks.Rows
                                ' 提取任务参数（按列名或索引访问，需匹配实际表结构）
                                Dim taskName As String = row("TaskName").ToString()       ' 任务名（假设列名）
                                Dim taskType As String = row("TaskType").ToString()       ' 任务类型
                                Dim icon As String = row("Icon").ToString()               ' 图标（原代码对应第5参数，需确认列顺序）
                                Dim programCode As String = row("ProgramCode").ToString() ' 程序代码
                                Dim showInToolbar As Boolean = Convert.ToBoolean(row("ShowInToolbar")) ' 显示状态
                                Dim csbStatus As String = row("CSBStatus").ToString()     ' CSB状态

                                ' 向集合添加任务（假设 gcol_LocalApps 的 Add 方法参数顺序匹配）
                                gcol_LocalApps.Add(
                                taskName,
                                taskType,
                                icon,
                                programCode,
                                showInToolbar,
                                csbStatus
                            )
                            Next
                        End If
                    End If
                Next
            End If

            Return True

        Catch ex As Exception
            MainErrorHandler("mdlMain.RecvTasks()", $"{ex.Message}-{ex.HResult}")
            Return False
        End Try
    End Function
    '=======================================================
    'Routine: mdlMain.RecvQualityStates(o)
    'Purpose: This requests the Quality States assinged to this
    '         NextCap PC.
    '
    'Globals:None
    '
    'Input: oSupervisor - A reference to the Supevisor object.
    '
    'Return: Boolean - true if the retrieval resulted in no
    '        errors.
    '
    'Modifications:
    '   03-17-1999 As written for Pass1.7
    '
    '
    '=======================================================
    Public Function RecvQualityStates(ByRef oSupervisor As clsSupervisor) As Boolean
        Dim dtQualityStates As DataTable
        Dim stationKey As clsStationKey

        Try
            If oSupervisor IsNot Nothing Then
                ' 遍历所有站点键（索引从 0 开始）
                For lngItem As Integer = 0 To gcol_StationKeys.Count - 1
                    stationKey = gcol_StationKeys(lngItem)

                    ' 获取质量状态数据（假设返回 DataTable）
                    dtQualityStates = oSupervisor.GetQualityStates(
                        stationKey.strLineType,
                        stationKey.nLineNumber,
                        stationKey.strSource
                    )

                    ' 验证 DataTable 有效性
                    If dtQualityStates IsNot Nothing AndAlso dtQualityStates.Rows.Count > 0 Then
                        ' 检查错误状态（假设第一列是状态码）
                        If Convert.ToString(dtQualityStates.Rows(0)(0)) = cServerError Then
                            MainErrorHandler("mdlMain.RecvQualityStates", Convert.ToString(dtQualityStates.Rows(0)(1)))
                        Else
                            ' 遍历 DataTable 的行
                            For Each row As DataRow In dtQualityStates.Rows
                                ' 提取质量状态参数（假设列顺序：状态名称、描述、代码）
                                Dim stateName As String = Convert.ToString(row(0))
                                Dim stateDesc As String = Convert.ToString(row(1))
                                Dim stateCode As String = Convert.ToString(row(2))

                                ' 调用模块方法创建质量状态对象
                                mdlChart.CreateQualityStatusObject(stateName, stateDesc, stateCode)
                            Next
                        End If
                    End If
                Next
            End If

            Return True

        Catch ex As Exception
            MainErrorHandler("mdlMain.RecvQualityStates()", $"{ex.Message}-{ex.HResult}")
            Return False
        End Try
    End Function

    '=======================================================
    'Routine: mdlMain.RecvStations(o)
    'Purpose: This requests the stations assinged to this
    '         NextCap PC.
    '
    'Globals:None
    '
    'Input: oSupervisor - A reference to the Supevisor object.
    '
    'Return: Boolean - true if the retrieval resulted in no
    '        errors.
    '
    'Modifications:
    '   11-18-1998 As written for Pass1.4
    '
    '
    '=======================================================
    Public Function RecvStations(ByRef oSupervisor As clsSupervisor) As Boolean
        Dim dtStations As DataTable
        Dim row As DataRow

        Try
            If oSupervisor IsNot Nothing Then
                ' 获取DataTable对象
                dtStations = oSupervisor.GetStations(mdlTools.CurrentMachineName)

                ' 验证DataTable有效性
                If dtStations IsNot Nothing AndAlso dtStations.Rows.Count > 0 Then
                    Dim value As Object = dtStations.Rows(0)(0)
                    ' 检查是否为错误状态（假设第一列第一个值是状态码）
                    'If Convert.ToString(dtStations.Columns.Contains("StatusCode") & dtStations.Rows(0)("StatusCode")) = cServerError Then
                    If Not IsDBNull(value) AndAlso value.ToString() = "ERROR" Then
                        MainErrorHandler("mdlMain.RecvStations", Convert.ToString(dtStations.Rows(0)("ErrorMessage")))
                        Return False
                    Else
                        ' 遍历DataTable的行
                        For Each row In dtStations.Rows
                            ' 添加站点键（根据实际列名调整索引或列名）
                            gcol_StationKeys.Add(
                            Convert.ToString(row(0)),
                            Convert.ToString(row(2)),
                            Convert.ToInt32(row(1))
                        )
                        Next

                        Return dtStations.Rows.Count > 0
                    End If
                End If
            End If

            Return False

        Catch ex As Exception
            MainErrorHandler("mdlMain.RecvStations", $"{ex.Message}-{ex.HResult}")
            Return False
        End Try
    End Function

    Private Sub GenerateLotManager()
        '/*Loop through and generate a LotManager object for each
        '/*Station-Key that we have.

        For nItem As Integer = 0 To gcol_StationKeys.Count - 1  ' 注意：VB.NET 集合索引从 0 开始
            '/*Set a reference to the selected Station-Key
            Dim stationKey = gcol_StationKeys(nItem)
            '/*Attempt to generate the instance
            mdlLotManager.GenerateLotManager(stationKey.strLineType, stationKey.nLineNumber, stationKey.strSource)
        Next nItem

        '/*Provided we have a LotManager available
        '/*set the LotManager Interface
        If gcol_LotManagers.Count > 0 Then  ' 简化判断逻辑
            '/*Make sure the client interface is on line
            frmLotManager.Refresh()

            '/*Call this module to that will seek and set
            '/*the server.LotManager reference in the
            '/*user interface for managing Lots.
            '/*---- WARNING!!!! --------------------------
            '/*This is making a big assumption that
            '/*this object in gcol_LotManagers will actually
            '/*be a frmLotManagerSink. It may be necessary
            '/*at some point to use a TypeOf() check on the
            '/*object before attempting to reference it.
            Dim firstLotManager = gcol_LotManagers(0)  ' Setting the Active Lot Manager
            mdlLotManager.SetActiveLotManager(
            firstLotManager.LotManager.LineType,
            firstLotManager.LotManager.LineNumber,
            firstLotManager.LotManager.Source)

            'mdlLotManager.SetActiveLotManager(gcol_LotManagers(1).LotManager.LineType,
            '                              gcol_LotManagers(1).LotManager.LineNumber,
            '                              gcol_LotManagers(1).LotManager.Source)

        End If
    End Sub

    Private Function EstablishConnection() As Boolean
        Try
            LogEvent("Connect to the server object specified in the system settings: " & go_clsSystemSettings.strBusinessServerUNC)

            ' Connect to the server object specified in the system settings
            ' NextCapServer Is the assembly name (i.e., the name of the compiled DLL), And the class clsBoot Is declared Like Public Class clsBoot inside the namespace  NextCapServer, And the DLL Is already referenced in your VB.NET project
            go_BusinessServer = mdlTools.CreateAutoObject("ncapsrv.clsBoot, ncapsrv", go_clsSystemSettings.strBusinessServerUNC)

            LogEvent("Refresh the Object Event trap (e.g. COM sink); this triggers the Form_Load that handles the initialization of the clsBoot object.")
            frmServerSink.Refresh()
            'Create the sink form (which will handle events)
            serverSinkForm = New frmServerSink()


            'Kick off the  startup logic in clsBoot which causes ReadyToGo to fire
            LogEvent("Trigger the server to do its WakeUp routine")
            go_BusinessServer.WakeUp(mdlTools.CurrentMachineName)

            LogEvent("If the object has been set return true")
            If go_Supervisor IsNot Nothing Then
                Return True
            End If

        Catch ex As Exception
            Console.WriteLine("mdlMain - EstablishConnection Failed")
            LogEvent("EstablishConnection_Err: " & ex.Message)
            MainErrorHandler("mdlMain.EstablishConnection", ex.Message)
        End Try

        Return False
    End Function

    Public Sub GenerateRunTimeParameters()
        Const cstrNone As String = "Undefined"
        Dim oRT As NextCapServer.clsRuntime = Nothing
        Dim rawValue As Object
        Dim strTemp As String = String.Empty
        Dim lngTemp As Integer

        Try
            ' Get a reference to the RunTime object
            oRT = go_BusinessServer.GetRTObject()

            ' Place the items into the SystemSettings object
            With go_clsSystemSettings
                '----------------------------------------------
                ' Base chart colors
                '----------------------------------------------
                rawValue = oRT("ColorOfBad")
                If rawValue IsNot Nothing AndAlso IsNumeric(rawValue) Then
                    .nColorOfBad = Convert.ToInt32(rawValue)
                End If

                rawValue = oRT("ColorOfGood")
                If rawValue IsNot Nothing AndAlso IsNumeric(rawValue) Then
                    .nColorOfGood = Convert.ToInt32(rawValue)
                End If

                '-----------------------------------------------
                ' Good/Bad Legends
                '-----------------------------------------------
                rawValue = oRT("LegendGood")
                strTemp = If(rawValue IsNot Nothing, rawValue.ToString().Trim(), String.Empty)
                If strTemp.Length > 0 AndAlso Not strTemp.Equals(cstrNone, StringComparison.OrdinalIgnoreCase) Then
                    .strLegendGood = strTemp
                Else
                    .strLegendGood = "Good"
                End If

                rawValue = oRT("LegendBad")
                strTemp = If(rawValue IsNot Nothing, rawValue.ToString().Trim(), String.Empty)
                If strTemp.Length > 0 AndAlso Not strTemp.Equals(cstrNone, StringComparison.OrdinalIgnoreCase) Then
                    .strLegendBad = strTemp
                Else
                    .strLegendBad = "Bad"
                End If

                '-------------------------------------------------
                ' Configurations for chart scrolling
                '-------------------------------------------------
                rawValue = oRT("MaxGroups")
                If rawValue IsNot Nothing AndAlso Integer.TryParse(rawValue.ToString(), lngTemp) Then
                    .lngMaxGroups = If(lngTemp < 10, 10, lngTemp)
                Else
                    .lngMaxGroups = 10
                End If

                rawValue = oRT("MaxVisibleGroups")
                If rawValue IsNot Nothing AndAlso Integer.TryParse(rawValue.ToString(), lngTemp) Then
                    .lngMaxVisibleGroups = If(lngTemp < 10, 10, lngTemp)
                Else
                    .lngMaxVisibleGroups = 10
                End If

                '----------------------------------------------
                ' Pen Id source and accessories
                '----------------------------------------------
                rawValue = oRT("PenIdSource")
                strTemp = If(rawValue IsNot Nothing, rawValue.ToString().Trim(), String.Empty)
                If strTemp.Length > 0 AndAlso Not strTemp.Equals(cstrNone, StringComparison.OrdinalIgnoreCase) Then
                    .strPenIdSource = strTemp

                    If strTemp.Equals("ACTIVEX", StringComparison.OrdinalIgnoreCase) Then
                        rawValue = oRT("ActiveXidServer")
                        strTemp = If(rawValue IsNot Nothing, rawValue.ToString().Trim(), String.Empty)
                        If strTemp.Length > 0 Then
                            .strActiveXserver_ID = strTemp
                            ' Generate the ActiveX server object
                            'SetActiveXIDserver(.strActiveXserver_ID)
                        Else
                            .strPenIdSource = "None"
                        End If

                    ElseIf strTemp.Equals("POLLED", StringComparison.OrdinalIgnoreCase) Then
                        rawValue = oRT("PenIdCaptureProgram")
                        strTemp = If(rawValue IsNot Nothing, rawValue.ToString().Trim(), String.Empty)
                        If strTemp.Length > 0 Then
                            .strPenIdCaptureProgram = strTemp

                            rawValue = oRT("PenIdTimeOut")
                            If rawValue IsNot Nothing AndAlso Integer.TryParse(rawValue.ToString(), lngTemp) Then
                                .lngPenIdTimeOut = If(lngTemp < 1000, 1000, lngTemp)
                            Else
                                .lngPenIdTimeOut = 1000
                            End If
                        Else
                            .strPenIdSource = "None"
                        End If

                    ElseIf strTemp.Equals("SCANNER", StringComparison.OrdinalIgnoreCase) Then
                        .bBarcodeScannerMode = True
                        frmNextCap.InputMode = strTemp
                    End If
                End If

                '------------------------------------------------
                ' PenIdMemory - maintain unique pen list
                '------------------------------------------------
                rawValue = oRT("PenIdMemory")
                If rawValue IsNot Nothing AndAlso Integer.TryParse(rawValue.ToString(), lngTemp) Then
                    If lngTemp > 0 Then
                        .nPenIdMemory = lngTemp
                        .bPenIdUnique = True
                    End If
                End If

                '-------------------------------------------------
                ' Get the required pen id flags
                '-------------------------------------------------
                rawValue = oRT("GoodPenIdRequired")
                .bGoodPenIdRequired = (rawValue IsNot Nothing AndAlso TypeOf rawValue Is Boolean AndAlso Convert.ToBoolean(rawValue))

                rawValue = oRT("BadPenIdRequired")
                .bBadPenIdRequired = (rawValue IsNot Nothing AndAlso TypeOf rawValue Is Boolean AndAlso Convert.ToBoolean(rawValue))

                '-------------------------------------------------
                ' Is the Number Of Good pens option enabled
                '-------------------------------------------------
                rawValue = oRT("AskForNumberOfGoodPens")
                .bAskForNumberOfGoodPens = (rawValue IsNot Nothing AndAlso TypeOf rawValue Is Boolean AndAlso Convert.ToBoolean(rawValue))

                '-------------------------------------------------
                ' Good/Bad pen feedback
                '-------------------------------------------------
                rawValue = oRT("GoodPenVisualFeedback")
                strTemp = If(rawValue IsNot Nothing, rawValue.ToString().Trim(), String.Empty)
                If System.IO.File.Exists(strTemp) Then
                    .bGoodPenFeedBack = True
                    .strGoodPenVisualFeedBack = strTemp
                    '  frmGoodPen.PictureBox1.Image = Image.FromFile(strTemp)
                End If

                rawValue = oRT("GoodPenAudioFeedback")
                strTemp = If(rawValue IsNot Nothing, rawValue.ToString().Trim(), String.Empty)
                If System.IO.File.Exists(strTemp) Then
                    .bGoodPenFeedBack = True
                    .strGoodPenAudioFeedBack = strTemp
                    'frmGoodPen.MMControl1.FileName = strTemp
                End If

                rawValue = oRT("BadPenVisualFeedback")
                strTemp = If(rawValue IsNot Nothing, rawValue.ToString().Trim(), String.Empty)
                If System.IO.File.Exists(strTemp) Then
                    .bBadPenFeedBack = True
                    .strBadPenVisualFeedBack = strTemp
                    ' frmBadPen.MMControl1.FileName = strTemp
                End If

                rawValue = oRT("BadPenAudioFeedback")
                strTemp = If(rawValue IsNot Nothing, rawValue.ToString().Trim(), String.Empty)
                If System.IO.File.Exists(strTemp) Then
                    .bBadPenFeedBack = True
                    .strBadPenAudioFeedBack = strTemp
                    'frmBadPen.PictureBox1.Image = Image.FromFile(strTemp)
                End If

                '-------------------------------------------------
                ' Production Date adjust function in use
                '-------------------------------------------------
                rawValue = oRT("ProductionDateAdjust")
                .bProductionDateAdjust = (rawValue IsNot Nothing AndAlso TypeOf rawValue Is Boolean AndAlso Convert.ToBoolean(rawValue))

                '--------------------------------------------------
                ' The status icon location
                '--------------------------------------------------
                rawValue = oRT("StatusIconLocation")
                strTemp = If(rawValue IsNot Nothing, rawValue.ToString().Trim(), String.Empty)
                .strStatusIconLocation = strTemp

                '--------------------------------------------------
                ' Client to Server heart beat time limit
                '--------------------------------------------------
                rawValue = 1 'oRT("ServerTimeOut")
                frmServerSink.Refresh()
                frmServerSink.ServerTimeOut = If(rawValue IsNot Nothing AndAlso Integer.TryParse(rawValue.ToString(), lngTemp), lngTemp, 0)

                '--------------------------------------------------
                ' Flag for whether Suspended list is visible on LotManager
                '--------------------------------------------------
                rawValue = oRT("DisableSuspended")
                frmLotManager.bSuspendedUsed = Not (rawValue IsNot Nothing AndAlso TypeOf rawValue Is Boolean AndAlso Convert.ToBoolean(rawValue))

                '--------------------------------------------------
                ' Flag for whether Closed list is visible on LotManager
                '--------------------------------------------------
                rawValue = oRT("DisableClosed")
                frmLotManager.bClosedUsed = Not (rawValue IsNot Nothing AndAlso TypeOf rawValue Is Boolean AndAlso Convert.ToBoolean(rawValue))

                '--------------------------------------------------
                ' Flag for whether SampleCount resets on product change
                '--------------------------------------------------
                rawValue = oRT("ResetCounterWithProduct")

                .bResetCountOnProductChange = (rawValue IsNot Nothing AndAlso TypeOf rawValue Is Boolean AndAlso Convert.ToBoolean(rawValue))

                '--------------------------------------------------
                ' Flag for whether to AutoCreate Lot
                '--------------------------------------------------
                rawValue = oRT("AutoCreateLot")
                'frmLotManager.AutoCreateLot = (rawValue IsNot Nothing  AndAlso TypeOf rawValue Is Boolean AndAlso Convert.ToBoolean(rawValue))

                '--------------------------------------------------
                ' Get the login parameter for the Security object
                '--------------------------------------------------
                rawValue = oRT("RequireLogIn")
                go_Security.bLoginRequired = (rawValue IsNot Nothing AndAlso TypeOf rawValue Is Boolean AndAlso Convert.ToBoolean(rawValue))

                '--------------------------------------------------
                ' v2.0.1: Default disposition on defect entry screen
                ' Must go here to prevent race condition on frmAllContext.Load()
                '--------------------------------------------------
                rawValue = oRT("NoDefaultRunType")
                strTemp = If(rawValue IsNot Nothing, rawValue.ToString().Trim(), String.Empty)
                .bUseDefaultRunType = Not (If(String.IsNullOrEmpty(strTemp), False, True))

                '--------------------------------------------------
                ' Shift function allowed?
                '--------------------------------------------------
                rawValue = oRT("ShiftFunction")
                .strShiftFunction = If(rawValue IsNot Nothing, rawValue.ToString().Trim(), String.Empty)

                '--------------------------------------------------
                ' Shift function can log out current user?
                '--------------------------------------------------
                rawValue = oRT("LogOutAtShift")
                If go_Security.bLoginRequired Then
                    .bLogOutAtShift = (rawValue IsNot Nothing AndAlso Convert.ToBoolean(rawValue))
                End If

                '--------------------------------------------------
                ' Material mode
                '--------------------------------------------------
                rawValue = oRT("MaterialMode")
                .strMaterialMode = If(rawValue IsNot Nothing, rawValue.ToString().Trim(), String.Empty)

                '--------------------------------------------------
                ' Client SAX engine debug mode?
                '--------------------------------------------------
                rawValue = oRT("ClientCSBDebug")
                .bCSBDebug = (rawValue IsNot Nothing AndAlso TypeOf rawValue Is Boolean AndAlso Convert.ToBoolean(rawValue))

                '--------------------------------------------------
                ' Image Sync debug?
                '--------------------------------------------------
                rawValue = oRT("ImageSyncDebug")
                .bImageDebug = (rawValue IsNot Nothing AndAlso TypeOf rawValue Is Boolean AndAlso Convert.ToBoolean(rawValue))

                '--------------------------------------------------
                ' Only Primary client allowed to AutoCreate Lots?
                '--------------------------------------------------
                rawValue = oRT("PrimaryAutoCreateOnly")
                .bPrimaryCreateLotOnly = (rawValue IsNot Nothing AndAlso TypeOf rawValue Is Boolean AndAlso Convert.ToBoolean(rawValue))

                '--------------------------------------------------
                ' Recovery Step enabled on # of pens
                '--------------------------------------------------
                rawValue = oRT("EnableRecoveryStep")
                .bEnableRecoveryStep = (rawValue IsNot Nothing AndAlso TypeOf rawValue Is Boolean AndAlso Convert.ToBoolean(rawValue))

                '--------------------------------------------------
                ' Default disposition on defect entry screen again?
                '--------------------------------------------------
                rawValue = oRT("DispositionDefault")
                strTemp = If(rawValue IsNot Nothing, rawValue.ToString().Trim(), String.Empty)
                .strDispositionDefault = strTemp

                '--------------------------------------------------
                ' Recovery Step enabled on # of pens again?
                '--------------------------------------------------
                rawValue = oRT("UseLotPointer")
                .bUseLotPointer = (rawValue IsNot Nothing AndAlso TypeOf rawValue Is Boolean AndAlso Convert.ToBoolean(rawValue))

                rawValue = oRT("ClientEnableCSB-IDE")
                .bEnableCSBIDE = (rawValue IsNot Nothing AndAlso TypeOf rawValue Is Boolean AndAlso Convert.ToBoolean(rawValue))

                rawValue = oRT("ClientHelpURL")
                .strHelpURL = If(rawValue IsNot Nothing, rawValue.ToString().Trim(), String.Empty)

                ' -- Setting for multiple bad pen count on defect editor
                rawValue = oRT("EnableBadPenCount")
                .bBadPenCountEnabled = (rawValue IsNot Nothing AndAlso TypeOf rawValue Is Boolean AndAlso Convert.ToBoolean(rawValue))
            End With

        Catch ex As Exception
            ' Handle any unexpected error
            Console.WriteLine($"Error in GenerateRunTimeParameters: {ex.Message}")
        End Try
    End Sub

    Private Sub GenerateSystemObject()
        Dim strUNC As String

        Try
            ' 生成系统设置对象
            go_clsSystemSettings = New clsSystemSettings()

            ' 获取业务服务器位置（从注册表）
            'strUNC = My.Computer.Registry.GetValue(cstrRegName, "BusinessServer", Nothing)

            ' 如果未找到注册表项或值为空，则使用当前计算机名
            If String.IsNullOrEmpty(strUNC) Then
                strUNC = CurrentMachineName()
            End If

            ' 设置业务服务器UNC路径
            go_clsSystemSettings.strBusinessServerUNC = strUNC

            ' 设置当前计算机的UNC名称
            go_clsSystemSettings.strUNC_Name = CurrentMachineName()

        Catch ex As Exception
            ' 处理异常（例如记录日志）
            Debug.WriteLine("生成系统对象时出错: " & ex.Message)
            Throw
        End Try
    End Sub

    Private Sub GenerateLocalAppsObject()
        '/*Generate the global instance of colLocalApps
        gcol_LocalApps = New colLocalApps
    End Sub

    Public Function ShutDown(ByVal nUnloadMode As Integer) As Boolean
        Dim frmApp As Form
        Const strMsg As String = "Are you sure you want to Exit NextCap?"
        Dim lngResult As DialogResult

        ' 忽略错误继续执行（VB.NET 中不推荐这样使用，但为了接近原逻辑）
        Try
            '/*Fliter out the more extreme shutdown situation
            If nUnloadMode = 3 Then
                '/*Query for shutdown before continuing
                lngResult = MessageBox.Show(strMsg, "Exit", MessageBoxButtons.OKCancel, MessageBoxIcon.Question)
            Else
                lngResult = DialogResult.OK
            End If
            '/*Test the outcome of the question
            If lngResult = DialogResult.OK Then
                '/*Set the shutdown flag to true
                'ShutDownFlag = True
                '/*Save the window states for frmNextCap
                'SaveWindowPos(frmNextCap)
                'SaveWindowPos(frmDefectEditor)

                '/*Shutdown any active timers
                'frmServerSink.tmrDBmaint.Enabled = False
                'frmServerSink.tmrHeartBeat.Enabled = False
                'frmEndOfLot.tmrLotSequnce.Enabled = False

                '/*Inform NextCap that this is not an optional shutdown
                'frmNextCap.m_bForcedShutDown = True
                '/*Unload any active forms
                'For Each frmApp In Application.OpenForms
                '    frmApp.Close()
                'Next frmApp
                ShutDown = True

                '/*Unload the system objects
                go_CAPmain = Nothing
                go_clsSystemSettings = Nothing
                go_ChartConfigs = Nothing
                'gcol_LocalApps = Nothing
                'gcol_SAXhandlers = Nothing
                gcol_QualityStatus = Nothing
                go_BusinessServer = Nothing
                go_Supervisor = Nothing
                gcol_StationKeys = Nothing
                gcol_ChartData = Nothing
                go_ActiveLotManager = Nothing
                gcol_LotManagers = Nothing
                go_Context = Nothing
                'go_ActiveXIDserver = Nothing
                '/*Context object
                go_Context = Nothing
                'go_SharedIdInput = Nothing
            End If
        Catch ex As Exception
            ' 这里可以添加异常处理逻辑
        End Try

        Return ShutDown
    End Function
End Module
