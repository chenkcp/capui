Imports System.DirectoryServices.ActiveDirectory
Imports System.Globalization
Imports NextcapServer = ncapsrv
Module mdlLotManager
    Private Const cServerError As String = "ERROR"
    Public ActiveLotManagerForm As frmLotManager
    '
    '=======================================================
    'Routine: mdlLotManager.TransactionFailure()
    'Purpose: This tests for errors related to the Current
    'LotId in the global COntext to see why transactions
    'are failing.
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
    '   02-02-1998 As written for Pass1.7
    '
    '
    '=======================================================
    Public Sub TransactionFailure(Optional ByRef strCaller As String = "Unknown Module")
        Dim strMsg As String

        If go_Context IsNot Nothing Then
            If String.IsNullOrEmpty(go_Context.GroupId) Then
                ' Send an alert message to the user
                strMsg = "The Context has no open Lot." & vbCrLf &
                         "1) Create a new Lot with the Lot Manager." & vbCrLf & vbCrLf &
                         "mdlLotManager.TransactionFailure() Error #1"
            Else
                ' Send an alert message to the user
                strMsg = "Either an unknown error has occurred or the current Context has no open Lot. To Fix:" & vbCrLf &
                         "1) Create a new Lot, there is no open Lot.  Or" & vbCrLf &
                         "2) Shutdown and restart the client." & vbCrLf & vbCrLf &
                         "mdlLotManager.TransactionFailure() Error #2"
            End If
            'frmMessage.GenerateMessage(strMsg)
            MessageBox.Show(strMsg)
        End If
    End Sub
    '
    '=======================================================
    'Routine: frmLotManager.RecvComments(o)
    'Purpose: This queries the Lot Manager to get a list
    '         of the current Closed lots.
    '
    'Globals:None
    '
    'Input: oLotManager - The lot manager connection.
    '
    '       strGroupId - The group id to request comments for.
    '
    '       dtBirth - The date of the group.
    '
    '       vrtArray - Byref pass back of the result.
    '
    'Return: Boolean - Whhether the request was succesful.
    '
    'Modifications:
    '   02-11-1999 As written for Pass1.7
    '
    '=======================================================
    Public Function RecvComments(oLotManager As NextcapServer.clsLotManager, ByVal strGroupId As String, ByVal dtBirth As Date, ByRef vrtArray As DataTable) As Boolean
        ' 确保有有效的LotManager对象
        If oLotManager Is Nothing Then
            Return False ' 返回False表示操作失败
        Else
            Try
                ' 请求数据
                vrtArray = oLotManager.GetComments(strGroupId, dtBirth)

                ' 检查DataTable是否包含数据
                If vrtArray Is Nothing OrElse vrtArray.Rows.Count = 0 OrElse vrtArray.Columns.Count = 0 Then
                    ReportServerError("mdlLotManager.RecvComments", "返回的数据表为空")
                    Return False
                End If

                ' 检查第一行第一列是否为错误标识
                If Convert.ToString(vrtArray.Rows(0)(0)) = cServerError Then
                    ' 通过集中处理器报告错误
                    Dim errorMessage As String = Convert.ToString(vrtArray.Rows(0)(1))
                    ReportServerError("mdlLotManager.RecvComments", errorMessage)
                    Return False
                Else
                    ' 数据正常，返回True
                    Return True
                End If
            Catch ex As Exception
                ' 处理异常情况
                ReportServerError("mdlLotManager.RecvComments", "获取评论时发生异常: " & ex.Message)
                Return False
            End Try
        End If
    End Function

    '/*Centralized server error
    Private Sub ReportServerError(ByRef strCaller As String, ByRef strMsg As String)
        '/*Call main error handler
        MainErrorHandler(strCaller, strMsg, "Warning")
    End Sub

    '

    '
    '=======================================================
    'Routine: frmLotManager.RecvShipment(o)
    'Purpose: This queries the Lot Manager to get a list
    '         of the current Shipment data for the Group.
    '
    'Globals:None
    '
    'Input: oLotManager - The lot manager connection.
    '
    '       strGroupId - The group id to request comments for.
    '
    '       dtBirth - The date of the group.
    '
    '       vrtArray - Byref pass back of the result.
    '
    'Return: Boolean - Whhether the request was succesful.
    '
    'Tested: Hand tested 2-15-1999 C. Barker
    '
    'Modifications:
    '   02-15-1999 As written for Pass1.7
    '
    '=======================================================
    Public Function RecvShipment(ByVal oLotManager As NextcapServer.clsLotManager,
                             ByVal strGroupId As String,
                             ByVal dtBirth As Date,
                             ByRef vrtArray As DataTable) As Boolean
        ' 确保有有效的 LotManager 对象
        If oLotManager IsNot Nothing Then
            Try
                ' 请求数据（假设 GetShipmentData 返回 DataTable）
                vrtArray = oLotManager.GetShipmentData(strGroupId, dtBirth)

                ' 检查 DataTable 是否有效且包含数据
                If vrtArray IsNot Nothing AndAlso vrtArray.Rows.Count > 0 AndAlso vrtArray.Columns.Count > 0 Then
                    ' 检查第一行第一列是否为错误码
                    If Convert.ToString(vrtArray.Rows(0)(0)) = cServerError Then
                        ' 通过集中处理程序处理错误
                        ReportServerError("mdlLotManager.RecvShipment", Convert.ToString(vrtArray.Rows(0)(1)))
                    Else
                        Return True ' 数据有效
                    End If
                End If
            Catch ex As Exception
                ' 处理异常
                ReportServerError("mdlLotManager.RecvShipment", "获取数据时发生异常: " & ex.Message)
            End Try
        End If

        Return False ' 默认返回 False（未成功获取数据）
    End Function

    '
    '=======================================================
    'Routine: mdlMain.GenerateLotManager()
    'Purpose: This attaches to the LotManager and sets the
    'current global LotManager object. The Context for the
    'Open Lot is also set at this time.
    '
    'Globals: go_ActiveLotManager - The global LotManager object.
    '
    'Input:None
    '
    'Return: LotManager - A reference to a LotManager
    '        connection on the business server.
    '
    'Tested: hand tested 1-6-1999. Always returned the same
    '        lot manager regardless of the parameters passed.
    '
    'Modifications:
    '   01-06-1999 As written for Pass1.6
    '
    '
    '=======================================================
    Public Function GenerateLotManager(ByVal strLine As String, ByVal nLineNumber As Integer, ByVal strSource As String) As NextcapServer.clsLotManager
        Dim oTemporary As NextcapServer.clsLotManager
        Dim frmTemporary As frmLotManagerSink

        Try
            '/*Insure that the Supervisor object is initialized
            If Not (go_Supervisor Is Nothing) Then
                '/*Attach the Lot Manager handle
                oTemporary = go_Supervisor.GetLotManager(strLine, nLineNumber, strSource)
                '/*Make sure we received a valid reference
                If Not (oTemporary Is Nothing) Then
                    '/*Generate a new form
                    frmTemporary = New frmLotManagerSink()
                    ' Trigger the Load event manually
                    'frmTemporary.OnLoad(EventArgs.Empty)

                    '' Alternatively, call a custom initialization method
                    ''frmTemporary.Initialize()

                    '' Keep the form hidden
                    'frmLotManagerSink.Visible = False
                    '/*Now set the instance of the LotManager
                    frmTemporary.SetLotManager(oTemporary)
                    '/*Add the Form to the global collection
                    '/*of business server sinks so that we
                    '/*have an Index formed for searching
                    gcol_LotManagers.Add(frmTemporary)
                End If
            End If
        Catch ex As Exception
            'mdlMain.MainErrorHandler("mdlLotManager.GenerateLotManager()", ex.Message & " no=" & ex.HResult.ToString())
            MessageBox.Show("mdlLotManager.GenerateLotManager()," + ex.Message)
        Finally
            '/*Destroy the instance
            oTemporary = Nothing
        End Try

        Return oTemporary
    End Function

    '
    '=======================================================
    'Routine: mdlLotManager.SetActiveLotManager(str,n,str)
    'Purpose: This queries the LotManagerSink collection
    'and set the global Active LotManager.
    '
    'Globals:None
    '
    'Input: strLine - The name of the line type.
    '
    '       nLineNumber - The numeric id of the line.
    '
    '       strSource - The source type.
    '
    'Return: Boolean - True if an acceptable LotManager
    'object was found and connected to.
    '
    'Tested:
    '
    'Modifications:
    '   01-07-1998 As written for Pass1.6
    '
    '   01-13-1999 Added call to Set gcol_ChartData =
    '   oLotSink.ChartDataCol
    '
    '   07-06-1999 Added reference to the LotManagerSink
    '   to a global object. This is primarily to allow
    '   the SAX engine to access a LotManager's methods.
    '=======================================================
    Public Function SetActiveLotManager(strLine As String, nLineNumber As Integer, strSource As String) As Boolean
        '/*Loop through the collection of Lot Managers
        For Each oLotSink As frmLotManagerSink In gcol_LotManagers
            '/*If this has a positive station match set it as active
            If oLotSink.IsStation(strLine, nLineNumber, strSource) Then
                '/*This is the global instance of the LotManager
                go_ActiveLotManager = oLotSink.LotManager
                '/*Set a global reference to the LotManagerSink itself
                go_ActiveLotSink = oLotSink
                '/*Toggle the active flag on the global collection
                'gcol_ChartData.bActive = False

                '/*Switch the reference for the global ChartData group
                gcol_ChartData = oLotSink.ChartDataCol
                gcol_ChartData.bActive = True
                If gcol_ChartData IsNot Nothing And gcol_ChartData.Count > 0 Then ' InitializeData not run yet

                    '/*Remap the NextCap chart
                    mdlChart.SetChartActiveData()
                    '/*This is the user interface for managing Lots. It
                    '/*needs to have a reference to enable the different
                    '/*Lot type to stay sync'd with the server.
                    'frmLotManager.SetLotManager(go_ActiveLotManager, oLotSink)
                    ActiveLotManagerForm = New frmLotManager()
                    ActiveLotManagerForm.SetLotManager(go_ActiveLotManager, oLotSink)
                End If

                Return True
            End If

        Next
        Return False
    End Function
    '=======================================================
    'Routine: frmLotManager.RecvOpenLot(o)
    'Purpose: This queries the Lot Manager to get a list
    '         of the current Open lots.
    '
    'Globals:None
    '
    'Input: oLotManager - The lot manager connection.
    '
    'Return:None
    '
    'Modifications:
    '   12-03-1998 As written for Pass1.5
    '
    '   01-05-1999 Moved to mdlLotManager to facilitate
    '   running more then one instance of frmLotmanager.
    '
    '   02-03-1999 Switched to pass back vrtArray(0, 0) = ""
    '   for Empty Lot on Error branch from Server.
    '=======================================================
    Public Function RecvOpenLot(oLotManager As NextcapServer.clsLotManager, ByRef arr As DataTable) As Boolean

        ' Validate the manager
        If oLotManager Is Nothing Then
            Return False
        End If

        ' Retrieve the last created lot id, birthday as datatable from NextcapServer.clsLotManager's class memory
        arr = oLotManager.GetOpenLot()


        If arr Is Nothing OrElse arr.Rows.Count = 0 Then
            Return False ' no lot created yet
        End If
        ' Check for server error code in column 0
        If arr Is Nothing AndAlso arr.Rows(0)(0).ToString() = cServerError Then
            ReportServerError(
                "LotManagerModule.RecvOpenLot",
                vtoa(arr.Rows(1)(0))
            )


            Return False
        End If

        ' Success
        Return True
    End Function
    'Public Function RecvOpenLot(oLotManager As NextcapServer.clsLotManager, ByRef vrtArray As Object) As Boolean
    '    '/*insure that we have a valid LotManager
    '    If oLotManager Is Nothing Then
    '        '/*Do nothing for now
    '    Else
    '        '/*request the data
    '        vrtArray = oLotManager.GetOpenLot()

    '        '/*Check if the array is initialized and handle errors
    '        If vrtArray IsNot Nothing AndAlso TypeOf vrtArray Is Object(,) Then
    '            'Dim array2D As Object(,) = DirectCast(vrtArray, Object(,))

    '            '/*An array of dimension (0,0) indicates an error occured.
    '            If vrtArray(0, 0).ToString() = cServerError Then
    '                '/*Route the error through the centralized handler
    '                ReportServerError("mdlLotManager.RecvOpenLot", vtoa(vrtArray(0, 1).ToString()))

    '                '/*Set unknown
    '                ReDim vrtArray(1, 0)
    '                vrtArray(0, 0) = ""
    '                vrtArray(1, 0) = DateTime.Now
    '                RecvOpenLot = False
    '            Else
    '                RecvOpenLot = True
    '            End If
    '        Else
    '            '/*Handle unexpected array type or null
    '            RecvOpenLot = False
    '        End If
    '    End If
    'End Function
    '=======================================================
    'Routine: frmLotManager.RecvClosedLot(o)
    'Purpose: This queries the Lot Manager to get a list
    '         of the current Closed lots.
    '
    'Globals:None
    '
    'Input: oLotManager - The lot manager connection.
    '
    'Return:None
    '
    'Modifications:
    '   12-03-1998 As written for Pass1.5
    '
    '   01-05-1999 Moved to mdlLotManager to facilitate
    '   running more then one instance of frmLotmanager.
    '
    '
    '=======================================================
    Public Function RecvClosedLot(oLotManager As NextcapServer.clsLotManager, ByRef arr As DataTable) As Boolean
        ' Validate the manager
        If oLotManager Is Nothing Then
            Return False
        End If

        ' Retrieve the open‐lot table
        arr = oLotManager.GetClosedLots()

        ' If we didn’t get back a table or it has no rows, treat as failure
        ' Check for empty or error result
        If arr Is Nothing OrElse arr.Rows.Count = 0 Then
            Return False
        End If

        If arr.Rows(0)(0).ToString() = cServerError Then
            ReportServerError("LotManagerModule.RecvClosedLot", vtoa(arr.Rows(1)(0)))
            Return False
        End If

        ' Normalize Birthday column to prevent ambiguity
        If arr.Columns.Contains("Birthday") Then
            For Each row As DataRow In arr.Rows
                If Not IsDBNull(row("Birthday")) Then
                    Dim rawDate As DateTime = CType(row("Birthday"), DateTime)
                    ' Normalize to consistent format: MM/dd/yyyy HH:mm:ss
                    row("Birthday") = DateTime.ParseExact(
                    rawDate.ToString("MM/dd/yyyy HH:mm:ss", CultureInfo.InvariantCulture),
                    "MM/dd/yyyy HH:mm:ss",
                    CultureInfo.InvariantCulture
                )
                End If
            Next
        End If

        Return True


    End Function

    '=======================================================
    'Routine: frmLotManager.RecvCSuspendedLot(o)
    'Purpose: This queries the Lot Manager to get a list
    '         of the current Suspended lots.
    '
    'Globals:None
    '
    'Input: oLotManager - The lot manager connection.
    '
    'Return:None
    '
    'Modifications:
    '   12-03-1998 As written for Pass1.5
    '
    '   01-05-1999 Moved to mdlLotManager to facilitate
    '   running more then one instance of frmLotmanager.
    '
    '
    '=======================================================
    Public Function RecvSuspendedLot(oLotManager As NextcapServer.clsLotManager, ByRef arr As DataTable) As Boolean
        If oLotManager Is Nothing Then
            Return False
        End If

        arr = oLotManager.GetSuspendedLots()
        ' Check for empty or error result
        If arr Is Nothing OrElse arr.Rows.Count = 0 Then
            Return False
        End If

        If arr.Rows(0)(0).ToString() = cServerError Then
            ReportServerError("LotManagerModule.RecvSuspendedLot", vtoa(arr.Rows(1)(0)))
            Return False
        End If

        ' ✅ Normalize Birthday column to prevent ambiguity
        If arr.Columns.Contains("Birthday") Then
            For Each row As DataRow In arr.Rows
                If Not IsDBNull(row("Birthday")) Then
                    Dim rawDate As DateTime = CType(row("Birthday"), DateTime)
                    ' Normalize to consistent format: MM/dd/yyyy HH:mm:ss
                    row("Birthday") = DateTime.ParseExact(
                    rawDate.ToString("MM/dd/yyyy HH:mm:ss", CultureInfo.InvariantCulture),
                    "MM/dd/yyyy HH:mm:ss",
                    CultureInfo.InvariantCulture
                )
                End If
            Next
        End If

        Return True


    End Function

    Public Function GetNextGroupId(frmCaller As frmLotManagerSink) As String
        Dim strGroupId As String

        ' 尝试从 CSB 获取 ID
        strGroupId = mdlSAX.ExecuteCSB_CreateLotId()

        ' 如果从 CSB 未获取到 ID，则使用默认方法
        If String.IsNullOrEmpty(strGroupId) Then
            strGroupId = CreateDefaultLotId(frmCaller)
        End If

        ' 返回新 ID
        Return strGroupId
    End Function

    Public Function CreateDefaultLotId(Optional ByRef frmManager As frmLotManagerSink = Nothing) As String
        Dim strGroupId As String
        Dim dtNow As Date = DateTime.Now
        Dim nLimiter As Integer = 1

        Try
            ' 确保有有效的 LotManager 引用
            If frmManager.LotManager Is Nothing Then
                frmManager.LotManager = go_ActiveLotManager
                'Return Nothing
            End If

            ' 生成初始唯一 ID
            strGroupId = GenerateUniqueId(dtNow, frmManager)

            ' 检查该批次 ID 是否已被使用
            While (Not frmManager.LotManager.LotNameUnique(strGroupId)) AndAlso (nLimiter < 100)
                ' 尝试生成下一个批次 ID
                frmManager.oLotsToday.AddLot(dtNow)
                strGroupId = GenerateUniqueId(dtNow, frmManager)
                nLimiter += 1
            End While

            Return strGroupId

        Catch ex As Exception
            ' 简单错误处理，可根据需要扩展
            Debug.WriteLine("生成默认批次 ID 时出错: " & ex.Message)
            Return String.Empty
        End Try
    End Function

    Private Function GenerateUniqueId(ByRef dtNow As Date, ByRef frmManager As frmLotManagerSink) As String
        Dim strGroupId As String
        Dim strLine As String
        Dim strLineNumber As String

        ' 构建 YMMDD 格式
        strGroupId = dtNow.ToString("yy").Substring(1) &
                     dtNow.Month.ToString("00") &
                     dtNow.Day.ToString("00")

        ' 添加生产线类型的首字母（大写）
        strGroupId &= frmManager.LotManager.LineType.Substring(0, 1).ToUpper()

        ' 添加生产线编号（大写）
        strGroupId &= frmManager.LotManager.LineNumber.ToString().ToUpper()

        ' 添加当天的批次计数
        Return strGroupId & frmManager.oLotsToday.NextLotCount(dtNow).ToString("00")
    End Function
End Module
