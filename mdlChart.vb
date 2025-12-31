Imports System.Drawing
Imports System.IO
Imports System.Net.NetworkInformation
Imports LiveCharts
Imports LiveCharts.Definitions
Imports LiveCharts.WinForms
Imports LiveCharts.Wpf
Imports LiveCharts.Wpf.Charts.Base

Module mdlChart

    ' Constants
    Private Const m_cstrReset As String = "reset"
    Private Const EMPTYSTRING As String = ""
    Private Const BLANK As String = "blank"
    Private Const nFileMax As Integer = 20000 ' 20000 字节（约 20KB）


    ' Load chart configuration and set static properties
    Public Sub LoadChartConfig()
        Debug.WriteLine("Running LoadChartConfig")
        HarvestChartConfig()
        SetChartStaticProperties()
    End Sub
    Public Sub InitializeChart()
        ' Only proceed if there is an active lot manager
        If go_ActiveLotManager Is Nothing Then
            Return
        End If

        ' Loop through each sink in the collection and initialize it
        For Each sink As frmLotManagerSink In gcol_LotManagers
            Try
                Debug.WriteLine($"checking mdlChart.InitializeData: ")

                Dim idx As Integer = 1

                Dim lineType As String = sink.LotManager.LineType
                Dim lineNumber As Integer = sink.LotManager.LineNumber
                Dim source As String = sink.LotManager.Source
                Dim chartDataCount As Integer = 0
                If Not sink.ChartDataCol Is Nothing And sink.ChartDataCol.Count > 0 Then
                    chartDataCount = sink.ChartDataCol.Count
                End If
                Debug.WriteLine($"[{idx}] Before mdlChart.InitializeData: LineType={lineType}, LineNumber={lineNumber}, Source={source}, ChartData.Count={chartDataCount}")



                sink.InitializeData()

                If Not sink.ChartDataCol Is Nothing And sink.ChartDataCol.Count > 0 Then
                    chartDataCount = sink.ChartDataCol.Count
                End If
                Debug.WriteLine($"[{idx}] After mdlChart.InitializeData: LineType={lineType}, LineNumber={lineNumber}, Source={source}, ChartData.Count={chartDataCount}")

                idx += 1


            Catch ex As Exception
                ' Optionally log or handle individual sink errors without stopping the loop
                Debug.WriteLine($"InitializeData failed for sink: {ex.Message}")
            End Try
        Next
    End Sub
    Public Sub CreateQualityStatusObject(
    ByRef strName As String,
    ByRef strImage As String,
    ByRef strMessage As String
)
        Dim strImagePath As String = ""

        ' 确保名称未被使用
        If gcol_QualityStatus.IsUsed(strName) Then
            Exit Sub
        End If

        ' 去除字符串空格
        strImage = strImage.Trim()

        ' 设置图像路径（如果图像非空且非"BLANK"）
        If Len(strImage) > 0 AndAlso UCase(strImage) <> "BLANK" And (UCase(strImage) <> "UNKNOWN") Then
            strImagePath = AddImageToStatusIcons(strImage)
        End If

        ' 向集合添加质量状态（使用名称作为索引）
        gcol_QualityStatus.Add(strName, strImagePath, strMessage, strName)
    End Sub
    ' Harvest chart configuration from global objects
    Private Sub HarvestChartConfig()
        Dim nChartGroups As Integer
        Dim nDefectCount As Integer
        Dim oColDefects As colChartDefects

        With go_ChartConfigs
            '/*Note: this is hard coded, but could be opened up
            '/*to deal with mutilple chart configurations in
            '/*the future if needed.
            '/*Get the number of Charts setup
            nChartGroups = 0
            For nGroupItem = 0 To nChartGroups
                .ColCharts = New colCharts()
                oColDefects = New colChartDefects()
                .ColCharts.Add("", 1, 0, "", 1, 0, "", "Pen",
                    go_clsSystemSettings.strLegendGood,
                    go_clsSystemSettings.strLegendBad,
                    go_clsSystemSettings.nColorOfGood,
                    go_clsSystemSettings.nColorOfBad,
                    oColDefects)

                '/*This maps the Title and color of the items
                '/*in the FeatureClasses (Defect Class) object
                '/*to the Chart Config object.
                nDefectCount = gcol_FeatureClasses.Count
                '/*Generate a new collection of Defect types
                .ColCharts(nGroupItem).ColChartDefects = New colChartDefects()
                If .ColCharts(nGroupItem).ColChartDefects.Count = 0 Then
                    .ColCharts(0).ColChartDefects.Add("Good", 2, "")
                End If
                For nDefectItem = 0 To nDefectCount - 1
                    .ColCharts(nGroupItem).ColChartDefects.Add(
                        gcol_FeatureClasses(nDefectItem).strTitle,
                        gcol_FeatureClasses(nDefectItem).nColor, "")
                Next




            Next
        End With
    End Sub

    ' Set static chart properties on the form
    Public Sub SetChartStaticProperties()

        Try
            If go_ChartConfigs Is Nothing Then
                MainErrorHandler("mdlChart.SetChartStaticProperties", "Global object is nothing - go_ChartConfigs")
                Return
            End If
            Debug.WriteLine("Running SetChartStaticProperties")
            ' Retrieve configuration values
            Dim lngMaxPoints As Integer = go_ChartConfigs.LngMaxGroups
            Dim lngVisiblePoints As Integer = go_ChartConfigs.LngVisibleGroups

            ' Ensure at least one point
            If lngMaxPoints < 1 Then
                lngMaxPoints = 1
                go_ChartConfigs.LngMaxGroups = 1
            End If

            ' Calculate scroll window
            Dim nMin As Integer = Math.Max(0, lngMaxPoints - lngVisiblePoints)

            '' Configure the horizontal scrollbar
            'With frmNextCap.hscrNextCapGraph
            '    .Maximum = 32000
            '    .Value = CInt((32000.0 / lngMaxPoints) * nMin)
            'End With

            ' Update form‐level chart properties
            mdlMain.frmNextCapInstance.ChartMaxPoints = lngMaxPoints
            mdlMain.frmNextCapInstance.ChartMaxVisiblePoints = lngVisiblePoints

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
            ' Configure chart axes/ranges; •	The MinValue and MaxValue properties work as expected only if your chart is using indexed/categorical axes (which is the case when you use .Labels).
            'Dim chart = frmNextCap.grphNextCap
            'chart.AxisX(0).NumPoints = lngMaxPoints 
            'chart.AxisX(0).MaxValue = lngMaxPoints
            ' chart.AxisX(0).MinValue = nMin

            'If chart.Series IsNot Nothing AndAlso chart.Series.Count > 0 Then
            ' Ensure imgStatus list is resized and visible
            'For lngLoop = 1 To lngVisiblePoints - 1
            '    Dim pic As New PictureBox() With {
            '                .Size = New Drawing.Size(33, 41),
            '                .Location = New Drawing.Point(10 + (lngLoop * 50), 0),
            '                .Visible = True
            '    }

            '    frmNextCap.imgStatus.Add(pic)
            '    ' Add to the panelStatusIcons panel
            '    Dim panelStatusIcons = frmNextCap.Controls.OfType(Of Panel)().FirstOrDefault(Function(p) p.Name = "panelStatusIcons")
            '    If panelStatusIcons IsNot Nothing Then
            '        panelStatusIcons.Controls.Add(pic)
            '    End If
            'Next


        Catch ex As Exception
            MainErrorHandler("mdlChart.SetChartStaticProperties", ex.Message)
            LogToFile("Error.txt", "mdlChart.SetChartStaticProperties: " & ex.Message)
        End Try
    End Sub


    ' Add or update chart data
    Public Function AddData(strGroupId As String,
                        dtBirth As Date,
                        strDefectName As String,
                        nCount As Integer,
                        colChartData As colChartData,
                        strQualityStatus As String) As Integer

        Try
            ' Ignore if count is -1 (used to indicate non-existing defect)
            If nCount = -1 Then Return 0

            ' Defensive: Ensure dtBirth is not ambiguous
            If dtBirth = DateTime.MinValue Then
                LogToFile("Error.txt", $"Invalid birth date in AddData for GroupId: {strGroupId}")
                Return 0
            End If

            ' Determine defect order: 0 = Good, otherwise find index
            Dim nDefectOrder As Integer = If(strDefectName = "Good", 0,
                                         go_ChartConfigs.ColCharts(0).GetDefectIndex(strDefectName))

            ' Find if this group already exists
            Dim nGroupId As Integer = colChartData.GroupIdExist(strGroupId, dtBirth)
            Dim nNewCount As Integer

            If nGroupId > 0 Then
                ' Update existing data group
                nNewCount = colChartData.Item(nGroupId - 1).ChangeDataSeries(nDefectOrder, nCount)
            Else
                ' Create new group if missing
                nGroupId = CreateDataGroup(strGroupId, dtBirth, colChartData)
                nNewCount = colChartData.Item(nGroupId - 1).ChangeDataSeries(nDefectOrder, nCount)
            End If

            ' Update quality status icon
            If strQualityStatus.ToUpper() = "RESET" Then
                With colChartData.Item(nGroupId - 1)
                    .strIconName = String.Empty
                    .strIconPath = String.Empty
                End With
            ElseIf gcol_QualityStatus.IsUsed(strQualityStatus) Then
                Dim oChartData = colChartData.Item(nGroupId - 1)
                oChartData.strIconName = strQualityStatus
                oChartData.strIconPath = gcol_QualityStatus(strQualityStatus).strImagePath
            End If

            Return nGroupId

        Catch ex As Exception
            MainErrorHandler("mdlChart.AddData", ex.Message)
            LogToFile("Error.txt", $"mdlChart.AddData: {ex.Message}")
            Return 0
        End Try
    End Function



    ' Create a new data group
    Public Function CreateDataGroup(ByRef strGroupId As String, ByRef dtBirth As Date, ByRef colChartData As colChartData) As Integer
        Try
            ' Get the number of defect categories
            Dim nDefectCategory As Integer = go_ChartConfigs.ColCharts(0).ColChartDefects.Count
            If nDefectCategory = 0 Then nDefectCategory = 1

            ' Add the group to the chart data collection
            colChartData.Add(strGroupId, dtBirth, 0, 0, nDefectCategory)
            Dim nGroupCount As Integer = colChartData.Count

            '/*Remap the data to the chart
            'MapDataToChart(colChartData)


            Return nGroupCount
        Catch ex As Exception
            MainErrorHandler("mdlChart.CreateDataGroup", ex.Message)
            LogToFile("Error.txt", "mdlChart.CreateDataGroup: " & ex.Message)
            Return 0
        End Try
    End Function

    ' Remove oldest group if over max
    Private Sub AgeDataGroups(colChartData As colChartData)
        Dim nGroupCount As Integer = colChartData.Count
        If (nGroupCount + 1) > go_ChartConfigs.LngMaxGroups Then
            colChartData.Remove(0)
        End If
    End Sub

    ' Map data to chart control
    Public Sub MapDataToChart(colChartData As colChartData, Optional strChartType As String = "")
        If Not colChartData.bActive Then Exit Sub

        Debug.WriteLine("mdlChart.MapDataToChart: Mapping data to chart...")
        mdlMain.frmNextCapInstance.InitializeStatusIcons()


    End Sub


    ' Set status icon for a group
    Public Sub SetStatusIcon(strGroupId As String, dtBirth As Date, strIconName As String, colChartData As colChartData)
        Try
            Dim nDataIndex As Integer = colChartData.GroupIdExist(strGroupId, dtBirth)
            If nDataIndex = 0 Then
                'AgeDataGroups(colChartData)
                nDataIndex = CreateDataGroup(strGroupId, dtBirth, colChartData)
            End If

            If strIconName.ToUpper() = "RESET" Then
                ResetStatusIcon(nDataIndex)
                With colChartData.Item(nDataIndex)
                    .strIconName = String.Empty
                    .strIconPath = String.Empty
                End With
            ElseIf gcol_QualityStatus.IsUsed(strIconName) Then
                Dim oChartData = colChartData.Item(nDataIndex - 1)
                oChartData.strIconName = strIconName
                oChartData.strIconPath = gcol_QualityStatus(strIconName).strImagePath
                LoadStatusIcon(nDataIndex, oChartData.strIconPath)
            End If
        Catch ex As Exception
            LogToFile("Error.txt", "SetStatusIcon: " & ex.Message)
        End Try
    End Sub

    Public Sub ResetStatusIcon(nIndex As Integer)
        Try
            Dim imgIndex = nIndex - mdlMain.frmNextCapInstance.grphNextCap.AxisX(0).MinValue
            If imgIndex >= 0 AndAlso imgIndex < frmNextCap.imgStatus.Count Then
                mdlMain.frmNextCapInstance.imgStatus(imgIndex).Image = Nothing
                mdlMain.frmNextCapInstance.imgStatus(imgIndex).Tag = String.Empty
            End If
        Catch ex As Exception
            LogToFile("Error.txt", "ResetStatusIcon: " & ex.Message)
        End Try
    End Sub

    Public Sub LoadStatusIcon(nIndex As Integer, strIconPath As String)
        Try
            Dim imgIndex = nIndex - mdlMain.frmNextCapInstance.grphNextCap.AxisX(0).MinValue
            If imgIndex < 0 OrElse imgIndex >= mdlMain.frmNextCapInstance.imgStatus.Count Then Exit Sub

            If String.IsNullOrWhiteSpace(strIconPath) OrElse strIconPath.EndsWith("\") Then
                mdlMain.frmNextCapInstance.imgStatus(imgIndex).Image = Nothing
                mdlMain.frmNextCapInstance.imgStatus(imgIndex).Tag = String.Empty
            ElseIf File.Exists(strIconPath) Then
                mdlMain.frmNextCapInstance.imgStatus(imgIndex).Image = Image.FromFile(strIconPath)
                mdlMain.frmNextCapInstance.imgStatus(imgIndex).Tag = strIconPath
            Else
                mdlMain.frmNextCapInstance.imgStatus(imgIndex).Image = mdlMain.frmNextCapInstance.imglst_tbrNextCap.Images("MissingIcon")
                mdlMain.frmNextCapInstance.imgStatus(imgIndex).Tag = "ERROR"
            End If
        Catch ex As Exception
            LogToFile("Error.txt", "LoadStatusIcon: " & ex.Message)
        End Try
    End Sub

    Public Sub MapStatusIcons(Optional nMin As Integer = 0, Optional nMax As Integer = 0)
        Try
            If nMin = 0 Then nMin = mdlMain.frmNextCapInstance.grphNextCap.AxisX(0).MinValue
            If nMax = 0 Then nMax = mdlMain.frmNextCapInstance.grphNextCap.AxisX(0).MaxValue

            For i = nMin To nMax
                Dim imgIndex = i - mdlMain.frmNextCapInstance.grphNextCap.AxisX(0).MinValue
                If mdlMain.frmNextCapInstance.imgStatus.Count > 0 Then
                    If i > gcol_ChartData.Count Then
                        mdlMain.frmNextCapInstance.imgStatus(imgIndex).Image = Nothing
                        mdlMain.frmNextCapInstance.imgStatus(imgIndex).Tag = String.Empty
                    Else
                        Dim iconPath = gcol_ChartData(i - 1).strIconPath
                        If String.IsNullOrEmpty(iconPath) OrElse Not File.Exists(iconPath) Then
                            mdlMain.frmNextCapInstance.imgStatus(imgIndex).Image = Nothing
                            mdlMain.frmNextCapInstance.imgStatus(imgIndex).Tag = String.Empty
                        Else
                            mdlMain.frmNextCapInstance.imgStatus(imgIndex).Image = Image.FromFile(iconPath)
                            mdlMain.frmNextCapInstance.imgStatus(imgIndex).Tag = iconPath
                        End If
                    End If
                End If

            Next

            'To match each frmNextCap.imgStatus PictureBox location to the X pixel of each X-axis point
            ' Get the visible X range
            Dim minIndex As Integer = mdlMain.frmNextCapInstance.grphNextCap.AxisX(0).MinValue
            Dim maxIndex As Integer = mdlMain.frmNextCapInstance.grphNextCap.AxisX(0).MaxValue
            Dim visibleCount As Integer = maxIndex - minIndex + 1

            ' Get the width of the chart area (or the panel where icons are placed)
            Dim panelStatusIcons = mdlMain.frmNextCapInstance.Controls.OfType(Of Panel)().FirstOrDefault(Function(p) p.Name = "panelStatusIcons")
            If panelStatusIcons Is Nothing Then Return

            Dim chartWidth As Integer = panelStatusIcons.Width
            Dim barWidth As Double = chartWidth / visibleCount

            ' Position each PictureBox under its X-axis point
            For i As Integer = minIndex To maxIndex
                Dim relativeIndex As Integer = i - minIndex
                If relativeIndex < mdlMain.frmNextCapInstance.imgStatus.Count Then
                    ' Center the icon under the bar
                    Dim xPixel As Integer = CInt(relativeIndex * barWidth + barWidth / 2 - mdlMain.frmNextCapInstance.imgStatus(relativeIndex).Width / 2)
                    mdlMain.frmNextCapInstance.imgStatus(relativeIndex).Location = New Point(xPixel, 0)
                End If
            Next

        Catch ex As Exception
            Debug.WriteLine("Error.txt", "MapStatusIcons: " & ex.Message)
        End Try
    End Sub

    Public Sub SetLotPointerIcon(Optional ByRef nMin As Integer = 0, Optional ByRef nMax As Integer = 0)
        If Not go_clsSystemSettings.bUseLotPointer Then
            mdlMain.frmNextCapInstance.imgLotPointer.Visible = False
            Exit Sub
        End If

        Dim openLot As DataTable = go_ActiveLotManager.GetOpenLot()
        If openLot.Rows(0)(0).ToString() = "ERROR" Then
            mdlMain.frmNextCapInstance.imgLotPointer.Visible = False
            Exit Sub
        End If

        Dim lotIndex = gcol_ChartData.GroupIdExist(openLot.Rows(0)(0).ToString(), CDate(openLot.Rows(0)(1)))
        If lotIndex >= mdlMain.frmNextCapInstance.grphNextCap.AxisX(0).MinValue AndAlso lotIndex <= mdlMain.frmNextCapInstance.grphNextCap.AxisX(0).MaxValue Then
            Dim nLeftBase = mdlMain.frmNextCapInstance.grphNextCap.Left + 50
            Dim spacing = mdlMain.frmNextCapInstance.grphNextCap.Width / (mdlMain.frmNextCapInstance.grphNextCap.AxisX(0).MaxValue - mdlMain.frmNextCapInstance.grphNextCap.AxisX(0).MinValue + 1)
            Dim xPos = nLeftBase + spacing * (lotIndex - mdlMain.frmNextCapInstance.grphNextCap.AxisX(0).MinValue)

            With mdlMain.frmNextCapInstance.imgLotPointer
                .SetBounds(CInt(xPos), mdlMain.frmNextCapInstance.imgStatus(0).Top, 48, 48)
                .Visible = True
            End With
        Else
            mdlMain.frmNextCapInstance.imgLotPointer.Visible = False
        End If
    End Sub

    ' Utility: Trim long lot IDs
    Public Function TrimLotId(strLotId As String) As String
        If strLotId.Length > 15 Then
            Return strLotId.Substring(0, 12) & "..."
        End If
        Return strLotId
    End Function

    ' ... (other routines as needed, such as error handling, image file writing, etc.)
    '=======================================================
    'Routine: mdlChart.ExecuteImageCheck()
    'Purpose: This checks that the LotId/State data structures
    'agree with the Chart and Status icons. If it fails it
    'will lock down the system and attempt to reload the
    'all feedback.
    '
    'Globals:None
    '
    'Input:None
    '
    'Return:None
    '
    'Modifications:
    '   01-25-2000 Written for Beta3
    '
    '
    '=======================================================
    Public Sub ExecuteImageCheck()
        Static nRecursionCount As Integer

        Try
            ' Prevent infinite recursion
            nRecursionCount += 1
            If nRecursionCount > 2 Then
                nRecursionCount -= 1
                Exit Sub
            End If

            If gcol_ChartData Is Nothing OrElse gcol_ChartData.Count = 0 Then Exit Sub

            Dim labels = mdlMain.frmNextCapInstance.grphNextCap.AxisX(0).Labels
            Dim nLbound As Integer = mdlMain.frmNextCapInstance.grphNextCap.AxisX(0).MinValue
            Dim nUbound As Integer = mdlMain.frmNextCapInstance.grphNextCap.AxisX(0).MaxValue


            For i As Integer = nLbound To nUbound
                Dim relativeIndex = i - nLbound
                If i <= gcol_ChartData.Count Then
                    Dim expectedLabel As String = TrimLotId(gcol_ChartData(i).strGroupId)
                    Dim chartLabel As String = If(relativeIndex < labels.Count, labels(relativeIndex), "")

                    ' Check label mismatch
                    If expectedLabel <> chartLabel Then
                        LogToFile("Error.txt", $"ExecuteImageCheck:{i}:LotId, mismatch '{chartLabel}' vs expected '{expectedLabel}'")
                        Throw New ApplicationException($"LotId mismatch at position {i}")
                    End If

                    ' Check icon path mismatch
                    Dim expectedPath As String = gcol_ChartData(i).strIconPath
                    Dim currentTag As Object = If(relativeIndex < mdlMain.frmNextCapInstance.imgStatus.Count, mdlMain.frmNextCapInstance.imgStatus(relativeIndex).Tag, Nothing)
                    Dim currentPath As String = If(currentTag IsNot Nothing, currentTag.ToString(), "")

                    If expectedPath <> currentPath AndAlso currentPath <> "ERROR" Then
                        LogToFile("Error.txt", $"ExecuteImageCheck:{i}:Image, mismatch '{currentPath}' vs expected '{expectedPath}'")
                        Throw New ApplicationException($"Image mismatch at position {i}")
                    End If
                Else
                    ' Ensure label cleared
                    If relativeIndex < labels.Count AndAlso Not String.IsNullOrEmpty(labels(relativeIndex)) Then
                        labels(relativeIndex) = ""
                        LogToFile("Error.txt", $"ExecuteImageCheck: null lot had a label at {i}")
                    End If

                    ' Ensure icon reset
                    If relativeIndex < mdlMain.frmNextCapInstance.imgStatus.Count Then
                        Dim tag = mdlMain.frmNextCapInstance.imgStatus(relativeIndex).Tag?.ToString()
                        If tag <> "" AndAlso tag <> "BLANK" AndAlso tag <> "ERROR" Then
                            ResetStatusIcon(relativeIndex)
                            LogToFile("Error.txt", $"ExecuteImageCheck: null lot had icon at {i}")
                        End If
                    End If
                End If
            Next
        Catch ex As Exception
            If Not ExecuteImageCorrection() Then
                LogToFile("Error.txt", "ExecuteImageCheck: ExecuteImageCorrection Failed, shutting down system.")
                frmServerSink.ShutdownServer()
            End If
            LogToFile("Error.txt", $"ExecuteImageCheck exception: {ex.Message}")
        Finally
            nRecursionCount -= 1
        End Try
    End Sub

    '
    '=======================================================
    'Routine: mdlChart.ExecuteImageCorrection()
    'Purpose: This updates all the chart internal labels
    'to match the data structure and sets the ones above
    'the structure upper limit to "". As the lots are
    'processed, the server is queried for the current state
    'of each lot that is passed over.
    '
    'Globals: gcol_ChartData - The currently active ChartData
    '         collection set when the LotManager was placed
    '         as the ActiveLotManager.
    '
    'Input:None
    '
    'Return: Boolean - respond with a critical failure
    '
    'Modifications:
    '   01-25-2000 Written for Beta3
    '
    '
    '=======================================================
    Public Function ExecuteImageCorrection() As Boolean
        Try
            Debug.WriteLine("mdlChart.ExecuteImageCorrection: Correcting image data...")
            ' Lock the screen in an error state
            frmMessage.LockedMessage("Error detected in Lot names/icons" & vbCrLf & "ImageCheck #1", "Error")
            ' Mark the data offline
            gcol_ChartData.bActive = False
            ' Clear the internal data structures
            Do While go_ActiveLotSink.ChartDataCol.Count > 0
                go_ActiveLotSink.ChartDataCol.Remove(1)
            Loop
            ' Call the routines to reload the objects
            go_ActiveLotSink.InitializeData()
            ' Mark the object on-line
            gcol_ChartData.bActive = True
            ' Re-map the screen (map the data and the icons)
            SetChartActiveData()
            ' Unlock the client
            frmMessage.UnlockMessage()
            ExecuteImageCorrection = True
        Catch ex As Exception
            LogToFile("Error.txt", $"ExecuteImageCorrection: {ex.Message}")
            ExecuteImageCorrection = False
        End Try
    End Function

    '=======================================================
    'Routine: mdlChart.SetActiveChartData(str,str)
    'Purpose: This maps the series values into the Chart
    '         from the Chart data object.
    '
    'Globals:None
    '
    'Input:strStyle - This is either Good/Bad [GoodBad] or
    '       [Defects] as set from the option buttons on
    '       frmNextCap.
    '
    '       strUnitType - This is optional for selecting
    '       an enumeration of Chart Configs in case we
    '       are monitoring more then one type of object.
    '
    'Return:None
    '
    'Modifications:
    '   11-09-1998 As written for Pass1.3
    '
    '   01-25-2000 Check that the screen data is correct
    '
    '   02-10-2000 Lock out if another process is
    '   actively scanning the Image set.
    '=======================================================
    Public Sub SetChartActiveData(Optional ByRef strStyle As String = "", Optional ByRef strUnitType As String = Nothing)
        Try
            Debug.WriteLine("mdlChart.SetChartActiveData: Setting active chart data...")
            Dim idx As Integer = 1
            Dim lineType As String
            Dim lineNumber As Integer
            Dim source As String
            Dim chartDataCount As Integer

            If go_ActiveLotManager Is Nothing Then
                Debug.WriteLine("No active LotManager.")

            End If

            lineType = go_ActiveLotManager.LineType
            lineNumber = go_ActiveLotManager.LineNumber
            source = go_ActiveLotManager.Source

            '' ChartDataCol is typically accessed via go_ActiveLotSink
            'chartDataCount = 0
            'If Not go_ActiveLotSink Is Nothing AndAlso Not go_ActiveLotSink.ChartDataCol Is Nothing AndAlso go_ActiveLotSink.ChartDataCol.Count > 0 Then
            '    chartDataCount = go_ActiveLotSink.ChartDataCol.Count
            'End If
            Debug.WriteLine($"[{idx}] LineType={lineType}, LineNumber={lineNumber}, Source={source}, ChartData.Count={go_ActiveLotSink.ChartDataCol.Count}, gcol_ChartData= {gcol_ChartData.Count}")

            ' gcol_ChartData = go_ActiveLotSink.ChartDataCol
            ' Map the current chart data to the chart control
            'MapDataToChart(gcol_ChartData, strStyle)
            'update the visible picture boxes
            mdlMain.frmNextCapInstance.InitializeStatusIcons()

            ' Map the status icons below the chart
            'MapStatusIcons()

            ' Set the lot pointer icon if enabled
            'SetLotPointerIcon()

            ' Perform consistency check if needed
            'If gn_ImageCheckCount <= 0 Then
            'ExecuteImageCheck()
            'End If

        Catch ex As Exception
            LogToFile("Error.txt", $"mdlChart.SetChartActiveData: {ex.Message}")
        End Try
    End Sub
    Public Sub SetChartActiveStyle(ByRef strStyle As String, Optional strUnitType As String = "")
        Dim oChart As clsChart
        Dim nItem As Integer
        Dim nItemCount As Integer
        Dim strMsg As String

        '/*Set a reference to the Chart Config enumeration
        If Len(strUnitType) = 0 Then
            oChart = go_ChartConfigs.ColCharts(0)
        Else
            oChart = go_ChartConfigs.ColCharts(strUnitType)
        End If
        '/*Insure that the frmNextCap is initialized
        If mdlMain.frmNextCapInstance.Visible = False Then mdlMain.frmNextCapInstance.Refresh()
        '/*Set a reference to the graph on the form
        'With frmNextCap.grphNextCap
        '    '/*Switch depending on type of chart
        '    If strStyle = "GoodBad" Then
        '        '/*Set the number of sets to 2;Good + Bad
        '        .NumSets = 2
        '        '/*Set the legends for the data types
        '        .Legend(1) = oChart.StrLegendGood
        '        .Legend(2) = oChart.StrLegendBad
        '        '/*Set the color for the Good & Bad items
        '        .Color(1) = oChart.NColorGood
        '        .Color(2) = oChart.NColorBad
        '        '/*Set the style to stacked
        '        .GraphStyle = 2
        '    Else
        '        '/*get the number of Defect catagories
        '        nItemCount = oChart.ColChartDefects.Count
        '        '/*Set the number of values per group; basically the
        '        '/*number of series on the chart. This comes from the
        '        '/*count of [defect groups]
        '        .NumSets = nItemCount
        '        '/*Set the series legends and colors
        '        For nItem = 1 To nItemCount
        '            .Legend(nItem) = oChart.ColChartDefects(nItem).StrDesc
        '            .Color(nItem) = oChart.ColChartDefects(nItem).NColor
        '        Next nItem
        '        '/*Set the style to default
        '        .GraphStyle = 0
        '    End If

        '    '---------------------------------------
        '    '/*Set the Limit Lines if they are used
        '    '---------------------------------------
        '    '/*Both limit lines are used
        '    If oChart.lngLimitHighValue > 0 And oChart.lngLimitLowValue > 0 Then
        '        .LimitLines = graphBoth
        '        .LimitHighLabel = oChart.strLimitHighLabel
        '        .LimitHighValue = oChart.lngLimitHighValue
        '        .LimitLowLabel = oChart.strLimitLowLabel
        '        .LimitLowValue = oChart.lngLimitLowValue
        '        .LimitFillPattern = oChart.nLimitFillPattern

        '        '/*Only the high limit line is used
        '    ElseIf oChart.lngLimitHighValue > 0 Then
        '        .LimitLines = graphHigh
        '        .LimitHighLabel = oChart.strLimitHighLabel
        '        .LimitHighValue = oChart.lngLimitHighValue
        '        .LimitFillPattern = oChart.nLimitFillPattern

        '        '/*Only the low limit line is used
        '    ElseIf oChart.lngLimitLowValue > 0 Then
        '        .LimitLines = graphLow
        '        .LimitLowLabel = oChart.strLimitLowLabel
        '        .LimitLowValue = oChart.lngLimitLowValue
        '        .LimitFillPattern = oChart.nLimitFillPattern

        '        '/*No limit lines
        '    Else
        '        .LimitLines = 0
        '    End If
        'End With
    End Sub
    Public Function AddImageToStatusIcons(ByRef strImageName As String) As String
        Dim strPathFile As String
        Dim strSourcePath As String
        Dim strDestination As String

        Try
            ' 检查图像名称是否有效
            strImageName = strImageName.Trim()
            If String.IsNullOrEmpty(strImageName) Then
                Return "BLANK"
            End If

            ' 获取源路径
            strSourcePath = go_clsSystemSettings.strStatusIconLocation

            ' 处理源路径为空的情况（使用工作目录）
            If String.IsNullOrEmpty(strSourcePath) Then
                strSourcePath = Path.Combine(Application.StartupPath, "QIcons", strImageName)
                If File.Exists(strSourcePath) Then
                    Return strSourcePath
                End If
            Else
                ' 确保源路径以反斜杠结尾
                If Not strSourcePath.EndsWith("\") Then
                    strSourcePath &= "\"
                End If
                strSourcePath = Path.Combine(strSourcePath, strImageName)
            End If

            ' 构造目标路径（应用程序目录下的 QIcons 文件夹）
            strDestination = Path.Combine(Application.StartupPath, "QIcons", strImageName)

            ' 检查源文件是否存在
            If File.Exists(strSourcePath) Then
                ' 复制或处理图像文件（假设 ImageFile 是自定义方法，需根据实际逻辑实现）
                If ImageFile(strSourcePath, strDestination) Then
                    Return strDestination
                End If
            ElseIf File.Exists(strDestination) Then
                ' 使用已存在的目标文件
                Return strDestination
            Else
                ' 返回错误标识
                Return "error_image"
            End If

            ' 默认返回空或错误状态
            Return "BLANK"
        Catch ex As Exception
            MainErrorHandler("mdlChart.AddImageToStatusIcon", $"{ex.Message}")
            Return "BLANK"
        End Try
    End Function

    Public Function ImageFile(ByRef strSource As String, ByRef strDestination As String) As Boolean
        Dim fileData As Byte()

        Try
            ' 读取源文件为字节数组
            Using fsIn As New FileStream(strSource, FileMode.Open, FileAccess.Read)
                Using br As New BinaryReader(fsIn)
                    fileData = br.ReadBytes(CInt(fsIn.Length))
                End Using
            End Using

            ' 写入目标文件
            Using fsOut As New FileStream(strDestination, FileMode.Create, FileAccess.Write)
                Using bw As New BinaryWriter(fsOut)
                    bw.Write(fileData)
                End Using
            End Using

            Return True

        Catch ex As Exception
            LogToFile("Error.txt", $"mdlChart.ImageFile:{ex.HResult} {ex.Message}")
            Return False
        End Try
    End Function


End Module
