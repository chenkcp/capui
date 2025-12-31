Imports System.Configuration
Imports System.IO
Imports System.Reflection
Imports System.Runtime.CompilerServices.RuntimeHelpers
Imports System.Runtime.Intrinsics
Imports System.Security.Cryptography
Imports System.Timers
Imports System.Windows.Forms
Imports System.Windows.Media.TextFormatting
Imports LiveCharts
Imports LiveCharts.Wpf


Public Class frmNextCap
    Inherits Form
    ' UI Controls
    Public sbrNextCap As StatusStrip
    Public tbrNextCap As ToolStrip
    Public tmrContinuousBad As System.Windows.Forms.Timer
    Public WithEvents hscrNextCapGraph As HScrollBar
    Public picGraphNextCap As Panel
    Public panelStatusIcons As Panel
    Public panelInputOptions As Panel

    Public imgLotPointer As PictureBox
    Public imgStatus As New List(Of PictureBox)
    Public frameGraphStyle As GroupBox
    Public optDefects As RadioButton
    Public optGoodBad As RadioButton
    Public txtUnitId As TextBox
    Public shpBusy As Panel ' Substitute for Shape control in VB6
    ' Create a dictionary to map class names to their series
    Public seriesMap As New Dictionary(Of String, StackedColumnSeries)()
    Private StartIndex As Integer = 0


    ' Toolbar Buttons
    Private WithEvents btnGoodPen, btnBadPen, btnEditPen,
        btnDeletePen, btnUndo, btnReports,
        btnContext, btnLotMgt, btnPareto, btnLocal As ToolStripButton

    ' LiveCharts Graph - Replacement for grphNextCap
    Public WithEvents grphNextCap As New LiveCharts.WinForms.CartesianChart()
    Private selectedChartPoint As ChartPoint
    Public Property ChartMaxPoints As Integer
    Public Property ChartMaxVisiblePoints As Integer

    '/*Reference to the global Security object
    'Private WithEvents m_oSecurity As clsSecurity
    'Private m_nAuthority As Integer

    ' Data Series
    Private correctSeries As StackedColumnSeries
    Private errorSeries As StackedColumnSeries
    Private m_colChartData As colChartData ' You must define this elsewhere
    Public GoodSeries As StackedColumnSeries
    Public BadSeries As StackedColumnSeries
    Public ChartRangeMin As Integer = 1
    Public ChartRangeMax As Integer = 10

    Public imglst_tbrNextCap As ImageList

    Private WithEvents Menu_LotFunctions As New ContextMenuStrip()
    Private WithEvents mnuLotStatusEditor As New ToolStripMenuItem("Lot Status Editor")
    Private WithEvents mnuLotReport As New ToolStripMenuItem("Lot Report")


    Private menuStrip As MenuStrip
    Private mnuFile As ToolStripMenuItem
    Private mnuExternal As ToolStripMenuItem
    Private mnuUtilities As ToolStripMenuItem
    Private WithEvents mnuHelp As New ToolStripMenuItem("&Help...")
    Private WithEvents mnuAbout As New ToolStripMenuItem("&About...")
    ' File Menu
    Private WithEvents mnuSampleCount As New ToolStripMenuItem("Download Config")
    Private WithEvents mnuResetDefectHotlist As New ToolStripMenuItem("Upload Data")
    Private WithEvents mnuDatabase As New ToolStripMenuItem("Lot Management")
    Private WithEvents mnuPrint As New ToolStripMenuItem("Context")
    Private WithEvents mnuUpload As New ToolStripMenuItem("TBD")
    Private WithEvents mnuExit As New ToolStripMenuItem("E&xit")

    ' External Test Menu
    Private WithEvents mnuReadID As New ToolStripMenuItem("&Inspector")
    Private WithEvents mnuServer As New ToolStripMenuItem("&Film Burst")
    Private WithEvents mnuDownload As New ToolStripMenuItem("&Ink Weight")
    'Private WithEvents mnuDiagnostic As New ToolStripMenuItem("&Diagnostic")

    ' Report Menu
    Private WithEvents mnuStartRecording As New ToolStripMenuItem("&Failure Trend Plot")
    Private WithEvents mnuRegressionTest As New ToolStripMenuItem("&Burst Strength Plot")
    Private WithEvents mnuDebugLog As New ToolStripMenuItem("&Ink Weight Plot")


    '/*Minimum Height & Width of the screen before it is Unusable
    Const clngMinHeight As Long = 7230
    Const clngMinWidth As Long = 10035
    '/*Connection to Unit Input Object
    Private WithEvents m_oweUnitInput As clsUnitId

    '/*Constants to define state machine
    Private Const cnINPUT_READY As Integer = 0
    Private Const cnINPUT_GOOD As Integer = 1
    Private Const cnINPUT_BAD As Integer = 2
    Private Const cnINPUT_DELETE As Integer = 3
    Private Const cnINPUT_EDIT As Integer = 4
    Private Const cnINPUT_UNDO As Integer = 5
    Private Const cnINPUT_EXECUTE As Integer = 6
    '/*Current input state
    Private m_nState As Integer
    Private m_strLastButtonKey As String

    Private Const cstrScanMode As String = "SCANNER"

    Private m_nPanelRegression As Integer
    '/*Flag to indicate if we are contected to the practice database
    Private m_bPracticeDatabase As Boolean
    Private m_nPanelWarning As Integer
    '/*Storage for the new name of a Regression test
    Public m_RegressionName As String
    Public m_RegressionCopy As String
    '/*Access from other procedures to inform this form that
    '/*a forced shutdown is occuring
    Public m_bForcedShutDown As Boolean
    '/*Reference to the global Security object
    Private WithEvents m_oSecurity As Security

    '/*Properties to track the maximum visible Chart points
    Private m_lngChartPoints As Integer
    Private m_lngChartMaxPoints As Integer

    Private m_nAuthority As Integer

    '/*Constants to defeine the background colors
    '/*for the Busy circle
    Private Const m_clngOnGreen As Long = &HFF00&
    Private Const m_clngOffGrey As Long = &HE0E0E0

    '/*Constants for words used in Form_KeyDown; v1.1.3
    Private Const m_cstrGood As String = "btnGood"
    Private Const m_cstrBad As String = "btnBad"

    ' -- v2.0.8 member variable to track the last status icon that was right clicked
    Private m_nLotIndexHit As Integer

    ' This form's current stack ID
    Private m_strFrmId As String
    Private Const xWindowSize As Integer = 20
    Private xCount As Integer
    Private allGoodValues As New ChartValues(Of Double)
    Private allBadValues As New ChartValues(Of Double)
    Private allXLabels As New List(Of String)

    Private WithEvents lotTimer As System.Timers.Timer
    'Public Event TimeToCreateLot()
    Public ReadOnly Property StrFrmId As String
        Get
            Return m_strFrmId
        End Get
    End Property

    '/*Flag to indicate if the registry read function should
    '/*be triggered in the ShowMe() function
    ' Flag indicating that regression mode is on
    Private m_bRegression As Boolean

    ' Property to set the ReadRegistry flag
    Private m_bReadPersistence As Boolean
    Public Property ReadPersistence As Boolean
        Get
            Return m_bReadPersistence
        End Get
        Set(value As Boolean)
            m_bReadPersistence = value
        End Set
    End Property

    ' Setting for input mode
    Private m_strInputMode As String
    ' Property to set Input Mode
    Public Property InputMode As String
        Get
            Return m_strInputMode
        End Get
        Set(value As String)
            ' Initialize the scanner state
            If value = cstrScanMode Then
                txtUnitId.Enabled = False
                txtUnitId.BackColor = SystemColors.GrayText
            End If
            m_strInputMode = value
        End Set
    End Property

    ' Flag to indicate if we are in continuous bad pen mode
    Private m_bContinuousBadMode As Boolean
    ' Property to set Continuous Bad mode
    Public Property ContinuousBad As Boolean
        Get
            Return m_bContinuousBadMode
        End Get
        Set(value As Boolean)
            m_bContinuousBadMode = value
        End Set
    End Property

    ' Private copy of what the current chart type is
    Private m_strChartType As String
    Protected Overrides Sub WndProc(ByRef m As Message)
        'method is automatically called by the Windows Forms framework whenever a message is sent to the form. Need not to call it explicitly.
        Const WM_SYSKEYDOWN As Integer = &H104
        Const WM_SYSKEYUP As Integer = &H105

        ' Suppress menu activation for Alt key
        If m.Msg = WM_SYSKEYDOWN OrElse m.Msg = WM_SYSKEYUP Then
            If CInt(m.WParam) = Keys.Menu Then
                Return ' Do nothing for Alt key
            End If
        End If

        MyBase.WndProc(m)
    End Sub
    Private Function HandleCreateUnit(unitType As String) As Boolean
        If gb_InputBusyFlag Or gb_InputRequestFlag Then Return True
        gb_InputBusyFlag = True

        If go_ActiveLotManager IsNot Nothing Then
            If m_nAuthority < 4 AndAlso m_nAuthority > -1 Then
                MessageBox.Show("Insufficient authority to perform this action.")
                Return True
            End If

            CreateUnit(unitType)
            FocusInput()
        End If

        Return True
    End Function
    Protected Overrides Function ProcessCmdKey(ByRef msg As Message, keyData As Keys) As Boolean
        Debug.Print("Global key pressed: " & keyData.ToString())

        ' Handle Alt-key combinations
        If (keyData And Keys.Alt) = Keys.Alt Then
            If m_strInputMode <> cstrScanMode Then
                Dim keyWithoutAlt As Keys = keyData And Not Keys.Alt

                Select Case keyWithoutAlt
                    Case Keys.G
                        MessageBox.Show("Alt+G detected.")
                        DirectCast(tbrNextCap.Items(m_cstrGood), ToolStripButton).Checked = True

                        Return HandleCreateUnit("GOOD")

                    Case Keys.B
                        MessageBox.Show("Alt+B detected.")
                        Return HandleCreateUnit("BAD")

                        ' Add more Alt+Key cases here as needed
                End Select
            End If

        ElseIf m_strInputMode = cstrScanMode Then
            ' Handle SCANNER mode
            Select Case keyData
                Case Keys.Enter
                    MessageBox.Show("Enter detected.")

                    If m_nState = cnINPUT_READY Then
                        If gb_InputBusyFlag Or gb_InputRequestFlag Then Return True

                        If m_bContinuousBadMode Then
                            If InputState(cnINPUT_BAD) Then
                                m_strLastButtonKey = m_cstrBad
                                DirectCast(tbrNextCap.Items(m_cstrBad), ToolStripButton).Checked = True
                            End If
                        Else
                            If InputState(cnINPUT_GOOD) Then
                                m_strLastButtonKey = m_cstrGood
                                DirectCast(tbrNextCap.Items(m_cstrGood), ToolStripButton).Checked = True
                            End If
                        End If

                    Else
                        gb_InputBusyFlag = True
                        InputState(cnINPUT_EXECUTE)
                        gb_InputBusyFlag = False
                    End If

                Case Keys.Escape
                    MessageBox.Show("Escape detected.")

                    If m_bReadPersistence Then
                        Me.Close()
                    Else
                        txtUnitId.Text = ""
                    End If
            End Select
        End If

        Return MyBase.ProcessCmdKey(msg, keyData)
    End Function

    Public Sub New()


        ' This call is required by the designer.
        InitializeComponent()
        'Me.KeyPreview = True ' Ensure form receives key events before controls

        Debug.WriteLine("Checking in frmNextcap New()")
        ' Add any initialization after the InitializeComponent() call.
        SetupGraphPanel() ' DockStyle=Top
        SetupTopPanel() ' DockStyle=Top, machine status and scrollbar
        SetupMenu() ' DockStyle=Top 
        SetupStatusIconsPanel() 'DockStyle=Bottom
        SetupInputOptionsPanel() 'DockStyle=Bottom
        SetupToolStripButtons() 'DockStyle=Bottom


        'Form_New runs once each time the form Is created. It's New constructor runs once for each new instance of the form, before Load is called

        m_strFrmId = "frmNextCap" & Guid.NewGuid().ToString()
        'Debug.WriteLine($"New Form ID: {m_strFrmId}")
        ' Initialize the chart type to Good/Bad
        m_strChartType = "GoodBad"
        ' Initialize the chart max points
        m_lngChartMaxPoints = 1000
        m_lngChartPoints = 0
        ' Initialize the input mode to Scanner mode
        InputMode = "" 'cstrScanMode
        ' Initialize the continuous bad mode to False
        ContinuousBad = False

        Me.ActiveControl = Nothing

        lotTimer = New System.Timers.Timer() With {
          .Interval = 600000,
          .Enabled = True
      }
    End Sub
    ' Property Get procedure
    Public ReadOnly Property ChartType As String
        Get
            Return m_strChartType
        End Get
    End Property
    Private Sub Form1_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        '/*Set our references to nothing
        'm_oweUnitInput = Nothing
        Debug.WriteLine("Form unload")
    End Sub
    ' Handles the form's closing event (user clicks X, Alt+F4, etc.)
    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        ' Example: prompt the user before closing
        If MessageBox.Show("Are you sure you want to exit?", "Confirm Exit", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then
            e.Cancel = True ' Cancel the close
            Return
        End If

        ' Perform any cleanup here
        Debug.WriteLine("Form is closing. Reason: " & e.CloseReason.ToString())
    End Sub

    Private Sub frmNextCap_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Form_Load runs once each time the form Is created And shown for the first time. It's New constructor runs once for each new instance of the form, before Load is called


        Debug.WriteLine("Checking in frmNextcap Load()")
        Dim idx As Integer = 1
        Dim lineType As String
        Dim lineNumber As Integer
        Dim source As String
        Dim chartDataCount As Integer
        For Each oLotSink As frmLotManagerSink In gcol_LotManagers
            lineType = oLotSink.LotManager.LineType
            lineNumber = oLotSink.LotManager.LineNumber
            source = oLotSink.LotManager.Source
            chartDataCount = 0
            If Not oLotSink.ChartDataCol Is Nothing Then
                chartDataCount = oLotSink.ChartDataCol.Count
            End If
            Debug.WriteLine($"[{idx}] LineType={lineType}, LineNumber={lineNumber}, Source={source}, ChartData.Count={chartDataCount}")
            idx += 1
        Next

        ' Trigger the Good/Bad change event
        'optGoodBad_MouseUp(Nothing, New MouseEventArgs(MouseButtons.Left, 1, 0, 0, 0))

        Debug.WriteLine("Checking in frmNextcap Load() after InitializeStatusIcon")
        idx = 1
        For Each oLotSink As frmLotManagerSink In gcol_LotManagers
            lineType = oLotSink.LotManager.LineType
            lineNumber = oLotSink.LotManager.LineNumber
            source = oLotSink.LotManager.Source
            chartDataCount = 0
            If Not oLotSink.ChartDataCol Is Nothing Then
                chartDataCount = oLotSink.ChartDataCol.Count
            End If
            Debug.WriteLine($"[{idx}] LineType={lineType}, LineNumber={lineNumber}, Source={source}, ChartData.Count={chartDataCount}")
            idx += 1
        Next

        If go_ActiveLotManager Is Nothing Then
            Debug.WriteLine("No active LotManager.")
        Else
            Debug.WriteLine("Found active LotManager.")
        End If

        lineType = go_ActiveLotManager.LineType
        lineNumber = go_ActiveLotManager.LineNumber
        source = go_ActiveLotManager.Source

        ' ChartDataCol is typically accessed via go_ActiveLotSink
        chartDataCount = 0
        If Not go_ActiveLotSink Is Nothing AndAlso Not go_ActiveLotSink.ChartDataCol Is Nothing Then
            chartDataCount = go_ActiveLotSink.ChartDataCol.Count
        End If

        Debug.WriteLine($"Form Load - Active LotManager: LineType={lineType}, LineNumber={lineNumber}, Source={source}, ChartData.Count={chartDataCount}")
        mdlChart.SetChartActiveData("GoodBad")

        '/*Preset the Max points used in calculations
        m_lngChartMaxPoints = 1
        m_lngChartPoints = 1
        '/*Set are PCs name in the second statusbar pane
        'If Not bCSBUsed Then Call mdlNextcap.UpdateStatusBar(mdlTools.CurrentMachineName(), 2)
        mdlNextcap.UpdateStatusBar(mdlTools.CurrentMachineName(), 2)
        '/*Connect the security object
        m_oSecurity = go_Security
        '/In case we are connecting after the object startup

        If Not (m_oSecurity Is Nothing) Then m_oSecurity_SecurityChange(m_oSecurity.nAuthority)


    End Sub
    Private Sub SetupTopPanel()
        ' Create a panel for the scrollbar
        Dim panelScrollBar As New Panel() With {.Dock = DockStyle.Top, .Height = 20}
        hscrNextCapGraph = New HScrollBar() With {
            .Dock = DockStyle.Fill,
            .Name = "hscrNextCapGraph" ' Set the Name property
        }
        panelScrollBar.Controls.Add(hscrNextCapGraph)
        Me.Controls.Add(panelScrollBar)


        ' Create the line info panel and status strip
        Dim lineInfoPanel = New Panel() With {.Dock = DockStyle.Top, .Height = 36}
        sbrNextCap = New StatusStrip() With {.Dock = DockStyle.Fill}

        ' Add a spacer before the first label
        Dim spacer As New ToolStripLabel() With {
        .Text = "",
        .AutoSize = False,
        .Width = 24 ' Adjust width as needed for your desired space
    }
        sbrNextCap.Items.Add(spacer)
        ' Add status labels and separators
        Dim labelCount As Integer = 6
        For i As Integer = 1 To labelCount
            Dim label As New ToolStripLabel($"Panel {i}") With {
                .Name = $"Panel{i}" ' Set the Name property
            }
            sbrNextCap.Items.Add(label)
            If i < labelCount Then
                Dim sep As New ToolStripSeparator()
                sep.AutoSize = False
                sep.Width = 8
                sep.Height = 10
                sbrNextCap.Items.Add(sep)
            End If
        Next
        lineInfoPanel.Controls.Add(sbrNextCap)
        Me.Controls.Add(lineInfoPanel)


    End Sub
    Private Sub SetupGraphPanel()
        ' Create and configure the panel for the chart
        picGraphNextCap = New Panel() With {.Dock = DockStyle.Fill}
        grphNextCap.Dock = DockStyle.Fill
        grphNextCap.BackColor = Color.White

        ' Configure Y axis on the right
        grphNextCap.AxisY.Clear()
        grphNextCap.AxisY.Add(New LiveCharts.Wpf.Axis With {
        .MinValue = 0, ' Default minimum value
        .MaxValue = 60, ' Default maximum value
        .ShowLabels = False,
        .Separator = New LiveCharts.Wpf.Separator With {.Step = 1}
        })

        ' Configure the X axis
        grphNextCap.AxisX.Clear()
        grphNextCap.AxisX.Add(New LiveCharts.Wpf.Axis With {
        .Title = "X Axis",
        .Labels = New List(Of String) From {"Placeholder"}, ' Placeholder labels
        .Separator = New LiveCharts.Wpf.Separator With {.Step = 1},
        .LabelsRotation = 45 ' Rotate labels for better visibility
    })
        ' Add the chart to the panel
        picGraphNextCap.Controls.Add(grphNextCap)
        Me.Controls.Add(picGraphNextCap)

        ' when a user clicks on a data point (bar) in the grphNextCap chart, the grphNextCap_HotHit event is triggered.
        AddHandler grphNextCap.DataClick, AddressOf grphNextCap_HotHit

        ' Timer
        tmrContinuousBad = New System.Windows.Forms.Timer With {.Interval = 1000}
    End Sub

    Private Sub SetupStatusIconsPanel()
        Try
            'create an empty panel for status icon.  By default, the `Visible` property of a new panel is `True`.
            panelStatusIcons = New Panel() With {.Dock = DockStyle.Bottom, .Height = 44}
            Me.Controls.Add(panelStatusIcons)
        Catch ex As Exception
            Debug.WriteLine($"Error initializing panelStatusIcons: {ex.Message}")
        End Try
    End Sub

    Private Sub SetupInputOptionsPanel()
        panelInputOptions = New Panel() With {.Dock = DockStyle.Bottom, .Height = 44}

        ' TableLayoutPanel with 2 columns
        Dim inputLayout As New TableLayoutPanel With {
        .Dock = DockStyle.Fill,
        .ColumnCount = 2,
        .RowCount = 1
    }
        inputLayout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100)) ' TextBox column stretches
        inputLayout.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))     ' Radio group column auto-sizes

        ' TextBox (left)
        txtUnitId = New TextBox With {
            .Name = "txtUnitId",
            .Width = 100, 'textbox starts at 100px wide.
            .Anchor = AnchorStyles.Left Or AnchorStyles.Right Or AnchorStyles.Top,
            .Margin = New Padding(8, 10, 8, 10),
            .MaximumSize = New Size(300, 0) 'Keep the column as Percent for streteched panel, but set a MaximumSize. When the parent panel stretches, the textbox can grow, but never beyond 300px wide
        }
        inputLayout.Controls.Add(txtUnitId, 0, 0)

        ' FlowLayoutPanel for radio buttons (right)
        Dim radioPanel As New FlowLayoutPanel With {
            .FlowDirection = FlowDirection.LeftToRight,
            .AutoSize = True,
            .WrapContents = False,
            .Anchor = AnchorStyles.Right Or AnchorStyles.Top,
            .Margin = New Padding(0, 5, 8, 5)
        }
        optGoodBad = New RadioButton With {.Name = "optGoodBad", .Text = "Good/Bad", .Checked = True, .AutoSize = True, .Margin = New Padding(10, 10, 10, 10)}
        optDefects = New RadioButton With {.Name = "optDefects", .Text = "Defects", .AutoSize = True, .Margin = New Padding(10, 10, 10, 10)}
        radioPanel.Controls.Add(optGoodBad)
        radioPanel.Controls.Add(optDefects)

        ' After creating optGoodBad and optDefects
        AddHandler optGoodBad.MouseUp, AddressOf optGoodBad_MouseUp
        AddHandler optDefects.MouseUp, AddressOf optDefects_MouseUp

        inputLayout.Controls.Add(radioPanel, 1, 0)

        panelInputOptions.Controls.Clear()
        panelInputOptions.Controls.Add(inputLayout)
        Me.Controls.Add(panelInputOptions)

    End Sub
    Private Sub SetupToolStripButtons()
        tbrNextCap = New ToolStrip() With {.Dock = DockStyle.Bottom, .Size = New Size(64, 50)}
        Dim btnGoodPen = New ToolStripButton() With {.Name = "btnGood", .Size = New Size(64, 47), .Text = "Good Pen", .ToolTipText = "Enter good pen"}
        Dim btnBadPen = New ToolStripButton() With {.Name = "btnBad", .Size = New Size(64, 47), .Text = "Bad Pen", .ToolTipText = "Enter bad pen"}
        Dim btnEditPen = New ToolStripButton() With {.Name = "btnEdit", .Size = New Size(64, 47), .Text = "Edit Pen", .ToolTipText = "Edit previously entered pen"}
        Dim btnDeletePen = New ToolStripButton() With {.Name = "btnDelete", .Size = New Size(64, 47), .Text = "Delete Pen", .ToolTipText = "Delete previously entered pen"}
        Dim btnUndo = New ToolStripButton() With {.Name = "btnUndo", .Size = New Size(64, 47), .Text = "Undo", .ToolTipText = "Undo previously entered pen"} 'FFFF



        Dim btnContext = New ToolStripButton() With {.Name = "btnContext", .Size = New Size(64, 47), .Text = "Context", .ToolTipText = "Edit Context"}
        Dim btnLotMgt = New ToolStripButton() With {.Name = "LotMgt", .Size = New Size(64, 47), .Text = "Lot Mgt", .ToolTipText = "Edit Lot"}

        ' Create a custom dark separator
        Dim darkSeparator As New ToolStripSeparator()
        darkSeparator.AutoSize = False
        darkSeparator.Width = 16
        darkSeparator.BackColor = Color.DimGray


        tbrNextCap.Items.AddRange({
            btnGoodPen, darkSeparator.Clone(), 'refer to Module ToolStripExtensions.vb
            btnBadPen, darkSeparator.Clone(),
            btnEditPen, darkSeparator.Clone(),
            btnDeletePen, darkSeparator.Clone(),
            btnUndo, New ToolStripSeparator() With {.AutoSize = False, .Width = 32},
            btnContext, darkSeparator.Clone(),
            btnLotMgt
        })
        Me.Controls.Add(tbrNextCap)
        ' Attach event handlers
        AddHandler btnGoodPen.Click, AddressOf tbrNextCapButton_Click
        AddHandler btnBadPen.Click, AddressOf tbrNextCapButton_Click
        AddHandler btnEditPen.Click, AddressOf tbrNextCapButton_Click
        AddHandler btnDeletePen.Click, AddressOf tbrNextCapButton_Click
        AddHandler btnUndo.Click, AddressOf tbrNextCapButton_Click
        AddHandler btnContext.Click, AddressOf tbrNextCapButton_Click
        AddHandler btnLotMgt.Click, AddressOf tbrNextCapButton_Click
        btnUndo.Visible = False 'not implement after migrating from VB6

    End Sub

    Private Sub SetupMenu()
        Dim MenuStrip = New MenuStrip()
        ' File Menu

        ' Add all items to the File menu
        mnuFile = New ToolStripMenuItem("&Main")
        mnuFile.DropDownItems.AddRange({
            mnuSampleCount,
            mnuResetDefectHotlist,
            mnuDatabase,
            mnuPrint,
            mnuUpload,
            New ToolStripSeparator(),
            mnuExit ' Add the exit item with the handler attached
        })

        ' Utilities Menu
        mnuExternal = New ToolStripMenuItem("&External Test")
        mnuExternal.DropDownItems.AddRange({
            mnuReadID,
            mnuServer,
            mnuDownload
        })

        mnuUtilities = New ToolStripMenuItem("&Chart")
        mnuUtilities.DropDownItems.AddRange({
            mnuStartRecording,
            mnuRegressionTest,
            mnuDebugLog
        })
        ' Lot Functions Menu (Hidden)
        'Dim mnuLotFunctions As New ToolStripMenuItem("&Lot Functions") With {.Visible = False}
        'mnuLotFunctions.DropDownItems.AddRange(New ToolStripItem() {
        '    New ToolStripMenuItem("Lot Status Editor") With {.Name = "mnuLotStatusEditor"},
        '    New ToolStripMenuItem("Lot Report") With {.Name = "mnuLotReport"}
        '})


        ' Add all to the MenuStrip   ', mnuCSBIDE, mnuTest,mnuLotFunctions
        MenuStrip.Items.AddRange({mnuFile, mnuExternal, mnuUtilities, mnuHelp, mnuAbout})

        MenuStrip.Dock = DockStyle.Top
        Me.MainMenuStrip = MenuStrip
        Me.Controls.Add(MenuStrip)
    End Sub
    Public Sub RefreshChartAtCurrentIndex()

        hscrNextCapGraph.Minimum = 0
        hscrNextCapGraph.Maximum = gcol_ChartData.Count - 1
        hscrNextCapGraph.LargeChange = xWindowSize

        hscrNextCapGraph.Value = StartIndex
        UpdateChartWindow()

    End Sub

    '================================================================
    'Routine: frmNextCap.hscrNextCapGraph_Scroll
    'Purpose:
    ' This is an event handler for the horizontal scrollbar (hscrNextCapGraph) associated with your chart. Its purpose is to update the visible window of data in the chart when the user scrolls left or right. The scrollbar allows the user to move the visible "window" across the full set of X data points. When the user scrolls, the Scroll event is triggered, and hscrNextCapGraph_Scroll is called
    '
    'Globals:None
    '
    'Input: sender, ScrollEventArgs
    '
    'Return:None
    '
    'Modifications:
    '   06-20-2025 As written for Nextcap2025

    '======================================================================
    Private Sub hscrNextCapGraph_Scroll(sender As Object, e As ScrollEventArgs) Handles hscrNextCapGraph.Scroll

        If hscrNextCapGraph.Value < hscrNextCapGraph.Minimum OrElse hscrNextCapGraph.Value > hscrNextCapGraph.Maximum Then
            Return ' Prevent bounce-back
        End If
        Debug.WriteLine($"Scroll Event Triggered: Type={e.Type}, Value={hscrNextCapGraph.Value}")
        If e.Type = ScrollEventType.EndScroll Then
            Dim maxStart As Integer = GetMaxStartIndex() ' = gcol_ChartData.Count - xWindowSize
            'determines the new starting index based on the scrollbar's value, but ensures it does not exceed the maximum
            StartIndex = Math.Min(hscrNextCapGraph.Value, maxStart)

            'Change the current view window; update the chart and status icons to show the correct subset of data.
            'UpdateChartWindow(StartIndex)
            RefreshChartAtCurrentIndex()

            'grphNextCap.Refresh() ' Ensure the chart is visually updated
        End If

    End Sub
    Private Function GetMaxStartIndex() As Integer
        If gcol_ChartData Is Nothing OrElse gcol_ChartData.Count <= xWindowSize Then
            Return 0
        End If
        Return gcol_ChartData.Count - xWindowSize
    End Function

    'Private Sub hscrNextCapGraph_ValueChanged(sender As Object, e As EventArgs) Handles hscrNextCapGraph.ValueChanged
    '    Dim maxStart As Integer = Math.Max(0, xCount - xWindowSize)
    '    Dim startIndex As Integer = Math.Min(hscrNextCapGraph.Value, maxStart) ' Use the current Value property
    '    Debug.WriteLine($"Scrollbar Value Changed: New Value = {hscrNextCapGraph.Value}")
    '    UpdateChartWindow(startIndex)
    '    grphNextCap.Refresh() ' Ensure the chart is visually updated
    'End Sub
    '================================================================
    'Routine: frmNextCap.hscrNextcapGraphConfig
    'Purpose:
    ' This method configures the horizontal scrollbar (hscrNextCapGraph) that controls which portion of the chart's X data points are visible.
    '
    'Globals:None
    '
    'Input: None
    '
    'Return:None
    '
    'Modifications:
    '   06-20-2025 As written for Nextcap2025

    '======================================================================
    Private Sub hscrNextcapGraphConfig()
        Dim hscroll = Me.Controls.Find("hscrNextCapGraph", True).FirstOrDefault()
        xCount = gcol_ChartData.Count

        If hscroll IsNot Nothing Then
            ' Setup scrollbar for X axis window scrolling
            With hscrNextCapGraph
                .Minimum = 0 '  the lowest possible value of the scrollbar. It is the starting point of the scroll range.
                .Maximum = Math.Max(0, xCount - xWindowSize) 'the highest possible start index that still allows a full window of data to be shown defined by xWindowSize
                .LargeChange = 1 'Both set to 1, so each scroll step moves the window by one data point.
                .SmallChange = 1
                .Value = 0 'Initial position, represents the current position of the scrollbar's thumb (slider). if Value = 5, it shows the window starting at the 6th data point.
                .Enabled = xCount > xWindowSize ': Only enabled if there are more data points than can be shown at once (xCount > xWindowSize).
            End With
        End If
    End Sub
    '================================================================
    'Routine:frmNextCap_Shown
    'Purpose:
    ' Runs only once — specifically the first time the form becomes visible to the user.Sets initial scrollbar And graph view To rightmost. It runs after the form Is loaded And displayed, right after Load, Activated, And layout events complete.Even later hide And show the form again, Shown will Not run again.
    '
    '======================================================================
    Private Sub frmNextCap_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        StartIndex = GetMaxStartIndex()
        ' Set scrollbar
        hscrNextCapGraph.Minimum = 0
        hscrNextCapGraph.Maximum = gcol_ChartData.Count - 1
        hscrNextCapGraph.LargeChange = xWindowSize

        hscrNextCapGraph.Value = StartIndex
        RefreshChartAtCurrentIndex()
    End Sub
    '================================================================
    'Routine: frmNextCap.UpdateChartWindow (startIndex)
    'Purpose:
    ' This is to create new collections containing only the data points that should be displayed in the current visible window of the  chart. This allows the chart to efficiently display only the relevant data as the user scrolls left or right, without redrawing the entire dataset.
    '
    'Globals:None
    '
    'Input: startIndex - The index of the first X data point currently visible in your chart window
    '
    'Return:None
    '
    'Modifications:
    '   06-20-2025 As written for Nextcap2025

    '======================================================================
    Public Sub UpdateChartWindow()
        ' Ensure xCount is updated based on the current data
        xCount = gcol_ChartData.Count
        Dim CurrentStartIndex As Integer = StartIndex
        ' Calculate the maximum start index for the scrollbar
        Dim maxStart As Integer = Math.Max(0, xCount - xWindowSize)

        ' Exit if gcol_ChartData hasn't populated.
        If gcol_ChartData Is Nothing OrElse gcol_ChartData.Count = 0 Then
            ' Clear the chart and display a placeholder message
            grphNextCap.Series.Clear()
            grphNextCap.AxisX.Clear()
            grphNextCap.AxisX.Add(New LiveCharts.Wpf.Axis With {
            .Labels = New List(Of String) From {"No data available"},
            .Separator = New LiveCharts.Wpf.Separator With {.Step = 1},
            .LabelsRotation = 45
        })
            Debug.WriteLine("UpdateChartWindow: gcol_ChartData is empty or uninitialized.")
            Exit Sub
        End If

        ' Extract the visible range of data for GoodSeries and BadSeries
        Dim goodWindow = New ChartValues(Of Double)(
        gcol_ChartData.Skip(StartIndex).Take(xWindowSize).Select(Function(cd) CDbl(cd.nGoodCount))
    )
        Dim badWindow = New ChartValues(Of Double)(
        gcol_ChartData.Skip(StartIndex).Take(xWindowSize).Select(Function(cd) CDbl(cd.nBadCount))
    )
        Dim xLabelsWindow = gcol_ChartData.Skip(StartIndex).Take(xWindowSize).Select(Function(cd) cd.strGroupId).ToList()

        ' Update the X axis labels
        grphNextCap.AxisX.Clear()
        grphNextCap.AxisX.Add(New LiveCharts.Wpf.Axis With {
            .Labels = xLabelsWindow,
            .Separator = New LiveCharts.Wpf.Separator With {.Step = 1},
            .LabelsRotation = 45
        })


        ' Update the Y axis dynamically
        grphNextCap.AxisY.Clear()
        grphNextCap.AxisY.Add(New LiveCharts.Wpf.Axis With {
            .MinValue = 0,
            .MaxValue = goodWindow.Max() + 10, ' Adjust max value dynamically
            .Separator = New LiveCharts.Wpf.Separator With {.Step = 10}
        })


        If m_strChartType = "GoodBad" Or m_strChartType Is Nothing Then
            ' Clear all existing series
            grphNextCap.Series.Clear()
            GoodSeries = New StackedColumnSeries With {
                .Title = "Good",
                .Values = goodWindow,
                .DataLabels = True
            }
            grphNextCap.Series.Add(GoodSeries)

            BadSeries = New StackedColumnSeries With {
                 .Title = "Bad",
                 .Values = badWindow,
                 .DataLabels = True
            }
            grphNextCap.Series.Add(BadSeries)

            Debug.WriteLine($"Series Count: {grphNextCap.Series.Count}")

        Else
            ' Remove GoodSeries and BadSeries manually
            grphNextCap.Series.Clear()

            ' Create or update defectSeries
            Dim defectCount As Integer = go_ChartConfigs.ColCharts(0).ColChartDefects.Count
            For i = 1 To defectCount - 1
                Dim seriesTitle = go_ChartConfigs.ColCharts(0).ColChartDefects(i).StrDesc
                Dim existingSeries = grphNextCap.Series.FirstOrDefault(Function(s) CType(s, StackedColumnSeries).Title = seriesTitle)

                Dim defectValues = New ChartValues(Of Double)(
                    gcol_ChartData.Select(Function(cd) Convert.ToDouble(cd.colIndivBad(i))).Skip(StartIndex).Take(xWindowSize)
                )

                Dim defectSeries = New StackedColumnSeries With {
                        .Title = seriesTitle,
                        .Values = defectValues,
                        .DataLabels = True
                    }
                grphNextCap.Series.Add(defectSeries)

            Next

        End If

        ' Update the visible picture boxes
        InitializeStatusIcons()
        MakeGroupVisible(CurrentStartIndex)


    End Sub
    '================================================================
    'Routine: frmNextCap.MakeGroupVisible (startIndex)
    'Purpose:
    ' The method is called whenever the visible window of the chart changes (e.g., when the user scrolls or resizes the chart).
    ' It calculates which data points (columns) are currently visible in the chart window, based on startIndex and xWindowSize.
    ' For each status icon (PictureBox in imgStatus)
    ' •	If its index Is within the visible window, it Is positioned horizontally to align with its corresponding chart column And made visible.
    ' •	If it Is outside the visible window, it Is hidden.
    '
    'Globals:None
    '
    'Input: startIndex - The index of the first X data point currently visible in your chart window
    '
    'Return:None
    '
    'Modifications:
    '   06-20-2025 As written for Nextcap2025

    '======================================================================
    Public Sub MakeGroupVisible(startIndex As Integer)
        ' Call this at the end of UpdateChartWindow
        ' Get chart area width
        Dim chartWidth As Integer = grphNextCap.Width
        Dim leftPadding As Integer = 10 ' Adjust for Y axis/labels if needed
        Dim rightPadding As Integer = 10
        Dim usableWidth As Integer = chartWidth - leftPadding - rightPadding
        Debug.WriteLine($"Running MakeGroupVisible")
        ' Calculate X positions for visible columns
        For i = 0 To xCount - 1
            Dim pb = imgStatus(i)
            If i >= startIndex AndAlso i < startIndex + xWindowSize Then
                Dim visibleIndex As Integer = i - startIndex
                Dim xPos As Integer = leftPadding + CInt((usableWidth / xWindowSize) * (visibleIndex + 0.5)) - pb.Width \ 2
                pb.Left = xPos
                pb.Top = (panelStatusIcons.Height - pb.Height) \ 2
                pb.Visible = True
            Else
                pb.Visible = False
            End If
        Next
        panelStatusIcons.Visible = True
        Me.Refresh()
        Debug.WriteLine($"gcol_ChartData Count: {gcol_ChartData?.Count}")
        Debug.WriteLine($"imgStatus Count: {imgStatus.Count}")
        Debug.WriteLine($"panelStatusIcons Visible: {panelStatusIcons.Visible}")
        Debug.WriteLine($"panelStatusIcons Height: {panelStatusIcons.Height}")
        Debug.WriteLine($"Parent Visible: {panelStatusIcons.Parent?.Visible}")
        If panelStatusIcons.InvokeRequired Then
            panelStatusIcons.Invoke(Sub() panelStatusIcons.Visible = True)
        Else
            panelStatusIcons.Visible = True
        End If
    End Sub
    '================================================================
    'Routine: frmNextCap.InitializeStatusIcons ()
    'Purpose:
    ' The method is to (re)create and initialize the list of status icon PictureBox controls (imgStatus) that visually represent the status of each X data point in your chart. It iterates through the current imgStatus list, removes each PictureBox from the panelStatusIcons panel, detaches its event handler, and disposes of it to free resources.
    ' For each index from 0 to total number of X data points  - 1, it creates a new PictureBox (the status icon), sets its size, color, and visibility, attaches the click event handler, and adds it to both the imgStatus list and the panelStatusIcons pane
    '
    'Globals:None
    '
    'Input: None
    '
    'Return:None
    '
    'Modifications:
    '   06-20-2025 As written for Nextcap2025

    '======================================================================
    Public Sub InitializeStatusIcons()
        ' Call this once after initializing gcol_ChartData and panelStatusIcons
        ' Remove any existing PictureBoxes
        For Each pb In imgStatus
            If panelStatusIcons.Controls.Contains(pb) Then
                panelStatusIcons.Controls.Remove(pb)
            End If
            RemoveHandler pb.MouseClick, AddressOf StatusIcon_Click
            pb.Dispose()
        Next
        imgStatus.Clear()
        xCount = gcol_ChartData.Count

        Try
            ' Create one PictureBox for each X point 
            For i = 0 To xCount - 1
                Dim pb As New PictureBox With {
                    .Width = 16,
                    .Height = 16,
                    .BackColor = Color.LightGray, ' Customize as needed
                    .Visible = False ' Only visible when in the current window
                }
                If i >= gcol_ChartData.Count Then
                    Debug.WriteLine($"Index {i} exceeds gcol_ChartData.Count ({gcol_ChartData.Count}).")
                    Exit For
                End If
                Debug.WriteLine($"frmNextcap.InitializeStatusIcons: iconPath {gcol_ChartData(i).strIconPath}")
                ' Set the image from strIconPath if the file exists
                Dim iconPath As String = gcol_ChartData(i).strIconPath
                If Not String.IsNullOrEmpty(iconPath) AndAlso IO.File.Exists(iconPath) Then
                    Try
                        '' Load the icon and convert to a resized bitmap
                        'Using ico As New Icon(iconPath, pb.Width, pb.Height)
                        '    pb.Image = ico.ToBitmap()
                        'End Using
                        ' For PNG, JPG, BMP, etc.
                        Using img As Image = Image.FromFile(iconPath)
                            pb.Image = New Bitmap(img, pb.Width, pb.Height)
                        End Using
                    Catch ex As Exception
                        pb.Image = Nothing ' Optionally handle error
                    End Try
                End If
                pb.Tag = iconPath
                AddHandler pb.MouseClick, AddressOf StatusIcon_Click
                imgStatus.Add(pb)
                panelStatusIcons.Controls.Add(pb)
            Next

        Catch ex As Exception
            Debug.WriteLine($"Error {ex}") ' Optionally handle error
        End Try

    End Sub

    '================================================================
    'Routine: frmNextCap.grphNextCap_Resize
    'Purpose:
    ' This is an event handler that responds to the Resize event of the grphNextCap chart control. Its purpose is to realign and reposition the status icons (imgStatus) whenever the chart is resized. It calculates the current startIndex (the first visible data point in the chart window). Then calls MakeGroupVisible(startIndex) to reposition And show/hide the status icons so they remain correctly aligned with the visible chart columns.
    '
    'Globals:None
    '
    'Input: sender, EventArgs
    '
    'Return:None
    '
    'Modifications:
    '   06-20-2025 As written for Nextcap2025

    '======================================================================
    Private Sub grphNextCap_Resize(sender As Object, e As EventArgs) Handles grphNextCap.Resize
        ' Prevent running if hscrNextCapGraph is not yet created
        If hscrNextCapGraph Is Nothing Or imgStatus.Count = 0 Then Exit Sub
        ' Use the current startIndex (visible window) for correct alignment.
        startIndex = Math.Min(hscrNextCapGraph.Value, Math.Max(0, xCount - xWindowSize))
        MakeGroupVisible(startIndex)
    End Sub
    '================================================================
    'Routine: frmNextCap.StatusIcon_Click
    'Purpose:
    ' This is an event handler that responds to the mouse click on the picturebox of the status icons.
    '
    'Globals:None
    '
    'Input: sender, MouseEventArgs
    '
    'Return:None
    '
    'Modifications:
    '   06-20-2025 As written for Nextcap2025

    '======================================================================
    Private Sub StatusIcon_Click(sender As Object, e As MouseEventArgs)
        Dim pb As PictureBox = CType(sender, PictureBox)
        Dim nLotIndex As Integer = imgStatus.IndexOf(pb)
        Dim frmActiveSink As frmLotManagerSink



        '/*Get a reference to the ActiveSink
        frmActiveSink = go_ActiveLotSink 'frmLotManager.ActiveLotManager

        If e.Button = MouseButtons.Left Then
            'Debug.WriteLine($"StatusIcon[{nLotIndex}]: {frmActiveSink.ChartDataCol(nLotIndex).strGroupId} selected")
            Dim form As New frmLotStatusEditor()
            form.lblGroupId.Text = frmActiveSink.ChartDataCol(nLotIndex).strGroupId
            form.imgQuality.Image = pb.Image
            form.lblCurrent.Text = frmActiveSink.ChartDataCol(nLotIndex).strIconName
            ' 填充其他控件的数据，例如颜色列表等
            form.lstQualityStatus.Items.Add("red")
            form.lstQualityStatus.Items.Add("yellow")
            form.lstQualityStatus.Items.Add("green")

            form.SetGroup(frmActiveSink.ChartDataCol(nLotIndex).strGroupId, frmActiveSink.ChartDataCol(nLotIndex).dtBirth, frmActiveSink)
            form.ShowDialog()

        End If


    End Sub

    '================================================================
    'Routine: frmNextCap.grphNextCap_HotHit
    'Purpose:
    ' This is an event handler that responds to the mouse click on the Chart columns.
    '
    'Globals:None
    '
    'Input: sender, LiveCharts.ChartPoint
    '
    'Return:None
    '
    'Modifications:
    '   06-20-2025 As written for Nextcap2025

    '======================================================================
    Private Sub grphNextCap_HotHit(sender As Object, chartPoint As LiveCharts.ChartPoint)
        Dim oReport As clsReportData
        Try
            ' Assume your chart's DataContext or Tag holds the collection similar to gcol_ChartData
            'Debug.WriteLine(grphNextCap.Tag)
            'Dim chartDataList As List(Of ChartData) = CType(grphNextCap.Tag, List(Of ChartData))
            ' Get the index of the clicked bar
            Dim pointIndex As Integer = CInt(chartPoint.X)
            ' Retrieve the full label from the X-Axis
            Dim fullLabel As String = grphNextCap.AxisX(0).Labels(pointIndex)


            ' Adjust for visible range if you have paging/scrolling logic
            ' For example, if you have a RangeMin property:
            ' Dim nGroup As Integer = pointIndex + (RangeMin - 1)
            Dim nGroup As Integer = pointIndex ' Adjust as needed

            If nGroup >= 0 AndAlso nGroup < gcol_ChartData.Count Then
                Dim groupData = gcol_ChartData(nGroup + StartIndex)

                ' Left mouse button: show report
                If Control.MouseButtons = MouseButtons.Left Then
                    Debug.WriteLine($"left clicked {groupData.strGroupId}")
                    oReport = mdlReports.CreateLotSummary(groupData.strGroupId, groupData.dtBirth)
                    'EditLotPen(groupData.strGroupId, groupData.dtBirth) 'FFFF

                    ' Create and show the summary form
                    Dim summaryForm As New frmLotSummary(groupData.strGroupId, groupData.dtBirth)
                    summaryForm.SetReport(oReport)
                    Dim x = summaryForm.ShowDialog()
                    If x = DialogResult.OK OrElse x = DialogResult.Yes Then
                        Me.RefreshChartAtCurrentIndex()
                    End If

                    'OpenBarDetailsWindow() 'FFFF
                    ' Create and show the report form
                    'Dim report As New LotSummaryReport(groupData.GroupId, groupData.BirthDate)
                    'report.ShowDialog()
                    ' Right mouse button: show tooltip
                ElseIf Control.MouseButtons = MouseButtons.Right Then
                    Debug.WriteLine($"Right clicked {groupData.strGroupId}")
                    EditLotStatus(groupData.strGroupId, groupData.dtBirth) 'FFFF

                    'Dim nCount As Integer = groupData.BadCount + groupData.GoodCount
                    'Dim toolTipText As String = $"Group: {groupData.GroupId} - Count: {nCount}"
                    'ToolTip1.Show(toolTipText, grphNextCap)
                End If
            Else
                MessageBox.Show("NextCap encountered an internal error attempting to display the requested Lot Report", "Warning")
            End If
        Catch ex As Exception
            MessageBox.Show($"Error: {ex.Message}", "Error")
        End Try

    End Sub

    '%%%%%%%%%%%%
    Private Sub EditLotStatus(lotId As String, bDate As DateTime) 'FFFF
        MessageBox.Show($"Showing details for Lot ID: {lotId}", "Details", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Private Sub EditLotPen(lotId As String, bDate As DateTime)
        Try
            ' Validate the DateTime value directly
            If bDate = DateTime.MinValue Then
                MessageBox.Show("Invalid date format. Please use MM/dd/yyyy h:mm:ss tt.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If

            ' Create and show the summary form
            Dim summaryForm As New frmLotSummary(lotId, bDate)

            Dim x = summaryForm.ShowDialog()
            If x = DialogResult.OK OrElse x = DialogResult.Yes Then
                Me.RefreshChartAtCurrentIndex()
            End If
        Catch ex As Exception
            MessageBox.Show($"Error: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub OpenBarDetailsWindow() 'FFFF


        If selectedChartPoint IsNot Nothing Then
            Dim lotLabel As String = grphNextCap.AxisX(0).Labels(CInt(selectedChartPoint.X)) ' Get label

            Dim barDetailsForm As New BarDetailsForm()
            barDetailsForm.LoadBarDetails(lotLabel)

            ' set StartPosition as Manual
            barDetailsForm.StartPosition = FormStartPosition.Manual
            ' set window to top
            barDetailsForm.TopMost = True
            ' set window location as mouse location
            barDetailsForm.Location = New Point(Cursor.Position.X, Cursor.Position.Y)

            'barDetailsForm.Show()
            ' 以模态对话框的形式显示小窗体
            barDetailsForm.ShowDialog()
        End If
    End Sub
    Private Sub frmNextcap_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        Debug.Print("Form key down triggered: " & e.KeyCode.ToString())
        'Form_KeyDown in VB.NET is the form-level event for handling key presses, using the KeyDown event and the KeyEventArgs parameter. 
        'The logic Is driven by the application's state, not by UI focus.
        'The toolbar button's .Value is set programmatically to reflect the action, but the actual event handler (tbrNextCap_ButtonClick) is not called by Form_KeyDown.
        'Instead, the state machine (InputState And related variables) And direct calls to CreateUnit Or other methods perform the required action.
        If m_strInputMode = cstrScanMode Then
            If e.KeyCode = Keys.Enter Then
                ' Handle Enter key
                MessageBox.Show("Enter key pressed on the form.")
                '/*Put us into the good state by default
                If m_nState = cnINPUT_READY Then
                    '/*Do not proceed if we are running or are about to start the LotSequence manager
                    If gb_InputBusyFlag Or gb_InputRequestFlag Then Exit Sub


                    '/*If we are in continuous bad mode route
                    '/*the incoming item to the default type
                    If m_bContinuousBadMode Then
                        If InputState(cnINPUT_BAD) Then
                            m_strLastButtonKey = m_cstrBad

                            ' Set the button as pressed (checked);does not trigger the ButtonClick event. 
                            DirectCast(tbrNextCap.Items(m_cstrBad), ToolStripButton).Checked = True

                        End If
                    Else
                        If InputState(cnINPUT_GOOD) Then
                            m_strLastButtonKey = m_cstrGood
                            ' Set the button as pressed (checked);does not trigger the ButtonClick event. 
                            DirectCast(tbrNextCap.Items(m_cstrGood), ToolStripButton).Checked = True
                        End If
                    End If
                    '/*If we are not in the ready state
                    '/*then we need to excute the current state
                Else
                    '/*Set a lock on the client
                    gb_InputBusyFlag = True
                    InputState(cnINPUT_EXECUTE)
                    '/*Unlock the client
                    gb_InputBusyFlag = False
                End If
            ElseIf e.KeyCode = Keys.Escape Then
                ' Handle Escape key
                If m_bReadPersistence Then
                    '/*If we are in the read persistence mode, then we should close the form
                    Me.Close()
                Else
                    '/*Otherwise, we should just clear the UnitId text box
                    txtUnitId.Text = ""
                End If
            End If
        End If
    End Sub
    '================================================================
    'Routine: frmNextCap.tbrNextCapButton_Click
    'Purpose:
    ' This is an event handler that responds to the mouse click on the ToolStripButtons.
    '
    'Globals:None
    '
    'Input: sender, EventArgs
    '
    'Return:None
    '
    'Modifications:
    '   06-20-2025 As written for Nextcap2025

    '======================================================================
    ' Shared event handler for all ToolStripButtons
    Private Sub tbrNextCapButton_Click(sender As Object, e As EventArgs)
        Dim btn As ToolStripButton = CType(sender, ToolStripButton)
        Debug.WriteLine($"Button Name: {btn.Name}, Button Text: {btn.Text}")

        '/*Do not proceed if we are running or are about to start the LotSequence manager
        'If Gb_InputBusyFlag Or gb_InputRequestFlag Then Exit Sub
        '/*Set a lock on the client
        gb_InputBusyFlag = True

        Select Case btn.Name 'FFFF
            Case "btnGood"
                ' Handle Good Pen button click

                CreateUnit("GOOD")
                StartIndex = GetMaxStartIndex()
                Me.RefreshChartAtCurrentIndex()
                Me.FocusInput()
            Case "btnBad"
                ' Handle Bad Pen button click

                CreateUnit("BAD")
                StartIndex = GetMaxStartIndex()
                Me.RefreshChartAtCurrentIndex()
                Me.FocusInput()
                ' Add more cases for other buttons
            Case "btnEdit"
                ' Handle Good Pen button click

                CreateUnit("EDIT")
                Me.FocusInput()
            Case "btnDelete"
                ' Handle Bad Pen button click

                CreateUnit("DELETE")
                Me.FocusInput()
            Case "btnUndo"
                ' Handle Good Pen button click

                CreateUnit("UNDO")
                Me.FocusInput()
            Case "btnContext"
                ' Handle Bad Pen button click

                frmAllContext.GetInstance().ShowDialog() ' Show the context form

            Case "LotMgt"
                ' Handle Bad Pen button click

                MATERIAL_Click()
        End Select
        '/*Set the current button
        m_strLastButtonKey = btn.Name
    End Sub
    Private Sub MATERIAL_Click()
        If (go_ActiveLotManager Is Nothing) Then
            MessageBox.Show("frmNextCap.tbrNextCap_ButtonClick()" & Chr(13) & "No LotManager Available")
        Else
            '/*Insure that we are not in continuous mode
            'If go_clsSystemSettings.strMaterialMode <> mdlGlobal.gcstrContinuous Then
            '    Dim lotMangerForm As New frmLotManager()
            '    lotMangerForm.ShowDialog()
            'Else
            '    gb_InputBusyFlag = False
            'End If
            ' Use the existing instance of frmLotManager
            Dim x = mdlLotManager.ActiveLotManagerForm.ShowDialog()
            If x = DialogResult.OK OrElse x = DialogResult.Yes Then
                StartIndex = GetMaxStartIndex()
                Me.RefreshChartAtCurrentIndex()
            End If
        End If
    End Sub
    Private Function InputState(ByRef nNewState As Integer) As Boolean
        '/*Handle request to execute state
        If nNewState = cnINPUT_EXECUTE Then
            If m_nState <> cnINPUT_EXECUTE And m_nState <> cnINPUT_READY Then
                Select Case m_nState
                    Case cnINPUT_BAD
                        CreateUnit("BAD")
                    Case cnINPUT_GOOD
                        CreateUnit("GOOD")
                    Case cnINPUT_DELETE
                        CreateUnit("DELETE")
                    Case cnINPUT_EDIT
                        CreateUnit("EDIT")
                    Case cnINPUT_UNDO
                        CreateUnit("UNDO")
                End Select
                '/*Insure the last button is unpressed
                DirectCast(tbrNextCap.Items(m_strLastButtonKey), ToolStripButton).Checked = False
                txtUnitId.Enabled = False
                txtUnitId.BackColor = SystemColors.GrayText
                m_nState = cnINPUT_READY
            End If
        Else
            If nNewState = cnINPUT_READY Then
                '/*Reset the pressed button
                If Len(m_strLastButtonKey) <> 0 Then
                    DirectCast(tbrNextCap.Items(m_strLastButtonKey), ToolStripButton).Checked = False
                End If
                '/*Reset the input box to disabled state
                txtUnitId.Enabled = False
                txtUnitId.BackColor = SystemColors.GrayText
                txtUnitId.Text = ""
            Else
                txtUnitId.Enabled = True
                txtUnitId.BackColor = SystemColors.Window
                'updated 2/25/01 for Mutex
                If txtUnitId.Visible = True And txtUnitId.Enabled = True Then txtUnitId.Focus()
            End If
            '/*Set the current state
            m_nState = nNewState
            '/*Return succesful state change
            InputState = True
        End If
    End Function
    Public Sub ShowMe()
        Try
            Debug.WriteLine("Checking in frmNextcap.showMe()")
            If go_ActiveLotManager Is Nothing Then
                Debug.WriteLine("No active LotManager.")
            End If

            ' Execute required steps for a login condition
            If m_bReadPersistence OrElse Not (go_clsSystemSettings.bLogOutAtShift OrElse go_Security.bLoginRequired) Then
                ' Active Lotmanager is set here! Set the context by running the standard form update routine
                Call frmAllContext.GetInstance().ReadInRegistry()
            End If


            ' Now gcol_chartdata can setup hscrNextCapGraph 
            hscrNextcapGraphConfig()

            ' Initialize this Form's connection to the global UnitId Input class
            ' m_oweUnitInput = go_CAPmain.NextCapIdInput

            ' Trigger the Good/Bad change event
            optGoodBad_MouseUp(Nothing, New MouseEventArgs(MouseButtons.Left, 1, 0, 0, 0))


            ' Change the current user if either login option is set
            ' Change the name to match the name on the frmPassword user list
            'If go_clsSystemSettings.bLogOutAtShift OrElse go_Security.bLoginRequired Then
            '    frmAllContext.cmbOperator.SelectedIndex = frmPassword.cmbUser.SelectedIndex
            'End If

            ' Clear the Id input box
            'txtUnitId.Text = ""

            ' Pop up this window
            'frmAllContext.StrFrmId = mdlWindow.AddForm(Me)
        Catch ex As Exception
            ' Optional: handle/log exception as needed
        End Try
    End Sub
    Private Sub m_oSecurity_SecurityChange(ByVal nNewAuthority As Integer)
        '/*Log the new security level
        m_nAuthority = nNewAuthority
        '/*Make the necessary modifications
        If nNewAuthority > 7 Or nNewAuthority = -1 Then
            '/*Utility menu
            mnuDownload.Enabled = True
            mnuStartRecording.Enabled = True
            mnuRegressionTest.Enabled = True
            mnuDebugLog.Enabled = True


            '/*File Menu
            mnuPrint.Enabled = True
            mnuResetDefectHotlist.Enabled = True
            mnuSampleCount.Enabled = True

        ElseIf nNewAuthority > 3 And nNewAuthority < 8 Then
            '/*Utility menu
            mnuDownload.Enabled = True
            mnuStartRecording.Enabled = True
            mnuRegressionTest.Enabled = False
            mnuDebugLog.Enabled = True

            '/*File Menu
            mnuPrint.Enabled = True
            mnuResetDefectHotlist.Enabled = True
            mnuSampleCount.Enabled = True


        ElseIf nNewAuthority < 2 Then
            '/*Utility menu
            mnuDownload.Enabled = False
            mnuStartRecording.Enabled = True
            mnuRegressionTest.Enabled = False
            mnuDebugLog.Enabled = True


            '/*File Menu
            mnuPrint.Enabled = False
            mnuResetDefectHotlist.Enabled = False
            mnuSampleCount.Enabled = False

        End If
    End Sub
    Private Sub mnuDownload_Click(sender As Object, e As EventArgs) Handles mnuDownload.Click
        '/*This calls the supervisory system to request a shutdown
        frmServerSink.ShutdownServer()
    End Sub
    Private Sub mnuAbout_Click(sender As Object, e As EventArgs) Handles mnuAbout.Click
        ' Create and display the About window
        Using aboutForm As New frmAbout()
            aboutForm.ShowWindow()
        End Using
    End Sub

    Private Sub mnuHelp_Click(sender As Object, e As EventArgs) Handles mnuHelp.Click
        ' Read help URL from App.config (prioritize config file settings)
        Dim helpUrl As String = ConfigurationManager.AppSettings("HelpUrl")

        If String.IsNullOrEmpty(helpUrl) Then
            helpUrl = go_clsSystemSettings.strHelpURL
        End If

        ' Validate URL existence
        If String.IsNullOrEmpty(helpUrl) OrElse helpUrl.Equals("Undefined", StringComparison.OrdinalIgnoreCase) Then
            MessageBox.Show("Help documentation URL is not configured. Please contact your system administrator.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        Try
            ' Open URL in system's default browser
            Process.Start(helpUrl)
        Catch ex As FileNotFoundException
            MessageBox.Show("Default browser is not found, please check system settings", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Catch ex As Exception
            MessageBox.Show($"Failed to open Help URL：{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub mnuExit_Click(sender As Object, e As EventArgs) Handles mnuExit.Click
        '/*Call are singular exit for this form via the Form_QueryUnload
        '/*that will be triggered when we unload this form.
        Me.Close()
    End Sub
    Private Sub mnuSampleCount_Click(sender As Object, e As EventArgs) Handles mnuSampleCount.Click
        '/*Call are singular exit for this form via the Form_QueryUnload
        '/*that will be triggered when we unload this form.
        Debug.WriteLine("mnuSampleCount clicked")
    End Sub
    Private Sub mnuResetDefectHotlist_Click(sender As Object, e As EventArgs) Handles mnuResetDefectHotlist.Click
        '/*Call the defect routine in the global
        '/*defect object
        gcol_FeatureClasses.ResetHotlist()
        '/*Call the HOtlist Reset on the DefectEditor Form
        frmDefectEditor.ResetHotlist
    End Sub
    Private Sub mnuDatabase_Click(sender As Object, e As EventArgs) Handles mnuDatabase.Click
        ' Toggle between the live database and the practice database
        If m_bPracticeDatabase Then
            mnuDatabase.Text = "Standard Database"
            mnuDatabase.Checked = False
            m_bPracticeDatabase = False
            go_Supervisor.Practice(False)
            ' Modify the system warning
            RemoveDisplayWarning("PracticeDatabase")
        Else
            mnuDatabase.Text = "Practice Database"
            mnuDatabase.Checked = True
            m_bPracticeDatabase = True
            go_Supervisor.Practice(True)
            ' Modify the system warning
            DisplayWarningBar("PracticeDatabase")
        End If
    End Sub
    Private Sub mnuPrint_Click(sender As Object, e As EventArgs) Handles mnuPrint.Click
        '/*Call are singular exit for this form via the Form_QueryUnload
        '/*that will be triggered when we unload this form.
        Debug.WriteLine("mnuPrint clicked")
    End Sub
    Private Sub mnuUpload_Click(sender As Object, e As EventArgs) Handles mnuUpload.Click
        frmServerSink.InitiateUpload()
    End Sub
    Private Sub mnuDebugLog_Click(sender As Object, e As EventArgs) Handles mnuDebugLog.Click
        '/*If it is debugging is on switch it off
        If go_clsSystemSettings.bDebug Then
            go_clsSystemSettings.bDebug = False
            mnuDebugLog.Text = "Enable Debugging"
            mnuDebugLog.Checked = False
            'frmSAX.SaxBasicEntNoUI1.Trace(0)
        Else
            go_clsSystemSettings.bDebug = True
            mnuDebugLog.Text = "Disable Debugging"
            mnuDebugLog.Checked = True
            'frmSAX.SaxBasicEntNoUI1.Trace(255)
        End If
    End Sub
    ' -- Attempt to execute the server's diagnostic routines, then
    ' -- execute all of our rountines.
    '
    ' v2.0.8 Chris Barker 07-08-2002 As written
    '
    'Private Sub mnuDiagnostic_Click(sender As Object, e As EventArgs) Handles mnuDiagnostic.Click
    '    Dim oTestObject As Object
    '    Dim sResult As String
    '    Dim sServer As String
    '    Dim sErrFile As String
    '    Dim process As System.Diagnostics.Process

    '    ' 设置错误日志文件名称（使用当前时间戳）
    '    sErrFile = $"{DateTime.Now:yyyyMMddHHmmss}_client_error.txt"
    '    mdlTools.LogToFile(sErrFile, "Begining error trace for client")

    '    ' 获取服务器名称
    '    sServer = go_clsSystemSettings.strBusinessServerUNC
    '    If String.IsNullOrEmpty(sServer) Then
    '        mdlTools.LogToFile(sErrFile, "No server name listed for client ...")
    '        sServer = mdlTools.CurrentMachineName()
    '        mdlTools.LogToFile(sErrFile, $"Switching server to local machine id:{sServer} for testing")
    '    End If

    '    ' 记录连接尝试
    '    mdlTools.LogToFile(sErrFile, "Attempting to connect to server ...")

    '    ' 尝试创建服务器对象（替代VB6的CreateObject）
    '    oTestObject = mdlTools.CreateAutoObject("NextCapServer.clsBoot", sServer)

    '    ' 检查对象是否创建成功
    '    If oTestObject Is Nothing Then
    '        mdlTools.LogToFile(sErrFile, $"Server on the PC named {sServer} could not be successfully attached.")
    '        mdlTools.LogToFile(sErrFile, "Make sure that the server Ncapsrv.exe is installed and that the type libraries are consistent with the client applications")
    '    Else
    '        mdlTools.LogToFile(sErrFile, $"Server on the PC named {sServer} successfully attached.")

    '        ' 调用服务器诊断方法（后期绑定）
    '        sResult = oTestObject.diagnostic("")

    '        ' 记录诊断结果
    '        mdlTools.LogToFile(sErrFile, $"listing server diagnostic response{Environment.NewLine}{sResult}")

    '        ' 用记事本打开日志文件（替代VB6的Shell函数）
    '        Dim filePath As String = Path.Combine(Application.StartupPath, sErrFile)
    '        process = New System.Diagnostics.Process()
    '        process.StartInfo.FileName = "notepad.exe"
    '        process.StartInfo.Arguments = $"""{filePath}""" ' 正确的引号处理
    '        process.Start()

    '        oTestObject = Nothing
    '    End If
    'End Sub
    Private Sub optGoodBad_MouseUp(sender As Object, e As MouseEventArgs)
        Const cstrStyle As String = "GoodBad"
        m_strChartType = cstrStyle
        ' mdlChart.SetChartActiveStyle(cstrStyle)
        Debug.WriteLine($"optGoodBad_MouseUp: mdlChart.SetChartActiveData {cstrStyle}")
        mdlChart.SetChartActiveData(cstrStyle)
        ' Initial chart display
        'UpdateChartWindow()
        RefreshChartAtCurrentIndex()
    End Sub
    Private Sub optDefects_MouseUp(sender As Object, e As MouseEventArgs)
        Const cstrStyle As String = "Defect"
        m_strChartType = cstrStyle
        'mdlChart.SetChartActiveStyle(cstrStyle)
        Debug.WriteLine($"optDefects_MouseUp: mdlChart.SetChartActiveData {cstrStyle}")
        'mdlChart.SetChartActiveData(cstrStyle)
        ' Initial chart display
        'UpdateChartWindow()
        RefreshChartAtCurrentIndex()
    End Sub
    Public Sub CreateUnit(ByVal strType As String)
        Dim strUnitId As String
        Dim oClsPen As clsPen

        If strType = "BAD" Then
            '/*See if a pen id is required, if so test the ID
            '/*if not continue to create a Bad pen
            If AquireId(strUnitId) And (Len(strUnitId) > 0 Or Not go_clsSystemSettings.bBadPenIdRequired) Then
                '/*Add on a transaction ID if one was not given
                If Len(strUnitId) = 0 Then strUnitId = GetIdDttm()

                If frmUnitCapture.VerifyUnitId(strUnitId) Then
                    '/*Pass a Unit Id to the defect editor window
                    'frmDefectEditor.SetUnitId(strUnitId)
                    '/*Raise the defect editor window
                    'frmDefectEditor.ShowDialog()

                    Dim frmDefectEditorInstance As New frmDefectEditor
                    frmDefectEditorInstance.SetUnitId(strUnitId)
                    frmDefectEditorInstance.ShowDialog()
                Else
                    '/*Warn user that the ID is bad
                    frmMessage.GenerateMessage("frmNextCap.CreateUnit()" & Chr(13) & "Error #2 Duplicate Id (" & strUnitId & ") detected.")

                End If
            Else
                '/*See if the error window is already in place
                If Me.Enabled And Not frmMessage.Visible Then
                    '/*Warn user that the ID is bad
                    frmMessage.GenerateMessage("frmNextCap.CreateUnit()" & Chr(13) & "Error #1 Unit Id Required.")

                End If
            End If
        ElseIf strType = "GOOD" Then
            '/*See if a pen id is required, if so test the ID
            '/*if not continue to create a Good pen
            If AquireId(strUnitId) And (Len(strUnitId) > 0 Or Not go_clsSystemSettings.bGoodPenIdRequired) Then
                '/*Add on a transaction ID if one was not given
                If Len(strUnitId) = 0 Then strUnitId = GetIdDttm()

                If frmUnitCapture.VerifyUnitId(strUnitId) Then
                    '/*Show the good pen screen if it is needed
                    If go_clsSystemSettings.bGoodPenEnterUsed Then
                        '/*Pass a Unit Id to the defect editor window
                        frmEnterGoodPen.SetUnitId(strUnitId)
                        '/*Raise the Enter Good Pen window to allow for count
                        '/*and setting recovery step etc.
                        frmEnterGoodPen.ShowWindow()
                    Else
                        '/*Generate a Good Unit
                        oClsPen = New clsPen
                        '/*Set the standard features
                        oClsPen.strPenId = strUnitId
                        oClsPen.nCount = 1
                        oClsPen.strRecoveryStep = frmEnterGoodPen.SelectedRecoveryStep
                        oClsPen.strDisposition = "G"

                        '/*Now set the objects standard properties
                        '/*This is done via ByRef on the passed pointer object
                        mdlCreatePen.SetUnitStdProperties(oClsPen)

                        '/*Transmit the unit(s) to the business object
                        If mdlCreatePen.TransmitUnit(oClsPen) Then
                            '/*Show the feeback if enabled
                            If go_clsSystemSettings.bGoodPenFeedBack Then
                                'frmGoodPen.Show
                                'If frmGoodPen.Visible Then
                                '    If frmGoodPen.Enabled Then
                                '        frmGoodPen.SetFocus
                                '    End If
                                'End If
                                mdlTools.ShowSetFocus(Me) '8-10-1999
                            End If
                            '/*Release are Lock on the Client
                            gb_InputBusyFlag = False
                        Else
                            '/*Error handling routine incase something went wrong
                            '/*during the update of the unit
                            mdlLotManager.TransactionFailure("mdlCreatePen.UpdateUnit()")
                        End If
                    End If
                Else
                    '/*Warn user that the ID is bad
                    frmMessage.GenerateMessage("frmNextCap.CreateUnit()" & Chr(13) & "Error #2 Duplicate Id (" & strUnitId & ") detected.")

                End If
            Else
                '/*See if the error window is already in place
                If Me.Enabled And Not frmMessage.Visible Then
                    '/*Warn user that the ID is bad
                    frmMessage.GenerateMessage("frmNextCap.CreateUnit()" & Chr(13) & "Error #1 Unit Id Required")

                End If
            End If
        ElseIf strType = "EDIT" Then
            '/*See if a pen id is required, if so test the ID
            '/*if not continue to create a Edit pen
            If AquireId(strUnitId, True) Then
                '/*Attempt #1 - Aquire the Unit from the business server
                oClsPen = RetrievePen(strUnitId, go_Context.GroupBirthDay, go_Context.GroupId)
                '/*Attempt #2 - If no pen was found make a second attempt using the Unknown GroupId
                If oClsPen Is Nothing Then
                    If mdlWindow.GetTopForm IsNot Nothing AndAlso mdlWindow.GetTopForm.Name = "frmMessage" Then
                        mdlWindow.RemoveTopForm()
                    End If
                    oClsPen = RetrievePen(strUnitId, , "CLIENT_UNKNOWN")

                    '/*Test if a pen was actually found
                    If Not (oClsPen Is Nothing) Then
                        '/*Pass the Unit object to the Defect Editor
                        'frmDefectEditor.SetUnit(oClsPen)
                        '/*Raise the defect editor window
                        'frmDefectEditor.ShowDialog()
                        '/*Attempt #3 - Try to get a remote pen check out
                        Dim frmDefectEditorInstance As New frmDefectEditor
                        frmDefectEditorInstance.SetUnit(oClsPen)
                        frmDefectEditorInstance.ShowDialog()
                    Else
                        If mdlWindow.GetTopForm IsNot Nothing AndAlso mdlWindow.GetTopForm.Name = "frmMessage" Then mdlWindow.RemoveTopForm()
                        Call frmRemoteUpdate.PromptPenSearch(strUnitId, Me, True)
                    End If
                    '/*Pen found on Attempt #1, display the unit
                Else
                    '/*Pass the Unit object to the Defect Editor
                    'frmDefectEditor.SetUnit(oClsPen)
                    '/*Raise the defect editor window
                    'frmDefectEditor.ShowDialog()
                    Dim frmDefectEditorInstance As New frmDefectEditor
                    frmDefectEditorInstance.SetUnit(oClsPen)
                    frmDefectEditorInstance.ShowDialog()
                End If
            Else
                '/*Warn user that the ID is bad
                frmMessage.GenerateMessage("frmNextCap.CreateUnit()" & Chr(13) & "Error #1 Unit Id Required")

            End If
        ElseIf strType = "UNDO" Then
            '/*Aquire the Unit from the business server
            oClsPen = mdlCreatePen.Undo()
            '/*Test if a pen was actually found
            If Not (oClsPen Is Nothing) Then
                '/*Pass the Unit object to the Defect Editor
                frmPenDelete.SetUnit(oClsPen, 2)
                '/*Raise the defect editor window
                frmPenDelete.ShowDialog()
            Else
                frmMessage.GenerateMessage("frmNextCap.CreateUnit()" & Chr(13) & "Nothing to undo")
            End If
        ElseIf strType = "DELETE" Then
            '/*See if a pen id is required, if so test the ID
            '/*if not continue to create a Edit pen
            If AquireId(strUnitId, True) Then
                '/*Aquire the Unit from the business server
                oClsPen = RetrievePen(strUnitId, go_Context.GroupBirthDay, go_Context.GroupId)
                '/*If no pen was found make a second attempt using the Unknown GroupId
                If oClsPen Is Nothing Then
                    oClsPen = RetrievePen(strUnitId, , "CLIENT_UNKNOWN")
                End If
                '/*Test if a pen was actually found
                If Not (oClsPen Is Nothing) Then
                    '/*Pass the Unit object to the Defect Editor
                    frmPenDelete.SetUnit(oClsPen, 1)
                    '/*Raise the defect editor window
                    frmPenDelete.ShowDialog()
                End If
            Else
                '/*Warn user that the ID is bad
                frmMessage.GenerateMessage("frmNextCap.CreateUnit()" & Chr(13) & "Error #1 Unit Id Required")

            End If
        End If
        '/*Clear the TextBox in case it was the input point for the Id
        txtUnitId.Text = ""
        '/*Destroy the local copy of the Unit
        oClsPen = Nothing
    End Sub
    Private Function AquireId(ByRef strId As String, Optional ByRef bFlag As Boolean = False) As Boolean
        ' See if there is a manual entry
        If txtUnitId.Text.Length <> 0 OrElse Not (go_clsSystemSettings.strPenIdSource = "POLLED" OrElse go_clsSystemSettings.strPenIdSource = "ACTIVEX") Then
            ' Set the id according to the input box's value
            strId = txtUnitId.Text
            If String.IsNullOrEmpty(strId) Then
                Return False
            End If
            ' Execute the UnitId modification CSB
            strId = mdlSAX.ExecuteCSB_AugmentPenId(strId)
            ' Verify the unit ID against a user CSB
            If Not bFlag AndAlso strId.Length > 0 Then
                If mdlSAX.ExecuteCSB_VerifyId(strId) Then
                    Return True
                End If
            Else
                    Return True
            End If
        Else
            ' This is now the only option left since we already have txtUnitId covered and Async does not route through frmNextCap
            If go_clsSystemSettings.strPenIdSource = "POLLED" Then
                strId = frmUnitCapture.ExecutePolledId()
                ' Execute the UnitId modification CSB
                strId = mdlSAX.ExecuteCSB_AugmentPenId(strId)
                ' Verify the unit ID against a user CSB
                If Not bFlag Then
                    If mdlSAX.ExecuteCSB_VerifyId(strId) Then
                        Return True
                    End If
                Else
                    Return True
                End If
            ElseIf go_clsSystemSettings.strPenIdSource = "ACTIVEX" Then
                ' Request the Id from the server
                'strId = mdlUnitCapture.GetIdFromServer(go_ActiveXIDserver)
                ' Exit if we detect an error
                If strId = "ERROR" Then Return False
                ' Execute the UnitId modification CSB
                strId = mdlSAX.ExecuteCSB_AugmentPenId(strId)
                ' Verify the unit ID against a user CSB
                If Not bFlag Then
                    If mdlSAX.ExecuteCSB_VerifyId(strId) Then
                        Return True
                    End If
                Else
                    Return True
                End If
            Else
                frmMessage.GenerateMessage("frmNextCap.CreateUnit()" & Chr(13) & "Enter an Unit Id to proceed")
            End If
        End If
        Return False
    End Function
    Public Sub FocusInput()
        'Try
        '    If Me.Visible AndAlso Me.Enabled AndAlso txtUnitId.Enabled Then
        '        txtUnitId.Focus()
        '    End If
        'Catch
        '    ' Optionally log or handle the error
        'End Try
    End Sub
    Private Sub imgStatus_DblClick(Index As Integer)
        Dim nLotIndex As Integer
        Dim frmActiveSink As frmLotManagerSink

        On Error GoTo imgStatus_Err
        '/*Get a reference to the ActiveSink
        frmActiveSink = go_ActiveLotSink 'frmLotManager.ActiveLotManager
        '/*Determine who the Lot is that was clicked on
        nLotIndex = grphNextCap.AxisX(0).MinValue + Index

        '/*Setup the LotStatusEditor
        frmLotStatusEditor.SetGroup(frmActiveSink.ChartDataCol(nLotIndex).strGroupId,
                                frmActiveSink.ChartDataCol(nLotIndex).dtBirth,
                                frmActiveSink)
        '/*Show the window
        frmLotStatusEditor.ShowWindow()
        Exit Sub
imgStatus_Err:
        Err.Clear()
        Exit Sub
    End Sub
    Private Sub imgStatus_MouseDown(sender As Object, e As MouseEventArgs)
        If e.Button = MouseButtons.Right Then
            ' Find the index of the clicked PictureBox
            Dim index As Integer = imgStatus.IndexOf(DirectCast(sender, PictureBox))
            If index = -1 Then Return

            ' Determine which Lot was clicked
            m_nLotIndexHit = grphNextCap.AxisX(0).MinValue + index

            ' Show the context menu at the mouse position
            Menu_LotFunctions.Show(DirectCast(sender, PictureBox), e.Location)
        End If
    End Sub

    Private Sub RepositionStatusImages()
        ' Assume imgStatus is a List(Of PictureBox) or PictureBox()
        ' grphNextCap is a LiveCharts.WinForms.CartesianChart
        'If imgStatus Is Nothing OrElse imgStatus.Count = 0 Then Return

        '' Calculate positions and sizes
        'Dim nTop As Integer = grphNextCap.Location.Y + grphNextCap.Height + 20
        'Dim nLeft As Integer = grphNextCap.Location.X + CInt(0.055 * grphNextCap.Width)
        'Dim nLeftIncrement As Integer = CInt((0.86 * grphNextCap.Width) / Math.Max(1, m_lngChartPoints))

        'Dim nWidth As Integer
        'If nLeftIncrement > 615 Then
        '    nWidth = 615
        'Else
        '    nWidth = CInt(0.95 * nLeftIncrement)
        'End If
        'Dim nHeight As Integer = nWidth

        '' Position each status image
        'For i As Integer = 0 To imgStatus.Count - 1
        '    imgStatus(i).Location = New Point(nLeft, nTop)
        '    imgStatus(i).Size = New Size(nWidth, nHeight)
        '    nLeft += nLeftIncrement
        'Next

        ' Optionally, update the lot pointer icon if you have a similar method
        ' mdlChart.SetLotPointerIcon(grphNextCap.AxisX(0).MinValue, grphNextCap.AxisX(0).MaxValue)
    End Sub
    Private Sub lotTimer_Elapsed(sender As Object, e As ElapsedEventArgs) Handles lotTimer.Elapsed
        Try
            ExecuteLotCreation()
        Catch ex As Exception
            Console.WriteLine($"Error in lotTimer_Elapsed tasks: {ex.Message}")
        End Try
    End Sub

    Private Sub ExecuteLotCreation()
        ' 通过单例属性获取当前frmLotManager实例
        Dim targetForm As frmLotManager = mdlLotManager.ActiveLotManagerForm

        ' 检查实例是否有效
        If targetForm IsNot Nothing AndAlso Not targetForm.IsDisposed Then
            targetForm.AutoCreateLot = go_AutoCreateLot
            ' 线程安全调用：确保在UI线程执行
            If targetForm.InvokeRequired Then
                ' 使用Invoke切换到UI线程，避免跨线程异常
                targetForm.Invoke(Sub()
                                      ' 确认窗体未被释放后调用方法
                                      If Not targetForm.IsDisposed Then
                                          'targetForm.m_oLotManager_TimeToCreateLot()
                                          'RaiseEvent TimeToCreateLot()
                                          targetForm.m_oLotManager_TimeToCreateLot()
                                          frmLotManagerSink.m_oLotManager_TimeToCreateLot()
                                      End If
                                  End Sub)
            Else
                ' 已在UI线程，直接调用
                targetForm.m_oLotManager_TimeToCreateLot()
                frmLotManagerSink.m_oLotManager_TimeToCreateLot()
            End If
            ' try to refresh frmNextCap chart
            If Me.InvokeRequired Then
                ' 若当前在后台线程，通过自身的Invoke切换到frmNextCap的UI线程
                StartIndex = GetMaxStartIndex()
                Me.Invoke(Sub() RefreshChartAtCurrentIndex())
            Else
                ' 已在frmNextCap的UI线程，直接刷新
                StartIndex = GetMaxStartIndex()
                RefreshChartAtCurrentIndex()
            End If
        Else
            Console.WriteLine($"frmLotManager does not exist or release , end auto create lot")
        End If
    End Sub
    Private Sub RemoveDisplayWarning(ByRef strType As String)

        If strType = "PracticeDatabase" Then

            ' If this index is equal to count then we can just remove it

            If m_nPanelWarning <> sbrNextCap.Items.Count - 1 Then

                m_nPanelRegression -= 1

            End If

            ' Remove the panel and separator

            If m_nPanelWarning >= 0 AndAlso m_nPanelWarning < sbrNextCap.Items.Count Then

                sbrNextCap.Items.RemoveAt(m_nPanelWarning)

                sbrNextCap.Items.RemoveAt(m_nPanelWarning - 1)

            End If

            m_nPanelWarning = 0

        ElseIf strType = "Regression" Then

            ' If this index is equal to count then we can just remove it

            If m_nPanelRegression <> sbrNextCap.Items.Count - 1 Then

                m_nPanelWarning -= 1

            End If

            ' Remove the panel and separator

            If m_nPanelRegression >= 0 AndAlso m_nPanelRegression < sbrNextCap.Items.Count Then

                sbrNextCap.Items.RemoveAt(m_nPanelRegression)

                sbrNextCap.Items.RemoveAt(m_nPanelRegression - 1)

            End If

            m_nPanelRegression = 0

        End If

    End Sub
    '=======================================================
    'Routine: frmNextCap.DisplayWarningBar(str)
    'Purpose: This controls the positioning of warning
    'notices in teh status bar for items such as Practice
    'Database and Regression modes.
    '
    'Globals:None
    '
    'Input: strType - The specific type of status
    '
    'Return: None
    '
    'Tested:
    '
    'Modifications:
    '   02-24-1999 As written for Pass1.7
    '
    '
    '=======================================================
    Private Sub DisplayWarningBar(ByRef strType As String)
        If strType = "PracticeDatabase" Then
            ' m_nPanelWarning is index of new panel
            m_nPanelWarning = sbrNextCap.Items.Count + 1

            ' nPanel is number of new Panel name like Panel1, Panel2 etc.
            Dim nPanel = (sbrNextCap.Items.Count \ 2) + 1

            mdlNextcap.UpdateStatusBar("Practice Mode", nPanel)

            'Dim targetIndex As Integer = sbrNextCap.Items.Count - 1

            'sbrNextCap.Items(targetIndex).Image = imglst_tbrNextCap.Images(14)

        ElseIf strType = "Regression" Then
            ' m_nPanelRegression is index of new panel
            m_nPanelRegression = sbrNextCap.Items.Count + 1

            ' nPanel is number of new Panel name like Panel1, Panel2 etc.
            Dim nPanel = (sbrNextCap.Items.Count \ 2) + 1

            mdlNextcap.UpdateStatusBar("Regression", nPanel)

            'Dim targetIndex As Integer = sbrNextCap.Items.Count - 1

            'sbrNextCap.Items(targetIndex).Image = imglst_tbrNextCap.Images(16)

        End If
    End Sub
    '=======================================================
    'Routine: Menu_Utility_ReadID_Click
    'Purpose: This allows the user to read a unit ID
    'if they are using an external id application/interface.
    '
    'Globals:None
    '
    'Input:None
    'Return:None
    '
    'Tested:
    '
    'Modifications:
    '   06-21-1999 Phase3.2 Low Severity Bug list
    '
    '
    '=======================================================
    Private Sub mnuReadID_Click(sender As Object, e As EventArgs) Handles mnuReadID.Click

        Dim strId As String = ""

        Try

            AquireId(strId)

            frmMessage.GenerateMessage("ID Result: " & strId, "Read ID")

        Catch ex As Exception

        End Try

    End Sub

End Class

