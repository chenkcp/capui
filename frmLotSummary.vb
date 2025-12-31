Imports System.Drawing
Imports System.Windows.Forms
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.ListView
Imports LiveCharts.Wpf

Public Class frmLotSummary
    Inherits Form

    ' Declare UI controls
    Private lblGroup As Label
    Private txtRowHeight As TextBox
    Private frameComments As GroupBox
    Private frameSummary As GroupBox
    Private lstvwComments As ListView
    Private lstvwParamDetail As ListView
    Private lstParameter As ListBox
    Private penDetail As DataGridView
    Private WithEvents cmdOK As Button
    Private WithEvents cmdPrint As Button
    Private WithEvents cmdEdit As Button
    Private WithEvents cmdDelete As Button

    '/*Window state constants
    Private Const cnSTATE_READY As Integer = 0
    Private Const cnSTATE_OK As Integer = 1
    Private Const cnSTATE_EDIT As Integer = 2
    Private Const cnSTATE_DELETE As Integer = 3
    Private Const cnSTATE_PRINT As Integer = 4
    Private Const cnSTATE_FLEXGRID As Integer = 5
    Private Const cnSTATE_EXECUTE As Integer = 6
    Private m_nState As Integer

    Dim strGroupId As String
    Dim dtBirth As DateTime

    ' Declare class-level variables
    Private m_colParameter As List(Of Collection)
    Private m_oReport As clsReportData
    Private m_nReportType As Integer
    Private m_strFrmId As String
    Private m_nSecurity As Integer

    ' Constants
    Private Const clngMinWidth As Integer = 7860
    Private Const clngMinHeight As Integer = 7485
    Private lotSUmmaryData As DataTable
    Private lotComments As DataTable
    Private lotContext As DataTable
    Private pensData As DataTable

    Public Sub New(ByVal sLotId As String, ByVal dtBirth As DateTime)
        ' Initialize form
        Me.Text = "Summary"
        Me.Size = New Size(800, 780)
        Me.StartPosition = FormStartPosition.CenterScreen
        prepData(sLotId, dtBirth)
        Debug.WriteLine("frmLotSummary: dtBirth " & dtBirth)
        strGroupId = sLotId
        Me.dtBirth = dtBirth
        ' Create GroupBox for Parameter Details
        Dim frameParameterDetails As New GroupBox With {
            .Text = "Lot Summary",
            .Size = New Size(700, 200),
            .Top = 20,
            .Left = 50
        }

        ' Initialize ListView
        lstvwParamDetail = New ListView With {
            .View = View.Details,
            .FullRowSelect = True,
            .Size = New Size(400, 160),
            .Top = 20,
            .Left = 10
        }

        ' Add columns to lstvwParamDetail
        lstvwParamDetail.Columns.Add("Parameter", 150, HorizontalAlignment.Left)
        lstvwParamDetail.Columns.Add("Value", 250, HorizontalAlignment.Left)

        ' Add ListView to GroupBox
        frameParameterDetails.Controls.Add(lstvwParamDetail)

        ' Add GroupBox to Form
        Me.Controls.Add(frameParameterDetails)

        ' Populate ListView with DataTable Data
        ConfigureParameters()

        ' Add lstParameter and lstvwParamDetail to the frameParameterDetails GroupBox
        ' frameParameterDetails.Controls.Add(lstParameter)
        frameParameterDetails.Controls.Add(lstvwParamDetail)

        ' Create GroupBox for Comments
        Dim frameComments As New GroupBox With {
            .Text = "LotComments",
            .Size = New Size(700, 180), ' Adjust size to fit the ListView
            .Top = frameParameterDetails.Top + frameParameterDetails.Height + 10, ' Position below frameParameterDetails
            .Left = 50
        }

        ' Create lstvwComments inside frameComments
        lstvwComments = New ListView With {
            .View = View.Details,
            .FullRowSelect = True,
            .Size = New Size(660, 140),
            .Top = 20,  ' Position inside the group box
            .Left = 10
        }

        ' Add columns to lstvwComments
        lstvwComments.Columns.Add("Date", 150, HorizontalAlignment.Left)
        lstvwComments.Columns.Add("User", 250, HorizontalAlignment.Left)
        lstvwComments.Columns.Add("Comment", 350, HorizontalAlignment.Left)

        ' Add lstvwComments to the frameComments GroupBox
        frameComments.Controls.Add(lstvwComments)

        ' Create GroupBox for Pen Details
        Dim framePenDetails As New GroupBox With {
            .Text = "Pens",
            .Size = New Size(700, 260),
            .Top = frameComments.Top + frameComments.Height + 10,
            .Left = 50
        }


        ' Create DataGridView inside framePenDetails
        penDetail = New DataGridView With {
            .Size = New Size(680, 220),
            .AllowUserToResizeColumns = True,
            .ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize,
            .Top = 20,
            .Left = 10
        }

        ' Add penDetail to framePenDetails
        framePenDetails.Controls.Add(penDetail)

        ' Initialize Buttons
        cmdOK = New Button With {.Text = "OK", .Size = New Size(100, 30)}
        cmdPrint = New Button With {.Text = "Print", .Size = New Size(100, 30)}
        cmdEdit = New Button With {.Text = "Edit", .Size = New Size(100, 30)}
        cmdDelete = New Button With {.Text = "Delete", .Size = New Size(100, 30)}

        ' Position Buttons Below Pen Details
        Dim buttonTop As Integer = framePenDetails.Top + framePenDetails.Height + 10
        cmdOK.Top = buttonTop
        cmdPrint.Top = buttonTop
        cmdEdit.Top = buttonTop
        cmdDelete.Top = buttonTop

        cmdOK.Left = 50
        cmdPrint.Left = cmdOK.Left + cmdOK.Width + 20
        cmdEdit.Left = cmdPrint.Left + cmdPrint.Width + 20
        cmdDelete.Left = cmdEdit.Left + cmdEdit.Width + 20

        ' Add GroupBoxes to the form
        Me.Controls.Add(frameParameterDetails)
        Me.Controls.Add(frameComments)
        Me.Controls.Add(framePenDetails)
        ' Add Buttons to the form (Now correctly positioned below `penDetail`)
        Me.Controls.Add(cmdOK)
        Me.Controls.Add(cmdPrint)
        Me.Controls.Add(cmdEdit)
        Me.Controls.Add(cmdDelete)
    End Sub

    ' **Handles OK button click**
    Private Sub cmdOK_Click(sender As Object, e As EventArgs) Handles cmdOK.Click
        cmdOK.Enabled = False
        ProcessState(cnSTATE_OK)
        cmdOK.Enabled = True
        Me.Close()
        Me.DialogResult = DialogResult.OK
    End Sub

    ' **Handles Print button click**
    Private Sub cmdPrint_Click(sender As Object, e As EventArgs) Handles cmdPrint.Click
        MessageBox.Show("Printing report...")
    End Sub

    ' **Handles Edit button click**
    Private Sub cmdEdit_Click(sender As Object, e As EventArgs) Handles cmdEdit.Click
        Dim oClsPen As clsPen
        Dim strUnitId As String


        ' Check if a cell is selected
        If penDetail.CurrentCell IsNot Nothing Then
            ' Check if the selected cell is in the first column (index 0)
            If penDetail.CurrentCell.ColumnIndex = 0 Then
                ' Get the value from the selected cell
                strUnitId = penDetail.CurrentCell.Value

                ' Ensure the value is not null
                If strUnitId IsNot Nothing Then
                    oClsPen = RetrievePen(strUnitId, dtBirth, strGroupId)
                    If Not (oClsPen Is Nothing) Then
                        '/*Pass the Unit object to the Defect Editor
                        Dim frmDefectEditorInstance As New frmDefectEditor
                        frmDefectEditorInstance.SetUnit(oClsPen)
                        frmDefectEditorInstance.ShowDialog()

                    End If
                    '/*Destroy our instance of the Unit
                    oClsPen = Nothing

                Else
                    MessageBox.Show("The selected cell in the first column is empty.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                End If
            Else
                MessageBox.Show("Please select a cell from the first column.", "Invalid Selection", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If
        Else
            MessageBox.Show("Please select a cell first.", "No Selection", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If
    End Sub

    Private Sub cmdDelete_Click(sender As Object, e As EventArgs) Handles cmdDelete.Click
        Dim oClsPen As clsPen
        Dim strUnitId As String

        '/*Process the state
        ProcessState(cnSTATE_DELETE)

        '/*Set the flex grid Col to 0 so we can access the Id
        penDetail.CurrentCell = penDetail.Rows(penDetail.CurrentRow.Index).Cells(0)
        '/*Get the Unit Id from the Flex Grid in
        '/*the User selected row at column 0
        strUnitId = penDetail.CurrentCell.Value.ToString()

        '/*Initiate the Edit process
        '/*Aquire the Unit from the business server
        oClsPen = RetrievePen(strUnitId, m_oReport.BirthDate, m_oReport.GroupId)

        '/*Test if a pen was actually found
        If oClsPen IsNot Nothing Then
            '/*Pass the Unit object to the Defect Editor
            frmPenDelete.SetUnit(oClsPen, 1)
            '/*Raise the defect editor window
            frmPenDelete.ShowDialog()
        End If

        '/*Destroy our instance of the Unit
        oClsPen = Nothing
    End Sub

    ' **Populate Report Data**
    Public Sub SetReport(oReport As clsReportData)
        m_oReport = oReport
        'lblGroup.Text = oReport.GroupId & " " & oReport.BirthDate.ToString("yyyy-MM-dd h:mm tt")

        ' Clear previous data
        'lstvwParamDetail.Items.Clear()
        'penDetail.Rows.Clear()

        ' Add report parameters
        'For Each param In oReport.Counts
        '    Dim item As New ListViewItem(param.Title)
        '    item.SubItems.Add(param.Value)
        '    lstvwParamDetail.Items.Add(item)
        'Next
    End Sub
    Private Sub prepData(ByVal sLotId As String, ByVal dtBirth As DateTime)
        Dim nStyle As Integer = 0

        Try
            ' get the pens in the lot
            If nStyle = 0 Then
                '/*Get a report by severity
                pensData = go_ActiveLotManager.GetLotSummaryBySeverity(sLotId, dtBirth)
            Else
                '/*Get a report by date
                pensData = go_ActiveLotManager.GetLotSummaryByDate(sLotId, CDate(dtBirth))
            End If

            ' get LotComments only if there are pen data
            If pensData IsNot Nothing AndAlso pensData.Rows.Count > 0 Then

                lotComments = go_ActiveLotManager.GetComments(sLotId, CDate(dtBirth))

                ' get Context from pen table only if there are pen data
                lotContext = go_ActiveLotManager.GetSummaryByContext(sLotId, CDate(dtBirth))


            Else
                Debug.WriteLine("fmLotSummary: prepData Failed")
            End If



        Catch ex As Exception
            ' Handle any errors
            Debug.WriteLine("Error retrieving report data: " & ex.Message)
        End Try
    End Sub
    ' Call this method in Form Load to disable the Close button
    Private Sub frmLotSummary_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DisableCloseX(Me) ' Disables the Close ("X") button when the form loads
        ' Check if DataTable is not null and set it as DataSource
        If pensData IsNot Nothing Then
            penDetail.DataSource = pensData
        Else
            MessageBox.Show("No data available to display.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If

        If lotComments IsNot Nothing AndAlso lotComments.Rows.Count > 0 Then
            PopulateListView(lotComments)
        Else
            MessageBox.Show("No comments available.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If

    End Sub
    ' Method to populate lstvwComments from DataTable
    Private Sub PopulateListView(ByVal dt As DataTable)
        ' Clear existing items
        lstvwComments.Items.Clear()
        lstvwComments.Columns.Clear()

        ' Add columns (assuming DataTable has "Date", "User", "Comment" columns)
        lstvwComments.Columns.Add("Date", 150, HorizontalAlignment.Left)
        lstvwComments.Columns.Add("User", 250, HorizontalAlignment.Left)
        lstvwComments.Columns.Add("Comment", 350, HorizontalAlignment.Left)

        ' Populate rows
        For Each row As DataRow In dt.Rows
            Dim listItem As New ListViewItem(row("CommentDate").ToString()) ' First column
            listItem.SubItems.Add(row("User").ToString()) ' Second column
            listItem.SubItems.Add(row("LotComment").ToString()) ' Third column
            lstvwComments.Items.Add(listItem)
        Next
    End Sub
    ' Configure ListBox with parameters from SummaryByContext
    Private Sub ConfigureParameters()
        ' Clear existing items
        lstvwParamDetail.Items.Clear()

        ' Add LotId
        Dim lotItem As New ListViewItem("LotId") ' First column
        lotItem.SubItems.Add(strGroupId) ' Second column (value)
        lstvwParamDetail.Items.Add(lotItem)

        ' Add Birthday (Convert DateTime to String)
        Dim birthItem As New ListViewItem("Birthday") ' First column
        birthItem.SubItems.Add(dtBirth.ToString("M/d/yyyy h:mm:ss tt")) ' Format date as string
        lstvwParamDetail.Items.Add(birthItem)

        ' Ensure DataTable is not empty
        If lotContext IsNot Nothing AndAlso lotContext.Rows.Count > 0 Then
            Dim firstRow As DataRow = lotContext.Rows(0) ' Get the first row
            ' Populate ListView with column names and values
            For Each col As DataColumn In lotContext.Columns
                Dim listItem As New ListViewItem(col.ColumnName) ' First column
                listItem.SubItems.Add(firstRow(col).ToString()) ' Second column (value)
                lstvwParamDetail.Items.Add(listItem)
            Next
        Else
            MessageBox.Show("No data available.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub

    Private Function ProcessState(ByRef nNewState As Integer) As Boolean
        Select Case nNewState
            Case cnSTATE_OK
                Me.Hide()
                '/*Release the lock on the client
                gb_InputBusyFlag = False
                ProcessState = True
        '/*If the user selects the FlexGrid and this is a Lot based report
            Case cnSTATE_FLEXGRID
                If m_nReportType = 1 Then
                    '/*Check the authority level
                    If m_nSecurity < 0 OrElse m_nSecurity >= 2 Then
                        cmdOK.Enabled = True
                        cmdEdit.Enabled = True
                        cmdDelete.Enabled = True
                        cmdPrint.Enabled = True
                        '/*set the requested state
                        m_nState = nNewState
                        ProcessState = True
                    End If
                End If
            Case Else
                cmdOK.Enabled = True
                cmdEdit.Enabled = False
                cmdDelete.Enabled = False
                cmdPrint.Enabled = True
                '/*set the requested state
                m_nState = nNewState
                ProcessState = True
        End Select
    End Function
End Class

