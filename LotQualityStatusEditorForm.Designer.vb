<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmLotStatusEditor
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim DataGridViewCellStyle2 As DataGridViewCellStyle = New DataGridViewCellStyle()
        lblGroupId = New TextBox()
        imgQuality = New PictureBox()
        txtBatchCount = New TextBox()
        txtUnitShipped = New TextBox()
        chk100Percent = New CheckBox()
        Label100Inspection = New Label()
        lblUnitBatch = New Label()
        lblUnitShipped = New Label()
        lstvwComments = New DataGridView()
        Column1 = New DataGridViewTextBoxColumn()
        Column2 = New DataGridViewTextBoxColumn()
        Column3 = New DataGridViewTextBoxColumn()
        cmbEvents = New ComboBox()
        txtComment = New RichTextBox()
        cmdOK = New Button()
        cmdDelete = New Button()
        cmdEdit = New Button()
        cmdAbort = New Button()
        Label3 = New GroupBox()
        lstQualityStatus = New ListBox()
        lblCurrent = New GroupBox()
        frmLotOptions = New GroupBox()
        lblCommonEvent = New GroupBox()
        lblComments = New GroupBox()
        CType(imgQuality, ComponentModel.ISupportInitialize).BeginInit()
        CType(lstvwComments, ComponentModel.ISupportInitialize).BeginInit()
        Label3.SuspendLayout()
        lblCurrent.SuspendLayout()
        frmLotOptions.SuspendLayout()
        lblCommonEvent.SuspendLayout()
        lblComments.SuspendLayout()
        SuspendLayout()
        ' 
        ' lblGroupId
        ' 
        lblGroupId.BackColor = SystemColors.ActiveBorder
        lblGroupId.Font = New Font("Segoe UI", 12F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        lblGroupId.Location = New Point(30, 12)
        lblGroupId.Name = "lblGroupId"
        lblGroupId.ReadOnly = True
        lblGroupId.Size = New Size(511, 34)
        lblGroupId.TabIndex = 0
        lblGroupId.TextAlign = HorizontalAlignment.Center
        ' 
        ' imgQuality
        ' 
        imgQuality.Location = New Point(0, 53)
        imgQuality.Name = "imgQuality"
        imgQuality.Size = New Size(66, 93)
        imgQuality.TabIndex = 2
        imgQuality.TabStop = False
        ' 
        ' txtBatchCount
        ' 
        txtBatchCount.Location = New Point(146, 94)
        txtBatchCount.Name = "txtBatchCount"
        txtBatchCount.Size = New Size(125, 27)
        txtBatchCount.TabIndex = 5
        ' 
        ' txtUnitShipped
        ' 
        txtUnitShipped.Location = New Point(146, 53)
        txtUnitShipped.Name = "txtUnitShipped"
        txtUnitShipped.Size = New Size(125, 27)
        txtUnitShipped.TabIndex = 4
        ' 
        ' chk100Percent
        ' 
        chk100Percent.AutoSize = True
        chk100Percent.Location = New Point(146, 138)
        chk100Percent.Name = "chk100Percent"
        chk100Percent.Size = New Size(103, 24)
        chk100Percent.TabIndex = 3
        chk100Percent.Text = "CheckBox1"
        chk100Percent.UseVisualStyleBackColor = True
        ' 
        ' Label100Inspection
        ' 
        Label100Inspection.AutoSize = True
        Label100Inspection.Location = New Point(10, 138)
        Label100Inspection.Name = "Label100Inspection"
        Label100Inspection.Size = New Size(117, 20)
        Label100Inspection.TabIndex = 2
        Label100Inspection.Text = "100% Inspection"
        ' 
        ' lblUnitBatch
        ' 
        lblUnitBatch.AutoSize = True
        lblUnitBatch.Location = New Point(10, 94)
        lblUnitBatch.Name = "lblUnitBatch"
        lblUnitBatch.Size = New Size(79, 20)
        lblUnitBatch.TabIndex = 1
        lblUnitBatch.Text = "Pens In Lot"
        ' 
        ' lblUnitShipped
        ' 
        lblUnitShipped.AutoSize = True
        lblUnitShipped.Location = New Point(10, 53)
        lblUnitShipped.Name = "lblUnitShipped"
        lblUnitShipped.Size = New Size(97, 20)
        lblUnitShipped.TabIndex = 0
        lblUnitShipped.Text = "Pens Shipped"
        ' 
        ' lstvwComments
        ' 
        lstvwComments.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        lstvwComments.BackgroundColor = SystemColors.ControlLightLight
        lstvwComments.CellBorderStyle = DataGridViewCellBorderStyle.None
        lstvwComments.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single
        DataGridViewCellStyle2.Alignment = DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle2.BackColor = SystemColors.ControlLight
        DataGridViewCellStyle2.Font = New Font("Segoe UI", 9F)
        DataGridViewCellStyle2.ForeColor = SystemColors.WindowText
        DataGridViewCellStyle2.SelectionBackColor = SystemColors.Highlight
        DataGridViewCellStyle2.SelectionForeColor = SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = DataGridViewTriState.True
        lstvwComments.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle2
        lstvwComments.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize
        lstvwComments.Columns.AddRange(New DataGridViewColumn() {Column1, Column2, Column3})
        lstvwComments.EnableHeadersVisualStyles = False
        lstvwComments.Location = New Point(30, 262)
        lstvwComments.Name = "lstvwComments"
        lstvwComments.RowHeadersBorderStyle = DataGridViewHeaderBorderStyle.None
        lstvwComments.RowHeadersVisible = False
        lstvwComments.RowHeadersWidth = 51
        lstvwComments.Size = New Size(511, 173)
        lstvwComments.TabIndex = 4
        ' 
        ' Column1
        ' 
        Column1.HeaderText = "Date"
        Column1.MinimumWidth = 6
        Column1.Name = "Column1"
        ' 
        ' Column2
        ' 
        Column2.HeaderText = "User"
        Column2.MinimumWidth = 6
        Column2.Name = "Column2"
        ' 
        ' Column3
        ' 
        Column3.HeaderText = "Comment"
        Column3.MinimumWidth = 6
        Column3.Name = "Column3"
        ' 
        ' cmbEvents
        ' 
        cmbEvents.FormattingEnabled = True
        cmbEvents.Location = New Point(0, 26)
        cmbEvents.Name = "cmbEvents"
        cmbEvents.Size = New Size(511, 28)
        cmbEvents.TabIndex = 5
        ' 
        ' txtComment
        ' 
        txtComment.Location = New Point(0, 26)
        txtComment.Name = "txtComment"
        txtComment.Size = New Size(511, 114)
        txtComment.TabIndex = 6
        txtComment.Text = ""
        ' 
        ' cmdOK
        ' 
        cmdOK.Location = New Point(30, 705)
        cmdOK.Name = "cmdOK"
        cmdOK.Size = New Size(94, 29)
        cmdOK.TabIndex = 7
        cmdOK.Text = "OK"
        cmdOK.UseVisualStyleBackColor = True
        ' 
        ' cmdDelete
        ' 
        cmdDelete.Location = New Point(166, 705)
        cmdDelete.Name = "cmdDelete"
        cmdDelete.Size = New Size(94, 29)
        cmdDelete.TabIndex = 8
        cmdDelete.Text = "Delete"
        cmdDelete.UseVisualStyleBackColor = True
        ' 
        ' cmdEdit
        ' 
        cmdEdit.Location = New Point(305, 705)
        cmdEdit.Name = "cmdEdit"
        cmdEdit.Size = New Size(94, 29)
        cmdEdit.TabIndex = 9
        cmdEdit.Text = "Edit"
        cmdEdit.UseVisualStyleBackColor = True
        ' 
        ' cmdAbort
        ' 
        cmdAbort.Location = New Point(447, 705)
        cmdAbort.Name = "cmdAbort"
        cmdAbort.Size = New Size(94, 29)
        cmdAbort.TabIndex = 10
        cmdAbort.Text = "Abort"
        cmdAbort.UseVisualStyleBackColor = True
        ' 
        ' Label3
        ' 
        Label3.Controls.Add(lstQualityStatus)
        Label3.Location = New Point(30, 73)
        Label3.Name = "Label3"
        Label3.Size = New Size(147, 174)
        Label3.TabIndex = 11
        Label3.TabStop = False
        Label3.Text = "Lot Quality Status"
        ' 
        ' lstQualityStatus
        ' 
        lstQualityStatus.FormattingEnabled = True
        lstQualityStatus.Location = New Point(0, 24)
        lstQualityStatus.Name = "lstQualityStatus"
        lstQualityStatus.Size = New Size(147, 144)
        lstQualityStatus.TabIndex = 0
        ' 
        ' lblCurrent
        ' 
        lblCurrent.Controls.Add(imgQuality)
        lblCurrent.Location = New Point(183, 73)
        lblCurrent.Name = "lblCurrent"
        lblCurrent.Size = New Size(66, 174)
        lblCurrent.TabIndex = 12
        lblCurrent.TabStop = False
        lblCurrent.Text = "Current"
        ' 
        ' frmLotOptions
        ' 
        frmLotOptions.Controls.Add(txtUnitShipped)
        frmLotOptions.Controls.Add(txtBatchCount)
        frmLotOptions.Controls.Add(lblUnitShipped)
        frmLotOptions.Controls.Add(Label100Inspection)
        frmLotOptions.Controls.Add(chk100Percent)
        frmLotOptions.Controls.Add(lblUnitBatch)
        frmLotOptions.Location = New Point(255, 73)
        frmLotOptions.Name = "frmLotOptions"
        frmLotOptions.Size = New Size(286, 174)
        frmLotOptions.TabIndex = 13
        frmLotOptions.TabStop = False
        frmLotOptions.Text = "Shipment Data"
        ' 
        ' lblCommonEvent
        ' 
        lblCommonEvent.Controls.Add(cmbEvents)
        lblCommonEvent.Location = New Point(30, 450)
        lblCommonEvent.Name = "lblCommonEvent"
        lblCommonEvent.Size = New Size(511, 64)
        lblCommonEvent.TabIndex = 14
        lblCommonEvent.TabStop = False
        lblCommonEvent.Text = "Common Events"
        ' 
        ' lblComments
        ' 
        lblComments.Controls.Add(txtComment)
        lblComments.Location = New Point(30, 531)
        lblComments.Name = "lblComments"
        lblComments.Size = New Size(511, 146)
        lblComments.TabIndex = 15
        lblComments.TabStop = False
        lblComments.Text = "Comments"
        ' 
        ' frmLotStatusEditor
        ' 
        AutoScaleDimensions = New SizeF(8F, 20F)
        AutoScaleMode = AutoScaleMode.Font
        ClientSize = New Size(582, 753)
        Controls.Add(lblComments)
        Controls.Add(lblCommonEvent)
        Controls.Add(frmLotOptions)
        Controls.Add(lblCurrent)
        Controls.Add(Label3)
        Controls.Add(cmdAbort)
        Controls.Add(cmdEdit)
        Controls.Add(cmdDelete)
        Controls.Add(cmdOK)
        Controls.Add(lstvwComments)
        Controls.Add(lblGroupId)
        Name = "frmLotStatusEditor"
        Text = "Lot Quality Status Editor"
        CType(imgQuality, ComponentModel.ISupportInitialize).EndInit()
        CType(lstvwComments, ComponentModel.ISupportInitialize).EndInit()
        Label3.ResumeLayout(False)
        lblCurrent.ResumeLayout(False)
        frmLotOptions.ResumeLayout(False)
        frmLotOptions.PerformLayout()
        lblCommonEvent.ResumeLayout(False)
        lblComments.ResumeLayout(False)
        ResumeLayout(False)
        PerformLayout()
    End Sub

    Friend WithEvents lblGroupId As TextBox
    Friend WithEvents imgQuality As PictureBox
    Friend WithEvents lblUnitShipped As Label
    Friend WithEvents lblUnitBatch As Label
    Friend WithEvents Label100Inspection As Label
    Friend WithEvents txtBatchCount As TextBox
    Friend WithEvents txtUnitShipped As TextBox
    Friend WithEvents chk100Percent As CheckBox
    Friend WithEvents lstvwComments As DataGridView
    Friend WithEvents Column1 As DataGridViewTextBoxColumn
    Friend WithEvents Column2 As DataGridViewTextBoxColumn
    Friend WithEvents Column3 As DataGridViewTextBoxColumn
    Friend WithEvents cmbEvents As ComboBox
    Friend WithEvents txtComment As RichTextBox
    Friend WithEvents cmdOK As Button
    Friend WithEvents cmdDelete As Button
    Friend WithEvents cmdEdit As Button
    Friend WithEvents cmdAbort As Button
    Friend WithEvents Label3 As GroupBox
    Friend WithEvents lstQualityStatus As ListBox
    Friend WithEvents lblCurrent As GroupBox
    Friend WithEvents frmLotOptions As GroupBox
    Friend WithEvents lblCommonEvent As GroupBox
    Friend WithEvents lblComments As GroupBox
End Class
