<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class BarDetailsForm
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
        DataGridViewBarDetails = New DataGridView()
        okButton = New Button()
        editButton = New Button()
        deleteButton = New Button()
        TableLayoutPanel1 = New TableLayoutPanel()
        LabelLotAndTime = New Label()
        CType(DataGridViewBarDetails, ComponentModel.ISupportInitialize).BeginInit()
        TableLayoutPanel1.SuspendLayout()
        SuspendLayout()
        ' 
        ' DataGridViewBarDetails
        ' 
        DataGridViewBarDetails.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewBarDetails.Dock = DockStyle.Fill
        DataGridViewBarDetails.Location = New Point(0, 0)
        DataGridViewBarDetails.Name = "DataGridViewBarDetails"
        DataGridViewBarDetails.RowHeadersWidth = 51
        DataGridViewBarDetails.Size = New Size(654, 199)
        DataGridViewBarDetails.TabIndex = 0
        ' 
        ' okButton
        ' 
        okButton.Anchor = AnchorStyles.None
        okButton.Location = New Point(62, 5)
        okButton.Name = "okButton"
        okButton.Size = New Size(94, 29)
        okButton.TabIndex = 0
        okButton.Text = "OK"
        okButton.UseVisualStyleBackColor = True
        ' 
        ' editButton
        ' 
        editButton.Anchor = AnchorStyles.None
        editButton.Location = New Point(280, 5)
        editButton.Name = "editButton"
        editButton.Size = New Size(94, 29)
        editButton.TabIndex = 1
        editButton.Text = "Edit"
        editButton.UseVisualStyleBackColor = True
        ' 
        ' deleteButton
        ' 
        deleteButton.Anchor = AnchorStyles.None
        deleteButton.Location = New Point(498, 5)
        deleteButton.Name = "deleteButton"
        deleteButton.Size = New Size(94, 29)
        deleteButton.TabIndex = 2
        deleteButton.Text = "Delete"
        deleteButton.UseVisualStyleBackColor = True
        ' 
        ' TableLayoutPanel1
        ' 
        TableLayoutPanel1.ColumnCount = 3
        TableLayoutPanel1.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 33.3333321F))
        TableLayoutPanel1.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 33.3333321F))
        TableLayoutPanel1.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 33.3333321F))
        TableLayoutPanel1.Controls.Add(deleteButton, 2, 0)
        TableLayoutPanel1.Controls.Add(okButton, 0, 0)
        TableLayoutPanel1.Controls.Add(editButton, 1, 0)
        TableLayoutPanel1.Dock = DockStyle.Bottom
        TableLayoutPanel1.Location = New Point(0, 159)
        TableLayoutPanel1.Name = "TableLayoutPanel1"
        TableLayoutPanel1.RowCount = 1
        TableLayoutPanel1.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))
        TableLayoutPanel1.Size = New Size(654, 40)
        TableLayoutPanel1.TabIndex = 3
        ' 
        ' LabelLotAndTime
        ' 
        LabelLotAndTime.Dock = DockStyle.Top
        LabelLotAndTime.Font = New Font("Microsoft Sans Serif", 9.0!, FontStyle.Bold, GraphicsUnit.Point)
        LabelLotAndTime.Location = New Point(0, 0)
        LabelLotAndTime.Name = "LabelLotAndTime"
        LabelLotAndTime.Size = New Size(654, 30)
        LabelLotAndTime.TabIndex = 4
        LabelLotAndTime.TextAlign = ContentAlignment.MiddleCenter
        ' 
        ' BarDetailsForm
        ' 
        AutoScaleDimensions = New SizeF(8.0F, 20.0F)
        AutoScaleMode = AutoScaleMode.Font
        ClientSize = New Size(654, 199)
        Controls.Add(TableLayoutPanel1)
        Controls.Add(LabelLotAndTime)
        Controls.Add(DataGridViewBarDetails)
        Name = "BarDetailsForm"
        Text = "Summary"
        CType(DataGridViewBarDetails, ComponentModel.ISupportInitialize).EndInit()
        TableLayoutPanel1.ResumeLayout(False)
        ResumeLayout(False)
    End Sub

    Friend WithEvents DataGridViewBarDetails As DataGridView
    Friend WithEvents okButton As Button
    Friend WithEvents editButton As Button
    Friend WithEvents deleteButton As Button
    Friend WithEvents TableLayoutPanel1 As TableLayoutPanel
    Friend WithEvents LabelLotAndTime As Label
End Class
