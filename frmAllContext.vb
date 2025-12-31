Imports System.Runtime.InteropServices.JavaScript.JSType
Imports System.Drawing.Color

Public Class frmAllContext
    'Inherits Form
    Inherits CustomForm
    Private Shared _instance As frmAllContext

    Private TabStrip1 As TabControl
    Private cmdOK As Button
    Private cmdHelp As Button
    Private cmdAbort As Button
    Private Frame1_User, Frame1_Material, Frame1_PhysicalLine As Panel

    ' Controls inside "Material"
    Private cmbFamily, cmbPartName, cmbPartNumber, cmbRunType As ComboBox
    Private txtTFLot, txtExpId As TextBox

    ' Controls inside "Physical Line"
    Private txtProdDate As TextBox
    Private cmbAccumulator, cmbSource, cmbLineId, cmbLineType As ComboBox

    ' Controls inside "User"
    Public cmbAuthority, cmbShift, cmbOperator As ComboBox
    '/*Storage for the ComboBox to trigger login
    Private m_strUser As String
    Private m_nUserIndex As Integer
    '/*Windowing variables
    Private m_clsWindow As clsWindowState
    Private m_frmParent As Form
    '/*Text field colors
    Dim clngEnabledText As Color = SystemColors.Window
    Dim clngDisabledText As Color = SystemColors.GrayText
    '/*Functions to facilitate Production Date Adjust option
    Private m_bProductionDate As Boolean
    Private m_strOriginalProdDate As String
    '/*Security object
    Private WithEvents m_oSecurity As Security
    'Attribute m_oSecurity.VB_VarHelpID = -1
    '/*One-shot variable to override activity associated
    '/*with the reading of teh operator by the registry
    Private m_bOneShotOperator As Boolean

    ' -- Constant ID for calls to LotManagerChange CSB
    Const m_csMe As String = "ContextEditor"

    '/*This forms current stack ID
    Private m_strFrmId As String

    Public Property StrFrmId As String
        Get
            Return m_strFrmId
        End Get
        Set(value As String)
            m_strFrmId = value
        End Set
    End Property


    '/*Accessor to set the OneShotOperator override
    Public Sub SetOneShotOperator()
        m_bOneShotOperator = True
    End Sub
    '/*Accessor to read the state of the flag and reset it
    Public Function GetOneShotOperator() As Boolean
        GetOneShotOperator = m_bOneShotOperator 'Return value
        m_bOneShotOperator = False              'Turn off
    End Function
    Public Property ProductionDate As String
        Get
            Return txtProdDate.Text
        End Get
        Set(value As String)
            If m_bProductionDate Then
                m_strOriginalProdDate = value
                txtProdDate.Text = value
            Else
                txtProdDate.Text = ""
            End If
        End Set
    End Property
    Public Shared Function GetInstance() As frmAllContext
        If _instance Is Nothing OrElse _instance.IsDisposed Then
            _instance = New frmAllContext()
        End If
        Return _instance
    End Function
    Private Sub New()
        ' 窗体基础设置
        Me.Text = "All Context"
        Me.Size = New Size(500, 580) ' 适当放大窗体，增加留白
        Me.Name = "frmAllContext"
        Me.Padding = New Padding(10) ' 窗体边缘留白
        Me.MinimumSize = New Size(450, 500) ' 宽度不小于450，高度不小于500

        ' === 初始化TabControl ===
        TabStrip1 = New TabControl With {
            .Location = New Point(10, 10),
            .Size = New Size(465, 450), ' 适配窗体大小
            .Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right Or AnchorStyles.Bottom, ' 随窗体缩放
            .Padding = New Point(10, 5) ' 标签页内边距
        }
        TabStrip1.TabPages.Add("User")
        TabStrip1.TabPages.Add("Material")
        TabStrip1.TabPages.Add("Physical Line")
        Me.Controls.Add(TabStrip1)

        ' === 初始化User标签页 ===
        Frame1_User = New Panel With {
            .Dock = DockStyle.Fill,
            .Padding = New Padding(0) ' 增加内边距，让内容不紧贴边缘,左上右下
        }

        ' 使用TableLayoutPanel布局User标签页控件（2列：标签+输入控件）
        Dim tlpUser As New TableLayoutPanel With {
            .Dock = DockStyle.None, ' 只占据顶部必要空间
            .Padding = New Padding(0),
            .RowCount = 2,
            .ColumnCount = 2,
            .Margin = New Padding(0), ' 顶部margin使内容垂直居中
            .Height = 120
        }
        ' 列宽比例：标签列占30%，输入控件列占70%
        'tlpUser.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 30))
        'tlpUser.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 70))
        tlpUser.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize)) ' 标签列：自动适应内容
        tlpUser.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100)) ' 输入列：占剩余宽度

        ' 行高自动适应内容
        tlpUser.RowStyles.Add(New RowStyle(SizeType.Percent, 50)) ' 第一行高度
        tlpUser.RowStyles.Add(New RowStyle(SizeType.Percent, 50)) ' 第二行高度

        ' 添加User标签页控件
        Dim lblOperator As New Label With {
            .Text = "Operator",
            .TextAlign = ContentAlignment.MiddleLeft,
            .Anchor = AnchorStyles.Left Or AnchorStyles.Right,
            .Font = New Font("Segoe UI", 10, FontStyle.Bold), ' 加粗字体
            .ForeColor = Color.Black, ' 黑色文字
            .Margin = New Padding(10, 10, 5, 10),
            .MinimumSize = New Size(115, 0)
        }
        tlpUser.Controls.Add(lblOperator, 0, 0)

        cmbOperator = New ComboBox With {
            .DropDownStyle = ComboBoxStyle.DropDownList,
            .Anchor = AnchorStyles.Left Or AnchorStyles.Right,
            .Margin = New Padding(5, 10, 10, 10),
            .Height = 35 ' 增加高度
        }
        AddHandler cmbOperator.SelectedIndexChanged, AddressOf cmbOperator_SelectedIndexChanged
        tlpUser.Controls.Add(cmbOperator, 1, 0)

        Dim lblShift As New Label With {
            .Text = "Shift",
            .TextAlign = ContentAlignment.MiddleLeft,
            .Anchor = AnchorStyles.Left Or AnchorStyles.Right,
            .Font = New Font("Segoe UI", 10, FontStyle.Bold), ' 加粗字体
            .ForeColor = Color.Black, ' 黑色文字
            .Margin = New Padding(10, 10, 5, 10),
            .MinimumSize = New Size(115, 0)
        }
        tlpUser.Controls.Add(lblShift, 0, 1)

        cmbShift = New ComboBox With {
            .DropDownStyle = ComboBoxStyle.DropDownList,
            .Anchor = AnchorStyles.Left Or AnchorStyles.Right,
            .Margin = New Padding(5, 10, 10, 10),
            .Height = 35 ' 增加高度
        }
        tlpUser.Controls.Add(cmbShift, 1, 1)

        ' 隐藏的Authority控件（保持原有逻辑）
        cmbAuthority = New ComboBox With {
            .DropDownStyle = ComboBoxStyle.DropDownList,
            .Visible = False
        }
        AddHandler cmbAuthority.SelectedIndexChanged, AddressOf cmbAuthority_SelectedIndexChanged
        tlpUser.Controls.Add(cmbAuthority)

        AddHandler Frame1_User.SizeChanged, Sub(sender, e)
                                                ' 左右留白（各1/10）
                                                Dim sideMargin As Integer = Frame1_User.ClientSize.Width \ 10
                                                tlpUser.Width = Frame1_User.ClientSize.Width - 2 * sideMargin

                                                ' 计算严格垂直居中的Y坐标
                                                Dim centerY As Integer = (Frame1_User.ClientSize.Height - tlpUser.Height) \ 2
                                                ' 向上调整10%（基于可用垂直空间的10%）
                                                Dim availableVerticalSpace As Integer = Frame1_User.ClientSize.Height - tlpUser.Height
                                                Dim adjustUp As Integer = CInt(availableVerticalSpace * 0.1) ' 10%的向上偏移量
                                                Dim finalY As Integer = centerY - adjustUp ' 最终Y坐标 = 居中Y - 偏移量

                                                ' 设置位置（水平居中，垂直偏高10%）
                                                tlpUser.Location = New Point(sideMargin, finalY)
                                            End Sub

        Frame1_User.Controls.Add(tlpUser)
        Frame1_User.PerformLayout()
        TabStrip1.TabPages(0).Controls.Add(Frame1_User)

        ' === 初始化Material标签页 ===
        Frame1_Material = New Panel With {
            .Dock = DockStyle.Fill,
            .Padding = New Padding(0)
        }
        ' 使用 TableLayoutPanel 布局 Material 标签页控件
        Dim tlpMaterial As New TableLayoutPanel With {
            .Dock = DockStyle.None, ' 不贴边，手动定位
            .Padding = New Padding(0),
            .RowCount = 5, ' 5 行控件
            .ColumnCount = 2,
            .Margin = New Padding(0),
            .Height = 200 ' 总高度（5 行 ×40，预留边距）
        }

        ' 列宽比例与 User 标签页一致（30%:70%）
        'tlpMaterial.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 30))
        'tlpMaterial.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 70))
        tlpMaterial.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize)) ' 标签列：自动适应内容
        tlpMaterial.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100)) ' 输入列：占剩余宽度

        ' 行高调整为 40（增加高度，避免内容拥挤）
        For i As Integer = 0 To 4
            tlpMaterial.RowStyles.Add(New RowStyle(SizeType.Percent, 20)) ' 5 行平均分配高度
        Next

        ' 添加 Material 标签页控件（统一样式：黑色加粗、边距一致）
        ' 1. Run Type
        Dim lblRunType As New Label With {
            .Text = "Run Type",
            .TextAlign = ContentAlignment.MiddleLeft,
            .Anchor = AnchorStyles.Left Or AnchorStyles.Right,
            .Font = New Font("Segoe UI", 10, FontStyle.Bold), ' 与 User 标签页字体一致
            .ForeColor = Color.Black,
            .Margin = New Padding(10, 10, 5, 10),
            .MinimumSize = New Size(115, 0)
        }
        tlpMaterial.Controls.Add(lblRunType, 0, 0)

        cmbRunType = New ComboBox With {
            .DropDownStyle = ComboBoxStyle.DropDownList,
            .Anchor = AnchorStyles.Left Or AnchorStyles.Right,
            .Margin = New Padding(5, 10, 10, 10),
            .Height = 35 ' 与 User 标签页下拉框高度一致
        }
        AddHandler cmbRunType.SelectedIndexChanged, AddressOf cmbRunType_SelectedIndexChanged
        tlpMaterial.Controls.Add(cmbRunType, 1, 0)

        ' 2. Experiment ID
        Dim lblExpId As New Label With {
            .Text = "Experiment ID",
            .TextAlign = ContentAlignment.MiddleLeft,
            .Anchor = AnchorStyles.Left Or AnchorStyles.Right,
            .Font = New Font("Segoe UI", 10, FontStyle.Bold),
            .ForeColor = Color.Black,
            .Margin = New Padding(10, 10, 5, 10),
            .MinimumSize = New Size(115, 0)
        }
        tlpMaterial.Controls.Add(lblExpId, 0, 1)

        txtExpId = New TextBox With {
            .Anchor = AnchorStyles.Left Or AnchorStyles.Right,
            .Margin = New Padding(5, 10, 10, 10),
            .Height = 35 ' 文本框高度与下拉框一致
        }
        tlpMaterial.Controls.Add(txtExpId, 1, 1)

        ' 3. Part Number
        Dim lblPartNumber As New Label With {
            .Text = "Part Number",
            .TextAlign = ContentAlignment.MiddleLeft,
            .Anchor = AnchorStyles.Left Or AnchorStyles.Right,
            .Font = New Font("Segoe UI", 10, FontStyle.Bold),
            .ForeColor = Color.Black,
            .Margin = New Padding(10, 10, 5, 10),
            .MinimumSize = New Size(115, 0)
        }
        tlpMaterial.Controls.Add(lblPartNumber, 0, 2)

        cmbPartNumber = New ComboBox With {
            .DropDownStyle = ComboBoxStyle.DropDownList,
            .Anchor = AnchorStyles.Left Or AnchorStyles.Right,
            .Margin = New Padding(5, 10, 10, 10),
            .Height = 35
        }
        AddHandler cmbPartNumber.SelectedIndexChanged, AddressOf cmbPartNumber_SelectedIndexChanged
        tlpMaterial.Controls.Add(cmbPartNumber, 1, 2)

        ' 4. Part Name
        Dim lblPartName As New Label With {
            .Text = "Part Name",
            .TextAlign = ContentAlignment.MiddleLeft,
            .Anchor = AnchorStyles.Left Or AnchorStyles.Right,
            .Font = New Font("Segoe UI", 10, FontStyle.Bold),
            .ForeColor = Color.Black,
            .Margin = New Padding(10, 10, 5, 10),
            .MinimumSize = New Size(115, 0)
        }
        tlpMaterial.Controls.Add(lblPartName, 0, 3)

        cmbPartName = New ComboBox With {
            .DropDownStyle = ComboBoxStyle.DropDownList,
            .Anchor = AnchorStyles.Left Or AnchorStyles.Right,
            .Margin = New Padding(5, 10, 10, 10),
            .Height = 35
        }
        AddHandler cmbPartName.SelectedIndexChanged, AddressOf cmbPartName_SelectedIndexChanged
        tlpMaterial.Controls.Add(cmbPartName, 1, 3)

        ' 5. Thin Film Lot
        Dim lblTFLot As New Label With {
            .Text = "Thin Film Lot",
            .TextAlign = ContentAlignment.MiddleLeft,
            .Anchor = AnchorStyles.Left Or AnchorStyles.Right,
            .Font = New Font("Segoe UI", 10, FontStyle.Bold),
            .ForeColor = Color.Black,
            .Margin = New Padding(10, 10, 5, 10),
            .MinimumSize = New Size(115, 0)
        }
        tlpMaterial.Controls.Add(lblTFLot, 0, 4)

        txtTFLot = New TextBox With {
            .Anchor = AnchorStyles.Left Or AnchorStyles.Right,
            .Margin = New Padding(5, 10, 10, 10),
            .Height = 35
        }
        tlpMaterial.Controls.Add(txtTFLot, 1, 4)

        ' 隐藏的 Family 控件
        cmbFamily = New ComboBox With {
            .DropDownStyle = ComboBoxStyle.DropDownList,
            .Visible = False
        }
        tlpMaterial.Controls.Add(cmbFamily)

        ' 核心：与 User 标签页一致的定位逻辑（左右留 1/10 边距，垂直居中偏上 10%）
        AddHandler Frame1_Material.SizeChanged, Sub(sender, e)
                                                    ' 左右留白（各 1/10 宽度）
                                                    Dim sideMargin As Integer = Frame1_Material.ClientSize.Width \ 10
                                                    tlpMaterial.Width = Frame1_Material.ClientSize.Width - 2 * sideMargin

                                                    ' 垂直位置：居中偏上 10%
                                                    Dim centerY As Integer = (Frame1_Material.ClientSize.Height - tlpMaterial.Height) \ 2
                                                    Dim availableVerticalSpace As Integer = Frame1_Material.ClientSize.Height - tlpMaterial.Height
                                                    Dim adjustUp As Integer = CInt(availableVerticalSpace * 0.1) ' 向上偏移 10%
                                                    Dim finalY As Integer = centerY - adjustUp

                                                    ' 设置位置
                                                    tlpMaterial.Location = New Point(sideMargin, finalY)
                                                End Sub

        ' 添加到容器并刷新布局
        Frame1_Material.Controls.Add(tlpMaterial)
        Frame1_Material.PerformLayout()
        TabStrip1.TabPages(1).Controls.Add(Frame1_Material)

        ' === 初始化Physical Line标签页 ===
        Frame1_PhysicalLine = New Panel With {
            .Dock = DockStyle.Fill,
            .Padding = New Padding(0)
        }
        ' 使用 TableLayoutPanel 布局 Physical Line 标签页控件
        Dim tlpPhysical As New TableLayoutPanel With {
            .Dock = DockStyle.None, ' 取消填充，手动定位
            .Padding = New Padding(0),
            .RowCount = 5, ' 5 行控件
            .ColumnCount = 2,
            .Margin = New Padding(0),
            .Height = 200 ' 总高度适配 5 行内容（含边距）
        }

        ' 列宽比例统一为 30%:70%（与其他标签页一致）
        'tlpPhysical.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 30))
        'tlpPhysical.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 70))
        tlpPhysical.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize)) ' 标签列：自动适应内容
        tlpPhysical.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100)) ' 输入列：占剩余宽度

        ' 行高使用百分比分配（5 行各占 20%），确保内容垂直居中
        For i As Integer = 0 To 4
            tlpPhysical.RowStyles.Add(New RowStyle(SizeType.Percent, 20))
        Next

        ' 添加 Physical Line 标签页控件（统一样式）
        ' 1. Line Type
        Dim lblLineType As New Label With {
            .Text = "Line Type",
            .TextAlign = ContentAlignment.MiddleLeft,
            .Anchor = AnchorStyles.Left Or AnchorStyles.Right,
            .Font = New Font("Segoe UI", 10, FontStyle.Bold), ' 黑色加粗字体
            .ForeColor = Color.Black,
            .Margin = New Padding(10, 10, 5, 10), ' 统一边距
            .MinimumSize = New Size(120, 0) ' 最小宽度100（根据文字长度调整）
        }
        tlpPhysical.Controls.Add(lblLineType, 0, 0)

        cmbLineType = New ComboBox With {
            .DropDownStyle = ComboBoxStyle.DropDownList,
            .Anchor = AnchorStyles.Left Or AnchorStyles.Right,
            .Margin = New Padding(5, 10, 10, 10),
            .Height = 35 ' 控件高度统一
        }
        tlpPhysical.Controls.Add(cmbLineType, 1, 0)

        ' 2. Line ID
        Dim lblLineId As New Label With {
            .Text = "Line ID",
            .TextAlign = ContentAlignment.MiddleLeft,
            .Anchor = AnchorStyles.Left Or AnchorStyles.Right,
            .Font = New Font("Segoe UI", 10, FontStyle.Bold),
            .ForeColor = Color.Black,
            .Margin = New Padding(10, 10, 5, 10),
            .MinimumSize = New Size(120, 0) ' 最小宽度100（根据文字长度调整）
        }
        tlpPhysical.Controls.Add(lblLineId, 0, 1)

        cmbLineId = New ComboBox With {
            .DropDownStyle = ComboBoxStyle.DropDownList,
            .Anchor = AnchorStyles.Left Or AnchorStyles.Right,
            .Margin = New Padding(5, 10, 10, 10),
            .Height = 35
        }
        tlpPhysical.Controls.Add(cmbLineId, 1, 1)

        ' 3. Source
        Dim lblSource As New Label With {
            .Text = "Source",
            .TextAlign = ContentAlignment.MiddleLeft,
            .Anchor = AnchorStyles.Left Or AnchorStyles.Right,
            .Font = New Font("Segoe UI", 10, FontStyle.Bold),
            .ForeColor = Color.Black,
            .Margin = New Padding(10, 10, 5, 10),
            .MinimumSize = New Size(120, 0) ' 最小宽度100（根据文字长度调整）
        }
        tlpPhysical.Controls.Add(lblSource, 0, 2)

        cmbSource = New ComboBox With {
            .DropDownStyle = ComboBoxStyle.DropDownList,
            .Anchor = AnchorStyles.Left Or AnchorStyles.Right,
            .Margin = New Padding(5, 10, 10, 10),
            .Height = 35
        }
        tlpPhysical.Controls.Add(cmbSource, 1, 2)

        ' 4. Accumulator
        Dim lblAccumulator As New Label With {
            .Text = "Accumulator",
            .TextAlign = ContentAlignment.MiddleLeft,
            .Anchor = AnchorStyles.Left Or AnchorStyles.Right,
            .Font = New Font("Segoe UI", 10, FontStyle.Bold),
            .ForeColor = Color.Black,
            .Margin = New Padding(10, 10, 5, 10),
            .MinimumSize = New Size(120, 0) ' 最小宽度100（根据文字长度调整）
        }
        tlpPhysical.Controls.Add(lblAccumulator, 0, 3)

        cmbAccumulator = New ComboBox With {
            .DropDownStyle = ComboBoxStyle.DropDownList,
            .Anchor = AnchorStyles.Left Or AnchorStyles.Right,
            .Margin = New Padding(5, 10, 10, 10),
            .Height = 35
        }
        tlpPhysical.Controls.Add(cmbAccumulator, 1, 3)

        ' 5. Production Date
        Dim lblProdDate As New Label With {
            .Text = "Production Date",
            .TextAlign = ContentAlignment.MiddleLeft,
            .Anchor = AnchorStyles.Left Or AnchorStyles.Right,
            .Font = New Font("Segoe UI", 10, FontStyle.Bold),
            .ForeColor = Color.Black,
            .Margin = New Padding(10, 10, 5, 10),
            .MinimumSize = New Size(120, 0) ' 最小宽度100（根据文字长度调整）
        }
        tlpPhysical.Controls.Add(lblProdDate, 0, 4)

        txtProdDate = New TextBox With {
            .Anchor = AnchorStyles.Left Or AnchorStyles.Right,
            .Margin = New Padding(5, 10, 10, 10),
            .Height = 35, ' 文本框高度与下拉框统一
            .Text = "txtProdDate"
        }
        tlpPhysical.Controls.Add(txtProdDate, 1, 4)

        ' 核心：与其他标签页完全一致的定位逻辑
        AddHandler Frame1_PhysicalLine.SizeChanged, Sub(sender, e)
                                                        ' 左右留 1/10 边距
                                                        Dim sideMargin As Integer = Frame1_PhysicalLine.ClientSize.Width \ 10
                                                        tlpPhysical.Width = Frame1_PhysicalLine.ClientSize.Width - 2 * sideMargin

                                                        ' 垂直居中偏上 10%
                                                        Dim centerY As Integer = (Frame1_PhysicalLine.ClientSize.Height - tlpPhysical.Height) \ 2
                                                        Dim availableVerticalSpace As Integer = Frame1_PhysicalLine.ClientSize.Height - tlpPhysical.Height
                                                        Dim adjustUp As Integer = CInt(availableVerticalSpace * 0.1) ' 向上偏移 10%
                                                        Dim finalY As Integer = centerY - adjustUp

                                                        ' 设置位置
                                                        tlpPhysical.Location = New Point(sideMargin, finalY)
                                                    End Sub

        ' 添加到容器并刷新布局
        Frame1_PhysicalLine.Controls.Add(tlpPhysical)
        Frame1_PhysicalLine.PerformLayout()
        TabStrip1.TabPages(2).Controls.Add(Frame1_PhysicalLine)

        ' === OK按钮 ===
        cmdOK = New Button With {
            .Text = "OK",
            .Size = New Size(120, 40),
            .FlatStyle = FlatStyle.Flat
        }
        cmdOK.FlatAppearance.BorderSize = 1
        cmdOK.FlatAppearance.BorderColor = Color.LightGray

        AddHandler cmdOK.Click, AddressOf cmdOK_Click

        cmdHelp = New Button With {
        .Text = "Help",
        .Size = New Size(120, 40),
        .FlatStyle = FlatStyle.Flat
        }
        cmdHelp.FlatAppearance.BorderSize = 1
        cmdHelp.FlatAppearance.BorderColor = Color.LightGray
        'AddHandler cmdHelp.Click, AddressOf cmdHelp_Click

        cmdAbort = New Button With {
        .Text = "Abort",
        .Size = New Size(120, 40),
        .FlatStyle = FlatStyle.Flat
        }
        cmdAbort.FlatAppearance.BorderSize = 1
        cmdAbort.FlatAppearance.BorderColor = Color.LightGray
        AddHandler cmdAbort.Click, AddressOf cmdAbort_Click

        Me.Controls.Add(cmdOK)
        Me.Controls.Add(cmdHelp)
        Me.Controls.Add(cmdAbort)

        ' 处理窗体大小变化时的按钮位置调整
        AddHandler Me.SizeChanged, AddressOf AdjustButtonPositions
        ' 初始调整位置
        AdjustButtonPositions(Nothing, EventArgs.Empty)
    End Sub


    'Modifications:
    '   08-15-2001 New call to make sure the TopForm is not blocked.
    Private Sub frmAllContext_Paint(sender As Object, e As PaintEventArgs) Handles Me.Paint
        mdlWindow.CheckWhosTop(Me)
    End Sub

    '=======================================================
    'Routine: Property Let EnableProductionDate( bool )
    'Purpose: This is a property to allow for changing the production date.
    'It also handles configuring the screen.
    '
    'Globals: go_clsSystemSettings
    '
    'Input:None
    '
    'Return:None
    '
    'Tested:
    '
    'Modifications:
    '   01-22-2001 - This is checking the the system
    '   setting for this and only updating when the two
    '   are out of sync to avoid setting up a cycle. The
    '   reason is that this could be updated by another
    '   caller or by the system object who is already
    '   calling this property.
    '
    '=======================================================
    Public Property EnableProductionDate As Boolean
        Set(bProdDate As Boolean)
            '/*Set the txt box to enabled
            If bProdDate Then
                '/*Set the text box's properties
                txtProdDate.Enabled = True
                txtProdDate.BackColor = clngEnabledText
            Else
                '/*Set the text box's properties
                txtProdDate.Enabled = False
                txtProdDate.BackColor = clngDisabledText
            End If
            '/*Clear the boxes content
            txtProdDate.Text = ""
            '/*Set the system setting incase this was set by an caller other
            '/*then the system object
            If go_clsSystemSettings.bProductionDateAdjust <> bProdDate Then
                go_clsSystemSettings.bProductionDateAdjust = bProdDate
            End If
            '/*Set are internal toggle
            m_bProductionDate = bProdDate
        End Set
        Get
            Return m_bProductionDate
        End Get
    End Property



    '
    '=======================================================
    'Routine: cmbAuthority_Click()
    'Purpose: This is triggered when the Index of the dropdown
    'list is changed since it isn't actually visible to
    'the users. If login is required, any necessary Authority
    'level changes will originate here from our list
    'of values taken agianst the Security objects values.
    '
    'Globals: go_Security - The global security object for
    '         this application.
    '
    'Input:None
    '
    'Return:None
    '
    'Tested:
    '
    'Modifications:
    '   04-27-1999 As written for Pass1.8
    '
    '
    '=======================================================
    Private Sub cmbAuthority_SelectedIndexChanged(sender As Object, e As EventArgs)
        If cmbAuthority.SelectedIndex = -1 Then Exit Sub
        go_Security.strUserAuthority = cmbAuthority.Items(cmbAuthority.SelectedIndex).ToString()
    End Sub

    '
    '=======================================================
    'Routine: cmbOperator_Click()
    ' *Standard command click event
    'Purpose: If a login is not required then the list is
    'changed and the Authority level with it. However, if
    'a login is required the change of name is handled by
    'the password form which will feed back the name on a
    'successful Name-Password pair. This will also handle
    'the change in the Authority level.
    '
    'Note: The static Boolean is used to exit this rountine
    'when it is triggered during the reset of the current index.
    '
    'Globals: go_Security.bLoginRequired - True:Login is required
    '
    'Input:None
    '
    'Return:None
    '
    'Tested:
    '
    'Modifications:
    '   04-27-1999 As written for Pass1.8
    '
    '=======================================================
    Private Sub cmbOperator_SelectedIndexChanged(sender As Object, e As EventArgs)
        Static bActive As Boolean
        Dim strNewUser As String

        '/*Avert re-enterant triggering
        If bActive Then Exit Sub

        If cmbOperator.SelectedIndex = -1 Then Exit Sub

        '/*See if a login is required to change this field
        If go_Security.bLoginRequired And Me.Visible And frmPassword.Visible = False Then
            bActive = True
            '/*Revert to previous user
            strNewUser = cmbOperator.Items(cmbOperator.SelectedIndex)
            frmPassword.PasswordDisplay(strNewUser, Me, False)
            bActive = False
        Else
            '/*Sync the Authority with this dropdowns Index
            If cmbOperator.SelectedIndex > -1 Then
                cmbAuthority.SelectedIndex = cmbOperator.SelectedIndex
                m_nUserIndex = cmbOperator.SelectedIndex
            End If
        End If
    End Sub

    Private Sub cmbPartName_SelectedIndexChanged(sender As Object, e As EventArgs)
        '/*Sync the Part Name with this dropdowns Index
        If cmbPartName.SelectedIndex > -1 Then
            cmbPartNumber.SelectedIndex = cmbPartName.SelectedIndex
            cmbFamily.SelectedIndex = cmbPartName.SelectedIndex
        End If
    End Sub

    Private Sub cmbPartNumber_SelectedIndexChanged(sender As Object, e As EventArgs)
        '/*Sync the Part Name with this dropdowns Index
        If cmbPartNumber.SelectedIndex > -1 Then
            cmbPartName.SelectedIndex = cmbPartNumber.SelectedIndex
            cmbFamily.SelectedIndex = cmbPartNumber.SelectedIndex
        End If
    End Sub

    Private Sub cmbRunType_SelectedIndexChanged(sender As Object, e As EventArgs)
        Dim bEnabled As Boolean

        If cmbRunType.SelectedIndex = -1 Then Exit Sub

        '/*Set the enable/disable of the Engineering Lot ID
        bEnabled = txtExpId.Enabled = IIf(cmbRunType.Text = "Engineering", True, False)
        '/*Set the fillcolor
        If bEnabled Then
            txtExpId.BackColor = clngEnabledText
        Else
            txtExpId.BackColor = clngDisabledText
            txtExpId.Text = ""
        End If
    End Sub



    Private Sub cmdAbort_Click(sender As Object, e As EventArgs)
        '/*Close this window and reactivate
        HideWindow()
        '/*Release the lock on the client
        gb_InputBusyFlag = False
    End Sub

    Private Sub cmdOK_Click(sender As Object, e As EventArgs)
        MessageBox.Show("OK clicked!")

        '/*Update the context
        If UpdateContext() Then
            '/*Update the server's copy of the context
            frmServerSink.UpdateServerContext()
            '/*Close this window and reactivate
            HideWindow()
            '/*Set a warning if there is no Open lot
            If String.IsNullOrEmpty(go_Context.GroupId) Then
                '/*Send an alert message to the user
                Dim strMsg As String = "frmAllContext.cmdOK()" & Chr(13) & "Warning! The current " &
            "Context has no open Lot. Go to the Lot Manager screen to create a New Lot."
                frmMessage.GenerateMessage(strMsg, "Warning")
            End If
        Else
            '/*Release the lock on the client
            gb_InputBusyFlag = False
        End If

    End Sub

    '=======================================================
    'Routine: frmAllContext.Form_Load
    'Purpose: Default routine. This has the code for preping
    '         the location of the Frames and which one is
    '         the starting Frame.
    '
    'Globals:None
    '
    'Input:None
    '
    'Return:None
    '
    '
    'Modifications:
    '   9-23-1998 As written for Pass1.0
    '
    '   02-08-1999 Added guarenteed base RunTypes.
    '
    '   04-27-1999 Added in reference to the security
    '   object to catch the Security change event.
    '
    '   10-15-2001 Chris Barker: Add check against system
    '   parameters to see if the Default Run Types has
    '   been suppressed.
    '=======================================================
    Private Sub Form_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Dim nLoop As Integer

        'TabStrip1.Top = 50
        'TabStrip1.Left = 50

        ''/*Adjust the width and the height of the screen
        'Me.Width = TabStrip1.Width + 190
        'Me.Height = cmdOK.Top + cmdOK.Height + (cmdOK.Top - TabStrip1.Height) * 2

        ''/*Loop through the Frames and set their positions
        'Frame1_User.Location = New Point(TabStrip1.Location.X, TabStrip1.Location.Y)
        'Frame1_User.FlatStyle = FlatStyle.Flat
        'Frame1_User.ForeColor = Color.Black

        'Frame1_Material.Location = New Point(TabStrip1.Location.X, TabStrip1.Location.Y)
        'Frame1_Material.FlatStyle = FlatStyle.Flat
        'Frame1_Material.ForeColor = Color.Black

        'Frame1_PhysicalLine.Location = New Point(TabStrip1.Location.X, TabStrip1.Location.Y)
        'Frame1_PhysicalLine.FlatStyle = FlatStyle.Flat
        'Frame1_PhysicalLine.ForeColor = Color.Black

        ''/*Set the top most Frame
        'Frame1_User.BringToFront()

        '/*Tidy up some text boxes
        txtTFLot.Text = ""
        txtExpId.Text = ""
        'txtProdDate.Text = "txtProdDate"

        '/*We need to guarentee that the base RunTypes will be available
        If go_clsSystemSettings.bUseDefaultRunType Then
            AddRunType("Production")
            AddRunType("Engineering")
        End If

        '/*Hide the family type
        'cmbFamily.Visible = False
        '/*Connect to the security object
        m_oSecurity = go_Security
        'In case we are connecting after the object startup
        If m_oSecurity IsNot Nothing Then
            m_oSecurity_SecurityChange(m_oSecurity.nAuthority)
        End If

        ShowWindow()
    End Sub

    '
    '=======================================================
    'Routine: frmAllContext.Form_QueryUnLoad
    'Purpose: This is the default Query_Unload and has
    '         been setup to trap the System Menu Close
    '         event. When this occurs it Asks the operator
    '         whether they want to save the changes.
    '         **To Do: Set Private Variable to Trap when Form
    '         is Dirty
    '
    'Globals:None
    '
    'Input: Cancel - Allows event to continue if 0; set
    '                to any other value to circumvent
    '                the form closing.
    '      UnloadMode - Indicates who Raised this routine.
    '                   We are only trapping the System Menu.
    '
    'Return:None
    '
    'Modifications:
    '   9-23-1998 As written for Pass1.0
    '
    '   01-28-1999 Added call to run a CSB to update the
    '   status bar (provided there is a functional one
    '   resident in the system). C. Barker
    '=======================================================
    Private Sub frmAllContext_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        Const cstrmsg As String = "Update the changes to the context?"

        ' Only handle user/system close (not app exit, etc.)
        If e.CloseReason = CloseReason.UserClosing Then
            ' Uncomment below to prompt user for confirmation if needed:
            'If MessageBox.Show(cstrmsg, "Update Changes", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.OK Then
            ' Update the context
            If UpdateContext() Then
                ' Process the update if needed
            End If
            'End If

            ' Prevent the form from actually closing
            e.Cancel = True
            ' Hide this form and reactivate any other windows
            HideWindow()
            ' Unlock the client
            gb_InputBusyFlag = False
        ElseIf e.Cancel Then
            PersistToUserRegistry()
        End If
    End Sub

    '/*This is the security change event
    Private Sub m_oSecurity_SecurityChange(ByVal nNewAuthority As Integer) Handles m_oSecurity.SecurityChange
        '/*Level 2 access
        If nNewAuthority > -1 AndAlso nNewAuthority < 3 Then
            cmbLineType.Enabled = True
            cmbLineId.Enabled = True
            cmbSource.Enabled = True
            cmbRunType.Enabled = False
            txtExpId.Enabled = False
            cmbPartNumber.Enabled = False
            cmbPartName.Enabled = False
            txtTFLot.Enabled = False

            '2.0.14
            'read only operator used able to change user name
            cmbOperator.Enabled = False
            cmbOperator.Enabled = True  ' 注：此行可能是代码遗留问题，连续的 False/True 可能需要确认逻辑

            cmbShift.Enabled = False
            cmbAccumulator.Enabled = False
            EnableProductionDate = False
            '/*Level 4 access
        ElseIf nNewAuthority > 3 OrElse nNewAuthority = -1 Then
            cmbLineType.Enabled = True
            cmbLineId.Enabled = True
            cmbSource.Enabled = True
            cmbRunType.Enabled = True
            txtExpId.Enabled = True
            cmbPartNumber.Enabled = True
            cmbPartName.Enabled = True
            txtTFLot.Enabled = True
            cmbOperator.Enabled = True
            cmbShift.Enabled = True
            cmbAccumulator.Enabled = True
            '/*Make sure the date is valid
            If go_clsSystemSettings.bProductionDateAdjust Then
                EnableProductionDate = True
            End If
        End If
    End Sub

    '=======================================================
    'Routine: frmAllContext.UPdateContext
    'Purpose: To act as a centrel point for propigating
    '         the Users screen changes to the Applicaiton's
    '         Objects (screen & collection)
    '
    'Globals:None
    '
    'Input:None
    '
    'Return:None
    '
    'Modifications:
    '   9-22-1998 As written for Pass1.0
    '
    '   11-12-1998 Added the production date field
    '   and switched this to a function to feedback
    '   update failure (particularly for Adjust Date).
    '
    '   03-03-1999 Added procedure to apply a Default Item
    '   Type to the Context for generating units.
    '
    '   08-24-2001 Icluded call to CSB for change v1.1.5
    '   chris barker
    '
    '=======================================================
    Public Function UpdateContext() As Boolean
        Dim strMsg As String
        Dim strSource As String
        Dim strLineNumber As String
        Dim strLineType As String
        Dim bLotManagerChanged As Boolean = False
        Dim vrtOpenLot As Object
        Dim sLotId As String = ""
        Dim dtLotBirth As Date = Date.MinValue
        Dim bdirty As Boolean = False

        ' Read out the combo settings that will be used more than once
        'strLineType = cmbLineType.Text
        'strLineNumber = cmbLineId.Text
        'strSource = cmbSource.Text
        strLineType = cmbLineType.Items(cmbLineType.SelectedIndex).ToString()
        strLineNumber = cmbLineId.Items(cmbLineId.SelectedIndex).ToString()
        strSource = cmbSource.Items(cmbSource.SelectedIndex).ToString()

        ' Check for unsaved changes
        With go_Context
            bdirty = False

            If .PartName <> cmbPartName.Items(cmbPartName.SelectedIndex).ToString() Then bdirty = True
            If .PartNumber <> cmbPartNumber.Items(cmbPartNumber.SelectedIndex).ToString() Then bdirty = True
            If .RunType <> cmbRunType.Items(cmbRunType.SelectedIndex).ToString() Then bdirty = True
            If .Source <> strSource Then bdirty = True
            If .Operator <> cmbOperator.Items(cmbOperator.SelectedIndex).ToString() Then bdirty = True
            If .LineType <> strLineType Then bdirty = True
            If .LineNumber <> strLineNumber Then bdirty = True
            If .Accumulator <> cmbAccumulator.Items(cmbAccumulator.SelectedIndex).ToString() Then bdirty = True
            If .Shift <> cmbShift.Items(cmbShift.SelectedIndex).ToString() Then bdirty = True
            If .ExperimentId <> txtExpId.Text Then bdirty = True
            If .ThinFilmLot <> txtTFLot.Text Then bdirty = True
            If .ProductType <> cmbFamily.Items(cmbFamily.SelectedIndex).ToString() Then bdirty = True

            If bdirty Then
                Dim strMessage As String = "You have changed the Context of Nextcap Application. Click Yes to Confirm."
                If MessageBox.Show(strMessage, "Context Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then
                    UpdateContext = False
                    Exit Function
                End If
            Else
                UpdateContext = True
                Exit Function
            End If

            ' If the Station-Key has changed attempt to redirect the ActiveLotManager
            If strLineType <> .LineType OrElse strLineNumber <> .LineNumber OrElse strSource <> .Source Then
                If Not mdlLotManager.SetActiveLotManager(strLineType, Convert.ToInt32(strLineNumber), strSource) Then
                    mdlMain.MainErrorHandler("frmContextAll.UpdateContext()", "The requested combination of Line, Line Number and Source is not valid. Change or Abort please.")
                    Exit Function
                End If
                bLotManagerChanged = True
            End If

            ' Set the global object context
            .PartName = cmbPartName.Items(cmbPartName.SelectedIndex).ToString()
            .PartNumber = cmbPartNumber.Items(cmbPartNumber.SelectedIndex).ToString()
            .RunType = cmbRunType.Items(cmbRunType.SelectedIndex).ToString()
            .Source = strSource
            .Operator = cmbOperator.Items(cmbOperator.SelectedIndex).ToString()
            .LineType = strLineType
            .LineNumber = strLineNumber
            .Accumulator = cmbAccumulator.Items(cmbAccumulator.SelectedIndex).ToString()
            .Shift = cmbShift.Items(cmbShift.SelectedIndex).ToString()
            .ExperimentId = txtExpId.Text
            .ThinFilmLot = txtTFLot.Text
            .ProductType = cmbFamily.Items(cmbFamily.SelectedIndex).ToString()

            ' Set the default item type
            If gcol_StationKeys IsNot Nothing Then
                .DefaultUnitType = gcol_StationKeys.GetDefaultItemType(strLineType, Convert.ToInt32(strLineNumber), strSource)
            End If
        End With

        ' Set the Adjusted production date if used
        If m_bProductionDate Then
            If IsDate(txtProdDate.Text) Then
                If m_strOriginalProdDate <> txtProdDate.Text Then
                    go_Context.ProductionDate = Convert.ToDateTime(txtProdDate.Text)
                End If
            Else
                strMsg = "The date in 'Production Date' is not" & vbCrLf &
                     "in a recognized format. Please" & vbCrLf &
                     "enter it again."
                MainErrorHandler(Me.Name & ".UpdateContext()", strMsg)
                Exit Function
            End If
        End If

        ' If the flag for CSB used to update the status bar is enabled, execute that CSB
        If go_Context.bCSBUsed Then mdlSAX.ExecuteCSB_UpdateStatusBar()

        ' Call context change CSB if a LotManager changed
        If bLotManagerChanged Then
            vrtOpenLot = go_ActiveLotManager.GetOpenLot()
            If vrtOpenLot(0, 0) <> "ERROR" Then
                sLotId = vrtOpenLot(0, 0)
                dtLotBirth = vrtOpenLot(1, 0)
            End If
            ExecuteCSB_LotManagerChange(sLotId, dtLotBirth, m_csMe)
        End If

        ' Persist the context information to the registry
        PersistToUserRegistry()
        UpdateContext = True

    End Function
    '=======================================================
    'Routine: frmAllContext.ShowWindow()
    'Purpose: Show this Form either as simulated Modal or
    '         just to show the form. This also handles setup
    '         of the screen prior to the visible state
    '
    'Globals:None
    '
    'Input: bModal - Show this screen in Modal Form
    '
    '       frmParent - The parent form to set focus on
    '       when this form is closed.
    '
    'Return:None
    '
    '
    'Modifications:
    '   10-15-1998 As written for Pass1.1
    '
    '
    '=======================================================
    Public Sub ShowWindow()
        '/*Initialize this window
        WindowInit()
        m_strFrmId = mdlWindow.AddForm(Me)
    End Sub


    '=======================================================
    'Routine: frmAllContext.HideWindow()
    'Purpose: Hide this Form and remove any applied modality
    '
    'Globals:None
    '
    'Input:None
    '
    'Return:None
    '
    '
    'Modifications:
    '   10-15-1998 As written for Pass1.1
    '
    '
    '=======================================================
    Public Sub HideWindow()
        ' Remove the form from any custom tracking (if needed)
        'mdlWindow.RemoveForm(m_strFrmId)
        m_strFrmId = ""
        ' Hide the form
        Me.Hide()
    End Sub
    '
    '=======================================================
    'Routine: frmAllContext.Match_(str)
    'Purpose: These routines search the specified combo
    '         box and set the list index.
    '
    'Globals:None
    '
    'Input: str - The item in the list to be matched.
    '
    'Return:None
    '
    'Modifications:
    '   11-25-1998 As written for Pass1.5
    '
    '
    '=======================================================
    ' Match the accumulator in the ComboBox and select it, or select the first item if not found
    Private Sub MatchAccumulator(ByVal strCompare As String)
        Dim bFound As Boolean = False

        Try
            ' Scan the list for the item
            For nItems As Integer = 0 To cmbAccumulator.Items.Count - 1
                If String.Equals(cmbAccumulator.Items(nItems).ToString(), strCompare, StringComparison.OrdinalIgnoreCase) Then
                    cmbAccumulator.SelectedIndex = nItems
                    bFound = True
                    Exit For
                End If
            Next

            ' Set the default if not found or empty
            If Not bFound OrElse String.IsNullOrEmpty(strCompare) Then
                If cmbAccumulator.Items.Count > 0 Then
                    cmbAccumulator.SelectedIndex = 0
                End If
            End If
        Catch ex As Exception
            ' Optionally log or handle the error
        End Try
    End Sub

    '
    ' Match the line type in the ComboBox and select it, or select the first item if not found
    Private Sub MatchLineType(ByVal strCompare As String)
        Dim bFound As Boolean = False

        Try
            ' Scan the list for the item
            For nItems As Integer = 0 To cmbLineType.Items.Count - 1
                If String.Equals(cmbLineType.Items(nItems).ToString(), strCompare, StringComparison.OrdinalIgnoreCase) Then
                    cmbLineType.SelectedIndex = nItems
                    bFound = True
                    Exit For
                End If
            Next

            ' Set the default if not found or empty
            If Not bFound OrElse String.IsNullOrEmpty(strCompare) Then
                If cmbLineType.Items.Count > 0 Then
                    cmbLineType.SelectedIndex = 0
                End If
            End If
        Catch ex As Exception
            ' Optionally log or handle the error
        End Try
    End Sub

    '
    ' Match the line number in the ComboBox and select it, or select the first item if not found
    Private Sub MatchLineNumber(ByVal strCompare As String)
        Dim bFound As Boolean = False

        Try
            ' Scan the list for the item
            For nItems As Integer = 0 To cmbLineId.Items.Count - 1
                If String.Equals(cmbLineId.Items(nItems).ToString(), strCompare, StringComparison.OrdinalIgnoreCase) Then
                    cmbLineId.SelectedIndex = nItems
                    bFound = True
                    Exit For
                End If
            Next

            ' Set the default if not found or empty
            If Not bFound OrElse String.IsNullOrEmpty(strCompare) Then
                If cmbLineId.Items.Count > 0 Then
                    cmbLineId.SelectedIndex = 0
                End If
            End If
        Catch ex As Exception
            ' Optionally log or handle the error
        End Try
    End Sub

    '
    ' Match the source in the ComboBox and select it, or select the first item if not found
    Private Sub MatchSource(ByVal strCompare As String)
        Dim bFound As Boolean = False

        Try
            ' Scan the list for the item
            For nItems As Integer = 0 To cmbSource.Items.Count - 1
                If String.Equals(cmbSource.Items(nItems).ToString(), strCompare, StringComparison.OrdinalIgnoreCase) Then
                    cmbSource.SelectedIndex = nItems
                    bFound = True
                    Exit For
                End If
            Next

            ' Set the default if not found or empty
            If Not bFound OrElse String.IsNullOrEmpty(strCompare) Then
                If cmbSource.Items.Count > 0 Then
                    cmbSource.SelectedIndex = 0
                End If
            End If
        Catch ex As Exception
            ' Optionally log or handle the error
        End Try
    End Sub

    '
    '----------------------------------------------------------
    'Match the operator name if in our list.
    'Modifications:
    '   06-03-1999 Moved the ListIndex=0 to the end
    '   and put the constraint of no match (bFound = False)
    '   on it to control over triggering of the cmbAuthority
    '   object. CB
    '
    '-----------------------------------------------------------
    ' Match the operator name in the ComboBox and select it, or select the first item if not found
    Private Sub MatchOperator(ByVal strCompare As String)
        Dim bFound As Boolean = False

        Try
            ' Scan the list for the item
            For nItems As Integer = 0 To cmbOperator.Items.Count - 1
                If String.Equals(cmbOperator.Items(nItems).ToString(), strCompare, StringComparison.OrdinalIgnoreCase) Then
                    cmbOperator.SelectedIndex = nItems
                    bFound = True
                    Exit For
                End If
            Next

            ' Set the default if not found or empty
            If Not bFound OrElse String.IsNullOrEmpty(strCompare) Then
                If cmbOperator.Items.Count > 0 Then
                    cmbOperator.SelectedIndex = 0
                End If
            End If
        Catch ex As Exception
            ' Optionally log or handle the error
        End Try
    End Sub

    Private Sub MatchShift(ByVal strCompare As String)
        Try
            ' Set the default
            cmbShift.SelectedIndex = 0
            ' Scan the list for the item
            For nItems As Integer = 0 To cmbShift.Items.Count - 1
                If String.Equals(cmbShift.Items(nItems).ToString(), strCompare, StringComparison.OrdinalIgnoreCase) Then
                    cmbShift.SelectedIndex = nItems
                    Exit For
                End If
            Next
        Catch ex As Exception
            ' Optionally log or handle the error
        End Try
    End Sub
    Private Sub MatchRunType(ByVal strCompare As String)
        Try
            ' Set the default
            cmbRunType.SelectedIndex = 0
            ' Scan the list for the item
            For nItems As Integer = 0 To cmbRunType.Items.Count - 1
                If String.Equals(cmbRunType.Items(nItems).ToString(), strCompare, StringComparison.OrdinalIgnoreCase) Then
                    cmbRunType.SelectedIndex = nItems
                    Exit For
                End If
            Next
        Catch
            ' Swallow errors as in original code
        End Try
    End Sub

    Private Sub MatchPartName(ByVal strCompare As String)
        Try
            ' Set the default
            cmbPartName.SelectedIndex = 0
            ' Scan the list for the item
            For nItems As Integer = 0 To cmbPartName.Items.Count - 1
                If String.Equals(cmbPartName.Items(nItems).ToString(), strCompare, StringComparison.OrdinalIgnoreCase) Then
                    cmbPartName.SelectedIndex = nItems
                    Exit For
                End If
            Next
        Catch
            ' Swallow errors as in original code
        End Try
    End Sub

    Private Sub MatchPartNumber(ByVal strCompare As String)
        Try
            ' Set the default
            cmbPartNumber.SelectedIndex = 0
            ' Scan the list for the item
            For nItems As Integer = 0 To cmbPartNumber.Items.Count - 1
                If String.Equals(cmbPartNumber.Items(nItems).ToString(), strCompare, StringComparison.OrdinalIgnoreCase) Then
                    cmbPartNumber.SelectedIndex = nItems
                    Exit For
                End If
            Next
        Catch
            ' Swallow errors as in original code
        End Try
    End Sub

    Private Sub MatchFamily(ByVal strCompare As String)
        Try
            ' Set the default
            cmbFamily.SelectedIndex = 0
            ' Scan the list for the item
            For nItems As Integer = 0 To cmbFamily.Items.Count - 1
                If String.Equals(cmbFamily.Items(nItems).ToString(), strCompare, StringComparison.OrdinalIgnoreCase) Then
                    cmbFamily.SelectedIndex = nItems
                    Exit For
                End If
            Next
        Catch
            ' Swallow errors as in original code
        End Try
    End Sub

    '
    '=============================================================
    'Routine: Addxxxxx(str)
    'Purpose: This next set of properties is for adding
    '         items to the context lists. The general procedure
    '         is that we scan the the combo box to see
    '         if we already have the item. If we don't have
    '         the item we add it to the list. See the note
    '         for AddPart(str,str) which deviates due
    '         to adding two items.
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
    '==============================================================
    '/*Set list item
    Public Sub AddLineType(strData As String)
        ' Check if the item already exists (case-insensitive)
        For Each item As Object In cmbLineType.Items
            If String.Equals(item.ToString(), strData, StringComparison.OrdinalIgnoreCase) Then
                Exit Sub
            End If
        Next
        ' If not found, add the item
        cmbLineType.Items.Add(strData)
    End Sub
    '/*Set list item
    Public Sub AddLineNumber(strData As String)
        ' Check if the item already exists (case-insensitive)
        For Each item As Object In cmbLineId.Items
            If String.Equals(item.ToString(), strData, StringComparison.OrdinalIgnoreCase) Then
                Exit Sub
            End If
        Next
        ' If not found, add the item
        cmbLineId.Items.Add(strData)
    End Sub
    '/*Set list item
    Public Sub AddSource(strData As String)
        ' Check if the item already exists (case-insensitive)
        For Each item As Object In cmbSource.Items
            If String.Equals(item.ToString(), strData, StringComparison.OrdinalIgnoreCase) Then
                Exit Sub
            End If
        Next
        ' If not found, add the item
        cmbSource.Items.Add(strData)


    End Sub
    '/*Set list item
    Public Sub AddAccumulator(strData As String)
        ' Check if the item already exists (case-insensitive)
        For Each item As Object In cmbAccumulator.Items
            If String.Equals(item.ToString(), strData, StringComparison.OrdinalIgnoreCase) Then
                Exit Sub
            End If
        Next
        ' If not found, add the item
        cmbAccumulator.Items.Add(strData)


    End Sub
    '/*Set list item
    Public Sub AddShifts(strData As String)
        ' Check if the item already exists (case-insensitive)
        For Each item As Object In cmbShift.Items
            If String.Equals(item.ToString(), strData, StringComparison.OrdinalIgnoreCase) Then
                Exit Sub
            End If
        Next
        ' If not found, add the item
        cmbShift.Items.Add(strData)

    End Sub
    '/*Set list item
    Public Sub AddRunType(strData As String)
        ' Check if the item already exists (case-insensitive)
        For Each item As Object In cmbRunType.Items
            If String.Equals(item.ToString(), strData, StringComparison.OrdinalIgnoreCase) Then
                Exit Sub
            End If
        Next
        ' If not found, add the item
        cmbRunType.Items.Add(strData)


    End Sub
    '/*Set list item
    Public Sub AddOperator(strUser As String, strAuthority As String)
        ' Check if the item already exists (case-insensitive)
        For Each item As Object In cmbOperator.Items
            If String.Equals(item.ToString(), strUser, StringComparison.OrdinalIgnoreCase) Then
                Exit Sub
            End If
        Next
        ' If not found, add the item
        cmbOperator.Items.Add(strUser)
        cmbAuthority.Items.Add(strAuthority)

    End Sub
    '
    '/*--------------------------------------------------------
    '/*Note:
    '/*This is slightly different. I figured we can gurantee
    '/*that the part number will be unique, but that the
    '/*name used may not necssarily be so we only enforce
    '/*the PartNumber when loading these.
    '/*---------------------------------------------------------
    Public Sub AddPart(strName As String, strNumber As String, strFamily As String)
        ' Check if the item already exists (case-insensitive)
        For Each item As Object In cmbPartNumber.Items
            If String.Equals(item.ToString(), strNumber, StringComparison.OrdinalIgnoreCase) Then
                Exit Sub
            End If
        Next
        ' If not found, add the item
        cmbPartNumber.Items.Add(strNumber)
        cmbPartName.Items.Add(strName)
        cmbFamily.Items.Add(strFamily)


    End Sub
    '
    '=======================================================
    'Routine: frmAllContext.WindowInit()
    'Purpose: Perform any required window initialization
    '         bfore showing the screen to the user.
    '
    'Globals:None
    '
    'Input:None
    '
    'Return:None
    '
    '
    'Modifications:
    '   9-23-1998 As written for Pass1.0
    '
    '   05-24-1999 Added prep for the produciton date
    '   when used. CB
    '
    '   07-06-2001 Chris Barker
    '   Added in code to enumerate fields by parameter
    '   encoding and test for enabled/disabled against
    '   a user hooked CSB (ContextLock).
    '
    '=======================================================
    Private Sub WindowInit()
        ' Make sure the form is correct
        SyncContextToForm()

        ' Get the current production date if the feature is enabled
        If m_bProductionDate Then
            If go_ActiveLotManager IsNot Nothing Then
                Dim prodDateObj = go_ActiveLotManager.ProductionDate
                If IsDate(prodDateObj) Then
                    ProductionDate = CDate(prodDateObj).ToString("MM-dd-yyyy")
                Else
                    ProductionDate = ""
                End If
            End If
        End If

        'If the ContextLock CSB Is used, Then test Each parameter against it For locking state
        If mdlSAX.ContextLockUsed Then
            ' Operator field
            If mdlSAX.ExecuteCSB_ContextLock("Operator") Then
                cmbOperator.Enabled = False
            Else
                cmbOperator.Enabled = True
            End If

            ' Shift
            If mdlSAX.ExecuteCSB_ContextLock("Shift") Then
                cmbShift.Enabled = False
            Else
                cmbShift.Enabled = True
            End If

            ' Part Number and Part Name
            If mdlSAX.ExecuteCSB_ContextLock("PartNumber") Then
                cmbPartNumber.Enabled = False
                cmbPartName.Enabled = False
            Else
                cmbPartNumber.Enabled = True
                cmbPartName.Enabled = True
            End If

            ' RunType
            If mdlSAX.ExecuteCSB_ContextLock("RunType") Then
                cmbRunType.Enabled = False
            Else
                cmbRunType.Enabled = True
            End If

            ' Experiment Id, only test when Exp Id is permitted
            If cmbRunType.Text = "Engineering" Then
                If mdlSAX.ExecuteCSB_ContextLock("ExpId") Then
                    txtExpId.Enabled = False
                Else
                    txtExpId.Enabled = True
                End If
            End If

            ' Thin Film Lot Id
            If mdlSAX.ExecuteCSB_ContextLock("ThinFilm") Then
                txtTFLot.Enabled = False
            Else
                txtTFLot.Enabled = True
            End If

            ' Line Type
            If mdlSAX.ExecuteCSB_ContextLock("LineType") Then
                cmbLineType.Enabled = False
            Else
                cmbLineType.Enabled = True
            End If

            ' Line Number
            If mdlSAX.ExecuteCSB_ContextLock("LineNumber") Then
                cmbLineId.Enabled = False
            Else
                cmbLineId.Enabled = True
            End If

            ' Line Source
            If mdlSAX.ExecuteCSB_ContextLock("Source") Then
                cmbSource.Enabled = False
            Else
                cmbSource.Enabled = True
            End If

            ' Production Date
            If m_bProductionDate Then
                If mdlSAX.ExecuteCSB_ContextLock("ProductionDate") Then
                    txtProdDate.Enabled = False
                Else
                    txtProdDate.Enabled = True
                End If
            End If
        End If

        ' Add any other screen prep calls here
    End Sub
    '=======================================================
    'Routine: frmAllContext.SyncContextToForm()
    'Purpose: This syncronizes the Context object with
    '         the values listed on the form. If the Context
    '         object is blank the Match_ routines default to
    '         the 0th element of the list.
    '
    'Globals: go_Context - The global context object.
    '
    'Input:None
    '
    'Return:None
    '
    'Modifications:
    '   12-03-1998 As written for Pass1.5
    '
    '   06-03-1999 Wrapped the operator match with
    '   the condition that we are not running
    '   in a secure mode.
    '=======================================================
    Public Sub SyncContextToForm()
        ' Ensure the selection boxes match the global context object
        With go_Context
            MatchAccumulator(.Accumulator)
            MatchLineType(.LineType)
            MatchLineNumber(.LineNumber)
            MatchSource(.Source)
            ' Bypass if secured setup
            If Not (go_clsSystemSettings.bLogOutAtShift Or go_Security.bLoginRequired) Then
                MatchOperator(.Operator)
                ' Set private member (used to trigger log-in changes)
                m_strUser = .Operator
            End If
            MatchShift(.Shift)
            MatchRunType(.RunType)
            txtExpId.Text = .ExperimentId
            MatchPartName(.PartName)
            txtTFLot.Text = .ThinFilmLot
        End With
    End Sub

    '
    '=======================================================
    'Routine: PersistToUserRegistry()
    'Purpose: This writes all of the current Context
    'settings to the User Registry.
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
    '   05-26-1999 As written for Pass1.8
    '
    '   08-09-2001 Changed to persist the registry from
    '   the global object instead of this form which
    '   may be out of date. Chris Barker
    '=======================================================
    Public Sub PersistToUserRegistry()
        Dim strRegName As String
        Dim strSection As String

        Try
            strRegName = mdlMain.cstrRegName
            strSection = Me.Name

            With go_Context
                ' Write out the settings
                SaveSettingValue(strRegName, strSection, "Operator", .Operator)
                SaveSettingValue(strRegName, strSection, "Shift", .Shift)
                SaveSettingValue(strRegName, strSection, "RunType", .RunType)
                ' Only set one of these since it will trigger all others
                SaveSettingValue(strRegName, strSection, "PartName", .PartName)
                SaveSettingValue(strRegName, strSection, "ExperimentId", .ExperimentId)
                SaveSettingValue(strRegName, strSection, "ThinFilmLot", .ThinFilmLot)
                ' The location tab
                SaveSettingValue(strRegName, strSection, "ProductionDate", .ProductionDate.ToString())
                SaveSettingValue(strRegName, strSection, "LineType", .LineType)
                SaveSettingValue(strRegName, strSection, "LineId", .LineNumber)
                SaveSettingValue(strRegName, strSection, "Source", .Source)
                SaveSettingValue(strRegName, strSection, "Accumulator", .Accumulator)
                SaveSettingValue(strRegName, strSection, "PartNumber", .PartNumber)
                SaveSettingValue(strRegName, strSection, "Family", .ProductType)

            End With
        Catch
            ' Swallow errors as in original On Error Resume Next
        End Try
    End Sub

    '
    '=======================================================
    'Routine: ReadInRegistry()
    'Purpose: This reads in the registry values for the
    'Context settings and attempts to set them. If the item
    'in the registry can not be found it is set to the 0th
    'list item or "" for fields.
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
    '   05-03-1999 As written for Pass1.8
    '
    '   06-03-1999 Included bypass for reading in the
    '   persisted user if we are using security or forcing
    '   logout (should be the same situation).
    '=======================================================
    Public Sub ReadInRegistry()
        Dim strRegName As String
        Dim strSection As String
        Dim strSource As String
        Dim strLineNumber As String
        Dim strLineType As String

        Try
            strRegName = mdlMain.cstrRegName
            strSection = Me.Name

            ' Read in the registry settings and pass them to the Match routines
            MatchAccumulator(GetSettingValue(strRegName, strSection, "Accumulator", ""))
            MatchLineType(GetSettingValue(strRegName, strSection, "LineType", ""))
            MatchLineNumber(GetSettingValue(strRegName, strSection, "LineId", ""))
            MatchSource(GetSettingValue(strRegName, strSection, "Source", ""))

            If Not (go_clsSystemSettings.bLogOutAtShift Or go_Security.bLoginRequired) Then
                MatchOperator(GetSettingValue(strRegName, strSection, "Operator", ""))
                m_strUser = GetSettingValue(strRegName, strSection, "Operator", "")
            Else
                MatchOperator(go_Context.Operator)
                m_strUser = go_Context.Operator
            End If

            MatchShift(GetSettingValue(strRegName, strSection, "Shift", ""))
            MatchRunType(GetSettingValue(strRegName, strSection, "RunType", ""))
            txtExpId.Text = GetSettingValue(strRegName, strSection, "ExperimentId", "")
            MatchPartName(GetSettingValue(strRegName, strSection, "PartName", ""))
            txtTFLot.Text = GetSettingValue(strRegName, strSection, "ThinFilmLot", "")
            If m_bProductionDate Then
                txtProdDate.Text = GetSettingValue(strRegName, strSection, "ProductionDate", "")
            End If
            MatchPartNumber(GetSettingValue(strRegName, strSection, "PartNumber", ""))
            MatchFamily(GetSettingValue(strRegName, strSection, "Family", ""))

            ' Duplicate of UpdateContext to allow for removing any user feedback
            With go_Context
                strLineType = cmbLineType.Items(cmbLineType.SelectedIndex).ToString()
                strLineNumber = cmbLineId.Items(cmbLineId.SelectedIndex).ToString()
                strSource = cmbSource.Items(cmbSource.SelectedIndex).ToString()


                If strLineType <> .LineType OrElse strLineNumber <> .LineNumber OrElse strSource <> .Source Then
                    If gcol_LotManagers IsNot Nothing Then
                        mdlLotManager.SetActiveLotManager(strLineType, Convert.ToInt32(strLineNumber), strSource)
                    End If
                End If

                .PartName = cmbPartName.Items(cmbPartName.SelectedIndex).ToString()
                .PartNumber = cmbPartNumber.Items(cmbPartNumber.SelectedIndex).ToString() ' Already set elsewhere
                .RunType = cmbRunType.Items(cmbRunType.SelectedIndex).ToString()

                If strSource Is Nothing Then
                    Debug.WriteLine("Source is null, setting to default value.")
                Else
                    .Source = strSource.ToString()
                End If

                .Operator = cmbOperator.Items(cmbOperator.SelectedIndex).ToString()
                .LineType = strLineType
                .LineNumber = strLineNumber
                .Accumulator = cmbAccumulator.Items(cmbAccumulator.SelectedIndex).ToString()
                .Shift = cmbShift.Items(cmbShift.SelectedIndex).ToString()
                .ExperimentId = txtExpId.Text
                .ThinFilmLot = txtTFLot.Text
                .ProductType = cmbFamily.Items(cmbFamily.SelectedIndex).ToString()
                If gcol_StationKeys IsNot Nothing Then
                    .DefaultUnitType = gcol_StationKeys.GetDefaultItemType(strLineType, Convert.ToInt32(strLineNumber), strSource)
                End If
            End With

            If m_bProductionDate Then
                If Date.TryParse(txtProdDate.Text, Nothing) Then
                    If m_strOriginalProdDate <> txtProdDate.Text Then
                        go_Context.ProductionDate = Convert.ToDateTime(txtProdDate.Text)
                    End If
                End If
            End If
        Catch ex As Exception
            Debug.WriteLine($"error {ex}")
            ' Swallow errors as in original On Error Resume Next
        End Try
    End Sub

    '
    '=======================================================
    'Routine: frmAllContext.ChangeShift(str)
    'Purpose: This is the interface for changing the
    'shift when the client is running in automatic change
    'mode.
    '   clsSystemSettings.strShiftFunction="Disable|AutoChange
    '   |MsgOnly"
    'Called from frmSupervisorSink.TEMP_ChangeShift()
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
    '   05-03-1999 As written for Pass1.8
    '
    '
    '=======================================================
    Public Sub ChangeShift(ByVal strShift As String)
        Try
            ' Set the form item
            MatchShift(strShift)
            ' Update the context
            go_Context.Shift = cmbShift.Items(cmbShift.SelectedIndex).ToString()
            ' If the flag for CSB used to update the status bar is enabled, execute that CSB
            If go_Context.bCSBUsed Then
                mdlSAX.ExecuteCSB_UpdateStatusBar()
            End If
        Catch
            ' Swallow any error, as in original On Error Resume Next
        End Try
    End Sub

    '=======================================================
    'Routine: frmAllCOntext.EnableShiftChange(b)
    'Purpose: This configures the enabled property on the
    'shift change dropdown.
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
    '   05-03-1999 As written for Pass1.8
    '
    '
    '=======================================================
    Public Sub EnableShiftChange(ByRef bEnabled As Boolean)
        cmbShift.Enabled = bEnabled
    End Sub

    Private Function GetSettingValue(appName As String, section As String, key As String, defaultValue As Object) As Object
        Dim regKey As Microsoft.Win32.RegistryKey
        Dim value As Object

        Try
            ' 打开或创建注册表项
            regKey = Microsoft.Win32.Registry.CurrentUser.CreateSubKey($"SOFTWARE\{appName}\{section}")

            ' 获取值或默认值
            value = regKey.GetValue(key, defaultValue)

            ' 关闭注册表项
            regKey.Close()

            Return value
        Catch ex As Exception
            Debug.WriteLine($"读取注册表设置时出错: {ex.Message}")
            Return defaultValue
        End Try
    End Function

    Private Sub SaveSettingValue(appName As String, section As String, key As String, defaultValue As Object)
        Dim regKey As Microsoft.Win32.RegistryKey

        Try
            ' 打开或创建注册表项
            regKey = Microsoft.Win32.Registry.CurrentUser.CreateSubKey($"SOFTWARE\{appName}\{section}")

            ' 获取值或默认值
            regKey.SetValue(key, defaultValue)

            ' 关闭注册表项
            regKey.Close()

        Catch ex As Exception
            Debug.WriteLine($"读取注册表设置时出错: {ex.Message}")
        End Try
    End Sub

    ' 调整按钮位置使其在底部居中排列
    Private Sub AdjustButtonPositions(sender As Object, e As EventArgs)
        ' 按钮之间的间距
        Dim spacing As Integer = 15
        ' 底部边距
        Dim bottomMargin As Integer = 20

        ' 计算三个按钮的总宽度（含间距）
        Dim totalWidth As Integer = cmdOK.Width + cmdHelp.Width + cmdAbort.Width + (spacing * 2)
        ' 计算起始X坐标（居中显示）
        Dim startX As Integer = (Me.ClientSize.Width - totalWidth) \ 2

        ' 设置每个按钮的位置
        cmdOK.Location = New Point(startX, Me.ClientSize.Height - cmdOK.Height - bottomMargin)
        cmdHelp.Location = New Point(cmdOK.Right + spacing, cmdOK.Top)
        cmdAbort.Location = New Point(cmdHelp.Right + spacing, cmdOK.Top)
    End Sub

End Class
