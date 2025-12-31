Imports System.Collections.ObjectModel
Imports System.DirectoryServices.ActiveDirectory
Imports System.Windows.Documents
Imports Microsoft.Win32
Imports NextcapServer = ncapsrv

Public Class frmLotManager
    'Inherits Form
    Inherits CustomForm

    Private WithEvents chkSortSuspendedList, chkSortClosedList As CheckBox
    Private WithEvents txtOpenLot As TextBox
    Private WithEvents cmdCreateNew, cmdDelete, cmdAbort, cmdOK As Button
    Private WithEvents C2O, S2O, C2S, O2C, S2C, O2S As Button
    Private WithEvents lstSuspended, lstClosed As ListBox
    Private lblOpen, lblSuspended, lblClosed As Label
    Private WithEvents tmrCheck As Timer
    'Private menuStrip As MenuStrip
    'Private lotStatusParentMenu As ToolStripMenuItem
    'Private WithEvents lotStatusEditorMenuItem As ToolStripMenuItem

    Private lotStatusContextMenu As ContextMenuStrip
    Private WithEvents lotStatusEditorMenuItem As ToolStripMenuItem

    Private WithEvents m_oLotManager As New NextcapServer.clsLotManager
    Private m_frmActiveSink As New frmLotManagerSink
    Private m_clsWindow As clsWindowState
    Private m_frmParent As Form
    Private m_colClosed As List(Of String)
    Private m_strOpenLot As String
    Private m_dtOpen As Date
    Private m_colSuspended As List(Of String)

    Private m_State As Integer
    Private Const cnSTATE_READY As Integer = 0
    Private Const cnSTATE_OK As Integer = 1
    Private Const cnSTATE_ABORT As Integer = 2
    Private Const cnSTATE_OPEN As Integer = 3
    Private Const cnSTATE_CLOSED As Integer = 4
    Private Const cnSTATE_SUSPENDED As Integer = 5
    Private Const cnSTATE_DELETEEMPTY As Integer = 6
    Private Const cnSTATE_CREATENEW As Integer = 7

    Private m_nStatusSelection As Integer
    Private Const cnOpen As Integer = 0
    Private Const cnClosed As Integer = 1
    Private Const cnSuspended As Integer = 2

    Private Const clngMinHeight As Long = 6135
    Private Const clngMinWidth As Long = 7290

    Private m_bSuspended As Boolean
    Private m_bClosed As Boolean

    Private Const cServerError As String = "ERROR"

    Private WithEvents m_oSecurity As Security
    Private m_nSecurity As Integer

    Private mb_AutoCreateLot As Boolean

    Private Const strEmptyString As String = ""

    Const cnNull As Integer = 0

    Private m_oOpenQueue As New colLotMangerQueue()

    Private m_oClosedQueue As colLotMangerQueue
    Private m_oSuspendedQueue As colLotMangerQueue
    Private m_strFrmId As String
    Public ReadOnly Property strFrmId() As String
        Get
            Return m_strFrmId
        End Get
    End Property

    Public Sub New()
        Me.Text = "Lot Manager"
        Me.ClientSize = New Size(711, 567)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.Name = "frmLotManager"

        ' === Timer ===
        tmrCheck = New Timer With {
            .Interval = 7000,
            .Enabled = False
        }

        ' === Open Lot TextBox ===
        txtOpenLot = New TextBox With {
            .Enabled = True,
            .Font = New Font("MS Serif", 13.5, FontStyle.Bold),
            .TextAlign = HorizontalAlignment.Center,
            .Location = New Point(36, 72),
            .Size = New Size(637, 42)
        }

        ' === Buttons ===
        cmdOK = New Button With {.Text = "OK", .Location = New Point(36, 480), .Size = New Size(133, 49)}
        cmdDelete = New Button With {.Text = "Delete Empty Lot", .Location = New Point(192, 480), .Size = New Size(157, 49), .Enabled = False}
        cmdCreateNew = New Button With {.Text = "Create New Lot", .Location = New Point(372, 480), .Size = New Size(145, 49)}
        cmdAbort = New Button With {.Text = "Abort", .Location = New Point(540, 480), .Size = New Size(133, 49)}

        ' Placeholder graphical buttons (add images if needed)
        Dim basePath As String = System.IO.Path.Combine(Application.StartupPath, "Resources")

        C2O = New Button With {
            .Location = New Point(144, 144),
            .Size = New Size(61, 61),
            .FlatStyle = FlatStyle.Flat,
            .BackgroundImage = Image.FromFile(System.IO.Path.Combine(basePath, "ArrowUp.png")),'.Image = Image.FromFile(System.IO.Path.Combine(basePath, "ArrowDown.png")),
            .BackgroundImageLayout = ImageLayout.Stretch '.ImageAlign = ContentAlignment.MiddleCenter
        }

        S2O = New Button With {
            .Location = New Point(576, 144),
            .Size = New Size(61, 61),
            .FlatStyle = FlatStyle.Flat,
            .BackgroundImage = Image.FromFile(System.IO.Path.Combine(basePath, "ArrowUp.png")),
            .BackgroundImageLayout = ImageLayout.Stretch
        }

        C2S = New Button With {
            .Location = New Point(324, 252),
            .Size = New Size(61, 61),
            .FlatStyle = FlatStyle.Flat,
            .BackgroundImage = Image.FromFile(System.IO.Path.Combine(basePath, "ArrowRight.png")),
            .BackgroundImageLayout = ImageLayout.Stretch
        }

        O2C = New Button With {
            .Location = New Point(72, 144),
            .Size = New Size(61, 61),
            .FlatStyle = FlatStyle.Flat,
            .BackgroundImage = Image.FromFile(System.IO.Path.Combine(basePath, "ArrowDown.png")),
            .BackgroundImageLayout = ImageLayout.Stretch
        }

        S2C = New Button With {
            .Location = New Point(324, 324),
            .Size = New Size(61, 61),
            .FlatStyle = FlatStyle.Flat,
            .BackgroundImage = Image.FromFile(System.IO.Path.Combine(basePath, "ArrowLeft.png")),
            .BackgroundImageLayout = ImageLayout.Stretch
        }

        O2S = New Button With {
            .Location = New Point(504, 144),
            .Size = New Size(61, 61),
            .FlatStyle = FlatStyle.Flat,
            .BackgroundImage = Image.FromFile(System.IO.Path.Combine(basePath, "ArrowDown.png")),
            .BackgroundImageLayout = ImageLayout.Stretch
        }


        ' === ListBoxes ===
        lstClosed = New ListBox With {
            .Font = New Font("MS Serif", 9.75),
            .Location = New Point(36, 228),
            .Size = New Size(253, 126)
        }

        lstSuspended = New ListBox With {
            .Font = New Font("MS Serif", 9.75),
            .Location = New Point(420, 228),
            .Size = New Size(253, 126)
        }

        ' === CheckBoxes ===
        chkSortClosedList = New CheckBox With {.Name = "chkSortClosedList", .Text = "Sort by Name", .Location = New Point(36, 360), .Size = New Size(253, 25)}
        chkSortSuspendedList = New CheckBox With {.Name = "chkSortSuspendedList", .Text = "Sort by Name", .Location = New Point(420, 360), .Size = New Size(253, 25)}

        ' === Labels ===
        lblOpen = New Label With {
            .Text = "Open",
            .Font = New Font("MS Sans Serif", 8.25, FontStyle.Bold),
            .TextAlign = ContentAlignment.MiddleCenter,
            .Location = New Point(264, 24),
            .Size = New Size(181, 25)
        }

        lblClosed = New Label With {
            .Text = "Closed",
            .Font = New Font("MS Sans Serif", 8.25, FontStyle.Bold),
            .TextAlign = ContentAlignment.MiddleCenter,
            .Location = New Point(36, 420),
            .Size = New Size(241, 25)
        }

        lblSuspended = New Label With {
            .Text = "Suspended",
            .Font = New Font("MS Sans Serif", 8.25, FontStyle.Bold),
            .TextAlign = ContentAlignment.MiddleCenter,
            .Location = New Point(420, 420),
            .Size = New Size(253, 25)
        }

        '' === MenuStrip ===
        '' === MenuStrip Initialization ===
        'menuStrip = New MenuStrip With {.Visible = False}

        '' === Lot Status Editor parent menu ===
        'lotStatusParentMenu = New ToolStripMenuItem("LotStatusEditor") With {.Visible = False}

        '' === Lot Status Editor submenu ===
        'lotStatusEditorMenuItem = New ToolStripMenuItem("Lot Status Editor")
        'lotStatusEditorMenuItem.DropDownItems.Add("View Lot Status", Nothing, AddressOf ViewLotStatus_Click)
        'lotStatusEditorMenuItem.DropDownItems.Add("Edit Lot Status", Nothing, AddressOf EditLotStatus_Click)
        'lotStatusEditorMenuItem.DropDownItems.Add("Refresh Status List", Nothing, AddressOf RefreshStatusList_Click)

        'lotStatusParentMenu.DropDownItems.Add(lotStatusEditorMenuItem)
        'menuStrip.Items.Add(lotStatusParentMenu)

        'Me.Controls.Add(menuStrip)
        'Me.MainMenuStrip = menuStrip

        ' 创建上下文菜单
        lotStatusContextMenu = New ContextMenuStrip()

        ' 创建菜单项
        lotStatusEditorMenuItem = New ToolStripMenuItem("Lot Status Editor")

        ' 添加子菜单项
        lotStatusEditorMenuItem.DropDownItems.Add("View Lot Status", Nothing, AddressOf ViewLotStatus_Click)
        lotStatusEditorMenuItem.DropDownItems.Add("Edit Lot Status", Nothing, AddressOf EditLotStatus_Click)
        lotStatusEditorMenuItem.DropDownItems.Add("Refresh Status List", Nothing, AddressOf RefreshStatusList_Click)

        ' 将菜单项添加到上下文菜单
        lotStatusContextMenu.Items.Add(lotStatusEditorMenuItem)

        ' 将上下文菜单绑定到需要右键功能的控件
        txtOpenLot.ContextMenuStrip = lotStatusContextMenu
        lstClosed.ContextMenuStrip = lotStatusContextMenu
        lstSuspended.ContextMenuStrip = lotStatusContextMenu


        ' === Add to Form ===
        'Me.Controls.AddRange({
        '    txtOpenLot, cmdOK, cmdDelete, cmdCreateNew, cmdAbort,
        '    C2O, S2O, C2S, O2C, S2C, O2S,
        '    lstClosed, lstSuspended,
        '    chkSortClosedList, chkSortSuspendedList,
        '    lblOpen, lblClosed, lblSuspended,
        '    MenuStrip
        '})
        'Me.MainMenuStrip = MenuStrip

        Me.Controls.AddRange({
            txtOpenLot, cmdOK, cmdDelete, cmdCreateNew, cmdAbort,
            C2O, S2O, C2S, O2C, S2C, O2S,
            lstClosed, lstSuspended,
            chkSortClosedList, chkSortSuspendedList,
            lblOpen, lblClosed, lblSuspended
        })
    End Sub

    Public Sub ShowWindow()
        ' 初始化此窗口
        WindowInit()

        ' 设置对新窗口状态的引用
        If Not String.IsNullOrEmpty(m_strFrmId) Then
            mdlWindow.RemoveForm(m_strFrmId)
        End If

        m_strFrmId = mdlWindow.AddForm(Me)
    End Sub

    Public Sub HideWindow()
        ' 隐藏此窗口
        'mdlWindow.RemoveForm(m_strFrmId)
        Me.Hide()
        m_strFrmId = String.Empty
    End Sub

    Private Sub WindowInit()
        ' 锁定图像检查扫描
        gn_ImageCheckCount += 1

        ' 强制当前状态为就绪
        m_State = cnSTATE_READY

        ' 设置当前状态
        ProcessState(cnSTATE_OPEN)

        ' 如果当前没有批次，显示默认批次
        If String.IsNullOrEmpty(m_strOpenLot) OrElse String.Equals(m_strOpenLot, "There is currently no open lot", StringComparison.OrdinalIgnoreCase) Then
            ' 同步屏幕与成员变量
            ' 调用获取下一个组ID的过程
            txtOpenLot.Text = mdlLotManager.GetNextGroupId(m_frmActiveSink)
        End If
    End Sub

    Public Property AutoCreateLot() As Boolean
        Get
            Return mb_AutoCreateLot
        End Get
        Set(value As Boolean)
            mb_AutoCreateLot = value
        End Set
    End Property

    ' 返回已关闭批次日期集合的引用
    Public ReadOnly Property ClosedLotDate() As List(Of String)
        Get
            Return m_colClosed
        End Get
    End Property

    ' 返回已暂停批次日期集合的引用
    Public ReadOnly Property SuspendedLotDate() As List(Of String)
        Get
            Return m_colSuspended
        End Get
    End Property

    ' 是否使用暂停列表的属性
    Public Property bSuspendedUsed() As Boolean
        Get
            Return m_bSuspended
        End Get
        Set(value As Boolean)
            ' 如果请求显示且当前状态为隐藏，则重新配置屏幕
            If value Then
                lstSuspended.Visible = True
                chkSortSuspendedList.Visible = True
                C2S.Visible = True
                S2C.Visible = True
                O2S.Visible = True
                S2O.Visible = True
                lblSuspended.Visible = True
            Else
                lstSuspended.Visible = False
                chkSortSuspendedList.Visible = False
                C2S.Visible = False
                S2C.Visible = False
                O2S.Visible = False
                S2O.Visible = False
                lblSuspended.Visible = False
            End If

            ' 设置私有变量值
            m_bSuspended = value

            ' 触发窗体大小调整事件
            OnResize(EventArgs.Empty)
        End Set
    End Property

    ' 是否使用已关闭列表的属性
    ' 如果暂停列表已隐藏，则忽略此设置以确保至少有一个可见列表
    Public Property bClosedUsed() As Boolean
        Get
            Return m_bClosed
        End Get
        Set(value As Boolean)
            ' 暂停列表可见
            If m_bSuspended = True Then
                ' 如果请求显示且当前状态为隐藏，则重新配置屏幕
                If value Then
                    lstClosed.Visible = True
                    chkSortClosedList.Visible = True
                    C2S.Visible = True
                    S2C.Visible = True
                    O2C.Visible = True
                    C2O.Visible = True
                    lblClosed.Visible = True
                Else
                    lstClosed.Visible = False
                    chkSortClosedList.Visible = False
                    C2S.Visible = False
                    S2C.Visible = False
                    O2C.Visible = False
                    C2O.Visible = False
                    lblClosed.Visible = False
                End If

                ' 设置私有变量值
                m_bClosed = value

                ' 触发窗体大小调整事件
                OnResize(EventArgs.Empty)
            Else
                ' 如果暂停列表不可见，则确保此列表可见
                m_bClosed = True
            End If
        End Set
    End Property

    ' 取消按钮点击事件
    Private Sub cmdAbort_Click(sender As Object, e As EventArgs) Handles cmdAbort.Click
        cmdOk_Click(sender, e)
    End Sub
    'Modifications:
    '   02-10-2000 Included decrement on ImageCheck to release
    '   its lockout and then Execute the image check when we are
    '   done with this window.
    '
    Private Sub cmdOk_Click(sender As Object, e As EventArgs) Handles cmdOK.Click
        Dim strMsg As String

        ' 关闭此窗口
        HideWindow()

        ' 减少图像检查锁计数
        gn_ImageCheckCount -= 1

        ' 执行图像扫描
        'mdlChart.ExecuteImageCheck()

        ' 如果没有打开的批次，显示警告
        If String.IsNullOrEmpty(m_strOpenLot) OrElse String.Equals(m_strOpenLot, "There is currently no open lot") Then
            ' 向用户发送警报消息
            strMsg = "ERROR NO OPEN LOT" & Environment.NewLine &
                 "The current Context has no open Lot." & Environment.NewLine &
                 "Go to the Lot Manager screen " &
                 "to create a New Lot." & Environment.NewLine &
                 "frmLotManager.OK() Error#1"

            MessageBox.Show(strMsg)
        Else
            ' 释放客户端锁
            gb_InputBusyFlag = False
        End If

        ' 触发批次指针例程
        'mdlChart.SetLotPointerIcon()
        Me.DialogResult = DialogResult.OK
    End Sub

    Private Sub C2O_Click(sender As Object, e As EventArgs) Handles C2O.Click
        Dim nIndex As Integer

        ' 检查列表项是否有选中值
        If Not String.IsNullOrEmpty(lstClosed.Text) Then
            nIndex = lstClosed.SelectedIndex

            ' CSB 否决状态变更检查
            If Not mdlSAX.ExecuteCSB_VetoLotStatusChange(
            lstClosed.Text,
            m_oClosedQueue(nIndex).dtBirth,
            mdlSAX.m_csClose,
            mdlSAX.m_csOpen) Then

                ' 打开选中的批次
                m_oLotManager.OpenLot(lstClosed.Text, m_oClosedQueue(nIndex).dtBirth)
            End If
        End If
    End Sub

    Private Sub C2S_Click(sender As Object, e As EventArgs) Handles C2S.Click
        Dim nIndex As Integer

        ' 检查列表项是否有选中值
        If Not String.IsNullOrEmpty(lstClosed.Text) Then
            nIndex = lstClosed.SelectedIndex

            ' CSB 否决状态变更检查
            If Not mdlSAX.ExecuteCSB_VetoLotStatusChange(
            lstClosed.Text,
            m_oClosedQueue(nIndex).dtBirth,
            mdlSAX.m_csClose,
            mdlSAX.m_csSuspend) Then

                ' 暂停选中的批次
                m_oLotManager.SuspendLot(lstClosed.Text, m_oClosedQueue(nIndex).dtBirth)
            End If
        End If
    End Sub

    '=======================================================
    'Routine: frmLotManager.ResumeCreate()
    'Purpose: This is for the frmWizard to call back to
    'restart the lot creation process.
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
    '=======================================================
    Public Sub ResumeCreate(ByVal strLotId As String)
        txtOpenLot.Text = strLotId
        cmdCreateNew_Click()
    End Sub
    '--------------------------------------------------------
    ' -- Trigger the change in the sort order of this list
    '
    ' Tested:
    '
    ' Modifications:
    '   04-24-2002 v2.0.8 Chris Barker: as written
    '
    '--------------------------------------------------------
    Private Sub chkSortClosedList_Click() Handles chkSortClosedList.Click
        ' 写入注册表
        Dim regKey As RegistryKey
        regKey = Registry.CurrentUser.CreateSubKey($"SOFTWARE\{mdlMain.cstrRegName}\{Me.Name}")
        regKey.SetValue(chkSortClosedList.Name, chkSortClosedList.Checked)
        regKey.Close()

        ' 按需排序列表（修改为支持倒序）
        If chkSortClosedList.Checked Then
            'm_oClosedQueue = mergeSort(m_oClosedQueue)
            'm_oClosedQueue = reverseQueue(m_oClosedQueue)
            'writeQueueToList(m_oClosedQueue, lstClosed)

            ' reverse m_oClosedQueue
            m_oClosedQueue.SortDescending()

            ' 1. disable default sort
            lstClosed.Sorted = False

            ' 2. save lstclosed to items
            Dim items(lstClosed.Items.Count - 1) As Object
            lstClosed.Items.CopyTo(items, 0)

            ' 3. sort items
            Array.Sort(items)

            ' 4. reverse sort
            Array.Reverse(items)

            ' 5. clear lstClosed and add sorted items
            lstClosed.Items.Clear()
            lstClosed.Items.AddRange(items)
        Else
            lstClosed.Sorted = False
        End If
    End Sub
    '=======================================================
    'Routine: frmLotManager.cmdCreateNew_Click()
    'Purpose: This is a standard button event. It
    'passes the call to the server to generate a new
    'Group Id. It also disables the NewLot wait timer
    'on the attached LotMangerSink so that we don't
    'get a double attempt at creating a new lot.
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
    '   07-08-1999 This was missing the call to PreCreateLot
    '   in the CSB collection.
    '
    '   01-22-2001 Added in call to CSB for Lot Id verfication.
    '   Chris Barker
    '
    '   07-06-2001 Chris Barker
    '   ExecuteLotWizard called if csb is active.
    '=======================================================
    Private Sub cmdCreateNew_Click() Handles cmdCreateNew.Click
        Dim strCurrentOpenLot As String
        Dim strLot As String

        ' 执行批次向导（如果启用）
        If mdlSAX.LotWizardUsed Then
            ' 如果向导尚未运行，则显示向导
            ' 否则继续执行
            'If frmWizard.nWizardPhase = 0 Then
            '    frmWizard.SetCaller(Me.Name, txtOpenLot.Text, Me)
            '    frmWizard.ShowDialog()
            '    Exit Sub
            'Else
            '    frmWizard.nWizardPhase = 0
            'End If
        End If

        ' 让 CSB 有机会修改批次 ID
        strLot = mdlSAX.ExecuteCSB_PreCreateLot(txtOpenLot.Text)

        ' 验证批次 ID 是否符合配置规则
        If ExecuteCSB_VerifyLotId(strLot) Then 'continue here,now use engineer mode to verify lotid, check why m_strOpenLot is nothing. remember delete this comment 
            If Not String.IsNullOrEmpty(strLot) Then
                ' 处理期间暂时禁用输入
                txtOpenLot.Enabled = False
                cmdCreateNew.Enabled = False

                ' 调用批次管理器的 CreateLot 函数
                If Not m_oLotManager.CreateLot(strLot) Then
                    ' 尝试获取批次管理器的打开批次
                    RecvOpenLot()

                    ' 检查是否获得有效结果
                    If String.IsNullOrEmpty(m_strOpenLot) OrElse String.Equals(m_strOpenLot, "There is currently no open lot") Then
                        ' 清除当前批次的私有变量，表示没有打开的批次
                        m_strOpenLot = String.Empty

                        ' 如果批次创建失败，启用输入框和命令按钮
                        txtOpenLot.Enabled = True
                        cmdCreateNew.Enabled = True

                        ' 显示错误消息
                        MessageBox.Show($"The requested Lot Id was rejected by the server ({strLot})." & Environment.NewLine & "CreateLot() #1")
                    End If
                End If
            End If
            For Each item As clsChartData In gcol_ChartData
                Debug.WriteLine($"ChartData: {item.strGroupId} - {item.dtBirth}")
            Next
        Else
            ' 显示错误消息
            MessageBox.Show($"CreateNewLot: The requested Lot Id failed VerifyLotId ({strLot})." & Environment.NewLine & "Check with IT Support.", "frmLotManager", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub


    Private Sub cmdDelete_Click(sender As Object, e As EventArgs) Handles cmdDelete.Click
        ' 禁用删除按钮防止重复点击
        cmdDelete.Enabled = False

        Try
            ' 检查并确保处于可删除的有效状态
            If m_State = cnSTATE_CLOSED Then
                ' 请求删除已关闭列表中选中的批次
                Dim selectedIndex As Integer = lstClosed.SelectedIndex
                If selectedIndex >= 0 Then
                    m_oLotManager.DeleteLot(
                    lstClosed.Text,
                    m_oClosedQueue(selectedIndex).dtBirth
                )
                End If
            ElseIf m_State = cnSTATE_SUSPENDED Then
                ' 请求删除已暂停列表中选中的批次
                Dim selectedIndex As Integer = lstSuspended.SelectedIndex
                If selectedIndex >= 0 Then
                    m_oLotManager.DeleteLot(
                    lstSuspended.Text,
                    m_oSuspendedQueue(selectedIndex).dtBirth
                )
                End If
            End If

            ' 处理当前状态（可能触发界面刷新）
            ProcessState(m_State)

        Catch ex As Exception
            ' 错误处理：记录日志并恢复按钮状态
            Debug.WriteLine($"删除批次时出错: {ex.Message}")
            MessageBox.Show("删除批次失败，请重试。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            ' 恢复按钮可用性
            cmdDelete.Enabled = True
        End Try
    End Sub

    Private Sub Form_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        WindowInit()

        ' 设置对新窗口状态的引用
        If Not String.IsNullOrEmpty(m_strFrmId) Then
            'mdlWindow.RemoveForm(m_strFrmId)
        End If

        'm_strFrmId = mdlWindow.AddForm(Me)

        ' 禁用关闭按钮(X)
        mdlTools.DisableCloseX(Me)

        ' 强制当前状态为就绪
        m_State = cnSTATE_READY
        'm_bSuspended = True

        ' 连接到安全对象并获取当前权限
        m_oSecurity = go_Security
        m_nSecurity = m_oSecurity.nAuthority
        'm_nSecurity = 6 ' hardcode

        ' 读取复选框控件的排序设置
        chkSortClosedList.Checked = CBool(GetSettingValue(mdlMain.cstrRegName, Me.Name, chkSortClosedList.Name, False))
        chkSortSuspendedList.Checked = CBool(GetSettingValue(mdlMain.cstrRegName, Me.Name, chkSortSuspendedList.Name, False))
    End Sub
    '============================================================
    'Routine: frmLotManager.Form_MouseDown(n,s,s)
    'Purpose: This is used to trap the right mouse button
    'button down event for the txtOpenLot object since
    'it will not receive any mouse events when it is disabled.
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
    '   02-15-1998 As written for Pass1.7
    '
    '
    '=============================================================
    Private Sub Form_MouseDown(sender As Object, e As MouseEventArgs) Handles Me.MouseDown
        ' 捕获右键点击（VB.NET 中 MouseButtons.Right 对应值为 2）
        If e.Button = MouseButtons.Right Then
            ' 检查 txtOpenLot 是否禁用（Enabled = False）
            If Not txtOpenLot.Enabled Then
                ' 获取 txtOpenLot 的坐标区域（注意 VB.NET 使用 ClientRectangle）
                Dim controlRect As Rectangle = txtOpenLot.ClientRectangle
                Dim clientPoint As Point = New Point(e.Location.X, e.Location.Y)

                ' 检查鼠标位置是否在控件区域内（VB.NET 坐标为控件相对坐标）
                If controlRect.Contains(clientPoint) Then
                    ' 设置状态选择
                    m_nStatusSelection = cnOpen

                    ' 显示上下文菜单（假设 Menu_LotStatusEditor 已转换为 ContextMenuStrip）
                    If txtOpenLot.ContextMenuStrip IsNot Nothing Then
                        txtOpenLot.ContextMenuStrip.Show(txtOpenLot, e.Location)
                    End If
                End If
            End If
        End If
    End Sub
    '=======================================================
    'Routine: frmLotManager.Form_Resize()
    'Purpose: This is the standard form resize event. We
    'need to make sure that we adjust this according to
    'what controls are visible.
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
    '   01-25-1998 As written for Pass1.6
    '
    '   03-22-2000 Added in an if..then to the
    '   closed/suspended lists to allow for interpreting
    '   the DisableClosed runtime value. This makes
    '   the assumption that suspended will be set before
    '   the closed is set.
    '
    '   04-24-2002 Chris Barker: Added in sizing for sort
    '   check box controls.
    '=======================================================
    'Private Sub Form_Resize(sender As Object, e As EventArgs) Handles Me.Resize
    '    ' 确保窗口状态为正常（非最小化或最大化）
    '    If Me.WindowState = FormWindowState.Normal Then
    '        ' 确保窗口不小于最小尺寸
    '        If Me.Width < clngMinWidth Then Me.Width = clngMinWidth
    '        If Me.Height < clngMinHeight Then Me.Height = clngMinHeight

    '        ' 移动"打开批次"文本框
    '        txtOpenLot.Width = Me.Width - 920

    '        ' 移动"确定"按钮
    '        cmdOK.Top = Me.Height - 1332

    '        ' 移动"取消"按钮
    '        cmdAbort.Location = New Point(Me.Width - 1890, cmdOK.Top)

    '        ' 计算中间位置
    '        Dim nMiddle As Integer = Me.Width \ 2

    '        ' 移动"新建"和"删除"按钮
    '        cmdCreateNew.Location = New Point(nMiddle + 75, cmdOK.Top)
    '        cmdDelete.Location = New Point(nMiddle - 1725, cmdOK.Top)

    '        ' 如果同时显示暂停和关闭列表
    '        If m_bSuspended AndAlso m_bClosed Then
    '            ' 重新定位标签
    '            lblClosed.Top = Me.Height - 1935
    '            lblSuspended.Location = New Point(Me.Width - 3090, lblClosed.Top)
    '            lblOpen.Left = nMiddle - (lblOpen.Width \ 2 + 100)

    '            ' 中间位置偏移-100
    '            ' 关闭-暂停按钮
    '            C2S.Left = nMiddle - 406
    '            S2C.Left = C2S.Left

    '            ' 移动列表框
    '            lstClosed.Bounds = New Rectangle(360, 2280, nMiddle - (406 + 360 + 348), Me.Height - 4515)
    '            lstSuspended.Bounds = New Rectangle(nMiddle + (306 + 348), 2280, lstClosed.Width - 100, lstClosed.Height)

    '            ' 移动排序复选框
    '            chkSortClosedList.Bounds = New Rectangle(lstClosed.Left, lstClosed.Top + lstClosed.Height, lstClosed.Width, chkSortClosedList.Height)
    '            chkSortSuspendedList.Bounds = New Rectangle(lstSuspended.Left, lstSuspended.Top + lstSuspended.Height, lstSuspended.Width, chkSortSuspendedList.Height)

    '            ' 打开-暂停按钮
    '            O2S.Left = Me.Width - 2250
    '            S2O.Left = Me.Width - 1530

    '            ' 打开-关闭按钮
    '            O2C.Left = 720
    '            C2O.Left = 1440

    '            ' 关闭列表隐藏
    '        ElseIf Not m_bClosed Then
    '            ' 重新定位标签
    '            lblSuspended.Location = New Point(nMiddle - lblSuspended.Width \ 2, Me.Height - 1935)

    '            ' 移动暂停列表框
    '            lstSuspended.Bounds = New Rectangle(360, 2280, Me.Width - 920, Me.Height - 4515)

    '            ' 移动排序复选框
    '            chkSortSuspendedList.Bounds = New Rectangle(lstSuspended.Left, lstSuspended.Top + lstSuspended.Height, lstSuspended.Width, chkSortSuspendedList.Height)

    '            ' 移动打开-关闭按钮
    '            O2S.Left = nMiddle - O2S.Width - 100
    '            S2O.Left = nMiddle + 100

    '            ' 暂停列表隐藏
    '        Else
    '            ' 重新定位标签
    '            lblClosed.Location = New Point(nMiddle - lblClosed.Width \ 2, Me.Height - 1935)

    '            ' 移动关闭列表框
    '            lstClosed.Bounds = New Rectangle(360, 2280, Me.Width - 920, Me.Height - 4515)

    '            ' 移动排序复选框
    '            chkSortClosedList.Bounds = New Rectangle(lstClosed.Left, lstClosed.Top + lstClosed.Height, lstClosed.Width, chkSortClosedList.Height)

    '            ' 移动打开-关闭按钮
    '            O2C.Left = nMiddle - O2C.Width - 100
    '            C2O.Left = nMiddle + 100
    '        End If
    '    End If
    'End Sub
    Private Sub Form_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        Try
            ' 释放对象引用
            'm_oLotManager = Nothing
        Catch ex As Exception
            ' 记录错误（根据需要）
            Debug.WriteLine($"窗体关闭时出错: {ex.Message}")
        End Try
    End Sub
    '=======================================================
    'Routine: frmLotManager.lstClosed_Click()
    'Purpose: This uses the Mouse Click event to see
    'if the selected Lot is empty and hence valid to
    'be deleted with the Delete Empty Lot button.
    '
    'Globals:None
    '
    'Input:None
    '
    'Return:None
    '
    'Tested: Hand tested 01-21-1999 C. Barker
    '
    'Modifications:
    '   01-21-1998 As written for Pass1.6
    '
    '
    '=======================================================
    Private Sub lstClosed_Click(sender As Object, e As EventArgs) Handles lstClosed.Click
        ' 利用此过程的错误处理
        Dim selectedIndex As Integer = lstClosed.SelectedIndex
        If selectedIndex >= 0 Then
            Dim selectedLotId As String = lstClosed.SelectedItem.ToString()
            Dim lotBirthDate As DateTime = m_oClosedQueue(selectedIndex).dtBirth

            If LotIsEmpty(selectedLotId, lotBirthDate) AndAlso m_nSecurity > 5 Then
                cmdDelete.Enabled = True
            Else
                cmdDelete.Enabled = False
            End If
        Else
            cmdDelete.Enabled = False
        End If
    End Sub
    '=======================================================
    'Routine: frmLotManager.lstClosed_MouseDown()
    'Purpose: We need to trap the right mouse down on this
    'to enable poping the form to modify the Lot Status
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
    Private Sub lstClosed_MouseDown(sender As Object, e As MouseEventArgs) Handles lstClosed.MouseDown
        ' 捕获右键点击
        If e.Button = MouseButtons.Right Then
            ' 设置状态选择
            m_nStatusSelection = cnClosed

            ' 显示上下文菜单
            If lstClosed.ContextMenuStrip IsNot Nothing Then
                lstClosed.ContextMenuStrip.Show(lstClosed, e.Location)
            End If
        End If
    End Sub
    '=======================================================
    'Routine: frmLotManager.lstSuspended_MouseDown()
    'Purpose: We need to trap the right mouse down on this
    'to enable poping the form to modify the Lot Status
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
    Private Sub lstSuspended_MouseDown(sender As Object, e As MouseEventArgs) Handles lstSuspended.MouseDown
        ' 捕获右键点击
        If e.Button = MouseButtons.Right Then
            ' 设置状态选择
            m_nStatusSelection = cnSuspended

            ' 显示上下文菜单
            If lstSuspended.ContextMenuStrip IsNot Nothing Then
                lstSuspended.ContextMenuStrip.Show(lstSuspended, e.Location)
            End If
        End If
    End Sub


    Private Sub lstClosed_MouseUp(sender As Object, e As MouseEventArgs) Handles lstClosed.MouseUp
        ' 处理鼠标释放事件
        If e.Button = MouseButtons.Left Then  ' 仅处理左键抬起（根据原代码意图调整）
            ProcessState(cnSTATE_CLOSED)
        End If
    End Sub
    '=======================================================
    'Routine: frmLotManager.lstSuspended_Click()
    'Purpose: This is checking to see if the selected
    'Lot is empty and if so it sets the Delete Lot button
    'to enabled.
    '
    'Globals:None
    '
    'Input:None
    '
    'Return:None
    '
    'Tested: Hand tested 01-21-1999 C. Barker
    '
    'Modifications:
    '   01-21-1998 As written for Pass1.6
    '   Bypassed for now
    '
    '=======================================================
    Private Sub lstSuspended_Click(sender As Object, e As EventArgs) Handles lstSuspended.Click
        ' 利用此过程的错误处理
        Dim selectedIndex As Integer = lstSuspended.SelectedIndex
        If selectedIndex >= 0 Then
            ' 原代码被注释，保持禁用状态
            cmdDelete.Enabled = False

            ' 若需要启用删除逻辑，取消注释以下代码：
            ' Dim selectedLotId As String = lstSuspended.SelectedItem.ToString()
            ' Dim lotData As Object = m_colSuspended(selectedIndex) ' 假设集合索引从0开始
            ' If LotIsEmpty(selectedLotId, lotData) Then
            '     cmdDelete.Enabled = True
            ' Else
            '     cmdDelete.Enabled = False
            ' End If
        Else
            cmdDelete.Enabled = False
        End If
    End Sub

    Private Sub lstSuspended_MouseUp(sender As Object, e As MouseEventArgs) Handles lstSuspended.MouseUp
        ' 处理鼠标释放事件
        If e.Button = MouseButtons.Left Then  ' 仅处理左键抬起（根据原代码意图调整）
            ProcessState(cnSTATE_SUSPENDED)
        End If
    End Sub
    '/*This is used when the AutoCreateLot flag is set
    '/
    '/ 09-02-1999 Added in error handling for CloseLot failure
    Private Sub m_oLotManager_TimeToCloseLot(ByVal strLotId As String, ByVal dateBirthday As Date) Handles m_oLotManager.TimeToCloseLot
        Try
            If mb_AutoCreateLot AndAlso m_nSecurity > 3 Then
                m_oLotManager.CloseLot(strLotId, dateBirthday)
            End If
        Catch ex As Exception
            ' 记录错误日志
            LogToFile("Error.txt", $"frmLotManagerSink.m_oLotManager_TimeToCloseLot:{ex.HResult} {ex.Message}")
        End Try
    End Sub
    '/*If the AutoCreateLot flag has been set then attempt to create the next Lot
    '/*
    '/*Modifications:
    '/* 09-09-1999 Modified loop to check if a Lot
    '/* exists to prevent infinite loop.
    '/*--------------------------------------------------------------------
    Public Sub m_oLotManager_TimeToCreateLot() Handles m_oLotManager.TimeToCreateLot
        Dim strLotId As String
        Dim nIncrement As Integer
        Dim bResult As Boolean

        Try
            If mb_AutoCreateLot AndAlso m_nSecurity > 3 Then
                If Not Me.Visible Then
                    strLotId = mdlLotManager.GetNextGroupId(m_frmActiveSink)
                    strLotId = mdlSAX.ExecuteCSB_PreCreateLot(strLotId)

                    If Not ValidOpenLot() Then
                        bResult = m_oLotManager.CreateLot(strLotId)

                        ' 修正 Left/Right 方法的使用
                        While strLotId.Length > 2 AndAlso Not bResult AndAlso Not ValidOpenLot()
                            ' 提取最后两位字符（VB.NET 中需指定长度参数）
                            Dim lastTwoChars As String = strLotId.Substring(strLotId.Length - 2)

                            ' 转换为整数（使用 TryParse 避免异常）
                            If Integer.TryParse(lastTwoChars, nIncrement) Then
                                nIncrement += 1
                            Else
                                nIncrement = 1 ' 处理转换失败的默认值
                            End If

                            ' 生成新的批次 ID（保留前缀并补零）
                            strLotId = strLotId.Substring(0, strLotId.Length - 2) & nIncrement.ToString("00")

                            Application.DoEvents()
                            bResult = m_oLotManager.CreateLot(strLotId)
                        End While
                    End If
                End If
            End If
        Catch ex As Exception
            LogToFile("Error.txt", $"frmLotManagerSink.m_oLotManager_TimeToCreateLot:{ex.HResult} {ex.Message}")
            MessageBox.Show("A Lot could not be created")
        End Try
    End Sub
    '/*Change in the Authority level of the system
    Private Sub m_oSecurity_SecurityChange(ByVal nNewAuthority As Integer) Handles m_oSecurity.SecurityChange
        ' 设置私有副本用于流程控制
        m_nSecurity = nNewAuthority

        ' 处理当前状态
        ProcessState(cnSTATE_OPEN)
    End Sub
    Private Sub lotStatusEditorMenuItem_Click(sender As Object, e As EventArgs) Handles lotStatusEditorMenuItem.Click
        ' 显示当前选择组的批次信息
        If m_nStatusSelection = cnClosed Then
            ' 确保列表已选择项
            If lstClosed.SelectedIndex >= 0 Then
                ' 设置批次状态编辑器
                frmLotStatusEditor.SetGroup(
                lstClosed.SelectedItem.ToString(),
                m_oClosedQueue(lstClosed.SelectedIndex).dtBirth,
                m_frmActiveSink
            )

                ' 显示窗口
                frmLotStatusEditor.Show()
            End If
        ElseIf m_nStatusSelection = cnSuspended Then
            ' 确保列表已选择项
            If lstSuspended.SelectedIndex >= 0 Then
                ' 设置批次状态编辑器
                frmLotStatusEditor.SetGroup(
                lstSuspended.SelectedItem.ToString(),
                m_oSuspendedQueue(lstSuspended.SelectedIndex).dtBirth,
                m_frmActiveSink
            )

                ' 显示窗口
                frmLotStatusEditor.Show()
            End If
        Else
            ' 设置批次状态编辑器
            frmLotStatusEditor.SetGroup(m_strOpenLot, m_dtOpen, m_frmActiveSink)

            ' 显示窗口
            frmLotStatusEditor.Show()
        End If
    End Sub
    '===============================================================
    '   01-19-1999 Changed to inculde no paramters
    '   on the LotManager Call
    '
    '   08-16-2001 Added in Veto transaction CSB hook CB
    '===============================================================
    Private Sub O2C_Click(sender As Object, e As EventArgs) Handles O2C.Click
        Dim sOpen As String
        Dim dtOpen As Date

        If Not String.IsNullOrEmpty(m_strOpenLot) AndAlso Not String.Equals(m_strOpenLot, "There is currently no open lot") Then
            sOpen = m_strOpenLot
            dtOpen = m_dtOpen

            If Not mdlSAX.ExecuteCSB_VetoLotStatusChange(sOpen, dtOpen, mdlSAX.m_csOpen, mdlSAX.m_csClose) Then
                Console.WriteLine(m_dtOpen.ToString("yyyy-MM-dd HH:mm:ss"))
                ' 执行批次管理器关闭功能
                m_oLotManager.CloseLot(m_strOpenLot, m_dtOpen)
                CheckOpenMove(cnClosed, sOpen, dtOpen)
            End If
        End If
    End Sub
    '===============================================================
    '   01-19-1999 Changed to inculde no paramters
    '   on the LotManager Call. Reversed.
    '===============================================================
    Private Sub O2S_Click(sender As Object, e As EventArgs) Handles O2S.Click
        Dim sOpen As String
        Dim dtOpen As Date

        If Not String.IsNullOrEmpty(m_strOpenLot) AndAlso Not String.Equals(m_strOpenLot, "There is currently no open lot") Then
            sOpen = m_strOpenLot
            dtOpen = m_dtOpen

            ' 检查是否允许状态变更
            If Not mdlSAX.ExecuteCSB_VetoLotStatusChange(sOpen, dtOpen, mdlSAX.m_csOpen, mdlSAX.m_csSuspend) Then
                ' 执行批次挂起操作
                m_oLotManager.SuspendLot(m_strOpenLot, m_dtOpen)
                CheckOpenMove(cnSuspended, sOpen, dtOpen)
            End If
        End If
    End Sub
    Private Sub S2C_Click(sender As Object, e As EventArgs) Handles S2C.Click
        Dim nIndex As Integer

        If lstSuspended.SelectedIndex >= 0 Then
            nIndex = lstSuspended.SelectedIndex

            ' 检查是否允许状态变更
            If Not mdlSAX.ExecuteCSB_VetoLotStatusChange(
            lstSuspended.SelectedItem.ToString(),
            m_oSuspendedQueue(nIndex).dtBirth,
            mdlSAX.m_csSuspend,
            mdlSAX.m_csClose) Then

                ' 执行批次关闭操作
                m_oLotManager.CloseLot(
                lstSuspended.SelectedItem.ToString(),
                m_oSuspendedQueue(nIndex).dtBirth)
            End If
        End If
    End Sub
    Private Sub S2O_Click(sender As Object, e As EventArgs) Handles S2O.Click
        Dim nIndex As Integer
        Dim strLotId As String

        If lstSuspended.SelectedIndex >= 0 Then
            nIndex = lstSuspended.SelectedIndex
            strLotId = lstSuspended.SelectedItem.ToString()

            ' 检查是否允许状态变更
            If Not mdlSAX.ExecuteCSB_VetoLotStatusChange(
            strLotId,
            m_oSuspendedQueue(nIndex).dtBirth,
            mdlSAX.m_csSuspend,
            mdlSAX.m_csOpen) Then

                ' 执行批次打开操作
                m_oLotManager.OpenLot(strLotId, m_oSuspendedQueue(nIndex).dtBirth)
            End If
        End If
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

    Private Sub chkSortSuspendedList_Click(sender As Object, e As EventArgs) Handles chkSortSuspendedList.Click
        ' 写入注册表
        Dim regKey As RegistryKey
        regKey = Registry.CurrentUser.CreateSubKey($"SOFTWARE\{mdlMain.cstrRegName}\{Me.Name}")
        regKey.SetValue(chkSortSuspendedList.Name, chkSortSuspendedList.Checked)
        regKey.Close()

        ' v2.0.8：如果需要，对列表进行排序
        If chkSortSuspendedList.Checked Then
            'm_oSuspendedQueue = mergeSort(m_oSuspendedQueue)
            'm_oSuspendedQueue = reverseQueue(m_oSuspendedQueue)
            'writeQueueToList(m_oSuspendedQueue, lstSuspended)

            ' reverse m_oSuspendedQueue
            m_oSuspendedQueue.SortDescending()

            ' 1. disable default sort
            lstSuspended.Sorted = False

            ' 2. save lstSuspended to items
            Dim items(lstSuspended.Items.Count - 1) As Object
            lstSuspended.Items.CopyTo(items, 0)

            ' 3. sort items
            Array.Sort(items)

            ' 4. reverse sort
            Array.Reverse(items)

            ' 5. clear lstSuspended and add sorted items
            lstSuspended.Items.Clear()
            lstSuspended.Items.AddRange(items)
        Else
            lstSuspended.Sorted = False
        End If
    End Sub

    Private Sub ViewLotStatus_Click(sender As Object, e As EventArgs)
        MessageBox.Show("Viewing lot status...", "Lot Status Editor")
    End Sub

    Private Sub EditLotStatus_Click(sender As Object, e As EventArgs)
        MessageBox.Show("Editing lot status...", "Lot Status Editor")
    End Sub

    Private Sub RefreshStatusList_Click(sender As Object, e As EventArgs)
        MessageBox.Show("Status list refreshed.", "Lot Status Editor")
    End Sub
    '=======================================================
    'Routine: frmLotManager.CreateLotManager()
    'Purpose: This creates the LotManager from the business
    'server if possible and then runs the required setup
    'routines such as RecvLots.
    '
    'Globals:None
    '
    'Input:None
    '
    'Return: Boolean - Indicates if a succesful connection
    '        was generated.
    '
    'Tested:
    '
    'Modifications:
    '   01-06-1998 As written for Pass1.6
    '
    '   10-06-1999 Added constraint on empty lot error
    '   message.
    '
    '=======================================================
    Public Sub SetLotManager(ByRef oLotManager As NextcapServer.clsLotManager, ByRef frmActiveLotManager As frmLotManagerSink)
        Dim strMsg As String

        '/*Connect the COM sink for the B:server
        If Not (oLotManager Is Nothing) Then
            m_oLotManager = oLotManager
            m_frmActiveSink = frmActiveLotManager
            If m_oLotManager.LotIdCount > 0 Then 'FFFF
                '/*Query the Business server for the Lot Manager information
                RecvOpenLot()
                RecvClosedLot()
                RecvSuspendedLot()
            End If
            '/*Set the current state; We need to set closed
            '/*prior to Open to trigger the disabling of Open
            '/*if it shakes out that way.
            '/*If we don't have a lot set the state to Open
            If String.IsNullOrEmpty(m_strOpenLot) AndAlso go_clsSystemSettings.strMaterialMode = mdlGlobal.gcstrLot Then 'tttt
                'If (String.IsNullOrEmpty(m_strOpenLot) OrElse String.Equals(m_strOpenLot, "There is currently no open lot", StringComparison.OrdinalIgnoreCase)) AndAlso go_clsSystemSettings.strMaterialMode = mdlGlobal.gcstrLot Then
                '/*Generate error
                mdlLotManager.TransactionFailure("frmLotManager.SetLotManager()")
                '/*Set the state
                ProcessState(cnSTATE_OPEN)
            ElseIf String.Equals(m_strOpenLot, "There is currently no open lot", StringComparison.OrdinalIgnoreCase) AndAlso go_clsSystemSettings.strMaterialMode = mdlGlobal.gcstrLot Then
                ProcessState(cnSTATE_OPEN)
            Else
                'lstClosed_MouseUp(0, 0, 0, 0)
            End If
        Else
            '/*If the connnection failed disable
            '/*all of the screen funcitons except for abort
            cmdDelete.Enabled = False
            C2S.Enabled = False
            S2C.Enabled = False
            O2C.Enabled = False
            O2S.Enabled = False
            '/*Enable the text box
            txtOpenLot.Enabled = False
            '/*Set the state of the create lot button
            cmdCreateNew.Enabled = False
            C2O.Enabled = False
            S2O.Enabled = False
            lstClosed.Enabled = False
            lstSuspended.Enabled = False
        End If
    End Sub

    Public ReadOnly Property ActiveLotManager() As frmLotManagerSink
        Get
            Return m_frmActiveSink
        End Get
    End Property

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
    '   01-19-1999 Added the sync of go_Context.LotId
    '   and dtBirth with this form so that it will
    '   be continuously matched up.
    '=======================================================
    Public Sub RecvOpenLot()
        '/*Get the information from the server
        m_frmActiveSink.RecvOpenLots(m_strOpenLot, m_dtOpen)

        '/*Syncronize the global Context object
        '/*Push the open lot information to the Context object
        '/*so that any exterior processes can access it.
        go_Context.GroupId = m_strOpenLot
        go_Context.GroupBirthDay = m_dtOpen

        '/*Set the new items into the list
        'If String.Equals(m_strOpenLot, "There is currently no open lot") Then 'tttt
        '    m_strOpenLot = mdlLotManager.GetNextGroupId(Nothing)
        '    currentNoOpen = True
        'End If
        txtOpenLot.Text = m_strOpenLot
    End Sub
    '=======================================================
    'Routine: frmLotManager.RecvClosedLot()
    'Purpose: This queries the Lot Manager to get a list
    '         of the current Closed lots.
    '
    'Globals:None
    '
    'Input: None
    '
    'Return:None
    '
    'Modifications:
    '   01-07-1999 As written for Pass1.6
    '
    '   04-24-2002 Chris Barker: Changing from list control
    '   and collection to colLotMangerQueue for tracking
    '   the visual position of Lots.
    '
    '   07-08-2002 Chris Barker: Reverse the queue order to
    '   descending sort.
    '=======================================================
    Private Sub RecvClosedLot()
        Dim nIndex As Integer
        Dim colLots As New List(Of String)
        Dim colDates As New List(Of Date)

        '/*Clear the current contents of the private collection
        m_colClosed = New List(Of String)

        '/*Request the Lot information from the Server Sink
        m_frmActiveSink.RecvClosedLots(colLots, colDates)

        ' -- v2.0.8 Switch collections into queue
        m_oClosedQueue = New colLotMangerQueue()
        'For nIndex = 1 To colLots.Count
        For nIndex = 0 To colLots.Count - 1
            m_oClosedQueue.Add(colLots(nIndex), cnNull, cnNull, colDates(nIndex))
        Next nIndex

        chkSortClosedList.Checked = CBool(GetSettingValue(mdlMain.cstrRegName, Me.Name, chkSortClosedList.Name, False))
        ' -- v2.0.8 Sort the list if required
        If chkSortClosedList.Checked Then
            'm_oClosedQueue = mergeSort(m_oClosedQueue)
            ' -- 07-08-2002 reverse sort to descending
            'm_oClosedQueue = reverseQueue(m_oClosedQueue)
            m_oClosedQueue.SortDescending()
        End If

        ' -- v2.0.8 write the lot names to the list
        writeQueueToList(m_oClosedQueue, lstClosed)

        '/*Destroy the temporary Lot holder
        colLots = Nothing
        colDates = Nothing
    End Sub

    '=======================================================
    'Routine: mergeSort(o)
    'Purpose: Executes merge sort on a lot collection queue
    '
    'Globals:None
    '
    'Input: o - a colLotManger queue object
    '
    'Return: a colLotManger queue object
    '
    'Tested:
    '   04-24-2002 Interactive for cases with set sizes 1,2,3,4
    '
    'Modifications:
    '   04-24-2002 Chris Barker: as written
    '
    '---------------- Algorithm -------------------------------------------
    '    if (data.length == 1) return data;       // only one element
    '    int middleIdx = data.length / 2;
    '    int data1[] = new int[middleIdx];
    '    for (int i = 0; i < middleIdx; i++) {
    '      data1[i] = data[i];        // copy data[0] ... data[middleIdx - 1]
    '    }
    '    int data2[] = new int[data.length - middleIdx];
    '    for (int i = 0; i < data2.length; i++) {
    '      data2[i] = data[middleIdx + i]; // copy data[moddleIdx] ...
    '    }
    '    data1 = mergeSort(data1);
    '    data2 = mergeSort(data2);
    '    return merge(data1, data2);
    '=======================================================
    Public Function mergeSort(ByVal o As colLotMangerQueue) As colLotMangerQueue
        Dim Idx As Integer
        Dim middleIdx As Integer
        Dim o1 As New colLotMangerQueue()
        Dim o2 As New colLotMangerQueue()

        Try
            If o.Count < 2 Then
                '-- return the input
                Return o
            Else
                '--
                middleIdx = o.Count \ 2 ' 使用整除运算符

                '-- Divide to first partition
                For Idx = 1 To middleIdx
                    o1.Add(o.Item(Idx).sLotId, cnNull, cnNull, o.Item(Idx).dtBirth)
                Next Idx

                '-- divide to second partition
                For Idx = middleIdx + 1 To o.Count
                    o2.Add(o.Item(Idx).sLotId, cnNull, cnNull, o.Item(Idx).dtBirth)
                Next Idx

                '-- Sort the new partitions
                o1 = mergeSort(o1)
                o2 = mergeSort(o2)

                '-- merge the resulting two sorted sets
                Return merge(o1, o2)
            End If
        Catch ex As Exception
            ' 处理异常（原代码中仅清除错误）
            Debug.WriteLine($"Error in mergeSort: {ex.Message}")
        End Try
    End Function
    '=======================================================
    'Routine: merge(o,o)
    'Purpose: Executes merge on two lot collection queues
    '
    'Globals:None
    '
    'Input: o1 - a colLotManger queue object
    '       o2 - a colLotManger queue object
    '
    'Return: a colLotManger queue object
    '
    'Tested:
    '   04-24-2002 Interactive for cases with set sizes 1,2,3,4
    '
    'Modifications:
    '   04-24-2002 Chris Barker: as written
    '
    '
    '------------------- Algorithm -------------------------------------
    '    int[] result = new int[data1.length + data2.length];  // for merged values
    '    int idx1 = 0;      // index to point element in data1
    '    int idx2 = 0;      // index to point element in data2
    '    int idx = 0;       // index to point slot in result[]
    '
    '    while (idx1 < data1.length || idx2 < data2.length) {
    '      if (idx1 < data1.length && idx2 < data2.length &&
    '          data1[idx1] <= data2[idx2] || idx2 == data2.length) {
    '        result[idx] = data1[idx1]; // copy value from data1
    '        idx1++;
    '      } else {
    '        result[idx] = data2[idx2]; // copy value from data2
    '        idx2++;
    '      }
    '      idx++;
    '    }
    '=======================================================
    Public Function merge(ByVal o1 As colLotMangerQueue, ByVal o2 As colLotMangerQueue) As colLotMangerQueue
        Dim o As New colLotMangerQueue()
        Dim Idx As Integer
        Dim Idx1 As Integer
        Dim Idx2 As Integer

        Try
            ' -- init the pointers
            Idx = 1
            Idx1 = 1
            Idx2 = 1

            ' -- perform the sorting
            ' -- We have to perform the criteria in this manner since
            '    there is no short circuit Boolean And operation in VB
            Do While (Idx1 <= o1.Count Or Idx2 <= o2.Count)
                If Idx1 <= o1.Count Then
                    If Idx2 <= o2.Count Then
                        If o1.Item(Idx1).sLotId <= o2.Item(Idx2).sLotId Then
                            ' -- place at idx in this set
                            o.Add(o1.Item(Idx1).sLotId, cnNull, cnNull, o1.Item(Idx1).dtBirth)
                            Idx1 = Idx1 + 1
                        Else
                            ' -- place at idx in this set
                            o.Add(o2.Item(Idx2).sLotId, cnNull, cnNull, o2.Item(Idx2).dtBirth)
                            Idx2 = Idx2 + 1
                        End If
                    Else
                        ' -- place at idx in this set
                        o.Add(o1.Item(Idx1).sLotId, cnNull, cnNull, o1.Item(Idx1).dtBirth)
                        Idx1 = Idx1 + 1
                    End If
                Else
                    ' -- place at idx in this set
                    o.Add(o2.Item(Idx2).sLotId, cnNull, cnNull, o2.Item(Idx2).dtBirth)
                    Idx2 = Idx2 + 1
                End If
                'Idx = Idx + 1
            Loop

            ' -- return the result set
            Return o
        Catch ex As Exception
            ' 处理异常（原代码中仅清除错误）
            Debug.WriteLine($"Error in merge: {ex.Message}")
        End Try
    End Function

    '=======================================================
    'Routine: reverseQueue(o)
    'Purpose: Reverses the current order of the queue
    '
    'Globals:None
    '
    'Input: o - a colLotManger queue object
    '
    'Return: a colLotManger queue object
    '
    'Tested:
    '
    'Modifications:
    '   07-08-2002 Chris Barker: as written
    '=======================================================
    Public Function reverseQueue(ByRef o As colLotMangerQueue) As colLotMangerQueue
        Dim l As Long
        Dim temp As New colLotMangerQueue()

        Try
            'add items to temp as read in reverse order from the current queue
            For l = o.Count To 1 Step -1
                temp.Add(o.Item(l).sLotId, o.Item(l).nFrom, o.Item(l).nTo, o.Item(l).dtBirth)
            Next l

            If temp.Count = o.Count Then
                Return temp
            Else
                Return o
            End If
        Catch ex As Exception
            MessageBox.Show("Error.txt", "frmLotManager.reverseQueue:" & ex.HResult.ToString() & " " & ex.Message)
            'LogToFile("Error.txt", "frmLotManager.reverseQueue:" & ex.HResult.ToString() & " " & ex.Message)
            Err.Clear()
        End Try
    End Function
    '=======================================================
    'Routine: writeQueueToList(o,o)
    'Purpose: Writes the specified Lot Queue to the list.
    '
    'Globals:None
    '
    'Input: o - a colLotManger queue object
    '       lst - a list control
    '
    'Return: nothing
    '
    'Tested:
    '
    'Modifications:
    '   04-24-2002 Chris Barker: as written
    '
    '   07-08-2002 Chris Barker: Make sure the first
    '   index is focused on after addition.
    '=======================================================
    Public Sub writeQueueToList(ByRef o As colLotMangerQueue, ByRef lst As ListBox)
        Try
            '/*Clear the current list contents
            lst.Items.Clear()

            ' -- v2.0.8 write the lot names to the list
            For nIndex As Integer = 0 To o.Count - 1
                lst.Items.Add(o.Item(nIndex).sLotId)
            Next nIndex

            ' -- 07-08-2002 Try to focus the index to the first position
            If lst.Enabled AndAlso lst.Visible Then
                lst.SelectedIndex = 0  ' 设置选中索引（ListIndex 已过时，推荐使用 SelectedIndex）
                lst.Focus()            ' 可选：将焦点移到列表框
            End If

        Catch ex As Exception
            ' 处理异常（原代码仅清除错误）
            Debug.WriteLine($"Error in writeQueueToList: {ex.Message}")
        End Try
    End Sub
    '=======================================================
    'Routine: frmLotManager.RecvSuspendedLot()
    'Purpose: This queries the Lot Manager to get a list
    '         of the current Suspended lots.
    '
    'Globals:None
    '
    'Input: None
    '
    'Return:None
    '
    'Modifications:
    '   01-07-1999 As written for Pass1.6
    '
    '   04-24-2002 Chris Barker: Changing from list control
    '   and collection to colLotMangerQueue for tracking
    '   the visual position of Lots.
    '
    '   07-08-2002 Chris Barker: Reverse the queue order to
    '   descending sort.
    '=======================================================
    Private Sub RecvSuspendedLot()
        Dim nIndex As Integer
        Dim colLots As New List(Of String)
        Dim colDates As New List(Of Date)

        '/*Generate a Lot to bring the Lot Id s across in
        '/*Clear the current contents of the collection
        m_colSuspended = New List(Of String)
        '/*Request the Lot information from the Server Sink
        m_frmActiveSink.RecvSuspendedLots(colLots, colDates)

        ' -- v2.0.8 Switch collections into queue
        m_oSuspendedQueue = New colLotMangerQueue()
        'For nIndex = 1 To colLots.Count
        For nIndex = 0 To colLots.Count - 1
            m_oSuspendedQueue.Add(colLots(nIndex), cnNull, cnNull, colDates(nIndex))
        Next nIndex

        chkSortSuspendedList.Checked = CBool(GetSettingValue(mdlMain.cstrRegName, Me.Name, chkSortSuspendedList.Name, False))
        ' -- v2.0.8 Sort the list if required
        If chkSortSuspendedList.Checked Then
            'm_oSuspendedQueue = mergeSort(m_oSuspendedQueue)
            'm_oSuspendedQueue = reverseQueue(m_oSuspendedQueue)
            m_oSuspendedQueue.SortDescending()
        End If

        ' -- v2.0.8 write the lot names to the list
        writeQueueToList(m_oSuspendedQueue, lstSuspended)

        '/*Destroy the temporary Lot holder
        colLots = Nothing
        colDates = Nothing
    End Sub
    '=======================================================
    'Routine: ProcessState()
    'Purpose: This is the state machine for the LotManager
    'screen. It controls the on/off state of the buttons
    'and what object/list has focus.
    '
    'Globals:None
    '
    'Input: nNewState - A requested state per the listed
    '       Constants in the declaration of this form.
    '
    'Return:None
    '
    'Modifications:
    '   12-03-1998 As written for Pass1.5
    '
    '=======================================================
    Private Function ProcessState(ByRef nNewState As Integer) As Boolean
        Dim nZorder As Integer

        '/*Process a request for Open as active
        If nNewState = cnSTATE_OPEN Then
            '------------------------------------------
            '/*Set the button states
            '------------------------------------------
            If m_nSecurity > 5 Then cmdDelete.Enabled = False
            '-----------------------------------------------
            'Closed to Suspended
            '-----------------------------------------------
            C2S.Enabled = False
            S2C.Enabled = False

            '-----------------------------------------------
            'Out of Open buttons
            '-----------------------------------------------
            '/*Test for whether we have a Lot currently open
            If String.IsNullOrEmpty(m_strOpenLot) OrElse String.Equals(m_strOpenLot, "There is currently no open lot", StringComparison.OrdinalIgnoreCase) Then 'tttt
                O2C.Enabled = False
                O2S.Enabled = False
                '/*Enable the text box
                txtOpenLot.Enabled = True
                '/*Set focus to this text box if we are visible
                If Me.Visible Then
                    If txtOpenLot.Enabled Then
                        txtOpenLot.Focus()
                    End If
                End If
                '/*Set the state of the create lot button
                If m_nSecurity > 5 Then cmdCreateNew.Enabled = True
            Else
                '/*Enable the text box
                txtOpenLot.Enabled = False
                '/*Set the state of the create lot button
                cmdCreateNew.Enabled = False
                O2C.Enabled = True
                O2S.Enabled = True
            End If
            '---------------------------------------------------
            'Into Open buttons
            '---------------------------------------------------
            'Test if we have a lot that can come up from Closed
            C2O.Enabled = False
            S2O.Enabled = False
        ElseIf nNewState = cnSTATE_CLOSED Then
            '/*Set the state of the create lot button
            If (String.IsNullOrEmpty(m_strOpenLot) OrElse String.Equals(m_strOpenLot, "There is currently no open lot")) AndAlso m_nSecurity > 5 Then
                cmdCreateNew.Enabled = True
            Else
                cmdCreateNew.Enabled = False
            End If


            '--------------------------------------------
            'Open and Closed
            '--------------------------------------------
            '/*Set the state of the create lot button
            cmdCreateNew.Enabled = False
            '/*See if there is a currently open lot
            'If Not String.IsNullOrEmpty(m_strOpenLot) Then 'tttt
            If Not String.IsNullOrEmpty(m_strOpenLot) AndAlso Not String.Equals(m_strOpenLot, "There is currently no open lot") Then
                '/*Set the txtOpenLot to the real open lot
                txtOpenLot.Text = m_strOpenLot
                '/*Disable all buttons associated with Open
                '/*Transfer between Open and Closed
                O2C.Enabled = True
                '/*Shutoff Closed to Open since we already have a Lot open
                C2O.Enabled = False
                '/*Enable the text box
                txtOpenLot.Enabled = False
            Else
                '/*Enable the text box
                txtOpenLot.Enabled = True
                '/*Can not transfer a closed lot on to an open lot
                O2C.Enabled = False
                If Not (m_colClosed Is Nothing) OrElse (lstClosed.Items.Count) > 0 Then C2O.Enabled = True Else C2O.Enabled = False
            End If
            '-----------------------------------------------
            '/*Disable transfer between Open and Suspended
            '-----------------------------------------------
            O2S.Enabled = False
            S2O.Enabled = False

            '------------------------------------------------
            '/*Set state between Closed and Suspended
            '------------------------------------------------
            '/*Test if we have a lot that can come up from Closed
            If lstClosed.Items.Count > 0 Then C2S.Enabled = True Else C2S.Enabled = False
            '/*Switch off Suspended to Closed
            S2C.Enabled = False
        ElseIf nNewState = cnSTATE_SUSPENDED Then
            '/*Set the state of the create lot button
            If (String.IsNullOrEmpty(m_strOpenLot) OrElse String.Equals(m_strOpenLot, "There is currently no open lot")) AndAlso m_nSecurity > 5 Then
                cmdCreateNew.Enabled = True
            Else
                cmdCreateNew.Enabled = False
            End If

            '---------------------------------------------
            'Set the state on Open to Closed
            '---------------------------------------------
            O2C.Enabled = False
            C2O.Enabled = False

            '---------------------------------------------
            'Set the state of Closed to suspended
            '---------------------------------------------
            'Switch off Closed to Suspened
            C2S.Enabled = False
            '/*Test if we have a lot that can come up from Suspended
            If lstSuspended.Items.Count > 0 Then S2C.Enabled = True Else S2C.Enabled = False

            '---------------------------------------------
            'Set the state on Open to Suspended
            '---------------------------------------------
            '/*Set the state of the create lot button
            cmdCreateNew.Enabled = False
            '/*See if there is a currently open lot
            If Not String.IsNullOrEmpty(m_strOpenLot) AndAlso Not String.Equals(m_strOpenLot, "There is currently no open lot") Then
                '/*Set the txtOpenLot to the real open lot
                txtOpenLot.Text = m_strOpenLot
                '/*Disable all buttons associated with Open
                '/*Transfer between Open and Closed
                O2S.Enabled = True
                '/*Shutoff Closed to Open since we already have a Lot open
                S2O.Enabled = False
                '/*Enable the text box
                txtOpenLot.Enabled = False
            Else
                '/*Enable the text box
                txtOpenLot.Enabled = True
                '/*There is no Lot to transfer in
                O2S.Enabled = False
                '/*Test if we have a lot that can come up from Suspended
                If lstSuspended.Items.Count > 0 Then S2O.Enabled = True Else S2O.Enabled = False
            End If
        End If
        '/*Set the new state
        m_State = nNewState
        ProcessState = True
    End Function
    Private Sub lstClosed_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
        Call ProcessState(cnSTATE_CLOSED)
    End Sub
    '=======================================================
    'Routine: frmLotManager.LotIsEmpty(str,dt)
    'Purpose: Use the empty lot collection to look up
    '         if the questioned lot is empty.
    '
    'Globals:None
    '
    'Input: strGroupId - The of the Group from the selected
    '       list item.
    '
    '       dtBirth - The associated Birthday of the Lot.
    '
    'Return: Boolean - True=the lot is empty.
    '
    'Modifications:
    '   12-03-1998 As written for Pass1.5
    '
    '
    '=======================================================
    Private Function LotIsEmpty(ByRef strGroupId As String, ByRef dtBirth As Date) As Boolean
        Try
            ' 测试批次是否可删除
            If m_oLotManager IsNot Nothing Then
                Return m_oLotManager.IsLotEmpty(strGroupId, dtBirth)
            Else
                Return False ' 或根据业务逻辑返回其他默认值
            End If
        Catch ex As Exception
            ' 记录错误日志
            LogToFile("Error.txt", $"frmLotManager.LotIsEmpty:{ex.HResult} {ex.Message}")
            Return False ' 发生异常时返回默认值
        End Try
    End Function
    '=======================================================
    'Routine: ValidOpenLot()
    'Purpose: This allows other functions to test if there
    'is a presently open valid lot before.
    '
    'Globals:None
    '
    'Input:None
    '
    'Return: True - There is an open lot.
    '
    'Tested:
    '
    'Modifications:
    '   07-22-1999 As written for Beta2 Phase3.2
    '
    '
    '=======================================================
    Public Function ValidOpenLot() As Boolean
        ' 检查打开的批次 ID 是否有值
        If String.IsNullOrEmpty(m_strOpenLot) Then
            Return False
        ElseIf String.Equals(m_strOpenLot, "There is currently no open lot") Then
            Return False
        Else
            Return True
        End If

        'Return Not String.IsNullOrEmpty(m_strOpenLot)
    End Function
    '=======================================================
    'Routine: CheckOpenMove()
    'Purpose: Insure that a requested move is not lost
    'in event calls.
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
    '   02-16-2000 February 2000 release
    '
    '
    '=======================================================
    Public Sub CheckOpenMove(ByRef nTo As Integer, ByRef sOpenLotId As String, ByRef dtBirth As Date)
        ' 添加到开放队列
        m_oOpenQueue.Add(sOpenLotId, cnOpen, nTo, dtBirth)

        ' 启用检查计时器
        tmrCheck.Enabled = True
    End Sub
    '=======================================================
    'Routine: tmrCheck_Timer()
    'Purpose: After a manual move is performed on the
    'open lot this procedure will trigger 7 seconds
    'later to make sure that the Lot appears to be in the
    'right place.
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
    Private Sub tmrCheck_Tick(sender As Object, e As EventArgs) Handles tmrCheck.Tick
        Try
            ' 处理开放队列中的所有项目
            While m_oOpenQueue.Count > 0
                Dim queueItem = m_oOpenQueue(0) ' VB.NET 集合索引从 0 开始

                ' 检查批次是否仍处于开放状态
                If queueItem.nFrom = cnOpen Then
                    If queueItem.sLotId = m_strOpenLot AndAlso queueItem.dtBirth = m_dtOpen Then
                        ' 移动失败，不执行任何操作
                    Else
                        ' 检查批次是否已移动到目标状态
                        If queueItem.nTo = cnClosed Then
                            ' 反向查找已关闭列表
                            Dim found As Boolean = False
                            For nIndex = lstClosed.Items.Count - 1 To 0 Step -1
                                If lstClosed.Items(nIndex).ToString() = queueItem.sLotId Then
                                    If m_oClosedQueue(nIndex).dtBirth = queueItem.dtBirth Then
                                        found = True
                                        Exit For
                                    End If
                                End If
                            Next

                            ' 未找到，重新加载已关闭批次
                            If Not found Then RecvClosedLot()
                        ElseIf queueItem.nFrom = cnSuspended Then
                            ' 反向查找已暂停列表
                            Dim found As Boolean = False
                            For nIndex = lstSuspended.Items.Count - 1 To 0 Step -1
                                If lstSuspended.Items(nIndex).ToString() = queueItem.sLotId Then
                                    If m_oSuspendedQueue(nIndex).dtBirth = queueItem.dtBirth Then
                                        found = True
                                        Exit For
                                    End If
                                End If
                            Next

                            ' 未找到，重新加载已暂停批次
                            If Not found Then RecvSuspendedLot()
                        End If
                    End If
                End If

                ' 移除处理过的项目
                m_oOpenQueue.Remove(0)
            End While

            ' 处理完成后禁用计时器
            tmrCheck.Enabled = False
        Catch ex As Exception
            ' 记录错误日志
            LogToFile("Error.txt", $"frmLotManager.tmrCheck:{ex.HResult} {ex.Message}")
        End Try
    End Sub
    '=======================================================
    'Routine: frmLotManager.txtOpenLot_MouseDown()
    'Purpose: We need to trap the right mouse down on this
    'to enable poping the form to modify the Lot Status
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
    Private Sub txtOpenLot_MouseDown(sender As Object, e As MouseEventArgs) Handles txtOpenLot.MouseDown
        ' 捕获右键点击
        If e.Button = MouseButtons.Right Then
            ' 显示上下文菜单
            If txtOpenLot.ContextMenuStrip IsNot Nothing Then
                txtOpenLot.ContextMenuStrip.Show(txtOpenLot, e.Location)
            End If
        End If
    End Sub

    Private Sub txtOpenLot_MouseUp(sender As Object, e As MouseEventArgs) Handles txtOpenLot.MouseUp
        ' 处理鼠标释放事件
        If e.Button = MouseButtons.Left Then  ' 仅处理左键抬起（根据原代码意图调整）
            ProcessState(cnSTATE_OPEN)
        End If
    End Sub
    '=======================================================
    'Routine: frmLotManager.m_oLotManager_LotDeleted(str,dt)
    'Purpose: This is the event raised by the business
    'server when a lot is deleted. It is then up to this
    'procedure to syncronize the this Form and the global
    'chart object.
    '
    'Globals: go_ChartData - The data related to Groups
    '         currently being tracked by the main Chart.
    '
    'Input: strGroupId - The name of the batch.
    '
    '       dtBirth - The birthday for the lot (part of
    '       the Id,Date Lot key).
    '
    'Return:None
    '
    'Modifications:
    '   12-10-1998 As written for Pass1.5
    '
    '
    '=======================================================
    Private Sub m_oLotManager_LotDeleted(ByVal strGroupId As String, ByVal dtBirth As Date) Handles m_oLotManager.LotDeleted
        ' 扫描列表并从屏幕列表对象和数组跟踪对象中移除批次
        For nListIndex As Integer = lstClosed.Items.Count - 1 To 0 Step -1
            ' 先查找匹配的名称
            If lstClosed.Items(nListIndex).ToString() = strGroupId Then
                ' 再测试日期集合（注意：VB.NET 集合索引从 0 开始）
                If m_oClosedQueue(nListIndex).dtBirth = dtBirth Then
                    ' 与业务服务器同步
                    lstClosed.Items.RemoveAt(nListIndex)
                    m_oClosedQueue.Remove(nListIndex)
                    ' 退出循环
                    Exit For
                End If
            End If
        Next nListIndex
    End Sub
    '==========================================================
    'Routine: frmLotManager.m_oLotManager_LotSuspended(str,dt)
    'Purpose: This is the event raised by the business
    'server when a lot is suspended. It is then up to this
    'procedure to syncronize this Form and remove the Group
    'from the appropriate object.
    '
    'Globals: gcol_ChartData - The data related to Groups
    '         currently being tracked by the main Chart.
    '
    'Input: strGroupId - The name of the batch.
    '
    '       dtBirth - The birthday for the lot (part of
    '       the Id,Date Lot key).
    '
    'Return:None
    '
    'Tested: Hand tested 1-21-1999 C. Barker
    '
    'Modifications:
    '   12-10-1998 As written for Pass1.5
    '
    '   05-11-1999 If the security is less than 6 than we
    '   need to automatically create the next Lot Id
    '
    '   08-09-2001 Added in csb when this triggers a Lot
    '   changing to suspended state.
    '
    '   07-08-2002 Chris Barker: Change the v2.0.8 insert
    '   sequence to do a descending order sort. Also,
    '   make sure the 1st position is focused.
    '==========================================================
    Private Sub m_oLotManager_LotSuspended(ByVal strGroupId As String, ByVal dtBirth As Date) Handles m_oLotManager.LotSuspended
        Dim nListIndex As Integer
        Dim nIdx As Integer
        Dim m_oQueueTemp As New colLotMangerQueue ' 假设为自定义队列类
        Dim bFound As Boolean = False

        Try
            ' 检查当前打开的批次
            If m_strOpenLot = strGroupId AndAlso m_dtOpen = dtBirth Then
                ' 清除打开的批次
                m_strOpenLot = String.Empty

                ' 同步界面：获取下一个批次 ID
                txtOpenLot.Text = mdlLotManager.GetNextGroupId(m_frmActiveSink)

                ' 自动生成新批次（根据安全级别）
                If m_nSecurity < 6 Then
                    cmdCreateNew_Click()
                End If

                ' 处理状态（确保窗体可见）
                If Me.Visible Then
                    ProcessState(cnSTATE_OPEN)
                End If
            Else
                ' 从已关闭列表中移除批次（从后往前遍历避免索引问题）
                For nListIndex = lstClosed.Items.Count - 1 To 0 Step -1
                    If lstClosed.Items(nListIndex).ToString() = strGroupId Then
                        ' 假设 m_oClosedQueue 元素为 clsLotMangerQueue 类型
                        Dim closedItem As clsLotManagerQueue = TryCast(m_oClosedQueue(nListIndex), clsLotManagerQueue)
                        If closedItem IsNot Nothing AndAlso closedItem.dtBirth = dtBirth Then
                            ' 同步界面和队列
                            If lstClosed.InvokeRequired Then
                                lstClosed.Invoke(Sub() lstClosed.Items.RemoveAt(nListIndex))
                            Else
                                lstClosed.Items.RemoveAt(nListIndex)
                            End If

                            m_oClosedQueue.Remove(nListIndex) ' 假设支持 RemoveAt
                            Exit For
                        End If
                    End If
                Next
            End If

            ' 添加到已暂停列表（支持排序或追加）
            If chkSortSuspendedList.Checked Then ' 假设 Value=1 对应 Checked=True
                bFound = False
                m_oQueueTemp = New colLotMangerQueue() ' 重置临时队列

                For nIdx = 1 To m_oSuspendedQueue.Count + 1
                    If nIdx > m_oSuspendedQueue.Count Then
                        If Not bFound Then
                            If lstSuspended.InvokeRequired Then
                                lstSuspended.Invoke(Sub() lstSuspended.Items.Add(strGroupId))
                            Else
                                lstSuspended.Items.Add(strGroupId)
                            End If
                            ' 假设 Add 方法接受参数（根据 clsLotMangerQueue 定义调整）
                            m_oQueueTemp.Add(strGroupId, cnNull, cnNull, dtBirth)
                        End If
                    Else
                        ' 从队列中获取元素（假设索引从 0 开始，nIdx-1 为 VB.NET 索引）
                        Dim currentItem As clsLotManagerQueue = TryCast(m_oSuspendedQueue(nIdx - 1), clsLotManagerQueue)
                        If currentItem IsNot Nothing AndAlso strGroupId > currentItem.sLotId AndAlso Not bFound Then
                            If lstSuspended.InvokeRequired Then
                                lstSuspended.Invoke(Sub() lstSuspended.Items.Insert(nIdx - 1, strGroupId))
                            Else
                                lstSuspended.Items.Insert(nIdx - 1, strGroupId)
                            End If
                            m_oQueueTemp.Add(strGroupId, cnNull, cnNull, dtBirth)
                            bFound = True
                        End If
                        ' 添加现有元素到临时队列
                        If currentItem IsNot Nothing Then
                            m_oQueueTemp.Add(currentItem.sLotId, currentItem.nFrom, currentItem.nTo, currentItem.dtBirth)
                        End If
                    End If
                Next

                'sort lstSuspended and m_oSuspendedQueue
                Dim items(lstSuspended.Items.Count - 1) As Object
                lstSuspended.Items.CopyTo(items, 0)
                Array.Sort(items)
                Array.Reverse(items)
                lstSuspended.Items.Clear()
                lstSuspended.Items.AddRange(items)

                m_oSuspendedQueue = m_oQueueTemp ' 赋值临时队列
                m_oSuspendedQueue.SortDescending()
            Else
                If lstSuspended.InvokeRequired Then
                    lstSuspended.Invoke(Sub() lstSuspended.Items.Add(strGroupId))
                Else
                    lstSuspended.Items.Add(strGroupId)
                End If
                ' 假设 Add 方法接受参数（根据 clsLotMangerQueue 定义调整）
                m_oSuspendedQueue.Add(strGroupId, cnNull, cnNull, dtBirth)
            End If

            ' 触发批次挂起事件
            mdlSAX.ExecuteCSB_LotSuspended(strGroupId)

        Catch ex As Exception
            LogToFile("Error.txt", $"frmLotManager.m_oLotManager_LotSuspended:{ex.HResult} {ex.Message}")
            Err.Clear()
        End Try
    End Sub
    '==========================================================
    'Routine: frmLotManager.m_oLotManager_LotClosed(str,dt)
    'Purpose: This is the event raised by the business
    'server when a lot is closed. It is then up to this
    'procedure to syncronize this Form and remove the Group
    'from the appropriate object.
    '
    'Globals: gcol_ChartData - The data related to Groups
    '         currently being tracked by the main Chart.
    '
    'Input: strGroupId - The name of the batch.
    '
    '       dtBirth - The birthday for the lot (part of
    '       the Id,Date Lot key).
    '
    'Return:None
    '
    'Tested: Hand tested 1-21-1999 C. Barker
    '
    'Modifications:
    '   12-10-1998 As written for Pass1.5
    '
    '   05-11-1999 If the security is less than 6 than we
    '   need to automatically create the next Lot Id
    '
    '   08-10-1999 Added code to ensure that the global
    '   context lot id is blanked out when the active
    '   lot is closed.
    '
    '   08-02-2001 Added in code to hook the EndOfLot csb
    '   to the end of this routine. v1.1.1 C. Barker
    '
    '   07-08-2002 Chris Barker: Change the v2.0.8 insert
    '   sequence to do a descending order sort. Also,
    '   make sure the 1st position is focused.
    '==========================================================
    Private Sub m_oLotManager_LotClosed(ByVal strGroupId As String, ByVal dtBirth As Date) Handles m_oLotManager.LotClosed
        Dim nListIndex As Integer
        Dim nIdx As Integer
        Dim m_oQueueTemp As New colLotMangerQueue
        Dim bFound As Boolean = False

        Try
            ' 检查当前打开的批次
            If m_strOpenLot = strGroupId AndAlso m_dtOpen = dtBirth Then
                ' 清除打开的批次
                m_strOpenLot = String.Empty
                go_Context.GroupId = String.Empty

                ' 同步界面：获取下一个批次 ID
                txtOpenLot.Text = mdlLotManager.GetNextGroupId(m_frmActiveSink)

                ' 自动生成新批次（根据安全级别）
                If m_nSecurity < 6 Then
                    cmdCreateNew_Click()
                End If

                ' 处理状态（确保窗体可见）
                If Me.Visible Then
                    ProcessState(cnSTATE_OPEN)
                End If
            Else
                ' 从已暂停列表中移除批次（从后往前遍历避免索引问题）
                For nListIndex = lstSuspended.Items.Count - 1 To 0 Step -1
                    If lstSuspended.Items(nListIndex).ToString() = strGroupId Then
                        ' 假设 m_oSuspendedQueue 元素为 clsLotMangerQueue 类型
                        Dim suspendedItem As clsLotManagerQueue = TryCast(m_oSuspendedQueue(nListIndex), clsLotManagerQueue)
                        If suspendedItem IsNot Nothing AndAlso suspendedItem.dtBirth = dtBirth Then
                            ' 同步界面和队列
                            If lstSuspended.InvokeRequired Then
                                lstSuspended.Invoke(Sub() lstSuspended.Items.RemoveAt(nListIndex))
                            Else
                                lstSuspended.Items.RemoveAt(nListIndex)
                            End If

                            m_oSuspendedQueue.Remove(nListIndex) ' 假设支持 RemoveAt
                            Exit For
                        End If
                    End If
                Next
            End If

            ' 添加到已关闭列表（支持排序或追加）
            If chkSortClosedList.Checked Then ' 假设 Value=1 对应 Checked=True
                bFound = False
                m_oQueueTemp = New colLotMangerQueue() ' 重置临时队列

                For nIdx = 1 To m_oClosedQueue.Count + 1
                    If nIdx > m_oClosedQueue.Count Then
                        If Not bFound Then
                            If lstClosed.InvokeRequired Then
                                lstClosed.Invoke(Sub() lstClosed.Items.Add(strGroupId))
                            Else
                                lstClosed.Items.Add(strGroupId)
                            End If
                            m_oQueueTemp.Add(strGroupId, cnNull, cnNull, dtBirth)
                        End If
                    Else
                        ' 从队列中获取元素（假设索引从 0 开始，nIdx-1 为 VB.NET 索引）
                        Dim currentItem As clsLotManagerQueue = TryCast(m_oClosedQueue(nIdx - 1), clsLotManagerQueue)
                        If currentItem IsNot Nothing AndAlso strGroupId > currentItem.sLotId AndAlso Not bFound Then
                            If lstClosed.InvokeRequired Then
                                lstClosed.Invoke(Sub() lstClosed.Items.Insert(nIdx - 1, strGroupId))
                            Else
                                lstClosed.Items.Insert(nIdx - 1, strGroupId)
                            End If
                            m_oQueueTemp.Add(strGroupId, cnNull, cnNull, dtBirth)
                            bFound = True
                        End If
                        ' 添加现有元素到临时队列
                        If currentItem IsNot Nothing Then
                            m_oQueueTemp.Add(currentItem.sLotId, currentItem.nFrom, currentItem.nTo, currentItem.dtBirth)
                        End If
                    End If
                Next
                'sort lstClosed and m_oClosedQueue
                Dim items(lstClosed.Items.Count - 1) As Object
                lstClosed.Items.CopyTo(items, 0)
                Array.Sort(items)
                Array.Reverse(items)
                lstClosed.Items.Clear()
                lstClosed.Items.AddRange(items)

                m_oClosedQueue = m_oQueueTemp ' 赋值临时队列
                m_oClosedQueue.SortDescending()
            Else
                If lstClosed.InvokeRequired Then
                    lstClosed.Invoke(Sub() lstClosed.Items.Add(strGroupId))
                Else
                    lstClosed.Items.Add(strGroupId)
                End If
                m_oClosedQueue.Add(strGroupId, cnNull, cnNull, dtBirth)
            End If

            ' 调用 CSB 钩子 EndOfLot
            mdlSAX.ExecuteCSB_EndOfLot()

        Catch ex As Exception
            LogToFile("Error.txt", $"frmLotManager.m_oLotManager_LotClosed:{ex.HResult} {ex.Message}")
            Err.Clear()
        End Try
    End Sub
    '==========================================================
    'Routine: frmLotManager.m_oLotManager_LotOpened(str,dt)
    'Purpose: This is the event raised by the business
    'server when a lot is Opened. It is then up to this
    'procedure to syncronize this Form and remove the Group
    'from the appropriate object.
    '
    'Globals: gcol_ChartData - The data related to Groups
    '         currently being tracked by the main Chart.
    '
    'Input: strGroupId - The name of the batch.
    '
    '       dtBirth - The birthday for the lot (part of
    '       the Id,Date Lot key).
    '
    'Return:None
    '
    'Tested: Hand tested 1-21-1999 C. Barker
    '
    'Modifications:
    '   12-10-1998 As written for Pass1.5
    '
    '   01-19-1999 Added the sync of go_Context.LotId
    '   and dtBirth with this form so that it will
    '   be continuously matched up. This is also present
    '   in RecvOpenLot().
    '
    '   01-19-1999 Added the ProcessState(cnSTATE_OPEN) to
    '   the end of this procedure since it rings false
    '   from the Button as there is no open Lot when
    '   it occurs.
    '
    '   08-09-2001 Added in csb hooks for Closed->Open
    '   Suspended -> Open and ->Open. C Barker
    '==========================================================
    Private Sub m_oLotManager_LotOpened(ByVal strGroupId As String, ByVal dtBirth As Date) Handles m_oLotManager.LotOpened
        Dim nListIndex As Integer
        Dim matchFound As Boolean = False

        Try
            ' 先扫描已关闭列表
            For nListIndex = lstClosed.Items.Count - 1 To 0 Step -1 ' 从后往前遍历避免索引问题
                If lstClosed.Items(nListIndex).ToString() = strGroupId Then
                    If m_oClosedQueue(nListIndex).dtBirth = dtBirth Then ' VB.NET 索引从 0 开始，移除 +1 偏移
                        lstClosed.Items.RemoveAt(nListIndex)
                        m_oClosedQueue.Remove(nListIndex)
                        mdlSAX.ExecuteCSB_ClosedLotOpened(strGroupId)
                        matchFound = True
                        Exit For ' 找到后退出循环
                    End If
                End If
            Next

            If Not matchFound Then ' 未在已关闭列表找到，扫描已暂停列表
                For nListIndex = lstSuspended.Items.Count - 1 To 0 Step -1
                    If lstSuspended.Items(nListIndex).ToString() = strGroupId Then
                        If m_oSuspendedQueue(nListIndex).dtBirth = dtBirth Then
                            lstSuspended.Items.RemoveAt(nListIndex)
                            m_oSuspendedQueue.Remove(nListIndex)
                            mdlSAX.ExecuteCSB_SuspendedLotOpened(strGroupId)
                            matchFound = True
                            Exit For
                        End If
                    End If
                Next
            End If

            ' 设置打开批次信息
            m_strOpenLot = strGroupId
            txtOpenLot.Text = m_strOpenLot
            m_dtOpen = dtBirth
            go_Context.GroupId = m_strOpenLot
            go_Context.GroupBirthDay = m_dtOpen
            mdlSAX.ExecuteCSB_LotOpened(strGroupId)

            ' 处理状态（确保窗体可见）
            If Me.Visible Then
                ProcessState(cnSTATE_OPEN)
            End If

        Catch ex As Exception
            LogToFile("Error.txt", $"frmLotManager.m_oLotManager_LotOpened:{ex.HResult} {ex.Message}")
            Err.Clear()
        End Try
    End Sub

End Class
