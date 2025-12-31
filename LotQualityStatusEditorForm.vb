Imports NextCapServer = ncapsrv

Public Class frmLotStatusEditor

    '/*Storage for are current state
    'Private m_clsWindow As clsWindowState
    Private m_frmParent As Form

    '/*Storage for a reference to the LotManager that holds this Group
    Private WithEvents m_oLotManager As NextCapServer.clsLotManager
    Private m_oLotManagerSink As frmLotManagerSink

    '/*Private stash of the Group's Birthday
    Private m_dtBirth As Date

    '/*Window dimension Mins before screen breakdown
    Private Const clngMinHeight As Integer = 7575
    Private Const clngMinWidth As Integer = 6210

    '/*State for tracking if a Comment is being edited
    Private m_nState As Integer
    Private Const cnSTATE_READY As Integer = 0
    Private Const cnSTATE_ABORT As Integer = 8
    Private Const cnSTATE_COMMENT As Integer = 1
    Private Const cnSTATE_NEWDEFAULTCOMMENT As Integer = 5
    Private Const cnSTATE_NEWMANUALCOMMENT As Integer = 6
    Private Const cnSTATE_EDITCOMMENT As Integer = 2
    Private Const cnSTATE_DELETECOMMENT As Integer = 3
    Private Const cnSTATE_STATUS As Integer = 4
    Private Const cnSTATE_LOTVALUES As Integer = 7

    Private toolTip As New ToolTip()

    '/*Comparison storage for Shipment Data
    Private m_strUnitShipped As String
    Private m_strBatchCount As String
    Private m_n100percent As Integer

    '/*Comparison for Quality Status
    Private m_strQuality As String
    Private m_bQualityChange As Boolean
    '/*Comparison for Comment
    Private m_bCommentSelect As Boolean

    '/*Security object
    Private WithEvents m_oSecurity As Security
    Private m_nSecurity As Integer

    '/*This forms current stack ID
    Private m_strFrmId As String

    Public ReadOnly Property strFrmId() As String
        Get
            Return m_strFrmId
        End Get
    End Property

    '
    '=======================================================
    'Routine: frmLotStatus.ShowWindow()
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
        '/*Set a reference to a new window state
        If Len(m_strFrmId) > 0 Then mdlWindow.RemoveForm(m_strFrmId)
        m_strFrmId = mdlWindow.AddForm(Me)
    End Sub

    '
    '=======================================================
    'Routine: frmLotStatus.HideWindow()
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
        '/*Hide this window
        mdlWindow.RemoveForm(m_strFrmId)
        m_strFrmId = ""
    End Sub

    '
    '=======================================================
    'Routine: frmLotStatus.WindowInit()
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
    '   02-15-1999 Added resume next to guard against
    '   errors for bad list selections on cmbEvents.
    '=======================================================
    Private Sub WindowInit()
        Try
            '/*Set the default on the drop down list
            'cmbEvents.SelectedIndex = 0
            '/*Set the default state
            If ProcessState(cnSTATE_READY) Then
                '/*Do nothing for now
            End If
        Catch ex As Exception
            ' 处理异常
        End Try
    End Sub

    '
    '=======================================================
    'Routine: frmLotStatus.UnitsShipped(b,str)
    'Purpose: This alllows an entry point for determining
    'what the title of the Units shipped label is and
    'whether it is visible.
    '
    'Globals:None
    '
    'Input: bVivisble - True/False for allowing this option
    '       to be visible on the screen.
    '
    '       strDesc - The caption for the label.
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
    Public Sub UnitsShipped(Optional ByVal bVisible As Boolean = True, Optional ByVal strDesc As String = "")
        '/*determine if this is used
        If bVisible = True Or Len(strDesc) > 0 Then
            '/*Insure the items are visible
            lblUnitShipped.Visible = True
            txtUnitShipped.Visible = True
        Else
            '/*Insure the items are visible
            lblUnitShipped.Visible = False
            txtUnitShipped.Visible = False
        End If
        '/*Set the description in the label caption regardless of visible
        lblUnitShipped.Text = strDesc
    End Sub

    '
    '=======================================================
    'Routine: frmLotStatus.UnitsInBatch(b,str)
    'Purpose: This alllows an entry point for determining
    'what the title of the Units in Batch label is and
    'whether it is visible.
    '
    'Globals:None
    '
    'Input: bVivisble - True/False for allowing this option
    '       to be visible on the screen.
    '
    '       strDesc - The caption for the label.
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
    Public Sub UnitsInBatch(Optional ByVal bVisible As Boolean = True, Optional ByVal strDesc As String = "")
        '/*determine if this is used
        If bVisible = True Or Len(strDesc) > 0 Then
            '/*Insure the items are visible
            lblUnitBatch.Visible = True
            txtBatchCount.Visible = True
        Else
            '/*Insure the items are visible
            lblUnitBatch.Visible = False
            txtBatchCount.Visible = False
        End If
        '/*Set the description in the label caption regardless of visible
        lblUnitBatch.Text = strDesc
    End Sub

    '
    '=======================================================
    'Routine: frmLotStatus.OneHundredPercent(b,str)
    'Purpose: This alllows an entry point for determining
    'what the title of the Units in Batch label is and
    'whether it is visible.
    '
    'Globals:None
    '
    'Input: bVivisble - True/False for allowing this option
    '       to be visible on the screen.
    '
    '       strDesc - The caption for the label.
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
    Public Sub OneHundredPercent(Optional ByVal bVisible As Boolean = True, Optional ByVal strDesc As String = "")
        '/*determine if this is used
        If bVisible = True Or Len(strDesc) > 0 Then
            '/*Insure the items are visible
            chk100Percent.Visible = True
        Else
            '/*Insure the items are visible
            chk100Percent.Visible = False
        End If
        '/*Set the description in the label caption regardless of visible
        chk100Percent.Text = strDesc
    End Sub

    Private Sub chk100Percent_MouseDown(sender As Object, e As MouseEventArgs)
        If ProcessState(cnSTATE_LOTVALUES) Then
            '/*Do nothing for now
        End If
    End Sub

    '/*Attempt to trap when a comment has been selected
    Private Sub cmbEvents_Click(sender As Object, e As EventArgs)
        If ProcessState(cnSTATE_NEWDEFAULTCOMMENT) Then
            m_bCommentSelect = True
        End If
    End Sub

    '/*Standard Click event for Abort button
    Private Sub cmdAbort_Click(sender As Object, e As EventArgs) Handles cmdAbort.Click
        If m_nState = cnSTATE_READY Then
            HideWindow()
            '/*Release the lock on the client
            gb_InputBusyFlag = False
        Else
            '/*Reset the values to original
            If m_nState = cnSTATE_LOTVALUES Then
                chk100Percent.Checked = m_n100percent = 1
                txtBatchCount.Text = m_strBatchCount
                txtUnitShipped.Text = m_strUnitShipped
            ElseIf m_nState = cnSTATE_NEWMANUALCOMMENT Then
                txtComment.Text = ""
            ElseIf m_nState = cnSTATE_DELETECOMMENT Then
                cmdDelete.Enabled = False
            End If
            ProcessState(cnSTATE_ABORT)
        End If
    End Sub

    Private Sub cmdDelete_Click(sender As Object, e As EventArgs) Handles cmdDelete.Click
        If ProcessState(cnSTATE_DELETECOMMENT) Then
            ProcessState(cnSTATE_READY)
        End If
    End Sub

    '/*Standard Click event for Abort button
    Private Sub cmdOK_Click(sender As Object, e As EventArgs) Handles cmdOK.Click
        If m_nState = cnSTATE_READY Then
            HideWindow()
            '/*Release the lock on the client
            gb_InputBusyFlag = False
        Else
            ProcessState(cnSTATE_READY)
        End If
        Me.Close()
    End Sub

    '/*Standard Form event
    Private Sub Form_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        WindowInit()

        '/*Call funciton to disable QueryUnload
        mdlTools.DisableCloseX(Me)
        '/*Populate the Qulaity List
        GenerateQualityList()
        '/*Set the FullRowSelect option on the DataGridView for comments
        mdlTools.SetDataGridViewStyle(lstvwComments)
        '/*Setup columns for the DataGridView
        If lstvwComments.Columns.Count = 0 Then
            lstvwComments.Columns.Add("DateColumn", "Date")
            lstvwComments.Columns.Add("UserColumn", "User")
            lstvwComments.Columns.Add("CommentColumn", "Comment")
        End If
        '/*Get a reference to the Security object
        m_oSecurity = go_Security
        If m_oSecurity IsNot Nothing Then
            m_nSecurity = m_oSecurity.nAuthority
        End If
    End Sub

    '
    '=======================================================
    'Routine: frmLotStatusEditor.SetGroup(str,dt)
    'Purpose: This is an external interface for a caller
    'to set the Group being modified in this window.
    '
    'Globals:None
    '
    'Input:None
    '
    'Return:None
    '
    'Modifications:
    '   12-14-1998 As written for Pass1.5
    '
    '
    '=======================================================
    Public Sub SetGroup(ByVal strGroupId As String, ByVal dtBirth As Date, frmActiveSink As frmLotManagerSink)
        Dim nIndex As Integer
        Dim strIcon As String

        '/*Set a reference to the ActiveSink that holds this group
        'm_oLotManager = frmActiveSink.LotManager
        m_oLotManager = go_ActiveLotManager
        '/*Set a reference to the LotManager
        m_oLotManagerSink = frmActiveSink
        '/*Set the group id
        lblGroupId.Text = strGroupId
        '/*Store the group's birthday
        m_dtBirth = dtBirth
        '/*Now lookup the current quality status
        nIndex = frmActiveSink.ChartDataCol.GroupIdExist(strGroupId, dtBirth)
        If nIndex > 0 Then
            m_strQuality = frmActiveSink.ChartDataCol(nIndex).strIconName
            strIcon = frmActiveSink.ChartDataCol(nIndex).strIconPath
            '/*Set the Quality displayed on the screen
            SetQualityState(m_strQuality, strIcon)
        End If
        '/*Now get the current set of comments
        GenerateCurrentCommentsList(frmActiveSink)
        '/*Get the current values for shipment options
        GenerateCurrentShipment(frmActiveSink)
    End Sub

    '
    '=======================================================
    'Routine: frmLotStatusEditor.strQualityState(str)
    'Purpose: This sets the screens ques to let the user
    'know what the current state is.
    '
    'Globals:None
    '
    'Input: strQuality - The quality idnetifier.
    '
    'Return:None
    '
    'Tested:
    '
    'Modifications:
    '   04-13-1999 As written for Pass1.8
    '
    '
    '=======================================================
    Private Sub SetQualityState(ByVal strQuality As String, ByVal strIcon As String)
        Try
            '/*Set the image quality icon
            If Len(strIcon) = 0 Or UCase(strIcon) = "BLANK" Then
                '/*Set the image to blank
                imgQuality.Image = Nothing
            ElseIf FileExist(strIcon) Then
                '/*Set the image to the specified file
                imgQuality.Image = Image.FromFile(strIcon)
            Else
                '/*Set the image as missing
                'imgQuality.Image = Form1.imglst_tbrNextCap.Images(15)
            End If


            '/*Set a reference
            With lstQualityStatus
                '/*Now select the item in the list
                For nItem As Integer = 0 To .Items.Count - 1
                    '/*See if this matches the passed in name
                    If .Items(nItem).ToString() = strQuality Then
                        '/*Select the item
                        .SetSelected(nItem, True)
                        Exit For
                    End If
                Next nItem
            End With

            '/*Set the current label
            lblCurrent.Text = strQuality
        Catch ex As Exception
            ' 处理异常
        End Try
    End Sub

    '
    '===============================================================
    'Routine: frmLotStatusEditor.GenerateCurrentCommentsList(frm)
    'Purpose: This fills the list for the current set of comments.
    '
    'Globals:None
    '
    'Input: frmActiveSink - This gives us access to the LotManager.
    '
    'Return:None
    '
    'Tested:
    '
    'Modifications:
    '   02-02-1998 As written for Pass1.7
    '
    '
    '=================================================================
    Private Sub GenerateCurrentCommentsList(ByRef frmActiveSink As frmLotManagerSink)
        Dim vrtArray As DataTable

        ' 清空评论列表
        lstvwComments.Rows.Clear()

        ' 获取评论列表（如果有）
        If mdlLotManager.RecvComments(frmActiveSink.LotManager, lblGroupId.Text, m_dtBirth, vrtArray) Then
            ' 添加评论到列表
            If vrtArray IsNot Nothing AndAlso vrtArray.Rows.Count > 0 Then
                ' 确保DataGridView有足够的列
                If lstvwComments.Columns.Count < 3 Then
                    lstvwComments.Columns.Clear()
                    lstvwComments.Columns.Add("DateColumn", "Date")
                    lstvwComments.Columns.Add("UserColumn", "User")
                    lstvwComments.Columns.Add("CommentColumn", "Comment")
                End If

                ' 添加每一行数据
                For Each row As DataRow In vrtArray.Rows
                    ' 提取数据
                    Dim commentDate As DateTime = Convert.ToDateTime(row(0))
                    Dim userName As String = Convert.ToString(row(1))
                    Dim commentText As String = Convert.ToString(row(2))

                    ' 添加到DataGridView
                    lstvwComments.Rows.Add(commentDate.ToString("yyyy-MM-dd hh:mm:ss tt"), userName, commentText)
                Next
            End If
        End If
    End Sub

    '
    '===============================================================
    'Routine: frmLotStatusEditor.GenerateCurrentShipment(frm)
    'Purpose: This fills out the options for Group data. Note that
    'the use of a loop to unload the variant array is protection
    'against an unexpected array size.
    '
    'Globals:None
    '
    'Input: frmActiveSink - This gives us access to the LotManager.
    '
    'Return:None
    '
    'Tested:
    '
    'Modifications:
    '   02-15-1998 As written for Pass1.7
    '
    '
    '=================================================================
    Private Sub GenerateCurrentShipment(ByRef frmActiveSink As frmLotManagerSink)
        Dim vrtArray As DataTable
        Dim nItem As Integer

        ' 获取发货信息
        If mdlLotManager.RecvShipment(frmActiveSink.LotManager, lblGroupId.Text, m_dtBirth, vrtArray) Then
            ' 检查 DataTable 是否有效且包含数据
            If vrtArray IsNot Nothing AndAlso vrtArray.Rows.Count > 0 Then
                ' 处理每一行数据（根据原代码逻辑）
                For nItem = 0 To vrtArray.Rows.Count - 1
                    Dim row As DataRow = vrtArray.Rows(nItem)

                    ' 提取并设置各值（注意：每次循环会覆盖前一次的值，最终只保留最后一行的数据）
                    txtUnitShipped.Text = Convert.ToString(row(0)) ' 批次发货数量
                    m_strUnitShipped = txtUnitShipped.Text

                    txtBatchCount.Text = Convert.ToString(row(1))  ' 批次总数量
                    m_strBatchCount = txtBatchCount.Text

                    ' 将布尔值转换为复选框状态
                    Dim inspect100Percent As Boolean = Convert.ToBoolean(row(2))
                    chk100Percent.Checked = inspect100Percent  ' 使用 Checked 属性
                    m_n100percent = If(inspect100Percent, 1, 0)  ' 存储为 1/0
                Next
            End If
        End If
    End Sub

    '
    '=======================================================
    'Routine: frmLotStatus.GenerateQualityList()
    'Purpose: This builds up the listing for the Quality
    'Statuses.
    '
    'Globals: gcol_QualityStatus - The collection of
    'known Quality States.
    '
    'Input:None
    '
    'Return:None
    '
    'Tested:
    '
    'Modifications:
    '   02-11-1998 As written for Pass1.7
    '
    '   04-06-1999 Changed from.strMessage to.strName
    '   for the QualityStates. Also added in the suppresion
    '   of the UNKNOWN state.
    '=======================================================
    Private Sub GenerateQualityList()
        Dim nItem As Integer
        Const strUnknown As String = "UNKNOWN"
        Dim strName As String
        Dim nCount As Integer

        '/*Set a reference to the list
        With lstQualityStatus
            '/*Make sure we start with an empty list
            .Items.Clear()
            '/*Loop through our listing of Quality Statuses
            For nItem = 0 To gcol_QualityStatus.Count - 1
                '/*Buffer to save on the number of calls to the object
                strName = gcol_QualityStatus(nItem).strName
                '/*Make sure we are not presenting the unknown state
                If strName <> strUnknown Then
                    '/*Translate the items at -1 to the collection
                    .Items.Add(strName)
                    nCount = nCount + 1
                End If
            Next nItem
        End With
    End Sub

    '
    '=======================================================
    'Routine: frmLotStatusEditor.GenerateCommentsList()
    'Purpose: This scans the listing of station-keys to
    'get all of the predefined comments for Lots. The
    'station-key object will return "" when it reaches the
    'end of its listing.
    '
    'Globals: gcol_StationKeys - The Station object that
    'contains all of the stations parameters returned during
    'client startup.
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
    Private Sub GenerateCommentsList()
        Dim nItem As Integer
        Dim nKey As Integer
        Dim strTemp As String

        '/*Make sure the drop down is cleared
        cmbEvents.Items.Clear()
        '/*Loop through the station-keys to get all of the Comments
        For nKey = 1 To gcol_StationKeys.Count
            nItem = 0 '/*Prime the index counter
            strTemp = gcol_StationKeys(nKey).LotComments(nItem) '/*Get the first Comment
            Do While Len(strTemp) > 0
                If Not CommentUsed(strTemp) Then cmbEvents.Items.Add(strTemp) '/*Add teh item ot the list
                strTemp = gcol_StationKeys(nKey).LotComments(nItem) '/*Get the next Comment
            Loop
        Next nKey
    End Sub

    '
    '=======================================================
    'Routine: frmLotStatusEditor.CommentUsed(str)
    'Purpose: This scans the Combo box for Comments and
    'returns true if a comment has been used.
    '
    'Globals:None
    '
    'Input: strComment - The comment to add to the list.
    '
    'Return: Boolean = True if the item is in the list.
    '
    'Tested:
    '
    'Modifications:
    '   02-02-1998 As written for Pass1.7
    '
    '
    '=======================================================
    Private Function CommentUsed(ByVal strComment As String) As Boolean
        Dim nItem As Integer

        '/*Scan the list and exit if found
        For nItem = 0 To cmbEvents.Items.Count - 1
            '/*Test for a match
            If cmbEvents.Items(nItem).ToString() = strComment Then
                '/*Set true and return right away
                CommentUsed = True
                Exit For
            End If
        Next nItem
        Return CommentUsed
    End Function

    'Modifications:
    '   08-15-2001 New call to make sure the TopForm is not blocked.
    Private Sub Form_Paint(sender As Object, e As PaintEventArgs)
        mdlWindow.CheckWhosTop(Me)
    End Sub

    '
    '=======================================================
    'Routine: frmLotStatusEditor.Form_Resize()
    'Purpose: This resizes all of the visible forms objects
    'when the user resizes the form.
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
    Private Sub Form_Resize(sender As Object, e As EventArgs)
        Dim nMiddle As Integer
        Dim nTop As Integer
        Dim nLeft As Integer

        '/*Insure the user is not Min or Max the window
        If Me.WindowState = FormWindowState.Normal Then
            '/*Insure that the user does not shrink below the minimum screen size
            If Me.Width < clngMinWidth Then Me.Width = clngMinWidth
            If Me.Height < clngMinHeight Then Me.Height = clngMinHeight

            '/*Resize the Group Id
            lblGroupId.Width = Me.Width - 915

            '/*Determine where the Top of the comments should be
            If frmLotOptions.Visible Then
                '/*Lot Quality Statuses
                'lstQualityStatus.Move 360, 960, Width * 0.3695, Height * 0.16237
                lstQualityStatus.Location = New Point(360, 960)
                lstQualityStatus.Width = CInt(Me.Width * 0.2343)
                lstQualityStatus.Height = CInt(Me.Height * 0.16237)

                '/*Move the image quality position
                imgQuality.Left = lstQualityStatus.Left + lstQualityStatus.Width + 50
                lblCurrent.Left = imgQuality.Left
                '/*The Lot options such as Samples, Batch Count, etc.
                'nLeft = lstQualityStatus.Left + lstQualityStatus.Width + 285
                nLeft = imgQuality.Left + imgQuality.Width + 285
                frmLotOptions.Location = New Point(nLeft, 840)
                frmLotOptions.Width = Me.Width - nLeft - 555
                frmLotOptions.Height = lstQualityStatus.Height + 105
                'nTop = frmLotOptions.Top + frmLotOptions.Height + 105
            Else
                '/*Lot Quality Statuses
                lstQualityStatus.Location = New Point(360, 960)
                lstQualityStatus.Width = lblGroupId.Width
                lstQualityStatus.Height = CInt(Me.Height * 0.085)
                'nTop = lstQualityStatus.Top + lstQualityStatus.Height + 75
            End If
            '/*Determine where the top of the next box will be
            nTop = frmLotOptions.Top + frmLotOptions.Height + 125
            '/*Now move the ListView of comments already attached to the Lot
            lstvwComments.Location = New Point(360, nTop)
            lstvwComments.Width = lblGroupId.Width
            lstvwComments.Height = Me.Height - 5760

            '/*Adjust the height of the label above the Common Event pull down
            lblCommonEvent.Top = lstvwComments.Top + lstvwComments.Height + 105
            '/*Adjust the pull down for the Common Events
            cmbEvents.Location = New Point(360, lblCommonEvent.Top + lblCommonEvent.Height - 15)
            cmbEvents.Width = lblGroupId.Width

            '/*Adjust the Top of the free form comment label
            lblComments.Top = cmbEvents.Top + cmbEvents.Height + 105
            '/*Adjust the position of the free form comment input box
            txtComment.Location = New Point(360, lblComments.Top + lblComments.Height - 15)
            txtComment.Width = lblGroupId.Width

            '/*Reposition the Top of the OK button
            cmdOK.Top = Me.Height - 1095

            '/*Reposition the Edit/Delete buttons
            nMiddle = Me.Width \ 2
            cmdEdit.Location = New Point(nMiddle, cmdOK.Top)
            cmdDelete.Location = New Point(nMiddle - cmdDelete.Width, cmdOK.Top)

            '/*Reposition the Abort button
            cmdAbort.Location = New Point(Me.Width - 1770, cmdOK.Top)
        End If
    End Sub

    '
    '=======================================================
    'Routine: frmLotStatusEditor.ProcessState(n)
    'Purpose: This process the requested state to track if
    'we are editing or requesting to delete a comment.
    '
    'Globals:None
    '
    'Input: nState - The requested new state.
    '
    'Return: Boolean - State change granted.
    '
    'Tested:
    '
    'Modifications:
    '   02-15-1998 As written for Pass1.7
    '
    '
    '=======================================================
    Private Function ProcessState(ByVal nState As Integer) As Boolean
        Dim strComment As String
        Dim nIndex As Integer

        '/*Make sure we are not requesting the existing state
        If m_nState = nState Then
            ProcessState = True
            Exit Function
        End If

        '/*Process the a request to switch the state to the Comment list
        If (nState = cnSTATE_COMMENT) And (m_nState <> cnSTATE_COMMENT) Then
            '/*Disable the commands
            'cmdEdit.Enabled = True
            cmdDelete.Enabled = True
            '/*Set the new state
            m_nState = cnSTATE_COMMENT
            '/*Return grant
            ProcessState = True
        ElseIf m_nState = cnSTATE_COMMENT And nState = cnSTATE_DELETECOMMENT And (m_nSecurity >= 6 Or m_nSecurity = -1) Then
            If lstvwComments.Rows.Count > 0 Then
                If DeleteComment() Then
                    nIndex = lstvwComments.SelectedRows(0).Index
                    lstvwComments.Rows.RemoveAt(nIndex)
                End If
            End If
            m_nState = cnSTATE_READY
            ProcessState = True
        ElseIf nState = cnSTATE_NEWDEFAULTCOMMENT And (m_nSecurity >= 4 Or m_nSecurity = -1) Then
            toolTip.SetToolTip(cmdOK, "Add Default Comment")
            m_nState = cnSTATE_NEWDEFAULTCOMMENT
            ProcessState = True
        ElseIf nState = cnSTATE_NEWMANUALCOMMENT And (m_nSecurity >= 4 Or m_nSecurity = -1) Then
            toolTip.SetToolTip(cmdOK, "Add Comment")
            m_nState = cnSTATE_NEWMANUALCOMMENT
            ProcessState = True
        ElseIf nState = cnSTATE_LOTVALUES Then
            toolTip.SetToolTip(cmdOK, "Change Shipment Data")
            m_nState = cnSTATE_LOTVALUES
            ProcessState = True
        ElseIf nState = cnSTATE_STATUS And (m_nSecurity >= 5 Or m_nSecurity = -1) Then
            toolTip.SetToolTip(cmdOK, "Change Status")
            m_nState = cnSTATE_STATUS
            ProcessState = True
        ElseIf nState = cnSTATE_ABORT Then
            '/*Toggle the tool tip for the OK button
            toolTip.SetToolTip(cmdOK, "Press OK to close screen")
            '/*Disable the commands
            cmdEdit.Enabled = False
            cmdDelete.Enabled = False
            '/*Set the new state
            m_nState = cnSTATE_READY
            '/*Return grant
            ProcessState = True
        ElseIf nState = cnSTATE_READY Then
            '/*Determine what the current state is and process
            If m_nState = cnSTATE_NEWDEFAULTCOMMENT Then 'Add a comment from the drop down
                If cmbEvents.Items.Count > 0 Then
                    strComment = cmbEvents.SelectedItem.ToString()
                    If Not AddComment(strComment) Then
                        Exit Function
                    End If
                End If
            ElseIf m_nState = cnSTATE_NEWMANUALCOMMENT Then 'Add a comment from the drop down
                If txtComment.Text.Length > 0 Then
                    If Not AddComment(txtComment.Text) Then
                        Exit Function
                    End If
                    txtComment.Text = ""
                End If
            ElseIf m_nState = cnSTATE_DELETECOMMENT Then
                If lstvwComments.Rows.Count > 0 Then
                    If DeleteComment() Then
                        nIndex = lstvwComments.SelectedRows(0).Index
                        lstvwComments.Rows.RemoveAt(nIndex)
                    End If
                End If
            ElseIf m_nState = cnSTATE_LOTVALUES Then
                If Not UpdateLotValues() Then
                    Exit Function
                End If
            ElseIf m_nState = cnSTATE_STATUS Then
                If Not UpdateStatus() Then
                    Exit Function
                End If
            End If
            '/*Toggle the tool tip for the OK button
            toolTip.SetToolTip(cmdOK, "Press OK to close screen")
            '/*Disable the commands
            cmdEdit.Enabled = False
            cmdDelete.Enabled = False
            '/*Set the new state
            m_nState = cnSTATE_READY
            '/*Return grant
            ProcessState = True
        End If

        Return ProcessState
    End Function

    Private Sub Form_Closing(sender As Object, e As FormClosingEventArgs)
        '/*Destroy the reference to the business server
        m_oLotManager = Nothing
    End Sub

    Private Sub frmLotOptions_MouseDown(sender As Object, e As MouseEventArgs)
        If ProcessState(cnSTATE_LOTVALUES) Then

        End If
    End Sub

    '/*This should indicate a change in the quality status
    Private Sub lstQualityStatus_Click(sender As Object, e As EventArgs)
        If ProcessState(cnSTATE_STATUS) Then
            m_strQuality = lstQualityStatus.Items(lstQualityStatus.SelectedIndex).ToString()
            m_bQualityChange = True
        End If
    End Sub

    Private Sub lstQualityStatus_MouseDown(sender As Object, e As MouseEventArgs)
        If ProcessState(cnSTATE_STATUS) Then
            '/*Do nothing for now
        End If
    End Sub

    '
    '=======================================================
    'Routine: frmLotStatusEditor.lstvwComments_MouseDown()
    'Purpose: This is used to change the state of the Edit
    'and Delete buttons.
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
    '=======================================================
    Private Sub lstvwComments_MouseDown(sender As Object, e As MouseEventArgs)
        If ProcessState(cnSTATE_COMMENT) Then
            '/*Do nothing for now
        End If
    End Sub

    '=======================================================
    'Routine: m_oLotManager_ClosedLotMaterialStatus
    '         m_oLotManager_UpdateMaterialStatus
    '         ChangeStatusIcon
    '
    'Purpose: These are used to catch the change in the
    'status icon and name for the Lot while this screen is
    'active. It also acts as the second half of the change
    'Status procedure.
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
    '   04-19-1999 As written for Pass1.8
    '
    '
    '=======================================================
    Private Sub m_oLotManager_ClosedLotMaterialStatus(ByVal strLotId As String, ByVal dateBirthday As Date, ByVal strQualityStatus As String)
        ChangeStatusIcon(strLotId, dateBirthday, strQualityStatus)
    End Sub

    Private Sub m_oLotManager_UpdateMaterialStatus(ByVal strLotId As String, ByVal dateBirthday As Date, ByVal strQualityStatus As String)
        ChangeStatusIcon(strLotId, dateBirthday, strQualityStatus)
    End Sub

    Private Sub ChangeStatusIcon(ByVal strLotId As String, ByVal dateBirthday As Date, ByVal strQualityStatus As String)
        Dim strIcon As String

        '/*Trap any lots that are not what we are looking at
        If strLotId = lblGroupId.Text And dateBirthday = m_dtBirth Then
            '/*Check if we already have an Icon loaded for this
            '/*group of data
            If Len(strQualityStatus) <> 0 Then
                '/*Filter out any requests to reset the image
                If UCase(strQualityStatus) = "RESET" Then
                    strIcon = ""
                    m_strQuality = "Good"
                    '/*Get the icon path form the QualityStatus object
                ElseIf gcol_QualityStatus.IsUsed(strQualityStatus) Then
                    '/*Set the pathway that we are holding for the image
                    strIcon = gcol_QualityStatus(strQualityStatus).strImagePath
                    m_strQuality = strQualityStatus
                End If
            Else
                m_strQuality = "Unknown"
            End If
            '/*Set the Quality displayed on the screen
            SetQualityState(m_strQuality, strIcon)
        End If
    End Sub

    Private Sub m_oSecurity_SecurityChange(ByVal nNewAuthority As Integer)
        '/*Set the new Security state
        m_nSecurity = nNewAuthority
    End Sub

    Private Sub txtBatchCount_MouseDown(sender As Object, e As MouseEventArgs)
        If ProcessState(cnSTATE_LOTVALUES) Then
            '/*Do nothing for now
        End If
    End Sub

    Private Sub txtComment_MouseDown(sender As Object, e As MouseEventArgs)
        If ProcessState(cnSTATE_NEWMANUALCOMMENT) Then

        End If
    End Sub

    Private Sub txtUnitShipped_MouseDown(sender As Object, e As MouseEventArgs)
        If ProcessState(cnSTATE_LOTVALUES) Then
            '/*Do nothing for now
        End If
    End Sub

    '
    '=======================================================
    'Routine:
    'Purpose:
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
    '   04-19-1999 As written for Pass1.8
    '
    '
    '=======================================================
    Private Function UpdateStatus() As Boolean
        '/*Update the Quality Status if necessary
        If Not m_oLotManager.ChangeQualityStatus(lblGroupId.Text,
                                                m_dtBirth, m_strQuality) Then
            '/*If this fails abort the procedure and alert the user
            'frmMessage.GenerateMessage("Lot Status Error #2" & Environment.NewLine & "Quality Status update failed. Press abort to cancel update")
            MessageBox.Show("Lot Status Error #2" & Environment.NewLine & "Quality Status update failed. Press abort to cancel update")
            Return False
        Else
            UpdateStatus = True
            Return UpdateStatus
        End If
    End Function

    '
    '=======================================================
    'Routine:
    'Purpose:
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
    '   04-19-1999 As written for Pass1.8
    '
    '
    '=======================================================
    Private Function DeleteComment() As Boolean
        If lstvwComments.Rows.Count > 0 Then
            Dim dtComment As Date
            If Date.TryParse(lstvwComments.SelectedRows(0).Cells(0).Value.ToString(), dtComment) Then
                m_oLotManager.DeleteComment(lblGroupId.Text, m_dtBirth, dtComment)
                'GenerateCurrentCommentsList(m_oLotManager)
                GenerateCurrentCommentsList(m_oLotManagerSink)
            End If
        End If
        Return True
    End Function

    '
    '=======================================================
    'Routine:
    'Purpose:
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
    '   04-19-1999 As written for Pass1.8
    '
    '
    '=======================================================
    Private Function AddComment(ByVal strComment As String) As Boolean
        '/*Update the server to reflect the change
        If Not m_oLotManager.AddComment(lblGroupId.Text,
                                        m_dtBirth, strComment,
                                        go_Context.Operator) Then
            '/*If this fails abort the procedure and alert the user
            'frmMessage.GenerateMessage("Lot Status Error #1" & Environment.NewLine & "Comment could not be added. Press abort to cancel update")
            MessageBox.Show("Lot Status Error #1" & Environment.NewLine & "Comment could not be added. Press abort to cancel update")
            Return False
        Else
            '/*Rebuild the comment list
            GenerateCurrentCommentsList(m_oLotManagerSink)
            AddComment = True
            Return AddComment
        End If
    End Function

    '
    '=======================================================
    'Routine:
    'Purpose:
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
    '   04-19-1999 As written for Pass1.8
    '
    '
    '=======================================================
    Private Function UpdateLotValues() As Boolean
        Dim b100Percent As Boolean

        '/*Test to see if we need to update the Shipment Data
        If m_n100percent <> If(chk100Percent.Checked, 1, 0) Or
            m_strBatchCount <> txtBatchCount.Text Or
            m_strUnitShipped <> txtUnitShipped.Text Then

            '/*Make sure the Lot >= Shipped
            If Integer.Parse(txtUnitShipped.Text) <= Integer.Parse(txtBatchCount.Text) Then
                '/*Convert the check box to boolean
                b100Percent = chk100Percent.Checked
                '/*Update the server to reflect the change
                If Not m_oLotManager.SetShipmentData(lblGroupId.Text,
                                                m_dtBirth, Integer.Parse(txtUnitShipped.Text),
                                                Integer.Parse(txtBatchCount.Text), b100Percent) Then
                    '/*If this fails abort the procedure and alert the user
                    'frmMessage.GenerateMessage("Lot Status Error #1" & Environment.NewLine & "Shipment Data update failed. Press abort to cancel update")
                    MessageBox.Show("Lot Status Error #1" & Environment.NewLine & "Shipment Data update failed. Press abort to cancel update")
                    Return False
                Else
                    UpdateLotValues = True
                    Return UpdateLotValues
                End If
            Else
                '/*If this fails abort the procedure and alert the user
                'frmMessage.GenerateMessage("Lot Status Error #2" & Environment.NewLine & "Pens shipped is less than Lot count. Press abort to cancel update")
                MessageBox.Show("Lot Status Error #2" & Environment.NewLine & "Pens shipped is less than Lot count. Press abort to cancel update")
                Return False
            End If
        End If
        Return True
    End Function

    ' 以下是可能缺失的一些函数定义模拟，根据实际情况调整
    Private Function vtoa(ByVal value As Object) As String
        Return Convert.ToString(value)
    End Function

    Private Function vtob(ByVal value As Object) As Boolean
        Return Convert.ToBoolean(value)
    End Function

    Private Function vtoDate(ByVal value As String) As Date
        Dim result As Date
        If Date.TryParse(value, result) Then
            Return result
        End If
        Return Date.MinValue
    End Function

    Private Function atoi(ByVal value As String) As Integer
        Dim result As Integer
        If Integer.TryParse(value, result) Then
            Return result
        End If
        Return 0
    End Function

    Private Function FileExist(ByVal filePath As String) As Boolean
        Return System.IO.File.Exists(filePath)
    End Function

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles txtUnitShipped.TextChanged

    End Sub

End Class