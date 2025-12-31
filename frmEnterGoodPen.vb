Imports System.Windows.Forms

Public Class frmEnterGoodPen
    Inherits Form

    ' Controls
    Public WithEvents cmbRecovery As ComboBox
    Private WithEvents cmbNumber As ComboBox
    Private WithEvents cmdAbort As Button
    Private WithEvents cmdOK As Button
    Private lblNumber As Label
    Private lblRecovery As Label

    ' State
    Private m_oUnit As clsPen
    Private m_strFrmId As String

    ' External/global references (to be adapted as needed)
    ' Private gb_InputBusyFlag As Boolean
    ' Private go_clsSystemSettings As Object
    ' Private frmNextCap As frmNextCap
    ' Private mdlCreatePen As Object
    ' Private mdlLotManager As Object
    ' Private mdlTools As Object
    ' Private frmGoodPen As frmGoodPen

    Public Sub New()
        Me.Text = "Good Pen Entry"
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.StartPosition = FormStartPosition.CenterParent
        Me.KeyPreview = True
        Me.ClientSize = New Drawing.Size(609, 228)

        ' Initialize controls
        cmbRecovery = New ComboBox With {
            .DropDownStyle = ComboBoxStyle.DropDownList,
            .Location = New Drawing.Point(240, 84),
            .Size = New Drawing.Size(325, 31),
            .TabIndex = 5
        }
        cmbNumber = New ComboBox With {
            .Location = New Drawing.Point(36, 84),
            .Size = New Drawing.Size(157, 28),
            .TabIndex = 4,
            .Text = "cmbNumber"
        }
        cmdAbort = New Button With {
            .Text = "Abort",
            .Location = New Drawing.Point(468, 144),
            .Size = New Drawing.Size(97, 49),
            .TabIndex = 1
        }
        cmdOK = New Button With {
            .Text = "OK",
            .Location = New Drawing.Point(36, 144),
            .Size = New Drawing.Size(97, 49),
            .TabIndex = 0
        }
        lblNumber = New Label With {
            .Text = "Number of Good Pens",
            .Location = New Drawing.Point(36, 36),
            .Size = New Drawing.Size(169, 25),
            .TabIndex = 3
        }
        lblRecovery = New Label With {
            .Text = "Pen Recovery Procedure",
            .Location = New Drawing.Point(240, 36),
            .Size = New Drawing.Size(265, 25),
            .TabIndex = 2
        }

        ' Add controls to form
        Me.Controls.AddRange({cmbRecovery, cmbNumber, cmdAbort, cmdOK, lblNumber, lblRecovery})
    End Sub
    Public ReadOnly Property SelectedRecoveryStep As String
        Get
            If cmbRecovery.SelectedItem IsNot Nothing Then
                Return cmbRecovery.SelectedItem.ToString()
            Else
                Return String.Empty
            End If
        End Get
    End Property
    Public ReadOnly Property strFrmId As String
        Get
            Return m_strFrmId
        End Get
    End Property

    Public Sub ShowWindow()
        WindowInit()
        ' m_strFrmId = mdlWindow.AddForm(Me) ' Implement as needed
        Me.Show()
        cmdOK.Focus()
    End Sub

    Public Sub HideWindow()
        ' mdlWindow.RemoveForm(m_strFrmId) ' Implement as needed
        m_strFrmId = ""
        Me.Hide()
    End Sub

    Private Sub WindowInit()
        If cmbNumber.Items.Count > 0 Then cmbNumber.SelectedIndex = 0
        If cmbRecovery.Items.Count > 0 Then cmbRecovery.SelectedIndex = 0
        cmdOK.Focus()
    End Sub

    Private Sub cmbNumber_Click(sender As Object, e As EventArgs) Handles cmbNumber.SelectedIndexChanged
        cmdOK.Focus()
    End Sub

    Private Sub cmbRecovery_Click(sender As Object, e As EventArgs) Handles cmbRecovery.SelectedIndexChanged
        cmdOK.Focus()
    End Sub

    Private Sub cmdAbort_Click(sender As Object, e As EventArgs) Handles cmdAbort.Click
        HideWindow()
        Gb_InputBusyFlag = False
        mdlMain.frmNextCapInstance.FocusInput()
    End Sub

    Private Sub cmdOK_Click(sender As Object, e As EventArgs) Handles cmdOK.Click
        cmdOK.Enabled = False
        If m_oUnit IsNot Nothing Then
            m_oUnit.nCount = cmbNumber.Text
            m_oUnit.strRecoveryStep = cmbRecovery.Text
            HideWindow()
            If mdlCreatePen.TransmitUnit(m_oUnit) Then
                Gb_InputBusyFlag = False
                mdlMain.frmNextCapInstance.FocusInput()
                If go_clsSystemSettings.bGoodPenFeedBack Then
                    frmGoodPen.ShowWindow()
                End If
            Else
                mdlLotManager.TransactionFailure("mdlCreatePen.TransmitUnit()")
            End If
        End If
        cmdOK.Enabled = True
    End Sub

    Private Sub frmEnterGoodPen_GotFocus(sender As Object, e As EventArgs) Handles Me.GotFocus
        cmdOK.Focus()
    End Sub

    Protected Overrides Sub OnKeyDown(e As KeyEventArgs)
        MyBase.OnKeyDown(e)
        If e.KeyCode = Keys.Enter Then
            cmdOK_Click(cmdOK, EventArgs.Empty)
        ElseIf e.KeyCode = Keys.Escape Then
            cmdAbort_Click(cmdAbort, EventArgs.Empty)
        End If
    End Sub

    Private Sub frmEnterGoodPen_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        mdlTools.DisableCloseX(Me)
        ' Me.Icon = frmNextCap.imglst_tbrNextCap.Images(10) ' Set icon as needed
    End Sub

    Public Sub SetUnitId(Optional ByVal strUnitId As String = "")
        If String.IsNullOrEmpty(strUnitId) Then strUnitId = GetIdDttm()
        m_oUnit = New clsPen()
        m_oUnit.strPenId = strUnitId
        mdlCreatePen.SetUnitStdProperties(m_oUnit)
    End Sub

    Public Sub AddRecoveryStep(strData As String)
        If Not cmbRecovery.Items.Contains(strData) Then
            cmbRecovery.Items.Add(strData)
        End If
    End Sub

    Public Sub ClearRecoverySteps()
        cmbRecovery.Items.Clear()
    End Sub
End Class
