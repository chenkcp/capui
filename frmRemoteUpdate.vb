Imports System.Windows.Forms

Public Class frmRemoteUpdate
    Inherits Form

    ' Controls
    Private lblCaption As Label
    Private WithEvents cmdOK As Button
    Private WithEvents cmdCancel As Button

    ' State
    Private m_strId As String
    Private m_strFrmId As String

    ' Property: Stack ID
    Public ReadOnly Property strFrmId As String
        Get
            Return m_strFrmId
        End Get
    End Property

    ' Property: SearchId
    Public Property SearchId As String
        Get
            Return m_strId
        End Get
        Set(value As String)
            m_strId = value
        End Set
    End Property

    Public Sub New()
        ' Form settings
        Me.Text = "Remote Update"
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.ClientSize = New Drawing.Size(411, 156)
        Me.ShowInTaskbar = False

        ' Label
        lblCaption = New Label() With {
            .BorderStyle = BorderStyle.FixedSingle,
            .Text = "Perform remote search from Informix for pen id to edit? WARNING, this can take a while.",
            .Location = New Drawing.Point(24, 12),
            .Size = New Drawing.Size(361, 73),
            .TextAlign = Drawing.ContentAlignment.MiddleLeft,
            .AutoSize = False
        }

        ' OK Button
        cmdOK = New Button() With {
            .Text = "OK",
            .Location = New Drawing.Point(24, 96),
            .Size = New Drawing.Size(157, 49),
            .TabIndex = 0
        }

        ' Cancel Button
        cmdCancel = New Button() With {
            .Text = "Cancel",
            .Location = New Drawing.Point(228, 96),
            .Size = New Drawing.Size(157, 49),
            .TabIndex = 1
        }

        ' Add controls
        Me.Controls.Add(lblCaption)
        Me.Controls.Add(cmdOK)
        Me.Controls.Add(cmdCancel)
    End Sub

    ' Show the form and manage stack ID
    Public Sub ShowWindow()
        WindowInit()
        If Not String.IsNullOrEmpty(m_strFrmId) Then
            mdlWindow.RemoveForm(m_strFrmId)
        End If
        m_strFrmId = mdlWindow.AddForm(Me)
        Me.Show()
    End Sub

    ' Prompt for pen search
    Public Sub PromptPenSearch(ByRef strId As String, frmParent As Form, Optional ByVal bModal As Boolean = False)
        SearchId = strId
        ShowWindow()
    End Sub

    ' Hide the form and remove from stack
    Public Sub HideWindow()
        mdlWindow.RemoveForm(m_strFrmId)
        m_strFrmId = ""
        Me.Hide()
    End Sub

    ' Window initialization (add any prep here)
    Private Sub WindowInit()
        ' Add any other screen prep calls here
    End Sub

    ' Cancel button click: hide and release lock
    Private Sub cmdCancel_Click(sender As Object, e As EventArgs) Handles cmdCancel.Click
        HideWindow()
        Gb_InputBusyFlag = False
    End Sub

    ' OK button click: perform remote search, show result in defect editor
    Private Sub cmdOK_Click(sender As Object, e As EventArgs) Handles cmdOK.Click
        cmdOK.Enabled = False
        cmdCancel.Enabled = False

        Dim oUnit As clsPen = mdlCreatePen.RetrieveRemotePen(m_strId)
        HideWindow()

        If oUnit IsNot Nothing Then
            frmDefectEditor.SetUnit(oUnit, True)
            frmDefectEditor.ShowWindow()
        Else
            Gb_InputBusyFlag = False
        End If

        cmdOK.Enabled = True
        cmdCancel.Enabled = True
    End Sub

    ' Prevent closing from the X button
    Protected Overrides Sub OnLoad(e As EventArgs)
        MyBase.OnLoad(e)
        mdlTools.DisableCloseX(Me)
    End Sub
End Class
