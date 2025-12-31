Imports System.Windows.Forms

Public Class frmMessage
    Inherits Form

    Private WithEvents cmdOK As New Button()
    Public lblMessage As New Label()

    Private m_bSuppress As Boolean
    Private m_strFrmId As String = ""

    Public Property strFrmId() As String
        Get
            Return m_strFrmId
        End Get
        Private Set(value As String)
            m_strFrmId = value
        End Set
    End Property

    Public Sub New()
        Me.FormBorderStyle = FormBorderStyle.SizableToolWindow
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.Text = "Message"
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.ShowInTaskbar = False
        Me.ClientSize = New Drawing.Size(508, 232)

        cmdOK.Text = "OK"
        cmdOK.Size = New Drawing.Size(241, 37)
        cmdOK.Location = New Drawing.Point(132, 192)

        lblMessage.Text = "Label1"
        lblMessage.BackColor = Drawing.Color.FromArgb(255, 255, 128)
        lblMessage.BorderStyle = BorderStyle.FixedSingle
        lblMessage.Size = New Drawing.Size(481, 169)
        lblMessage.Location = New Drawing.Point(12, 12)
        lblMessage.AutoSize = False

        Me.Controls.Add(cmdOK)
        Me.Controls.Add(lblMessage)
    End Sub

    Public Sub GenerateMessage(strMsg As String, Optional strTitle As String = "Error", Optional bSuppressUnlock As Boolean = False)
        Me.Text = strTitle
        lblMessage.BackColor = If(strTitle = "Error", Drawing.Color.Yellow, SystemColors.Control)
        lblMessage.Text = strMsg
        m_bSuppress = bSuppressUnlock
        Me.ShowDialog()
    End Sub

    Public Sub LockedMessage(strMsg As String, Optional strTitle As String = "Error")
        cmdOK.Enabled = False
        Me.Text = strTitle
        lblMessage.Text = strMsg
        m_bSuppress = True
        Me.ShowDialog()
    End Sub

    Public Sub UnlockMessage()
        cmdOK.Enabled = True
        Me.Close()
    End Sub

    Private Sub cmdOK_Click(sender As Object, e As EventArgs) Handles cmdOK.Click
        If Not m_bSuppress Then
            ' Example: global flag reset if needed
            ' gb_InputBusyFlag = False
        End If
        Me.Close()
    End Sub
End Class
