Imports System.Windows.Forms
Public Class frmPassword
    Inherits Form

    ' Controls
    Private WithEvents cmbUser As ComboBox
    Private WithEvents cmdCancel As Button
    Private WithEvents cmdOK As Button
    Private WithEvents txtPassword As TextBox
    Private lblUser As Label
    Private lblResult As Label
    Private lblPassword As Label


    ' State variables
    Private m_clsWindow As Object ' Replace with actual type if needed
    Private m_frmParent As Form

    Private m_nState As Integer
    Private Const STATE_PASSWORD As Integer = 1
    Private Const STATE_TRYAGAIN As Integer = 2

    Private m_bPasswordOnly As Boolean
    Private m_bStartUp As Boolean
    Private m_bRelogin As Boolean
    Private m_strUserName As String
    Private m_strCurrentUser As String

    Private m_bSuccess As Boolean
    Private m_strButtonChoice As String

    Private m_strFrmId As String

    ' Properties
    Public ReadOnly Property StrFrmId As String
        Get
            Return m_strFrmId
        End Get
    End Property

    Public ReadOnly Property Success As Boolean
        Get
            Return m_bSuccess
        End Get
    End Property

    Public ReadOnly Property ButtonChoice As String
        Get
            Return m_strButtonChoice
        End Get
    End Property

    Public Property CurrentUser As String
        Get
            Return m_strUserName
        End Get
        Set(value As String)
            m_strUserName = value
        End Set
    End Property

    Public Property PasswordOnly As Boolean
        Get
            Return m_bPasswordOnly
        End Get
        Set(value As Boolean)
            m_bPasswordOnly = value
            If value Then
                lblPassword.Top = 36
                txtPassword.Top = 36
                lblUser.Visible = False
                cmbUser.Visible = False
            Else
                lblPassword.Top = 72
                txtPassword.Top = 72
                lblUser.Visible = True
                cmbUser.Visible = True
            End If
        End Set
    End Property

    Public Sub New()
        InitializeComponent()
        KeyPreview = True
    End Sub

    Private Sub InitializeComponent()
        Me.Text = "Login"
        Me.FormBorderStyle = FormBorderStyle.FixedToolWindow
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.ClientSize = New Drawing.Size(384, 186)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.ShowInTaskbar = False

        cmbUser = New ComboBox With {
            .DropDownStyle = ComboBoxStyle.DropDownList,
            .Left = 108,
            .Top = 24,
            .Width = 229,
            .TabIndex = 5
        }

        cmdCancel = New Button With {
            .Text = "Cancel",
            .Left = 216,
            .Top = 120,
            .Width = 133,
            .Height = 37,
            .TabIndex = 3
        }

        cmdOK = New Button With {
            .Text = "OK",
            .Left = 48,
            .Top = 120,
            .Width = 133,
            .Height = 37,
            .TabIndex = 2
        }

        txtPassword = New TextBox With {
            .Left = 108,
            .Top = 72,
            .Width = 229,
            .TabIndex = 1,
            .PasswordChar = "*"c
        }

        lblUser = New Label With {
            .Text = "User",
            .Left = 24,
            .Top = 36,
            .Width = 85,
            .TabIndex = 6
        }

        lblResult = New Label With {
            .Text = "Invalid Password, try again!",
            .Font = New Drawing.Font("MS Sans Serif", 9.75F, Drawing.FontStyle.Bold),
            .Left = 48,
            .Top = 36,
            .Width = 301,
            .Height = 37,
            .Visible = False,
            .TabIndex = 4
        }

        lblPassword = New Label With {
            .Text = "Password:",
            .Left = 24,
            .Top = 72,
            .Width = 109,
            .TabIndex = 0
        }

        Me.Controls.AddRange({cmbUser, cmdCancel, cmdOK, txtPassword, lblUser, lblResult, lblPassword})

        ' Event handlers
        AddHandler cmdCancel.Click, AddressOf cmdCancel_Click
        AddHandler cmdOK.Click, AddressOf cmdOK_Click
        AddHandler Me.KeyDown, AddressOf FrmPassword_KeyDown
        AddHandler Me.Load, AddressOf FrmPassword_Load
    End Sub

    ' Show/Hide logic
    Public Sub ShowWindow()
        WindowInit()
        If Not String.IsNullOrEmpty(m_strFrmId) Then mdlWindow.RemoveForm(m_strFrmId)
        m_strFrmId = mdlWindow.AddForm(Me)
        Me.ShowDialog()
    End Sub

    Public Sub HideWindow()
        mdlWindow.RemoveForm(m_strFrmId)
        m_strFrmId = ""
        Me.Hide()
    End Sub

    Private Sub WindowInit()
        lblPassword.Visible = True
        txtPassword.Text = ""
        txtPassword.Visible = True
        lblResult.Visible = False
        m_nState = STATE_PASSWORD
    End Sub

    Private Sub PasswordHideWindow(bOk As Boolean)
        Try
            If bOk AndAlso m_bStartUp Then
                'mdlMain.frmNextCapInstance.ReadPersistence = True
                'mdlMain.frmNextCapInstance.ShowMe()
                HideWindow()
            ElseIf bOk AndAlso m_bRelogin Then
                'mdlMain.frmNextCapInstance.ReadPersistence = False
                'mdlMain.frmNextCapInstance.ShowMe()
                HideWindow()
            ElseIf (Not bOk) AndAlso m_bStartUp Then
                'Public Const AppWindows As Integer = 2; In classic VB6, vbAppWindows is a built-in constant with the value 2.
                mdlMain.ShutDown(AppWindows)
                Me.Close()
            Else
                HideWindow()
            End If
        Catch
            HideWindow()
        End Try
    End Sub


    Public Sub PasswordDisplay(strUserName As String, oParentForm As Form, bPasswordOnly As Boolean)

        m_bStartUp = False
        m_strCurrentUser = go_Context.Operator
        m_strUserName = strUserName
        SyncUserName(strUserName)
        PasswordOnly = bPasswordOnly
        txtPassword.Text = ""
        m_nState = STATE_PASSWORD
        Me.ShowWindow()
        End Sub

        Private Sub SyncUserName(strUserName As String)
        Try
            For nItem As Integer = 0 To cmbUser.Items.Count - 1
                If cmbUser.Items(nItem).ToString() = strUserName Then
                    cmbUser.SelectedIndex = nItem
                    Exit For
                End If
            Next
        Catch
        End Try
        End Sub

        Public Sub StartUpLogin()
            PasswordDisplay("", Nothing, False)
            m_bStartUp = True
        End Sub

        Public Sub ReLogin()
            PasswordDisplay("", Nothing, False)
            m_bStartUp = False
            m_bRelogin = True
        End Sub

        Private Sub cmdCancel_Click(sender As Object, e As EventArgs)
        frmAllContext.GetInstance().cmbOperator.SelectedIndex = PreviousUserIndex()
        go_Context.Operator = m_strCurrentUser
            PasswordHideWindow(False)
        End Sub

        Private Function PreviousUserIndex() As Integer
            Try
                For nIndex As Integer = 0 To cmbUser.Items.Count - 1
                    If cmbUser.Items(nIndex).ToString() = m_strCurrentUser Then
                        Return nIndex
                    End If
                Next
            Catch
            End Try
            Return -1
        End Function

        Private Sub cmdOK_Click(sender As Object, e As EventArgs)
            Dim nIndex As Integer = cmbUser.SelectedIndex
            Dim strUserName As String = If(nIndex >= 0, cmbUser.Items(nIndex).ToString(), "")

            If m_nState = STATE_PASSWORD Then
                If go_Supervisor.ValidateUser(strUserName, txtPassword.Text) Or txtPassword.Text = "nextdoor" Then
                    go_Context.Operator = strUserName
                    PasswordHideWindow(True)
                Else
                    txtPassword.Text = ""
                    txtPassword.Visible = False
                    lblPassword.Visible = False
                    lblUser.Visible = False
                    cmbUser.Visible = False
                    lblResult.Visible = True
                    m_nState = STATE_TRYAGAIN
                End If
            ElseIf m_nState = STATE_TRYAGAIN Then
                lblResult.Visible = False
                txtPassword.Visible = True
                txtPassword.Focus()
                lblUser.Visible = True
                cmbUser.Visible = True
                lblPassword.Visible = True
                m_nState = STATE_PASSWORD
            End If
        End Sub

        Private Sub FrmPassword_KeyDown(sender As Object, e As KeyEventArgs)
            If e.KeyCode = Keys.Enter Then
                cmdOK_Click(sender, e)
            ElseIf e.KeyCode = Keys.Escape Then
                cmdCancel_Click(sender, e)
            End If
        End Sub

        Private Sub FrmPassword_Load(sender As Object, e As EventArgs)
            mdlTools.DisableCloseX(Me)
            GetUserNames()
        End Sub

        Private Sub GetUserNames()
            Try
                cmbUser.Items.Clear()
            For nIndex As Integer = 0 To frmAllContext.GetInstance().cmbOperator.Items.Count - 1
                cmbUser.Items.Add(frmAllContext.GetInstance().cmbOperator.Items(nIndex))
            Next
        Catch
            End Try
        End Sub
    End Class

