Imports System.Windows.Forms
Imports System.Media

Public Class frmGoodPen
    Inherits Form

    ' Controls
    Private picture1 As PictureBox
    Private label1 As Label
    Private timer1 As Timer

    ' Sound
    Private soundPlayer As SoundPlayer

    ' Stack ID
    Private m_strFrmId As String

    ' Property to get stack ID
    Public ReadOnly Property strFrmId As String
        Get
            Return m_strFrmId
        End Get
    End Property

    Public Sub New()
        ' Form settings
        Me.Text = "Feedback"
        Me.FormBorderStyle = FormBorderStyle.FixedToolWindow
        Me.StartPosition = FormStartPosition.CenterParent
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.ShowInTaskbar = False
        Me.ClientSize = New Drawing.Size(340, 556)

        ' PictureBox
        picture1 = New PictureBox() With {
            .Location = New Drawing.Point(-12, 0),
            .Size = New Drawing.Size(361, 577),
            .SizeMode = PictureBoxSizeMode.StretchImage
        }
        ' Set your image here if available:
        ' picture1.Image = Image.FromFile("your_image_path.png")

        ' Label
        label1 = New Label() With {
            .Text = "Good Pen!",
            .Font = New Drawing.Font("Arial Black", 19.5, Drawing.FontStyle.Regular),
            .ForeColor = SystemColors.HotTrack,
            .BackColor = Drawing.Color.Transparent,
            .TextAlign = Drawing.ContentAlignment.MiddleCenter,
            .Location = New Drawing.Point(60, 0),
            .Size = New Drawing.Size(241, 73)
        }

        ' Timer
        timer1 = New Timer() With {
            .Interval = 1000
        }
        AddHandler timer1.Tick, AddressOf Timer1_Tick

        ' Add controls
        picture1.Controls.Add(label1)
        Me.Controls.Add(picture1)
    End Sub

    ' Show the feedback window and play sound
    Public Sub ShowWindow()
        WindowInit()
        If Not String.IsNullOrEmpty(m_strFrmId) Then
            mdlWindow.RemoveForm(m_strFrmId)
        End If
        m_strFrmId = mdlWindow.AddForm(Me)
        Me.Show()
        PlayFeedbackSound()
        timer1.Start()
    End Sub

    ' Hide the feedback window and remove from stack
    Public Sub HideWindow()
        mdlWindow.RemoveForm(m_strFrmId)
        m_strFrmId = ""
        Me.Hide()
    End Sub

    ' Window initialization (add any prep here)
    Private Sub WindowInit()
        ' Add any other screen prep calls here
    End Sub

    ' Play feedback sound (replace path as needed)
    Private Sub PlayFeedbackSound()
        Try
            ' Update the path to your WAV file as needed
            Dim wavPath As String = "C:\WINDOWS\Media\Office97\Cashreg.wav"
            If IO.File.Exists(wavPath) Then
                soundPlayer = New SoundPlayer(wavPath)
                soundPlayer.Play()
            End If
        Catch
            ' Ignore sound errors
        End Try
    End Sub

    ' Timer event: stop sound and hide window
    Private Sub Timer1_Tick(sender As Object, e As EventArgs)
        timer1.Stop()
        If soundPlayer IsNot Nothing Then
            soundPlayer.Stop()
        End If
        Me.Hide()
    End Sub

    ' Optional: handle form activation to replay sound if needed
    Protected Overrides Sub OnActivated(e As EventArgs)
        MyBase.OnActivated(e)
        PlayFeedbackSound()
    End Sub
End Class
