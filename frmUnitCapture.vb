Imports System.IO
Imports System.Windows.Forms

Public Class frmUnitCapture
    Inherits Form

    ' Controls
    Private txtAugmentId As TextBox
    Private txtResult As TextBox
    Private txtCount As TextBox
    Private txtAcquiredId As TextBox
    Private tmrAsyncInput As Timer
    Private label4 As Label
    Private label3 As Label
    Private label2 As Label
    Private label1 As Label

    ' State
    Private m_oweUnitInput As clsUnitId
    Private m_colProcessedUnits As List(Of String)
    Private m_strIdFile As String

    Public Sub New()
        ' Form settings
        Me.Text = "Unit Id Capture"
        Me.ClientSize = New Drawing.Size(462, 314)
        Me.StartPosition = FormStartPosition.WindowsDefaultLocation

        ' Controls
        txtAugmentId = New TextBox() With {.Location = New Drawing.Point(132, 84), .Size = New Drawing.Size(314, 28), .TabIndex = 7, .Text = ""}
        txtResult = New TextBox() With {.Location = New Drawing.Point(132, 132), .Size = New Drawing.Size(314, 28), .TabIndex = 5, .Text = ""}
        txtCount = New TextBox() With {.Location = New Drawing.Point(132, 12), .Size = New Drawing.Size(194, 28), .TabIndex = 3, .Text = ""}
        txtAcquiredId = New TextBox() With {.Location = New Drawing.Point(132, 48), .Size = New Drawing.Size(314, 28), .TabIndex = 1, .Text = ""}

        tmrAsyncInput = New Timer() ' Not used in original code, but can be wired up if needed

        label4 = New Label() With {.Text = "Augment Id", .Location = New Drawing.Point(12, 84), .Size = New Drawing.Size(98, 25), .TabIndex = 6}
        label3 = New Label() With {.Text = "Last Result", .Location = New Drawing.Point(12, 132), .Size = New Drawing.Size(110, 25), .TabIndex = 4}
        label2 = New Label() With {.Text = "IDs Available", .Location = New Drawing.Point(12, 12), .Size = New Drawing.Size(110, 25), .TabIndex = 2}
        label1 = New Label() With {.Text = "Acquired Id", .Location = New Drawing.Point(12, 48), .Size = New Drawing.Size(86, 25), .TabIndex = 0}

        Me.Controls.AddRange({txtAugmentId, txtResult, txtCount, txtAcquiredId, label4, label3, label2, label1})

        ' State
        m_colProcessedUnits = New List(Of String)()
        m_strIdFile = Path.Combine(Application.StartupPath, "id_buffer.log")
    End Sub

    Private Sub frmUnitCapture_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Initialize connection to global UnitId Input class
        m_oweUnitInput = go_CAPmain.NextCapIdInput
        ' Initialize the collection of captured IDs
        m_colProcessedUnits = New List(Of String)()
        ' Set the ID Buffer file
        m_strIdFile = Path.Combine(Application.StartupPath, "id_buffer.log")
        ' Read in the id buffer
        InitializeIdBuffer()
        ' Trim off any extra buffer
        TrimBuffer()
    End Sub

    Private Sub frmUnitCapture_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        ' Flush out the id collection
        FlushIdBuffer()
    End Sub

    ' AsyncId event handler (simulate WithEvents for .NET)
    Public Sub OnAsyncId(strId As String)
        Test_ShowStats(strId)
    End Sub

    Private Sub Test_ShowStats(ByVal strId As String)
        Me.Show()
        txtAcquiredId.Text = strId
    End Sub

    ' Verify if a unit ID is unique and valid
    Public Function VerifyUnitId(ByVal strUnitId As String) As Boolean
        If go_clsSystemSettings.bPenIdUnique Then
            If IdNotUsed(strUnitId) Then
                ' Optionally: TrackId(strUnitId)
                Return True
            End If
        Else
            Return True
        End If
        Return False
    End Function

    ' Add an ID to the private collection, keeping length <= nPenIdMemory
    Public Sub TrackId(ByVal strId As String)
        If m_colProcessedUnits.Count > go_clsSystemSettings.nPenIdMemory Then
            m_colProcessedUnits.RemoveAt(0)
        End If
        m_colProcessedUnits.Add(strId)
    End Sub

    ' Remove an ID from the entered ID stack
    Public Sub RemoveId(ByVal strId As String)
        For i As Integer = m_colProcessedUnits.Count - 1 To 0 Step -1
            If m_colProcessedUnits(i) = strId Then
                m_colProcessedUnits.RemoveAt(i)
                Exit For
            End If
        Next
    End Sub

    ' Check if an ID is not used in the collection
    Private Function IdNotUsed(ByVal strId As String) As Boolean
        For i As Integer = m_colProcessedUnits.Count - 1 To 0 Step -1
            If m_colProcessedUnits(i) = strId Then
                Return False
            End If
        Next
        Return True
    End Function

    ' Execute and monitor the acquisition of a unit ID via an external executable
    Public Function ExecutePolledId() As String
        With go_clsSystemSettings
            If mdlTools.StartProcess(.strPenIdCaptureProgram, .lngPenIdTimeOut) Then
                Return m_oweUnitInput.GetPolledId()
            End If
        End With
        Return String.Empty
    End Function

    ' Write the ID collection to the buffer file
    Private Sub FlushIdBuffer()
        Try
            Using sw As New StreamWriter(m_strIdFile, False)
                For Each id As String In m_colProcessedUnits
                    sw.WriteLine(id)
                Next
            End Using
        Catch
            ' Ignore errors
        End Try
    End Sub

    ' Read the ID buffer file into the collection
    Private Sub InitializeIdBuffer()
        Try
            If File.Exists(m_strIdFile) Then
                m_colProcessedUnits.Clear()
                Using sr As New StreamReader(m_strIdFile)
                    While Not sr.EndOfStream
                        Dim strId As String = sr.ReadLine()
                        If Not String.IsNullOrEmpty(strId) Then
                            m_colProcessedUnits.Add(strId)
                        End If
                    End While
                End Using
            End If
        Catch
            ' Ignore errors
        End Try
    End Sub

    ' Trim the buffer to the allowed memory size
    Private Sub TrimBuffer()
        While m_colProcessedUnits.Count > go_clsSystemSettings.nPenIdMemory
            m_colProcessedUnits.RemoveAt(0)
        End While
    End Sub

    ' --- Add any additional event handlers or logic as needed ---

End Class