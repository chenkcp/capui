Option Explicit On
Imports System.Runtime.InteropServices

Public Class clsUnitIdSampleApp

    ' Declare the COM/OLE object with event handling
    Private WithEvents objUnitId As clsUnitId

    Public Sub New()
        ' Initialize clsUnitId instance (equivalent to Class_Initialize)
        objUnitId = New clsUnitId()
    End Sub

    Protected Overrides Sub Finalize()
        ' Cleanup (equivalent to Class_Terminate)
        objUnitId = Nothing
        MyBase.Finalize()
    End Sub

    ' Method to write a Unit ID using the OLE interface
    Public Sub WriteUnitId()
        Dim strUnitId As String = "12345ABC"
        Dim strInputType As String = "AUTO" ' "AUTO" or "POLLED"
        Dim result As Integer

        ' Call InputUnitId method on the COM object
        result = objUnitId.InputUnitId(strUnitId, strInputType)

        ' Interpret result
        Select Case result
            Case 1
                MessageBox.Show("Unit ID successfully written.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Case 2
                MessageBox.Show("System is shutting down. Cannot write Unit ID.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Case 3
                MessageBox.Show("No open lot available. Cannot write Unit ID.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Case Else
                MessageBox.Show("Unknown error occurred.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Select
    End Sub

    ' Event handler for AsyncId event
    Private Sub objUnitId_AsyncId(ByVal strId As String) Handles objUnitId.AsyncId
        MessageBox.Show("Async ID received: " & strId, "Async Event", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    ' Event handler for LastIdState event
    Private Sub objUnitId_LastIdState(ByVal strId As String, ByVal bResult As Boolean) Handles objUnitId.LastIdState
        Dim strMessage As String = $"Last ID: {strId}{vbCrLf}Result: {(If(bResult, "Accepted", "Rejected"))}"
        MessageBox.Show(strMessage, "Last ID State Event", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

End Class
