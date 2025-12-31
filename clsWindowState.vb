Public Class clsWindowState
    'local variable(s) to hold property value(s)
    Private mvarcolWindow As New List(Of clsWindow) 'local copy
    Private mvarfrmIn As Form 'The form that this was called by
    Private mvarbEnabled As Boolean
    Private mvarbVisible As Boolean
    Private mvarbFloat As Boolean
    Private mvarbModal As Boolean

    '
    '=======================================================
    'Routine: clsWindowState.HIdeWindow(Form)
    'Purpose: This hides a Form and sets all the window
    '         states back to where they were previous
    '         to this window showing.
    '
    'Globals:None
    '
    'Input: frmIn - A reference to the form being closed.
    '
    'Return:None
    '
    '
    'Modifications:
    '   10-15-1998 As written for Pass1.1
    '
    '   03-29-1999 Removed the window visible property
    '   from the windows collection to keep from
    '   re-showing dropped windows.
    '=======================================================
    Public Sub HideWindow()
        For Each winItem In mvarcolWindow
            'Reset each of the windows to its previous state
            winItem.frmIn.Enabled = winItem.bEnabled
            'winItem.frmIn.Visible = winItem.bVisible
            'If the window was floating for this state
            If winItem.bFloat Then
                'Refloat the window
                'Call mdlTools.FloatWindow(winItem.frmIn)
            End If
        Next winItem

        'Set the window properties to the way they
        'were when ShowWIndow() was called
        'mvarfrmIn.Visible = mvarbVisible
        mvarfrmIn.Enabled = mvarbEnabled
        If mvarbFloat Then
            'mdlTools.FloatWindow mvarfrmIn
        End If
    End Sub

    '
    '=======================================================
    'Routine: ShowWindow(Form)
    'Purpose: This sets the current state of the windows so
    '         that the requested window is the only one
    '         that the user can access.
    '
    'Globals:None
    '
    'Input: frmIn - The form to be loaded/shown.
    '
    'Return:None
    '
    '
    'Modifications:
    '   10-15-1998 As written for Pass1.1
    '
    '   02-09-1999 Added section at end for requests
    '   that are not Modal so that the window will be seen.
    '
    '   03-30-1999 Added DoEvents on either side of the
    '   FormThis.Enabled=False becuase this was not
    '   consistantly occuring. Seems to be related to the
    '   windows loading and losing the message to disable.
    '
    '   07-07-1999 [Bug] There is an error 401 occuring
    '   some where in the app. It appears at the end
    '   of a Lot (when the SamplesToSuspend is reached).
    '   The only documentation that I could find indicates
    '   that this can occur from in-process components
    '   trying to show Modeless forms. Hopefully these
    '   fixes (error handler and test for App.NonModalAllowed
    '   ) will cure the problem.
    '=======================================================
    Public Sub ShowWindow(ByRef frmIn As Form, Optional ByVal bModal As Boolean = False)
        Dim frmThis As Form
        Const cstrBlank As String = ""
        Dim strMsg As String

        Try
            'Get the state of the window to be shown
            mvarfrmIn = frmIn
            mvarbEnabled = frmIn.Enabled
            mvarbFloat = CBool(frmIn.Tag)
            mvarbVisible = frmIn.Visible

            'Do an inventory of the active window states
            'Loop through the Collection of Forms
            For Each frmThis In Application.OpenForms
                'Make sure this isn't are Modal window
                If frmThis.Name <> frmIn.Name Then
                    'Add a record of the window to are private collection
                    mvarcolWindow.Add(New clsWindow With {.frmIn = frmThis, .bEnabled = frmThis.Enabled, .bVisible = frmThis.Visible, .bFloat = CBool(frmThis.Tag)})
                    'If the form is visible remove its float state
                    'frmThis.Tag = cstrBlank
                    'If it is visible and the request is Modal then disable it
                    If bModal Then
                        'Suppress hiding the Test Routine Form
                        If frmThis.Name <> "frmTestRoutines" Then
                            'Wrapper to gaurentee that the form is disabled
                            Application.DoEvents()
                            If frmThis.Visible Then frmThis.Enabled = False
                            Application.DoEvents()
                        End If
                    End If
                End If
            Next frmThis

            'If this is a request for modal set it
            If bModal Then
                'Float the window to the Top of the Desktop
                'Call mdlTools.FloatWindow(frmIn)
                'If it is are modal window make sure that it is enabled.
                frmIn.Enabled = True
                'now show the form
                frmIn.ShowDialog()
                mvarbModal = bModal
                'Just show the window and attempt to set its focus
            Else
                'Trigger the windows standard Show method
                frmIn.Show()
                'Try to set its Focus
                If frmIn.Enabled And frmIn.Visible Then frmIn.Focus()
            End If
        Catch ex As Exception
            'Get the error message
            strMsg = "ShowWindow() Error:(" & ex.HResult & ")" & ex.Message
            'Ge the form that made this call
            If frmIn Is Nothing Then
                strMsg = strMsg & Environment.NewLine & "Form In:Nothing"
            Else
                strMsg = strMsg & Environment.NewLine & "Form In:" & frmIn.Name
            End If
            'This should show if we are cycling through the form collection
            If frmThis Is Nothing Then
                strMsg = strMsg & Environment.NewLine & "Current Form:Nothing"
            Else
                strMsg = strMsg & Environment.NewLine & "Current Form:" & frmThis.Name
            End If
            'Show the message
            'frmMessage.GenerateMessage(strMsg, "Fatal Error", True)
            MessageBox.Show("Fatal Error")
            Err.Clear()
        End Try
    End Sub
End Class
