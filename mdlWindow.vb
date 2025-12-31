Imports System.Collections.Generic

'/---------------------------------------------------------
'/*In old vb6 code,  m_colForms type is Collection
'/*But in vb.net , there is no Collection, we will change m_colForms to List(Of CustomForm)(),
'/*Here creating a new class CustomForm and add Property strFrmId , because function RemoveForm() will
'/*use m_colForms(nItem).strFrmId
'/---------------------------------------------------------
Public Class CustomForm
    Inherits Form
    Public Property strFrmId As String
End Class

Module mdlWindow
    '/*stack of Forms
    Private m_colForms As New List(Of CustomForm)()
    '/*Name of the error form
    Private Const cstrErrorForm As String = "frmMessage"

    '/---------------------------------------------------------
    '/*Add a new for to the stack
    '/*
    '/*Modifications:
    '/* 09-27-1999 Adding in check to insure that
    '/* the Error form stays on top when it is in the
    '/* stack. C Barker
    '/---------------------------------------------------------
    Public Function AddForm(ByVal frmIn As Form) As String
        '/*Make sure the error form is on top
        If frmIn.Name <> cstrErrorForm AndAlso GetTopFormName() = cstrErrorForm AndAlso m_colForms.Count > 0 Then
            '/*Insert the form before the current top window
            m_colForms.Insert(m_colForms.Count - 1, frmIn)
            '/*Dsiplay the new form
            ShowSetFocus(frmIn)
            '/*Now disable the new form
            frmIn.Enabled = False
            '/*Redisplay the error message form
            'ShowSetFocus(frmMessage)
            MessageBox.Show("Redispaly the error message form")
        Else
            '/*Disable the current form
            If m_colForms.Count > 0 Then
                m_colForms(m_colForms.Count - 1).Enabled = False
            End If
            '/*Display the form
            m_colForms.Add(frmIn)
            ShowSetFocus(frmIn)
        End If
        Return mdlCreatePen.GetIdDttm()
    End Function

    '/*Remove the top most form
    Public Sub RemoveTopForm()
        If m_colForms.Count > 0 Then
            Dim nIndex As Integer = m_colForms.Count - 1
            m_colForms(nIndex).Hide()
            m_colForms.RemoveAt(nIndex)
            EnableTopForm()
        End If
    End Sub

    '/*Remove a form in the middle of the collection
    Public Sub RemoveForm(ByVal strFrmId As String)
        Dim nIndex As Integer = -1
        For nItem As Integer = m_colForms.Count - 1 To 0 Step -1
            If m_colForms(nItem).strFrmId = strFrmId Then
                nIndex = nItem
                Exit For
            End If
        Next nItem
        If nIndex > -1 Then
            m_colForms(nIndex).Hide()
            m_colForms.RemoveAt(nIndex)
            If nIndex - 1 = m_colForms.Count - 1 Then
                EnableTopForm()
            End If
        End If
    End Sub

    '/*Return reference to the top form
    Public Function GetTopForm() As Form
        If m_colForms.Count > 0 Then
            Return m_colForms(m_colForms.Count - 1)
        End If
        Return Nothing
    End Function

    '/*Return the name of hte top form if there is one
    Public Function GetTopFormName() As String
        If m_colForms.Count > 0 Then
            Return m_colForms(m_colForms.Count - 1).Name
        End If
        Return ""
    End Function

    '/*Enable the top form
    Private Sub EnableTopForm()
        Try
            If m_colForms.Count > 0 Then
                m_colForms(m_colForms.Count - 1).Enabled = True
                m_colForms(m_colForms.Count - 1).Focus()
            End If
        Catch ex As Exception
            ' 处理异常
        End Try
    End Sub

    '
    '========================================================
    'Routine: ShowSetFocus(frm)
    'Purpose: This is an isolated routine to show a form
    'and set its focus if possible.
    '
    'Globals:None
    '
    'Input: frmIn - the form to show and attempt to
    '       setfocus on
    '
    'Return:None
    '
    'Tested:
    '
    'Modifications:
    '   08-10-1999 As written for Beta1 Phase 3.2
    '
    '
    '=======================================================
    Public Sub ShowSetFocus(ByVal frmIn As Form)
        Try
            With frmIn
                .Enabled = True
                If Not .Visible Then
                    .Show()
                    .Visible = True
                End If
                '/*Now set our focus
                If .Visible AndAlso .Enabled Then
                    .Focus()
                End If
            End With
        Catch ex As Exception
            ' 处理异常
        End Try
    End Sub

    '
    '========================================================
    'Routine: CheckWhosTop(frm)
    'Purpose: This is called from forms that show in the
    'taskbar to see if they are supposed to be on top when
    'their paint method is called. This situation occurrs
    'when a TopWindow locks the lower forms, but a user clicks
    'on the icon in the status bar for a lower form. It also
    'occurrs when a user alt+tab's through the windows. The paint
    'method is guranteed to be triggered when the wrong window
    'moves to the top of the painted screen layers.
    '
    'Globals:None
    '
    'Input: frmIn - the form to check focus on.
    '
    'Return:None
    '
    'Tested:
    '
    'Modifications:
    '   08-15-2001 As written v1.1.3 Chris Barker
    '
    '
    '=======================================================
    Public Sub CheckWhosTop(ByVal frmIn As Form)
        Try
            If GetTopFormName() <> frmIn.Name AndAlso Not frmIn.Enabled Then
                Dim topForm As Form = GetTopForm()
                If topForm IsNot Nothing AndAlso topForm.Enabled Then
                    topForm.Focus()
                End If
            End If
        Catch ex As Exception
            mdlcommon.LogEvent("mdlWindow.CheckWhosTop: Form=" & frmIn.Name & " Error=" & ex.HResult.ToString() & " " & ex.Message)
        End Try
    End Sub
End Module
