Imports System.Windows.Forms

Public Class frmPenDelete
    Inherits Form

    ' Controls
    Private lblMessage As Label
    Private WithEvents cmdDelete As Button
    Private WithEvents cmdAbort As Button
    Private txtUnitId As TextBox
    Private fgFeatures As DataGridView

    ' State
    Private m_strFrmId As String
    Private m_iconIndex As Integer
    'Private m_pen As clsPen
    'Private m_mode As Integer

    '/*storage of current Modality
    Private m_clsWindow As clsWindowState
    Private m_frmParent As Form

    '/*Object for building a pen defect
    Private m_colDefects As colDefect

    Private m_PrimaryRow As Integer = -1

    '/*This is the pen object for operating with
    '/*and passing to the CreateUnit() operation
    Private m_oUnit As clsPen

    '/*Cause Item mapping parameters
    Private m_nCauseIndex As Integer = -1
    Private m_nCauseItem As Integer = -1
    Private m_nCauseSubItem As Integer = -1

    '/*Current icon index
    Private m_nCurrentIcon As Integer

    '/*Flag to indicate the mode of this screen
    '/*e.g. Delete=1 or Undo=2
    Private m_nViewingMode As Integer
    Private gb_InputBusyFlag As Boolean

    ' Property: Stack ID
    Public ReadOnly Property strFrmId As String
        Get
            Return m_strFrmId
        End Get
    End Property

    ' Property: IconIndex (for compatibility)
    Public Property IconIndex As Integer
        Get
            Return m_iconIndex
        End Get
        Set(value As Integer)
            m_iconIndex = value
            ' Optionally update form icon or visuals here
        End Set
    End Property

    Public Sub SetUnit(ByRef oUnitEdit As clsPen, ByVal nViewingMode As Integer)
        Dim lngItem As Long

        '/*Capture the mode that we are in so we know about releasing
        m_nViewingMode = nViewingMode
        '/*Set flag to indicate we are in an edit mode
        'm_bUnitEdit = True
        '/*Set the txtBox on the screen
        txtUnitId.Text = oUnitEdit.strPenId
        '/*set a reference to the clsPen object itself
        m_oUnit = oUnitEdit
        '/*Set a reference to the Unit's Defect collection
        m_colDefects = oUnitEdit.colPenDefects

        '/*Clear the FlexGrid; this is required for fgAddFeatures that
        '/*will be used down in FixEditItems
        mdlDefectEditor.MapInFeatureHeaders(Me, Me.fgFeatures)
        Me.fgFeatures.EditMode = DataGridViewEditMode.EditOnEnter

        '/*Loop through each item in colPenDefects
        '/*and translate the code into Description,Severity & Index's
        For lngItem = 1 To m_colDefects.Count
            FixEditItems(lngItem)
        Next lngItem
    End Sub

    Public Sub New()
        InitializeComponents()
        '' Form settings
        'Me.Text = "Delete/Undo Pen"
        'Me.FormBorderStyle = FormBorderStyle.FixedDialog
        'Me.StartPosition = FormStartPosition.CenterScreen
        'Me.MaximizeBox = False
        'Me.MinimizeBox = False
        'Me.ClientSize = New Drawing.Size(400, 160)
        'Me.ShowInTaskbar = False

        '' Label
        'lblMessage = New Label() With {
        '    .BorderStyle = BorderStyle.FixedSingle,
        '    .Text = "Delete this pen?",
        '    .Location = New Drawing.Point(24, 20),
        '    .Size = New Drawing.Size(350, 60),
        '    .TextAlign = Drawing.ContentAlignment.MiddleLeft,
        '    .AutoSize = False
        '}

        '' Delete Button
        'cmdDelete = New Button() With {
        '    .Text = "Delete",
        '    .Location = New Drawing.Point(50, 100),
        '    .Size = New Drawing.Size(120, 40),
        '    .TabIndex = 0
        '}

        '' Abort Button
        'cmdAbort = New Button() With {
        '    .Text = "Abort",
        '    .Location = New Drawing.Point(220, 100),
        '    .Size = New Drawing.Size(120, 40),
        '    .TabIndex = 1
        '}

        '' Add controls
        'Me.Controls.Add(lblMessage)
        'Me.Controls.Add(cmdDelete)
        'Me.Controls.Add(cmdAbort)
    End Sub

    Private Sub InitializeComponents()
        ' Form settings
        Me.Text = "Delete Pen"
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.StartPosition = FormStartPosition.CenterParent
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.ShowInTaskbar = False
        Me.ClientSize = New Drawing.Size(600, 400)

        ' txtUnitId
        txtUnitId = New TextBox() With {
            .Text = "100012394592A23",
            .Enabled = False,
            .TextAlign = HorizontalAlignment.Center,
            .Font = New Font("Microsoft Sans Serif", 9.0F, FontStyle.Regular, GraphicsUnit.Point),
            .Location = New Drawing.Point(50, 24),
            .Size = New Drawing.Size(500, 28),
            .TabIndex = 2
        }

        ' fgFeatures
        fgFeatures = New DataGridView() With {
            .Location = New Drawing.Point(50, 88),
            .Size = New Drawing.Size(500, 180),
            .TabIndex = 0,
            .AllowUserToAddRows = False,
            .AllowUserToDeleteRows = False,
            .ReadOnly = True,
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            .GridColor = Color.LightGray,
            .RowHeadersVisible = False,
            .ColumnCount = 6
        }
        Dim columnHeaders As String() = {"Primary", "Type", "Feature", "SubFeature", "Cause", "SubCause"}
        For i As Integer = 0 To columnHeaders.Length - 1
            fgFeatures.Columns(i).Name = columnHeaders(i)
            fgFeatures.Columns(i).Width = CInt(fgFeatures.Width / 6)
        Next

        lblMessage = New Label() With {
            .BorderStyle = BorderStyle.FixedSingle,
            .Text = "Delete this pen?",
            .Location = New Drawing.Point(50, 290),
            .Size = New Drawing.Size(500, 40),
            .TextAlign = ContentAlignment.MiddleLeft,
            .AutoSize = False,
            .TabIndex = 3
        }

        cmdDelete = New Button() With {
            .Text = "Delete Pen",
            .Location = New Drawing.Point(120, 350),
            .Size = New Drawing.Size(120, 33),
            .TabIndex = 1,
            .FlatStyle = FlatStyle.Standard
        }

        cmdAbort = New Button() With {
            .Text = "Abort",
            .Location = New Drawing.Point(360, 350),
            .Size = New Drawing.Size(120, 33),
            .TabIndex = 4,
            .FlatStyle = FlatStyle.Standard
        }

        Me.Controls.AddRange({txtUnitId, fgFeatures, lblMessage, cmdDelete, cmdAbort})
    End Sub

    Private Sub form_load(sender As Object, e As EventArgs) Handles MyBase.Load
        DisableCloseX(Me)
        WindowInit()
        If Not String.IsNullOrEmpty(m_strFrmId) Then
            mdlWindow.RemoveForm(m_strFrmId)
        End If
        'm_strFrmId = mdlWindow.AddForm(Me)
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

    ' Delete button click: perform delete/undo
    Private Sub cmdDelete_Click(sender As Object, e As EventArgs) Handles cmdDelete.Click
        '/*Disable the buttons on the screen
        cmdAbort.Enabled = False
        cmdDelete.Enabled = False

        '/*Make sure we have a valid unit
        If m_oUnit IsNot Nothing Then
            '/*Relaese the pen if this is a delete
            mdlCreatePen.ReleaseUnit(m_oUnit)

            '/*Attempt to delete the current Unit
            If mdlCreatePen.DeleteUnit(m_oUnit) Then
                '/*Succesful, so hide the window and show the result
                HideWindow()

                '/*Prime the message
                frmMessage.GenerateMessage("Pen REMOVED from database", "Results")
            End If
        End If

        '/*Enable the buttons on the screen
        cmdAbort.Enabled = True
        cmdDelete.Enabled = True

        '/*Destory the unit
        m_oUnit = Nothing
    End Sub

    ' Prevent closing from the X button
    Protected Overrides Sub OnLoad(e As EventArgs)
        MyBase.OnLoad(e)
        mdlTools.DisableCloseX(Me)
    End Sub

    ' Abort button click: hide and release lock
    Private Sub cmdAbort_Click(sender As Object, e As EventArgs) Handles cmdAbort.Click
        '/*Disable the buttons on the screen
        cmdAbort.Enabled = False
        cmdDelete.Enabled = False

        If m_oUnit IsNot Nothing Then
            '/*Relaese the pen if this is a delete
            If m_nViewingMode = 1 Then
                If Not mdlCreatePen.ReleaseUnit(m_oUnit) Then
                    '/*See if the user would like to try to release again
                End If
            End If

            '/*If we are aborting on an Undo we need to correct
            '/*the modification to count
            If m_nViewingMode = 2 Then
                mdlCreatePen.AddSampleCount(m_oUnit)
            End If
        End If

        '/*Close this window
        HideWindow()

        '/*Enable the buttons on the screen
        cmdAbort.Enabled = True
        cmdDelete.Enabled = True

        '/*Destroy the unit
        m_oUnit = Nothing

        '/*Release the lock on the client
        gb_InputBusyFlag = False
    End Sub

    '=======================================================
    'Routine: frmPenDelete.FixEditItems(lng)
    'Purpose: This takes a passed index and attempts to
    'translate the Unit's Features into Index coordinates.
    'This also attaches the severity,descriptions and calls
    'the Flexgrid mapping function in the process.
    '
    'Globals:None
    '
    'Input: lngItem - The item index of the Defect to translate.
    '
    'Return:None
    '
    'Modifications:
    '   12-16-1998 As written for Pass1.5
    '
    '
    '=======================================================
    Private Sub FixEditItems(ByRef lngItem As Long)
        Dim nClassIndex As Integer
        Dim nCauseIndex As Integer
        Dim nCauseItem As Integer
        Dim nCauseSubItem As Integer

        '/*Set a reference to the colDefects being operated on
        With m_colDefects(lngItem)
            '/*Get the index of the Class if available
            nClassIndex = ClassCodeToIndex(.strFeatureClassDesc)

            '/*Make sure an Index was returned
            If nClassIndex >= 0 Then
                '/*Map in the known items
                .lngSeverity = gcol_FeatureClasses(nClassIndex).lngSeverity
                .nFeatureClass = nClassIndex

                '/*Try to match up the Features
                SeekLevelCodes(m_colDefects(lngItem), nClassIndex)

                '/*Try to match up any Cause items
                If .strCauseCd.Length <> 0 Then
                    nCauseIndex = CauseCodeToIndex()

                    '/*If the function reurned a valid Index
                    If nCauseIndex >= 0 Then
                        SeekCauseCodes(m_colDefects(lngItem), nCauseIndex, nCauseItem, nCauseSubItem)

                        If nCauseItem >= 0 Then
                            '/*Set the Level-1 cause code
                            m_nCauseIndex = nCauseIndex
                            m_nCauseItem = nCauseItem

                            '/*if the Level-2 code is not zero set it
                            If nCauseSubItem >= 0 Then
                                m_nCauseSubItem = nCauseSubItem
                            End If
                        End If
                    End If
                End If
            Else
                .nFeatureClass = -1
            End If

            '/*Check for any erroneuos returns from the SeekCodes()
            If .nFeatureIndex < 0 Or (.nSubFeatureIndex < 0 And
           m_colDefects(lngItem).strSubFeatureCd.Length > 0) Then
                frmMessage.GenerateMessage("Error encountered translating codes to descriptions, can not display pen correctly.")
            Else
                '/*Map the unit to the FlexGrid
                '/*Call the Grid mapping function
                fgFeatureAddItem(nClassIndex, .nFeatureIndex, .nSubFeatureIndex, .bPrimary)
            End If
        End With
    End Sub

    '=======================================================
    'Routine: frmPenDelete.fgFeatureAddItem(n,n,n,b)
    'Purpose: Add a Feature row to the FlexGrid at the
    '         bottom of the screen.
    '
    'Globals:None
    '
    'Input: nGroup - integer pointer to the index of the top
    '                level Feature Structure class
    '       nFeature - pointer to the index of the Feature within
    '                  the class.
    '       nSubFeature - pointer to a sub feature if there is
    '                     one. 0 will indicate there is not
    '                     since it is a base 1 collection.
    '       bPrimary - Mark the primary column "Yes"
    '
    'Return:None
    '
    '
    'Modifications:
    '   10-07-1998 As written for Pass1.1
    '
    '   12-16-1998 Added the CauseIndex resets
    '   and the Comment & NumericInput resets to this
    '   routine. This is due to Edit Pens bypassing
    '   the AddUnit routine.
    '
    '   12-17-1998 Copied from frmDefectEditor for
    '   use in this form.
    '=======================================================
    Private Sub fgFeatureAddItem(ByRef nGroup As Integer, ByRef nFeature As Integer, ByRef nSubFeature As Integer, ByRef bPrimary As Boolean)
        Dim gridRowData As New List(Of String)

        '/*Build up a Tab delimeted string
        '/*Set the Primary flag
        If bPrimary Then
            gridRowData.Add("Yes")
            '/*If there is a primary marking
            If m_PrimaryRow <> -1 AndAlso fgFeatures.Rows.Count > m_PrimaryRow Then
                '/*Set the position
                '/*Remove primary mark from the current row.
                fgFeatures.Rows(m_PrimaryRow).Cells(0).Value = ""
            End If
            '/*Set the new primary row
            m_PrimaryRow = fgFeatures.Rows.Count
        Else
            gridRowData.Add("")
        End If

        '/*Set reference to the group
        With gcol_FeatureClasses.Item(nGroup)
            '/*Add on the Type
            gridRowData.Add(.strTitle)
            '/*Add on the Feature to the output string
            gridRowData.Add(.colClassFeatures.Item(nFeature).strDesc)
            '/*Switch on whether the Feature has a SubFeature
            If nSubFeature >= 0 Then
                gridRowData.Add(.colClassFeatures.Item(nFeature).colSub.Item(nSubFeature).strDesc)
            Else
                gridRowData.Add("")
            End If
        End With

        '/*Tack on any Cause items if the Cause Index Member has been used
        If m_nCauseIndex >= 0 Then
            '/*Get a reference to the object
            With gcol_FeatureClasses.Item(m_nCauseIndex).colClassFeatures.Item(m_nCauseItem)
                '/*Add teh description to the string
                gridRowData.Add(.strDesc)
                '/*get any subitems used by this Cause
                If m_nCauseSubItem >= 0 Then
                    '/*Append on the SubItem
                    gridRowData.Add(.colSub.Item(m_nCauseSubItem).strDesc)
                Else
                    gridRowData.Add("")
                End If
            End With
            '/*Insure that the items are now reset
            m_nCauseIndex = -1
            m_nCauseItem = -1
            m_nCauseSubItem = -1
        Else
            gridRowData.Add("")
            gridRowData.Add("")
        End If

        With fgFeatures
            '/*Add the item to the list
            Dim newRowIndex As Integer = .Rows.Add(gridRowData.ToArray())
            '/*Need to set a toggle that supresses this when the rows
            '/*do not extend below the bottom of the Grid window
            '/*This is set to 6 rows since we never resize the
            '/*height of this control.
            If .Rows.Count > 6 Then
                '/*Set the grid focus to the Item to reassure the user
                '/*that things are funcitoning as expected.
                .FirstDisplayedScrollingRowIndex = newRowIndex
                .CurrentCell = .Rows(newRowIndex).Cells(0)
            End If
        End With
    End Sub
End Class
