Imports System.Drawing
Imports System.Windows.Documents
Imports System.Windows.Forms
Imports System.Windows.Ink

Public Class frmDefectEditor
    Inherits Form

    ' Declare Controls


    Private grpDisposition As GroupBox
    Private WithEvents optGood As RadioButton
    Private WithEvents optReclaim As RadioButton
    Private WithEvents optScrap As RadioButton
    Private lblCosmetic As Label
    Private lblFunctional As Label
    Private WithEvents lstCosmetic As ListBox
    Private WithEvents lstFunctional As ListBox

    Private WithEvents dgvFeatures As DataGridView

    ' Declare Controls
    Private lblLotId As Label
    Private lblUnitId As Label

    Private WithEvents txtPenCount As TextBox
    Private txtLotId As TextBox
    Private txtNumericInput As New TextBox()
    Private txtComment As New TextBox()
    Private txtUnitId As TextBox
    Private cmbTestBedType As ComboBox
    Private cmbTestBedNumber As ComboBox
    Private cmbComponent As ComboBox
    Private WithEvents cmdOK As Button
    Private WithEvents cmdCancel As Button
    Private WithEvents cmdHelp As Button
    Private WithEvents cmdPrimary As Button
    Private WithEvents cmdDelete As Button
    Private WithEvents cmdComment As Button
    Private WithEvents cmdNumericInput As Button
    Private WithEvents chkContLogging As CheckBox
    Private WithEvents chkNotShipped As CheckBox
    Friend WithEvents lstItemExtended As ListBox
    Friend WithEvents lstItems As List(Of ListBox)
    Friend WithEvents lstHotItems As List(Of ListBox)
    Private lblPenCount As Label
    Friend lblList As List(Of Label)
    Private WithEvents fgFeatures As DataGridView

    ' Declare other variables
    Private m_State As Integer
    'Private m_clsWindow As clsWindowState
    Private m_frmParent As Form
    Private m_nLastIndex As Integer
    Private m_nURLIndex As Integer = -1
    Private m_nURLItem As Integer = -1
    Private m_nListIndex As Integer = -1
    Private m_nListItem As Integer = -1
    Private m_nListSubItem As Integer = -1
    Private m_nCauseIndex As Integer = -1
    Private m_nCauseItem As Integer = -1
    Private m_nCauseSubItem As Integer = -1
    Private m_PrimaryRow As Integer
    Private m_PrimaryWeight As Long
    Private m_colDefects As New colDefect()
    Private m_PrimaryIndex As Integer
    Private m_bUnitEdit As Boolean
    Private m_bUnitEditRemote As Boolean
    Private m_oUnit As clsPen
    'Private WithEvents m_oSecurity As Security
    Private m_strFrmId As String
    Private m_strDisposition As String

    ' Constants
    Private Const cnSTATE_READY As Integer = 0
    Private Const cnSTATE_LIST As Integer = 1
    Private Const cnSTATE_SUBLIST As Integer = 2
    Private Const cnSTATE_SUBLISTDONE As Integer = 6
    Private Const cnSTATE_DETAIL As Integer = 3
    Private Const cnSTATE_COMMENT As Integer = 4
    Private Const cnSTATE_NUMERIC As Integer = 5
    Private Const clngMinHeight As Long = 6765
    Private Const clngMinWidth As Long = 8785
    Private Const cstrClassExclude As String = "CAUSE"


    Public Sub New()
        Me.Text = "Defect Editor"
        Me.Size = New Size(880, 700)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        InitializeControls()
    End Sub


    Public Sub SetUnit(ByRef oUnitEdit As clsPen, Optional ByRef bRemote As Boolean = False)
        ' Set edit mode flags
        If bRemote Then
            m_bUnitEditRemote = True
            m_bUnitEdit = False
        Else
            m_bUnitEditRemote = False
            m_bUnitEdit = True
        End If

        ' Set unit fields
        txtUnitId.Text = oUnitEdit.strPenId
        txtLotId.Text = oUnitEdit.strGroupId

        ' Save reference to pen and its defect collection
        m_oUnit = oUnitEdit
        m_colDefects = oUnitEdit.colPenDefects

        ' Reset tracking values
        m_PrimaryWeight = -1
        m_PrimaryRow = 0
        m_PrimaryIndex = 0

        ' Setup grid headers
        mdlDefectEditor.MapInFeatureHeaders(Me, Me.fgFeatures)
        Me.fgFeatures.EditMode = DataGridViewEditMode.EditOnEnter

        ' Populate grid with existing defects
        For i As Integer = 1 To m_colDefects.Count
            FixEditItems(i)
        Next

        ' Set additional unit display properties
        SetEditProperties(oUnitEdit)

        ' Set pen count if feature is enabled
        'If go_clsSystemSettings.bBadPenCountEnabled Then
        '    txtPenCount.Text = oUnitEdit.nCount.ToString()
        'End If
    End Sub
    Private Sub FixEditItems(ByRef lngItem As Long)
        Dim nClassIndex As Integer
        Dim nCauseIndex As Integer
        Dim nCauseItem As Integer
        Dim nCauseSubItem As Integer

        ' Reference to the current defect item
        With m_colDefects(lngItem)
            ' Get class index from description
            nClassIndex = ClassCodeToIndex(.strFeatureClassDesc)

            If nClassIndex >= 0 Then
                ' Map severity and feature class
                .lngSeverity = gcol_FeatureClasses(nClassIndex).lngSeverity
                .nFeatureClass = nClassIndex

                ' Match feature levels
                SeekLevelCodes(m_colDefects(lngItem), nClassIndex)

                ' Try matching cause codes
                If Not String.IsNullOrEmpty(.strCauseCd) Then
                    nCauseIndex = CauseCodeToIndex()

                    If nCauseIndex >= 0 Then
                        SeekCauseCodes(m_colDefects(lngItem), nCauseIndex, nCauseItem, nCauseSubItem)

                        If nCauseItem >= 0 Then
                            m_nCauseIndex = nCauseIndex
                            m_nCauseItem = nCauseItem

                            If nCauseSubItem >= 0 Then
                                m_nCauseSubItem = nCauseSubItem
                            End If
                        End If
                    End If
                End If
            End If

            ' Validate mapping results
            If .nFeatureIndex < 0 OrElse
           (.nSubFeatureIndex < 0 AndAlso Not String.IsNullOrEmpty(.strSubFeatureCd)) Then
                'frmMessage.GenerateMessage("Error encountered translating codes to descriptions, cannot display pen correctly.")
                MessageBox.Show("Error encountered translating codes to descriptions, cannot display pen correctly.")
            Else
                ' Map item to the FlexGrid
                fgFeatureAddItem(nClassIndex, .nFeatureIndex, .nSubFeatureIndex, .bPrimary)
            End If

            ' Set comment and numeric input
            'If Not String.IsNullOrEmpty(.strComment) Then SetComment(.strComment)
            'SetNumeric(.strNumericInput)

            ' Set primary indicators if this is the primary item
            If .bPrimary Then
                m_PrimaryWeight = .lngSeverity
                m_PrimaryIndex = lngItem
            End If
        End With
    End Sub

    Private Sub SetEditProperties(ByRef oEditUnit As clsPen)
        Try
            Dim nIndex As Integer
            Dim strTestBedNumber As String = ""
            Dim strDisposition As String = ""

            ' If ComboBox has more than 1 item, try to match test bed
            'If cmbTestBedNumber.Items.Count > 1 Then
            '    strTestBedNumber = oEditUnit.strTestBed

            '    ' Extract value before "-"
            '    If Not String.IsNullOrEmpty(strTestBedNumber) AndAlso strTestBedNumber.Contains("-") Then
            '        strTestBedNumber = strTestBedNumber.Substring(0, strTestBedNumber.IndexOf("-"))
            '    End If

            '    ' Look for matching item in ComboBox
            '    For nIndex = 0 To cmbTestBedNumber.Items.Count - 1
            '        If cmbTestBedNumber.Items(nIndex).ToString() = strTestBedNumber Then
            '            cmbTestBedNumber.SelectedIndex = nIndex
            '            Exit For
            '        End If
            '    Next
            'End If

            ' Set Disposition OptionButtons
            strDisposition = oEditUnit.strDisposition
            optGood.Checked = (strDisposition = "G")
            optReclaim.Checked = (strDisposition = "R")
            optScrap.Checked = (strDisposition = "S")

            ' Set "Not Shipped" CheckBox
            chkNotShipped.Checked = oEditUnit.bPenNotShipped

        Catch ex As Exception
            ' Optionally log or handle the error
            ' MessageBox.Show("Error in SetEditProperties: " & ex.Message)
        End Try
    End Sub

    Private Sub InitializeControls()
        ' Initialize controls and set properties
        lblLotId = New Label() With {.Text = "Lot ID", .Location = New Point(20, 20), .AutoSize = True}
        txtLotId = New TextBox() With {.Location = New Point(70, 18), .Size = New Size(180, 24)}

        lblUnitId = New Label() With {.Text = "Pen ID", .Location = New Point(270, 20), .AutoSize = True}
        txtUnitId = New TextBox() With {.Location = New Point(330, 18), .Size = New Size(180, 24)}

        grpDisposition = New GroupBox() With {.Text = "Pen Disposition", .Location = New Point(530, 10), .Size = New Size(235, 60)}
        optGood = New RadioButton() With {.Text = "Good", .Location = New Point(10, 20), .AutoSize = True}
        optReclaim = New RadioButton() With {.Text = "Reclaim", .Location = New Point(80, 20), .AutoSize = True}
        optScrap = New RadioButton() With {.Text = "Scrap", .Location = New Point(165, 20), .AutoSize = True}
        grpDisposition.Controls.AddRange({optGood, optReclaim, optScrap})

        chkNotShipped = New CheckBox() With {.Text = "Not Shipped", .Location = New Point(20, 60), .AutoSize = True}

        ' Section Labels and ListBoxes
        lblList = New List(Of Label)()
        lstItems = New List(Of ListBox)()
        lstHotItems = New List(Of ListBox)()
        Dim sectionTitles() As String = {"Cosmetic", "Risk", "Functional"}
        Dim xPositions() As Integer = {24, 304, 584}
        Dim yLabel As Integer = 100
        Dim yList As Integer = 120
        Dim listBoxHeight As Integer = 100
        Dim hotListBoxHeight As Integer = 65

        Dim bottomOfListArea As Integer = 0

        For i As Integer = 0 To 2
            Dim lbl As New Label With {.Text = sectionTitles(i), .Location = New Point(xPositions(i), yLabel), .AutoSize = True}
            lblList.Add(lbl)
            Me.Controls.Add(lbl)

            Dim lst As New ListBox With {.Location = New Point(xPositions(i), yList), .Size = New Size(250, listBoxHeight)}
            lstItems.Add(lst)
            Me.Controls.Add(lst)
            AddHandler lst.Click, AddressOf lstItems_Click
            AddHandler lst.DoubleClick, AddressOf lstItems_DoubleClick

            Dim lstHotY As Integer = yList + listBoxHeight + 5 ' position just below lstItems
            Dim lstHot As New ListBox With {.Location = New Point(xPositions(i), lstHotY), .Size = New Size(250, hotListBoxHeight), .Visible = True}
            lstHotItems.Add(lstHot)
            Me.Controls.Add(lstHot)
            AddHandler lstHot.Click, AddressOf lstHotItems_Click
            AddHandler lstHot.DoubleClick, AddressOf lstHotItems_DoubleClick

            ' Track the bottom Y of the last control in this section
            bottomOfListArea = Math.Max(bottomOfListArea, lstHotY + hotListBoxHeight)
        Next

        Dim yAfterLists As Integer = bottomOfListArea + 20 ' Add spacing after last list

        cmdOK = New Button() With {.Text = "OK", .Location = New Point(20, yAfterLists), .Size = New Size(100, 30)}
        AddHandler cmdOK.Click, AddressOf cmdOK_Click

        cmdCancel = New Button() With {.Text = "Cancel", .Location = New Point(130, yAfterLists), .Size = New Size(100, 30)}
        AddHandler cmdCancel.Click, AddressOf cmdCancel_Click

        cmdHelp = New Button() With {.Text = "Help", .Location = New Point(240, yAfterLists), .Size = New Size(100, 30)}

        chkContLogging = New CheckBox() With {.Text = "Enable Continuous Logging", .Location = New Point(20, yAfterLists + 40), .AutoSize = True}

        fgFeatures = New DataGridView() With {.Location = New Point(20, yAfterLists + 70), .Size = New Size(800, 200), .ColumnCount = 6}
        fgFeatures.Columns(0).Name = "Primary"
        fgFeatures.Columns(1).Name = "Type"
        fgFeatures.Columns(2).Name = "Defect"
        fgFeatures.Columns(3).Name = "SubDefect"
        fgFeatures.Columns(4).Name = "Comment"
        fgFeatures.Columns(5).Name = "Numeric"
        AddHandler fgFeatures.SelectionChanged, AddressOf fgFeatures_SelectionChanged
        fgFeatures.Columns(0).ReadOnly = True
        fgFeatures.Columns(1).ReadOnly = True
        fgFeatures.Columns(2).ReadOnly = True
        fgFeatures.Columns(3).ReadOnly = True
        fgFeatures.Columns(4).ReadOnly = False 'Editable column Comment
        fgFeatures.Columns(5).ReadOnly = False 'Editable column Numeric

        fgFeatures.EditMode = DataGridViewEditMode.EditOnEnter ' choose cell will enable edit mode

        ' Add a new ListBox named lstItemExtended width 250+584
        lstItemExtended = New ListBox() With {
            .Size = New Size((250 * 3) + 60, listBoxHeight + 5 + hotListBoxHeight), ' Default size, will be adjusted dynamically
            .Location = New Point(24, yList),
            .Visible = False
        }
        AddHandler lstItemExtended.DoubleClick, AddressOf lstItemExtended_DoubleClick
        AddHandler lstItemExtended.Click, AddressOf lstItemExtended_Click
        ' Add lstItemExtended to the form
        Me.Controls.Add(lstItemExtended)

        cmdPrimary = New Button() With {.Text = "Primary", .Location = New Point(100, yAfterLists + 280), .Size = New Size(100, 30)}
        cmdDelete = New Button() With {.Text = "Delete", .Location = New Point(210, yAfterLists + 280), .Size = New Size(100, 30)}
        cmdComment = New Button() With {.Text = "Comment", .Location = New Point(320, yAfterLists + 280), .Size = New Size(100, 30)}
        cmdNumericInput = New Button() With {.Text = "Numeric Input", .Location = New Point(430, yAfterLists + 280), .Size = New Size(120, 30)}


        ' Add Controls
        Me.Controls.AddRange({lblLotId, txtLotId, lblUnitId, txtUnitId, chkNotShipped, grpDisposition,
                              cmdOK, cmdCancel, cmdHelp, chkContLogging,
                              fgFeatures, cmdPrimary, cmdDelete, cmdComment, cmdNumericInput})




    End Sub

    Private Sub lstHotItems_Click(sender As Object, e As EventArgs)
        '/*Feature structure indexes
        Dim nListIndex As Integer
        Dim nListItem As Integer
        Dim nSubListItem As Integer
        Dim lstHotItem As ListBox = CType(sender, ListBox)
        Dim Index As Integer = lstHotItems.IndexOf(lstHotItem)

        If lstHotItems(Index).SelectedIndex < 0 Then
            Exit Sub
        End If

        '/*Request to set the current state
        If ProcessState(cnSTATE_LIST) Then
            '/*Increment for base 0 -> base 1 offset
            'nListIndex = Index + 1
            'nListItem = lstHotItems(Index).SelectedIndex + 1
            nListIndex = Index
            nListItem = lstHotItems(Index).SelectedIndex
        End If
    End Sub

    Private Sub lstHotItems_DoubleClick(sender As Object, e As EventArgs)
        '/*Feature structure indexes
        Dim nListIndex As Integer
        Dim nListItem As Integer
        Dim nSubListItem As Integer
        Dim lstHotItem As ListBox = CType(sender, ListBox)
        Dim Index As Integer = lstHotItems.IndexOf(lstHotItem)

        If lstHotItems(Index).SelectedIndex < 0 Then
            Exit Sub
        End If

        If ProcessState(cnSTATE_LIST) Then
            '/*Increment for base 0 -> base 1 offset
            'nListIndex = Index + 1
            'nListItem = lstHotItems(Index).SelectedIndex + 1
            nListIndex = Index
            nListItem = lstHotItems(Index).SelectedIndex

            'If nListItem > 0 Then
            If nListItem >= 0 Then
                '/*Filter out any excluded classes
                If Not (mdlDefectEditor.IsExcludedClass(nListIndex)) Then
                    '/*Translate the Hot Item indexes and return byref
                    mdlDefectEditor.LookUpHotItemCoord(nListIndex, nListItem, nSubListItem)
                    '/*Add the item
                    AddFeature(nListIndex, nListItem, nSubListItem)
                    '/*Unselect the user's item
                    'lstHotItems(Index).SelectedIndex = -1
                    lstHotItem.ClearSelected()
                    '/*Add on a comment
                Else
                    '/*Translate the Hot Item indexes and return byref
                    mdlDefectEditor.LookUpHotItemCoord(nListIndex, nListItem, nSubListItem)
                    '/*Add the item
                    AddCause(nListIndex, nListItem, nSubListItem)
                    '/*Unselect the user's item
                    'lstHotItems(Index).SelectedIndex = -1
                    lstHotItem.ClearSelected()
                End If
            End If
        End If
    End Sub

    Private Sub lstItemExtended_Click()
        If lstItemExtended.SelectedIndex < 0 Then
            Exit Sub
        End If
        '/*This is for use in the Help button
        'm_nListSubItem = lstItemExtended.SelectedIndex + 1
        m_nListSubItem = lstItemExtended.SelectedIndex
    End Sub

    Private Sub lstItemExtended_DoubleClick()
        Dim nIndex As Integer

        If lstItemExtended.SelectedIndex < 0 Then
            Exit Sub
        End If

        '/*Generate the ListIndex offset
        'nIndex = lstItemExtended.SelectedIndex + 1
        nIndex = lstItemExtended.SelectedIndex
        '/*We need to abort any processing if the user has
        '/*not actually selected a List item
        If nIndex >= 0 Then
            '/*Filter out any excluded classes
            If mdlDefectEditor.IsExcludedClass(m_nListIndex) Then
                '/*Set the private member
                'm_nCauseSubItem = nIndex
                '/*Add the item to the Grid
                AddCause(m_nListIndex, m_nListItem, nIndex)
            Else
                '/*Add the item to the Grid
                AddFeature(m_nListIndex, m_nListItem, nIndex)
            End If
            '/*Hide the SubFeature List
            lstItemExtended.Visible = False
            '/*Make all the Lables visible
            For Each lbl As Label In lblList
                lbl.Visible = True
            Next
            For i = 0 To lstItems.Count - 1

                lstItems(i).Visible = True

            Next
            For i = 0 To lstHotItems.Count - 1

                lstHotItems(i).Visible = True

            Next
            ProcessState(cnSTATE_SUBLISTDONE)
        End If
    End Sub

    Private Sub lstItems_Click(sender As Object, e As EventArgs)
        Dim nListIndex As Integer
        Dim nListItem As Integer
        Dim lstItem As ListBox = CType(sender, ListBox)
        ' index of ListBox , cosmetic=0 , risk=1 , functional=2
        Dim Index As Integer = lstItems.IndexOf(lstItem)

        If lstItems(Index).SelectedIndex < 0 Then
            Exit Sub
        End If

        '/*Track the currently focused index
        '/*for stately 'Enter' translating to cmdOK
        m_nLastIndex = Index
        '/*Set the tracking variables for cmdHelp
        'm_nURLIndex = Index + 1
        'm_nURLItem = lstItems(Index).SelectedIndex + 1
        m_nURLIndex = Index
        m_nURLItem = lstItems(Index).SelectedIndex

        '/*Set the state
        If ProcessState(cnSTATE_LIST) Then
            '/*Increment for base 0 -> base 1 offset
            'nListIndex = Index + 1
            'nListItem = lstItems(Index).SelectedIndex + 1
            nListIndex = Index
            nListItem = lstItems(Index).SelectedIndex
        End If
    End Sub

    Private Sub frmDefectEditor_load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim nTopLevel As Integer
        Dim nMidLevel As Integer
        Dim nListIndex As Integer

        '/*Remove the Query_Unload button from the window
        mdlTools.DisableCloseX(Me)

        For nTopLevel = 1 To gcol_FeatureClasses.Count
            ' Get the current control array upper limit
            nListIndex = nTopLevel - 1
            ' If we need more than the initial list, add another set of list objects to the screen
            If (nListIndex + 1) > Me.lblList.Count Then
                ' Make the object add call
                AddDefectEditorList()
            End If

            ' Set the list object title
            Me.lblList(nListIndex).Text = gcol_FeatureClasses.Item(nListIndex).strTitle

            ' Set a reference to the list we are working on
            With Me.lstItems(nListIndex)
                ' Add the actual descriptions to the List
                For nMidLevel = 0 To gcol_FeatureClasses.Item(nListIndex).colClassFeatures.Count - 1
                    ' Add the item to the list; be sure to set the index since this was experiencing a strange anomaly
                    ' that was causing overwriting of the 1st and 2nd list items
                    .Items.Add(gcol_FeatureClasses.Item(nListIndex).colClassFeatures.Item(nMidLevel).strDesc)
                Next nMidLevel
            End With
        Next nTopLevel
        'disable buttons and gridView until needed
        fgFeatures.ClearSelection()
        cmdPrimary.Enabled = False
        cmdDelete.Enabled = False
        cmdComment.Enabled = False
        cmdNumericInput.Enabled = False
        fgFeatures.Enabled = False

        '/*Set a reference to the Feature
        With gcol_FeatureClasses.Item(0).colClassFeatures.Item(0)
            '/*Cat the Feature description
            '  strOut = .strDesc
            '/*If there is a SubFeature
            'If nSubFeature > 0 Then
            '    '/*Cat the SubFeature Description
            '    strOut = strOut & "->" & .colSub.Item(nSubFeature).strDesc
            'End If
        End With
        '/*Return the string
        'ConvertFeatureCoordToWord = "" 'strOut

        'Trigger the default disposition calculation
        UpdateDisposition()

    End Sub
    Private Sub AddDefectEditorList()
        Dim nIndex As Integer

        ' Increment the index by +1
        nIndex = Me.lblList.Count
        '-------------------------------------
        ' Label of the Listing Type
        ' Create and add a new Label control
        Dim newLabel As New Label With {
            .Visible = True
        }
        Me.lblList.Add(newLabel)
        Me.Controls.Add(newLabel)
        '-------------------------------------
        ' The Hot list
        ' Create and add a new ListBox control for lstHotItems
        Dim newLstHotItems As New ListBox With {
            .Visible = True
        }
        ' Assuming lstHotItems is a List(Of ListBox)
        Me.lstHotItems.Add(newLstHotItems)
        Me.Controls.Add(newLstHotItems)
        '-------------------------------------
        ' The actual defect list
        ' Create and add a new ListBox control for lstItems
        Dim newLstItems As New ListBox With {
            .Visible = True
        }
        Me.lstItems.Add(newLstItems)
        Me.Controls.Add(newLstItems)
    End Sub
    Private Sub cmdOK_Click(sender As Object, e As EventArgs)
        Dim bResult As Boolean

        ' Switch on the current state of Control focus
        Select Case m_State
            Case cnSTATE_COMMENT
                ' Trigger the update of the Comment
                ' Log the information
                SetComment()
                txtComment.Visible = False
            Case cnSTATE_NUMERIC
                If mdlDefectEditor.ValidateNumericInput(txtNumericInput.Text) Then
                    ' Add the data to the proper m_colDefect Feature
                    SetNumeric()
                    ' Switch the viewed controls back
                    txtNumericInput.Visible = False
                    fgFeatures.Visible = True
                Else
                    ' Raise some sort of alert
                    Exit Sub
                End If
            Case cnSTATE_LIST
                ' Translate the OK key into the last selected
                ' list item that would be double clicked
                lstItems_DoubleClick(lstItems(m_nLastIndex), EventArgs.Empty)
            Case cnSTATE_SUBLIST
                ' translate the OK key to the list double click
                lstItemExtended_DoubleClick()
            Case Else ' User is done adding Features
                ' Trap any event that may have opened us
                ' inappropriately
                If m_oUnit Is Nothing Then
                    ' Just hide the window
                    Me.Hide()
                Else
                    ' Toggle the OK button off to disable double triggering
                    cmdOK.Enabled = False
                    cmdCancel.Enabled = False
                    cmdHelp.Enabled = False
                    'txtPenCount.Enabled = False

                    ' -- Test the pen count input if in use and raise error if
                    ' -- out of range
                    If go_clsSystemSettings.bBadPenCountEnabled AndAlso Integer.TryParse(txtPenCount.Text, New Integer()) <= 0 Then
                        ' -- display error
                        'frmMessage.GenerateMessage("Quantity must be greater than zero.", "Error")
                        'frmMessage.ShowWindow()
                        MessageBox.Show("Error: Quantity must be greater than zero.")
                    Else
                        '-----------------------------------
                        ' Pass the unit on for processing
                        '-----------------------------------

                        ' -- Apply pen count if in use
                        If go_clsSystemSettings.bBadPenCountEnabled Then
                            Integer.TryParse(txtPenCount.Text, m_oUnit.nCount)
                        End If

                        ' Apply the window options to the pen
                        ' Set the objects defect collection to the private collection
                        m_oUnit.colPenDefects = m_colDefects
                        ' Set the shipped /not shipped flag
                        m_oUnit.bPenNotShipped = chkNotShipped.Checked

                        ' Add on the Test Bed value
                        'm_oUnit.strTestBed = TestBed()

                        ' Adjust processing if we are in an Edit Mode
                        If m_bUnitEdit Then
                            ' If the update failed abort back to the screen
                            bResult = mdlCreatePen.UpdateUnit(m_oUnit)
                        ElseIf m_bUnitEditRemote Then
                            ' If the update failed abort back to the screen
                            bResult = mdlCreatePen.UpdateRemoteUnit(m_oUnit)
                        Else
                            ' Make sure there are actually some defects
                            ' Associated with the pen
                            If m_colDefects.Count > 0 Then
                                ' Transmit the pen to the business server
                                bResult = mdlCreatePen.TransmitUnit(m_oUnit)

                                '------------------------------------------------------
                                ' See if we are in continuous mode and we don't
                                ' require a user ID, if so set the timer for return
                                ' from frmNextCap
                                '------------------------------------------------------
                                'If chkContLogging.Checked AndAlso Not go_clsSystemSettings.bBadPenIdRequired Then
                                '    Form1.tmrContinuousBad.Enabled = True
                                'End If
                            Else
                                ' Toggle the OK button off to disable double triggering
                                cmdOK.Enabled = True
                                cmdCancel.Enabled = True
                                cmdHelp.Enabled = True
                                txtPenCount.Enabled = True

                                ' We are not in edit mode
                                ' and no defects were selected
                                ' -- display error
                                'frmMessage.GenerateMessage("Bad entry must have defects.", "Error")
                                'frmMessage.ShowWindow()
                                MessageBox.Show("Error: Bad entry must have defects.")
                                Exit Sub
                            End If
                        End If

                        ' Close the window
                        Me.Hide()
                        'Form1.focusInput()
                        'MessageBox.Show("Close the window")

                        ' Generate error if the transaction failed
                        If bResult Then
                            ' Show the feeback if enabled
                            'If go_clsSystemSettings.bBadPenFeedBack Then
                            '    'frmBadPen.Show()
                            '    'frmBadPen.Focus()
                            '    MessageBox.Show("show frmBadPen")
                            'End If
                            ' Release the lock on the client
                            gb_InputBusyFlag = False
                        Else
                            ' INsure we are switching between eidt and create
                            If m_bUnitEdit OrElse m_bUnitEditRemote Then
                                mdlLotManager.TransactionFailure("mdlCreatePen.UpdateUnit()")
                            Else
                                mdlLotManager.TransactionFailure("mdlCreatePen.TransmitUnit()")
                            End If
                        End If
                    End If

                    ' Toggle the OK button off to disable double triggering
                    cmdOK.Enabled = True
                    cmdCancel.Enabled = True
                    cmdHelp.Enabled = True
                    'txtPenCount.Enabled = True

                    MessageBox.Show("Close the window")

                End If
        End Select

        ' Clear the state flag
        ProcessState(cnSTATE_READY)
        ' Set the global button selected for this form
        'gn_frmDefectEditorBttnReturn = 1
    End Sub

    Private Sub lstItems_DoubleClick(sender As Object, e As EventArgs)
        Dim listBox As ListBox = CType(sender, ListBox)
        Dim index As Integer = lstItems.IndexOf(listBox)

        If listBox.SelectedIndex < 0 Then
            Exit Sub
        End If

        If ProcessState(cnSTATE_LIST) Then
            'Dim nListIndex As Integer = index + 1
            'Dim nListItem As Integer = listBox.SelectedIndex + 1
            Dim nListIndex As Integer = index
            Dim nListItem As Integer = listBox.SelectedIndex

            If nListItem >= 0 Then
                If Not mdlDefectEditor.IsExcludedClass(nListIndex) Then
                    If Not HasSecondary(nListIndex, nListItem) Then
                        AddFeature(nListIndex, nListItem)
                    End If
                    listBox.ClearSelected()
                    'CauseUnselect()
                Else
                    If Not HasSecondary(nListIndex, nListItem) Then
                        AddCause(nListIndex, nListItem)
                    End If
                    listBox.ClearSelected()
                    'CauseUnselect()
                End If
            End If
        End If
    End Sub

    Private Function HasSecondary(ByVal nListIndex As Integer, ByVal nListItem As Integer) As Boolean
        Dim secondaryCount = gcol_FeatureClasses(nListIndex).colClassFeatures(nListItem).colSub.Count
        If secondaryCount > 0 Then
            m_nListIndex = nListIndex
            m_nListItem = nListItem

            lstItemExtended.Items.Clear()
            For Each subItem In gcol_FeatureClasses(nListIndex).colClassFeatures(nListItem).colSub
                lstItemExtended.Items.Add(subItem.strDesc)
            Next
            For i = 0 To lstItems.Count - 1

                lstItems(i).Visible = False

            Next
            For i = 0 To lstHotItems.Count - 1

                lstHotItems(i).Visible = False

            Next

            ProcessState(cnSTATE_SUBLIST)
            lstItemExtended.Visible = True


            Return True
        End If
        Return False
    End Function

    Private Sub AddFeature(ByRef nGroup As Integer, ByRef nFeature As Integer, Optional ByRef nSubFeature As Integer = -1)
        Dim bPrimary As Boolean

        '/*Insure that the Feature has not already been selected
        If Not (IsFeatureUsed(nGroup, nFeature, nSubFeature)) Then
            '/*Add the item to the Hot List
            UnitAddHotFeature(nGroup, nFeature, nSubFeature)
            '/*Test for rule based Primary
            bPrimary = IsFeaturePrimary(nGroup)
            '/*Update the Defect member
            UnitAddFeature(nGroup, nFeature, nSubFeature, bPrimary)
            '/*Update the Grid visual listing
            fgFeatureAddItem(nGroup, nFeature, nSubFeature, bPrimary)
        End If
    End Sub

    Private Sub AddCause(ByRef nGroup As Integer, ByRef nFeature As Integer, Optional ByRef nSubFeature As Integer = -1)
        Dim nItem As Integer

        '/*Use the currently selected row
        If fgFeatures.CurrentCell IsNot Nothing Then
            nItem = fgFeatures.CurrentCell.RowIndex
            If nItem >= 0 Then
                '/*Add the item to the Hot List
                UnitAddHotFeature(nGroup, nFeature, nSubFeature)
                '/*Update the Defect member
                UnitAddCause(nGroup, nFeature, nSubFeature, nItem)
            End If
        End If
    End Sub

    Private Sub UnitAddCause(ByRef nGroup As Integer, ByRef nFeature As Integer, ByRef nSubFeature As Integer, ByRef nItem As Integer)
        Dim strSFCd As String = String.Empty
        Dim strSFDesc As String = String.Empty
        Const cnCauseCol As Integer = 4
        Const cnSubCauseCol As Integer = 5

        '/*Make sure we are passed a vlid feature
        If nItem >= 0 AndAlso nItem <= m_colDefects.Count - 1 Then
            '/*Get the Cause codes from the list
            With gcol_FeatureClasses.Item(nGroup)
                '/*Sort out the possible sub items
                If nSubFeature >= 0 Then
                    strSFCd = .colClassFeatures.Item(nFeature).colSub.Item(nSubFeature).strCode
                    strSFDesc = .colClassFeatures.Item(nFeature).colSub.Item(nSubFeature).strDesc
                End If

                m_colDefects(nItem + 1).strCauseCd = .colClassFeatures.Item(nFeature).strCode
                m_colDefects(nItem + 1).strCauseDesc = .colClassFeatures.Item(nFeature).strDesc
                m_colDefects(nItem + 1).strSubCauseCd = strSFCd
                m_colDefects(nItem + 1).strSubCauseDesc = strSFDesc
            End With

            '/*Transfer the comment to the screen
            fgFeatures.Rows(nItem).Cells(cnCauseCol).Value = m_colDefects(nItem + 1).strCauseDesc
            fgFeatures.Rows(nItem).Cells(cnSubCauseCol).Value = strSFDesc
        End If
    End Sub

    Private Sub fgFeatureAddItem(ByRef nGroup As Integer, ByRef nFeature As Integer, ByRef nSubFeature As Integer, ByRef bPrimary As Boolean)
        Dim rowItems As New List(Of String)
        '----------------------------------------
        '/*Build up a Tab delimeted string
        '/*Set the Primary flag
        If bPrimary Then
            rowItems.Add("Yes")
            '/*If there is a primary marking
            If m_PrimaryRow >= 0 AndAlso fgFeatures.Rows.Count > m_PrimaryRow Then
                fgFeatures.Rows(m_PrimaryRow).Cells(0).Value = ""
            End If
            '/*Set the new primary row
            m_PrimaryRow = fgFeatures.Rows.Count - 1
        Else
            rowItems.Add("")
        End If

        '/*Set reference to the group
        With gcol_FeatureClasses(nGroup)
            '/*Add on the Type
            rowItems.Add(.strTitle)
            '/*Add on the Feature to the output string
            rowItems.Add(.colClassFeatures(nFeature).strDesc)
            '/*Switch on whether the Feature has a SubFeature
            'If nSubFeature > 0 Then
            If nSubFeature >= 0 Then
                rowItems.Add(.colClassFeatures(nFeature).colSub(nSubFeature).strDesc)
            Else
                rowItems.Add("")
            End If
        End With

        '/*Tack on any Cause items if the Cause Index Member has been used
        If m_nCauseIndex >= 0 Then
            '/*Get a reference to the object
            With gcol_FeatureClasses(m_nCauseIndex).colClassFeatures(m_nCauseItem)
                '/*Add teh description to the string
                rowItems.Add(.strDesc)
                '/*get any subitems used by this Cause
                If m_nCauseSubItem >= 0 Then
                    '/*Append on the SubItem
                    rowItems.Add(.colSub(m_nCauseSubItem).strDesc)
                Else
                    rowItems.Add("")
                End If
            End With
            '/*Insure that the items are now reset
            m_nCauseIndex = -1 : m_nCauseItem = -1 : m_nCauseSubItem = -1
        End If

        '/*Add the item to the list
        fgFeatures.Rows.Add(rowItems.ToArray())

        '/*Need to set a toggle that supresses this when the rows
        '/*do not extend below the bottom of the Grid window
        '/*This is set to 6 rows since we never resize the
        '/*height of this control.
        If fgFeatures.Rows.Count > 6 Then
            fgFeatures.FirstDisplayedScrollingRowIndex = fgFeatures.Rows.Count - 1
            fgFeatures.CurrentCell = fgFeatures.Rows(fgFeatures.Rows.Count - 1).Cells(0)
        End If

        '/*Reset the comment and the Numeric Input
        txtComment.Text = ""
        txtNumericInput.Text = ""
    End Sub

    Private Sub UnitAddFeature(ByRef nGroup As Integer, ByRef nFeature As Integer, ByRef nSubFeature As Integer, ByRef bPrimary As Boolean)
        Dim strSFCd As String = String.Empty
        Dim strSFDesc As String = String.Empty
        Dim strSubCCd As String = String.Empty
        Dim strSubCDesc As String = String.Empty
        Dim strCCd As String = String.Empty
        Dim strCDesc As String = String.Empty

        '/*Get any associated cause code
        If m_nCauseIndex >= 0 Then
            With gcol_FeatureClasses.Item(m_nCauseIndex).colClassFeatures.Item(m_nCauseItem)
                strCCd = .strCode
                strCDesc = .strDesc
                '/*Get any sub items
                If m_nCauseSubItem >= 0 Then
                    strSubCCd = .colSub.Item(m_nCauseSubItem).strCode
                    strSubCDesc = .colSub.Item(m_nCauseSubItem).strDesc
                End If
            End With
            CauseUnselect()

            '/*Process the Hot List for the cause
            UnitAddHotFeature(m_nCauseIndex, m_nCauseItem, m_nCauseSubItem)
            '/*Reset the Cause values
            'm_nCauseIndex = 0
            'm_nCauseItem = 0
            'm_nCauseSubItem = 0
        End If

        With gcol_FeatureClasses.Item(nGroup)
            '/*Sort out the possible sub items
            If nSubFeature >= 0 Then
                strSFCd = .colClassFeatures.Item(nFeature).colSub.Item(nSubFeature).strCode
                strSFDesc = .colClassFeatures.Item(nFeature).colSub.Item(nSubFeature).strDesc
            End If

            'If txtComment Is Nothing Then
            '    txtComment = New TextBox()
            '    txtComment.Text = String.Empty
            'End If

            '/*Add the new Feature to the pen
            m_colDefects.Add(.lngSeverity, bPrimary, nGroup, nFeature, nSubFeature, .colClassFeatures.Item(nFeature).strCode, .colClassFeatures.Item(nFeature).strDesc, strSFCd, strSFDesc, strCCd, strCDesc, strSubCCd, strSubCDesc, .strClassType, .strTitle, txtComment.Text, String.Empty)

            '/*Reset the comment and the Numeric Input
            'txtComment.Text = String.Empty
        End With

        '/*Clear any existing primary
        If bPrimary Then
            '/*Remove the current setting
            If m_PrimaryIndex > 0 Then m_colDefects.Item(m_PrimaryIndex).bPrimary = False
            '/*Store the new setting in are private variable
            m_PrimaryIndex = m_colDefects.Count
        End If
    End Sub

    Private Sub UnitAddHotFeature(ByRef nClass As Integer, ByRef nFeature As Integer, ByRef nSubFeature As Integer)
        '/*Make sure the Item is not alreay on the Hot list
        If Not IsHotFeatureUsed(nClass, nFeature, nSubFeature) Then
            '/*Set a reference to the entry point
            With gcol_FeatureClasses.Item(nClass)
                '/*See if we are at the MAX number of Hot Items
                If .colHotFeature.Count >= .lngHotMax Then
                    '/*Remove the Feature from teh structure
                    .colHotFeature.Remove(1)
                    '/*Remove the first Item from the on-screen list
                    lstHotItems(nClass).Items.RemoveAt(0)
                End If
                '/*Add the item to the Hot collection
                .colHotFeature.Add(Convert.ToInt64(nFeature), Convert.ToInt64(nSubFeature))
                '/*Tack the Feature on to the Hot list; this requires
                '/*an adjustment to place it on the right list.
                'lstHotItems(nClass - 1).Items.Add(mdlDefectEditor.ConvertFeatureCoordToWord(nClass, nFeature, nSubFeature))
                lstHotItems(nClass).Items.Add(mdlDefectEditor.ConvertFeatureCoordToWord(nClass, nFeature, nSubFeature))
            End With
        End If
    End Sub

    Private Function IsHotFeatureUsed(ByVal nClass As Integer, ByVal nFeature As Integer, ByVal nSubFeature As Integer) As Boolean
        '/*Set a reference to the entry point
        With gcol_FeatureClasses.Item(nClass).colHotFeature
            '/*Loop through teh current set of Hot Features
            For lngLoop As Integer = 0 To .Count - 1
                '/*See if this is a member of the set
                If .Item(lngLoop).lngFeature = CLng(nFeature) AndAlso .Item(lngLoop).lngSubFeature = CLng(nSubFeature) Then
                    '/*Set return to True and exit function
                    Return True
                End If
            Next
        End With

        Return False
    End Function

    Private Function IsFeatureUsed(ByRef nClass As Integer, ByRef nFeature As Integer, ByRef nSubFeature As Integer) As Boolean
        Dim nLoop As Integer

        If m_colDefects IsNot Nothing AndAlso m_colDefects.Count > 0 Then
            '/*Loop through the list
            For nLoop = 1 To m_colDefects.Count
                With m_colDefects.Item(nLoop)
                    '/*Test the items Feature indexes to see if they match
                    If .nFeatureClass = nClass AndAlso .nFeatureIndex = nFeature AndAlso .nSubFeatureIndex = nSubFeature Then
                        Return True
                    End If
                End With
            Next nLoop
        End If
        Return False
    End Function

    Private Function IsFeaturePrimary(ByRef nGroup As Integer) As Boolean
        Dim weight As Long = gcol_FeatureClasses(nGroup).lngSeverity
        If m_PrimaryWeight < weight Then
            m_PrimaryWeight = weight
            Return True
        End If
        Return False
    End Function
    Private Sub SetNumeric(Optional ByRef strNumericInput As String = "")
        Dim nCol As Integer
        Dim nRow As Integer

        ' Determine where the numeric input column is
        If fgFeatures.ColumnCount = 8 Then
            nCol = 7
        Else
            nCol = 5
        End If

        If String.IsNullOrEmpty(strNumericInput) Then
            ' Use the currently selected row
            nRow = fgFeatures.CurrentCell.RowIndex

            ' Transfer the numeric input to the screen
            fgFeatures.Rows(nRow).Cells(nCol).Value = txtNumericInput.Text

            ' Transfer the numeric input to the private Defect structure
            m_colDefects(nRow + 1).strNumericInput = txtNumericInput.Text
        Else
            ' Use the last row in the DataGridView
            nRow = fgFeatures.RowCount - 1
            fgFeatures.Rows(nRow).Cells(nCol).Value = strNumericInput
        End If
    End Sub

    Private Sub CauseUnselect()
        '/*filter out any out of range selections
        If m_nCauseIndex >= 0 OrElse m_nCauseItem >= 0 Then
            '/*Clear the Hot List as well
            ClearHotList(m_nCauseIndex)
            '/*Unselect the Cause Item from the List Object
            lstItems(m_nCauseIndex).SetSelected(m_nCauseItem, False)
        End If
    End Sub

    Private Function TestBed() As String
        Try
            ' Add on the Test Bed if we have any
            If cmbTestBedNumber.Items.Count > 0 Then
                Return cmbTestBedNumber.Items(cmbTestBedNumber.SelectedIndex).ToString() & "-" & cmbTestBedType.Items(cmbTestBedType.SelectedIndex).ToString()
            Else
                ' Return no list values to use
                Return "0-NONE"
            End If
        Catch ex As Exception
            Return "0-ERROR"
        End Try
    End Function
    Private Sub ClearHotList(Optional ByRef nIndex As Integer = 0)
        Dim nLoop As Integer

        Try
            '/*Set a reference to the List
            With lstHotItems(nIndex)
                For nLoop = 0 To .Items.Count - 1
                    If .GetSelected(nLoop) Then .SetSelected(nLoop, False)
                Next nLoop
            End With
        Catch ex As Exception

        End Try
    End Sub

    Private Sub SetComment(Optional ByRef strComment As String = "")
        Dim nCol As Integer
        Dim nRow As Integer

        ' Determine where the comment column is
        If fgFeatures.ColumnCount = 8 Then
            nCol = 6
        Else
            nCol = 4
        End If

        If String.IsNullOrEmpty(strComment) Then
            ' Use the currently selected row
            nRow = fgFeatures.CurrentCell.RowIndex

            ' Transfer the comment to the screen
            fgFeatures.Rows(nRow).Cells(nCol).Value = txtComment.Text

            ' Transfer the comment to the private Defect structure
            m_colDefects(nRow + 1).strComment = txtComment.Text
        Else
            ' Use the last row in the DataGridView
            nRow = fgFeatures.RowCount - 1
            fgFeatures.Rows(nRow).Cells(nCol).Value = strComment
        End If
    End Sub
    Private Sub cmdCancel_Click(sender As Object, e As EventArgs)
        Dim lngLoop As Long

        If m_State = cnSTATE_SUBLIST Then
            lstItemExtended.Visible = False
            ' Make all the Labels visible
            For lngLoop = 0 To lblList.Count - 1
                lblList(lngLoop).Visible = True
            Next lngLoop
            ' Force the state off
            ProcessState(cnSTATE_SUBLISTDONE)
        ElseIf m_State = cnSTATE_COMMENT Then
            txtComment.Text = ""
            txtComment.Visible = False
        ElseIf m_State = cnSTATE_NUMERIC Then
            txtNumericInput.Text = ""
            txtNumericInput.Visible = False
            fgFeatures.Visible = True
        Else
            ' If we are in an Edit mode then we need to reset
            ' the flag and Release the pen that was under edit
            If m_bUnitEdit Or m_bUnitEditRemote Then
                ' Make sure the pen was not released
                ' by an external process such as Update
                If m_oUnit IsNot Nothing Then
                    ' Try to release the pen
                    If Not mdlCreatePen.ReleaseUnit(m_oUnit) Then
                        ' if the release failed and we returned abort
                        ' back to the defect editor.
                        Exit Sub
                    End If
                End If
                m_bUnitEdit = False
                m_bUnitEditRemote = False
            End If

            ' Remove the continuous logging if cancel is pressed
            If chkContLogging.Checked Then
                chkContLogging.Checked = False
            End If
            ' Release the lock on the client
            gb_InputBusyFlag = False

            ' Drop the window
            Me.Hide()
            '' -- Refocus the cursor in the main window
            'Form1.focusInput()
            'MessageBox.Show("Close the window")
        End If
        ' Clear the state flag
        ProcessState(cnSTATE_READY)
    End Sub
    Private Function ProcessState(ByRef nNewState As Integer) As Boolean
        Dim lngLoop As Integer

        ' Process the state according to rules
        If (m_State = cnSTATE_COMMENT Or m_State = cnSTATE_NUMERIC) And nNewState <> cnSTATE_READY Then
            ' Do nothing for now
        ElseIf (nNewState = cnSTATE_LIST Or nNewState = cnSTATE_READY) And m_State = cnSTATE_SUBLIST Then
            ' Do nothing for now
        Else
            ' Hide the sublist if it is the present state and the new state is not the sublist
            If m_State = cnSTATE_SUBLIST And nNewState <> cnSTATE_SUBLIST Then
                lstItemExtended.Visible = False
                ' Make all the Labels visible

                For lngLoop = 0 To lblList.Count - 1
                    lblList(lngLoop).Visible = True
                Next lngLoop
            End If

            If nNewState = cnSTATE_DETAIL Then
                ' Control the state of the units feature buttons by setting the appropriate enabled
                cmdPrimary.Enabled = True
                cmdDelete.Enabled = True
                cmdComment.Enabled = True
                cmdNumericInput.Enabled = True
            Else
                ' Control the state of the units feature buttons by setting the appropriate enabled
                cmdPrimary.Enabled = False
                cmdDelete.Enabled = False
                cmdComment.Enabled = False
                cmdNumericInput.Enabled = False
            End If
            ' If the state is ready set the control focus to the cmdOK button
            If Me.Visible And cmdOK.Enabled Then cmdOK.Focus()
            ' Return that the caller has control
            ProcessState = True
            ' Set the new state
            m_State = nNewState
        End If
    End Function
    '=======================================================
    'Routine: AddUnitType(str) and AddTestBed
    'Purpose: This is an interface for adding Unit Types
    '         and Test Beds to the DefectEditor screen.
    '         Basically it takes the received string and
    '         insures that it has not already been placed
    '         in the drop down list.
    '
    'Globals:None
    '
    'Input: strData - The item to add to the list.
    '
    'Return:None
    '
    'Modifications:
    '   11-23-1998 As written for Pass1.5
    '
    '
    '=======================================================
    Public Sub AddUnitType(strData As String)
        ' 检查ComboBox中是否已存在该值
        If cmbComponent.Items.Contains(strData) Then
            Return
        End If

        ' 添加新项目
        cmbComponent.Items.Add(strData)

        ' 确保选中第一个项目
        If cmbComponent.Items.Count > 0 Then
            cmbComponent.SelectedIndex = 0
        End If

        ' 确保下拉列表可见
        If cmbComponent.Items.Count > 1 Then
            cmbComponent.Visible = True
        End If
    End Sub
    Public Sub AddTestBed(strTestBedNumber As String, strTestBedType As String)
        ' 检查测试台编号是否已存在
        If cmbTestBedNumber.Items.Contains(strTestBedNumber) Then
            Return
        End If

        ' 添加新的测试台编号和类型
        cmbTestBedNumber.Items.Add(strTestBedNumber)
        cmbTestBedType.Items.Add(strTestBedType)

        ' 确保选中第一个项目
        If cmbTestBedNumber.Items.Count > 0 Then
            cmbTestBedNumber.SelectedIndex = 0
            cmbTestBedType.SelectedIndex = 0
        End If

        ' 确保下拉列表可见
        If cmbTestBedNumber.Items.Count > 1 Then
            cmbTestBedNumber.Visible = True
        End If
    End Sub

    Public Sub ShowWindow()
        Try
            ' Initialize the window (implement WindowInit as needed)
            WindowInit()
            ' Remove previous form from stack if needed
            If Not String.IsNullOrEmpty(m_strFrmId) Then
                mdlWindow.RemoveForm(m_strFrmId)
            End If
            ' Add this form to the stack and get new ID
            m_strFrmId = mdlWindow.AddForm(Me)
            ' Show the form (non-modal)
            Me.Show()
            ' Set focus to OK button if visible
            If Me.Visible AndAlso cmdOK.Visible Then
                cmdOK.Focus()
            End If
        Catch
            ' Optionally log or handle the error
        End Try
    End Sub

    Private Sub WindowInit()
        '/*Add any other screen prep calls here
    End Sub

    Public Sub SetUnitId(Optional ByRef strUnitId As String = "")
        '/*Remove the Edit flag from this form
        m_bUnitEdit = False

        ' If the unit comes in blank, assign a time stamp to it
        If String.IsNullOrEmpty(strUnitId) Then
            strUnitId = GetIdDttm()
        End If

        ' Generate the private clsPen object
        m_oUnit = New clsPen()

        ' Set the Unit's Id
        m_oUnit.strPenId = strUnitId

        '/*Set the default count
        m_oUnit.nCount = 1

        ' -- Set the user pen count input
        If go_clsSystemSettings.bBadPenCountEnabled Then
            txtPenCount.Text = "1"
        End If

        '/*Now set the objects standard properties
        '/*This is done via ByRef on the passed pointer object
        mdlCreatePen.SetUnitStdProperties(m_oUnit)

        '/*Set the txtBox on the screen
        txtUnitId.Text = strUnitId
        txtLotId.Text = m_oUnit.strGroupId

        '/*Set the default disposition
        If String.IsNullOrEmpty(m_strDisposition) OrElse m_strDisposition = "G" Then
            optGood.Checked = True
            m_oUnit.strDisposition = "G"
        ElseIf m_strDisposition = "R" Then
            optReclaim.Checked = True
            m_oUnit.strDisposition = "R"
        ElseIf m_strDisposition = "S" Then
            optScrap.Checked = True
            m_oUnit.strDisposition = "S"
        End If
    End Sub

    Public Sub UpdateDisposition()
        Dim strDisposition As String

        Try
            If go_clsSystemSettings IsNot Nothing Then
                strDisposition = go_clsSystemSettings.strDispositionDefault

                Select Case strDisposition
                    Case "Reclaim"
                        m_strDisposition = "R"
                    Case "Scrap"
                        m_strDisposition = "S"
                    Case Else ' default
                        m_strDisposition = "G"
                End Select
            End If
        Catch ex As Exception
            Debug.WriteLine($"Error in UpdateDisposition: {ex.Message}")
        End Try
    End Sub

    '/*Standard Method; set the value on the Unit object for the
    '/*unit disposition.
    Private Sub optGood_MouseUp(sender As Object, e As MouseEventArgs) Handles optGood.MouseUp
        m_oUnit.strDisposition = "G"
    End Sub

    '/*Standard Method; set the value on the Unit object for the
    '/*unit disposition.
    Private Sub optReclaim_MouseUp(sender As Object, e As MouseEventArgs) Handles optReclaim.MouseUp
        m_oUnit.strDisposition = "R"
    End Sub

    '/*Standard Method; set the value on the Unit object for the
    '/*unit disposition.
    Private Sub optScrap_MouseUp(sender As Object, e As MouseEventArgs) Handles optScrap.MouseUp
        m_oUnit.strDisposition = "S"
    End Sub

    Private Sub fgFeatures_SelectionChanged(sender As Object, e As EventArgs) Handles fgFeatures.SelectionChanged
        Dim hasSelection As Boolean = fgFeatures.CurrentRow IsNot Nothing AndAlso fgFeatures.CurrentRow.Index >= 0
        cmdPrimary.Enabled = hasSelection
        cmdDelete.Enabled = hasSelection
        cmdComment.Enabled = hasSelection
        cmdNumericInput.Enabled = hasSelection
    End Sub
    Private Sub cmdPrimary_Click(sender As Object, e As EventArgs) Handles cmdPrimary.Click
        'If fgFeatures.SelectedRows.Count > 0 Then
        '    Dim nRowSel As Integer = fgFeatures.SelectedRows(0).Index
        '    SetPrimary(nRowSel)
        'End If
        Dim nRowSel As Integer = -1

        If fgFeatures.SelectedCells.Count > 0 Then
            nRowSel = fgFeatures.SelectedCells(0).RowIndex
        ElseIf fgFeatures.CurrentCell IsNot Nothing Then
            nRowSel = fgFeatures.CurrentCell.RowIndex
        End If

        If nRowSel >= 0 AndAlso Not String.IsNullOrEmpty(fgFeatures.Rows(nRowSel).Cells(1).Value) Then
            SetPrimary(nRowSel)
        End If
    End Sub
    Private Sub SetPrimary(ByVal nRowSel As Integer)
        ' 注意：Collection 是 1 基索引，所以需要 +1
        Dim currentDefect As clsDefect = CType(m_colDefects(m_PrimaryIndex), clsDefect)
        Dim newDefect As clsDefect = CType(m_colDefects(nRowSel + 1), clsDefect)
        If currentDefect.lngSeverity <= newDefect.lngSeverity Then
            ' 清空当前 Primary 单元格
            fgFeatures.Rows(m_PrimaryRow).Cells(0).Value = ""

            ' 移除旧 Primary
            currentDefect.bPrimary = False

            ' 设置新的 Primary
            fgFeatures.Rows(nRowSel).Cells(0).Value = "Yes"
            newDefect.bPrimary = True

            ' 更新跟踪变量
            m_PrimaryRow = nRowSel
            m_PrimaryIndex = nRowSel + 1
        End If
    End Sub
    Private Sub fgFeatures_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles fgFeatures.CellEndEdit

        ' clear error message

        fgFeatures.Rows(e.RowIndex).ErrorText = ""

        Dim nRow As Integer = e.RowIndex

        Dim cellValue As String = fgFeatures.Rows(nRow).Cells(e.ColumnIndex).Value?.ToString().Trim()

        If nRow < 0 OrElse nRow >= m_colDefects.Count Then

            Return

        End If

        ' handle Comment and Numeric columns

        Select Case e.ColumnIndex

            Case 4 ' Comment

                'ProcessState(cnSTATE_COMMENT)

                HandleCommentEdit(nRow, cellValue)

            Case 5 ' Numeric

                'ProcessState(cnSTATE_NUMERIC)

                HandleNumericEdit(nRow, cellValue)

        End Select

    End Sub

    Private Sub HandleCommentEdit(nRow As Integer, inputValue As String)

        ' if input message only space , set as default value ""

        Dim processedValue As String = If(String.IsNullOrWhiteSpace(inputValue), "", inputValue)

        ' update dataGridView fgFeatures

        fgFeatures.Rows(nRow).Cells(4).Value = processedValue

        ' call setComment to save to m_colDefects

        SetComment(nRow, processedValue)

    End Sub
    Private Sub HandleNumericEdit(nRow As Integer, inputValue As String)

        Dim displayValue As String

        Dim storageValue As String

        ' not-numeric/space/empty set as default value "0"

        If Not Integer.TryParse(inputValue, Nothing) Then

            ' display as empty but m_colDefects save "0"

            displayValue = ""

            storageValue = "0"

        Else

            ' space/empty display empty, m_colDefects save "0"

            If String.IsNullOrWhiteSpace(inputValue) Then

                displayValue = ""

                storageValue = "0"

            Else

                displayValue = inputValue

                storageValue = inputValue

            End If

        End If

        ' update fgFeatures

        fgFeatures.Rows(nRow).Cells(5).Value = displayValue

        SetNumeric(nRow, storageValue)

    End Sub

    Private Sub SetComment(nRow As Integer, Optional strComment As String = "")

        Dim nCol As Integer = If(fgFeatures.Columns.Count = 8, 6, 4)

        ' update fgFeatures

        fgFeatures.Rows(nRow).Cells(nCol).Value = strComment

        ' save to m_colDefects

        Dim collectionIndex As Integer = nRow + 1 ' m_colDefects is 1-based index

        If collectionIndex <= m_colDefects.Count Then

            m_colDefects.Item(collectionIndex).strComment = strComment

        End If

    End Sub

    Private Sub SetNumeric(nRow As Integer, Optional strNumericInput As String = "")

        Dim nCol As Integer = If(fgFeatures.Columns.Count = 8, 7, 5)

        ' update fgFeatures
        fgFeatures.Rows(nRow).Cells(nCol).Value = fgFeatures.Rows(nRow).Cells(5).Value

        ' save to m_colDefects
        Dim collectionIndex As Integer = nRow + 1 ' m_colDefects is 1-based index
        If collectionIndex <= m_colDefects.Count Then
            m_colDefects.Item(collectionIndex).strNumericInput = strNumericInput
        End If
    End Sub
    Private Sub fgFeatures_CellValidating(sender As Object, e As DataGridViewCellValidatingEventArgs) Handles fgFeatures.CellValidating
        If e.ColumnIndex = 5 Then
            Dim input As String = e.FormattedValue?.ToString().Trim()
            If Not String.IsNullOrWhiteSpace(input) AndAlso Not Integer.TryParse(input, Nothing) Then
                fgFeatures.Rows(e.RowIndex).ErrorText = "Please input valid numeric"
                e.Cancel = True ' Prevent editing from completing until valid input is enter
            End If
        End If
    End Sub
    Private Sub cmdDelete_Click(sender As Object, e As EventArgs) Handles cmdDelete.Click
        '/*Insure that this is not the header row
        If fgFeatures.CurrentCell IsNot Nothing AndAlso fgFeatures.CurrentCell.RowIndex <> -1 Then
            Dim selectedRowIndex As Integer = fgFeatures.CurrentCell.RowIndex
            DeleteFeature(selectedRowIndex)
        End If
    End Sub
    Private Sub DeleteFeature(ByRef lngRowIndex As Integer)
        Dim nRow As Integer
        Dim bPrimary As Boolean
        ' m_colDefects is 1-based
        Dim collectionIndex As Integer = lngRowIndex + 1

        ' '/*Store off the Primary value
        If collectionIndex <= m_colDefects.Count Then
            bPrimary = m_colDefects(collectionIndex).bPrimary
        Else
            Return
        End If

        '/*Switch correctly to remove pens from the FlexGrid
        '/*This is due to the fact that you must have n+1 rows
        '/*to have a FixedRow for some operations
        If fgFeatures.Rows.Count > 0 AndAlso lngRowIndex < fgFeatures.Rows.Count Then
            fgFeatures.Rows.RemoveAt(lngRowIndex)
        End If

        '/*Remove form the defect set
        m_colDefects.Remove(collectionIndex)

        '/*If this was the only row left
        If fgFeatures.Rows.Count = 0 Then
            '/*Reset all Primary holders
            m_PrimaryRow = 0
            m_PrimaryIndex = 0
            m_PrimaryWeight = -1
        Else
            '/*See if this was the primary symptom
            If bPrimary Then
                '/*Reset all Primary holders
                m_PrimaryRow = 0
                m_PrimaryIndex = 0
                m_PrimaryWeight = -1

                '/*Find the most severe defect from top to bottom of the list
                For nRow = 1 To m_colDefects.Count
                    '/*Is this the primary?
                    If IsFeaturePrimary(m_colDefects.Item(nRow).nFeatureClass) Then
                        '/*Wipe any older Primaries
                        If m_PrimaryIndex > 0 Then
                            '/*Clear the current cell
                            fgFeatures.Rows(m_PrimaryIndex - 2).Cells(0).Value = ""
                            '/*Remove the old primary
                            m_colDefects(m_PrimaryIndex - 1).bPrimary = False
                        End If

                        '/*Set the ne primary
                        m_colDefects(nRow).bPrimary = True
                        ' '/*Store the new setting in are private variable（keep 1-based）
                        m_PrimaryIndex = nRow + 1
                        '/*Swith the private members to the new item
                        m_PrimaryRow = nRow
                        '/*Set the new cell to "Yes"
                        fgFeatures.Rows(nRow - 1).Cells(0).Value = "Yes"
                    End If
                Next nRow
            Else
                '/*If this was not the primary then we need to
                '/*find where the primary is at now
                For nRow = 1 To m_colDefects.Count
                    If m_colDefects(nRow).bPrimary Then
                        ' '/*Store the new setting in are private variable（keep 1-based）
                        m_PrimaryIndex = nRow + 1
                        m_PrimaryRow = nRow
                        Exit For
                    End If
                Next nRow
            End If
        End If
    End Sub
    Private Sub fgFeatures_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles fgFeatures.CellClick
        If e.RowIndex <> -1 Then
            ProcessState(cnSTATE_DETAIL) ' enable fgFeatures buttons
        End If
    End Sub

    '=======================================================
    'Routine: frmDefectEditor.ResetHotlist()
    'Purpose: This clears all of the items listed in
    '         each of the Hotlists. The use for this
    '         is to function in conjunction with the
    '         colFeatureCalsses.ResetHotlist routine.
    '
    'Globals:None
    '
    'Input:None
    '
    'Return:None
    '
    'Modifications:
    '   11-16-1998 As written for Pass1.4
    '
    '
    '=======================================================
    Public Sub ResetHotlist()
        For Each listBox As ListBox In lstHotItems
            If listBox IsNot Nothing Then
                listBox.Items.Clear()
            End If
        Next
    End Sub

End Class
