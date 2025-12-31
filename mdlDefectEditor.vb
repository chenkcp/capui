Imports System.IO
Imports System.Windows.Documents

Module mdlDefectEditor
    ' Temporary MAX for Hot Feature List
    Private Const m_lngHotMax As Integer = 3
    ' Flag to indicate that there is no defect structure
    Private m_bDataAvailable As Boolean = False
    ' Reserved word for Cause category
    Public Const cstrClassExclude As String = "CAUSE"
    'Public Const cstrClassExclude As String = "FUNCTIONAL"
    ' Storage for the Header Titles used on the fgFeature Flex Grid
    Private m_colfgFeatureHeader As New List(Of String)
    ' Flag to indicate that Cause category is in use
    Private m_bCauseEnabled As Boolean = False
    Public gcol_FeatureClasses As colFeatureClasses
    Public Sub ConfigureDefectEditor()
        '/*Remove any Defect/Feature Classes that hvae zero items
        CleanFeatureObject()
        '/*Configure the header map used for the FlexGrid
        Configure_fgFeature()
        '/*Set flag that gcol_FeatureClasses is alive
        m_bDataAvailable = True
        '/*Map the gcol_FEatureClasses object descriptions
        '/*to the List objects on frmDefectEditor
        FillDefectEditorLists()
    End Sub
    Public Sub CleanFeatureObject()
        Dim lngItem As Long
        Dim colRemove As New Collection

        '/*Cycle through the collection and mark any
        '/*any zero sets for removal
        For lngItem = 1 To gcol_FeatureClasses.Count
            '/*There is no use in a Class with zero Level-1 descriptions
            If gcol_FeatureClasses.Item(lngItem - 1).colClassFeatures.Count = 0 Then
                colRemove.Add(gcol_FeatureClasses.Item(lngItem - 1).strIndex)
            End If
        Next lngItem

        '/*Execute the removal
        For lngItem = 1 To colRemove.Count
            gcol_FeatureClasses.Remove(colRemove(lngItem))
        Next lngItem
    End Sub
    Public ReadOnly Property CauseEnabled() As Boolean
        Get
            Return m_bCauseEnabled
        End Get
    End Property

    ' Return the header collection for a pen
    Public ReadOnly Property DefaultPenHeader() As List(Of String)
        Get
            Return m_colfgFeatureHeader
        End Get
    End Property

    ' Builds the Feature Structure
    Public Function Test_BuildFeatureStructure() As Boolean
        Dim strFile As String = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "testfiles\feature.txt")

        If File.Exists(strFile) Then
            gcol_FeatureClasses = New colFeatureClasses()

            Using sr As New StreamReader(strFile)
                Dim strIn As String
                Dim strTopKey As String = ""
                Dim strMidKey As String = ""

                Do While Not sr.EndOfStream
                    strIn = sr.ReadLine()
                    If strIn.StartsWith("###") Then
                        Test_AddSubFeature(strIn, strTopKey, strMidKey)
                    ElseIf strIn.StartsWith("##") Then
                        Test_AddFeature(strIn, strTopKey, strMidKey)
                    ElseIf strIn.StartsWith("#") Then
                        Test_AddFeatureClass(strIn)
                        strTopKey = strIn
                    End If
                Loop
            End Using

            Configure_fgFeature()
            Test_BuildFeatureStructure = True
            m_bDataAvailable = True
        End If

        Return Test_BuildFeatureStructure
    End Function

    Private Sub Test_AddFeatureClass(ByRef strIn As String)
        Dim parts As String() = strIn.Substring(1).Split("|")
        If parts.Length < 3 Then Exit Sub

        Dim strTitle As String = parts(0)
        Dim strType As String = parts(1).ToUpper()
        Dim lngSeverity As Integer = Convert.ToInt32(parts(2))

        gcol_FeatureClasses.Add(strTitle, strType, lngSeverity, AddFeatureCol(), AddHotFeatureCol(), m_lngHotMax, "Pen", strTitle & "+" & strType)
    End Sub

    Private Sub Test_AddFeature(ByRef strIn As String, ByVal strTopKey As String, ByRef strKey As String)
        Dim parts As String() = strIn.Substring(2).Split("|")
        If parts.Length < 2 Then Exit Sub

        Dim strCode As String = parts(0)
        Dim strDesc As String = parts(1)
        Dim strURL As String = If(parts.Length > 2, parts(2), "")

        strKey = strCode & "+" & strDesc
        gcol_FeatureClasses.Item(strTopKey).colClassFeatures.Add(strDesc, strCode, strURL, AddSubFeatureCol(), strKey)
    End Sub

    Private Sub Test_AddSubFeature(ByVal strIn As String, ByVal strTopKey As String, ByVal strMidKey As String)
        Dim parts As String() = strIn.Substring(3).Split("|")
        If parts.Length < 2 Then Exit Sub

        Dim strCode As String = parts(0)
        Dim strDesc As String = parts(1)
        Dim strURL As String = If(parts.Length > 2, parts(2), "")

        gcol_FeatureClasses.Item(strTopKey).colClassFeatures.Item(strMidKey).colSub.Add(strDesc, strCode, strURL)
    End Sub

    Public Sub Configure_fgFeature()
        m_colfgFeatureHeader.Clear()
        m_colfgFeatureHeader.AddRange({"Primary", "Type", "Defect", "SubDefect"})

        If gcol_FeatureClasses.ToList().Any(Function(c) c.strClassType.ToUpper() = cstrClassExclude) Then
            m_bCauseEnabled = True
            m_colfgFeatureHeader.AddRange({"Cause", "SubCause"})
        End If

        m_colfgFeatureHeader.AddRange({"Comment", "Numeric"})
    End Sub

    Public Function AddFeatureCol() As colFeatures
        Return New colFeatures()
    End Function

    Public Function AddHotFeatureCol() As colHotFeatures
        Return New colHotFeatures()
    End Function

    Public Function AddSubFeatureCol() As colSubFeatures
        Return New colSubFeatures()
    End Function

    Public Function IsExcludedClass(ByVal nClassIndex As Integer) As Boolean
        Return gcol_FeatureClasses.Item(nClassIndex).strClassType.ToUpper() = cstrClassExclude
    End Function

    Public Function ConvertFeatureCoordToWord(ByVal nClass As Integer, ByVal nFeature As Integer, Optional ByVal nSubFeature As Integer = -1) As String
        Dim strOut As String = String.Empty

        Try
            '/*Set a reference to the Feature
            With gcol_FeatureClasses.Item(nClass).colClassFeatures.Item(nFeature)
                '/*Cat the Feature description
                strOut = .strDesc
                '/*If there is a SubFeature
                If nSubFeature >= 0 Then
                    '/*Cat the SubFeature Description
                    strOut &= "->" & .colSub.Item(nSubFeature).strDesc
                End If
            End With
            '/*Return the string
            Return strOut
        Catch ex As Exception
            ' handle error
            Return String.Empty
        End Try
    End Function

    Public Function ValidateNumericInput(ByRef strIn As String) As Boolean
        Return IsNumeric(strIn)
    End Function

    Public Sub LookUpHotItemCoord(ByRef nClass As Integer, ByRef nFeature As Integer, nSubFeature As Integer)
        Try
            '/*Set a reference to the Feature
            With gcol_FeatureClasses.Item(nClass).colHotFeature.Item(nFeature)
                '/*Return the Feature
                nFeature = CInt(.lngFeature)
                '/*REturn the subfeature
                nSubFeature = CInt(.lngSubFeature)
            End With
            '/*Return success
            'LookUpHotItemCoord = True
        Catch ex As Exception

        End Try
    End Sub

    Public Function ClassCodeToIndex(ByVal strCode As String) As Integer
        Dim searchCode As String = strCode ' Ensure correct scoping
        Return gcol_FeatureClasses.ToList().FindIndex(Function(c) c.strTitle = searchCode)
    End Function

    Public Function FeatureCodeToIndex(ByVal nClassIndex As Integer, ByVal strCode As String) As Integer
        ' Ensure valid index to prevent runtime errors
        If nClassIndex < 0 OrElse nClassIndex >= gcol_FeatureClasses.Count Then
            Return -1 ' Return an error code if index is out of range
        End If

        Dim searchCode As String = strCode ' Ensure proper scoping

        Return gcol_FeatureClasses.Item(nClassIndex).colClassFeatures.ToList().FindIndex(Function(f) f.strCode = searchCode)
    End Function

    Public Function SubFeatureCodeToIndex(ByVal nClassIndex As Integer, ByVal nL1Index As Integer, ByVal strCode As String) As Integer
        Return gcol_FeatureClasses.ToList() _
    .Item(nClassIndex) _
    .colClassFeatures.ToList() _
    .Item(nL1Index) _
    .colSub.ToList() _
    .FindIndex(Function(s) s.strCode = strCode)

    End Function

    Public Function CauseCodeToIndex() As Integer
        Return gcol_FeatureClasses.ToList().FindIndex(Function(c) c.strTitle.ToUpper() = cstrClassExclude)
    End Function
    Public Sub FillDefectEditorListsDynamic()
        Dim yLabel As Integer = 20
        Dim yList As Integer = 50
        Dim xSpacing As Integer = 270
        Dim startX As Integer = 10

        Dim xPositions As New List(Of Integer)
        For i = 0 To gcol_FeatureClasses.Count - 1
            xPositions.Add(startX + i * xSpacing)
        Next

        ' Clear any existing controls and lists if needed
        'frmDefectEditor.lblList.Clear()
        'frmDefectEditor.lstItems.Clear()
        'frmDefectEditor.lstHotItems.Clear()



        ' Dynamically generate controls based on gcol_FeatureClasses
        For i As Integer = 0 To gcol_FeatureClasses.Count - 1
            Dim featureClass = gcol_FeatureClasses(i)
            Dim x = xPositions(i)

            ' Create and add Label
            Dim lbl As New Label With {
            .Text = featureClass.strTitle,
            .Location = New Point(x, yLabel),
            .AutoSize = True
        }
            frmDefectEditor.lblList.Add(lbl)
            frmDefectEditor.Controls.Add(lbl)

            ' Create and add ListBox for standard items
            Dim lst As New ListBox With {
            .Location = New Point(x, yList),
            .Size = New Size(250, 100)
}
            frmDefectEditor.lstItems.Add(lst)
            frmDefectEditor.Controls.Add(lst)

            ' Add feature descriptions to lst
            For j As Integer = 0 To featureClass.colClassFeatures.Count - 1
                lst.Items.Add(featureClass.colClassFeatures(j).strDesc)
            Next

            ' Create and add a hidden ListBox (hot items, if applicable)
            Dim lstHot As New ListBox With {
            .Location = New Point(x, yList - 60),
            .Size = New Size(250, 50),
            .Visible = False
        }
            frmDefectEditor.lstHotItems.Add(lstHot)
            frmDefectEditor.Controls.Add(lstHot)
        Next
    End Sub

    Public Sub FillDefectEditorLists()
        Dim nTopLevel As Integer
        Dim nMidLevel As Integer
        Dim nListIndex As Integer

        ' Loop through each main category of list objects
        For nTopLevel = 1 To gcol_FeatureClasses.Count
            ' Get the current control array upper limit
            nListIndex = nTopLevel - 1
            ' If we need more than the initial list, add another set of list objects to the screen
            If (nListIndex + 1) > frmDefectEditor.lblList.Count Then
                ' Make the object add call
                AddDefectEditorList()
            End If

            ' Set the list object title
            frmDefectEditor.lblList(nListIndex).Text = gcol_FeatureClasses.Item(nListIndex).strTitle

            ' Set a reference to the list we are working on
            With frmDefectEditor.lstItems(nListIndex)
                ' Add the actual descriptions to the List
                For nMidLevel = 0 To gcol_FeatureClasses.Item(nListIndex).colClassFeatures.Count - 1
                    ' Add the item to the list; be sure to set the index since this was experiencing a strange anomaly
                    ' that was causing overwriting of the 1st and 2nd list items
                    .Items.Add(gcol_FeatureClasses.Item(nListIndex).colClassFeatures.Item(nMidLevel).strDesc)
                Next nMidLevel
            End With
        Next nTopLevel
    End Sub
    Private Sub AddDefectEditorList()
        Dim nIndex As Integer

        ' Increment the index by +1
        nIndex = frmDefectEditor.lblList.Count
        '-------------------------------------
        ' Label of the Listing Type
        ' Create and add a new Label control
        Dim newLabel As New Label With {
            .Visible = True
        }
        frmDefectEditor.lblList.Add(newLabel)
        frmDefectEditor.Controls.Add(newLabel)
        '-------------------------------------
        ' The Hot list
        ' Create and add a new ListBox control for lstHotItems
        Dim newLstHotItems As New ListBox With {
            .Visible = True
        }
        ' Assuming lstHotItems is a List(Of ListBox)
        frmDefectEditor.lstHotItems.Add(newLstHotItems)
        frmDefectEditor.Controls.Add(newLstHotItems)
        '-------------------------------------
        ' The actual defect list
        ' Create and add a new ListBox control for lstItems
        Dim newLstItems As New ListBox With {
            .Visible = True
        }
        frmDefectEditor.lstItems.Add(newLstItems)
        frmDefectEditor.Controls.Add(newLstItems)
    End Sub

    '======================================================= 
    'Routine: mdlDefectEditor.MapInFeatureHeaders()
    'Purpose: This clears the current contents of the Flex
    '         Grid and resets the Fixed header row.
    '
    'Globals: None
    '
    'Input: Private m_colfgFeatureHeader - The header titles
    '
    'Return: None
    '
    'Modifications:
    '   10-07-1998 As written for Pass1.1
    '=======================================================

    Public Sub MapInFeatureHeaders(ByRef frmIn As Form, ByRef fgFlexGrid As DataGridView)
        ' Ensure m_colfgFeatureHeader and m_bDataAvailable are declared at module or class level

        ' Refresh the form in case it's not visible or not repainted
        frmIn.Refresh()
        ' Clear existing grid content
        fgFlexGrid.Rows.Clear()
        fgFlexGrid.Columns.Clear()

        ' Exit if no data available
        If Not m_bDataAvailable Then Exit Sub

        ' Create columns from header collection
        For lngLoop As Integer = 1 To m_colfgFeatureHeader.Count
            Dim headerText As String = m_colfgFeatureHeader.Item(lngLoop - 1).ToString()
            fgFlexGrid.Columns.Add($"col{lngLoop}", headerText)
            fgFlexGrid.Columns(lngLoop - 1).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
            fgFlexGrid.Columns(lngLoop - 1).DefaultCellStyle.BackColor = fgFlexGrid.ColumnHeadersDefaultCellStyle.BackColor
        Next

        ' set the columns Comment and Numeric readonly
        If fgFlexGrid.Columns.Count >= 6 Then ' 假设至少有6列（索引0-5）
            fgFlexGrid.Columns(0).ReadOnly = True
            fgFlexGrid.Columns(1).ReadOnly = True
            fgFlexGrid.Columns(2).ReadOnly = True
            fgFlexGrid.Columns(3).ReadOnly = True
            fgFlexGrid.Columns(4).ReadOnly = False ' Comment can edit
            fgFlexGrid.Columns(5).ReadOnly = False ' Numeric can edit
        Else
            Debug.WriteLine("warning：columns less than 6，can not set Comment/Numeric Readonly level")
        End If

        ' Optionally add an empty row if needed
        'fgFlexGrid.Rows.Add()
    End Sub

    Public Sub SeekLevelCodes(ByRef colDefects As clsDefect, ByVal nClassIndex As Integer)
        Dim nL1Index As Integer
        Dim nL2Index As Integer

        ' Set a reference to the colDefects being operated on
        With colDefects
            ' Attempt to convert the Level-1 description
            nL1Index = FeatureCodeToIndex(nClassIndex, .strFeatureCd)
            ' Check if a match was found
            If nL1Index >= 0 Then
                ' Map the known items
                .nFeatureIndex = nL1Index
                .strFeatureDesc = gcol_FeatureClasses(nClassIndex).colClassFeatures(nL1Index).strDesc

                ' Attempt to convert the SubItem
                If Not String.IsNullOrEmpty(.strSubFeatureCd) Then
                    nL2Index = SubFeatureCodeToIndex(nClassIndex, nL1Index, .strSubFeatureCd)
                    ' Map known items into the indexes
                    If nL2Index >= 0 Then
                        .nSubFeatureIndex = nL2Index
                        .strSubFeatureDesc = gcol_FeatureClasses(nClassIndex).colClassFeatures(nL1Index).colSub(nL2Index).strDesc
                    Else
                        .nSubFeatureIndex = -1
                    End If
                Else
                    .nSubFeatureIndex = -1
                End If
            Else
                .nFeatureIndex = -1 'pppp
                .nSubFeatureIndex = -1
            End If
        End With
    End Sub

    Public Sub SeekCauseCodes(ByRef oDefects As clsDefect, ByVal nClassIndex As Integer, ByRef nL1Index As Integer, ByRef nL2Index As Integer)
        ' Set a reference to the colDefects being operated on
        With oDefects
            ' Attempt to convert the Level-1 description
            nL1Index = FeatureCodeToIndex(nClassIndex, .strCauseCd)
            ' Check if a match was found
            If nL1Index >= 0 Then
                ' Map the known items
                .strCauseDesc = gcol_FeatureClasses(nClassIndex).colClassFeatures(nL1Index).strDesc

                ' Attempt to convert the SubItem
                If Not String.IsNullOrEmpty(.strSubFeatureCd) Then
                    nL2Index = SubFeatureCodeToIndex(nClassIndex, nL1Index, .strSubCauseCd)
                    ' Map known items into the indexes
                    If nL2Index >= 0 Then
                        .strSubCauseDesc = gcol_FeatureClasses(nClassIndex).colClassFeatures(nL1Index).colSub(nL2Index).strDesc
                    End If
                End If
            End If
        End With
    End Sub
    '
    '=======================================================
    'Routine: mdlDefectEditor.CreateUnit()
    'Purpose: Generate a new Defect Collection for the
    '         defect editor to add Features on to.
    '
    'Globals:None
    '
    'Input:None
    '
    'Return:None
    '
    '
    'Modifications:
    '   10-07-1998 As written for Pass1.1
    '
    '
    '=======================================================
    Public Function CreateUnit() As colDefect
        Dim oDefect As New colDefect()
        Return oDefect
    End Function
End Module
