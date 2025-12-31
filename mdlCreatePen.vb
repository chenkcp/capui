Imports System.DirectoryServices.ActiveDirectory
Imports System.Globalization
Imports System.Runtime.InteropServices
Imports System.Windows.Forms

Module mdlCreatePen
    ' Constant to define server error command
    Private Const cServerError As String = "ERROR"
    Public frmUnitCaptureInstance As New frmUnitCapture()

    ' Structure for SYSTEMTIME (used in GetIdDttm)
    <StructLayout(LayoutKind.Sequential)>
    Public Structure SYSTEMTIME
        Public wYear As Short
        Public wMonth As Short
        Public wDayOfWeek As Short
        Public wDay As Short
        Public wHour As Short
        Public wMinute As Short
        Public wSecond As Short
        Public wMilliseconds As Short
    End Structure

    ' Declare external function to get system time
    <DllImport("kernel32.dll", SetLastError:=True)>
    Public Sub GetLocalTime(ByRef lpSystemTime As SYSTEMTIME)
    End Sub

    ' Sample count tracking
    Private m_oPenCounter As New clsSampleCount
    Private m_strLastUnitId As String = String.Empty

    ' Property to set the PenCounter object
    Public Property PenCounter() As clsSampleCount
        Get
            Return m_oPenCounter
        End Get
        Set(value As clsSampleCount)
            m_oPenCounter = value
        End Set
    End Property
    Public Function ReleaseUnit(ByRef oClsPen As clsPen) As Boolean
        Dim strGroupId As String
        Dim dtBirth As Date
        Dim strId As String

        ' Set the needed items
        With oClsPen
            strGroupId = .strGroupId
            dtBirth = .dtBirthDay
            strId = .strPenId
        End With

        ' Make sure we have somewhere to go with this
        If go_ActiveLotManager Is Nothing Then
            ' Do nothing for now
            Return False
        Else
            ' Send out the release request via Global LotManager object
            If Not go_ActiveLotManager.ReleasePen(strGroupId, dtBirth, strId) Then
                ' Wait before second release attempt
                Pause()

                ' Attempt the release again
                If go_ActiveLotManager.ReleasePen(strGroupId, dtBirth, strId) Then
                    ' Add the number of samples to the counter
                    If oClsPen.strPenId = m_strLastUnitId Then
                        AddSampleCount(oClsPen)
                        m_strLastUnitId = String.Empty
                    End If
                    ' Clear the pen and return success
                    ' oClsPen = Nothing ' Not needed in .NET (Garbage collection)
                    Return True
                End If
            Else
                ' Add the number of samples to the counter
                If oClsPen.strPenId = m_strLastUnitId Then
                    AddSampleCount(oClsPen)
                    m_strLastUnitId = String.Empty
                End If
                ' Clear the pen and return success
                ' oClsPen = Nothing ' Not needed in .NET (Garbage collection)
                Return True
            End If
        End If

        Return False
    End Function
    Public Function DeleteUnit(ByRef oClsPen As clsPen) As Boolean
        Dim strGroupId As String
        Dim dtBirth As Date
        Dim strId As String

        Try
            ' Set the needed items
            With oClsPen
                strGroupId = .strGroupId
                dtBirth = .dtBirthDay
                strId = .strPenId
            End With

            ' Make sure we have somewhere to go with this
            If go_ActiveLotManager Is Nothing Then
                ' Do nothing for now
                Return False
            Else
                ' Send out the delete request via Global LotManager object
                If Not go_ActiveLotManager.DeletePen(strGroupId, dtBirth, strId) Then
                    ' Wait
                    Pause()

                    ' Attempt it again
                    If go_ActiveLotManager.DeletePen(strGroupId, dtBirth, strId) Then
                        ' Add the number of samples to the counter
                        DeleteSampleCount(oClsPen)

                        ' Remove the unit id from the verification stack
                        frmUnitCaptureInstance.RemoveId(strId)

                        ' Clear the pen and return success
                        oClsPen = Nothing ' No need to use `Set` in .NET
                        Return True
                    End If
                Else
                    ' Add the number of samples to the counter
                    DeleteSampleCount(oClsPen)

                    ' Remove the unit id from the verification stack
                    frmUnitCaptureInstance.RemoveId(strId)

                    ' Clear the pen and return success
                    oClsPen = Nothing ' No need to use `Set` in .NET
                    Return True
                End If
            End If

        Catch ex As Exception
            ' Log the error
            Console.WriteLine("Error.txt", "mdlCreatePen.DeleteUnit(): " & ex.Message)
            ' mdlMain.MainErrorHandler("mdlCreatePen.DeleteUnit()", "Error: " & ex.Message)
        End Try

        Return False
    End Function

    ' Method to set standard properties of a Pen
    Public Sub SetUnitStdProperties(ByRef oClsPen As clsPen)
        With go_Context
            oClsPen.strOperator = .Operator
            oClsPen.strExperimentId = .ExperimentId
            oClsPen.strPartName = .PartName
            oClsPen.strPartNumber = .PartNumber
            oClsPen.strRunType = .RunType
            oClsPen.strShift = .Shift
            oClsPen.strThinFilmLotId = .ThinFilmLot
            oClsPen.strAccumulator = .Accumulator
            oClsPen.strLineType = .LineType
            oClsPen.strLineId = .LineNumber
            oClsPen.strSource = .Source
            oClsPen.dtInspectionDate = DateTime.Now
            oClsPen.strGroupId = .GroupId
            oClsPen.dtBirthDay = .GroupBirthDay
            oClsPen.strDisposition = "G"
            oClsPen.strUnitType = .ProductType
        End With
    End Sub

    ' Function to get formatted date-time string
    Public Function GetIdDttm() As String
        Dim sysTime As New SYSTEMTIME()
        GetLocalTime(sysTime)
        Return $"{sysTime.wYear}-{sysTime.wMonth}-{sysTime.wDay} {sysTime.wHour}:{sysTime.wMinute}:{sysTime.wSecond}.{sysTime.wMilliseconds}"
    End Function

    ' Function to pause execution
    Public Sub Pause()
        Dim waitUntil As DateTime = DateTime.Now.AddSeconds(4)
        Do While DateTime.Now < waitUntil
            Application.DoEvents()
        Loop
    End Sub

    ' Function to transmit a Pen object
    Public Function TransmitUnit(ByRef oClsPen As clsPen) As Boolean
        Try
            ' Test for a LotId
            If String.IsNullOrEmpty(oClsPen.strGroupId) AndAlso go_clsSystemSettings.strMaterialMode = mdlGlobal.gcstrLot Then
                Throw New Exception("Pen has no Lot ID")
                ' Test for Pen Id
            ElseIf String.IsNullOrEmpty(oClsPen.strPenId) Then
                Throw New Exception("Pen has no ID")
                ' Prop up the User ID
            ElseIf String.IsNullOrEmpty(oClsPen.strOperator) Then
                oClsPen.strOperator = "No Name"
            End If

            ' Pack the pen information into array form
            Dim vntArray As Object = mdlCreatePen.PackUnitArray(oClsPen)

            ' Make sure we have some where to go with this
            If go_ActiveLotManager Is Nothing Then
                ' Do nothing for now
            Else
                ' Send out pen data via Global LotManager object
                If Not go_ActiveLotManager.AddPen(vntArray) Then
                    ' Wait
                    Pause()
                    ' An error occured logging the pen;
                    ' attempt to deal with it
                    If go_ActiveLotManager.AddPen(vntArray) Then
                        ' Log the Id to the Context object
                        go_Context.LastUnitId = oClsPen.strPenId
                        ' Add the number of samples to the counter
                        AddSampleCount(oClsPen)
                        ' Add the unit id to the verification stack
                        'frmUnitCapture.TrackId(oClsPen.strPenId)
                        ' Execute post-entry CSBs
                        PostEntryCSB(oClsPen)
                        ' Clear the pen and return success
                        oClsPen = Nothing
                        Return True
                    End If
                Else
                    ' Log the Id to the Context object
                    go_Context.LastUnitId = oClsPen.strPenId
                    ' Add the number of samples to the counter
                    AddSampleCount(oClsPen)
                    ' Add the unit id to the verification stack
                    'frmUnitCapture.TrackId(oClsPen.strPenId)
                    ' Execute post-entry CSBs
                    PostEntryCSB(oClsPen)
                    ' Clear the pen and return success
                    oClsPen = Nothing
                    Return True
                End If
            End If
        Catch ex As Exception
            Return False
        End Try
        Return False
    End Function

    ' Function to retrieve a Pen object
    Public Function RetrievePen(ByVal strId As String, Optional ByVal dtBirthDay As Date = #1/1/1980#, Optional ByVal strGroupId As String = "CLIENT_UNKNOWN") As clsPen
        Dim oClsPen As New clsPen

        Try
            ' Ensure the ActiveLotManager is available
            If go_ActiveLotManager IsNot Nothing Then
                ' Normalize the group ID
                If strGroupId = "CLIENT_UNKNOWN" Then strGroupId = "UNKNOWN"

                ' Retrieve the pen data from the ActiveLotManager
                Dim vrtArray As Object = go_ActiveLotManager.GetPen(strGroupId, dtBirthDay, strId)

                ' Debugging: Log the retrieved array for inspection
                For i As Integer = 0 To vrtArray.GetUpperBound(0)
                    For j As Integer = 0 To vrtArray.GetUpperBound(1)
                        Debug.WriteLine($"vrtArray({i},{j}) = {vrtArray(i, j)}")
                    Next
                Next

                ' Check for server error and unpack the array if valid
                If vrtArray(0, 0).ToString() <> cServerError Then
                    oClsPen = UnPackUnitArray(vrtArray)
                    m_strLastUnitId = oClsPen.strPenId
                    Return oClsPen
                End If
            End If

        Catch ex As Exception
            ' Log the error for debugging purposes
            Debug.WriteLine($"Error in RetrievePen: {ex.Message}")
        End Try

        ' Return Nothing if retrieval fails
        Return Nothing
    End Function

    ' Function to pack a Pen object into an array
    Public Function PackUnitArray(oClsPen As clsPen) As Object
        Dim lngDefects As Integer
        Dim vrtArray As Object(,)
        Dim lngItem As Integer
        Dim lngDefectSet As Integer
        Const cnBASESIZE As Integer = 20
        Const cnITEMSperDEFECT As Integer = 9

        ' Determine the number of defects attached to the pen
        lngDefects = oClsPen.colPenDefects.Count

        ' Initialize the array dimensions
        ReDim vrtArray(cnBASESIZE + (lngDefects * cnITEMSperDEFECT), 0)

        '---------------------------------
        ' Pack the basic pen body
        '---------------------------------
        ' number of defects attached
        vrtArray(0, 0) = lngDefects
        vrtArray(1, 0) = oClsPen.strLineType 'Line Type
        vrtArray(2, 0) = oClsPen.strLineId 'Line Number
        vrtArray(3, 0) = oClsPen.strSource 'Station Source Code
        vrtArray(4, 0) = oClsPen.strGroupId 'The batch name
        vrtArray(5, 0) = oClsPen.dtBirthDay 'The date of birth for this unit
        vrtArray(6, 0) = oClsPen.strPenId 'The test/units unique id
        vrtArray(7, 0) = oClsPen.dtInspectionDate 'date this inspection was performed
        vrtArray(8, 0) = oClsPen.nCount 'units counted for this test
        vrtArray(9, 0) = oClsPen.strOperator 'user/operator conducting the test
        vrtArray(10, 0) = oClsPen.strShift 'shift that the test was performed on
        vrtArray(11, 0) = oClsPen.strDisposition 'disposition of the test
        vrtArray(12, 0) = oClsPen.strTestBed 'test bed used for the test (Number - Type)
        vrtArray(13, 0) = oClsPen.bPenNotShipped 'whether the pen was shippable
        vrtArray(14, 0) = oClsPen.strRecoveryStep 'the recovery step used to make a good pen
        vrtArray(15, 0) = oClsPen.strRunType 'The run mode during the test
        vrtArray(16, 0) = oClsPen.strExperimentId 'Id if this is an experiment
        vrtArray(17, 0) = oClsPen.strPartName 'the name of the part type
        vrtArray(18, 0) = oClsPen.strPartNumber 'the part number
        vrtArray(19, 0) = oClsPen.strUnitType 'type of unit/part
        vrtArray(20, 0) = oClsPen.strThinFilmLotId 'thinfilm batch id if known

        ' Add on the defects to the unit
        For lngItem = 1 To lngDefects
            ' Mark the new ubound for as an index to the defect fields
            lngDefectSet = cnBASESIZE + (lngItem * cnITEMSperDEFECT)
            ' Set a reference to the current defect
            Dim currentDefect = oClsPen.colPenDefects(lngItem)
            vrtArray(lngDefectSet - 8, 0) = lngItem 'Defect number
            vrtArray(lngDefectSet - 7, 0) = currentDefect.strFeatureClassDesc 'The class of the feature group
            vrtArray(lngDefectSet - 6, 0) = currentDefect.bPrimary 'is this the primary feature
            vrtArray(lngDefectSet - 5, 0) = currentDefect.strComment 'Any comment from the user
            vrtArray(lngDefectSet - 4, 0) = atod(currentDefect.strNumericInput) 'numeric value for this
            vrtArray(lngDefectSet - 3, 0) = currentDefect.strFeatureCd 'The code for the feature
            vrtArray(lngDefectSet - 2, 0) = currentDefect.strSubFeatureCd 'The sub feature code viewed
            vrtArray(lngDefectSet - 1, 0) = currentDefect.strCauseCd 'any specified cause for the feature
            vrtArray(lngDefectSet, 0) = currentDefect.strSubCauseCd 'an associated sub - cause for the feature
        Next

        ' 打印 vrtArray 的内容
        For i As Integer = 0 To vrtArray.GetUpperBound(0)
            For j As Integer = 0 To vrtArray.GetUpperBound(1)
                Debug.WriteLine(vrtArray(i, j))
            Next
            Console.WriteLine() ' 换行
        Next

        Return vrtArray
        '' Convert the 2D array to a 1D array for return
        'Dim resultArray As Object() = New Object(vrtArray.GetLength(0) - 1) {}
        'For i As Integer = 0 To vrtArray.GetLength(0) - 1
        '    resultArray(i) = vrtArray(i, 0)
        'Next

        '' Return the packed array
        'Return resultArray
    End Function

    ' Function to unpack an array into a Pen object
    ' Function to unpack an array into a Pen object
    Public Function UnPackUnitArray(ByVal vrtArray As Object(,)) As clsPen
        Dim oClsPen As clsPen = Nothing
        Dim nItems As Integer
        Dim nItem As Integer
        Dim nCheck As Integer
        Dim nUbound As Integer ' Upper bound of the 2nd dimension of the array

        '-- Translation variables for defects --
        Dim nOrder As Integer
        Dim bPrimary As Boolean
        Dim strFeatureCd As String
        Dim strSubFeatureCd As String
        Dim strCauseCd As String
        Dim strSubCauseCd As String
        Dim strFeatureClassDesc As String
        Dim strComment As String
        Dim strNumericInput As String

        ' Check if the first item in the array is not an error
        If vrtArray(0, 0) IsNot Nothing AndAlso vrtArray(0, 0).ToString() <> cServerError Then
            oClsPen = New clsPen()

            ' Get the defect count
            nItems = vtoi(vrtArray(0, 0))

            ' Verify the setting
            nCheck = (vrtArray.GetUpperBound(0) - 20) \ 9 ' Ensure integer division

            ' Exit if this appears to be corrupted
            If nCheck <> nItems Then
                Return Nothing
            End If

            ' Unload the basic 19 fields
            oClsPen.strLineType = vtoa(vrtArray(1, 0))
            oClsPen.strLineId = vtoa(vrtArray(2, 0))
            oClsPen.strSource = vtoa(vrtArray(3, 0))
            oClsPen.strGroupId = vtoa(vrtArray(4, 0))
            oClsPen.dtBirthDay = vrtArray(5, 0)
            oClsPen.strPenId = vtoa(vrtArray(6, 0))
            oClsPen.dtInspectionDate = vrtArray(7, 0)
            oClsPen.nCount = vtoi(vrtArray(8, 0))
            oClsPen.strOperator = vtoa(vrtArray(9, 0))
            oClsPen.strShift = vtoa(vrtArray(10, 0))
            oClsPen.strDisposition = vtoa(vrtArray(11, 0))
            oClsPen.strTestBed = vtoa(vrtArray(12, 0))
            oClsPen.bPenNotShipped = vtob(vrtArray(13, 0))
            oClsPen.strRecoveryStep = vtoa(vrtArray(14, 0))
            oClsPen.strRunType = vtoa(vrtArray(15, 0))
            oClsPen.strExperimentId = vtoa(vrtArray(16, 0))
            oClsPen.strPartName = vtoa(vrtArray(17, 0))
            oClsPen.strPartNumber = vtoa(vrtArray(18, 0))
            oClsPen.strUnitType = vtoa(vrtArray(19, 0))
            oClsPen.strThinFilmLotId = vtoa(vrtArray(20, 0))

            ' Set the variable for the upper bound
            nUbound = vrtArray.GetUpperBound(0)

            ' Unload the defect section
            nItem = 21
            While nItem < nUbound
                nOrder = vtoi(vrtArray(nItem, 0))
                strFeatureClassDesc = vtoa(vrtArray(nItem + 1, 0))
                bPrimary = vtob(vrtArray(nItem + 2, 0))
                strComment = vtoa(vrtArray(nItem + 3, 0))
                strNumericInput = vtoa(vrtArray(nItem + 4, 0))
                strFeatureCd = vtoa(vrtArray(nItem + 5, 0))
                strSubFeatureCd = vtoa(vrtArray(nItem + 6, 0))
                strCauseCd = vtoa(vrtArray(nItem + 7, 0))
                strSubCauseCd = vtoa(vrtArray(nItem + 8, 0))



                ' Add the observed feature to the pen collection
                oClsPen.colPenDefects.Add(0, bPrimary, 0, 0, 0,
                          strFeatureCd, "", strSubFeatureCd, "",
                          strCauseCd, "", strSubCauseCd, "",
                           "", strFeatureClassDesc, strComment, strNumericInput,
                          nOrder)


                ' Increment the array pointer by +9
                nItem += 9
            End While
        End If

        Return oClsPen
    End Function

    ' Function to add sample count
    Public Sub AddSampleCount(ByVal oUnit As clsPen)
        If oUnit IsNot Nothing Then
            If oUnit.colPenDefects.Count > 0 Then
                m_oPenCounter.AddBad(oUnit.nCount)
            Else
                m_oPenCounter.AddGood(oUnit.nCount)
            End If
        End If
    End Sub

    ' Function to delete sample count
    Public Sub DeleteSampleCount(ByVal oUnit As clsPen)
        If oUnit IsNot Nothing Then
            If oUnit.colPenDefects.Count > 0 Then
                m_oPenCounter.SubtractBad(oUnit.nCount)
            Else
                m_oPenCounter.SubtractGood(oUnit.nCount)
            End If
        End If
    End Sub

    Public Function UpdateUnit(ByRef oClsPen As clsPen) As Boolean
        Dim vntArray As Object

        Try
            ' Pack the pen information into array form
            vntArray = mdlCreatePen.PackUnitArray(oClsPen)

            ' Make sure we have some where to go with this
            If go_ActiveLotManager IsNot Nothing Then
                ' Send out pen data via Global LotManager object
                If Not (go_ActiveLotManager.UpdatePen(vntArray)) Then
                    ' Wait
                    Pause()
                    ' An error occured logging the pen; attempt to deal with it
                    If go_ActiveLotManager.UpdatePen(vntArray) Then
                        ' Log the Id to the Context object
                        go_Context.LastUnitId = oClsPen.strPenId
                        ' Add the number of samples to the counter
                        AddSampleCount(oClsPen)
                        ' Execute post-entry CSBs
                        PostEntryCSB(oClsPen)
                        ' Release the object
                        oClsPen = Nothing
                        ' Return success
                        Return True
                    End If
                Else
                    ' Log the Id to the Context object
                    go_Context.LastUnitId = oClsPen.strPenId
                    ' Add the number of samples to the counter
                    AddSampleCount(oClsPen)
                    ' Execute post-entry CSBs
                    PostEntryCSB(oClsPen)
                    ' Release the object
                    oClsPen = Nothing
                    ' Return success
                    Return True
                End If
            End If
        Catch ex As Exception
            'LogToFile("Error.txt", "mdlCreatePen.UpdateUnit(): " & ex.HResult.ToString() & " " & ex.Message)
            'mdlMain.MainErrorHandler("mdlCreatePen.UpdateUnit()", "Error: " & ex.Message & " " & ex.HResult.ToString())
            MessageBox.Show("Error: mdlCreatePen.UpdateUnit()")
        End Try

        Return False
    End Function

    Public Function UpdateRemoteUnit(ByRef oClsPen As clsPen) As Boolean
        Dim vntArray As Object
        Try
            ' Pack the pen information into array form
            vntArray = mdlCreatePen.PackUnitArray(oClsPen)

            ' Make sure we have somewhere to go with this
            If go_ActiveLotManager IsNot Nothing Then
                ' Send out pen data via Global LotManager object
                If Not go_ActiveLotManager.UpdatePenRemote(vntArray) Then
                    ' Wait
                    Pause()
                    ' An error occurred logging the pen; attempt to deal with it
                    If go_ActiveLotManager.UpdatePenRemote(vntArray) Then
                        ' Log the Id to the Context object
                        go_Context.LastUnitId = oClsPen.strPenId
                        ' Add the number of samples to the counter
                        AddSampleCount(oClsPen)
                        ' Execute post-entry CSBs
                        PostEntryCSB(oClsPen)
                        ' Release the object
                        oClsPen = Nothing
                        ' Return success
                        Return True
                    End If
                Else
                    ' Log the Id to the Context object
                    go_Context.LastUnitId = oClsPen.strPenId
                    ' Add the number of samples to the counter
                    AddSampleCount(oClsPen)
                    ' Execute post-entry CSBs
                    PostEntryCSB(oClsPen)
                    ' Release the object
                    oClsPen = Nothing
                    ' Return success
                    Return True
                End If
            End If
        Catch ex As Exception
            'LogToFile("Error.txt", $"mdlCreatePen.UpdateRemoteUnit(): {ex.HResult} {ex.Message}")
            'mdlMain.MainErrorHandler("mdlCreatePen.UpdateRemoteUnit()", $"Error: {ex.Message} {ex.HResult}")
            MessageBox.Show("Error: mdlCreatePen.UpdateRemoteUnit()")
        End Try
        Return False
    End Function


    Private Sub PostEntryCSB(oUnit As clsPen)
        ' Ensure that this is actually a Unit
        If oUnit IsNot Nothing Then
            ' Send a signal to increment the Unit Counter
            If oUnit.colPenDefects.Count > 0 Then
                'mdlSAX.ExecuteCSB_PostBadPenEntry()
            Else
                'mdlSAX.ExecuteCSB_PostGoodPenEntry()
            End If
        End If
    End Sub

    Public Function Undo() As clsPen
        Dim oClsPen As clsPen = Nothing
        Dim vrtArray As Object = Nothing

        Try
            ' Ensure there is somewhere to go to
            If go_ActiveLotManager Is Nothing Then
                ' Do nothing for now
                Return Nothing
            End If

            ' Request the unit from the business server
            vrtArray = go_ActiveLotManager.Undo()

            ' Ensure that there were no errors fetching the pen
            If Not vrtArray(0, 0).ToString().Equals("ERROR", StringComparison.OrdinalIgnoreCase) Then
                ' See if the return is a message or Unit
                If UBound(vrtArray, 1) = 0 Then
                    ' Show a message
                    frmMessage.GenerateMessage(vtoa(vrtArray(0, 1)), "Undo Successful")
                    Return Nothing
                Else
                    ' Unpack the retrieved pen array
                    oClsPen = UnPackUnitArray(vrtArray)
                    ' Decrement the count of the Sample Count object
                    ' (DeleteSampleCount is intentionally not called here, see original comment)
                    ' Set the last pen retrieved object
                    m_strLastUnitId = oClsPen.strPenId
                    Return oClsPen
                End If
            Else
                mdlMain.MainErrorHandler("mdlCreatePen.Undo()", "BusinessServer Error:" & vtoa(vrtArray(0, 1)))
                Return Nothing
            End If
        Catch ex As Exception
            LogToFile("Error.txt", "mdlCreatePen.Undo():" & ex.Message)
            mdlMain.MainErrorHandler("mdlCreatePen.Undo()", "Error:" & ex.Message)
            Return Nothing
        End Try
    End Function

    Public Function RetrieveRemotePen(ByVal strId As String) As clsPen
        Dim oClsPen As clsPen = Nothing
        Dim vrtArray As Object = Nothing

        If go_ActiveLotManager Is Nothing Then
            ' Do nothing for now
            Return Nothing
        Else
            ' Request the unit from the business server
            vrtArray = go_ActiveLotManager.GetPenRemote(strId)

            ' Ensure that there were no errors fetching the pen
            If Not vrtArray(0, 0).ToString().Equals(cServerError, StringComparison.OrdinalIgnoreCase) Then
                ' Unpack the retrieved pen array
                oClsPen = UnPackUnitArray(vrtArray)
                ' Decrement the count of the Sample Count object
                DeleteSampleCount(oClsPen)
                ' Set the last pen retrieved object
                m_strLastUnitId = oClsPen.strPenId
                Return oClsPen
            Else
                mdlMain.MainErrorHandler("mdlCreatePen.RetrieveRemotePen", "BusinessServer Error:" & vtoa(vrtArray(0, 1)))
                Return Nothing
            End If
        End If
    End Function
End Module
