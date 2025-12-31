Public Class clsUnitId
    ' Event to signal a Unit ID is ready for processing
    Public Event AsyncId(strId As String)
    ' Event to broadcast the Last Id entered
    Public Event LastIdState(strId As String, result As Boolean)

    ' Storage for an ID from a polled interface
    Private m_LastPolledId As String

    ' InputUnitId: OLE interface for other applications to write Unit Ids to NextCap
    Public Function InputUnitId(strData As String, strInputType As String) As Integer
        Try
            If mdlMain.ShutDownFlag Then
                Return 2
            ElseIf Not frmLotManager.ValidOpenLot Then
                Return 3
            End If

            If strInputType = "POLLED" Then
                m_LastPolledId = strData
                Return 1
            Else
                RaiseEvent AsyncId(strData)
                EnterGoodUnit(strData)
                Return 1
            End If
        Catch
            Return 0
        End Try
    End Function

    ' ReturnId: Entry point to return the result of the Last Id Acquired
    Friend Sub ReturnId(strId As String, result As Boolean)
        RaiseEvent LastIdState(strId, result)
    End Sub

    ' GetPolledId: Return the last entered Polled Id
    Public ReadOnly Property GetPolledId As String
        Get
            Return m_LastPolledId
        End Get
    End Property

    ' SubmitID: Expanded version of InputId for external applications
    Public Function SubmitID(strEntryMethod As String, strId As String, strDate As String, vrtCodes As Object) As Integer
        Const cstrAUTO As String = "AUTO"
        Const cstrDISPLAY As String = "DISPLAY"
        Try
            If mdlMain.ShutDownFlag Then
                Return 2
            ElseIf Not frmLotManager.ValidOpenLot Then
                Return 3
            End If

            If strEntryMethod = cstrDISPLAY Then
                mdlMain.frmNextCapInstance.txtUnitId.Text = strId
                Return 1
            ElseIf strEntryMethod = cstrAUTO Then
                If vrtCodes(0, 0) <> "--OK--" Then
                    If EnterBadUnit(strId, strDate, vrtCodes) Then Return 1
                Else
                    If EnterGoodUnit(strId, strDate) Then Return 1
                End If
            End If
            Return 0
        Catch
            Return 0
        End Try
    End Function

    ' EnterGoodUnit: Performs the Good pen transaction
    Private Function EnterGoodUnit(ByRef strUnitId As String, Optional ByRef strDate As String = "") As Boolean
        Try
            If mdlMain.ShutDownFlag OrElse Not frmLotManager.ValidOpenLot Then Return False

            strUnitId = mdlSAX.ExecuteCSB_AugmentPenId(strUnitId)

            '/*Execute the UnitId modification CSB
            If mdlSAX.ExecuteCSB_VerifyId(strUnitId) Then
                If String.IsNullOrEmpty(strUnitId) Then strUnitId = GetIdDttm()
                If frmUnitCapture.VerifyUnitId(strUnitId) Then
                    Dim oClsPen As New clsPen()
                    oClsPen.strPenId = strUnitId
                    oClsPen.nCount = 1
                    oClsPen.strRecoveryStep = frmEnterGoodPen.cmbRecovery.Text
                    oClsPen.strDisposition = "G"
                    mdlCreatePen.SetUnitStdProperties(oClsPen)
                    If Not String.IsNullOrEmpty(strDate) Then oClsPen.dtInspectionDate = vtoDate(strDate)
                    If mdlCreatePen.TransmitUnit(oClsPen) Then
                        Return True
                    End If
                End If
            End If
            Return False
        Catch
            Return False
        End Try
    End Function

    ' EnterBadUnit: Performs the Bad Unit entry transaction
    Private Function EnterBadUnit(ByRef strUnitId As String, ByRef strDate As String, ByRef vrtCodes As Object) As Boolean
        Try
            If mdlMain.ShutDownFlag OrElse Not frmLotManager.ValidOpenLot Then Return False

            strUnitId = mdlSAX.ExecuteCSB_AugmentPenId(strUnitId)
            If mdlSAX.ExecuteCSB_VerifyId(strUnitId) Then
                If String.IsNullOrEmpty(strUnitId) Then strUnitId = GetIdDttm()
                If frmUnitCapture.VerifyUnitId(strUnitId) Then
                    Dim oClsPen As New clsPen()
                    oClsPen.colPenDefects = mdlDefectEditor.CreateUnit()
                    oClsPen.strPenId = strUnitId
                    oClsPen.bPenNotShipped = False
                    oClsPen.nCount = 1
                    oClsPen.strDisposition = "G"
                    mdlCreatePen.SetUnitStdProperties(oClsPen)
                    oClsPen.dtInspectionDate = vtoDate(strDate)
                    If AddCodes(oClsPen, vrtCodes) Then
                        If mdlCreatePen.TransmitUnit(oClsPen) Then
                            Return True
                        End If
                    End If
                End If
            End If
            Return False
        Catch
            Return False
        End Try
    End Function

    ' AddCodes: Adds the Codes array to the Defect collection for the unit
    Private Function AddCodes(ByRef oClsPen As clsPen, ByRef vrtCodes As Object) As Boolean
        Try
            With oClsPen.colPenDefects
                For nItem As Integer = 0 To UBound(vrtCodes, 2)
                    Dim bPrimary As Boolean = (nItem = 0)
                    .Add(1, bPrimary, 0, 0, 0, vtoa(vrtCodes(1, nItem)), "", vtoa(vrtCodes(2, nItem)), "", "", "", "", "", vtoa(vrtCodes(0, nItem)), vtoa(vrtCodes(0, nItem)), "", "0")
                Next
            End With
            Return True
        Catch
            Return False
        End Try
    End Function
End Class
