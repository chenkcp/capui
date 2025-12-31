Public Module mdlUnitCapture

    ''' <summary>
    ''' Queries an external late-bound ActiveX server to get the pen id.
    ''' </summary>
    ''' <param name="oServer">The ActiveX server object (must have a GetId method).</param>
    ''' <returns>The pen ID as a string, or "ERROR" if an error occurs.</returns>
    Public Function GetIdFromServer(ByVal oServer As Object) As String
        Const cstrError As String = "ERROR"
        Try
            If oServer Is Nothing Then
                frmMessage.GenerateMessage("mdlUnitCaptureId()-Error#1: No ActiveX Server Available")
                Return cstrError
            Else
                Dim strResult As String = oServer.GetId("")
                If strResult IsNot Nothing AndAlso strResult.ToUpper().Contains(cstrError) Then
                    frmMessage.GenerateMessage("GetIdFromServer() Error#1" & vbCrLf & "ActiveX Id server reported error:" & vbCrLf & strResult)
                    Return cstrError
                Else
                    Return strResult
                End If
            End If
        Catch
            ' Optionally log the error here
            Return cstrError
        End Try
    End Function

End Module
