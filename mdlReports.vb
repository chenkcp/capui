Module mdlReports
    Public Function CreateLotSummary(ByRef strGroupId As String, ByRef dtBirth As Date, Optional ByRef nStyle As Integer = 0) As clsReportData
        Dim oReportData As clsReportData

        '/*Generate the Report object
        oReportData = New clsReportData()
        '/*Set the Lot properties
        oReportData.GroupId = strGroupId
        oReportData.BirthDate = dtBirth

        Return oReportData
    End Function
End Module
