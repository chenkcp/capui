Imports System
Imports System.Collections.Generic
Imports System.Data

Public Class clsReportData
    Private m_colContext As Dictionary(Of String, colParameters)
    Private m_colContextIdx As List(Of String)
    Private m_colStrHeaderInfo As List(Of String)
    Private m_colStrCounts As colParameters
    Private m_colStrDefectCounts As colParameters
    Private m_colStrTimeDetail As colParameters
    Private m_colStrQualityState As colParameters
    Private m_colStrLotStatus As colParameters
    Private m_strGroupId As String
    Private m_dtBirth As Date
    Private m_strShift As String
    Private m_dtStart As Date
    Private m_dtEnd As Date
    Private m_bSuccess As Boolean
    Private m_vrtArrayData As DataTable
    Private m_vrtStrL1Descriptions As Object
    Private m_vrtStrL2Descriptions As Object

    Public Property ContextIdx As List(Of String)
        Get
            Return m_colContextIdx
        End Get
        Set(value As List(Of String))
            m_colContextIdx = value
        End Set
    End Property

    Public Property Context As Dictionary(Of String, colParameters)
        Get
            Return m_colContext
        End Get
        Set(value As Dictionary(Of String, colParameters))
            m_colContext = value
        End Set
    End Property

    Public Property Success As Boolean
        Get
            Return m_bSuccess
        End Get
        Set(value As Boolean)
            m_bSuccess = value
        End Set
    End Property

    'Public Function GetLotContexts(ByVal sLotId As String, ByVal dtBirthday As Date, ByVal colIdx As List(Of String)) As List(Of colParameters)
    '    Dim result As New List(Of colParameters)()
    '    Dim v As DataTable = go_ActiveLotManager.GetSummaryByContext(sLotId, dtBirthday)

    '    If v IsNot Nothing AndAlso v.Rows.Count > 0 Then
    '        Dim vTitles As List(Of String) = go_ActiveLotManager.GetSummaryHeader()

    '        For Each title As String In vTitles
    '            If title <> "Count" Then
    '                Dim colPS As New colParameters()
    '                For Each row As DataRow In v.Rows
    '                    Dim value As String = row(title).ToString()
    '                    If Not colPS.Contains(value) Then
    '                        colPS.Add(value, value)
    '                    End If
    '                Next
    '                result.Add(colPS)
    '                colIdx.Add(title)
    '            End If
    '        Next

    '        For Each row As DataRow In v.Rows
    '            Dim colPS As New colParameters()
    '            For Each title As String In vTitles
    '                colPS.Add(title, row(title).ToString())
    '            Next
    '            result.Add(colPS)
    '            colIdx.Add("Context " & (result.Count).ToString())
    '        Next
    '    End If

    '    Return result
    'End Function

    Public Property EndDate As Date
        Get
            Return m_dtEnd
        End Get
        Set(value As Date)
            m_dtEnd = value
        End Set
    End Property

    Public Property StartDate As Date
        Get
            Return m_dtStart
        End Get
        Set(value As Date)
            m_dtStart = value
        End Set
    End Property

    Public Property Shift As String
        Get
            Return m_strShift
        End Get
        Set(value As String)
            m_strShift = value
        End Set
    End Property

    Public Property BirthDate As Date
        Get
            Return m_dtBirth
        End Get
        Set(value As Date)
            m_dtBirth = value
        End Set
    End Property

    Public Property GroupId As String
        Get
            Return m_strGroupId
        End Get
        Set(value As String)
            m_strGroupId = value
        End Set
    End Property
    ' Converted col_strDefectCounts from VB6
    Public Property DefectCounts As colParameters
        Get
            Return m_colStrDefectCounts
        End Get
        Set(value As colParameters)
            m_colStrDefectCounts = value
        End Set
    End Property

    ' Converted col_strCounts from VB6
    Public Property Counts As colParameters
        Get
            Return m_colStrCounts
        End Get
        Set(value As colParameters)
            m_colStrCounts = value
        End Set
    End Property

    ' Converted vrtArrayData from VB6
    Public Property ArrayData As DataTable
        Get
            Return m_vrtArrayData
        End Get
        Set(value As DataTable)
            If value IsNot Nothing Then
                m_vrtArrayData = value
            Else
                m_vrtArrayData = New DataTable()
            End If
        End Set
    End Property

    ' Converted vrt_strL1Descriptions from VB6
    Public Property L1Descriptions As Object
        Get
            Return m_vrtStrL1Descriptions
        End Get
        Set(value As Object)
            m_vrtStrL1Descriptions = value
        End Set
    End Property

    ' Converted vrt_strL2Descriptions from VB6
    Public Property L2Descriptions As Object
        Get
            Return m_vrtStrL2Descriptions
        End Get
        Set(value As Object)
            m_vrtStrL2Descriptions = value
        End Set
    End Property

    ' Converted col_strHeaderInfo from VB6
    Public Property HeaderInfo As List(Of String)
        Get
            Return m_colStrHeaderInfo
        End Get
        Set(value As List(Of String))
            m_colStrHeaderInfo = value
        End Set
    End Property


    Public Sub New()
        m_colContext = New Dictionary(Of String, colParameters)()
        m_colContextIdx = New List(Of String)()
        m_colStrHeaderInfo = New List(Of String)()
        m_colStrCounts = New colParameters()
        m_colStrDefectCounts = New colParameters()
        m_colStrTimeDetail = New colParameters()
        m_colStrQualityState = New colParameters()
        m_colStrLotStatus = New colParameters()
        m_vrtArrayData = New DataTable()
    End Sub
End Class
