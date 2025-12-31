Public Class Context
    ' Private variables
    Private m_LineType As String
    Private m_LineNumber As String
    Private m_Source As String
    Private m_Accumulator As String
    Private m_ProductionDate As Date
    Private m_Operator As String
    Private m_Shift As String
    Private m_RunType As String
    Private m_ExperimentId As String
    Private m_PartNumber As String
    Private m_PartName As String
    Private m_ThinFilmLot As String
    Private m_GroupId As String
    Private m_GroupBirthDay As Date
    Private m_bCSBUsed As Boolean
    Private m_DefaultUnitType As String
    Private m_ProductType As String
    Private m_LastUnitId As String

    ' Collections for synchronization
    Private m_colPartName As New List(Of String)
    Private m_colPartNumber As New List(Of String)
    Private m_colProductType As New List(Of String)

    ' Properties
    Public Property LastUnitId As String
        Get
            Return m_LastUnitId
        End Get
        Set(value As String)
            m_LastUnitId = value
        End Set
    End Property

    Public Property ProductType As String
        Get
            Return m_ProductType
        End Get
        Set(value As String)
            m_ProductType = value
        End Set
    End Property

    Public Property DefaultUnitType As String
        Get
            Return m_DefaultUnitType
        End Get
        Set(value As String)
            m_DefaultUnitType = value
        End Set
    End Property

    Public Property bCSBUsed As Boolean
        Get
            Return m_bCSBUsed
        End Get
        Set(value As Boolean)
            m_bCSBUsed = value
        End Set
    End Property

    Public Property GroupBirthDay As Date
        Get
            Return m_GroupBirthDay
        End Get
        Set(value As Date)
            m_GroupBirthDay = value
        End Set
    End Property

    Public Property GroupId As String
        Get
            Return m_GroupId
        End Get
        Set(value As String)
            m_GroupId = value
        End Set
    End Property

    Public Property ThinFilmLot As String
        Get
            Return m_ThinFilmLot
        End Get
        Set(value As String)
            m_ThinFilmLot = value
        End Set
    End Property

    Public Property PartName As String
        Get
            Return m_PartName
        End Get
        Set(value As String)
            m_PartName = value
            ' Sync with part number and product type
            'Dim index As Integer = m_colPartName.IndexOf(value)
            'If index >= 0 Then
            '    m_PartNumber = m_colPartNumber(index)
            '    m_ProductType = m_colProductType(index)
            'End If

            ' Update status bar if applicable
            'If Not m_bCSBUsed Then
            mdlNextcap.UpdateStatusBar(value, 1)
            'End If
            '' Reset count if necessary
            'If go_clsSystemSettings.bResetCountOnProductChange Then
            '    go_SampleCount.ResetCount()
            'End If
        End Set
    End Property

    Public Property PartNumber As String
        Get
            Return m_PartNumber
        End Get
        Set(value As String)
            m_PartNumber = value
            ' Sync with part name and product type
            Dim index As Integer = m_colPartNumber.IndexOf(value)
            If index >= 0 Then
                m_PartName = m_colPartName(index)
                m_ProductType = m_colProductType(index)
            End If
        End Set
    End Property

    Public Property ExperimentId As String
        Get
            Return m_ExperimentId
        End Get
        Set(value As String)
            m_ExperimentId = value
        End Set
    End Property

    Public Property RunType As String
        Get
            Return m_RunType
        End Get
        Set(value As String)
            m_RunType = value
            'If Not m_bCSBUsed Then
            mdlNextcap.UpdateStatusBar(value, 6)
            'End If
        End Set
    End Property

    Public Property Shift As String
        Get
            Return m_Shift
        End Get
        Set(value As String)
            m_Shift = value
        End Set
    End Property

    Public Property [Operator] As String
        Get
            Return m_Operator
        End Get
        Set(value As String)
            m_Operator = value
            'If Not m_bCSBUsed Then
            mdlNextcap.UpdateStatusBar(value, 4)

            'End If
        End Set
    End Property

    Public Property ProductionDate As Date
        Get
            Return m_ProductionDate
        End Get
        Set(value As Date)
            m_ProductionDate = value
            ' Update Active LotManager
            If go_ActiveLotManager IsNot Nothing Then
                go_ActiveLotManager.ProductionDate = value
            End If
        End Set
    End Property

    Public Property Accumulator As String
        Get
            Return m_Accumulator
        End Get
        Set(value As String)
            m_Accumulator = value
        End Set
    End Property

    Public Property Source As String
        Get
            Return m_Source
        End Get
        Set(value As String)
            m_Source = value
            'If Not m_bCSBUsed Then
            mdlNextcap.UpdateStatusBar(value, 3)
            'End If
        End Set
    End Property

    Public Property LineNumber As String
        Get
            Return m_LineNumber
        End Get
        Set(value As String)
            m_LineNumber = value
            'If Not m_bCSBUsed Then
            mdlNextcap.UpdateStatusBar(m_LineType & value, 5)
            'End If
        End Set
    End Property

    Public Property LineType As String
        Get
            Return m_LineType
        End Get
        Set(value As String)
            m_LineType = value
            'If Not m_bCSBUsed Then
            mdlNextcap.UpdateStatusBar(value & m_LineNumber, 5)
            'End If
        End Set
    End Property

    ' AddParts function
    Public Sub AddParts(ByVal sName As String, ByVal sNumber As String, ByVal sFamily As String)
        Try
            ' Check if part already exists
            If Not m_colPartName.Contains(sName) OrElse Not m_colPartNumber.Contains(sNumber) OrElse Not m_colProductType.Contains(sFamily) Then
                m_colPartName.Add(sName)
                m_colPartNumber.Add(sNumber)
                m_colProductType.Add(sFamily)
            End If
        Catch ex As Exception
            ' Handle errors silently
        End Try
    End Sub
End Class
