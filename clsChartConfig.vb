Public Class clsChartConfig
    Private mvarlngMaxGroups As Long
    Private mvarlngVisibleGroups As Long
    Private mvarcolCharts As colCharts

    Public Property ColCharts As colCharts
        Get
            Return mvarcolCharts
        End Get
        Set(value As colCharts)
            mvarcolCharts = value
        End Set
    End Property

    Public Property LngVisibleGroups As Long
        Get
            Return mvarlngVisibleGroups
        End Get
        Set(value As Long)
            mvarlngVisibleGroups = value
        End Set
    End Property

    Public Property LngMaxGroups As Long
        Get
            Return mvarlngMaxGroups
        End Get
        Set(value As Long)
            mvarlngMaxGroups = value
        End Set
    End Property

    ' Constructor
    Public Sub New()
        mvarlngMaxGroups = go_clsSystemSettings.lngMaxGroups
        mvarlngVisibleGroups = go_clsSystemSettings.lngMaxVisibleGroups
    End Sub


End Class

