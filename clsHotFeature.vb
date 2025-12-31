Public Class clsHotFeature
    Private m_lngFeature As Long
    Private m_lngSubFeature As Long

    ' Property for lngSubFeature
    Public Property lngSubFeature As Long
        Get
            Return m_lngSubFeature
        End Get
        Set(ByVal value As Long)
            m_lngSubFeature = value
        End Set
    End Property

    ' Property for lngFeature
    Public Property lngFeature As Long
        Get
            Return m_lngFeature
        End Get
        Set(ByVal value As Long)
            m_lngFeature = value
        End Set
    End Property
End Class

