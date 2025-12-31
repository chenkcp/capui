Imports System.IO


Public Class clsFeatureClass
    Public Property strTitle As String
    Public Property strClassType As String
    Public Property lngSeverity As Integer
    Public Property colClassFeatures As colFeatures
    Public Property colHotFeature As colHotFeatures
    Public Property lngHotMax As Integer
    Public Property strUnitType As String
    Public Property strIndex As String
    Public Property nColor As Integer

    Public Sub New()
        colClassFeatures = New colFeatures()
        colHotFeature = New colHotFeatures()
    End Sub
End Class
