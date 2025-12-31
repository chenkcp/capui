Imports System

Public Class clsParameter

    ' Local variables to hold property values
    Private m_strTitle As String
    Private m_strValue As String

    Public Property Value As String
        Get
            Return m_strValue
        End Get
        Set(ByVal vData As String)
            m_strValue = vData
        End Set
    End Property

    Public Property Title As String
        Get
            Return m_strTitle
        End Get
        Set(ByVal vData As String)
            m_strTitle = vData
        End Set
    End Property

    Public Sub New()
        ' Default constructor
    End Sub
End Class
