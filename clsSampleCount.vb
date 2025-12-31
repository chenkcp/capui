Public Class clsSampleCount
    ' Private storage for the current count
    Private m_nGoodCount As Integer
    Private m_nBadCount As Integer

    ' Event that count has changed
    Public Event NewCount(ByVal nGood As Integer, ByVal nBad As Integer)

    ' Add to the current good count, if it is about to overflow, reset the count.
    Public Sub AddGood(Optional ByVal nCount As Integer = 1)
        If m_nGoodCount < 32766 Then
            m_nGoodCount += nCount
        Else
            m_nGoodCount = 0
        End If
        RaiseEvent NewCount(m_nGoodCount, m_nBadCount)
    End Sub

    ' Add to the current bad count, if it is about to overflow, reset the count.
    Public Sub AddBad(Optional ByVal nCount As Integer = 1)
        If m_nBadCount < 32766 Then
            m_nBadCount += nCount
        Else
            m_nBadCount = 0
        End If
        RaiseEvent NewCount(m_nGoodCount, m_nBadCount)
    End Sub

    ' Subtract from the current good count, if it is less than zero, reset the count.
    Public Sub SubtractGood(Optional ByVal nCount As Integer = 1)
        If (m_nGoodCount - nCount) > 0 Then
            m_nGoodCount -= nCount
        Else
            m_nGoodCount = 0
        End If
        RaiseEvent NewCount(m_nGoodCount, m_nBadCount)
    End Sub

    ' Subtract from the current bad count, if it is less than zero, reset the count.
    Public Sub SubtractBad(Optional ByVal nCount As Integer = 1)
        If (m_nBadCount - nCount) > 0 Then
            m_nBadCount -= nCount
        Else
            m_nBadCount = 0
        End If
        RaiseEvent NewCount(m_nGoodCount, m_nBadCount)
    End Sub

    ' Clear the counters
    Public Sub ResetCount()
        m_nGoodCount = 0
        m_nBadCount = 0
        RaiseEvent NewCount(m_nGoodCount, m_nBadCount)
    End Sub

    ' Force NewCount Event
    Public Sub GetCount()
        RaiseEvent NewCount(m_nGoodCount, m_nBadCount)
    End Sub

    ' Get the value of the counter stored in the registry
    Public Sub New()
        m_nGoodCount = 20 ' Convert.ToInt32(Microsoft.VisualBasic.GetSetting(mdlMain.cstrRegName, "SampleCount", "Good", "0"))
        m_nBadCount = 0 'Convert.ToInt32(Microsoft.VisualBasic.GetSetting(mdlMain.cstrRegName, "SampleCount", "Bad", "0"))
    End Sub

    ' Save the current count of samples
    Protected Overrides Sub Finalize()
        Console.WriteLine(m_nGoodCount)
        Console.WriteLine(m_nBadCount)
        'Microsoft.VisualBasic.SaveSetting(mdlMain.cstrRegName, "SampleCount", "Good", m_nGoodCount.ToString())
        'Microsoft.VisualBasic.SaveSetting(mdlMain.cstrRegName, "SampleCount", "Bad", m_nBadCount.ToString())
        MyBase.Finalize()
    End Sub

End Class
