Public Class CAPmain
    ' Returns a shared UnitId input object, creating it if necessary
    Public ReadOnly Property NextCapIdInput As clsUnitId
        Get
            If go_SharedIdInput Is Nothing Then
                go_SharedIdInput = New clsUnitId()
            End If
            Return go_SharedIdInput
        End Get
    End Property

    ' Returns the global Context object
    Public Function GetContext() As Context
        If go_Context IsNot Nothing Then
            Return go_Context
        Else
            Return Nothing
        End If
    End Function

    ' Returns the global SystemSettings object
    Public Function GetSystemSettings() As clsSystemSettings
        If go_clsSystemSettings IsNot Nothing Then
            Return go_clsSystemSettings
        Else
            Return Nothing
        End If
    End Function
End Class


