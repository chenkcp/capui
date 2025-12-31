' --- clsSAXhandler.vb ---
Public Class clsSAXhandler
    Public Property bLoaded As Boolean
    Public Property strType As String
    Public Property strCode As String
    Public Property strFileName As String
    Public Property oHandler As clsSAXEngineHandler
    Public Property strProtoType As String
    Public Property sProtoReturnType As String
    Public Property strLocation As String
    Public Property strIndex As String

    ' NEW: carry the parameters parsed from the prototype
    Public Property oParams As colProtoTypeParam  ' Of clsProtoTypeParam
End Class
