Public Class clsProtoTypeParam
    Public ReadOnly Property sName As String
    Public ReadOnly Property sType As String
    Private _value As Object

    ' Property to get/set the value with type validation
    Public Property Value As Object
        Get
            Return _value
        End Get
        Set(value As Object)
            If ValidateType(value) Then
                _value = value
            Else
                Throw New ArgumentException($"Value type does not match expected type '{sType}'")
            End If
        End Set
    End Property

    ' Constructor with name and type string
    Public Sub New(name As String, type As String)
        sName = name
        sType = type.ToLowerInvariant()
        _value = Nothing
    End Sub

    ' Validate if a value matches the expected type
    Private Function ValidateType(value As Object) As Boolean
        If value Is Nothing Then Return True ' Allow null values
        Select Case sType
            Case "string"
                Return TypeOf value Is String
            Case "integer", "int"
                Return TypeOf value Is Integer
            Case "double"
                Return TypeOf value Is Double
            Case "boolean", "bool"
                Return TypeOf value Is Boolean
            Case "object"
                Return True
            Case "connectionstring", "sqlconnectionstring", "oledbconnectionstring", "odbcconnectionstring"
                Return TypeOf value Is IDbConnection  ' Connection strings are stored as strings
            Case Else
                Return True ' Unknown types are allowed
        End Select
    End Function

    ' Helper conversion setters
    Public Sub SetTypedValue(value As Object)
        Select Case sType
            Case "string"
                _value = value?.ToString()
            Case "integer", "int"
                If IsNumeric(value) Then
                    _value = CInt(value)
                Else
                    Throw New ArgumentException("Cannot convert value to Integer")
                End If
            Case "double"
                If IsNumeric(value) Then
                    _value = CDbl(value)
                Else
                    Throw New ArgumentException("Cannot convert value to Double")
                End If
            Case "boolean", "bool"
                _value = CBool(value)
             ' 🔑 Handle connection types properly
            Case "connection", "sqlconnection"
                If TypeOf value Is String Then
                    _value = New SqlClient.SqlConnection(value.ToString())
                ElseIf TypeOf value Is IDbConnection Then
                    _value = CType(value, IDbConnection)
                Else
                    Throw New ArgumentException("Value must be a connection string or IDbConnection")
                End If

            Case "oledbconnection"
                If TypeOf value Is String Then
                    _value = New OleDb.OleDbConnection(value.ToString())
                ElseIf TypeOf value Is IDbConnection Then
                    _value = CType(value, IDbConnection)
                End If

            Case "odbcconnection"
                If TypeOf value Is String Then
                    _value = New Odbc.OdbcConnection(value.ToString())
                ElseIf TypeOf value Is IDbConnection Then
                    _value = CType(value, IDbConnection)
                End If

            Case Else
                _value = value
        End Select
    End Sub

    ' Typed getters
    Public Function GetStringValue() As String
        Return _value?.ToString()
    End Function

    Public Function GetIntegerValue() As Integer?
        If IsNumeric(_value) Then Return CInt(_value)
        Return Nothing
    End Function

    Public Function GetDoubleValue() As Double?
        If IsNumeric(_value) Then Return CDbl(_value)
        Return Nothing
    End Function

    Public Function GetBooleanValue() As Boolean?
        If TypeOf _value Is Boolean Then Return CBool(_value) ' Fixed: was *value
        Return Nothing
    End Function

    ' Connection string getter
    Public Function GetConnectionString() As IDbConnection
        If sType.Contains("connectionstring") Then
            Return TryCast(_value, IDbConnection)
        End If
        Return Nothing
    End Function

    ' State check
    Public Function HasValue() As Boolean
        Return _value IsNot Nothing
    End Function

    Public Overrides Function ToString() As String
        Return $"Name: {sName}, Type: {sType}, Value: {If(_value, "Nothing")}"
    End Function
End Class