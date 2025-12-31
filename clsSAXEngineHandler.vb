Imports System.Data.Common
Imports System.Data.Odbc
Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports System.Security.Cryptography
Imports System.Text
Imports Microsoft.CodeAnalysis.CSharp.Scripting
Imports Microsoft.CodeAnalysis.Scripting
Imports System.Configuration

Public Class clsSAXEngineHandler
    Public Property FunctionName As String
    Public Property IsFunction As Boolean
    Public Property ActiveCode As String



    Public Sub New(sCode As String, functionName As String, isFunction As Boolean)

        Me.FunctionName = functionName
        Me.IsFunction = isFunction
        Me.ActiveCode = sCode
    End Sub


    ' Parameterless constructor
    Public Sub New()
        Me.IsFunction = True
    End Sub

    ' Evaluate function without parameters - directly runs the script
    Public Function Evaluate() As Object
        'If Not IsFunction Then Return Nothing
        Return RunScript(Nothing)
    End Function

    ' Evaluate function with Dictionary only (no handler parameters)
    Public Function Evaluate(vars As colProtoTypeParam) As Object
        'If Not IsFunction Then Return Nothing
        Return RunScript(vars)
    End Function

    ' Original Evaluate function that accepts Dictionary and handler parameters
    Public Function Evaluate(vars As colProtoTypeParam, handler_param As colProtoTypeParam) As Object
        If handler_param Is Nothing OrElse handler_param.Count = 0 Then
            ' No handler parameters to process, just run script with vars
            Return Evaluate(vars)
        End If

        If vars Is Nothing Then
            Throw New ApplicationException($"clsSAXEngineHandler: No vars provided, but handler_param has {handler_param.Count} items.")
        End If

        Dim varCount As Integer = vars.Count
        Dim handlerCount As Integer = handler_param.Count

        Try
            ' --- 1. Map overlapping items ---
            Dim i As Integer
            For i = 0 To Math.Min(varCount, handlerCount) - 1
                Dim varItem As clsProtoTypeParam = vars(i)
                Dim handlerItem As clsProtoTypeParam = handler_param(i)

                ' Map value from vars into handler_param
                handlerItem.Value = varItem.Value
            Next
            If varCount > handlerCount Then
                For j As Integer = handlerCount To varCount - 1
                    Dim varItem As clsProtoTypeParam = vars(j)

                    If varItem.sType.Contains("connection") Then
                        handler_param.Add(New clsProtoTypeParam(varItem.sName, varItem.sType))

                    End If
                Next
            End If

            ' --- 3. If there are more handler_params than vars, error out ---
            If handlerCount > varCount Then
                Throw New ApplicationException($"clsSAXEngineHandler: handler_param has {handlerCount - varCount} extra item(s) with no matching vars.")
            End If

            ' --- 4. Run the script with fully prepared handler_param ---
            Return RunScript(handler_param)

        Catch ex As Exception
            Debug.WriteLine($"Evaluate error: {ex}")
            Console.WriteLine($"Script execution error: {ex.Message}")
            Return Nothing
        End Try

    End Function



    ' Helper function to determine type string from object
    Private Function GetTypeString(value As Object) As String
        If value Is Nothing Then Return "Object"

        Select Case value.GetType().Name
            Case "String"
                Return "String"
            Case "Int32"
                Return "Integer"
            Case "Double"
                Return "Double"
            Case "Boolean"
                Return "Boolean"
            Case "DateTime"
                Return "DateTime"
            Case Else
                Return "Object"
        End Select
    End Function
    Private Function MakeValidIdentifier(name As String) As String
        If String.IsNullOrEmpty(name) Then Return "arg"
        Dim sb As New StringBuilder()
        Dim first As Char = name(0)
        If Char.IsLetter(first) OrElse first = "_"c Then
            sb.Append(first)
        Else
            sb.Append("_"c)
        End If
        For i As Integer = 1 To name.Length - 1
            Dim c As Char = name(i)
            If Char.IsLetterOrDigit(c) OrElse c = "_"c Then
                sb.Append(c)
            Else
                sb.Append("_"c)
            End If
        Next
        Return sb.ToString()
    End Function
    Private Function RunScript(handler_param As colProtoTypeParam) As Object
        Dim connections As New List(Of IDisposable)()

        Try
            ' Configure script options
            Dim options = ScriptOptions.Default _
            .AddReferences(
                GetType(Object).Assembly,
                GetType(System.Linq.Enumerable).Assembly,
                GetType(OdbcConnection).Assembly,
                GetType(SqlConnection).Assembly,
                GetType(OleDbConnection).Assembly,
                GetType(Microsoft.Win32.Registry).Assembly,
                GetType(MessageBox).Assembly,
                GetType(Form).Assembly,
                GetType(Dictionary(Of String, String)).Assembly
            ) _
            .AddImports("System", "System.Linq", "System.Data", "System.Data.Odbc",
                       "System.Data.SqlClient", "System.Data.OleDb", "Microsoft.Win32", "System.Windows.Forms", "System.Drawing", "System.Collections.Generic")

            ' Validate script code
            If String.IsNullOrWhiteSpace(ActiveCode) Then
                Debug.WriteLine("No script code available")
                Return Nothing
            End If

            For Each paramItem As clsProtoTypeParam In handler_param
                If paramItem IsNot Nothing AndAlso paramItem.sType IsNot Nothing AndAlso paramItem.sType.Contains("connection") Then
                    ' create the connection for this parameter
                    Dim connection As IDbConnection = CreateConnectionFromType(paramItem.sType, paramItem.sName)

                    If connection IsNot Nothing Then
                        If connection.State <> ConnectionState.Open Then
                            connection.Open()
                        End If

                        ' Replace the value in handler_param
                        handler_param.UpdateValue(paramItem.sName, connection)
                    End If
                End If
            Next
            ' Create globals object
            Dim globals = New ScriptGlobals(handler_param)
            ' Generate script with injected aliases
            Dim finalScript = GenerateScriptWithAliases(ActiveCode, handler_param)

            Debug.WriteLine($"Generated Script:{Environment.NewLine}{finalScript}")

            ' Execute the script
            Return CSharpScript.EvaluateAsync(Of Object)(finalScript, options, globals).GetAwaiter().GetResult()

        Catch ex As Exception
            Debug.WriteLine($"Script error: {ex}")
            Console.WriteLine($"Script execution error: {ex.Message}")
            Return Nothing
        Finally
            ' Ensure all connections are properly disposed
            For Each conn In connections
                Try
                    conn?.Dispose()
                Catch disposeEx As Exception
                    Debug.WriteLine($"Error disposing connection: {disposeEx.Message}")
                End Try
            Next
        End Try
    End Function
    ''' <summary>
    ''' Creates a database connection based on the parameter type and connection string
    ''' </summary>
    Private Function CreateConnectionFromType(paramType As String, connectionName As String) As IDbConnection
        Select Case paramType.ToLowerInvariant()
            Case "sqlconnection"
                Return New SqlConnection(GetSecureConnectionString(connectionName))
            Case "oledbconnection"
                Return New OleDbConnection($"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={ GetSecureFilePath(connectionName)};")
                'Case "odbcconnectionstring"
                '    Return New OdbcConnection(connectionString)
                'Case "connectionstring"
                '    ' Default to SqlConnection for generic connection string
                '    Return New SqlConnection(connectionString)
            Case Else
                Debug.WriteLine($"Unknown connection type: {paramType}")
                Return Nothing
        End Select
    End Function

    Private Function GetSecureConnectionString(connectionName As String) As String
        Try
            Dim CurrentSite As String = ConfigurationManager.AppSettings("SiteCode")
            connectionName = $"{CurrentSite}_{connectionName}"
            Dim configConnString = ConfigurationManager.ConnectionStrings(connectionName)?.ConnectionString
            If Not String.IsNullOrEmpty(configConnString) Then
                Return configConnString
            End If
        Catch ex As Exception
            Debug.WriteLine($"Error retrieving secure connection string for {connectionName}: {ex.Message}")
        End Try

        Return String.Empty
    End Function
    Private Function GetSecureFilePath(pathKey As String) As String
        Try
            ' Try app settings first
            Dim configPath = ConfigurationManager.AppSettings(pathKey)
            If Not String.IsNullOrEmpty(configPath) Then
                Return configPath
            End If

            ' Try environment variable
            Return Environment.GetEnvironmentVariable(pathKey)
        Catch ex As Exception
            Debug.WriteLine($"Error retrieving secure file path for {pathKey}: {ex.Message}")
            Return String.Empty
        End Try
    End Function

    Private Function GenerateScriptWithAliases(originalCode As String, vars As colProtoTypeParam) As String
        ' Find where to insert aliases (after using statements)
        Dim insertionPoint = FindEndOfUsingStatements(originalCode)

        Dim aliasLines As New StringBuilder()
        aliasLines.AppendLine("// <auto-injected-aliases>")

        ' Add parameter aliases
        If vars IsNot Nothing AndAlso vars.Count > 0 Then
            ' Only generates parameter aliases if vars exist
            For Each p As clsProtoTypeParam In vars
                Dim ident = MakeValidIdentifier(p.sName)
                'set the data type
                Dim typeLine = GenerateParameterAlias(ident, p.sName, p.sType)
                aliasLines.AppendLine(typeLine)
            Next
        End If

        ' Add database connection aliases
        'For Each kvp In connections
        '    Dim aliasName = $"{kvp.Key}Conn"
        '    aliasLines.AppendLine($"var {aliasName} = {kvp.Key}Connection;")
        'Next

        aliasLines.AppendLine("// </auto-injected-aliases>")

        Return aliasLines.ToString() & Environment.NewLine & originalCode
    End Function
    Private Function FindEndOfUsingStatements(code As String) As Integer
        Dim lines = code.Split({Environment.NewLine, vbNewLine, vbLf}, StringSplitOptions.None)
        Dim currentPosition = 0

        For i = 0 To lines.Length - 1
            Dim trimmedLine = lines(i).Trim()

            ' Skip empty lines and comments at the beginning
            If String.IsNullOrWhiteSpace(trimmedLine) OrElse trimmedLine.StartsWith("//") Then
                currentPosition += lines(i).Length + Environment.NewLine.Length
                Continue For
            End If

            ' If it's a using statement, move past it
            If trimmedLine.StartsWith("using ") AndAlso trimmedLine.EndsWith(";") Then
                currentPosition += lines(i).Length + Environment.NewLine.Length
                Continue For
            End If

            ' Found first non-using statement, return this position
            Exit For
        Next

        Return currentPosition
    End Function
    Private Function GenerateParameterAlias(identifier As String, paramName As String, paramType As String) As String
        Select Case paramType.ToLowerInvariant()
            Case "string"
                Return $"string {identifier} = GetVar(""{paramName}"") as string;"

            Case "integer", "int"
                Return $"int {identifier} = Convert.ToInt32(GetVar(""{paramName}""));"

            Case "double"
                Return $"double {identifier} = Convert.ToDouble(GetVar(""{paramName}""));"

            Case "boolean", "bool"
                Return $"bool {identifier} = Convert.ToBoolean(GetVar(""{paramName}""));"

            Case "connection", "sqlconnection", "oledbconnection", "odbcconnection", "idbconnection"
                Return $"System.Data.IDbConnection {identifier} = GetVar(""{paramName}"") as System.Data.IDbConnection;"

            Case Else
                Return $"var {identifier} = GetVar(""{paramName}"");"
        End Select
    End Function
End Class

'######################################
'######################################
' Enhanced ScriptGlobals class
'######################################
'######################################

Public Class ScriptGlobals
    Private ReadOnly _params As colProtoTypeParam
    Public Sub New(params As colProtoTypeParam)
        _params = If(params, New colProtoTypeParam())

    End Sub
    'only returns the value of the parameter; called by GenerateParameterAlias
    Public Function GetVar(name As String) As Object
        If _params Is Nothing Then Return Nothing

        Dim param As clsProtoTypeParam = Nothing
        If Not _params.TryGetValue(name, param) Then
            Return Nothing
        End If

        If param Is Nothing OrElse Not param.HasValue() Then
            Return Nothing
        End If

        Select Case param.sType.ToLowerInvariant()
            Case "string"
                Return param.GetStringValue()
            Case "integer", "int"
                Return param.GetIntegerValue()
            Case "double"
                Return param.GetDoubleValue()
            Case "boolean", "bool"
                Return param.GetBooleanValue()
            Case "connection", "sqlconnection", "oledbconnection", "odbcconnection", "idbconnection"
                ' ✅ Always return an IDbConnection for these types
                Return TryCast(param.Value, IDbConnection)
            Case Else
                ' Fallback: return raw object
                Return param.Value
        End Select
    End Function

    'returns the entire clsProtoTypeParam object.
    Public Function GetParam(name As String) As clsProtoTypeParam
        If _params Is Nothing Then Return Nothing

        Dim param As clsProtoTypeParam = Nothing
        Return If(_params.TryGetValue(name, param), param, Nothing)
    End Function

    Public Function RunQuery(conn As IDbConnection, sql As String) As DataTable
        If conn Is Nothing Then
            Throw New ArgumentNullException(NameOf(conn), "Database connection cannot be null.")
        End If

        If String.IsNullOrWhiteSpace(sql) Then
            Throw New ArgumentException("SQL query cannot be null or empty.", NameOf(sql))
        End If

        If conn.State <> ConnectionState.Open Then
            Throw New InvalidOperationException($"Connection is not open. Current state: {conn.State}")
        End If

        Dim dt As New DataTable()
        Try
            Using cmd As IDbCommand = conn.CreateCommand()
                cmd.CommandText = sql
                cmd.CommandTimeout = 30 ' Set reasonable timeout

                Using reader = cmd.ExecuteReader()
                    dt.Load(reader)
                End Using
            End Using
        Catch ex As Exception
            Throw New Exception($"Error executing query: {sql}", ex)
        End Try

        Return dt
    End Function

    Public Function RunNonQuery(conn As IDbConnection, sql As String) As Integer
        If conn Is Nothing Then
            Throw New ArgumentNullException(NameOf(conn), "Database connection cannot be null.")
        End If

        If String.IsNullOrWhiteSpace(sql) Then
            Throw New ArgumentException("SQL command cannot be null or empty.", NameOf(sql))
        End If

        If conn.State <> ConnectionState.Open Then
            Throw New InvalidOperationException($"Connection is not open. Current state: {conn.State}")
        End If

        Try
            Using cmd As IDbCommand = conn.CreateCommand()
                cmd.CommandText = sql
                cmd.CommandTimeout = 30
                Return cmd.ExecuteNonQuery()
            End Using
        Catch ex As Exception
            Throw New Exception($"Error executing non-query: {sql}", ex)
        End Try
    End Function
    ''' <summary>
    ''' Executes a parameterized query and returns a DataTable
    ''' </summary>
    ''' <param name="conn">Database connection</param>
    ''' <param name="sql">SQL query with parameter placeholders</param>
    ''' <param name="parameters">Dictionary of parameter names and values</param>
    ''' <returns>DataTable with query results</returns>
    Public Function RunParameterizedQuery(conn As IDbConnection, sql As String, parameters As Dictionary(Of String, Object)) As DataTable
        If conn Is Nothing Then
            Throw New ArgumentNullException(NameOf(conn), "Database connection cannot be null.")
        End If
        If String.IsNullOrWhiteSpace(sql) Then
            Throw New ArgumentException("SQL query cannot be null or empty.", NameOf(sql))
        End If
        If conn.State <> ConnectionState.Open Then
            Throw New InvalidOperationException($"Connection is not open. Current state: {conn.State}")
        End If

        Dim dt As New DataTable()
        Try
            Using cmd As IDbCommand = conn.CreateCommand()
                cmd.CommandText = sql
                cmd.CommandTimeout = 30

                ' Add parameters if provided
                If parameters IsNot Nothing Then
                    For Each kvp In parameters
                        Dim param As IDbDataParameter = cmd.CreateParameter()
                        param.ParameterName = kvp.Key
                        param.Value = If(kvp.Value, DBNull.Value)
                        cmd.Parameters.Add(param)
                    Next
                End If

                Using reader = cmd.ExecuteReader()
                    dt.Load(reader)
                End Using
            End Using
        Catch ex As Exception
            Throw New Exception($"Error executing parameterized query: {sql}", ex)
        End Try
        Return dt
    End Function

    ''' <summary>
    ''' Executes a parameterized non-query command (INSERT, UPDATE, DELETE)
    ''' </summary>
    ''' <param name="conn">Database connection</param>
    ''' <param name="sql">SQL command with parameter placeholders</param>
    ''' <param name="parameters">Dictionary of parameter names and values</param>
    ''' <returns>Number of affected rows</returns>
    Public Function RunParameterizedNonQuery(conn As IDbConnection, sql As String, parameters As Dictionary(Of String, Object)) As Integer
        If conn Is Nothing Then
            Throw New ArgumentNullException(NameOf(conn), "Database connection cannot be null.")
        End If
        If String.IsNullOrWhiteSpace(sql) Then
            Throw New ArgumentException("SQL command cannot be null or empty.", NameOf(sql))
        End If
        If conn.State <> ConnectionState.Open Then
            Throw New InvalidOperationException($"Connection is not open. Current state: {conn.State}")
        End If

        Try
            Using cmd As IDbCommand = conn.CreateCommand()
                cmd.CommandText = sql
                cmd.CommandTimeout = 30

                ' Add parameters if provided
                If parameters IsNot Nothing Then
                    For Each kvp In parameters
                        Dim param As IDbDataParameter = cmd.CreateParameter()
                        param.ParameterName = kvp.Key
                        param.Value = If(kvp.Value, DBNull.Value)
                        cmd.Parameters.Add(param)
                    Next
                End If

                Return cmd.ExecuteNonQuery()
            End Using
        Catch ex As Exception
            Throw New Exception($"Error executing parameterized non-query: {sql}", ex)
        End Try
    End Function

    ''' <summary>
    ''' Executes a parameterized scalar query (returns single value)
    ''' </summary>
    ''' <param name="conn">Database connection</param>
    ''' <param name="sql">SQL query with parameter placeholders</param>
    ''' <param name="parameters">Dictionary of parameter names and values</param>
    ''' <returns>Single scalar value or DBNull</returns>
    Public Function RunParameterizedScalar(conn As IDbConnection, sql As String, parameters As Dictionary(Of String, Object)) As Object
        If conn Is Nothing Then
            Throw New ArgumentNullException(NameOf(conn), "Database connection cannot be null.")
        End If
        If String.IsNullOrWhiteSpace(sql) Then
            Throw New ArgumentException("SQL command cannot be null or empty.", NameOf(sql))
        End If
        If conn.State <> ConnectionState.Open Then
            Throw New InvalidOperationException($"Connection is not open. Current state: {conn.State}")
        End If

        Try
            Using cmd As IDbCommand = conn.CreateCommand()
                cmd.CommandText = sql
                cmd.CommandTimeout = 30

                ' Add parameters if provided
                If parameters IsNot Nothing Then
                    For Each kvp In parameters
                        Dim param As IDbDataParameter = cmd.CreateParameter()
                        param.ParameterName = kvp.Key
                        param.Value = If(kvp.Value, DBNull.Value)
                        cmd.Parameters.Add(param)
                    Next
                End If

                Return cmd.ExecuteScalar()
            End Using
        Catch ex As Exception
            Throw New Exception($"Error executing parameterized scalar query: {sql}", ex)
        End Try
    End Function
    ' Add this simple InputBox method - accessible directly from C# scripts
    Public Function InputBox(prompt As String, title As String, Optional defaultText As String = "") As String
        Try
            Using form As New Form()
                Using label As New Label()
                    Using textBox As New TextBox()
                        Using buttonOk As New Button()
                            Using buttonCancel As New Button()
                                ' Form setup
                                form.Text = title
                                form.FormBorderStyle = FormBorderStyle.FixedDialog
                                form.StartPosition = FormStartPosition.CenterScreen
                                form.MaximizeBox = False
                                form.MinimizeBox = False
                                form.TopMost = True
                                form.Size = New Drawing.Size(400, 150)

                                ' Label setup
                                label.Text = prompt
                                label.Location = New Drawing.Point(12, 15)
                                label.Size = New Drawing.Size(360, 23)

                                ' TextBox setup
                                textBox.Text = defaultText
                                textBox.Location = New Drawing.Point(12, 40)
                                textBox.Size = New Drawing.Size(360, 23)

                                ' OK Button setup
                                buttonOk.Text = "OK"
                                buttonOk.DialogResult = DialogResult.OK
                                buttonOk.Location = New Drawing.Point(230, 75)
                                buttonOk.Size = New Drawing.Size(75, 25)

                                ' Cancel Button setup
                                buttonCancel.Text = "Cancel"
                                buttonCancel.DialogResult = DialogResult.Cancel
                                buttonCancel.Location = New Drawing.Point(315, 75)
                                buttonCancel.Size = New Drawing.Size(75, 25)

                                ' Set default and cancel buttons
                                form.AcceptButton = buttonOk
                                form.CancelButton = buttonCancel

                                ' Add controls to form
                                form.Controls.AddRange({label, textBox, buttonOk, buttonCancel})

                                ' Focus and select text
                                textBox.Select()
                                textBox.SelectAll()

                                ' Show dialog and return result
                                If form.ShowDialog() = DialogResult.OK Then
                                    Return textBox.Text
                                Else
                                    Return ""
                                End If
                            End Using
                        End Using
                    End Using
                End Using
            End Using
        Catch ex As Exception
            ' Fallback to simple message if InputBox fails
            MessageBox.Show($"InputBox error: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return defaultText
        End Try
    End Function
End Class