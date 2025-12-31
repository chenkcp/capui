Imports System
Imports System.Configuration
Imports System.Collections.Generic
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb
Public Class SiteAwareManager

    Private Shared _currentSite As String
    Private Shared _siteDetected As Boolean = False

    ''' <summary>
    ''' Gets the current site code
    ''' </summary>
    Public Shared ReadOnly Property CurrentSite As String
        Get
            Return ConfigurationManager.AppSettings("SiteCode")
        End Get
    End Property


    ''' <summary>
    ''' Gets site-specific connection string
    ''' </summary>
    Public Shared Function GetSiteConnectionString(databaseType As String) As String
        Try
            Dim connectionName = $"{CurrentSite}_{databaseType}"
            Dim connectionString = ConfigurationManager.ConnectionStrings(connectionName)

            If connectionString Is Nothing Then
                Throw New ConfigurationErrorsException($"Connection string '{connectionName}' not found for site '{CurrentSite}'")
            End If

            Console.WriteLine($"Retrieved connection: {connectionName}")
            Return connectionString.ConnectionString

        Catch ex As Exception
            Console.WriteLine($"Error getting site connection string: {ex.Message}")
            Throw
        End Try
    End Function

    ''' <summary>
    ''' Creates site-specific database connection
    ''' </summary>
    Public Shared Function CreateSiteConnection(databaseType As String) As IDbConnection
        Try
            Dim connectionString = GetSiteConnectionString(databaseType)
            Dim connectionName = $"{CurrentSite}_{databaseType}"
            Dim connInfo = ConfigurationManager.ConnectionStrings(connectionName)

            Select Case connInfo.ProviderName?.ToLowerInvariant()
                Case "system.data.sqlclient", Nothing
                    Return New SqlConnection(connectionString)
                Case "system.data.oledb"
                    Return New OleDbConnection(connectionString)
                Case Else
                    Return New SqlConnection(connectionString)
            End Select

        Catch ex As Exception
            Console.WriteLine($"Error creating site connection: {ex.Message}")
            Throw
        End Try
    End Function

    ''' <summary>
    ''' Lists all available connections for current site
    ''' </summary>
    Public Shared Function GetSiteConnections() As List(Of String)
        Dim siteConnections As New List(Of String)()
        Dim sitePrefix = $"{CurrentSite}_"

        For Each connString As ConnectionStringSettings In ConfigurationManager.ConnectionStrings
            If connString.Name.StartsWith(sitePrefix) Then
                siteConnections.Add(connString.Name)
            End If
        Next

        Return siteConnections
    End Function
End Class
