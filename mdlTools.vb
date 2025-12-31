Imports System.Globalization
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Windows.Forms

Module mdlTools
    ' Constants for disabling the close button
    Private Const SC_CLOSE As Integer = &HF060
    Private Const MF_BYCOMMAND As Integer = &H0
    Private Const MF_GRAYED As Integer = &H1 ' Grays out the menu item
    Private Const MAX_COMPUTERNAME_LENGTH As Integer = 15

    ' Import Windows API functions
    <DllImport("user32.dll", SetLastError:=True)>
    Private Function GetSystemMenu(hWnd As IntPtr, bRevert As Boolean) As IntPtr
    End Function

    <DllImport("user32.dll", SetLastError:=True)>
    Private Function EnableMenuItem(hMenu As IntPtr, uIDEnableItem As UInteger, uEnable As UInteger) As Integer
    End Function

    <DllImport("user32.dll", SetLastError:=True)>
    Private Function RemoveMenu(hMenu As IntPtr, uPosition As Integer, uFlags As Integer) As Boolean
    End Function

    <DllImport("user32.dll", SetLastError:=True)>
    Private Function DrawMenuBar(hWnd As IntPtr) As Boolean
    End Function
    ' Method to disable the Close (X) button
    Public Sub DisableCloseX(ByRef frmIn As Form)
        Try
            Dim hMenu As IntPtr = GetSystemMenu(frmIn.Handle, False)
            If hMenu <> IntPtr.Zero Then
                'RemoveMenu(hMenu, SC_CLOSE, MF_BYCOMMAND Or MF_GRAYED) ' Disables Close button
                EnableMenuItem(hMenu, SC_CLOSE, MF_BYCOMMAND Or MF_GRAYED)
                DrawMenuBar(frmIn.Handle) ' Refresh the menu bar to apply changes
            End If
        Catch ex As Exception
            MessageBox.Show("Error disabling close button: " & ex.Message)
        End Try
    End Sub

    Public Function atod(ByVal strIn As String) As Double
        Dim result As Double
        If Double.TryParse(strIn, result) Then
            Return result
        End If
        Return 0
    End Function

    Public Sub SetDataGridViewStyle(ByVal dgv As DataGridView)
        dgv.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        dgv.MultiSelect = False
    End Sub

    Public Function ReadProfile(strIniPath As String, strSection As String, strKey As String) As String
        Try
            ' Check if file exists
            If Not File.Exists(strIniPath) Then Return String.Empty

            ' Read all lines from the INI file
            Dim lines As String() = File.ReadAllLines(strIniPath, Encoding.Default)

            ' Find the section
            Dim inSection As Boolean = False
            For Each line As String In lines
                ' Trim whitespace and comments
                Dim trimmedLine As String = line.Trim()
                If trimmedLine.StartsWith(";") OrElse trimmedLine.StartsWith("#") Then Continue For

                ' Check for section header
                If trimmedLine.StartsWith("["c) AndAlso trimmedLine.EndsWith("]"c) Then
                    inSection = (trimmedLine.Substring(1, trimmedLine.Length - 2).Trim() = strSection)
                    If inSection Then Continue For
                End If

                ' If we're in the right section, look for the key
                If inSection Then
                    Dim keyValue As String() = line.Split(New Char() {"="c}, 2)
                    If keyValue.Length = 2 AndAlso keyValue(0).Trim() = strKey Then
                        Return keyValue(1).Trim()
                    End If
                End If
            Next

            ' Key not found
            Return String.Empty
        Catch ex As Exception
            ' Log error if needed
            Return String.Empty
        End Try
    End Function

    '============================================================
    'Routine: mdlTools.CurrentMachineName()
    'Purpose: This queries the machine name using Win32 API.
    '         More references to similar calls can be found
    '         at the Microsoft Knowledge Base in article
    '         Q148835 which lists API for WorkGroup, Domain etc.
    '
    'Globals:None
    '
    'Input:None
    '
    'Return: String - This returns the UNC for a PC.
    '
    'Tested:
    '   11-12-1998 Tested by hand. Chris Barker
    '
    'Modifications:
    '   11-12-1998 As written for Pass1.3
    '
    '
    '============================================================
    Public Function CurrentMachineName() As String
        Try
            ' 直接返回当前计算机名（推荐方式）
            'Return System.Environment.MachineName

            Return "M1LPQ-HESTIA-NX" 'Return "N-L2-Z1-NXTCAP"
        Catch ex As Exception
            ' 处理可能的异常（如权限不足）
            MainErrorHandler("CurrentMachineName", ex.Message)
            Return String.Empty
        End Try
    End Function

    Public Function CreateAutoObject(
    ByVal strTypeName As String,
    Optional ByVal strMachineName As String = "",
    Optional ByVal strUser As String = "",
    Optional ByVal strDomain As String = "",
    Optional ByVal strPassword As String = ""
) As Object
        Try
            LogEvent("Attempting to create object: " & strTypeName)

            ' Only strMachineName is not the current running machine else we are creating object in remote machines
            'If Not (String.IsNullOrEmpty(strMachineName) AndAlso strMachineName.Equals(Environment.MachineName, StringComparison.OrdinalIgnoreCase)) Then

            '    Throw New NotSupportedException("Remote instantiation is not supported. Only local creation is allowed.")
            'End If

            'Attempt to resolve the Type from the string
            Dim t As Type = Type.GetType(strTypeName, throwOnError:=True)
            If t Is Nothing Then
                Throw New TypeLoadException("Type not found: " & strTypeName)
            Else
                Debug.WriteLine(t.AssemblyQualifiedName)
            End If
            'Create a new instance using the default (parameterless) constructor
            go_BusinessServer = Activator.CreateInstance(t)


            Return go_BusinessServer

        Catch ex As Exception
            LogEvent("CreateAutoObject error: " & ex.Message)
            MessageBox.Show("Automation Error:" & vbCrLf & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return Nothing
        End Try
    End Function

    '========================================================
    'Routine: LogToFile(filename,msg)
    'Purpose: This logs a string to a file and insures that
    'the file size does not exceed 20k.
    '
    'Globals:None
    '
    'Input: sFile - the file name to log to
    '       sMsg  - the string to write to the file
    '
    'Return:None
    '
    'Tested:
    '
    'Modifications:
    '   01-24-2000 As written
    '
    '
    '=======================================================
    Public Sub LogToFile(sFile As String, sMsg As String)
        Const nFileMax As Integer = 20000
        Dim sDir As String

        Try
            ' Determine full file path
            sDir = Path.Combine(Application.StartupPath, sFile)

            ' If file exists and is too big, delete it
            If File.Exists(sDir) AndAlso New FileInfo(sDir).Length > nFileMax Then
                File.Delete(sDir)
            End If

            ' Append message to file
            Using writer As StreamWriter = File.AppendText(sDir)
                writer.WriteLine($"{Now:MM-dd HH:mm} {sMsg}")
            End Using

        Catch ex As Exception
            ' Optionally log or suppress the error
            ' You can write to EventLog or Debug here
        End Try
    End Sub
    Public Sub ShowSetFocus(frm As Form)
        If frm IsNot Nothing Then
            frm.Show()
            frm.BringToFront()
            frm.Focus()
        End If
    End Sub
    Public Function StartProcess(ByVal strCmdLine As String, ByVal lngSecondsWait As Integer) As Boolean
        Try
            Dim psi As New ProcessStartInfo()
            psi.FileName = strCmdLine
            psi.UseShellExecute = True

            Using proc As Process = Process.Start(psi)
                If proc Is Nothing Then Return False

                ' Wait for exit or timeout
                Dim timeoutMs As Integer = lngSecondsWait * 1000
                If proc.WaitForExit(timeoutMs) Then
                    Return True ' Process finished within timeout
                Else
                    Try
                        proc.Kill()
                    Catch
                        ' Ignore errors if process already exited
                    End Try
                    Return False ' Timeout occurred
                End If
            End Using
        Catch ex As Exception
            ' Optionally log error here
            Return False
        End Try
    End Function
    ' Common date-time format
    Public Sub DisplayCurrentCultureInfo()
        Dim currentCulture As CultureInfo = CultureInfo.CurrentCulture
        Dim cultureInfoDetails As String = $"Current Culture: {currentCulture.Name}" & vbCrLf &
                                       $"Display Name: {currentCulture.DisplayName}" & vbCrLf &
                                       $"Date Format: {currentCulture.DateTimeFormat.ShortDatePattern}" & vbCrLf &
                                       $"Time Format: {currentCulture.DateTimeFormat.ShortTimePattern}" & vbCrLf &
                                       $"Number Format: {currentCulture.NumberFormat.CurrencySymbol} {currentCulture.NumberFormat.CurrencyDecimalSeparator}" & vbCrLf &
                                       $"Decimal Separator: {currentCulture.NumberFormat.NumberDecimalSeparator}" & vbCrLf &
                                       $"List Separator: {currentCulture.TextInfo.ListSeparator}"

        MessageBox.Show(cultureInfoDetails, "Current Culture Info", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub
    Public Function FormatForDatabase(ByVal dateValue As DateTime) As String
        ' get the day and month based on region setting. en-SG has d/M/yyyy h:mm:ss tt
        Dim swappedDate As New DateTime(dateValue.Year, dateValue.Day, dateValue.Month, dateValue.Hour, dateValue.Minute, dateValue.Second)

        'Convert to Access-compatible date format (MM/dd/yyyy hh:mm : ss tt)
        'Return interpretedDate.ToString("MM/dd/yyyy h:mm:ss tt", System.Globalization.CultureInfo.InvariantCulture)
        Return swappedDate.ToString("M/d/yyyy h:mm:ss tt", CultureInfo.InvariantCulture)
    End Function

    Public Function FormatFromDatabase(ByVal dateString As DateTime) As DateTime
        Try

            Dim DateValue As String = (dateString).ToString("M/d/yyyy h:mm:ss tt", System.Globalization.CultureInfo.InvariantCulture)

            Return DateTime.ParseExact(DateValue, "d/M/yyyy h:mm:ss tt", System.Globalization.CultureInfo.InvariantCulture)
        Catch ex As Exception
            ' Log error and handle invalid format
            Debug.WriteLine($"Error parsing date: {dateString}. Exception: {ex.Message}")
            Return DateTime.MinValue ' Return a default value or handle appropriately
        End Try
    End Function
    Public Function ParseDateTime(ByVal dateString As String) As Boolean
        Dim resultDateTime As DateTime
        If Not DateTime.TryParse(dateString, CultureInfo.CurrentCulture, DateTimeStyles.None, resultDateTime) Then
            Throw New ArgumentException("Invalid date format.")
        End If
        Debug.WriteLine($"Parsed date: {resultDateTime.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture)}")
        Return True
    End Function
End Module
