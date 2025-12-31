Imports System
Imports System.IO
Imports System.Runtime.InteropServices

Module mdlcommon
    ' 定义 OSVERSIONINFO 结构体
    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Ansi)>
    Public Structure OSVERSIONINFO
        Public dwOSVersionInfoSize As Integer
        Public dwMajorVersion As Integer
        Public dwMinorVersion As Integer
        Public dwBuildNumber As Integer
        Public dwPlatformId As Integer
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=128)>
        Public szCSDVersion As String
    End Structure

    ' 定义常量
    Public Const VER_PLATFORM_WIN32_NT As Integer = 2
    Public Const VER_PLATFORM_WIN32_WINDOWS As Integer = 1
    Public Const VER_PLATFORM_WIN32s As Integer = 0

    ' 声明 GetVersionEx 函数
    <DllImport("kernel32.dll", CharSet:=CharSet.Ansi)>
    Public Function GetVersionEx(ByRef lpVersionInformation As OSVERSIONINFO) As Integer
    End Function

    ' 记录事件到日志文件
    Public Sub LogEvent(strMsg As String)
        Try
            Dim logFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "logevent.txt")
            File.AppendAllText(logFilePath, $"{DateTime.Now.ToString()}{vbTab}{strMsg}{vbCrLf}")
        Catch ex As Exception
            ' 可以根据需要添加异常处理逻辑，例如记录错误日志等
        End Try
    End Sub

    ' 创建新的日志文件
    Public Sub CreateLogEvent()
        Try
            Dim logFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "logevent.txt")
            If File.Exists(logFilePath) Then
                File.Delete(logFilePath)
            End If
            File.Create(logFilePath).Close()
        Catch ex As Exception
            ' 可以根据需要添加异常处理逻辑，例如记录错误日志等
        End Try
    End Sub
End Module
