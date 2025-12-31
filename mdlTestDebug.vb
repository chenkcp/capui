Imports System.IO
Module mdlTestDebug
    Public Sub FlushToFile(ByVal strFile As String, ByVal strInfo As String)
        Try
            ' 使用 File.AppendAllText 直接追加内容（线程安全）
            File.AppendAllText(strFile, strInfo & Environment.NewLine)

            ' 输出到调试窗口（等价于 VB6 的 Debug.Print）
            System.Diagnostics.Debug.WriteLine(strInfo)
        Catch ex As Exception
            ' 处理可能的异常（如文件权限问题）
            System.Diagnostics.Debug.WriteLine($"FlushToFile 错误: {ex.Message}")
        End Try
    End Sub
End Module
