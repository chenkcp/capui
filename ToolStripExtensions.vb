Imports System.Runtime.CompilerServices
Imports System.Windows.Forms

Module ToolStripExtensions
    <Extension()>
    Public Function Clone(ts As ToolStripSeparator) As ToolStripSeparator
        Dim sep As New ToolStripSeparator()
        sep.AutoSize = ts.AutoSize
        sep.Height = 20 ' Set to your desired height
        sep.Width = ts.Width
        sep.BackColor = ts.BackColor
        Return sep
    End Function
End Module