Public Module mdlNextcap
    '=================================================================
    'Routine:mdlNextCap.UpdateStatusBar
    'Purpose:Single point of entry for updating the panels
    '        of the status bar
    '
    'Globals:None
    '
    'Input:strText - The text to be placed in the panel
    '      nPanel  - The index of the panel to replace the text in.
    '
    'Return:None
    '
    'Modifications:
    '   9-21-1998 As written for Pass1.0
    '
    '   01-27-1999 Changed to allow for automaitc expansion of
    '   panel count if the requested panel is greater then the
    '   current upper bound.
    '
    '   08-03-2001 [Bug] - The panel expansion does not take into
    '   account a request for n more panels then current. I'm adding
    '   in a loop to expand by n panels rather than just 1.
    '==================================================================
    'Public Sub UpdateStatusBar(strText As String, nPanel As Integer)
    '    Try
    '        If mdlMain.frmNextCapInstance.sbrNextCap Is Nothing OrElse mdlMain.frmNextCapInstance.sbrNextCap.Items.Count = 0 Then
    '            Throw New InvalidOperationException("StatusStrip is not initialized.")
    '        End If

    '        Dim label As ToolStripLabel = CType(mdlMain.frmNextCapInstance.sbrNextCap.Items($"Panel{nPanel}"), ToolStripLabel)
    '        label.Text = strText



    '        '' Reference to the StatusStrip control
    '        'Dim statusStrip = frmNextCap.sbrNextCap
    '        'If statusStrip Is Nothing OrElse nPanel < 1 Then Return

    '        '' Ensure there are enough panels (ToolStripStatusLabels)
    '        'While nPanel > statusStrip.Items.Count
    '        '    Dim newLabel As New ToolStripLabel()
    '        '    newLabel.AutoSize = True
    '        '    statusStrip.Items.Add(newLabel)
    '        'End While

    '        '' Update the specified panel
    '        'Dim label As ToolStripLabel = CType(statusStrip.Items(nPanel - 1), ToolStripLabel)
    '        'label.Text = strText
    '        ''label.Spring = False ' Optional: use True to auto-fill remaining space
    '        'label.AutoSize = True

    '    Catch ex As Exception
    '        LogEvent($"Error: UpdateStatusBar() {ex.Message} Panel Request={nPanel} Count={frmNextCap.sbrNextCap.Items.Count}")
    '    End Try
    'End Sub
    Public Sub UpdateStatusBar(strText As String, nPanel As Integer)
        Try
            Dim statusStrip = mdlMain.frmNextCapInstance.sbrNextCap
            If statusStrip Is Nothing Then
                Throw New InvalidOperationException("StatusStrip is not initialized.")
            End If

            ' 正确计算当前面板总数：总项数 ÷ 2（因初始结构和新增规则确保总项数为偶数）
            ' 验证：12项 → 6个面板，14项 →7个面板，符合预期
            Dim currentPanelCount As Integer = statusStrip.Items.Count \ 2

            ' 校验面板编号合理性（1 ≤ nPanel ≤ 当前面板数 + 1）
            If nPanel < 1 OrElse nPanel > currentPanelCount + 1 Then
                Throw New ArgumentException($"Invalid panel number: {nPanel}. Valid range: 1 to {currentPanelCount + 1}")
            End If

            ' 目标面板索引公式（PanelN 位于 2n-1，验证：n=1→1，n=6→11，正确）
            Dim targetIndex As Integer = 2 * nPanel - 1

            ' 新增面板（当nPanel等于当前面板数+1时）
            If nPanel = currentPanelCount + 1 Then
                ' 添加分隔符（保持初始样式）
                Dim newSeparator As New ToolStripSeparator() With {
                .AutoSize = False,
                .Width = 8,
                .Height = 10
            }
                statusStrip.Items.Add(newSeparator)

                ' 添加新面板
                Dim newPanel As New ToolStripLabel() With {
                .Name = $"Panel{nPanel}",
                .Text = strText,
                .AutoSize = True
            }
                statusStrip.Items.Add(newPanel)
            Else
                ' 更新已有面板
                If targetIndex >= statusStrip.Items.Count Then
                    Throw New IndexOutOfRangeException($"Panel {nPanel} not found.")
                End If

                Dim targetPanel = TryCast(statusStrip.Items(targetIndex), ToolStripLabel)
                If targetPanel Is Nothing OrElse targetPanel.Name <> $"Panel{nPanel}" Then
                    Throw New InvalidOperationException($"Item at index {targetIndex} is not Panel{nPanel}.")
                End If

                targetPanel.Text = strText
            End If

        Catch ex As Exception
            Dim itemCount As Integer = If(mdlMain.frmNextCapInstance?.sbrNextCap?.Items IsNot Nothing,
                                     mdlMain.frmNextCapInstance.sbrNextCap.Items.Count, 0)
            LogEvent($"Error in UpdateStatusBar: {ex.Message} | Requested Panel: {nPanel} | Total Items: {itemCount}")
        End Try
    End Sub
End Module
