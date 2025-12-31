Imports System.DirectoryServices.ActiveDirectory

Public Class clsSystemSettings
    ' 本地变量用于保存属性值
    Private mvarstrBasicStyle As String ' 设置当前为 1XX 或 TICAP
    Private mvarbContinuousBad As Boolean
    Private mvarbGoodPenIdRequired As Boolean ' 本地副本
    Private mvarbBadPenIdRequired As Boolean ' 本地副本
    Private mvarbAskForNumberOfGoodPens As Boolean ' frmEnterGoodPen 上的 cmbNumber 被使用
    Private mvarbGoodPenEnterUsed As Boolean
    Private mvarnPenIdMemory As Integer ' 本地副本
    Private mvarlngPenIdTimeOut As Long ' 本地副本
    Private mvarstrPenIdCaptureProgram As String ' 本地副本
    Private mvarbPenIdUnique As Boolean ' 本地副本
    Private mvarstrPenIdSource As String ' 本地副本
    Private mvarbDebug As Boolean ' 本地副本
    Private mvarstrShiftFunction As String ' 本地副本
    Private mvarstrGoodPenVisualFeedBack As String ' 本地副本
    Private mvarstrGoodPenAudioFeedBack As String ' 本地副本
    Private mvarstrBadPenVisualFeedBack As String ' 本地副本
    Private mvarstrBadPenAudioFeedBack As String ' 本地副本
    Private mvarbGoodPenFeedBack As Boolean ' 本地副本
    Private mvarbBadPenFeedBack As Boolean ' 本地副本
    Private mvarbBarcodeScannerMode As Boolean ' 本地副本
    Private mvarbProductionDateAdjust As Boolean ' 本地副本
    Private mvarstrUNC_Name As String ' 本地副本
    Private mvarstrBusinessServerUNC As String ' 本地副本
    Private mvarnColorOfGood As Integer ' 图表上 Good 的默认颜色
    Private mvarnColorOfBad As Integer ' 图表上 Bad 的默认颜色
    Private mvarstrStatusIconLocation As String ' 本地副本
    Private mvarstrActiveXserver_ID As String ' 本地副本
    Private mvarstrLegendGood As String ' 本地副本
    Private mvarstrLegendBad As String ' 本地副本
    Private mvarlngMaxGroups As Long ' 本地副本
    Private mvarlngMaxVisibleGroups As Long ' 本地副本
    Private mvarbResetCountOnProductChange As Boolean ' 本地副本
    Private mvarbLogOutAtShift As Boolean ' 本地副本
    Private mvarstrMaterialMode As String ' 本地副本
    Private mvarbCSBDebug As Boolean ' 本地副本
    Private mvarbImageDebug As Boolean ' 本地副本
    Private mvarbPrimaryCreateLotOnly As Boolean ' 本地副本
    Private mvarbEnableRecoveryStep As Boolean ' 本地副本
    Private mvarstrDispositionDefault As String ' 本地副本
    Private mvarbUseDefaultRunType As Boolean ' 本地副本
    Private mvarbUseLotPointer As Boolean ' 本地副本
    Private mvarbEnableCSBIDE As Boolean ' 本地副本
    Private mvarstrHelpURL As String ' 本地副本
    Private mvarbBadPenCountEnabled As Boolean ' 在缺陷编辑器中显示坏笔计数窗口

    Public Property bBadPenCountEnabled() As Boolean
        Get
            Return mvarbBadPenCountEnabled
        End Get
        Set(ByVal value As Boolean)
            mvarbBadPenCountEnabled = value
        End Set
    End Property

    Public Property strHelpURL() As String
        Get
            Return mvarstrHelpURL
        End Get
        Set(ByVal value As String)
            mvarstrHelpURL = value
        End Set
    End Property

    Public Property bEnableCSBIDE() As Boolean
        Get
            Return mvarbEnableCSBIDE
        End Get
        Set(ByVal value As Boolean)
            mvarbEnableCSBIDE = value
        End Set
    End Property

    Public Property bUseLotPointer() As Boolean
        Get
            Return mvarbUseLotPointer
        End Get
        Set(ByVal value As Boolean)
            mvarbUseLotPointer = value
        End Set
    End Property

    Public Property bUseDefaultRunType() As Boolean
        Get
            Return mvarbUseDefaultRunType
        End Get
        Set(ByVal value As Boolean)
            mvarbUseDefaultRunType = value
        End Set
    End Property

    Public Property strDispositionDefault() As String
        Get
            Return mvarstrDispositionDefault
        End Get
        Set(ByVal value As String)
            mvarstrDispositionDefault = value

            If Not (frmDefectEditor Is Nothing) Then
                frmDefectEditor.UpdateDisposition()
            End If
        End Set
    End Property

    Public Property bEnableRecoveryStep() As Boolean
        Get
            Return mvarbEnableRecoveryStep
        End Get
        Set(ByVal value As Boolean)
            mvarbEnableRecoveryStep = value

            ' 设置输入屏幕使用标志
            If value Then
                mvarbGoodPenEnterUsed = True
                ' 确保屏幕未被恢复步骤屏幕使用
            ElseIf Not mvarbAskForNumberOfGoodPens Then
                mvarbGoodPenEnterUsed = False
            End If
            ' 调用屏幕重新计算例程
            'mdlMain.ConfigurePenEntry()
        End Set
    End Property

    Public Property bPrimaryCreateLotOnly() As Boolean
        Get
            If mvarbPrimaryCreateLotOnly Then
                ' 匹配我们的 UNC 名称与业务服务器的 UNC 名称
                If mvarstrUNC_Name = mvarstrBusinessServerUNC Then
                    Return mvarbPrimaryCreateLotOnly
                End If
            End If
            Return mvarbPrimaryCreateLotOnly
        End Get
        Set(ByVal value As Boolean)
            mvarbPrimaryCreateLotOnly = value
        End Set
    End Property

    Public Property bImageDebug() As Boolean
        Get
            Return mvarbImageDebug
        End Get
        Set(ByVal value As Boolean)
            mvarbImageDebug = value
        End Set
    End Property

    Public Property bCSBDebug As Boolean
        Get
            Return mvarbCSBDebug
        End Get
        Set(value As Boolean)
            mvarbCSBDebug = value
            ' 检查窗体是否存在并调用方法
            'If Not frmSAX Is Nothing Then
            '    frmSAX.UpdateCSBDebug()
            'End If
        End Set
    End Property

    Public Property strMaterialMode() As String
        Get
            Return mvarstrMaterialMode
        End Get
        Set(ByVal value As String)
            mvarstrMaterialMode = value
        End Set
    End Property

    Public Property bLogOutAtShift() As Boolean
        Get
            Return mvarbLogOutAtShift
        End Get
        Set(ByVal value As Boolean)
            mvarbLogOutAtShift = value
        End Set
    End Property

    Public Property bResetCountOnProductChange() As Boolean
        Get
            Return mvarbResetCountOnProductChange
        End Get
        Set(ByVal value As Boolean)
            mvarbResetCountOnProductChange = value
        End Set
    End Property

    Public Property lngMaxVisibleGroups() As Long
        Get
            Return mvarlngMaxVisibleGroups
        End Get
        Set(ByVal value As Long)
            mvarlngMaxVisibleGroups = value
        End Set
    End Property

    Public Property lngMaxGroups() As Long
        Get
            Return mvarlngMaxGroups
        End Get
        Set(ByVal value As Long)
            mvarlngMaxGroups = value
        End Set
    End Property

    Public Property strLegendBad() As String
        Get
            Return mvarstrLegendBad
        End Get
        Set(ByVal value As String)
            mvarstrLegendBad = value
        End Set
    End Property

    Public Property strLegendGood() As String
        Get
            Return mvarstrLegendGood
        End Get
        Set(ByVal value As String)
            mvarstrLegendGood = value
        End Set
    End Property

    Public Property strActiveXserver_ID() As String
        Get
            Return mvarstrActiveXserver_ID
        End Get
        Set(ByVal value As String)
            mvarstrActiveXserver_ID = value
        End Set
    End Property

    Public Property strStatusIconLocation() As String
        Get
            Return mvarstrStatusIconLocation
        End Get
        Set(ByVal value As String)
            mvarstrStatusIconLocation = value
        End Set
    End Property

    Public Property nColorOfBad() As Integer
        Get
            Return mvarnColorOfBad
        End Get
        Set(ByVal value As Integer)
            mvarnColorOfBad = value
        End Set
    End Property

    Public Property nColorOfGood() As Integer
        Get
            Return mvarnColorOfGood
        End Get
        Set(ByVal value As Integer)
            mvarnColorOfGood = value
        End Set
    End Property

    Public Property strBusinessServerUNC() As String
        Get
            Return mvarstrBusinessServerUNC
        End Get
        Set(ByVal value As String)
            mvarstrBusinessServerUNC = value
        End Set
    End Property

    Public Property strUNC_Name() As String
        Get
            Return mvarstrUNC_Name
        End Get
        Set(ByVal value As String)
            mvarstrUNC_Name = value
        End Set
    End Property

    Public Property bProductionDateAdjust() As Boolean
        Get
            Return mvarbProductionDateAdjust
        End Get
        Set(ByVal value As Boolean)
            Dim bTemp As Boolean
            ' 必须保存此值，以便我们知道是否触发表单上下文的更改，但我们需要 frmContext 抑制其写回此过程
            bTemp = mvarbProductionDateAdjust
            mvarbProductionDateAdjust = value
            If value <> bTemp Then
                frmAllContext.GetInstance().EnableProductionDate = value
            End If
        End Set
    End Property

    Public Property bBarcodeScannerMode() As Boolean
        Get
            Return mvarbBarcodeScannerMode
        End Get
        Set(ByVal value As Boolean)
            mvarbBarcodeScannerMode = value
        End Set
    End Property

    Public Property bBadPenFeedBack() As Boolean
        Get
            Return mvarbBadPenFeedBack
        End Get
        Set(ByVal value As Boolean)
            mvarbBadPenFeedBack = value
        End Set
    End Property

    Public Property bGoodPenFeedBack() As Boolean
        Get
            Return mvarbGoodPenFeedBack
        End Get
        Set(ByVal value As Boolean)
            mvarbGoodPenFeedBack = value
        End Set
    End Property

    Public Property strBadPenAudioFeedBack() As String
        Get
            Return mvarstrBadPenAudioFeedBack
        End Get
        Set(ByVal value As String)
            mvarstrBadPenAudioFeedBack = value
        End Set
    End Property

    Public Property strBadPenVisualFeedBack() As String
        Get
            Return mvarstrBadPenVisualFeedBack
        End Get
        Set(ByVal value As String)
            mvarstrBadPenVisualFeedBack = value
        End Set
    End Property

    Public Property strGoodPenAudioFeedBack() As String
        Get
            Return mvarstrGoodPenAudioFeedBack
        End Get
        Set(ByVal value As String)
            mvarstrGoodPenAudioFeedBack = value
        End Set
    End Property

    Public Property strGoodPenVisualFeedBack() As String
        Get
            Return mvarstrGoodPenVisualFeedBack
        End Get
        Set(ByVal value As String)
            mvarstrGoodPenVisualFeedBack = value
        End Set
    End Property

    Public Property strShiftFunction() As String
        Get
            Return mvarstrShiftFunction
        End Get
        Set(ByVal value As String)
            mvarstrShiftFunction = value
            ' 如果是自动更改，则禁用班次更改组合框
            If value = "AutoChange" Then
                frmAllContext.GetInstance().EnableShiftChange(False)
            Else
                frmAllContext.GetInstance().EnableShiftChange(True)
            End If
        End Set
    End Property

    Public Property bDebug() As Boolean
        Get
            Return mvarbDebug
        End Get
        Set(ByVal value As Boolean)
            mvarbDebug = value
        End Set
    End Property

    Public Property strPenIdSource() As String
        Get
            Return mvarstrPenIdSource
        End Get
        Set(ByVal value As String)
            mvarstrPenIdSource = value
        End Set
    End Property

    Public Property bPenIdUnique() As Boolean
        Get
            Return mvarbPenIdUnique
        End Get
        Set(ByVal value As Boolean)
            mvarbPenIdUnique = value
        End Set
    End Property

    Public Property strPenIdCaptureProgram() As String
        Get
            Return mvarstrPenIdCaptureProgram
        End Get
        Set(ByVal value As String)
            mvarstrPenIdCaptureProgram = value
        End Set
    End Property

    Public Property lngPenIdTimeOut() As Long
        Get
            Return mvarlngPenIdTimeOut
        End Get
        Set(ByVal value As Long)
            mvarlngPenIdTimeOut = value
        End Set
    End Property

    Public Property nPenIdMemory() As Integer
        Get
            Return mvarnPenIdMemory
        End Get
        Set(ByVal value As Integer)
            mvarnPenIdMemory = value
            If value > 0 Then
                mvarbPenIdUnique = True
            End If
        End Set
    End Property

    Public Property bGoodPenEnterUsed() As Boolean
        Get
            Return mvarbGoodPenEnterUsed
        End Get
        Set(ByVal value As Boolean)
            ' 仅当设置为使用好笔计数或恢复步骤时才允许
            If value AndAlso (mvarbEnableRecoveryStep OrElse mvarbAskForNumberOfGoodPens) Then
                mvarbGoodPenEnterUsed = value
            Else
                mvarbGoodPenEnterUsed = False
            End If
        End Set
    End Property

    Public Property bAskForNumberOfGoodPens() As Boolean
        Get
            Return mvarbAskForNumberOfGoodPens
        End Get
        Set(ByVal value As Boolean)
            mvarbAskForNumberOfGoodPens = value

            ' 设置输入屏幕使用标志
            If value Then
                mvarbGoodPenEnterUsed = True
                ' 确保屏幕未被恢复步骤屏幕使用
            ElseIf Not mvarbEnableRecoveryStep Then
                mvarbGoodPenEnterUsed = False
            End If
            ' 调用屏幕重新计算例程
            'mdlMain.ConfigurePenEntry()
        End Set
    End Property

    Public Property bBadPenIdRequired() As Boolean
        Get
            Return mvarbBadPenIdRequired
        End Get
        Set(ByVal value As Boolean)
            mvarbBadPenIdRequired = value
        End Set
    End Property

    Public Property bGoodPenIdRequired() As Boolean
        Get
            Return mvarbGoodPenIdRequired
        End Get
        Set(ByVal value As Boolean)
            mvarbGoodPenIdRequired = value
        End Set
    End Property

    Public Property strBasicStyle() As String
        Get
            Return mvarstrBasicStyle
        End Get
        Set(ByVal value As String)
            mvarstrBasicStyle = value
        End Set
    End Property

    Public Property bContinuousBad() As Boolean
        Get
            Return mvarbContinuousBad
        End Get
        Set(ByVal value As Boolean)
            mvarbContinuousBad = value
        End Set
    End Property

    ' 类初始化
    Public Sub New()
        Try
            ' 设置应用程序的默认操作模式
            mvarstrBasicStyle = "NextCap"
            ' 图表上 Good 和 Bad 的默认颜色
            mvarnColorOfGood = 10 ' 浅绿色
            mvarnColorOfBad = 12 ' 红色

            ' 其他默认值
            mvarstrLegendGood = "Good"
            mvarstrLegendBad = "Bad"
        Catch ex As Exception
            ' 处理异常
            Console.WriteLine($"初始化出错: {ex.Message}")
        End Try
    End Sub
End Class

