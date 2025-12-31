Imports System.Reflection.Metadata
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.ProgressBar
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.ToolTip
Imports MSScriptControl
Module mdlSAX
    '/*public constants for VetoLotStatusChange
    Public Const m_csClose As String = "CLOSE"
    Public Const m_csSuspend As String = "SUSPEND"
    Public Const m_csOpen As String = "OPEN"

    '/*Member pointers to SAX handler objects and their proto types
    Private mcol_SaxTranslations As colSAXTranslations

    '[v2.0.1] Change to name equivalence for matching up correct code block in the
    '         incoming module(s)
    Private Const cstrAugmentPenId As String = "AUGMENTPENID"
    '--/*Verify the unit id against a set of rules
    Private Const cstrVerifyId As String = "VERIFYID"
    '--/*Verify the Penid  against a set of rules
    Private Const cstrVerifyPenId As String = "VERIFYPENID"
    '--/*Qualify Lot  against a set of rules
    Private Const cstrProductMonitor As String = "PRODUCTMONITOR"
    '--/*Generate a new Lot Id using the system Context
    Private Const cstrCreateLotId As String = "CREATELOTID"
    '--/*Execute script when End of Lot triggered
    Private Const cstrEndOfLot As String = "ENDOFLOT"
    '--/*Script executed prior to new lot creation; triggered prior to the CreateLot call
    '--/*passed to the server.
    Private Const cstrPreCreateLot As String = "PRECREATELOT"
    '--/*Script executed after the entry of a good unit
    Private Const cstrPostGoodPenEntry As String = "POSTGOODPENENTRY"
    '--/*Script executed after the entry of a good unit
    Private Const cstrPostBadPenEntry As String = "POSTBADPENENTRY"
    '--/*Script to allow for user selection of data on the frmNextCap status bar
    Private Const cstrUpdateStatusBar As String = "UPDATESTATUSBAR"
    '--/*Script executed prior to automatic shift change
    Private Const cstrPreShiftChange As String = "PRESHIFTCHANGE"
    '--/*Script executed after shift change
    Private Const cstrPostShiftChange As String = "POSTSHIFTCHANGE"
    '--/*Script executed before Lot Id can be created to verify it against a rule
    Private Const cstrVerifyLotId As String = "VERIFYLOTID"
    '--/*Executed from frmAllContext.WindowInit() to allow custom field disable.
    '--/*Return of True --> field.Enabled = False
    Private Const cstrContextLock As String = "CONTEXTLOCK"
    '--/*Executed after CreateLotId() is started, but before PreCreateLot() CSB
    Private Const cstrLotWizard As String = "LOTWIZARD"
    '--/*Executed after CreateLotId() is started, but before PreCreateLot() CSB
    Private Const cstrClientSystemStartUp As String = "CLIENTSYSTEMSTARTUP"
    '--/*Executed after a lot is opened in the active LotManager
    Private Const cstrLotOpened As String = "LOTOPENED"
    '--/*Executed after a closed lot is opened in the active LotManager
    Private Const cstrClosedLotOpened As String = "CLOSEDLOTOPENED"
    '--/*Executed after a suspended lot is opened in the active LotManager
    Private Const cstrSuspendedLotOpened As String = "SUSPENDEDLOTOPENED"
    '--/*Executed after a lot is suspended in the active LotManager
    Private Const cstrLotSuspended As String = "LOTSUSPENDED"
    '--/*Executed after a lot is suspended in the active LotManager
    Private Const cstrVetoLotStatusChange As String = "VETOLOTSTATUSCHANGE"
    '--/*Executed after a LotManager Change and after the global Context has been modified
    Private Const cstrLotManagerChange As String = "LOTMANAGERCHANGE"

    '=======================================================
    'Routine: mdlSAX.ExecuteCSB_VetoLotStatusChange(str,dt,str,str)
    'Purpose: This allows a CSB to veto the change of a lot
    'to another state from where ever it is currently. This
    'is hooked into the frmLotManager, frmEndLot, frmNewLot
    '
    'Globals: None
    '
    'Input: strLotId - The lot being changed
    '
    '       dtBirth - The birthday of the lot
    '
    '       sOld - The current state
    '
    '       sNew - The requested state
    '
    '       vrtReturnCode - Any error that was generated
    '       by the attempt to or the operation of the SAX handler.
    '
    'Return: Boolean - False is allow, e.g. no Veto.
    '
    'Modifications:
    '   08-16-2001 v:1.1.3 As written Chris Barker
    '
    '
    '=======================================================
    Public Function ExecuteCSB_VetoLotStatusChange(
    ByRef strLotId As String,
    ByRef dtBirth As Date,
    ByRef sOld As String,
    ByRef sNew As String,
    Optional ByRef vrtReturnCode As Object = Nothing
) As Boolean
        Dim strCSBIndex As String
        Dim bSelfRegister As Boolean 'pppp
        Console.WriteLine(dtBirth.ToString("yyyy-MM-dd HH:mm:ss"))
        Try
            ' 从翻译集合获取CSB索引
            'strCSBIndex = mcol_SaxTranslations(CStr(cstrVetoLotStatusChange)).strCSBIndex
            'bSelfRegister = mcol_SaxTranslations(CStr(cstrVetoLotStatusChange)).bSelfRegister 'pppp

            '' 根据索引决定执行逻辑
            'If String.IsNullOrEmpty(strCSBIndex) OrElse Not bSelfRegister Then
            '    ' 无模块时返回False（保持原逻辑）
            '    Return False
            'Else
            '    ' 执行SAX处理程序
            '    Dim handler As clsSAXhandler = gcol_SAXhandlers(strCSBIndex)
            '    If handler IsNot Nothing AndAlso handler.bLoaded Then
            '        ' 调用Evaluate方法并传递多个参数
            '        Return CBool(handler.oHandler.Evaluate(strLotId, dtBirth, sOld, sNew))
            '    End If
            'End If

            ' 若处理程序未执行，返回默认值False
            Return False
        Catch ex As Exception
            ' 错误处理：记录错误并返回False
            vrtReturnCode = ex.HResult
            Err.Clear()
            Return False
        End Try
    End Function
    Private Function ParseReturnTypeFromPrototype(protoLine As String) As String
        If String.IsNullOrWhiteSpace(protoLine) Then Return String.Empty

        ' Strip trailing line comment (')
        Dim tick = protoLine.IndexOf("'"c)
        If tick >= 0 Then protoLine = protoLine.Substring(0, tick)

        ' Must contain Function
        Dim funcIdx = protoLine.IndexOf("Function", 0, StringComparison.OrdinalIgnoreCase)
        If funcIdx < 0 Then Return String.Empty

        ' Find balanced (...) after Function
        Dim openIdx = protoLine.IndexOf("("c, funcIdx)
        If openIdx < 0 Then Return String.Empty

        Dim depth = 0, closeIdx = -1
        For i = openIdx To protoLine.Length - 1
            Dim ch = protoLine(i)
            If ch = "("c Then
                depth += 1
            ElseIf ch = ")"c Then
                depth -= 1
                If depth = 0 Then
                    closeIdx = i
                    Exit For
                End If
            End If
        Next
        If closeIdx = -1 Then Return String.Empty

        ' Everything after params: look for "As <type>"
        Dim tail = protoLine.Substring(closeIdx + 1)

        Dim m = System.Text.RegularExpressions.Regex.Match(
        tail,
        "\bAs\b\s+(?<type>.+?)(?=\s+(Implements|Handles)\b|$)",
        System.Text.RegularExpressions.RegexOptions.IgnoreCase
    )
        If Not m.Success Then Return String.Empty
        Return m.Groups("type").Value.Trim()
    End Function
    Private ReadOnly RxSpaces As New Regex("\s+", RegexOptions.Compiled)
    Private ReadOnly RxEq As New Regex("\s*=\s*", RegexOptions.Compiled)
    Private ReadOnly RxComma As New Regex("\s*,\s*", RegexOptions.Compiled)
    Private ReadOnly RxOpen As New Regex("\(\s+", RegexOptions.Compiled)
    Private ReadOnly RxClose As New Regex("\s+\)", RegexOptions.Compiled)
    Private Function NormalizeHeaderLine(s As String) As String
        If String.IsNullOrWhiteSpace(s) Then Return String.Empty
        s = s.Trim()
        s = RxSpaces.Replace(s, " ")       ' collapse to single spaces
        s = RxEq.Replace(s, " = ")         ' make " = " (or use "=" if you prefer tight)
        s = RxComma.Replace(s, ", ")       ' tidy commas
        s = RxOpen.Replace(s, "(")         ' no space after "("
        s = RxClose.Replace(s, ")")        ' no space before ")"
        Return s
    End Function
    Function NormalizeNewlines(t As String) As String
        Return t.Replace(vbCrLf, vbLf).Replace(vbCr, vbLf)
    End Function
    Private Function NormalizeQuotes(t As String) As String
        ' Use Unicode code points to avoid char-literal issues in some editors/encodings
        Return t _
        .Replace(ChrW(&H201C), """"c).Replace(ChrW(&H201D), """"c).  ' “ ”
        Replace(ChrW(&H2018), "'"c).Replace(ChrW(&H2019), "'"c)    ' ‘ ’
    End Function
    Function CleanHeadTailNonAscii(t As String) As String
        ' keep your own RemoveNonAscii* if you have special logic;
        ' here we just trim BOM & odd whitespace safely.
        Dim u = t
        If u.Length > 0 AndAlso AscW(u(0)) = &HFEFF Then
            u = u.Substring(1)
        End If
        Return u.Trim()
    End Function
    ' Helper to set a parameter name if clsProtoTypeParam exposes a name property.
    ' Returns True if set; False if not supported (no such property).
    Private Function TrySetParamName(p As clsProtoTypeParam, paramName As String) As Boolean
        Try
            Dim prop = p.GetType().GetProperty("sName")
            If prop IsNot Nothing AndAlso prop.CanWrite Then
                prop.SetValue(p, paramName, Nothing)
                Return True
            End If
        Catch
            ' ignore
        End Try
        Return False
    End Function
    Public Function ParseScript(ByRef sScript As String) As colProtoTypes
        Dim colProtos As New colProtoTypes()

        ' --- Clean & normalize ---
        sScript = CleanHeadTailNonAscii(sScript)
        sScript = NormalizeQuotes(sScript)
        sScript = NormalizeNewlines(sScript)

        ' Split by LF safely
        Dim lines As String() = sScript.Split(New String() {vbLf}, StringSplitOptions.None)
        If lines.Length = 0 Then Return colProtos

        ' Prototype is guaranteed on the first line
        Dim firstLine As String = lines(0).Trim()
        Dim sFullProtoType As String = NormalizeHeaderLine(firstLine)

        Dim protoRegex As New Regex(
        "^\s*(?://\s*)?(?<kind>Function|Sub)\s+(?<name>[A-Za-z_]\w*)\s*(?<paren>\((?<params>[^\)]*)\))?",
        RegexOptions.IgnoreCase Or RegexOptions.Compiled
    )
        Dim protoMatch As Match = protoRegex.Match(sFullProtoType)
        If Not protoMatch.Success Then
            Debug.WriteLine("ParseScript: No valid prototype found in first line. Contact Support and halt app.")
            MessageBox.Show("No valid prototype found in first line. Contact Support and halt app.",
                        "Fatal Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Environment.Exit(1)
            Return colProtos
        End If

        Dim kind As String = protoMatch.Groups("kind").Value
        Dim fname As String = protoMatch.Groups("name").Value
        Dim paramsTextRaw As String = protoMatch.Groups("params").Value
        Dim normalizedParams As String = Regex.Replace(If(paramsTextRaw, ""), "\s+", " ").Trim()

        Dim sClassName As String = If(kind.Equals("Sub", StringComparison.OrdinalIgnoreCase), "Sub", "Function")
        Dim sFunctionName As String = fname

        ' Capture return type (if Function)
        Dim sProtoReturnType As String = ""
        If sClassName.Equals("Function", StringComparison.OrdinalIgnoreCase) Then
            sProtoReturnType = ParseReturnTypeFromPrototype(sFullProtoType) ' e.g. "Boolean"
        End If

        ' --- Parse parameters (supports both patterns) ---
        Dim oParams As New colProtoTypeParam()
        If normalizedParams.Length > 0 Then
            ' Try grouped form: "x, y As Integer"
            Dim groupMatch As Match = Regex.Match(
            normalizedParams,
            "^\s*(?<names>[A-Za-z_]\w*(?:\s*,\s*[A-Za-z_]\w*)*)\s+As\s+(?<type>[A-Za-z_]\w*)\s*$",
            RegexOptions.IgnoreCase
        )

            If groupMatch.Success Then
                Dim typeName As String = groupMatch.Groups("type").Value
                For Each nm In groupMatch.Groups("names").Value.Split(","c)
                    Dim name As String = nm.Trim()
                    oParams.Add(name, typeName)
                Next
            Else
                ' Per-item form: "x As String, y As Integer, z As String"
                For Each chunk As String In normalizedParams.Split(","c)
                    Dim itemMatch As Match = Regex.Match(
                    chunk,
                    "^\s*(?<name>[A-Za-z_]\w*)\s+As\s+(?<type>[A-Za-z_]\w*)\s*$",
                    RegexOptions.IgnoreCase
                )
                    If Not itemMatch.Success Then Continue For

                    Dim name As String = itemMatch.Groups("name").Value
                    Dim typeName As String = itemMatch.Groups("type").Value
                    oParams.Add(name, typeName)
                Next
            End If
        End If

        ' Body = everything after the first line
        Dim sb As New StringBuilder()
        For i As Integer = 1 To lines.Length - 1
            sb.AppendLine(lines(i))
        Next
        Dim sCodeBody As String = sb.ToString().TrimStart()

        ' Build result
        Dim oProto As New clsProtoType() With {
        .sClassName = sClassName,
        .sFunctionName = sFunctionName,
        .oParams = oParams,
        .sProtoReturnType = sProtoReturnType,
        .sFullProtoType = sFullProtoType, ' normalized header line
        .sCodeBody = sCodeBody
    }

        colProtos.Add(oProto, sFunctionName)
        Return colProtos
    End Function
    '=======================================================
    'Routine: mdlSAX.AddSAXmodule(str)
    'Purpose: Call the code analysis procedures and
    '         add the code to the global collection.
    '
    'Globals:None
    '
    'Input: strCode = Code for the SAX module.
    '
    '       strCSBIndex - Return byref the value for the
    '       CSB index if it is unknown.
    '
    'Return: Boolean - The module was sucessfully loaded
    '        into the SAX engine.
    '
    'Modifications:
    '   10-22-1998 As written for Pass1.2
    '
    '   08-31-2001 v2.0.1 Modified to use mdlCSBParser
    '   to find the rountines in a code block. This
    '   uses the new SAX translations to match the prototype
    '   based on name equivalence.
    '
    '   09-24-2001 v2.0.1 Chris Barker: Change GetSAXhandler
    '   to use sFullProtoType instead of the matching alias
    '   so that sax will recognize it.
    '
    '=======================================================
    Public Function AddSAXmodule(ByRef strScript As String, Optional ByRef strCSBIndex As String = Nothing) As Boolean

        Dim strProtoReturnType As String = ""
        Dim strProtoType As String = ""
        Dim strCode As String = ""
        Dim strFileName As String = ""
        Dim sClassName As String = ""
        Dim exists As Boolean = False
        Dim oparam As colProtoTypeParam = Nothing
        Dim colProtos As colProtoTypes = Nothing
        Dim lngTranslateIndex As Long = 0

        Try
            'colProtos = mdlCSBParser.ParseSourceCode(strScript)
            colProtos = ParseScript(strScript)
            If colProtos IsNot Nothing AndAlso colProtos.Count > 0 Then
                ' Take the first proto (parser yields one per module)
                Dim nIdx As Integer = 0

                strProtoType = mdlSAX.MatchProtoType(colProtos(nIdx), lngTranslateIndex) ' set mcol_SaxTranslations(nPtr).bSelfRegister = True
                If Not String.IsNullOrEmpty(strProtoType) Then
                    strCode = colProtos(nIdx).sCodeBody
                    sClassName = colProtos(nIdx).sClassName
                    strFileName = colProtos(nIdx).sFunctionName
                    oparam = colProtos(nIdx).oParams
                    strProtoReturnType = colProtos(nIdx).sFullProtoType
                Else
                    Debug.WriteLine($"mdlSAX.AddSAXmodule could not recognise the Script name {strProtoType}")
                    If go_clsSystemSettings.bCSBDebug Then
                        Dim filePath As String = IO.Path.Combine(Application.StartupPath, "SAXnoui.txt")
                        FlushToFile(filePath, $"ProtoType could not be identified:{Environment.NewLine}{strProtoType}")
                    End If
                End If

                If Not String.IsNullOrEmpty(strProtoType) Then
                    strCSBIndex = strFileName

                    If gcol_SAXhandlers IsNot Nothing Then
                        Dim existingHandler As clsSAXhandler = Nothing
                        exists = gcol_SAXhandlers.TryGetByKey(strCSBIndex, existingHandler)

                        If Not exists Then

                            Dim handlerObj As New clsSAXEngineHandler(strCode, strFileName, isFunction:=sClassName.Equals("Function", StringComparison.OrdinalIgnoreCase))

                            gcol_SAXhandlers.Add(
                            bLoaded:=True,
                            strType:=sClassName,
                            strCode:=strCode,
                            strFileName:=strFileName,
                            oParams:=oparam,
                            oHandler:=handlerObj,
                            strProtoType:=strProtoReturnType,
                            strLocation:="",
                            sKey:=strCSBIndex
                        )

                            Debug.WriteLine("Handler added for key: " & strCSBIndex)
                            exists = True
                        Else
                            Debug.WriteLine("Handler already exists for key: " & strCSBIndex)
                        End If
                    End If

                    AddSAXmodule = exists
                End If
            End If
        Catch ex As Exception
            MainErrorHandler("AddSAXmodule", ex.Message, "Error")
            AddSAXmodule = False
        End Try
    End Function
    '=======================================================
    'Routine: MatchProtoType(o)
    'Purpose: Just scan the SAX Translations object
    ' for a mathcing function name.
    '
    'Globals:None
    '
    'Input: oProtoType - A prototype object
    '
    '       lngIdx - pointer to the collection index (return
    '       byref for this param).
    '
    'Return: String - the index string to the routine.
    '
    'Tested:
    '
    'Modifications:
    '   08-31-2001 v2.0.1 As written Chris Barker
    '
    '
    '=======================================================
    Public Function MatchProtoType(ByRef oProtoType As clsProtoType, ByRef lngIdx As Integer) As String
        For nPtr As Integer = 0 To mcol_SaxTranslations.Count - 1
            If mcol_SaxTranslations(nPtr).strProtoType.ToUpper() = oProtoType.sFunctionName.ToUpper() Then
                lngIdx = nPtr + 1  ' VB6索引从1开始，VB.NET从0开始需+1
                mcol_SaxTranslations(nPtr).bSelfRegister = True
                Return mcol_SaxTranslations(nPtr).strProtoType
            End If
        Next nPtr

        Return String.Empty
    End Function
    '===========================================================
    'Routine: mdlSAX.GetNewModuleName()
    'Purpose: This scans the global collection for the next
    '         available module name.
    '
    'Globals:None
    '
    'Input: None
    '
    'Return: String - This returns a file name in the form of
    '                 '*###' where this is non-repeating during
    '                 a session.
    '
    'Modifications:
    '   10-22-1998 As written for Pass1.2
    '
    '
    '============================================================
    Private Function GetNewModuleName() As String
        If gcol_SAXhandlers.Count = 0 Then
            Return "*1"
        Else
            Dim lastItem = gcol_SAXhandlers(gcol_SAXhandlers.Count - 1)  ' VB.NET 索引从0开始
            Dim strTemp As String = lastItem.strFileName

            ' 移除开头的 "*"
            If strTemp.StartsWith("*") Then
                strTemp = strTemp.Substring(1)  ' 从索引1开始截取
            End If

            ' 转换为整数并递增（处理可能的非数字情况）
            Dim numericValue As Integer
            If Integer.TryParse(strTemp, numericValue) Then
                Return "*" & (numericValue + 1).ToString()
            Else
                ' 处理解析失败的情况（例如返回默认值或抛出异常）
                Return "*1"
            End If
        End If
    End Function
    '=======================================================
    'Routine: mdlSAX.GetSAXhandler(str,str)
    'Purpose: This creates the SAX Handler to the
    '         delcared prototype of the module. The module
    '         is then loaded into the NoUI SAX engine.
    '
    'Globals:None
    '
    'Input: strModuleName - The alias used for the module
    '       of code.
    '
    '       strProtoType - The function declaration that
    '       represents the entry point of the module.
    '
    'Return: Object - A pointer to the SAX Handler object.
    '
    'Modifications:
    '   10-29-1998 As written for Pass1.2
    '
    '
    '=======================================================
    Private Function GetSAXhandler(ByRef strModuleName As String, ByRef strProtoType As String, ByRef strCode As String) As Object
        'Dim oHand As Object = Nothing

        '' 创建处理器实例
        'oHand = frmSAX.SaxBasicEntNoUI1.CreateHandler(strModuleName & "|" & strProtoType)
        'If oHand Is Nothing Then Return Nothing  ' 直接返回 Nothing

        '' 设置活动代码
        'frmSAX.ActiveCode = strCode

        '' 加载模块
        'If frmSAX.SaxBasicEntNoUI1.LoadModule(strModuleName) Then
        '    Return oHand  ' 返回处理器实例
        'Else
        '    ' 记录错误
        '    If go_clsSystemSettings.bCSBDebug Then
        '        If frmSAXIDE.Visible Then
        '            frmMessage.GenerateMessage(frmSAX.SaxBasicEntNoUI1.ErrorText)
        '        Else
        '            frmSAX.LoadErrorCode(strModuleName, strProtoType, strCode)
        '        End If
        '        Debug.WriteLine(frmSAX.SaxBasicEntNoUI1.ErrorText)
        '    End If
        '    Return Nothing  ' 加载失败返回 Nothing
        'End If

        Return Nothing
    End Function
    '=======================================================
    'Routine: mdlSAX.CreateProtoTypes()
    'Purpose: This is where the set of known SAX Module
    '         proto types is initialized.
    '
    'Globals:None
    '
    'Input:None
    '
    'Return:None
    '
    'Modifications:
    '   11-03-1998 As written for Pass1.2
    '
    '   07/05/2001 Chris Barker
    '   Added in ContextLock and LotWizard
    '
    '   8-16-2001 Chris Barker
    '   Added in VetoLotStatusChange v1.1.3
    '
    '   8-24-2001 Chris Barker
    '   Added in LotManagerChange v1.1.5
    '=======================================================
    Public Sub CreateProtoTypes()
        ' 初始化 SAX 翻译集合
        mcol_SaxTranslations = New colSAXTranslations()

        ' 添加已知的原型类型
        mcol_SaxTranslations.Add(cstrAugmentPenId, False, "AugmentPenId")
        mcol_SaxTranslations.Add(cstrVerifyPenId, False, "VerifyPenId")
        mcol_SaxTranslations.Add(cstrCreateLotId, False, "CreateLotId")
        mcol_SaxTranslations.Add(cstrEndOfLot, False, "EndOfLot")
        mcol_SaxTranslations.Add(cstrPreCreateLot, False, "PreCreateLot")
        mcol_SaxTranslations.Add(cstrPostGoodPenEntry, False, "PostGoodPenEntry")
        mcol_SaxTranslations.Add(cstrUpdateStatusBar, False, "UpdateStatusBar")
        mcol_SaxTranslations.Add(cstrPreShiftChange, False, "PreShiftChange")
        mcol_SaxTranslations.Add(cstrPostShiftChange, False, "PostShiftChange")
        mcol_SaxTranslations.Add(cstrPostBadPenEntry, False, "PostBadPenEntry")
        mcol_SaxTranslations.Add(cstrVerifyLotId, False, "VerifyLotId")
        mcol_SaxTranslations.Add(cstrContextLock, False, "ContextLock")
        mcol_SaxTranslations.Add(cstrLotWizard, False, "LotWizard")
        mcol_SaxTranslations.Add(cstrClientSystemStartUp, False, "ClientSystemStartUp")
        mcol_SaxTranslations.Add(cstrLotOpened, False, "LotOpened")
        mcol_SaxTranslations.Add(cstrClosedLotOpened, False, "ClosedLotOpened")
        mcol_SaxTranslations.Add(cstrSuspendedLotOpened, False, "SuspendedLotOpened")
        mcol_SaxTranslations.Add(cstrLotSuspended, False, "LotSuspended")
        mcol_SaxTranslations.Add(cstrVetoLotStatusChange, False, "VetoLotStatusChange")
        mcol_SaxTranslations.Add(cstrLotManagerChange, False, "LotManagerChange")
    End Sub
    '/*Property to access use of LotWizard CSB
    Public ReadOnly Property LotWizardUsed As Boolean
        Get
            'Return Not String.IsNullOrEmpty(mcol_SaxTranslations(cstrLotWizard).strCSBIndex)
            Return False
        End Get
    End Property
    '=======================================================
    'Routine: mdlSAX.ExecuteCSB_EndOfLot(vrt)
    'Purpose: This merely executes the users script when
    'the End-Of-Lot is triggered.
    '
    'Globals:None
    '
    'Input: strLotId - The requested Lot Id to be created.
    '
    '       strCSBIndex - The index of the particular
    '       handler in the collection of SAX objects.
    '
    '       vrtReturnCode - Any error that was generated
    '       by the attempt to or the operation of the SAX handler.
    '
    'Return: String - The result of the created Lot Id.
    '
    'Modifications:
    '   01-27-1999 As written for Pass1.6
    '
    '
    '=======================================================
    Public Function ExecuteCSB_PreCreateLot(ByVal strLotId As String, Optional ByRef vrtReturnCode As Object = Nothing) As String
        Dim strCSBIndex As String
        Dim bSelfRegister As Boolean

        Try
            ' 获取翻译集合中的 CSB 索引
            strCSBIndex = mcol_SaxTranslations(cstrPreCreateLot).strCSBIndex
            bSelfRegister = mcol_SaxTranslations(cstrPreCreateLot).bSelfRegister

            '' 根据索引值决定返回结果
            'If String.IsNullOrEmpty(strCSBIndex) OrElse Not bSelfRegister Then
            '    ' 若无模块，直接返回输入值
            '    Return strLotId
            'Else
            '    ' 执行 SAX 处理器
            '    Dim handler = gcol_SAXhandlers(strCSBIndex)
            '    If handler IsNot Nothing AndAlso handler.bLoaded Then
            '        Return handler.oHandler.Evaluate(strLotId)
            '    Else
            '        Return strLotId ' 或抛出异常，根据业务逻辑决定
            '    End If
            'End If
            ' 根据索引决定执行逻辑
            If String.IsNullOrEmpty(strCSBIndex) OrElse Not bSelfRegister Then
                ' 无模块时返回True（保持原逻辑）
                Return strLotId
            End If

            ' 执行SAX处理程序
            Dim handler As clsSAXhandler = gcol_SAXhandlers(strCSBIndex)
            Dim paramCol As New colProtoTypeParam()
            paramCol.Add("LotId", "String")
            paramCol.UpdateValue("LotId", strLotId)

            If handler IsNot Nothing AndAlso handler.bLoaded Then
                Console.WriteLine("ExecuteCSB_PreCreateLot executing.")
                Return handler.oHandler.Evaluate(paramCol, handler.oParams)
            End If

            Console.WriteLine("ExecuteCSB_PreCreateLot didn't execute.")
            ' 若处理程序未执行，返回默认值True（保持原逻辑）
            Return strLotId
        Catch ex As Exception
            ' 错误处理逻辑
            Debug.WriteLine($"执行 PreCreateLot 时出错: {ex.Message}")
            Return strLotId ' 默认返回输入值，可根据需求修改
        End Try
    End Function
    '=======================================================
    'Routine: mdlSAX.ExcuteCSB_VerifyLotId(str,vrt)
    'Purpose:
    '
    'Globals:None
    '
    'Input:
    '       strLotId - The Pen Id to validate
    '
    '       vrtReturnCode - Any error that was generated
    '       by the attempt to or the operation of the SAX handler.
    '
    'Return: Boolean - Whether or not the Lot Id met specification.
    '
    'Modifications:
    '   01-22-2001 As written. Chris Barker
    '
    '
    '=======================================================
    Public Function ExecuteCSB_VerifyLotId(ByVal strLotId As String, Optional ByRef vrtReturnCode As Object = Nothing) As Boolean
        Dim strCSBIndex As String
        Dim bSelfRegister As Boolean

        Try
            ' 获取翻译集合中的 CSB 索引
            strCSBIndex = mcol_SaxTranslations(cstrVerifyLotId).strCSBIndex
            bSelfRegister = mcol_SaxTranslations(cstrVerifyLotId).bSelfRegister

            ' 根据索引值决定返回结果
            If String.IsNullOrEmpty(strCSBIndex) OrElse Not bSelfRegister Then
                '    ' 若无模块，默认返回 True
                Return True
            Else
                ' 执行 SAX 处理器
                Dim handler As clsSAXhandler = gcol_SAXhandlers(strCSBIndex)
                Dim paramCol As New colProtoTypeParam()
                paramCol.Add("LotId", "String")
                paramCol.UpdateValue("LotId", strLotId)
                paramCol.Add("PartNumber", "String")
                paramCol.UpdateValue("PartNumber", go_Context.PartNumber)
                paramCol.Add("RunType", "String")
                paramCol.UpdateValue("RunType", go_Context.RunType)
                paramCol.Add("lineNumber", "Integer")
                paramCol.UpdateValue("lineNumber", CInt(go_Context.LineNumber))

                paramCol.Add("source", "String")
                paramCol.UpdateValue("source", go_Context.Source)
                paramCol.Add("SqlNextcap", "sqlconnection")
                paramCol.Add("accessNextcap", "oledbconnection")
                paramCol.Add("accessMfg", "oledbconnection")

                If handler IsNot Nothing AndAlso handler.bLoaded Then
                    Console.WriteLine("ExecuteCSB_VerifyLotId executing.")
                    Return handler.oHandler.Evaluate(paramCol, handler.oParams)
                End If
            End If

            Console.WriteLine("ExecuteCSB_VerifyLotId didn't execute.")
            Return True
        Catch ex As Exception
            ' 错误处理：设置返回码并返回默认值
            vrtReturnCode = ex.HResult ' 或 ex.ToString()，根据原代码意图
            Return False ' 发生错误时返回 False
        End Try
    End Function
    '=======================================================
    'Routine: mdlSAX.ExecuteCSB_LotSuspended(str,vrt)
    'Purpose: This merely executes the users script after
    'a lot is suspended in the active LotManager
    '
    'Globals: None
    '
    'Input: strLotId - The lot being changed
    '
    '       vrtReturnCode - Any error that was generated
    '       by the attempt to or the operation of the SAX handler.
    '
    'Return: None
    '
    'Modifications:
    '   08-09-2001 v:1.1.2 As written Chris Barker
    '
    '
    '=======================================================
    Public Sub ExecuteCSB_LotSuspended(ByRef strLotId As String, Optional ByRef vrtReturnCode As Object = Nothing)
        Dim strCSBIndex As String
        Dim bSelfRegister As Boolean

        Try
            ' 获取 CSB 索引
            'strCSBIndex = mcol_SaxTranslations(CStr(cstrLotSuspended)).strCSBIndex
            'bSelfRegister = mcol_SaxTranslations(CStr(cstrLotSuspended)).bSelfRegister

            '' 根据索引值决定是否执行处理程序
            'If String.IsNullOrEmpty(strCSBIndex) OrElse Not bSelfRegister Then
            '    ' 如果没有模块，直接返回输入值
            'Else
            '    ' 执行 SAX 处理程序
            '    Dim handler As clsSAXhandler = TryCast(gcol_SAXhandlers(strCSBIndex), clsSAXhandler)
            '    If handler IsNot Nothing AndAlso handler.bLoaded Then
            '        handler.oHandler.Evaluate(strLotId)
            '    End If
            'End If
        Catch ex As Exception
            ' 返回错误信息
            vrtReturnCode = ex.HResult
            Err.Clear()
        End Try
    End Sub
    '=======================================================
    'Routine: mdlSAX.ExecuteCSB_EndOfLot(vrt)
    'Purpose: This merely executes the users script when
    'the End-Of-Lot is triggered.
    '
    'Globals:None
    '
    'Input: strCSBIndex - The index of the particular
    '       handler in the collection of SAX objects.
    '
    '       vrtReturnCode - Any error that was generated
    '       by the attempt to or the operation of the SAX handler.
    '
    'Return: String - The result of the created Lot Id.
    '
    'Modifications:
    '   01-27-1999 As written for Pass1.6
    '
    '
    '=======================================================
    Public Sub ExecuteCSB_EndOfLot(Optional ByRef vrtReturnCode As Object = Nothing)
        Dim strCSBIndex As String
        Dim bSelfRegister As Boolean

        Try
            ' 获取 CSB 索引
            'strCSBIndex = mcol_SaxTranslations(CStr(cstrEndOfLot)).strCSBIndex
            'bSelfRegister = mcol_SaxTranslations(CStr(cstrEndOfLot)).bSelfRegister

            '' 根据索引值决定是否执行处理程序
            'If String.IsNullOrEmpty(strCSBIndex) OrElse Not bSelfRegister Then
            '    ' 如果没有模块，直接返回（无需操作）
            'Else
            '    ' 执行 SAX 处理程序
            '    Dim handler As clsSAXhandler = gcol_SAXhandlers(strCSBIndex)
            '    If handler IsNot Nothing AndAlso handler.bLoaded Then
            '        ' 调用无参数的 Call 方法（假设 oHandler 是可调用对象）
            '        handler.oHandler.Evaluate() ' 若为方法调用
            '    End If
            'End If
        Catch ex As Exception
            ' 返回错误信息（兼容 Variant）
            vrtReturnCode = ex.HResult
            Err.Clear()
        End Try
    End Sub
    '=======================================================
    'Routine: mdlSAX.ExecuteCSB_ClosedLotOpened(str,vrt)
    'Purpose: This merely executes the users script after
    'a closed lot is opened in the active LotManager
    '
    'Globals: None
    '
    'Input: strLotId - The lot being changed
    '
    '       vrtReturnCode - Any error that was generated
    '       by the attempt to or the operation of the SAX handler.
    '
    'Return: None
    '
    'Modifications:
    '   08-09-2001 v:1.1.2 As written Chris Barker
    '
    '
    '=======================================================
    Public Sub ExecuteCSB_ClosedLotOpened(ByRef strLotId As String, Optional ByRef vrtReturnCode As Object = Nothing)
        Dim strCSBIndex As String
        Dim bSelfRegister As Boolean

        Try
            '' 获取 CSB 索引（假设 cstrClosedLotOpened 为字符串常量）
            'strCSBIndex = mcol_SaxTranslations(CStr(cstrClosedLotOpened)).strCSBIndex
            'bSelfRegister = mcol_SaxTranslations(CStr(cstrClosedLotOpened)).bSelfRegister

            '' 根据索引决定是否执行处理程序
            'If String.IsNullOrEmpty(strCSBIndex) OrElse Not bSelfRegister Then
            '    ' 无模块时不执行操作
            'Else
            '    ' 执行 SAX 处理程序
            '    Dim handler As clsSAXhandler = gcol_SAXhandlers(strCSBIndex)
            '    If handler IsNot Nothing AndAlso handler.bLoaded Then
            '        ' 调用带参数的 Call 方法（假设 oHandler 是可调用对象）
            '        handler.oHandler.Evaluate(strLotId) ' 直接调用方法
            '    End If
            'End If
        Catch ex As Exception
            ' 返回错误信息（兼容 Variant）
            vrtReturnCode = ex.HResult
            Err.Clear()
        End Try
    End Sub
    '=======================================================
    'Routine: mdlSAX.ExecuteCSB_SuspendedLotOpened(str,vrt)
    'Purpose: This merely executes the users script after
    'a lot is opened in the active LotManager
    '
    'Globals: None
    '
    'Input: strLotId - The lot being changed
    '
    '       vrtReturnCode - Any error that was generated
    '       by the attempt to or the operation of the SAX handler.
    '
    'Return: None
    '
    'Modifications:
    '   08-09-2001 v:1.1.2 As written Chris Barker
    '
    '
    '=======================================================
    Public Sub ExecuteCSB_SuspendedLotOpened(ByRef strLotId As String, Optional ByRef vrtReturnCode As Object = Nothing)
        Dim strCSBIndex As String
        Dim bSelfRegister As Boolean

        Try
            ' 获取 CSB 索引（假设 cstrSuspendedLotOpened 为字符串常量）
            'strCSBIndex = mcol_SaxTranslations(CStr(cstrSuspendedLotOpened)).strCSBIndex
            'bSelfRegister = mcol_SaxTranslations(CStr(cstrSuspendedLotOpened)).bSelfRegister

            '' 根据索引决定是否执行处理程序
            'If String.IsNullOrEmpty(strCSBIndex) OrElse Not bSelfRegister Then
            '    ' 无模块时不执行操作
            'Else
            '    ' 执行 SAX 处理程序
            '    Dim handler As clsSAXhandler = gcol_SAXhandlers(strCSBIndex)
            '    If handler IsNot Nothing AndAlso handler.bLoaded Then
            '        ' 调用带参数的 Call 方法（假设 oHandler 是可调用对象）
            '        handler.oHandler.Evaluate(strLotId) ' 直接调用方法
            '    End If
            'End If
        Catch ex As Exception
            ' 返回错误信息（兼容 Variant）
            vrtReturnCode = ex.HResult
            Err.Clear()
        End Try
    End Sub
    '=======================================================
    'Routine: mdlSAX.ExecuteCSB_LotOpened(str,vrt)
    'Purpose: This merely executes the users script after
    'a lot is opened in the active LotManager
    '
    'Globals: None
    '
    'Input: strLotId - The lot being changed
    '
    '       vrtReturnCode - Any error that was generated
    '       by the attempt to or the operation of the SAX handler.
    '
    'Return: None
    '
    'Modifications:
    '   08-09-2001 v:1.1.2 As written Chris Barker
    '
    '
    '=======================================================
    Public Sub ExecuteCSB_LotOpened(ByRef strLotId As String, Optional ByRef vrtReturnCode As Object = Nothing)
        Dim strCSBIndex As String
        Dim bSelfRegister As Boolean

        Try
            ' 从翻译集合获取CSB索引（假设cstrLotOpened为字符串常量）
            'strCSBIndex = mcol_SaxTranslations(CStr(cstrLotOpened)).strCSBIndex
            'bSelfRegister = mcol_SaxTranslations(CStr(cstrLotOpened)).bSelfRegister

            '' 根据索引决定执行逻辑
            'If String.IsNullOrEmpty(strCSBIndex) OrElse Not bSelfRegister Then
            '    ' 无模块时不执行任何操作（保持原逻辑）
            '    Exit Sub
            'Else
            '    ' 执行SAX处理程序
            '    Dim handler As clsSAXhandler = gcol_SAXhandlers(strCSBIndex)
            '    If handler IsNot Nothing AndAlso handler.bLoaded Then
            '        ' 调用Call方法（原代码中明确使用Call，此处保持一致）
            '        handler.oHandler.Evaluate(strLotId)
            '    End If
            'End If

        Catch ex As Exception
            ' 错误处理：记录错误并返回错误信息
            vrtReturnCode = ex.HResult
            Err.Clear()
        End Try
    End Sub
    '=======================================================
    'Routine: mdlSAX.ExecuteCSB_CreateLotId(vrt) str
    'Purpose: This allows the user to create a Lot Id based
    'on the system global Context.
    '
    'Globals:None
    '
    'Input: strCSBIndex - The index of the particular
    '       handler in the collection of SAX objects.
    '
    '       vrtReturnCode - Any error that was generated
    '       by the attempt to or the operation of the SAX handler.
    '
    'Return: String - The result of the created Lot Id.
    '
    'Modifications:
    '   01-27-1999 As written for Pass1.6
    '
    '
    '=======================================================
    Public Function ExecuteCSB_CreateLotId(Optional ByRef vrtReturnCode As Object = Nothing) As String
        Dim strCSBIndex As String
        Dim bSelfRegister As Boolean 'pppp

        Try
            ' 从翻译集合获取CSB索引（假设cstrCreateLotId为字符串常量）
            strCSBIndex = mcol_SaxTranslations(CStr(cstrCreateLotId)).strCSBIndex
            bSelfRegister = mcol_SaxTranslations(CStr(cstrCreateLotId)).bSelfRegister 'pppp

            ' 根据索引决定执行逻辑
            If String.IsNullOrEmpty(strCSBIndex) OrElse Not bSelfRegister Then 'pppp
                ' 无模块时返回空字符串（或根据需求调整默认值）
                Return String.Empty
            Else
                ' 执行 SAX 处理器
                Dim handler As clsSAXhandler = gcol_SAXhandlers(strCSBIndex)
                Dim paramCol As New colProtoTypeParam()
                paramCol.Add("partNumber", "String")
                paramCol.UpdateValue("partNumber", go_Context.PartNumber)
                paramCol.Add("lineNumber", "Integer")
                paramCol.UpdateValue("lineNumber", CInt(go_Context.LineNumber))
                paramCol.Add("accessMfg", "oledbconnection")

                If handler IsNot Nothing AndAlso handler.bLoaded Then
                    Console.WriteLine("ExecuteCSB_CreateLotId executing.")
                    Return handler.oHandler.Evaluate(paramCol, handler.oParams)
                End If
            End If

            Console.WriteLine("ExecuteCSB_CreateLotId didn't execute.")

            ' 若处理程序未执行，返回默认值
            Return String.Empty
        Catch ex As Exception
            ' 错误处理：记录错误并返回空字符串
            vrtReturnCode = ex.HResult
            Err.Clear()
            Return String.Empty
        End Try
    End Function

    Public Function ExecuteCSB_VerifyId(ByVal strPenId As String, Optional ByRef vrtReturnCode As Object = Nothing) As Boolean
        Dim strCSBIndex As String
        Dim bSelfRegister As Boolean

        Try
            strCSBIndex = mcol_SaxTranslations(CStr(cstrVerifyPenId)).strCSBIndex
            bSelfRegister = mcol_SaxTranslations(CStr(cstrVerifyPenId)).bSelfRegister

            ' 根据索引决定执行逻辑
            If String.IsNullOrEmpty(strCSBIndex) OrElse Not bSelfRegister Then
                ' 无模块时返回True（保持原逻辑）
                Return True
            End If

            ' 执行SAX处理程序
            Dim handler As clsSAXhandler = gcol_SAXhandlers(strCSBIndex)
            Dim paramCol As New colProtoTypeParam()
            paramCol.Add("penid", "String")
            paramCol.UpdateValue("penid", strPenId)
            paramCol.Add("partNumber", "String")
            paramCol.UpdateValue("partNumber", go_Context.PartNumber)
            paramCol.Add("runType", "String")
            paramCol.UpdateValue("runType", go_Context.RunType)
            paramCol.Add("source", "String")
            paramCol.UpdateValue("source", go_Context.Source)
            paramCol.Add("sqlNextcap", "sqlconnection")
            paramCol.Add("sqlMfg", "sqlconnection")
            paramCol.Add("accessNextcap", "oledbconnection")
            paramCol.Add("accessMfg", "oledbconnection")

            If handler IsNot Nothing AndAlso handler.bLoaded Then
                Console.WriteLine("ExecuteCSB_VerifyId executing.")
                Return CBool(handler.oHandler.Evaluate(paramCol, handler.oParams))
            End If

            Console.WriteLine("ExecuteCSB_VerifyId didn't execute.")
            ' 若处理程序未执行，返回默认值True（保持原逻辑）
            Return True
        Catch ex As Exception
            ' 错误处理：记录错误并返回False
            vrtReturnCode = ex.HResult
            Err.Clear()
            Return False
        End Try
    End Function
    Private Function IsReturnTypeCompatible(expected As String, carrier As Object) As Boolean
        ' If no constraint or “variant-like”, accept.
        If String.IsNullOrWhiteSpace(expected) Then Return True
        If expected.Equals("Variant", StringComparison.OrdinalIgnoreCase) _
       OrElse expected.Equals("Object", StringComparison.OrdinalIgnoreCase) Then
            Return True
        End If
        ' If carrier is Nothing, allow (caller may set later).
        If carrier Is Nothing Then Return True

        ' Compare against runtime type of the carrier
        Select Case expected.ToLowerInvariant()
            Case "string" : Return TypeOf carrier Is String
            Case "integer", "int32" : Return TypeOf carrier Is Integer
            Case "long", "int64" : Return TypeOf carrier Is Long Or TypeOf carrier Is Integer
            Case "double" : Return TypeOf carrier Is Double Or TypeOf carrier Is Single Or TypeOf carrier Is Decimal
            Case "decimal" : Return TypeOf carrier Is Decimal
            Case "boolean", "bool" : Return TypeOf carrier Is Boolean
            Case "date", "datetime" : Return TypeOf carrier Is Date
            Case Else : Return True ' unknown => be permissive
        End Select
    End Function
    '=======================================================
    'Routine: mdlSAX.ExcuteCSB_LotWizard(str,vrt)
    'Purpose:
    '
    'Globals:None
    '
    'Input:
    '       sParam - Context parameter
    '
    '       vrtReturnCode - Any error that was generated
    '       by the attempt to or the operation of the SAX handler.
    '
    'Return: Boolean - Whether or not the Field is asked, true = yes ask
    '
    'Modifications:
    '   07-05-2001 As written. Chris Barker
    '
    '
    '=======================================================
    Public Function ExecuteCSB_LotWizard(ByVal strParam As String, Optional ByRef vrtReturnCode As Object = Nothing) As Boolean
        Dim strCSBIndex As String
        Dim bSelfRegister As Boolean

        Try
            '' 从翻译集合获取CSB索引
            'strCSBIndex = mcol_SaxTranslations(CStr(cstrLotWizard)).strCSBIndex
            'bSelfRegister = mcol_SaxTranslations(CStr(cstrLotWizard)).bSelfRegister

            '' 根据索引决定执行逻辑
            'If String.IsNullOrEmpty(strCSBIndex) OrElse Not bSelfRegister Then
            '    ' 无模块时返回False
            '    Return False
            'Else
            '    ' 执行SAX处理程序
            '    Dim handler As clsSAXhandler = gcol_SAXhandlers(strCSBIndex)
            '    If handler IsNot Nothing AndAlso handler.bLoaded Then
            '        ' 调用Evaluate方法并返回结果（传递strParam参数）
            '        Return CBool(handler.oHandler.Evaluate(strParam))
            '    End If
            'End If

            ' 若处理程序未执行或加载失败，返回默认值False
            Return False

        Catch ex As Exception
            ' 错误处理：记录错误信息并返回False
            vrtReturnCode = New With {
                .ErrorCode = ex.HResult,
                .ErrorMessage = ex.Message,
                .StackTrace = ex.StackTrace
            }
            Return False
        End Try
    End Function
    '=======================================================
    'Routine: mdlSAX.ExecuteCSB_AugmentPenId(str,str,vrt)
    'Purpose:
    '
    'Globals:None
    '
    'Input: strCSBIndex - The index of the particular
    '       handler in the collection of SAX objects.
    '
    '       strPenId - The ID to operate on.
    '
    '       vrtReturnCode - Any error that was generated
    '       by the attempt to or the operation of the SAX handler.
    '
    'Return: String - The result of the modifications to the
    '        Pen Id
    '
    'Modifications:
    '   11-02-1998 As written for Pass1.2
    '
    '
    '=======================================================
    Public Function ExecuteCSB_AugmentPenId(ByVal strPenId As String, Optional ByRef vrtReturnCode As Object = Nothing) As String
        Dim strCSBIndex As String
        Dim bSelfRegister As Boolean

        Try
            ' 从翻译集合获取CSB索引（假设cstrAugmentPenId为字符串常量）
            strCSBIndex = mcol_SaxTranslations(CStr(cstrAugmentPenId)).strCSBIndex
            bSelfRegister = mcol_SaxTranslations(CStr(cstrAugmentPenId)).bSelfRegister
            ' 根据索引决定执行逻辑
            If String.IsNullOrEmpty(strCSBIndex) OrElse Not bSelfRegister Then

                ' 无模块时返回True（保持原逻辑）
                Return strPenId
            End If

            ' 执行SAX处理程序
            Dim handler As clsSAXhandler = gcol_SAXhandlers(strCSBIndex)
            ' Create a new colProtoTypeParam collection
            Dim paramCol As New colProtoTypeParam()

            ' Add strPenId as a parameter named "penId" (or use the appropriate name)
            paramCol.Add("penId", "String")
            paramCol.UpdateValue("penId", strPenId)

            If handler IsNot Nothing AndAlso handler.bLoaded Then
                Return (handler.oHandler.Evaluate(paramCol, handler.oParams))
            End If


            Console.WriteLine("ExecuteCSB_AugmentPenId completed successfully.")

            '' 根据索引决定执行逻辑
            'If String.IsNullOrEmpty(strCSBIndex) OrElse Not bSelfRegister Then
            '    Debug.WriteLine("AugmentPenId is not configured in RunTime Value, so pen id is from frmNextcap text box.")
            '    ' 无模块时返回原始输入值（保持原逻辑）
            '    Return strPenId
            'Else
            '    Debug.WriteLine("AugmentPenId is configured in RunTime Value, pen id is from ExecuteCSB_AugmentPenId.")
            '    ' 执行SAX处理程序
            '    Dim handler As clsSAXhandler = gcol_SAXhandlers(strCSBIndex)
            '    If handler IsNot Nothing AndAlso handler.bLoaded Then
            '        ' 调用Evaluate方法并返回结果（传递strPenId参数）
            '        Return CStr(handler.oHandler.Evaluate(strPenId))
            '    End If
            'End If

            ' 若处理程序未执行，返回默认值（保持原逻辑，即返回原始输入）

        Catch ex As Exception
            ' 错误处理：记录错误并返回空字符串
            vrtReturnCode = ex.HResult
            Err.Clear()
            Return String.Empty
        End Try
    End Function
    '=======================================================
    'Routine: mdlSAX.ExecuteCSB_PostGoodPenEntry(vrt)
    'Purpose: This merely executes the users script after
    'the Good pen entry is triggered.
    '
    'Globals: None
    '
    'Input: strCSBIndex - The index of the particular
    '       handler in the collection of SAX objects.
    '
    '       vrtReturnCode - Any error that was generated
    '       by the attempt to or the operation of the SAX handler.
    '
    'Return: None
    '
    'Modifications:
    '   01-27-1999 As written for Pass1.6
    '
    '
    '=======================================================
    Public Sub ExecuteCSB_PostGoodPenEntry(Optional ByRef vrtReturnCode As Object = Nothing)
        Dim strCSBIndex As String
        Dim bSelfRegister As Boolean

        Try
            '' 从翻译集合获取CSB索引
            'strCSBIndex = mcol_SaxTranslations(CStr(cstrPostGoodPenEntry)).strCSBIndex
            'bSelfRegister = mcol_SaxTranslations(CStr(cstrPostGoodPenEntry)).bSelfRegister

            '' 根据索引决定执行逻辑
            'If Not String.IsNullOrEmpty(strCSBIndex) AndAlso bSelfRegister Then
            '    ' 执行SAX处理程序
            '    Dim handler As clsSAXhandler = gcol_SAXhandlers(strCSBIndex)
            '    If handler IsNot Nothing AndAlso handler.bLoaded Then
            '        ' 调用无参Call方法（保持原逻辑）
            '        handler.oHandler.Evaluate()
            '    End If
            'End If
            '' 无模块时不执行任何操作（保持原逻辑）

        Catch ex As Exception
            ' 错误处理：记录错误
            vrtReturnCode = New With {
                .ErrorCode = ex.HResult,
                .Message = ex.Message,
                .StackTrace = ex.StackTrace
            }
            Err.Clear()
        End Try
    End Sub
    '=======================================================
    'Routine: mdlSAX.ExecuteCSB_PostGoodPenEntry(vrt)
    'Purpose: This merely executes the users script after
    'the Good pen entry is triggered.
    '
    'Globals: None
    '
    'Input: strCSBIndex - The index of the particular
    '       handler in the collection of SAX objects.
    '
    '       vrtReturnCode - Any error that was generated
    '       by the attempt to or the operation of the SAX handler.
    '
    'Return: None
    '
    'Modifications:
    '   01-27-1999 As written for Pass1.6
    '
    '
    '=======================================================
    Public Sub ExecuteCSB_UpdateStatusBar(Optional ByRef vrtReturnCode As Object = Nothing)
        Dim strCSBIndex As String
        Dim bSelfRegister As Boolean

        Try
            '' 从翻译集合获取CSB索引
            'strCSBIndex = mcol_SaxTranslations(CStr(cstrUpdateStatusBar)).strCSBIndex
            'bSelfRegister = mcol_SaxTranslations(CStr(cstrUpdateStatusBar)).bSelfRegister

            '' 根据索引决定执行逻辑
            'If Not String.IsNullOrEmpty(strCSBIndex) AndAlso bSelfRegister Then
            '    ' 执行SAX处理程序
            '    Dim handler As clsSAXhandler = gcol_SAXhandlers(strCSBIndex)
            '    If handler IsNot Nothing AndAlso handler.bLoaded Then
            '        ' 调用无参Call方法（保持原逻辑）
            '        handler.oHandler.Evaluate()
            '    End If
            'End If
            '' 无模块时不执行任何操作（保持原逻辑）

        Catch ex As Exception
            ' 错误处理：记录错误信息
            vrtReturnCode = New With {
                .ErrorCode = ex.HResult,
                .Message = ex.Message,
                .StackTrace = ex.StackTrace
            }
            Err.Clear()
        End Try
    End Sub
    '=======================================================
    'Routine: mdlSAX.ExecuteCSB_PreShiftChange(vrt)
    'Purpose: This merely executes the users script prior
    'to a shift change command taking place.
    '
    'Globals: None
    '
    'Input: strCSBIndex - The index of the particular
    '       handler in the collection of SAX objects.
    '
    '       vrtReturnCode - Any error that was generated
    '       by the attempt to or the operation of the SAX handler.
    '
    'Return: None
    '
    'Modifications:
    '   01-28-1999 As written for Pass1.7
    '
    '
    '=======================================================
    Public Sub ExecuteCSB_PreShiftChange(Optional ByRef vrtReturnCode As Object = Nothing)
        Dim strCSBIndex As String
        Dim bSelfRegister As Boolean

        Try
            '' 从翻译集合获取CSB索引
            'strCSBIndex = mcol_SaxTranslations(CStr(cstrPreShiftChange)).strCSBIndex
            'bSelfRegister = mcol_SaxTranslations(CStr(cstrPreShiftChange)).bSelfRegister

            '' 根据索引决定执行逻辑
            'If Not String.IsNullOrEmpty(strCSBIndex) AndAlso bSelfRegister Then
            '    ' 执行SAX处理程序
            '    Dim handler As clsSAXhandler = gcol_SAXhandlers(strCSBIndex)
            '    If handler IsNot Nothing AndAlso handler.bLoaded Then
            '        ' 调用无参Call方法（保持原逻辑）
            '        handler.oHandler.Evaluate()
            '    End If
            'End If
            '' 无模块时不执行任何操作（保持原逻辑）

        Catch ex As Exception
            ' 错误处理：记录错误信息
            vrtReturnCode = New With {
                .ErrorCode = ex.HResult,
                .Message = ex.Message,
                .StackTrace = ex.StackTrace
            }
            Err.Clear()
        End Try
    End Sub
    '=======================================================
    'Routine: mdlSAX.ExecuteCSB_PostShiftChange(vrt)
    'Purpose: This merely executes the users script prior
    'to a shift change command taking place.
    '
    'Globals: None
    '
    'Input: strCSBIndex - The index of the particular
    '       handler in the collection of SAX objects.
    '
    '       vrtReturnCode - Any error that was generated
    '       by the attempt to or the operation of the SAX handler.
    '
    'Return: None
    '
    'Modifications:
    '   01-28-1999 As written for Pass1.7
    '
    '
    '=======================================================
    Public Sub ExecuteCSB_PostShiftChange(Optional ByRef vrtReturnCode As Object = Nothing)
        Dim strCSBIndex As String
        Dim bSelfRegister As Boolean

        Try
            '' 从翻译集合获取CSB索引
            'strCSBIndex = mcol_SaxTranslations(CStr(cstrPostShiftChange)).strCSBIndex
            'bSelfRegister = mcol_SaxTranslations(CStr(cstrPostShiftChange)).bSelfRegister

            '' 根据索引决定执行逻辑
            'If Not String.IsNullOrEmpty(strCSBIndex) AndAlso bSelfRegister Then
            '    ' 执行SAX处理程序
            '    Dim handler As clsSAXhandler = gcol_SAXhandlers(strCSBIndex)
            '    If handler IsNot Nothing AndAlso handler.bLoaded Then
            '        ' 调用无参Call方法（保持原逻辑）
            '        handler.oHandler.Evaluate()
            '    End If
            'End If
            '' 无模块时不执行任何操作（保持原逻辑）

        Catch ex As Exception
            ' 错误处理：记录错误信息
            vrtReturnCode = New With {
                .ErrorCode = ex.HResult,
                .Message = ex.Message,
                .StackTrace = ex.StackTrace
            }
            Err.Clear()
        End Try
    End Sub
    '=======================================================
    'Routine: mdlSAX.ExecuteCSB_PostBadPenEntry(vrt)
    'Purpose: This merely executes the users script after
    'a bad pen is entered.
    '
    'Globals: None
    '
    'Input: strCSBIndex - The index of the particular
    '       handler in the collection of SAX objects.
    '
    '       vrtReturnCode - Any error that was generated
    '       by the attempt to or the operation of the SAX handler.
    '
    'Return: None
    '
    'Modifications:
    '   02-01-1999 As written for Pass1.7
    '
    '
    '=======================================================
    Public Sub ExecuteCSB_PostBadPenEntry(Optional ByRef vrtReturnCode As Object = Nothing)
        Dim strCSBIndex As String
        Dim bSelfRegister As Boolean

        Try
            '' 从翻译集合获取CSB索引
            'strCSBIndex = mcol_SaxTranslations(CStr(cstrPostBadPenEntry)).strCSBIndex
            'bSelfRegister = mcol_SaxTranslations(CStr(cstrPostBadPenEntry)).bSelfRegister

            '' 根据索引决定执行逻辑
            'If Not String.IsNullOrEmpty(strCSBIndex) AndAlso bSelfRegister Then
            '    ' 执行SAX处理程序
            '    Dim handler As clsSAXhandler = gcol_SAXhandlers(strCSBIndex)
            '    If handler IsNot Nothing AndAlso handler.bLoaded Then
            '        ' 调用无参Call方法（保持原逻辑）
            '        handler.oHandler.Evaluate()
            '    End If
            'End If
            '' 无模块时不执行任何操作（保持原逻辑）

        Catch ex As Exception
            ' 错误处理：记录错误信息
            vrtReturnCode = New With {
                .ErrorCode = ex.HResult,
                .Message = ex.Message,
                .StackTrace = ex.StackTrace
            }
            Err.Clear()
        End Try
    End Sub
    '=======================================================
    'Routine: mdlSAX.ExcuteCSB_ContextLock(str,vrt)
    'Purpose:
    '
    'Globals:None
    '
    'Input: strCSBIndex - The index of the particular
    '       handler in the collection of SAX objects.
    '
    '       strLotId - The Pen Id to validate
    '
    '       vrtReturnCode - Any error that was generated
    '       by the attempt to or the operation of the SAX handler.
    '
    'Return: Boolean - Whether or not the Field is enabled, true = disable
    '
    'Modifications:
    '   07-05-2001 As written. Chris Barker
    '
    '
    '=======================================================
    Public Function ExecuteCSB_ContextLock(ByVal strParam As String, Optional ByRef vrtReturnCode As Object = Nothing) As Boolean
        Dim strCSBIndex As String
        Dim bSelfRegister As Boolean

        Try
            '' 从翻译集合获取CSB索引
            'strCSBIndex = mcol_SaxTranslations(CStr(cstrContextLock)).strCSBIndex
            'bSelfRegister = mcol_SaxTranslations(CStr(cstrContextLock)).bSelfRegister

            '' 根据索引决定执行逻辑
            'If String.IsNullOrEmpty(strCSBIndex) OrElse Not bSelfRegister Then
            '    ' 无模块时返回False（保持原逻辑）
            '    Return False
            'Else
            '    ' 执行SAX处理程序
            '    Dim handler As clsSAXhandler = gcol_SAXhandlers(strCSBIndex)
            '    If handler IsNot Nothing AndAlso handler.bLoaded Then
            '        ' 调用Evaluate方法并返回结果（传递strParam参数）
            '        Return CBool(handler.oHandler.Evaluate(strParam))
            '    End If
            'End If

            ' 若处理程序未执行，返回默认值False
            Return False
        Catch ex As Exception
            ' 错误处理：记录错误并返回False
            vrtReturnCode = ex.HResult
            Err.Clear()
            Return False
        End Try
    End Function
    '=======================================================
    'Routine: mdlSAX.ExecuteCSB_ClientSystemStartUp(vrt)
    'Purpose: This merely executes the users script after
    'at the end of startup in mdlMain.Main()
    '
    'Globals: None
    '
    'Input: strCSBIndex - The index of the particular
    '       handler in the collection of SAX objects.
    '
    '       vrtReturnCode - Any error that was generated
    '       by the attempt to or the operation of the SAX handler.
    '
    'Return: None
    '
    'Modifications:
    '   08-03-2001 v:1.1.1 As written Chris Barker
    '
    '
    '=======================================================
    Public Sub ExecuteCSB_ClientSystemStartUp(Optional ByRef vrtReturnCode As Object = Nothing)
        Dim strCSBIndex As String
        Dim bSelfRegister As Boolean

        Try
            ' 从翻译集合获取CSB索引
            strCSBIndex = mcol_SaxTranslations(CStr(cstrClientSystemStartUp)).strCSBIndex
            bSelfRegister = mcol_SaxTranslations(CStr(cstrClientSystemStartUp)).bSelfRegister

            ' 根据索引决定执行逻辑
            If String.IsNullOrEmpty(strCSBIndex) OrElse Not bSelfRegister Then
                ' 无模块时返回True（保持原逻辑）

            Else
                ' 执行SAX处理程序
                Dim handler As clsSAXhandler = gcol_SAXhandlers(strCSBIndex)
                Dim paramCol As New colProtoTypeParam()
                paramCol.Add("sqlMfg", "sqlconnection")
                paramCol.Add("accessMfg", "oledbconnection")

                If handler IsNot Nothing AndAlso handler.bLoaded Then
                    handler.oHandler.Evaluate()
                End If
                Console.WriteLine("ExecuteCSB_ClientSystemStartUp completed successfully.")
            End If
        Catch ex As Exception
            ' 错误处理：记录错误信息
            vrtReturnCode = New With {
                .ErrorCode = ex.HResult,
                .Message = ex.Message,
                .StackTrace = ex.StackTrace
            }
            Err.Clear()
        End Try
    End Sub
    '=======================================================
    'Routine: mdlSAX.ExecuteCSB_LotManagerChange(str,dt,str,vrt)
    'Purpose: This executes after a LotManager change occurs
    'as prompted by the Context Editor and the Context Wizard.
    '
    'Globals: None
    '
    'Input: strLotId - The lot being changed
    '
    '       dtBirth - The brirthday of the lot
    '
    '       sWho - Where the event is generating from.
    '
    '       vrtReturnCode - Any error that was generated
    '       by the attempt to or the operation of the SAX handler.
    '
    'Return: None
    '
    'Modifications:
    '   08-24-2001 v:1.1.5 As written Chris Barker
    '
    '
    '=======================================================
    Public Sub ExecuteCSB_LotManagerChange(
    ByRef strLotId As String,
    ByRef dtBirth As Date,
    ByRef sWho As String,
    Optional ByRef vrtReturnCode As Object = Nothing
)
        Dim strCSBIndex As String
        Dim bSelfRegister As Boolean

        Try

            ' 从翻译集合获取CSB索引
            'strCSBIndex = mcol_SaxTranslations(CStr(cstrLotManagerChange)).strCSBIndex
            'bSelfRegister = mcol_SaxTranslations(CStr(cstrLotManagerChange)).bSelfRegister

            '' 根据索引决定执行逻辑
            'If Not String.IsNullOrEmpty(strCSBIndex) AndAlso bSelfRegister Then
            '    ' 执行SAX处理程序
            '    Dim handler As clsSAXhandler = gcol_SAXhandlers(strCSBIndex)
            '    If handler IsNot Nothing AndAlso handler.bLoaded Then
            '        ' 调用带参Call方法（传递strLotId, dtBirth, sWho参数）
            '        handler.oHandler.Evaluate(strLotId, dtBirth, sWho)
            '    End If
            'End If
            ' 无模块时不执行任何操作（保持原逻辑）

        Catch ex As Exception
            ' 错误处理：记录错误信息
            vrtReturnCode = New With {
                .ErrorCode = ex.HResult,
                .Message = ex.Message,
                .StackTrace = ex.StackTrace
            }
            Err.Clear()
        End Try
    End Sub

    Public ReadOnly Property ContextLockUsed() As Boolean


        Get

            ' 检查集合中是否存在指定键且对应值的 strCSBIndex 不为空
            If Not String.IsNullOrEmpty(mcol_SaxTranslations(cstrContextLock).strCSBIndex) Then
                Return True
            Else
                Return False
            End If
        End Get
    End Property

End Module
