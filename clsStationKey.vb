Public Class clsStationKey
    ' 私有字段用于存储属性值
    Private mvarstrLineType As String
    Private mvarstrSource As String
    Private mvarnLineNumber As Integer

    ' 私有数组字段
    Private m_strAccumulators() As String
    Private m_bAccumulatorsUsed As Boolean
    Private m_strItemTypes() As String
    Private m_bItemTypesUsed As Boolean
    Private m_strRunTypes() As String
    Private m_bRunTypesUsed As Boolean
    Private m_strTestBedTypes() As String
    Private m_bTestBedTypesUsed As Boolean
    Private m_strTestBeds() As String
    Private m_bTestBedsUsed As Boolean
    Private m_strPartName() As String
    Private m_strPartNumber() As String
    Private m_strPartFamily() As String
    Private m_bPartsUsed As Boolean
    Private m_strLotComments() As String
    Private m_bLotCommentsUsed As Boolean
    Private m_strDefectClasses() As String
    Private m_bDefectClassesUsed As Boolean

    ' nLineNumber 属性
    Public Property nLineNumber As Integer
        Get
            Return mvarnLineNumber
        End Get
        Set(ByVal value As Integer)
            mvarnLineNumber = value
        End Set
    End Property

    ' strSource 属性
    Public Property strSource As String
        Get
            Return mvarstrSource
        End Get
        Set(ByVal value As String)
            mvarstrSource = value
        End Set
    End Property

    ' strLineType 属性
    Public Property strLineType As String
        Get
            Return mvarstrLineType
        End Get
        Set(ByVal value As String)
            mvarstrLineType = value
        End Set
    End Property

    ' Accumulators 属性
    Public Property Accumulators(ByVal nIndex As Integer) As String
        Get
            If m_bAccumulatorsUsed AndAlso nIndex < m_strAccumulators.Length Then
                Return m_strAccumulators(nIndex)
            End If
            Return String.Empty
        End Get
        Set(ByVal value As String)
            If m_bAccumulatorsUsed Then
                If nIndex >= m_strAccumulators.Length Then
                    Array.Resize(m_strAccumulators, nIndex + 1)
                End If
                m_strAccumulators(nIndex) = value
            Else
                ReDim m_strAccumulators(nIndex)
                m_strAccumulators(nIndex) = value
                m_bAccumulatorsUsed = True
            End If
        End Set
    End Property

    ' 重置 Accumulators 数组的方法
    Public Sub ResetAccumulators()
        m_bAccumulatorsUsed = False
    End Sub

    ' ItemTypes 属性
    Public Property ItemTypes(ByVal nIndex As Integer) As String
        Get
            If m_bItemTypesUsed AndAlso nIndex < m_strItemTypes.Length Then
                Return m_strItemTypes(nIndex)
            End If
            Return String.Empty
        End Get
        Set(ByVal value As String)
            If m_bItemTypesUsed Then
                If nIndex >= m_strItemTypes.Length Then
                    Array.Resize(m_strItemTypes, nIndex + 1)
                End If
                m_strItemTypes(nIndex) = value
            Else
                ReDim m_strItemTypes(nIndex)
                m_strItemTypes(nIndex) = value
                m_bItemTypesUsed = True
            End If
        End Set
    End Property

    ' 新增方法，用于获取 ItemTypes 数组的长度
    Public Function GetItemTypesLength() As Integer
        If m_bItemTypesUsed Then
            Return m_strItemTypes.Length
        End If
        Return 0
    End Function

    ' 重置 ItemTypes 数组的方法
    Public Sub ResetItemTypes()
        m_bItemTypesUsed = False
    End Sub

    ' RunTypes 属性
    Public Property RunTypes(ByVal nIndex As Integer) As String
        Get
            If m_bRunTypesUsed AndAlso nIndex < m_strRunTypes.Length Then
                Return m_strRunTypes(nIndex)
            End If
            Return String.Empty
        End Get
        Set(ByVal value As String)
            If m_bRunTypesUsed Then
                If nIndex >= m_strRunTypes.Length Then
                    Array.Resize(m_strRunTypes, nIndex + 1)
                End If
                m_strRunTypes(nIndex) = value
            Else
                ReDim m_strRunTypes(nIndex)
                m_strRunTypes(nIndex) = value
                m_bRunTypesUsed = True
            End If
        End Set
    End Property

    ' 重置 RunTypes 数组的方法
    Public Sub ResetRunTypes()
        m_bRunTypesUsed = False
    End Sub

    ' TestBeds 属性
    Public Property TestBeds(ByVal nIndex As Integer) As String
        Get
            If m_bTestBedsUsed AndAlso nIndex < m_strTestBeds.Length Then
                Return m_strTestBeds(nIndex)
            End If
            Return String.Empty
        End Get
        Set(ByVal value As String)
            If m_bTestBedsUsed Then
                If nIndex >= m_strTestBeds.Length Then
                    Array.Resize(m_strTestBeds, nIndex + 1)
                End If
                m_strTestBeds(nIndex) = value
            Else
                ReDim m_strTestBeds(nIndex)
                m_strTestBeds(nIndex) = value
                m_bTestBedsUsed = True
            End If
        End Set
    End Property

    ' 重置 TestBeds 数组的方法
    Public Sub ResetTestBeds()
        m_bTestBedsUsed = False
    End Sub

    ' TestBedTypes 属性
    Public Property TestBedTypes(ByVal nIndex As Integer) As String
        Get
            If m_bTestBedTypesUsed AndAlso nIndex < m_strTestBedTypes.Length Then
                Return m_strTestBedTypes(nIndex)
            End If
            Return String.Empty
        End Get
        Set(ByVal value As String)
            If m_bTestBedTypesUsed Then
                If nIndex >= m_strTestBedTypes.Length Then
                    Array.Resize(m_strTestBedTypes, nIndex + 1)
                End If
                m_strTestBedTypes(nIndex) = value
            Else
                ReDim m_strTestBedTypes(nIndex)
                m_strTestBedTypes(nIndex) = value
                m_bTestBedTypesUsed = True
            End If
        End Set
    End Property

    ' LotComments 属性
    Public Property LotComments(ByVal nIndex As Integer) As String
        Get
            If m_bLotCommentsUsed AndAlso nIndex < m_strLotComments.Length Then
                Return m_strLotComments(nIndex)
            End If
            Return String.Empty
        End Get
        Set(ByVal value As String)
            If m_bLotCommentsUsed Then
                If nIndex >= m_strLotComments.Length Then
                    Array.Resize(m_strLotComments, nIndex + 1)
                End If
                m_strLotComments(nIndex) = value
            Else
                ReDim m_strLotComments(nIndex)
                m_strLotComments(nIndex) = value
                m_bLotCommentsUsed = True
            End If
        End Set
    End Property

    ' 重置 LotComments 数组的方法
    Public Sub ResetLotComments()
        m_bLotCommentsUsed = False
    End Sub

    ' DefectClasses 属性
    Public Property DefectClasses(ByVal nIndex As Integer) As String
        Get
            If m_bDefectClassesUsed AndAlso nIndex < m_strDefectClasses.Length Then
                Return m_strDefectClasses(nIndex)
            End If
            Return String.Empty
        End Get
        Set(ByVal value As String)
            If m_bDefectClassesUsed Then
                If nIndex >= m_strDefectClasses.Length Then
                    Array.Resize(m_strDefectClasses, nIndex + 1)
                End If
                m_strDefectClasses(nIndex) = value
            Else
                ReDim m_strDefectClasses(nIndex)
                m_strDefectClasses(nIndex) = value
                m_bDefectClassesUsed = True
            End If
        End Set
    End Property

    ' 重置 DefectClasses 数组的方法
    Public Sub ResetDefectClasses()
        m_bDefectClassesUsed = False
    End Sub

    ' LetParts 方法转换为设置 Parts 相关数组的方法
    Public Sub LetParts(ByVal strName As String, ByVal strNumber As String, ByVal strFamily As String)
        If m_bPartsUsed Then
            Dim nUbound As Integer = m_strPartName.Length
            Array.Resize(m_strPartName, nUbound + 1)
            Array.Resize(m_strPartNumber, nUbound + 1)
            Array.Resize(m_strPartFamily, nUbound + 1)
            m_strPartName(nUbound) = strName
            m_strPartNumber(nUbound) = strNumber
            m_strPartFamily(nUbound) = strFamily
        Else
            ReDim m_strPartName(0)
            ReDim m_strPartNumber(0)
            ReDim m_strPartFamily(0)
            m_strPartName(0) = strName
            m_strPartNumber(0) = strNumber
            m_strPartFamily(0) = strFamily
            m_bPartsUsed = True
        End If
    End Sub

    ' GetParts 方法转换为获取 Parts 相关数组值的方法
    Public Sub GetParts(ByVal nIndex As Integer, ByRef strName As String, ByRef strNumber As String, ByRef strFamily As String)
        If m_bPartsUsed AndAlso nIndex < m_strPartName.Length Then
            strName = m_strPartName(nIndex)
            strNumber = m_strPartNumber(nIndex)
            strFamily = m_strPartFamily(nIndex)
        Else
            strName = String.Empty
            strNumber = String.Empty
            strFamily = String.Empty
        End If
    End Sub

    ' 重置 Parts 数组的方法
    Public Sub ResetParts()
        m_bPartsUsed = False
    End Sub
End Class