Public Class colStationKeys
    ' 使用 Dictionary 来存储 clsStationKey 对象
    Private mCol As New Dictionary(Of String, clsStationKey)

    Public Function Add(ByVal strLineType As String, ByVal strSource As String, ByVal nLineNumber As Integer, Optional ByVal sKey As String = "") As clsStationKey
        ' 创建新对象
        Dim objNewMember As New clsStationKey()

        ' 设置对象属性
        objNewMember.strLineType = strLineType
        objNewMember.strSource = strSource
        objNewMember.nLineNumber = nLineNumber

        ' 添加到上下文列表
        With frmAllContext.GetInstance()
            .AddLineType(strLineType)
            .AddLineNumber(nLineNumber.ToString())
            .AddSource(strSource)
        End With

        ' 如果没有提供键，则生成一个默认键
        If String.IsNullOrEmpty(sKey) Then
            sKey = $"{strLineType}_{strSource}_{nLineNumber}"
        End If

        ' 将对象添加到字典
        mCol.Add(sKey, objNewMember)

        Return objNewMember
    End Function

    Default Public ReadOnly Property Item(ByVal vntIndexKey As Object) As clsStationKey
        Get
            If TypeOf vntIndexKey Is String Then
                Dim key As String = CStr(vntIndexKey)
                If mCol.ContainsKey(key) Then
                    Return mCol(key)
                End If
            ElseIf TypeOf vntIndexKey Is Integer Then
                Dim index As Integer = CInt(vntIndexKey)
                If index >= 0 AndAlso index < mCol.Count Then
                    Return mCol.ElementAt(index).Value
                End If
            End If
            Return Nothing
        End Get
    End Property

    Public ReadOnly Property Count() As Integer
        Get
            Return mCol.Count
        End Get
    End Property

    Public Sub Remove(ByVal vntIndexKey As Object)
        If TypeOf vntIndexKey Is String Then
            Dim key As String = CStr(vntIndexKey)
            If mCol.ContainsKey(key) Then
                mCol.Remove(key)
            End If
        ElseIf TypeOf vntIndexKey Is Integer Then
            Dim index As Integer = CInt(vntIndexKey)
            If index >= 0 AndAlso index < mCol.Count Then
                Dim keyToRemove = mCol.ElementAt(index).Key
                mCol.Remove(keyToRemove)
            End If
        End If
    End Sub

    Public Function GetEnumerator() As IEnumerator(Of KeyValuePair(Of String, clsStationKey))
        Return mCol.GetEnumerator()
    End Function

    Public Function IsKeyValid(ByVal LineType As String, ByVal LineNumber As Integer, ByVal Source As String) As Boolean
        Dim key As String = $"{LineType}_{Source}_{LineNumber}"
        Return mCol.ContainsKey(key)
    End Function

    Public Function GetDefaultItemType(ByVal LineType As String, ByVal LineNumber As Integer, ByVal Source As String) As String
        Dim key As String = $"{LineType}_{Source}_{LineNumber}"
        If mCol.ContainsKey(key) Then
            Dim stationKey As clsStationKey = mCol(key)
            ' 新增一个方法来获取 ItemTypes 数组的长度
            Dim itemTypesLength = stationKey.GetItemTypesLength()
            If itemTypesLength > 0 Then
                ' 假设要获取第一个元素，索引为 0
                Return stationKey.ItemTypes(0)
            End If
        End If
        Return "Unknown"
    End Function
End Class