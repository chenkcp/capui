Imports System.Collections
Imports System.Collections.Generic

Public Class colLotMangerQueue
    Implements IEnumerable(Of clsLotManagerQueue)

    ' 本地变量保存泛型列表
    Private mList As New List(Of clsLotManagerQueue)
    Private mKeyLookup As New Dictionary(Of String, clsLotManagerQueue)

    ' 添加项到集合
    Public Function Add(sLotId As String, nFrom As Integer, nTo As Integer, dtBirth As DateTime, Optional sKey As String = "") As clsLotManagerQueue
        ' 创建新对象
        Dim objNewMember As New clsLotManagerQueue()

        ' 设置属性
        objNewMember.sLotId = sLotId
        objNewMember.nFrom = nFrom
        objNewMember.nTo = nTo
        objNewMember.dtBirth = dtBirth
        Console.WriteLine(objNewMember.dtBirth.ToString("yyyy-MM-dd HH:mm:ss"))
        ' 添加到列表和键查找字典
        mList.Add(objNewMember)
        If Not String.IsNullOrEmpty(sKey) Then
            mKeyLookup(sKey) = objNewMember
        End If

        ' 返回创建的对象
        Return objNewMember
    End Function

    ' 按指定属性倒序排序
    Public Sub SortDescending(Optional sortBy As String = "sLotId")
        ' 根据指定属性排序后反转（倒序）
        Select Case sortBy.ToLower()
            Case "slotid"
                mList.Sort(Function(x, y) String.Compare(x.sLotId, y.sLotId))
            Case "nfrom"
                mList.Sort(Function(x, y) x.nFrom.CompareTo(y.nFrom))
            Case "nto"
                mList.Sort(Function(x, y) x.nTo.CompareTo(y.nTo))
            Case "dtbirth"
                mList.Sort(Function(x, y) x.dtBirth.CompareTo(y.dtBirth))
            Case Else
                ' 默认按sLotId排序
                mList.Sort(Function(x, y) String.Compare(x.sLotId, y.sLotId))
        End Select

        ' 反转列表实现倒序
        mList.Reverse()
    End Sub

    ' 通过索引或键获取项
    Default Public ReadOnly Property Item(vntIndexKey As Object) As clsLotManagerQueue
        Get
            If TypeOf vntIndexKey Is String Then
                Return mKeyLookup(CStr(vntIndexKey))
            ElseIf TypeOf vntIndexKey Is Integer Then
                Return mList(CInt(vntIndexKey)) ' 调整为 0 基索引
            Else
                Throw New ArgumentException("索引必须是整数或字符串")
            End If
        End Get
    End Property

    ' 获取集合项数
    Public ReadOnly Property Count() As Integer
        Get
            Return mList.Count
        End Get
    End Property

    ' 移除项
    Public Sub Remove(vntIndexKey As Object)
        If TypeOf vntIndexKey Is String Then
            Dim key As String = CStr(vntIndexKey)
            If mKeyLookup.ContainsKey(key) Then
                mList.Remove(mKeyLookup(key))
                mKeyLookup.Remove(key)
            End If
        ElseIf TypeOf vntIndexKey Is Integer Then
            Dim index As Integer = CInt(vntIndexKey) ' 调整为 0 基索引
            If index >= 0 AndAlso index < mList.Count Then
                Dim item = mList(index)
                mList.RemoveAt(index)

                ' 清理键查找
                Dim keyToRemove As String = Nothing
                For Each kvp In mKeyLookup
                    If kvp.Value Is item Then
                        keyToRemove = kvp.Key
                        Exit For
                    End If
                Next
                If keyToRemove IsNot Nothing Then
                    mKeyLookup.Remove(keyToRemove)
                End If
            End If
        Else
            Throw New ArgumentException("索引必须是整数或字符串")
        End If
    End Sub

    ' 实现 IEnumerable 接口
    Public Function GetEnumerator() As IEnumerator(Of clsLotManagerQueue) Implements IEnumerable(Of clsLotManagerQueue).GetEnumerator
        Return mList.GetEnumerator()
    End Function

    Private Function GetEnumerator1() As IEnumerator Implements IEnumerable.GetEnumerator
        Return Me.GetEnumerator()
    End Function
End Class
