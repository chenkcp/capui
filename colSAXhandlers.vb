Imports System.Collections.Generic
Imports System.Runtime.InteropServices
Public Class colSAXhandlers
    Implements IEnumerable(Of clsSAXhandler)

    Private mList As New List(Of clsSAXhandler)
    Private mDictionary As New Dictionary(Of String, clsSAXhandler)(StringComparer.OrdinalIgnoreCase)

    ' 添加元素
    Public Function Add(bLoaded As Boolean, strType As String, strCode As String,
                        strFileName As String, oParams As colProtoTypeParam, oHandler As clsSAXEngineHandler, strProtoType As String,
                        strLocation As String, Optional sKey As String = Nothing) As clsSAXhandler

        Dim objNewMember As New clsSAXhandler() With {
            .bLoaded = bLoaded,
            .strType = strType,
            .strCode = strCode,
            .strFileName = strFileName,
            .oHandler = oHandler,
            .strProtoType = strProtoType,
            .strLocation = strLocation,
            .strIndex = sKey,
            .oParams = If(oParams, New Dictionary(Of String, Object)()) ' never Nothing
        }

        ' maintain order
        mList.Add(objNewMember)

        ' index by key if provided
        If Not String.IsNullOrEmpty(sKey) Then
            mDictionary(sKey) = objNewMember
        End If

        Return objNewMember
    End Function

    ' 索引器（支持索引或键访问）
    Default Public ReadOnly Property Item(vntIndexKey As Object) As clsSAXhandler
        Get
            If TypeOf vntIndexKey Is Integer Then
                Dim index = CInt(vntIndexKey)
                If index >= 1 AndAlso index <= mList.Count Then
                    Return mList(index - 1)
                End If
                Return Nothing ' instead of throwing
            ElseIf TypeOf vntIndexKey Is String Then
                Dim key = CStr(vntIndexKey)
                Dim h As clsSAXhandler = Nothing
                If mDictionary.TryGetValue(key, h) Then Return h
                Return Nothing ' instead of throwing
            Else
                Throw New ArgumentException("参数必须是整数或字符串")
            End If
        End Get
    End Property
    Public Function TryGetByKey(key As String, ByRef handler As clsSAXhandler) As Boolean
        handler = Nothing
        If String.IsNullOrEmpty(key) Then Return False
        Return mDictionary.TryGetValue(key, handler)
    End Function

    Public Function ContainsKey(key As String) As Boolean
        If String.IsNullOrEmpty(key) Then Return False
        Return mDictionary.ContainsKey(key)
    End Function
    ' 获取元素数量
    Public ReadOnly Property Count As Integer
        Get
            Return mList.Count
        End Get
    End Property

    ' 移除元素（支持索引或键）
    Public Sub Remove(vntIndexKey As Object)
        If TypeOf vntIndexKey Is Integer Then
            Dim index As Integer = CInt(vntIndexKey)
            ' 检查索引范围（1-based）
            If index >= 1 AndAlso index <= mList.Count Then
                Dim item = mList(index - 1)
                mList.RemoveAt(index - 1)

                ' 从字典中移除对应的键（如果存在）
                Dim keyToRemove As String = Nothing
                For Each kvp In mDictionary
                    If kvp.Value Is item Then
                        keyToRemove = kvp.Key
                        Exit For
                    End If
                Next

                If Not String.IsNullOrEmpty(keyToRemove) Then
                    mDictionary.Remove(keyToRemove)
                End If
            Else
                Throw New IndexOutOfRangeException("索引超出范围")
            End If
        ElseIf TypeOf vntIndexKey Is String Then
            Dim key As String = CStr(vntIndexKey)
            ' 检查键是否存在
            If mDictionary.ContainsKey(key) Then
                Dim item = mDictionary(key)
                mDictionary.Remove(key)
                mList.Remove(item)
            Else
                Throw New KeyNotFoundException("键未找到")
            End If
        Else
            Throw New ArgumentException("参数必须是整数或字符串")
        End If
    End Sub

    ' 实现枚举接口
    Public Function GetEnumerator() As IEnumerator(Of clsSAXhandler) Implements IEnumerable(Of clsSAXhandler).GetEnumerator
        Return mList.GetEnumerator()
    End Function

    Private Function IEnumerable_GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
        Return GetEnumerator()
    End Function

    ' 构造函数（替代 Class_Initialize）
    Public Sub New()
        mList = New List(Of clsSAXhandler)
        mDictionary = New Dictionary(Of String, clsSAXhandler)(StringComparer.OrdinalIgnoreCase)
    End Sub

    ' 析构函数（替代 Class_Terminate）
    Protected Overrides Sub Finalize()
        mList.Clear()
        mDictionary.Clear()
        MyBase.Finalize()
    End Sub

End Class
