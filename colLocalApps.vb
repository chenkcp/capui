Imports System.Runtime.InteropServices
Imports System.Collections.Generic
Imports System.Runtime.CompilerServices

Public Class colLocalApps
    Private mCol As New List(Of clsLocalApp)

    ''' <summary>
    ''' 向集合添加新的 clsLocalApp 对象
    ''' </summary>
    Public Function Add(
        strDesc As String,
        strType As String,
        strIcon As String,
        strApp As String,
        bToolBar As Boolean,
        Optional strCSBStatus As String = "",
        Optional strCSBIndex As String = "",
        Optional sKey As String = ""
    ) As clsLocalApp

        Dim newItem As New clsLocalApp()
        With newItem
            .strDesc = strDesc
            .strType = strType
            .strIcon = strIcon
            .strApp = strApp
            .bToolBar = bToolBar
            .strCSBStatus = strCSBStatus
            .strCSBIndex = strCSBIndex
        End With

        mCol.Add(newItem)
        Return newItem
    End Function

    ''' <summary>
    ''' 按索引或键获取集合中的对象
    ''' </summary>
    Default Public ReadOnly Property Item(vntIndexKey As Object) As clsLocalApp
        Get
            If TypeOf vntIndexKey Is Integer Then
                Dim index As Integer = CInt(vntIndexKey) - 1
                If index >= 0 AndAlso index < mCol.Count Then
                    Return mCol(index)
                End If
            End If
            Throw New ArgumentException("索引无效", "vntIndexKey")
        End Get
    End Property

    ''' <summary>
    ''' 获取集合中的元素数量
    ''' </summary>
    Public ReadOnly Property Count As Integer
        Get
            Return mCol.Count
        End Get
    End Property

    ''' <summary>
    ''' 从集合中移除指定索引的对象
    ''' </summary>
    Public Sub Remove(vntIndexKey As Object)
        If TypeOf vntIndexKey Is Integer Then
            Dim index As Integer = CInt(vntIndexKey) - 1
            If index >= 0 AndAlso index < mCol.Count Then
                mCol.RemoveAt(index)
            End If
        End If
    End Sub

    ''' <summary>
    ''' 获取集合的枚举器，支持 For Each 循环
    ''' </summary>
    Public Function GetEnumerator() As IEnumerator(Of clsLocalApp)
        Return mCol.GetEnumerator()
    End Function

    ''' <summary>
    ''' 获取 COM 兼容的枚举器（支持 For Each 循环）
    ''' </summary>
    <ComVisible(True), DispId(-4)>
    Private Function GetIUnknownEnumerator() As System.Runtime.InteropServices.ComTypes.IEnumVARIANT
        Dim enumerable = CType(mCol, IEnumerable)
        Return CType(enumerable.GetEnumerator(), System.Runtime.InteropServices.ComTypes.IEnumVARIANT)
    End Function

    ''' <summary>
    ''' 类初始化时创建集合
    ''' </summary>
    Private Sub Class_Initialize()
        mCol = New List(Of clsLocalApp)
    End Sub

    ''' <summary>
    ''' 类终止时释放资源
    ''' </summary>
    Protected Overrides Sub Finalize()
        mCol = Nothing
        MyBase.Finalize()
    End Sub
End Class