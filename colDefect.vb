Public Class colDefect

    Private mCol As New Collection()

    Public Function Add(lngSeverity As Long, bPrimary As Boolean, nFeatureClass As Integer, nFeatureIndex As Integer, nSubFeatureIndex As Integer, strFeatureCd As String, strFeatureDesc As String, strSubFeatureCd As String, strSubFeatureDesc As String, strCauseCd As String, strCauseDesc As String, strSubCauseCd As String, strSubCauseDesc As String, strFeatureClassCd As String, strFeatureClassDesc As String, strComment As String, strNumericInput As String, Optional nEnteredOrder As Integer = 0, Optional sKey As String = "") As clsDefect
        ' Create a new object
        Dim objNewMember As New clsDefect()

        ' Set the properties passed into the method
        objNewMember.lngSeverity = lngSeverity
        objNewMember.bPrimary = bPrimary
        objNewMember.nFeatureClass = nFeatureClass
        objNewMember.nFeatureIndex = nFeatureIndex
        objNewMember.nSubFeatureIndex = nSubFeatureIndex
        objNewMember.strFeatureCd = strFeatureCd
        objNewMember.strFeatureDesc = strFeatureDesc
        objNewMember.strSubFeatureCd = strSubFeatureCd
        objNewMember.strSubFeatureDesc = strSubFeatureDesc
        objNewMember.strCauseCd = strCauseCd
        objNewMember.strCauseDesc = strCauseDesc
        objNewMember.strSubCauseCd = strSubCauseCd
        objNewMember.strSubCauseDesc = strSubCauseDesc
        objNewMember.strFeatureClassCd = strFeatureClassCd
        objNewMember.strFeatureClassDesc = strFeatureClassDesc
        objNewMember.strComment = strComment
        objNewMember.strNumericInput = strNumericInput
        objNewMember.nEnteredOrder = nEnteredOrder

        If String.IsNullOrEmpty(sKey) Then
            mCol.Add(objNewMember)
        Else
            mCol.Add(objNewMember, sKey)
        End If

        ' Return the object created
        Return objNewMember
    End Function

    Default Public ReadOnly Property Item(vntIndexKey As Object) As clsDefect
        Get
            Return mCol(vntIndexKey)
        End Get
    End Property

    Public ReadOnly Property Count() As Long
        Get
            Return mCol.Count
        End Get
    End Property

    Public Sub Remove(vntIndexKey As Object)
        mCol.Remove(vntIndexKey)
    End Sub

    Public ReadOnly Property NewEnum() As IEnumerator
        Get
            Return mCol.GetEnumerator()
        End Get
    End Property

    Public Sub New()
        ' Initializes the collection when this class is created
        mCol = New Collection()
    End Sub

End Class

