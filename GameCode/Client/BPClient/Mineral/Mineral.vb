
Public Class Mineral
    Inherits Base_GUID

    Public bDiscovered As Boolean

    Public MineralName As String
    Public bReceivedProps As Boolean = False
    Public bRequestedPropsExport As Boolean = False ' Unlike bRequestedProps this needs to stay set to True once requested.
	Public bRequestedProps As Boolean = False
	Public lLastMsgUpdate As Int32

	Public bArchived As Boolean = False

    Private Structure MineralPropVal
        Public lMineralPropertyID As Int32
        Public sValue As String
        Public lValueScore As Int32
        Public sKnowValue As String
    End Structure

    Private muMinPropVals() As MineralPropVal
    Private mlMinPropValUB As Int32 = -1
    Private mbLoaded As Boolean = False

    Public Sub SetMineralPropertyValue(ByVal lMineralPropertyID As Int32, ByVal sValue As String, ByVal sKnowValue As String, ByVal yValueScore As Byte)
        mbLoaded = True
        For X As Int32 = 0 To mlMinPropValUB
            If muMinPropVals(X).lMineralPropertyID = lMineralPropertyID Then
                muMinPropVals(X).sValue = sValue
                muMinPropVals(X).sKnowValue = sKnowValue
                muMinPropVals(X).lValueScore = CInt(yValueScore)
                Return
            End If
        Next X

        mlMinPropValUB += 1
        ReDim Preserve muMinPropVals(mlMinPropValUB)
        muMinPropVals(mlMinPropValUB).lMineralPropertyID = lMineralPropertyID
        muMinPropVals(mlMinPropValUB).sValue = sValue
        muMinPropVals(mlMinPropValUB).sKnowValue = sKnowValue
        muMinPropVals(mlMinPropValUB).lValueScore = CInt(yValueScore)
    End Sub

    Public Function GetPropertyIndex(ByVal lPropID As Int32) As Int32
        For X As Int32 = 0 To mlMinPropValUB
            If muMinPropVals(X).lMineralPropertyID = lPropID Then
                Return X
            End If
        Next X
        Return -1
    End Function

	'Public Function MinPropValue(ByVal lIndex As Int32) As String
	'    Return muMinPropVals(lIndex).sValue
	'End Function

	'Public Function MinPropKnownValue(ByVal lIndex As Int32) As String
	'    Return muMinPropVals(lIndex).sKnowValue
	'End Function

    Public Function HasNonZeroProperty() As Boolean
        If muMinPropVals Is Nothing Then Return False
        For X As Int32 = 0 To mlMinPropValUB
            If muMinPropVals(X).lValueScore > 0 Then Return True
        Next X
        'Fallback check in case of a 'zero' for all values mineral.
        If mbLoaded = True Then Return True
        Return False
    End Function

	Public Function MinPropValueScore(ByVal lIndex As Int32) As Int32
		If muMinPropVals Is Nothing Then Return 0
		If lIndex > muMinPropVals.GetUpperBound(0) Then Return 0
		If lIndex < 0 Then Return 0

		Return muMinPropVals(lIndex).lValueScore
	End Function

	Public Function MinPropKnownScore(ByVal lIndex As Int32) As Int32
		If muMinPropVals Is Nothing Then Return 0
		If lIndex > muMinPropVals.GetUpperBound(0) Then Return 0
		If lIndex < 0 Then Return 0
		If muMinPropVals(lIndex).sKnowValue Is Nothing Then Return 0
		If muMinPropVals(lIndex).sKnowValue.ToLower = "a full" Then Return 9

		Select Case muMinPropVals(lIndex).sKnowValue.ToLower
			Case "no"
				Return 0
			Case "a sliver of"
				Return 1
			Case "very little"
				Return 2
			Case "a small"
				Return 3
			Case "adequate"
				Return 4
			Case "ample"
				Return 5
			Case "a large"
				Return 6
			Case "an expansive"
				Return 7
			Case "a vast", "an unlimited"
				Return 8
				'Case "an unlimited"
				'    Return 9
		End Select
		Return 0
	End Function

	Public Function SentenceDescription(ByVal lIndex As Int32) As String
		If muMinPropVals Is Nothing Then Return ""
		If lIndex > muMinPropVals.GetUpperBound(0) Then Return ""
		If lIndex < 0 Then Return ""

		Dim sResult As String
		If lIndex = -1 Then Return ""
		With muMinPropVals(lIndex)
			If .sKnowValue = "" OrElse .sValue = "" Then Return ""

			sResult = "We have "

			sResult &= .sKnowValue
			sResult &= " understanding of this material's "

			For X As Int32 = 0 To glMineralPropertyUB
				If glMineralPropertyIdx(X) = .lMineralPropertyID Then
					sResult &= goMineralProperty(X).MineralPropertyName
					Exit For
				End If
			Next X

			sResult &= " and it has "
			sResult &= .sValue & " qualities."
		End With

		Return sResult
	End Function

	'Public Function IsFullyStudied() As Boolean
	'	Try
	'		For X As Int32 = 0 To mlMinPropValUB
	'			If MinPropKnownScore(X) + 1 <> 10 Then Return False
	'		Next X
	'	Catch
	'	End Try
	'	Return True
    'End Function

    Public Sub CheckRequestProps()
        If bRequestedPropsExport = True Then Return
        'If bRequestedProps = True Then Return
        If goUILib Is Nothing Then Return
        bRequestedPropsExport = True

        Me.lLastMsgUpdate = 0
        Me.bRequestedProps = True

        Dim lIdx As Int32
        For X As Int32 = 0 To glMineralUB
            If goMinerals(X).ObjectID = Me.ObjectID Then
                lIdx = X
                Exit For
            End If
        Next
        If lIdx = -1 Then Return
        Dim yMsg(5) As Byte
        System.BitConverter.GetBytes(GlobalMessageCode.eRequestMineral).CopyTo(yMsg, 0)
        System.BitConverter.GetBytes(glMineralIdx(0)).CopyTo(yMsg, 2)

        goUILib.SendMsgToPrimary(yMsg)


    End Sub
End Class
