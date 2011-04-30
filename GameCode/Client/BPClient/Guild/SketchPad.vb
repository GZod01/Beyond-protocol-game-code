Option Strict On

Public Class SketchPad
	Public Enum eySketchShapes As Byte
		Box = 0
		Circle = 1
		Line = 2
		Arrow = 3
		Cross = 4
		Text = 5
	End Enum

	Public Structure SketchPadItem
		Public yType As eySketchShapes
		Public fPtA_X As Single
		Public fPtA_Y As Single
		Public fPtB_X As Single
		Public fPtB_Y As Single
		Public yClrVal As Byte

		Public sText As String		'for textboxes
	End Structure

	Public lID As Int32 = -1
	Public sName As String
	Public lEnvirID As Int32
	Public iEnvirTypeID As Int16
	Public ViewID As Int32
	Public CameraX As Int32
	Public CameraY As Int32
	Public CameraZ As Int32
	Public CameraAtX As Int32
	Public CameraAtY As Int32
	Public CameraAtZ As Int32

	Public uItems() As SketchPadItem

	Public Sub AddItem(ByVal yType As eySketchShapes, ByVal fAX As Single, ByVal fAY As Single, ByVal fBX As Single, ByVal fBY As Single, ByVal yClrVal As Byte, ByVal sText As String)
		Dim uTmp As SketchPadItem
		With uTmp
			.fPtA_X = fAX
			.fPtA_Y = fAY
			.fPtB_X = fBX
			.fPtB_Y = fBY
			.sText = sText
			.yClrVal = yClrVal
			.yType = yType
		End With
		If uItems Is Nothing Then ReDim uItems(-1)
		ReDim Preserve uItems(uItems.GetUpperBound(0) + 1)
		uItems(uItems.GetUpperBound(0)) = uTmp
	End Sub

	Public Shared Function GetColorFromValue(ByVal lColorVal As Int32) As System.Drawing.Color
		Select Case lColorVal
			Case 0
				Return System.Drawing.Color.FromArgb(255, 255, 255, 255)
			Case 1
				Return System.Drawing.Color.FromArgb(255, 255, 0, 0)
			Case 2
				Return System.Drawing.Color.FromArgb(255, 0, 255, 0)
			Case 3
				Return System.Drawing.Color.FromArgb(255, 0, 0, 255)
			Case 4
				Return System.Drawing.Color.FromArgb(255, 255, 255, 0)
			Case 5
				Return System.Drawing.Color.FromArgb(255, 0, 255, 255)
			Case 6
				Return System.Drawing.Color.FromArgb(255, 255, 0, 255)
			Case 7
				Return System.Drawing.Color.FromArgb(255, 255, 128, 64)
			Case 8
				Return System.Drawing.Color.FromArgb(255, 128, 0, 128)
			Case 9
				Return System.Drawing.Color.FromArgb(255, 0, 0, 0)
			Case 10
				Return System.Drawing.Color.FromArgb(255, 128, 128, 128)
			Case 11
				Return System.Drawing.Color.FromArgb(255, 0, 128, 128)
			Case 12
				Return System.Drawing.Color.FromArgb(255, 64, 64, 64)
			Case 13
				Return System.Drawing.Color.FromArgb(255, 128, 0, 0)
			Case 14
				Return System.Drawing.Color.FromArgb(255, 0, 128, 0)
			Case 15
				Return System.Drawing.Color.FromArgb(255, 0, 0, 128)
		End Select
	End Function

	Public Function GetAsAddObjectMsg(ByVal lEventID As Int32) As Byte()
		Dim lItemCnt As Int32 = 0
		Dim lItemLen As Int32 = 0
		If uItems Is Nothing = False Then
			lItemCnt = uItems.Length

			For X As Int32 = 0 To uItems.GetUpperBound(0)
				lItemLen += 18
				If uItems(X).yType = eySketchShapes.Text Then
					lItemLen += 4
					If uItems(X).sText Is Nothing Then uItems(X).sText = ""
					lItemLen += uItems(X).sText.Length
				End If
			Next X
		End If

		Dim yMsg(68 + lItemLen) As Byte
		Dim lPos As Int32 = 0

		System.BitConverter.GetBytes(GlobalMessageCode.eAddEventAttachment).CopyTo(yMsg, lPos) : lPos += 2
		System.BitConverter.GetBytes(lID).CopyTo(yMsg, lPos) : lPos += 4
		System.BitConverter.GetBytes(lEventID).CopyTo(yMsg, lPos) : lPos += 4
		yMsg(lPos) = 0 : lPos += 1
		System.Text.ASCIIEncoding.ASCII.GetBytes(sName).CopyTo(yMsg, lPos) : lPos += 20
		System.BitConverter.GetBytes(lEnvirID).CopyTo(yMsg, lPos) : lPos += 4
		System.BitConverter.GetBytes(iEnvirTypeID).CopyTo(yMsg, lPos) : lPos += 2
		System.BitConverter.GetBytes(ViewID).CopyTo(yMsg, lPos) : lPos += 4
		System.BitConverter.GetBytes(CameraX).CopyTo(yMsg, lPos) : lPos += 4
		System.BitConverter.GetBytes(CameraY).CopyTo(yMsg, lPos) : lPos += 4
		System.BitConverter.GetBytes(CameraZ).CopyTo(yMsg, lPos) : lPos += 4
		System.BitConverter.GetBytes(CameraAtX).CopyTo(yMsg, lPos) : lPos += 4
		System.BitConverter.GetBytes(CameraAtY).CopyTo(yMsg, lPos) : lPos += 4
		System.BitConverter.GetBytes(CameraAtZ).CopyTo(yMsg, lPos) : lPos += 4


		System.BitConverter.GetBytes(lItemCnt).CopyTo(yMsg, lPos) : lPos += 4
		If uItems Is Nothing = False Then
			For X As Int32 = 0 To uItems.GetUpperBound(0)
				With uItems(X)
					yMsg(lPos) = .yType : lPos += 1
					yMsg(lPos) = .yClrVal : lPos += 1
					System.BitConverter.GetBytes(.fPtA_X).CopyTo(yMsg, lPos) : lPos += 4
					System.BitConverter.GetBytes(.fPtA_Y).CopyTo(yMsg, lPos) : lPos += 4
					System.BitConverter.GetBytes(.fPtB_X).CopyTo(yMsg, lPos) : lPos += 4
					System.BitConverter.GetBytes(.fPtB_Y).CopyTo(yMsg, lPos) : lPos += 4

					If .yType = eySketchShapes.Text Then
						If .sText Is Nothing Then .sText = ""
						System.BitConverter.GetBytes(.sText.Length).CopyTo(yMsg, lPos) : lPos += 4
						System.Text.ASCIIEncoding.ASCII.GetBytes(.sText).CopyTo(yMsg, lPos) : lPos += .sText.Length
					End If
				End With
			Next X
		End If


		Return yMsg
	End Function
End Class
