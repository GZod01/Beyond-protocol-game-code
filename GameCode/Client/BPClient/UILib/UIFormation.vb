Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

Public Class UIFormation
	Inherits UIControl

	Public Const ml_GRID_SIZE_WH As Int32 = 25		'number of squares in each direction - an odd number to provide a "middle"
	Private Const ml_SQUARE_SIZE As Int32 = 16		'Width and Height of each square

	Public AutoRefresh As Boolean = True

	Private mlSlots(,) As Int32

	Public Event SlotClick(ByVal lIndexX As Int32, ByVal lIndexY As Int32)
 
	Public Sub New(ByRef oUILib As UILib)
		MyBase.New(oUILib)

		ReDim mlSlots(ml_GRID_SIZE_WH - 1, ml_GRID_SIZE_WH - 1)

		For Y As Int32 = 0 To ml_GRID_SIZE_WH - 1
			For X As Int32 = 0 To ml_GRID_SIZE_WH - 1
				mlSlots(X, Y) = 0
			Next X
		Next Y

		Me.Width = (ml_SQUARE_SIZE * ml_GRID_SIZE_WH) + 2
		Me.Height = Me.Width
	End Sub

	'Private Sub UIFormation_OnMouseMove(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles Me.OnMouseMove
	'	If lButton = MouseButtons.Left Then
	'		Dim oLoc As System.Drawing.Point = Me.GetAbsolutePosition()
	'		Dim lTmpX As Int32 = lMouseX - oLoc.X - 1
	'		Dim lTmpY As Int32 = lMouseY - oLoc.Y - 1

	'		lTmpX \= ml_SQUARE_SIZE
	'		lTmpY \= ml_SQUARE_SIZE

	'		If lTmpX > -1 AndAlso lTmpY > -1 AndAlso lTmpX < ml_GRID_SIZE_WH AndAlso lTmpY < ml_GRID_SIZE_WH Then
	'			RaiseEvent SlotClick(lTmpX, lTmpY)
	'		End If
	'	End If
	'End Sub

	Private Sub UIFormation_OnMouseUp(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles Me.OnMouseUp
		Dim oLoc As System.Drawing.Point = Me.GetAbsolutePosition()
		Dim lTmpX As Int32 = lMouseX - oLoc.X - 1
		Dim lTmpY As Int32 = lMouseY - oLoc.Y - 1

		lTmpX \= ml_SQUARE_SIZE
		lTmpY \= ml_SQUARE_SIZE

		If lTmpX > -1 AndAlso lTmpY > -1 AndAlso lTmpX < ml_GRID_SIZE_WH AndAlso lTmpY < ml_GRID_SIZE_WH Then
			RaiseEvent SlotClick(lTmpX, lTmpY)
		End If
	End Sub

	Private Sub UIFormation_OnRender() Handles Me.OnRender
		Dim oLoc As System.Drawing.Point = GetAbsolutePosition()

		Dim oBorderLineVerts(4) As Vector2
		'Draw a box border around our window...
		With oBorderLineVerts(0)
			.X = oLoc.X
			.Y = oLoc.Y
		End With
		With oBorderLineVerts(1)
			.X = oLoc.X + Width
			.Y = oLoc.Y
		End With
		With oBorderLineVerts(2)
			.X = oLoc.X + Width
			.Y = oLoc.Y + Height
		End With
		With oBorderLineVerts(3)
			.X = oLoc.X
			.Y = oLoc.Y + Height
		End With
		With oBorderLineVerts(4)
			.X = oLoc.X
			.Y = oLoc.Y
		End With

		'Draw a box border around our window...
        Using moBorderLine As New Line(MyBase.moUILib.oDevice)
            moBorderLine.Antialias = True
            moBorderLine.Width = 2
            moBorderLine.Begin()
            moBorderLine.Draw(oBorderLineVerts, muSettings.InterfaceBorderColor)
            moBorderLine.End()
        End Using
		'End of the border drawing

		'Now, render our stuff
		Dim lTmpX As Int32
		Dim lTmpY As Int32

        If GFXEngine.gbPaused = False AndAlso GFXEngine.gbDeviceLost = False AndAlso MyBase.moUILib.oInterfaceTexture Is Nothing = False Then
            Using oSpr As Sprite = New Sprite(MyBase.moUILib.oDevice)

                Dim rcSrc As Rectangle = grc_UI(elInterfaceRectangle.eCheck_Unchecked)
                Dim rcDest As Rectangle
                Dim fX As Single
                Dim fY As Single

                oSpr.Begin(SpriteFlags.AlphaBlend)

                Dim lMidPoints As Int32 = ml_GRID_SIZE_WH \ 2

                For Y As Int32 = 0 To ml_GRID_SIZE_WH - 1
                    For X As Int32 = 0 To ml_GRID_SIZE_WH - 1
                        lTmpX = oLoc.X + (X * ml_SQUARE_SIZE) + 1
                        lTmpY = oLoc.Y + (Y * ml_SQUARE_SIZE) + 1

                        fX = lTmpX * (rcSrc.Width / CSng(16))
                        fY = lTmpY * (rcSrc.Height / CSng(16))

                        rcDest = Rectangle.FromLTRB(oLoc.X, oLoc.Y, oLoc.X + 16, oLoc.Y + 16)

                        If X = lMidPoints OrElse Y = lMidPoints Then
                            oSpr.Draw2D(MyBase.moUILib.oInterfaceTexture, rcSrc, rcDest, System.Drawing.Point.Empty, 0, New Point(CInt(fX), CInt(fY)), System.Drawing.Color.FromArgb(255, 32, 255, 32))
                        Else : oSpr.Draw2D(MyBase.moUILib.oInterfaceTexture, rcSrc, rcDest, System.Drawing.Point.Empty, 0, New Point(CInt(fX), CInt(fY)), System.Drawing.Color.FromArgb(255, 32, 128, 32))
                        End If

                    Next X
                Next Y
                oSpr.End()
                oSpr.Dispose()
            End Using

            Dim lMaxSlot As Int32 = Me.GetMaxSlotValue()

            Using oFont As Font = New Font(MyBase.moUILib.oDevice, New System.Drawing.Font("Arial", 10, FontStyle.Bold, GraphicsUnit.Point, 0))
                Dim sCap As String
                Dim rcDest As Rectangle
                Using oTextSpr As Sprite = New Sprite(MyBase.moUILib.oDevice)
                    oTextSpr.Begin(SpriteFlags.AlphaBlend)
                    For Y As Int32 = 0 To ml_GRID_SIZE_WH - 1
                        For X As Int32 = 0 To ml_GRID_SIZE_WH - 1
                            lTmpX = oLoc.X + 1 + (X * ml_SQUARE_SIZE)
                            lTmpY = oLoc.Y + 1 + (Y * ml_SQUARE_SIZE)

                            rcDest = Rectangle.FromLTRB(lTmpX, lTmpY, lTmpX + 16, lTmpY + 16)

                            If mlSlots(X, Y) = lMaxSlot Then
                                sCap = "D"
                            Else : sCap = mlSlots(X, Y).ToString
                            End If
                            If mlSlots(X, Y) > 0 AndAlso sCap <> "" Then
                                oFont.DrawText(oTextSpr, sCap, rcDest, DrawTextFormat.Center Or DrawTextFormat.VerticalCenter, Color.White)
                            End If
                        Next X
                    Next Y
                    oTextSpr.End()
                    oTextSpr.Dispose()
                End Using

                oFont.Dispose()
            End Using
        End If
    End Sub

	Protected Overrides Sub Finalize()
		If goResMgr Is Nothing = False Then goResMgr.DeleteTexture("HullBuilderIcons.dds")
		MyBase.Finalize()
	End Sub

	Public Function HasChanged(ByRef oFormationDef As FormationDef) As Boolean
		For Y As Int32 = 0 To ml_GRID_SIZE_WH - 1
			For X As Int32 = 0 To ml_GRID_SIZE_WH - 1
				If oFormationDef.mlSlots(X, Y) <> mlSlots(X, Y) Then Return True
			Next X
		Next Y
		Return False
	End Function

	Public Sub SetSlot(ByVal lX As Int32, ByVal lY As Int32, ByVal lValue As Int32)
		If lX > -1 AndAlso lY > -1 AndAlso lX < ml_GRID_SIZE_WH AndAlso lY < ml_GRID_SIZE_WH Then
			mlSlots(lX, lY) = lValue
			MyBase.IsDirty = True
		End If
	End Sub

	Public Sub InsertSlot(ByVal lX As Int32, ByVal lY As Int32, ByVal lValue As Int32)
		If lX > -1 AndAlso lY > -1 AndAlso lX < ml_GRID_SIZE_WH AndAlso lY < ml_GRID_SIZE_WH Then
			mlSlots(lX, lY) = lValue

			'Now, go and increment all slots above me...
			For X As Int32 = 0 To ml_GRID_SIZE_WH - 1
				For Y As Int32 = 0 To ml_GRID_SIZE_WH - 1
					If mlSlots(X, Y) >= lValue AndAlso (X <> lX OrElse Y <> lY) Then mlSlots(X, Y) += 1
				Next Y
			Next X

			MyBase.IsDirty = True
		End If
	End Sub

	Public Sub RemoveSlot(ByVal lValue As Int32)
		For X As Int32 = 0 To ml_GRID_SIZE_WH - 1
			For Y As Int32 = 0 To ml_GRID_SIZE_WH - 1
				If mlSlots(X, Y) = lValue Then mlSlots(X, Y) = 0
			Next Y
		Next X
		Me.IsDirty = True
	End Sub

	Public Sub RemoveAndShiftSlot(ByVal lIndexX As Int32, ByVal lIndexY As Int32)
		Dim lSlotVal As Int32 = GetSlot(lIndexX, lIndexY)
		If lSlotVal < 1 Then Return

		mlSlots(lIndexX, lIndexY) = 0

		For X As Int32 = 0 To ml_GRID_SIZE_WH - 1
			For Y As Int32 = 0 To ml_GRID_SIZE_WH - 1
				If mlSlots(X, Y) > lSlotVal Then mlSlots(X, Y) -= 1
			Next Y
		Next X
		Me.IsDirty = True
	End Sub

	Public Function GetSlot(ByVal lX As Int32, ByVal lY As Int32) As Int32
		If lX > -1 AndAlso lY > -1 AndAlso lX < ml_GRID_SIZE_WH AndAlso lY < ml_GRID_SIZE_WH Then
			Return mlSlots(lX, lY)
		End If
		Return -1
	End Function

	Public Function GetMaxSlotValue() As Int32
		Dim lMax As Int32 = 0
		For X As Int32 = 0 To ml_GRID_SIZE_WH - 1
			For Y As Int32 = 0 To ml_GRID_SIZE_WH - 1
				If mlSlots(X, Y) > lMax Then lMax = mlSlots(X, Y)
			Next
		Next
		Return lMax
	End Function

	Public Function GetSlotCount() As Int32
		Dim lCnt As Int32 = 0
		For X As Int32 = 0 To ml_GRID_SIZE_WH - 1
			For Y As Int32 = 0 To ml_GRID_SIZE_WH - 1
				If mlSlots(X, Y) > 0 Then lCnt += 1
			Next
		Next
		Return lCnt
	End Function

	Public Sub Clear()
		For X As Int32 = 0 To ml_GRID_SIZE_WH - 1
			For Y As Int32 = 0 To ml_GRID_SIZE_WH - 1
				mlSlots(X, Y) = 0
			Next
		Next
	End Sub

	Public Function Validate() As Int32
		Dim lMax As Int32 = GetMaxSlotValue()

		'Ok, now, go through them
		For lValue As Int32 = 1 To lMax
			Dim lCnt As Int32 = 0

			For X As Int32 = 0 To ml_GRID_SIZE_WH - 1
				For Y As Int32 = 0 To ml_GRID_SIZE_WH - 1
					If mlSlots(X, Y) = lValue Then lCnt += 1
				Next Y
			Next X

			If lCnt = 0 Then
				'ok, shift all values down
				For X As Int32 = 0 To ml_GRID_SIZE_WH - 1
					For Y As Int32 = 0 To ml_GRID_SIZE_WH - 1
						If mlSlots(X, Y) > lValue Then mlSlots(X, Y) -= 1
					Next Y
				Next X
			ElseIf lCnt > 1 Then
				Return lValue
			End If
		Next lValue
		Return 0
	End Function

	Public Function GetMsgData() As Byte()
		Dim lCnt As Int32 = 0
		For X As Int32 = 0 To ml_GRID_SIZE_WH - 1
			For Y As Int32 = 0 To ml_GRID_SIZE_WH - 1
				If mlSlots(X, Y) > 0 Then lCnt += 1
			Next Y
		Next X

		Dim yResp((lCnt * 4) + 3) As Byte
		Dim lPos As Int32 = 0

		System.BitConverter.GetBytes(lCnt).CopyTo(yResp, lPos) : lPos += 4

		For X As Int32 = 0 To ml_GRID_SIZE_WH - 1
			For Y As Int32 = 0 To ml_GRID_SIZE_WH - 1
				If mlSlots(X, Y) > 0 Then
					Dim iSlotIdx As Int16 = CShort((Y * ml_GRID_SIZE_WH) + X)
					Dim iSlotValue As Int16 = CShort(mlSlots(X, Y))

					System.BitConverter.GetBytes(iSlotIdx).CopyTo(yResp, lPos) : lPos += 2
					System.BitConverter.GetBytes(iSlotValue).CopyTo(yResp, lPos) : lPos += 2
				End If
			Next Y
		Next X

		Return yResp
	End Function

	Public Sub SetFromFormation(ByRef oFormation As FormationDef)
		Me.Clear()

		If oFormation.mlSlots Is Nothing = False Then
			For X As Int32 = 0 To ml_GRID_SIZE_WH - 1
				For Y As Int32 = 0 To ml_GRID_SIZE_WH - 1
					Me.mlSlots(X, Y) = oFormation.mlSlots(X, Y)
					'If Me.mlSlots(X, Y) > 0 Then Stop
				Next Y
			Next X
		End If
		Me.IsDirty = True
	End Sub
End Class