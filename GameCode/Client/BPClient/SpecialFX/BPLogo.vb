Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

Public Class BPLogo

	Private muBeyondVerts() As CustomVertex.TransformedColoredTextured
	Private muProtocolVerts() As CustomVertex.TransformedColoredTextured

	Private mlBeyondAlpha() As Int32
	Private mlProtocolAlpha() As Int32
	Private mlSpriteAlpha As Int32 = 0
	Private mbFadedIn As Boolean = False

	Private moBeyondTex As Texture = Nothing
	Private moProtocolTex As Texture = Nothing
	Private moBPTex As Texture = Nothing
    'Private moSprite As Sprite = Nothing

	Private mlIconWH As Int32 = 256

	Public Sub New(ByRef oDevice As Device, ByVal lVertUB As Int32)		'vertub should be 19
		ReDim muBeyondVerts(lVertUB)
		ReDim muProtocolVerts(lVertUB)
		ReDim mlBeyondAlpha(lVertUB)
		ReDim mlProtocolAlpha(lVertUB)

		Dim lTotalWidth As Int32 = CInt(oDevice.PresentationParameters.BackBufferWidth * 0.9F)

		mlIconWH = 128
		Dim lIconWd As Int32 = 82

		If oDevice.PresentationParameters.BackBufferHeight > 768 Then
			mlIconWH = 256
			lIconWd = 164
		End If


		'both images (beyond and protocol) are 512x64

		'we'll do the Beyond Verts first...
		Dim lOffset As Int32 = (oDevice.PresentationParameters.BackBufferWidth - lTotalWidth) \ 2
		'now, get our width of beyond...... BP icon is 164 x 256
		Dim lBeyondWidth As Int32 = oDevice.PresentationParameters.BackBufferWidth \ 2 - lOffset - (lIconWd \ 2)
		Dim lVertWidth As Int32 = lBeyondWidth \ (muBeyondVerts.GetUpperBound(0) \ 2)
		Dim lVertHeight As Int32 = CInt((lBeyondWidth / 512.0F) * 64)

		Dim lTop As Int32 = 16 + (mlIconWH \ 2) - (lVertHeight \ 2)
		Dim lBottom As Int32 = lTop + lVertHeight

		For X As Int32 = 0 To muBeyondVerts.GetUpperBound(0)
			With muBeyondVerts(X)
				.X = (X \ 2) * lVertWidth
				If X Mod 2 = 0 Then
					.Y = lTop
					.Tv = 0
				Else
					.Y = lBottom
					.Tv = 1
				End If
				.Z = 0
                .Rhw = 1
				.Tu = 0.2F + (.X / lBeyondWidth) * 0.8F
				.X += lOffset
				.Color = System.Drawing.Color.FromArgb(0, 255, 255, 255).ToArgb
			End With
			mlBeyondAlpha(X) = 0
		Next X

		'Now, for the Protocol Verts...
		Dim lTempVal As Int32 = lOffset
		lOffset = oDevice.PresentationParameters.BackBufferWidth \ 2
		lOffset += (lIconWd \ 2)
		Dim lProtocolWidth As Int32 = oDevice.PresentationParameters.BackBufferWidth - lOffset - lTempVal
		lVertWidth = lProtocolWidth \ (muProtocolVerts.GetUpperBound(0) \ 2)
		lVertHeight = CInt((lProtocolWidth / 512.0F) * 64)
		lTop = 16 + (mlIconWH \ 2) - (lVertHeight \ 2)
		lBottom = lTop + lVertHeight
		For X As Int32 = 0 To muProtocolVerts.GetUpperBound(0)
			With muProtocolVerts(X)
				.X = ((X \ 2) * lVertWidth)
				If X Mod 2 = 0 Then
					.Y = lTop
					.Tv = 0
				Else
					.Y = lBottom
					.Tv = 1
				End If
				.Z = 0
                .Rhw = 1
				.Tu = .X / lProtocolWidth
				.X += lOffset
				.Color = System.Drawing.Color.FromArgb(0, 255, 255, 255).ToArgb
			End With
			mlProtocolAlpha(X) = 0
		Next X
	End Sub

	Public Function Render(ByVal bFadeout As Boolean, ByRef oDevice As Device) As Boolean
		If bFadeout = True Then
			Dim bDone As Boolean = True
			For X As Int32 = 0 To muBeyondVerts.GetUpperBound(0)
				Dim lValue As Int32 = mlBeyondAlpha(X)
				lValue -= 10
				If lValue < 0 Then lValue = 0
				If lValue <> 0 Then bDone = False
				mlBeyondAlpha(X) = lValue
				muBeyondVerts(X).Color = System.Drawing.Color.FromArgb(lValue, 255, 255, 255).ToArgb
				lValue = mlProtocolAlpha(X)
				lValue -= 10
				If lValue < 0 Then lValue = 0
				If lValue <> 0 Then bDone = False
				mlProtocolAlpha(X) = lValue
				muProtocolVerts(X).Color = System.Drawing.Color.FromArgb(lValue, 255, 255, 255).ToArgb
			Next X
			mlSpriteAlpha -= 15
			If mlSpriteAlpha < 0 Then mlSpriteAlpha = 0
			If bDone = True Then Return True
		ElseIf mbFadedIn = False Then
			Dim bDone As Boolean = True
			For X As Int32 = 0 To muBeyondVerts.GetUpperBound(0)
				Dim lValue As Int32 = mlBeyondAlpha(X)
				lValue += 1
				If lValue > 255 Then lValue = 255
				If lValue <> 255 Then bDone = False
				mlBeyondAlpha(X) = lValue
				muBeyondVerts(X).Color = System.Drawing.Color.FromArgb(lValue, 255, 255, 255).ToArgb
				lValue = mlProtocolAlpha(X)
				lValue += 1
				If lValue > 255 Then lValue = 255
				If lValue <> 255 Then bDone = False
				mlProtocolAlpha(X) = lValue
				muProtocolVerts(X).Color = System.Drawing.Color.FromArgb(lValue, 255, 255, 255).ToArgb
			Next
			mlSpriteAlpha += 5
			If mlSpriteAlpha > 255 Then mlSpriteAlpha = 255
			If bDone = True Then mbFadedIn = True
		Else
			For X As Int32 = 0 To muBeyondVerts.GetUpperBound(0)
				Dim lValue As Int32 = mlBeyondAlpha(X)
                lValue += CInt(Rnd() * 32) - 16
                If lValue < 32 Then lValue = 32
                If lValue > 255 Then lValue = 255
                lValue = 255
                mlBeyondAlpha(X) = lValue
                muBeyondVerts(X).Color = System.Drawing.Color.FromArgb(lValue, 255, 255, 255).ToArgb

                lValue = mlProtocolAlpha(X)
                lValue += CInt(Rnd() * 32) - 16
				If lValue < 32 Then lValue = 32
				If lValue > 255 Then lValue = 255
				lValue = 255
				mlProtocolAlpha(X) = lValue
                muProtocolVerts(X).Color = System.Drawing.Color.FromArgb(lValue, 255, 255, 255).ToArgb
            Next X
		End If

		If moBeyondTex Is Nothing OrElse moBeyondTex.Disposed = True Then
			Device.IsUsingEventHandlers = False
			moBeyondTex = goResMgr.LoadScratchTexture("Beyond.dds", "Misc.pak")
			Device.IsUsingEventHandlers = True
		End If
		If moProtocolTex Is Nothing OrElse moProtocolTex.Disposed = True Then
			Device.IsUsingEventHandlers = False
			moProtocolTex = goResMgr.LoadScratchTexture("Protocol.dds", "Misc.pak")
			Device.IsUsingEventHandlers = True
		End If
		If moBPTex Is Nothing OrElse moBPTex.Disposed = True Then
			Device.IsUsingEventHandlers = False
			moBPTex = goResMgr.LoadScratchTexture("BP.dds", "Misc.pak")
			Device.IsUsingEventHandlers = True
		End If

		With oDevice
			.Transform.World = Matrix.Identity

			.SetTextureStageState(0, TextureStageStates.AlphaArgument1, TextureArgument.TextureColor)
			.SetTextureStageState(0, TextureStageStates.AlphaArgument2, TextureArgument.Diffuse)
			.SetTextureStageState(0, TextureStageStates.AlphaOperation, TextureOperation.Modulate)

			.RenderState.SourceBlend = Blend.SourceAlpha
			.RenderState.DestinationBlend = Blend.One
			.RenderState.AlphaBlendEnable = True
			.RenderState.ZBufferWriteEnable = False
			'.RenderState.Lighting = False

			.VertexFormat = CustomVertex.TransformedColoredTextured.Format
			.SetTexture(0, moBeyondTex)
			.DrawUserPrimitives(PrimitiveType.TriangleStrip, muBeyondVerts.GetUpperBound(0) - 1, muBeyondVerts)
			.SetTexture(0, moProtocolTex)
			.DrawUserPrimitives(PrimitiveType.TriangleStrip, muProtocolVerts.GetUpperBound(0) - 1, muProtocolVerts)

			'Then, reset our device...
			.RenderState.ZBufferWriteEnable = True
			'.RenderState.Lighting = True
			'.RenderState.PointSpriteEnable = False
			.RenderState.SourceBlend = Blend.SourceAlpha 'Blend.One
			.RenderState.DestinationBlend = Blend.InvSourceAlpha 'Blend.Zero
			.RenderState.AlphaBlendEnable = True

			.SetTextureStageState(0, TextureStageStates.AlphaArgument1, TextureArgument.TextureColor)
			.SetTextureStageState(0, TextureStageStates.AlphaArgument2, TextureArgument.Current)
			.SetTextureStageState(0, TextureStageStates.AlphaOperation, TextureOperation.SelectArg1)
        End With

        'If moSprite Is Nothing OrElse moSprite.Disposed = True Then
        '    Device.IsUsingEventHandlers = False
        '    moSprite = New Sprite(oDevice)
        '    Device.IsUsingEventHandlers = True
        'End If

        If moBPTex Is Nothing Then Return False

        Dim rcDest As Rectangle
        rcDest.X = oDevice.PresentationParameters.BackBufferWidth \ 2 - (mlIconWH \ 2)
        rcDest.Y = 16
        rcDest.Width = mlIconWH
        rcDest.Height = mlIconWH
        'moSprite.Begin(SpriteFlags.AlphaBlend)
        Dim ptTemp As Point = rcDest.Location
        If mlIconWH <> 128 Then
            Dim fTemp As Single = 128.0F / mlIconWH
            ptTemp.X = CInt(ptTemp.X * fTemp)
            ptTemp.Y = CInt(ptTemp.Y * fTemp)
        End If
        'moSprite.Draw2D(moBPTex, Rectangle.Empty, rcDest, Point.Empty, 0, ptTemp, System.Drawing.Color.FromArgb(mlSpriteAlpha, 255, 255, 255))
        'moSprite.End()
        Dim bFog As Boolean = False
        With oDevice.RenderState
            .Lighting = False
            bFog = .FogEnable
            .FogEnable = False
        End With
        Dim lImgWidth As Int32 = 128
        Dim lImgHeight As Int32 = 128
        With moBPTex.GetLevelDescription(0)
            lImgWidth = .Width
            lImgHeight = .Height
        End With
        BPSprite.Draw2DOnce(oDevice, moBPTex, New Rectangle(0, 0, lImgWidth, lImgHeight), rcDest, System.Drawing.Color.FromArgb(mlSpriteAlpha, 255, 255, 255), lImgWidth, lImgHeight)
        With oDevice.RenderState
            .Lighting = True
            .FogEnable = bFog
        End With


        Return False
    End Function

    Public Sub ReleaseSprite()
        Try
            'If moSprite Is Nothing = False Then moSprite.Dispose()
            'moSprite = Nothing
            If moBeyondTex Is Nothing = False Then moBeyondTex.Dispose()
            moBeyondTex = Nothing
            If moProtocolTex Is Nothing = False Then moProtocolTex.Dispose()
            moProtocolTex = Nothing
            If moBPTex Is Nothing = False Then moBPTex.Dispose()
            moBPTex = Nothing
        Catch
        End Try
    End Sub

	Protected Overrides Sub Finalize()
		If moBeyondTex Is Nothing = False Then moBeyondTex.Dispose()
		moBeyondTex = Nothing
		If moProtocolTex Is Nothing = False Then moProtocolTex.Dispose()
		moProtocolTex = Nothing
		If moBPTex Is Nothing = False Then moBPTex.Dispose()
		moBPTex = Nothing
        'If moSprite Is Nothing = False Then moSprite.Dispose()
        'moSprite = Nothing
		MyBase.Finalize()
	End Sub
End Class
