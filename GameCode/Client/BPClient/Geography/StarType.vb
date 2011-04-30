Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

Public Class StarType
    Public StarTypeID As Byte
    Public StarTypeName As String
    Private msStarTexture As String     'the actual star's texture
    Private moTexture As Texture
    Private mlStarRadius As Int32
    Public HeatIndex As Byte

    Public StarAttributes As Int32

    Public StarMaterial As Material     'set from MatDiffuse and MatEmissive
    Public StarLight As Light           'set from LightRange, LightDiffuse, LightAmbient, LightSpecular, LightAtt1(2,3)
    Public lStarMapRectIdx As Int32
    Private moStarMesh As Mesh

    Public Shared StarMapRect As Rectangle
    Public Shared StarMapTex As Texture
    Private Shared moLensSpr As Sprite
	Private Shared moLensTex As Texture
    Public Shared Sub ReleaseSprite()
        Try
            If moLensSpr Is Nothing = False Then moLensSpr.Dispose()
            moLensSpr = Nothing
            If moLensTex Is Nothing = False Then moLensTex.Dispose()
            moLensTex = Nothing
            If StarMapTex Is Nothing = False Then StarMapTex.Dispose()
            StarMapTex = Nothing
        Catch
        End Try
    End Sub
    Public Sub ClearResources()
        If moTexture Is Nothing = False Then moTexture.Dispose()
        moTexture = Nothing
        If moStarMesh Is Nothing = False Then moStarMesh.Dispose()
        moStarMesh = Nothing
    End Sub

#Region "Material and Texture Properties"
    Public Property MatDiffuse() As Int32
        Get
            Return StarMaterial.Diffuse.ToArgb
        End Get
        Set(ByVal Value As Int32)
            StarMaterial.Diffuse = System.Drawing.Color.FromArgb(Value)
        End Set
    End Property
    Public Property MatEmissive() As Int32
        Get
            Return StarMaterial.Emissive.ToArgb
        End Get
        Set(ByVal Value As Int32)
			StarMaterial.Emissive = System.Drawing.Color.FromArgb(Value)
        End Set
    End Property
    Public Property StarTexture() As String
        Get
            Return msStarTexture
        End Get
        Set(ByVal Value As String)
            msStarTexture = Value
			moTexture = goResMgr.GetTexture(msStarTexture, GFXResourceManager.eGetTextureType.NoSpecifics, "Stars.pak")
        End Set
    End Property
#End Region

#Region "Light Properties"
    Public Property LightRange() As Int32
        Get
            Return CInt(StarLight.Range)
        End Get
        Set(ByVal Value As Int32)
            StarLight.Range = Value
        End Set
    End Property
    Public Property LightDiffuse() As Int32
        Get
            Return StarLight.Diffuse.ToArgb
        End Get
        Set(ByVal Value As Int32)
            StarLight.Diffuse = System.Drawing.Color.FromArgb(Value)
        End Set
    End Property
    Public Property LightAmbient() As Int32
        Get
            Return StarLight.Ambient.ToArgb
        End Get
        Set(ByVal Value As Int32)
            StarLight.Ambient = System.Drawing.Color.FromArgb(Value)
        End Set
    End Property
    Public Property LightSpecular() As Int32
        Get
            Return StarLight.Specular.ToArgb
        End Get
        Set(ByVal Value As Int32)
            StarLight.Specular = System.Drawing.Color.FromArgb(Value)
        End Set
    End Property
    Public Property LightAtt0() As Single
        Get
            Return StarLight.Attenuation0
        End Get
        Set(ByVal Value As Single)
            StarLight.Attenuation0 = Value
        End Set
    End Property
    Public Property LightAtt1() As Single
        Get
            Return StarLight.Attenuation1
        End Get
        Set(ByVal Value As Single)
            StarLight.Attenuation1 = Value
        End Set
    End Property
    Public Property LightAtt2() As Single
        Get
            Return StarLight.Attenuation2
        End Get
        Set(ByVal Value As Single)
            StarLight.Attenuation2 = Value
        End Set
    End Property
#End Region

    Public Property StarRadius() As Int32
        Get
            Return mlStarRadius
        End Get
        Set(ByVal Value As Int32)
            mlStarRadius = Value
            moStarMesh = goResMgr.CreateTexturedSphere(mlStarRadius, 32, 32, 0)
        End Set
    End Property

    Public Sub New()
        StarLight = New Light()
        StarLight.Type = LightType.Point
        StarLight.Position = New Vector3(0, 0, 0)
        If StarMapTex Is Nothing OrElse StarMapTex.Disposed = True Then
            StarMapTex = goResMgr.GetTexture("whitestar.dds", GFXResourceManager.eGetTextureType.NoSpecifics)
        End If
    End Sub

    Public Sub Render(ByVal lEnvirType As Int32, ByVal lIndexID As Int32, ByVal lStar1Radius As Int32, ByVal lBaseHtAdjust As Int32)
        Dim moDevice As Device = GFXEngine.moDevice
        If (moStarMesh Is Nothing OrElse moStarMesh.Disposed = True) AndAlso mlStarRadius > 20 Then StarRadius = mlStarRadius
        If lEnvirType <> CurrentView.ePlanetView Then
            'Ok, render the star as a globe on the system view (normal)
            Dim dAngle As Single
            Dim matWorld As Matrix
            Dim matTemp As Matrix
            Dim matYaw As Matrix
            Dim matScale As Matrix

            Dim LocX As Int32
            Dim LocY As Int32
            Dim LocZ As Int32

            If lIndexID = 1 Then
                LocX = 0 : LocY = 0 : LocZ = 0
            ElseIf lIndexID = 2 Then
                LocX = lStar1Radius + StarRadius + 1000
                LocY = CInt((StarRadius / 2) - lStar1Radius)
                LocZ = 0
            Else
                LocX = -(lStar1Radius + StarRadius + 1000)
                LocY = -CInt(lStar1Radius / 2)
                LocZ = -LocX
            End If

            LocY -= lBaseHtAdjust

            If lEnvirType = CurrentView.eSystemMapView2 Then
                'LocX = CInt(LocX / 30)
                'LocY = CInt(LocY / 30)
                'LocZ = CInt(LocZ / 30)
                LocX \= 30
                LocY \= 30
                LocZ \= 30
            End If

            'First, render our sprite for the lens flare
            If moLensSpr Is Nothing OrElse moLensSpr.Disposed = True Then
                Device.IsUsingEventHandlers = False
                moLensSpr = New Sprite(moDevice)
                Device.IsUsingEventHandlers = True
            End If
            If moLensTex Is Nothing OrElse moLensTex.Disposed = True Then moLensTex = goResMgr.GetTexture("LensFlare.dds", GFXResourceManager.eGetTextureType.NoSpecifics)

            If moTexture Is Nothing OrElse moTexture.Disposed = True Then
                If msStarTexture Is Nothing = False AndAlso msStarTexture <> "" Then
                    StarTexture = msStarTexture
                End If
            End If

            Dim fScale As Single = CSng((StarRadius * 2) / 34.0F)
            If lEnvirType = CurrentView.eSystemMapView2 Then fScale /= 30.0F
            matTemp = Matrix.Scaling(fScale, fScale, fScale)
            matWorld = Matrix.Identity
            matWorld.Multiply(matTemp)
            matTemp = Matrix.Identity
            matTemp.Translate(LocX, LocY, LocZ)
            matWorld.Multiply(matTemp)


            Dim lPrevSrcBlend As Blend
            Dim lPrevDestBlend As Blend
            Dim lPrevOp As BlendOperation

            With moDevice
                lPrevSrcBlend = .RenderState.SourceBlend
                lPrevDestBlend = .RenderState.DestinationBlend
                lPrevOp = .RenderState.BlendOperation

                .RenderState.SourceBlend = Blend.SourceColor
                .RenderState.DestinationBlend = Blend.One

                .RenderState.BlendOperation = BlendOperation.Max
                .RenderState.AlphaBlendEnable = True
                '.RenderState.ZBufferWriteEnable = False
                .RenderState.ZBufferEnable = False
            End With

            moDevice.Transform.World = matWorld
            moLensSpr.SetWorldViewLH(matWorld, moDevice.Transform.View)
            moLensSpr.Begin(SpriteFlags.AlphaBlend Or SpriteFlags.Billboard Or SpriteFlags.ObjectSpace)

            moDevice.RenderState.SourceBlend = Blend.SourceColor
            moDevice.RenderState.DestinationBlend = Blend.One
            moDevice.RenderState.BlendOperation = BlendOperation.Add

            moLensSpr.Draw(moLensTex, System.Drawing.Rectangle.Empty, New Vector3(128, 128, 0), New Vector3(0, 0, 0), System.Drawing.Color.FromArgb(MatDiffuse)) '
            moLensSpr.End()

            moDevice.RenderState.ZBufferWriteEnable = True
            moDevice.RenderState.ZBufferEnable = True

            With moDevice
                ' .RenderState.ZBufferEnable = True
                .RenderState.SourceBlend = lPrevSrcBlend
                .RenderState.DestinationBlend = lPrevDestBlend
                .RenderState.BlendOperation = lPrevOp
            End With
            'END of SPRITE LENS FLARE RENDER

            'Rotate Yaw(Axis)
            dAngle = (900.0F / 10.0F) * (Math.PI / 180.0F)
            matYaw = Matrix.Identity
            matYaw.RotateX(CSng(dAngle))

            'Create our world matrix
            matWorld = Matrix.Identity
            If lEnvirType = CurrentView.eSystemMapView2 Then
                matTemp = Matrix.Identity
                fScale = 1 / 30
                matTemp = Matrix.Scaling(fScale, fScale, fScale)
                matWorld.Multiply(matTemp)
            End If
            matWorld.Multiply(matYaw)
            matTemp = Matrix.Identity
            matTemp.Translate(LocX, LocY, LocZ)
            matWorld.Multiply(matTemp)

            'set the world and render
            moDevice.Transform.World = matWorld
            moDevice.Material = StarMaterial
            moDevice.SetTexture(0, moTexture)
            moStarMesh.DrawSubset(0)

            fScale *= 2.0F
            matTemp = Matrix.Scaling(fScale, fScale, fScale)
            matWorld = Matrix.Identity
            matWorld.Multiply(matTemp)
            matTemp = Matrix.Identity
            matTemp.Translate(LocX, LocY, LocZ)
            matWorld.Multiply(matTemp)

            With moDevice
                lPrevSrcBlend = .RenderState.SourceBlend
                lPrevDestBlend = .RenderState.DestinationBlend
                lPrevOp = .RenderState.BlendOperation

                .RenderState.SourceBlend = Blend.SourceColor
                .RenderState.DestinationBlend = Blend.One

                .RenderState.BlendOperation = BlendOperation.Max
                .RenderState.AlphaBlendEnable = True
                '.RenderState.ZBufferWriteEnable = False
                .RenderState.ZBufferEnable = False

                .SetTextureStageState(0, TextureStageStates.AlphaArgument1, TextureArgument.TextureColor)
                .SetTextureStageState(0, TextureStageStates.AlphaArgument2, TextureArgument.Diffuse)
                .SetTextureStageState(0, TextureStageStates.AlphaOperation, TextureOperation.Modulate)
            End With

            moDevice.Transform.World = matWorld
            moLensSpr.SetWorldViewLH(matWorld, moDevice.Transform.View)
            moLensSpr.Begin(SpriteFlags.AlphaBlend Or SpriteFlags.Billboard Or SpriteFlags.ObjectSpace)

            moDevice.RenderState.SourceBlend = Blend.SourceColor
            moDevice.RenderState.DestinationBlend = Blend.One
            moDevice.RenderState.BlendOperation = BlendOperation.Add

            Dim lClrVal As Int32 = 255
            'lClrVal -= CInt(Rnd() * 3)
            Dim clr As System.Drawing.Color = System.Drawing.Color.FromArgb(255, lClrVal, lClrVal, lClrVal)

            moLensSpr.Draw(goResMgr.GetTexture("WpnFire.dds", GFXResourceManager.eGetTextureType.NoSpecifics), New Rectangle(96, 32, 31, 31), New Vector3(16, 16, 0), New Vector3(0, 0, 0), clr)
            moLensSpr.End()

            matScale = Nothing
            matWorld = Nothing
            matTemp = Nothing
            matYaw = Nothing

            moDevice.RenderState.ZBufferWriteEnable = True
            moDevice.RenderState.ZBufferEnable = True

            With moDevice
                ' .RenderState.ZBufferEnable = True
                .RenderState.SourceBlend = lPrevSrcBlend
                .RenderState.DestinationBlend = lPrevDestBlend
                .RenderState.BlendOperation = lPrevOp
                .SetTextureStageState(0, TextureStageStates.AlphaArgument1, TextureArgument.TextureColor)
                .SetTextureStageState(0, TextureStageStates.AlphaArgument2, TextureArgument.Current)
                .SetTextureStageState(0, TextureStageStates.AlphaOperation, TextureOperation.SelectArg1)
            End With
            'END of SPRITE LENS FLARE RENDER
        End If
    End Sub

End Class
