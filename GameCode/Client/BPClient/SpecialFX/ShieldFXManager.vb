Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D


Public Class ShieldFXManager
    Private Declare Function timeGetTime Lib "winmm.dll" Alias "timeGetTime" () As Int32

    Private Class ShldEffect
        Public Index As Int32

        Public oParent As RenderObject
        Public ShieldAlpha As Byte
        Public ShieldRed As Byte
        Public ShieldBlue As Byte
        Public ShieldGreen As Byte

        Public Event ShieldEnd(ByVal lIndex As Int32)

        Public Sub Render(ByVal bUpdateNoRender As Boolean)
            Dim oMat As Material
            Dim moDevice As Device = GFXEngine.moDevice

            'Update our colors
            If ShieldRed > 14 Then ShieldRed -= CByte(15) Else ShieldRed = 0
            If ShieldGreen > 14 Then ShieldGreen -= CByte(15) Else ShieldGreen = 0
            If ShieldBlue > 14 Then ShieldBlue -= CByte(15) Else ShieldBlue = 0
            If ShieldAlpha > 14 Then ShieldAlpha -= CByte(15) Else ShieldAlpha = 0

            If bUpdateNoRender = False AndAlso oParent.bCulled = False AndAlso oParent.yVisibility = eVisibilityType.Visible Then
                If oParent.oMesh Is Nothing = False Then
                    'If Not (oParent.oMesh Is Nothing OrElse oParent.oMesh.oShieldMesh Is Nothing) Then
                    'Set up our matrix

                    'matWorld = Matrix.Identity
                    'matTemp = Matrix.Identity
                    'matTemp.RotateY(((oParent.LocAngle / 10.0F) - 90.0F) * gdRadPerDegree)
                    'matWorld.Multiply(matTemp)
                    'matTemp = Matrix.Identity
                    'matTemp.Translate(oParent.LocX, oParent.LocY, oParent.LocZ)
                    'matWorld.Multiply(matTemp)
                    'matTemp = Nothing
                    'moDevice.Transform.World = matWorld

                    With oMat
                        .Emissive = System.Drawing.Color.FromArgb(ShieldAlpha, ShieldRed, ShieldGreen, ShieldBlue)
                        .Ambient = System.Drawing.Color.Black
                        .Diffuse = System.Drawing.Color.FromArgb(ShieldAlpha, ShieldRed, ShieldGreen, ShieldBlue)
                        .Specular = System.Drawing.Color.Black
                    End With
                    moDevice.Material = oMat

                    'oParent.oMesh.oShieldMesh.DrawSubset(0)

                    'oparent.oMesh.oMesh.DrawSubset(0)
                    moDevice.Transform.World = oParent.GetWorldMatrix()
                    With oParent.oMesh
                        For X As Int32 = 0 To .NumOfMaterials - 1
                            .oMesh.DrawSubset(X)
                        Next X

                        'Now, check for a turret...
                        If .bTurretMesh = True Then
                            moDevice.Transform.World = oParent.GetTurretMatrix()
                            .oTurretMesh.DrawSubset(0)
                        End If
                    End With
                End If
            End If

            If ShieldRed = 0 AndAlso ShieldBlue = 0 AndAlso ShieldGreen = 0 Then RaiseEvent ShieldEnd(Index)
        End Sub
    End Class

    Private moFX() As ShldEffect
    Private mlFXID() As Int32
    Private miFXType() As Int16
    Private mlFXUB As Int32 = -1

    Private Shared moTexture(3) As Texture
    Private mlCurrentTex As Int32 = 0
    Private mlLastTexUpdate As Int32

    Public Sub AddNewEffect(ByRef oObj As RenderObject, ByVal clrShields As System.Drawing.Color)
        Dim X As Int32
        Dim lFirstIndex As Int32 = -1
        Dim lIdx As Int32 = -1

        If muSettings.RenderShieldFX = False Then Return

        For X = 0 To mlFXUB
            If mlFXID(X) = oObj.ObjectID AndAlso miFXType(X) = oObj.ObjTypeID Then
                lIdx = X
                Exit For
            ElseIf lFirstIndex = -1 AndAlso mlFXID(X) = -1 Then
                lFirstIndex = X
            End If
        Next X

        If lIdx = -1 Then
            If lFirstIndex = -1 Then
                mlFXUB += 1
                ReDim Preserve moFX(mlFXUB)
                ReDim Preserve mlFXID(mlFXUB)
                ReDim Preserve miFXType(mlFXUB)
                lIdx = mlFXUB

                moFX(lIdx) = New ShldEffect()
            Else
                lIdx = lFirstIndex
                moFX(lIdx) = New ShldEffect()
            End If
        End If

        mlFXID(lIdx) = oObj.ObjectID
        miFXType(lIdx) = oObj.ObjTypeID
        moFX(lIdx).oParent = oObj
        moFX(lIdx).ShieldAlpha = clrShields.A
        moFX(lIdx).ShieldBlue = clrShields.B
        moFX(lIdx).ShieldGreen = clrShields.G
        moFX(lIdx).ShieldRed = clrShields.R
        moFX(lIdx).Index = lIdx

        AddHandler moFX(lIdx).ShieldEnd, AddressOf ShieldEffectEnd
    End Sub

    Public Sub RenderFX(ByVal bUpdateNoRender As Boolean)
        Dim X As Int32

        If muSettings.RenderShieldFX = False Then Return

        Dim moDevice As Device = GFXEngine.moDevice

        'ensure our textures are ready
        For X = 0 To 3
            If moTexture(X) Is Nothing OrElse moTexture(X).Disposed = True Then
                moTexture(X) = goResMgr.GetTexture("Shield" & (X + 1) & ".dds", GFXResourceManager.eGetTextureType.ModelTexture)
            End If
        Next X

        'Set up our texture...
        If timeGetTime - mlLastTexUpdate > 30 Then
            mlCurrentTex += 1
            mlLastTexUpdate = timeGetTime
            If mlCurrentTex > moTexture.Length - 1 Then mlCurrentTex = 0
        End If

        moDevice.SetTexture(0, moTexture(mlCurrentTex))

        'Ensure culling is on
        'moDevice.RenderState.CullMode = Cull.CounterClockwise
        moDevice.RenderState.CullMode = Cull.None
        'set our wrapping, material, texture
        'moDevice.RenderState.Wrap0 = WrapCoordinates.Zero

        moDevice.RenderState.BlendOperation = BlendOperation.Max
        moDevice.RenderState.DestinationBlend = Blend.DestinationColor
        moDevice.RenderState.SourceBlend = Blend.SourceColor

        glShieldFXRendered = 0
        For X = 0 To mlFXUB
            If mlFXID(X) <> -1 Then
                glShieldFXRendered += 1
                moFX(X).Render(bUpdateNoRender)
            End If
        Next X

        moDevice.RenderState.DestinationBlend = Blend.Zero
        moDevice.RenderState.SourceBlend = Blend.One
        moDevice.RenderState.BlendOperation = BlendOperation.Add

        'reset our wrapping
        'moDevice.RenderState.Wrap0 = 0 'WrapCoordinates.Two
        moDevice.RenderState.CullMode = Cull.CounterClockwise

    End Sub

    Private Sub ShieldEffectEnd(ByVal lIndex As Int32)
        moFX(lIndex) = Nothing
        mlFXID(lIndex) = -1
        miFXType(lIndex) = -1
    End Sub

    Public Sub TurnOffMgr()
        mlFXUB = -1
        Erase moFX
        Erase mlFXID
        Erase miFXType
    End Sub

    Public Shared Sub ClearTextures()
        If moTexture Is Nothing = False Then
            For X As Int32 = 0 To moTexture.GetUpperBound(0)
                If moTexture(X) Is Nothing = False Then moTexture(X).Dispose()
                moTexture(X) = Nothing
            Next X
        End If
    End Sub
End Class
