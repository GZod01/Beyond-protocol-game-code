Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

Public Structure BurnEffectPoint
    Public vecLoc As Vector3
    Public fBurnAlpha As Single
    Public fRotation As Single
    Public bActive As Boolean
    Public ySide As Byte
    Public fSize As Single
End Structure

Public Class EntityBurnMarkManager
    'ok, here's how this works...
    ' 2 matrices are used...
    ' so 2 sprites
    Private moForwardAligned As BPSprite
    Private moLeftAligned As BPSprite
    Public Shared moBurnTex As Texture

    Public Sub New()
        moForwardAligned = New BPSprite()
        moLeftAligned = New BPSprite
    End Sub
    Public Sub BeginFrame()

        If moBurnTex Is Nothing OrElse moBurnTex.Disposed = True Then
            moBurnTex = Nothing
            moBurnTex = goResMgr.GetTexture("p3.dds", GFXResourceManager.eGetTextureType.NoSpecifics, "xpl.pak")
        End If

        Dim moDevice As Device = GFXEngine.moDevice
        moForwardAligned.BeginRender(0, moBurnTex, moDevice)
        moLeftAligned.BeginRender(0, moBurnTex, moDevice)
    End Sub

    Public Sub RenderFromObj(ByRef oObj As BaseEntity)
        If oObj Is Nothing Then Return
        If oObj.muBurnMarks Is Nothing Then Return
        Dim bMatsSet As Boolean = False
        Dim matWorld As Matrix
        Dim matPerp As Matrix
        'If oObj.ObjectID = 1118028 Then Stop
        For X As Int32 = 0 To oObj.muBurnMarks.GetUpperBound(0)
            If oObj.muBurnMarks(X).bActive = True Then
                If bMatsSet = False Then
                    matWorld = oObj.GetWorldMatrix()
                    matPerp = oObj.GetPerpWorldMatrix

                    moLeftAligned.SetVecColFromMatrix(matWorld)
                    moForwardAligned.SetVecColFromMatrix(matPerp)
                    bMatsSet = True
                End If
                With oObj.muBurnMarks(X)
                    .fBurnAlpha -= 1.0F
                    If .fBurnAlpha < 0 Then
                        .bActive = False
                    Else
                        'Dim vecTemp As Vector3 = Vector3.TransformCoordinate(.vecLoc, matWorld)
                        If .ySide = 0 OrElse .ySide = 2 Then
                            Dim vecTemp As Vector3
                            Dim fX As Single = .vecLoc.X
                            Dim fY As Single = .vecLoc.Y
                            Dim fZ As Single = .vecLoc.Z
                            With matPerp
                                vecTemp.X = (fX * .M11) + (fY * .M21) + (fZ * .M31) + .M41
                                vecTemp.Y = (fX * .M12) + (fY * .M22) + (fZ * .M32) + .M42
                                vecTemp.Z = (fX * .M13) + (fY * .M23) + (fZ * .M33) + .M43
                            End With
                            moForwardAligned.Draw(vecTemp, .fSize, Color.FromArgb(CInt(.fBurnAlpha), 0, 0, 0), .fRotation)
                        Else
                            Dim vecTemp As Vector3
                            Dim fX As Single = .vecLoc.X
                            Dim fY As Single = .vecLoc.Y
                            Dim fZ As Single = .vecLoc.Z
                            With matWorld
                                vecTemp.X = (fX * .M11) + (fY * .M21) + (fZ * .M31) + .M41
                                vecTemp.Y = (fX * .M12) + (fY * .M22) + (fZ * .M32) + .M42
                                vecTemp.Z = (fX * .M13) + (fY * .M23) + (fZ * .M33) + .M43
                            End With
                            moLeftAligned.Draw(vecTemp, .fSize, Color.FromArgb(CInt(.fBurnAlpha), 0, 0, 0), .fRotation)
                        End If
                    End If
                End With
            End If
        Next X

    End Sub

    Public Sub EndFrame()
        moForwardAligned.EndRender()
        moLeftAligned.EndRender()
    End Sub

End Class
