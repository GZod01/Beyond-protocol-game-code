Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

Public Class BPSprite

    Private muVerts() As CustomVertex.PositionColoredTextured
    Private mlCurrVert As Int32 = -1
    Private mu2DVerts() As CustomVertex.TransformedColoredTextured
    Private mlCurr2DVert As Int32 = -1

    Private moTexture As Texture
    Private moDevice As Device

    Private mfImageWidth As Single
    Private mfImageHeight As Single

    Private vecCol1 As Vector3
    Private vecCol2 As Vector3
    Private vecCol3 As Vector3

    Public Sub SetVecColFromMatrix(ByVal mat As Matrix)
        'With matTemp
        '    .M11 = xAxis.X : .M12 = xAxis.Y : .M13 = xAxis.Z : .M14 = 0
        '    .M21 = yAxis.X : .M22 = yAxis.Y : .M23 = yAxis.Z : .M24 = 0
        '    .M31 = zAxis.X : .M32 = zAxis.Y : .M33 = zAxis.Z : .M34 = 0
        '    '.M41 = vecPosition.X : .M42 = vecPosition.Y : .M43 = vecPosition.Z : .M44 = 1
        'End With
        With mat
            vecCol1 = New Vector3(.M11, .M21, .M31) '  New Vector3(xAxis.X, yAxis.X, zAxis.X)
            vecCol2 = New Vector3(.M12, .M22, .M32) '  New Vector3(xAxis.Y, yAxis.Y, zAxis.Y)
            vecCol3 = New Vector3(.M13, .M23, .M33) ' New Vector3(xAxis.Z, yAxis.Z, zAxis.Z)
        End With
    End Sub

    Public Sub BeginRender(ByVal lItemCnt As Int32, ByVal oTexture As Texture, ByRef oDevice As Device)
        If lItemCnt < 0 Then lItemCnt = 0

        moDevice = oDevice
        moTexture = oTexture
        If lItemCnt = 0 Then
            ReDim muVerts(100)
            ReDim mu2DVerts(100)
        Else
            ReDim muVerts((lItemCnt * 6) - 1)
            ReDim mu2DVerts((lItemCnt * 6) - 1)
        End If
        mlCurrVert = -1
        mlCurr2DVert = -1

        If moTexture Is Nothing = False Then
            With moTexture.GetLevelDescription(0)
                mfImageWidth = .Width
                mfImageHeight = .Height
            End With
        End If

        Dim zAxis As Vector3 = New Vector3(goCamera.mlCameraX - goCamera.mlCameraAtX, goCamera.mlCameraY - goCamera.mlCameraAtY, goCamera.mlCameraZ - goCamera.mlCameraAtZ)
        zAxis.Normalize()
        Dim vecWorldUp As Vector3 = New Vector3(0, 1, 0)
        Dim xAxis As Vector3 = Vector3.Cross(zAxis, vecWorldUp)
        xAxis.Normalize()
        Dim yAxis As Vector3 = Vector3.Cross(xAxis, zAxis)
        yAxis.Normalize()

        vecCol1 = New Vector3(xAxis.X, yAxis.X, zAxis.X)
        vecCol2 = New Vector3(xAxis.Y, yAxis.Y, zAxis.Y)
        vecCol3 = New Vector3(xAxis.Z, yAxis.Z, zAxis.Z)
    End Sub

    Public Sub Draw(ByVal vecPosition As Vector3, ByVal fSize As Single, ByVal clrVal As System.Drawing.Color, ByVal fRotation As Single)
        'ok, let's determine points in space
        Dim fTempSize As Single = fSize * 0.5F

        Dim fPtX As Single = -fTempSize
        Dim fPtY As Single = fTempSize
        If fRotation <> 0 Then
            RotatePoint(0, 0, fPtX, fPtY, fRotation)
        End If

        Dim vecPt1 As Vector3 = New Vector3(-fPtX, fPtY, 0)
        Dim vecPt2 As Vector3 = New Vector3(-fPtY, -fPtX, 0)
        Dim vecPt3 As Vector3 = New Vector3(fPtY, fPtX, 0)
        Dim vecPt4 As Vector3 = New Vector3(fPtX, -fPtY, 0)

        'transform is DOT(vecPtX, New Vector3(M11,M21,M31))
        'DOT is (X1 * X2) + (Y1 * Y2) + (Z1 * Z2)
        'this is faster when dealing with large quantities of sprites... faster than DX's Vector3
        With vecPt1
            Dim fX As Single = (.X * vecCol1.X) + (.Y * vecCol1.Y) + (.Z * vecCol1.Z) + vecPosition.X
            Dim fY As Single = (.X * vecCol2.X) + (.Y * vecCol2.Y) + (.Z * vecCol2.Z) + vecPosition.Y
            Dim fZ As Single = (.X * vecCol3.X) + (.Y * vecCol3.Y) + (.Z * vecCol3.Z) + vecPosition.Z
            .X = fX : .Y = fY : .Z = fZ
        End With
        With vecPt2
            Dim fX As Single = (.X * vecCol1.X) + (.Y * vecCol1.Y) + (.Z * vecCol1.Z) + vecPosition.X
            Dim fY As Single = (.X * vecCol2.X) + (.Y * vecCol2.Y) + (.Z * vecCol2.Z) + vecPosition.Y
            Dim fZ As Single = (.X * vecCol3.X) + (.Y * vecCol3.Y) + (.Z * vecCol3.Z) + vecPosition.Z
            .X = fX : .Y = fY : .Z = fZ
        End With
        With vecPt3
            Dim fX As Single = (.X * vecCol1.X) + (.Y * vecCol1.Y) + (.Z * vecCol1.Z) + vecPosition.X
            Dim fY As Single = (.X * vecCol2.X) + (.Y * vecCol2.Y) + (.Z * vecCol2.Z) + vecPosition.Y
            Dim fZ As Single = (.X * vecCol3.X) + (.Y * vecCol3.Y) + (.Z * vecCol3.Z) + vecPosition.Z
            .X = fX : .Y = fY : .Z = fZ
        End With
        With vecPt4
            Dim fX As Single = (.X * vecCol1.X) + (.Y * vecCol1.Y) + (.Z * vecCol1.Z) + vecPosition.X
            Dim fY As Single = (.X * vecCol2.X) + (.Y * vecCol2.Y) + (.Z * vecCol2.Z) + vecPosition.Y
            Dim fZ As Single = (.X * vecCol3.X) + (.Y * vecCol3.Y) + (.Z * vecCol3.Z) + vecPosition.Z
            .X = fX : .Y = fY : .Z = fZ
        End With

        'now, billboard our position
        'lineangledegrees(gocamera.mlCameraX, gocamera.mlCameraY, vecposition.X, vecposition.Y

        Dim lClrVal As Int32 = clrVal.ToArgb
        If muVerts Is Nothing Then ReDim muVerts(-1)

        If mlCurrVert + 6 > muVerts.GetUpperBound(0) Then
            ReDim Preserve muVerts(mlCurrVert + 100)
        End If

        mlCurrVert += 1
        muVerts(mlCurrVert) = New CustomVertex.PositionColoredTextured(vecPt1, lClrVal, 0, 0)
        mlCurrVert += 1
        muVerts(mlCurrVert) = New CustomVertex.PositionColoredTextured(vecPt2, lClrVal, 0, 1)
        mlCurrVert += 1
        muVerts(mlCurrVert) = New CustomVertex.PositionColoredTextured(vecPt3, lClrVal, 1, 0)

        mlCurrVert += 1
        muVerts(mlCurrVert) = New CustomVertex.PositionColoredTextured(vecPt2, lClrVal, 0, 1)
        mlCurrVert += 1
        muVerts(mlCurrVert) = New CustomVertex.PositionColoredTextured(vecPt3, lClrVal, 1, 0)
        mlCurrVert += 1
        muVerts(mlCurrVert) = New CustomVertex.PositionColoredTextured(vecPt4, lClrVal, 1, 1)

        'position
        'color
        'texture

    End Sub

    Public Sub Draw(ByVal vecPosition As Vector3, ByVal fSize As Single, ByVal clrVal As System.Drawing.Color, ByVal fRotation As Single, ByVal srcRect As Rectangle)
        'ok, let's determine points in space
        Dim fTempSize As Single = fSize * 0.5F

        'With matTemp
        '    .M41 = vecPosition.X : .M42 = vecPosition.Y : .M43 = vecPosition.Z : .M44 = 1
        'End With

        Dim fPtX As Single = -fTempSize
        Dim fPtY As Single = fTempSize
        If fRotation <> 0 Then
            RotatePoint(0, 0, fPtX, fPtY, fRotation)
        End If

        Dim vecPt1 As Vector3 = New Vector3(-fPtX, fPtY, 0)
        Dim vecPt2 As Vector3 = New Vector3(-fPtY, -fPtX, 0)
        Dim vecPt3 As Vector3 = New Vector3(fPtY, fPtX, 0)
        Dim vecPt4 As Vector3 = New Vector3(fPtX, -fPtY, 0)

        'transform is DOT(vecPtX, New Vector3(M11,M21,M31))
        'DOT is (X1 * X2) + (Y1 * Y2) + (Z1 * Z2)
        'this is faster when dealing with large quantities of sprites... faster than DX's Vector3
        With vecPt1
            Dim fX As Single = (.X * vecCol1.X) + (.Y * vecCol1.Y) + (.Z * vecCol1.Z) + vecPosition.X
            Dim fY As Single = (.X * vecCol2.X) + (.Y * vecCol2.Y) + (.Z * vecCol2.Z) + vecPosition.Y
            Dim fZ As Single = (.X * vecCol3.X) + (.Y * vecCol3.Y) + (.Z * vecCol3.Z) + vecPosition.Z
            .X = fX : .Y = fY : .Z = fZ
        End With
        With vecPt2
            Dim fX As Single = (.X * vecCol1.X) + (.Y * vecCol1.Y) + (.Z * vecCol1.Z) + vecPosition.X
            Dim fY As Single = (.X * vecCol2.X) + (.Y * vecCol2.Y) + (.Z * vecCol2.Z) + vecPosition.Y
            Dim fZ As Single = (.X * vecCol3.X) + (.Y * vecCol3.Y) + (.Z * vecCol3.Z) + vecPosition.Z
            .X = fX : .Y = fY : .Z = fZ
        End With
        With vecPt3
            Dim fX As Single = (.X * vecCol1.X) + (.Y * vecCol1.Y) + (.Z * vecCol1.Z) + vecPosition.X
            Dim fY As Single = (.X * vecCol2.X) + (.Y * vecCol2.Y) + (.Z * vecCol2.Z) + vecPosition.Y
            Dim fZ As Single = (.X * vecCol3.X) + (.Y * vecCol3.Y) + (.Z * vecCol3.Z) + vecPosition.Z
            .X = fX : .Y = fY : .Z = fZ
        End With
        With vecPt4
            Dim fX As Single = (.X * vecCol1.X) + (.Y * vecCol1.Y) + (.Z * vecCol1.Z) + vecPosition.X
            Dim fY As Single = (.X * vecCol2.X) + (.Y * vecCol2.Y) + (.Z * vecCol2.Z) + vecPosition.Y
            Dim fZ As Single = (.X * vecCol3.X) + (.Y * vecCol3.Y) + (.Z * vecCol3.Z) + vecPosition.Z
            .X = fX : .Y = fY : .Z = fZ
        End With


        Dim lClrVal As Int32 = clrVal.ToArgb

        Dim fTU0 As Single = srcRect.X / mfImageWidth
        Dim fTV0 As Single = srcRect.Y / mfImageHeight
        Dim fTU1 As Single = srcRect.Right / mfImageWidth
        Dim fTV1 As Single = srcRect.Bottom / mfImageHeight

        If mlCurrVert + 6 > muVerts.GetUpperBound(0) Then
            ReDim Preserve muVerts(mlCurrVert + 100)
        End If

        mlCurrVert += 1
        muVerts(mlCurrVert) = New CustomVertex.PositionColoredTextured(vecPt1, lClrVal, fTU0, fTV0)
        mlCurrVert += 1
        muVerts(mlCurrVert) = New CustomVertex.PositionColoredTextured(vecPt2, lClrVal, fTU0, fTV1)
        mlCurrVert += 1
        muVerts(mlCurrVert) = New CustomVertex.PositionColoredTextured(vecPt3, lClrVal, fTU1, fTV0)

        mlCurrVert += 1
        muVerts(mlCurrVert) = New CustomVertex.PositionColoredTextured(vecPt2, lClrVal, fTU0, fTV1)
        mlCurrVert += 1
        muVerts(mlCurrVert) = New CustomVertex.PositionColoredTextured(vecPt3, lClrVal, fTU1, fTV0)
        mlCurrVert += 1
        muVerts(mlCurrVert) = New CustomVertex.PositionColoredTextured(vecPt4, lClrVal, fTU1, fTV1)

        'position
        'color
        'texture

    End Sub

    Public Sub EndRender()
        If moDevice Is Nothing Then Return

        If mlCurr2DVert > 2 AndAlso mu2DVerts Is Nothing = False Then
            moDevice.RenderState.CullMode = Cull.None
            moDevice.SetTexture(0, moTexture)
            moDevice.VertexFormat = CustomVertex.TransformedColoredTextured.Format
            moDevice.DrawUserPrimitives(PrimitiveType.TriangleList, (mlCurr2DVert + 1) \ 3, mu2DVerts)
            Erase mu2DVerts
        End If

        If muVerts Is Nothing Then Return
        If mlCurrVert < 3 Then Return


        With moDevice
            .SetTextureStageState(0, TextureStageStates.AlphaArgument1, TextureArgument.TextureColor)
            .SetTextureStageState(0, TextureStageStates.AlphaArgument2, TextureArgument.Diffuse)
            .SetTextureStageState(0, TextureStageStates.AlphaOperation, TextureOperation.Modulate)
        End With

        'do our draw call here
        Dim lCull As Cull = moDevice.RenderState.CullMode
        moDevice.RenderState.ZBufferWriteEnable = False
        moDevice.RenderState.CullMode = Cull.None
        moDevice.RenderState.Lighting = False
        moDevice.SetTexture(0, moTexture)
        moDevice.VertexFormat = CustomVertex.PositionColoredTextured.Format
        moDevice.DrawUserPrimitives(PrimitiveType.TriangleList, (mlCurrVert + 1) \ 3, muVerts)
        moDevice.RenderState.Lighting = True
        moDevice.RenderState.CullMode = lCull
        moDevice.RenderState.ZBufferWriteEnable = True

        With moDevice
            .SetTextureStageState(0, TextureStageStates.AlphaArgument1, TextureArgument.TextureColor)
            .SetTextureStageState(0, TextureStageStates.AlphaArgument2, TextureArgument.Current)
            .SetTextureStageState(0, TextureStageStates.AlphaOperation, TextureOperation.SelectArg1)
        End With

        'now kill our stuff
        Erase muVerts
        moTexture = Nothing
        moDevice = Nothing
    End Sub

    Public Sub EndRenderNoStateChange()
        If moDevice Is Nothing Then Return

        If mlCurr2DVert > 2 AndAlso mu2DVerts Is Nothing = False Then
            moDevice.RenderState.CullMode = Cull.None
            moDevice.SetTexture(0, moTexture)
            moDevice.VertexFormat = CustomVertex.TransformedColoredTextured.Format
            moDevice.DrawUserPrimitives(PrimitiveType.TriangleList, (mlCurr2DVert + 1) \ 3, mu2DVerts)
            Erase mu2DVerts
        End If

        If muVerts Is Nothing Then Return
        If mlCurrVert < 3 Then Return


        moDevice.SetTexture(0, moTexture)
        moDevice.VertexFormat = CustomVertex.PositionColoredTextured.Format
        moDevice.DrawUserPrimitives(PrimitiveType.TriangleList, (mlCurrVert + 1) \ 3, muVerts)

        'now kill our stuff
        Erase muVerts
        moTexture = Nothing
        moDevice = Nothing
    End Sub

    Public Sub DisposeMe()
        moDevice = Nothing
        moTexture = Nothing
        Erase muVerts
    End Sub

    Protected Overrides Sub Finalize()
        DisposeMe()
        MyBase.Finalize()
    End Sub

    Private Shared fStringTU(,) As Single
    Private Shared fStringTV(,) As Single
    Public Shared Function GetStringRectIDs(ByVal sString As String) As Int32()
        Dim lResult(sString.Length - 1) As Int32

        If fStringTU Is Nothing Then
            ReDim fStringTU(11, 5)
            ReDim fStringTV(11, 5)

            For X As Int32 = 0 To 11

                '128
                Dim rcSrc As Rectangle
                Select Case X
                    Case 0
                        rcSrc = New Rectangle(0, 2, 18, 22)
                    Case 1
                        rcSrc = New Rectangle(19, 2, 20, 22)
                    Case 2
                        rcSrc = New Rectangle(39, 2, 16, 22)
                    Case 3
                        rcSrc = New Rectangle(59, 2, 20, 22)
                    Case 4
                        rcSrc = New Rectangle(82, 2, 20, 22)
                    Case 5
                        rcSrc = New Rectangle(104, 2, 20, 22)
                    Case 6
                        rcSrc = New Rectangle(2, 28, 20, 22)
                    Case 7
                        rcSrc = New Rectangle(25, 28, 20, 22)
                    Case 8
                        rcSrc = New Rectangle(49, 28, 20, 22)
                    Case 9
                        rcSrc = New Rectangle(73, 28, 20, 22)
                    Case 10
                        rcSrc = New Rectangle(97, 28, 20, 22)
                    Case 11
                        rcSrc = New Rectangle(119, 28, 9, 22)
                End Select

                Dim fLeft As Single = rcSrc.X / 128.0F
                Dim fTop As Single = rcSrc.Y / 128.0F
                Dim fRight As Single = rcSrc.Right / 128.0F
                Dim fBottom As Single = rcSrc.Bottom / 128.0F

                '1,0
                fStringTU(X, 0) = fRight : fStringTV(X, 0) = fTop
                '0,0
                fStringTU(X, 1) = fLeft : fStringTV(X, 1) = fTop
                '1,1
                fStringTU(X, 2) = fRight : fStringTV(X, 2) = fBottom
                '0,0
                fStringTU(X, 3) = fLeft : fStringTV(X, 3) = fTop
                '1,1
                fStringTU(X, 4) = fRight : fStringTV(X, 4) = fBottom
                '0,1
                fStringTU(X, 5) = fLeft : fStringTV(X, 5) = fBottom
            Next X
        End If

        For X As Int32 = 0 To sString.Length - 1
            Select Case sString.Substring(X, 1)
                Case "+"
                    lResult(X) = 0
                Case "0"
                    lResult(X) = 1
                Case "1"
                    lResult(X) = 2
                Case "2"
                    lResult(X) = 3
                Case "3"
                    lResult(X) = 4
                Case "4"
                    lResult(X) = 5
                Case "5"
                    lResult(X) = 6
                Case "6"
                    lResult(X) = 7
                Case "7"
                    lResult(X) = 8
                Case "8"
                    lResult(X) = 9
                Case "9"
                    lResult(X) = 10
                Case ","
                    lResult(X) = 11
            End Select
        Next X

        Return lResult
    End Function
    Public Sub DrawStringIn3DSpace(ByVal lStringLocIDs() As Int32, ByVal fBaseX As Single, ByVal fBaseY As Single, ByVal fBaseZ As Single, ByRef oTex As Texture, ByVal clr As System.Drawing.Color, ByVal fFontSize As Single)
        ''ok, let's determine points in space
        Dim fTempSize As Single = fFontSize
        Dim fPtX As Single = -fTempSize
        Dim fPtY As Single = fTempSize
        Dim fRootX As Single = 0.0F



        'We want the vector between pt 1 and pt 4
        Dim lClrVal As Int32 = clr.ToArgb
        'Now, go thru each car
        If mlCurrVert + (6 * lStringLocIDs.Length) > muVerts.GetUpperBound(0) Then
            ReDim Preserve muVerts(Math.Max(mlCurrVert + 100, mlCurrVert + (6 * lStringLocIDs.Length)))
        End If
        Dim vecPosition As New Vector3(fBaseX, fBaseY, fBaseZ)

        For X As Int32 = 0 To lStringLocIDs.GetUpperBound(0)
            'Dim vecSkewVal As Vector3 = (vecSkew * X) + vecPosition

            Dim vecPt1 As Vector3 = New Vector3(fRootX - fPtX, fPtY, 0)
            Dim vecPt2 As Vector3 = New Vector3(fRootX - fPtY, -fPtX, 0)
            Dim vecPt3 As Vector3 = New Vector3(fRootX + fPtY, fPtX, 0)
            Dim vecPt4 As Vector3 = New Vector3(fRootX + fPtX, -fPtY, 0)

            'transform is DOT(vecPtX, New Vector3(M11,M21,M31))
            'DOT is (X1 * X2) + (Y1 * Y2) + (Z1 * Z2)
            'this is faster when dealing with large quantities of sprites... faster than DX's Vector3
            With vecPt1
                Dim fX As Single = (.X * vecCol1.X) + (.Y * vecCol1.Y) + (.Z * vecCol1.Z) + vecPosition.X
                Dim fY As Single = (.X * vecCol2.X) + (.Y * vecCol2.Y) + (.Z * vecCol2.Z) + vecPosition.Y
                Dim fZ As Single = (.X * vecCol3.X) + (.Y * vecCol3.Y) + (.Z * vecCol3.Z) + vecPosition.Z
                .X = fX : .Y = fY : .Z = fZ
            End With
            With vecPt2
                Dim fX As Single = (.X * vecCol1.X) + (.Y * vecCol1.Y) + (.Z * vecCol1.Z) + vecPosition.X
                Dim fY As Single = (.X * vecCol2.X) + (.Y * vecCol2.Y) + (.Z * vecCol2.Z) + vecPosition.Y
                Dim fZ As Single = (.X * vecCol3.X) + (.Y * vecCol3.Y) + (.Z * vecCol3.Z) + vecPosition.Z
                .X = fX : .Y = fY : .Z = fZ
            End With
            With vecPt3
                Dim fX As Single = (.X * vecCol1.X) + (.Y * vecCol1.Y) + (.Z * vecCol1.Z) + vecPosition.X
                Dim fY As Single = (.X * vecCol2.X) + (.Y * vecCol2.Y) + (.Z * vecCol2.Z) + vecPosition.Y
                Dim fZ As Single = (.X * vecCol3.X) + (.Y * vecCol3.Y) + (.Z * vecCol3.Z) + vecPosition.Z
                .X = fX : .Y = fY : .Z = fZ
            End With
            With vecPt4
                Dim fX As Single = (.X * vecCol1.X) + (.Y * vecCol1.Y) + (.Z * vecCol1.Z) + vecPosition.X
                Dim fY As Single = (.X * vecCol2.X) + (.Y * vecCol2.Y) + (.Z * vecCol2.Z) + vecPosition.Y
                Dim fZ As Single = (.X * vecCol3.X) + (.Y * vecCol3.Y) + (.Z * vecCol3.Z) + vecPosition.Z
                .X = fX : .Y = fY : .Z = fZ
            End With


            Dim lLocID As Int32 = lStringLocIDs(X)

            mlCurrVert += 1
            muVerts(mlCurrVert) = New CustomVertex.PositionColoredTextured(vecPt1, lClrVal, fStringTU(lLocID, 0), fStringTV(lLocID, 0))
            mlCurrVert += 1
            muVerts(mlCurrVert) = New CustomVertex.PositionColoredTextured(vecPt2, lClrVal, fStringTU(lLocID, 1), fStringTV(lLocID, 1))
            mlCurrVert += 1
            muVerts(mlCurrVert) = New CustomVertex.PositionColoredTextured(vecPt3, lClrVal, fStringTU(lLocID, 2), fStringTV(lLocID, 2))

            mlCurrVert += 1
            muVerts(mlCurrVert) = New CustomVertex.PositionColoredTextured(vecPt2, lClrVal, fStringTU(lLocID, 3), fStringTV(lLocID, 3))
            mlCurrVert += 1
            muVerts(mlCurrVert) = New CustomVertex.PositionColoredTextured(vecPt3, lClrVal, fStringTU(lLocID, 4), fStringTV(lLocID, 4))
            mlCurrVert += 1
            muVerts(mlCurrVert) = New CustomVertex.PositionColoredTextured(vecPt4, lClrVal, fStringTU(lLocID, 5), fStringTV(lLocID, 5))

            fRootX += (fTempSize * 2)
        Next X
    End Sub

#Region "  2D Sprite Rendering  "

    Public Sub Draw2D(ByVal srcRect As Rectangle, ByVal destRect As Rectangle, ByVal clr As System.Drawing.Color)
        Dim fLeftTu As Single = srcRect.X / mfImageWidth
        Dim fTopTv As Single = srcRect.Y / mfImageHeight
        Dim fRightTu As Single = srcRect.Right / mfImageWidth
        Dim fBottomTv As Single = srcRect.Bottom / mfImageHeight

        Dim lClrVal As Int32 = clr.ToArgb
        If mlCurr2DVert + 6 > mu2DVerts.GetUpperBound(0) Then
            ReDim Preserve mu2DVerts(mlCurr2DVert + 100)
        End If

        '0, 1, 2
        '1, 2, 3
        mlCurr2DVert += 1
        mu2DVerts(mlCurr2DVert) = New CustomVertex.TransformedColoredTextured(destRect.X, destRect.Bottom, 0, 1, lClrVal, fLeftTu, fBottomTv)
        mlCurr2DVert += 1
        mu2DVerts(mlCurr2DVert) = New CustomVertex.TransformedColoredTextured(destRect.X, destRect.Y, 0, 1, lClrVal, fLeftTu, fTopTv)
        mlCurr2DVert += 1
        mu2DVerts(mlCurr2DVert) = New CustomVertex.TransformedColoredTextured(destRect.Right, destRect.Bottom, 0, 1, lClrVal, fRightTu, fBottomTv)

        mlCurr2DVert += 1
        mu2DVerts(mlCurr2DVert) = New CustomVertex.TransformedColoredTextured(destRect.X, destRect.Y, 0, 1, lClrVal, fLeftTu, fTopTv)
        mlCurr2DVert += 1
        mu2DVerts(mlCurr2DVert) = New CustomVertex.TransformedColoredTextured(destRect.Right, destRect.Bottom, 0, 1, lClrVal, fRightTu, fBottomTv)
        mlCurr2DVert += 1
        mu2DVerts(mlCurr2DVert) = New CustomVertex.TransformedColoredTextured(destRect.Right, destRect.Y, 0, 1, lClrVal, fRightTu, fTopTv)


    End Sub

    Public Shared Sub Draw2DOnce(ByRef oDevice As Device, ByRef oTex As Texture, ByVal srcRect As Rectangle, ByVal destRect As Rectangle, ByVal clr As System.Drawing.Color, ByVal lTextureWidth As Int32, ByVal lTextureHeight As Int32)
        Dim uVerts(3) As CustomVertex.TransformedColoredTextured

        Dim fLeftTu As Single = CSng(srcRect.X / lTextureWidth)
        Dim fTopTv As Single = CSng(srcRect.Y / lTextureHeight)
        Dim fRightTu As Single = CSng(srcRect.Right / lTextureWidth)
        Dim fBottomTv As Single = CSng(srcRect.Bottom / lTextureHeight)

        Dim lClrVal As Int32 = clr.ToArgb
        uVerts(0) = New CustomVertex.TransformedColoredTextured(destRect.X, destRect.Bottom, 0, 1, lClrVal, fLeftTu, fBottomTv)
        uVerts(1) = New CustomVertex.TransformedColoredTextured(destRect.X, destRect.Y, 0, 1, lClrVal, fLeftTu, fTopTv)
        uVerts(2) = New CustomVertex.TransformedColoredTextured(destRect.Right, destRect.Bottom, 0, 1, lClrVal, fRightTu, fBottomTv)
        uVerts(3) = New CustomVertex.TransformedColoredTextured(destRect.Right, destRect.Y, 0, 1, lClrVal, fRightTu, fTopTv)

        oDevice.RenderState.CullMode = Cull.None
        oDevice.SetTexture(0, oTex)
        oDevice.VertexFormat = CustomVertex.TransformedColoredTextured.Format
        oDevice.DrawUserPrimitives(PrimitiveType.TriangleStrip, 2, uVerts)
    End Sub

    Public Shared Sub Draw2DOnceRotation(ByRef oDevice As Device, ByRef oTex As Texture, ByVal srcRect As Rectangle, ByVal destRect As Rectangle, ByVal clr As System.Drawing.Color, ByVal lTextureWidth As Int32, ByVal lTextureHeight As Int32, ByVal fRotation As Single)
        Dim uVerts(3) As CustomVertex.TransformedColoredTextured

        Dim fLeftTu As Single = CSng(srcRect.X / lTextureWidth)
        Dim fTopTv As Single = CSng(srcRect.Y / lTextureHeight)
        Dim fRightTu As Single = CSng(srcRect.Right / lTextureWidth)
        Dim fBottomTv As Single = CSng(srcRect.Bottom / lTextureHeight)

        Dim fX1 As Single = destRect.X
        Dim fY1 As Single = destRect.Y
        Dim fX2 As Single = destRect.Right
        Dim fY2 As Single = destRect.Bottom
        Dim lCX As Int32 = destRect.X + (destRect.Width \ 2)
        Dim lCY As Int32 = destRect.Y + (destRect.Height \ 2)

        Dim fPt0X As Single = fX1
        Dim fPt0Y As Single = fY2
        RotatePoint(lCX, lCY, fPt0X, fPt0Y, fRotation)

        Dim fPt1X As Single = fX1
        Dim fPt1Y As Single = fY1
        RotatePoint(lCX, lCY, fPt1X, fPt1Y, fRotation)

        Dim fPt2X As Single = fX2
        Dim fPt2Y As Single = fY2
        RotatePoint(lCX, lCY, fPt2X, fPt2Y, fRotation)

        Dim fPt3X As Single = fX2
        Dim fPt3Y As Single = fY1
        RotatePoint(lCX, lCY, fPt3X, fPt3Y, fRotation)

        Dim lClrVal As Int32 = clr.ToArgb
        uVerts(0) = New CustomVertex.TransformedColoredTextured(fPt0X, fPt0Y, 0, 1, lClrVal, fLeftTu, fBottomTv)
        uVerts(1) = New CustomVertex.TransformedColoredTextured(fPt1X, fPt1Y, 0, 1, lClrVal, fLeftTu, fTopTv)
        uVerts(2) = New CustomVertex.TransformedColoredTextured(fPt2X, fPt2Y, 0, 1, lClrVal, fRightTu, fBottomTv)
        uVerts(3) = New CustomVertex.TransformedColoredTextured(fPt3X, fPt3Y, 0, 1, lClrVal, fRightTu, fTopTv)

        oDevice.RenderState.CullMode = Cull.None
        oDevice.SetTexture(0, oTex)
        oDevice.VertexFormat = CustomVertex.TransformedColoredTextured.Format
        oDevice.DrawUserPrimitives(PrimitiveType.TriangleStrip, 2, uVerts)
    End Sub

#End Region

End Class