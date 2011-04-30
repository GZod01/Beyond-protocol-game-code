Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

Public Class BaseMesh
    Public lModelID As Int32
	Public oMesh As Mesh

    Public oSharedTex() As GFXResourceManager.SharedTexture
    Public sSharedTexName() As String
    Public lSharedTexOpt() As Int32
    Public bMTTexCapable As Boolean = False
    Public lDefaultSharedTexIdx As Int32 = 0

	Public oNormalMap As Texture
	Public sNormalMap As String
	Public sNormalMapPak As String
	Public oIllumMap As Texture
	Public sIllumMap As String
	Public sIllumMapPak As String

    Public oShieldMesh As Mesh

    Public NumOfMaterials As Int32 = -1
    Public Materials() As Material
    Public Textures() As Texture
	Public sTextures() As String		'for getting back to the res mgr
	Public sTexturePak() As String

    Public YMidPoint As Int32
	Public PlanetYAdjust As Int32

	'for y-level adjusting
	Public YLevelEntityHeight As Int32
	Public YLevelRectSize As Int32

    Public EngineFXCnt As Int32
    Public EngineFXOffset() As Vector3
    Public EngineFXPCnt() As Int32

    Public ShieldXZRadius As Single

    Public RangeOffset As Int32         'a value added to all ranges...

    Public bLandBased As Boolean = False

    Public bTurretMesh As Boolean = False
    Public oTurretMesh As Mesh
    Public lTurretZOffset As Int32

    Public bHasLights As Boolean = False
    Public Lights() As LightInfo

    Public muFireLocs()() As FireFromLoc

    Public muBurnLocs() As BurnFXData

    'For Death Sequencing
    Public vecDeathSeqSize As Vector3
    Public vecHalfDeathSeqSize As Vector3       'to reduce a lot of calculations
    Public lDeathSeqExpCnt As Int32
    Public bDeathSeqFinale As Boolean
    Public sDeathSeqWav() As String

    Public sRoarSFX As String = ""

#Region "  Fire to Locs  "
    Public mvecMeshPoints()() As Vector3
    Private mbMeshPointsLoaded As Boolean = False
    Public Sub LoadMeshPoints()
        If mbMeshPointsLoaded = True Then Return
        mbMeshPointsLoaded = True

        'Ok, we need to lock the mesh... let's see what the declaration of the mesh is...
        Dim decl() As VertexElement = oMesh.Declaration
        Dim lPosOffset As Int32 = 0
        Dim lNormalOffset As Int32 = 0
        Dim lTotal As Int32 = 0
        If decl Is Nothing = False Then
            For X As Int32 = 0 To decl.GetUpperBound(0)
                If Object.Equals(decl(X), VertexElement.VertexDeclarationEnd) Then

                    If X <> 0 Then
                        lTotal = decl(X - 1).Offset
                        Select Case decl(X - 1).DeclarationType
                            Case DeclarationType.Float2
                                lTotal += 8
                            Case DeclarationType.Float3
                                lTotal += 12
                            Case DeclarationType.Float4
                                lTotal += 16
                        End Select
                    End If

                    Exit For
                End If
                If decl(X).DeclarationUsage = DeclarationUsage.Position Then
                    lPosOffset = decl(X).Offset
                ElseIf decl(X).DeclarationUsage = DeclarationUsage.Normal Then
                    lNormalOffset = decl(X).Offset
                End If
            Next X
        End If

        ReDim mvecMeshPoints(3) '(lTotalElems) '()
        Dim lTotalElems As Int32 = 0
        If lTotal < 1 Then
            'add 0,0 to each side...
            Dim vecTemp(0) As Vector3
            vecTemp(0) = New Vector3(0, 0, 0)
            For X As Int32 = 0 To 3
                mvecMeshPoints(X) = vecTemp
            Next X
            Return
        End If

        Dim yData() As Byte = Nothing

        Try
            Dim ranks(0) As Int32
            ranks(0) = oMesh.NumberVertices * lTotal
            yData = CType(oMesh.VertexBuffer.Lock(0, New Byte().GetType(), LockFlags.ReadOnly, ranks), Byte())
            lTotalElems = yData.Length \ lTotal
            If lTotalElems <> oMesh.NumberVertices Then lTotalElems = Math.Min(lTotalElems, oMesh.NumberVertices)
        Catch
            'add 0,0 to each side...
            Dim vecTemp(0) As Vector3
            vecTemp(0) = New Vector3(0, 0, 0)
            For X As Int32 = 0 To 3
                mvecMeshPoints(X) = vecTemp
            Next X
            Return
        End Try

        'give enough room for the entire list...
        For X As Int32 = 0 To 3
            Dim vecTemp(lTotalElems) As Vector3
            mvecMeshPoints(X) = vecTemp
        Next X

        'For storing which index in the array we are on...
        Dim lSideIdx() As Int32 = {-1, -1, -1, -1}
        For i As Int32 = 0 To lTotalElems - 1
            Try
                Dim lVertexPos As Int32 = i * lTotal
                Dim lPos As Int32 = lVertexPos + lPosOffset
                Dim fLocX As Single = System.BitConverter.ToSingle(yData, lPos) : lPos += 4
                Dim fLocY As Single = System.BitConverter.ToSingle(yData, lPos) : lPos += 4
                Dim fLocZ As Single = System.BitConverter.ToSingle(yData, lPos) : lPos += 4
                lPos = lVertexPos + lNormalOffset
                Dim fNx As Single = System.BitConverter.ToSingle(yData, lPos) : lPos += 4
                Dim fNy As Single = System.BitConverter.ToSingle(yData, lPos) : lPos += 4
                Dim fNz As Single = System.BitConverter.ToSingle(yData, lPos) : lPos += 4

                'if this vertex points "up" or "down" essentially... than disregard it
                If Math.Abs(fNy) < 0.3F Then
                    'Now, determine where to place this... first... left vs. right
                    If Math.Abs(fNz) < 0.3F Then    'disregard more "forward" or "rear" facing items
                        Dim lSideArc As Int32 = -1
                        If fNx < 0 AndAlso fLocX < 0 Then
                            'rightside
                            lSideArc = UnitArcs.eRightArc
                        ElseIf fNx > 0 AndAlso fLocX > 0 Then
                            'leftside
                            lSideArc = UnitArcs.eLeftArc
                        End If
                        If lSideArc <> -1 Then
                            lSideIdx(lSideArc) += 1
                            mvecMeshPoints(lSideArc)(lSideIdx(lSideArc)) = New Vector3(fLocX, fLocY, fLocZ)
                        End If
                    End If
                    'Now, determine front vs. back
                    If Math.Abs(fNx) < 0.3F Then    'disregard more "left" or "right" facing items
                        Dim lSideArc As Int32 = -1
                        If fNz < 0 AndAlso fLocZ < 0 Then
                            'forward?
                            lSideArc = UnitArcs.eForwardArc
                        ElseIf fNz > 0 AndAlso fLocZ > 0 Then
                            'rear?
                            lSideArc = UnitArcs.eBackArc
                        End If
                        If lSideArc <> -1 Then
                            lSideIdx(lSideArc) += 1
                            mvecMeshPoints(lSideArc)(lSideIdx(lSideArc)) = New Vector3(fLocX, fLocY, fLocZ)
                        End If
                    End If
                End If
            Catch
            End Try
        Next i

        For X As Int32 = 0 To 3
            Dim vecTemp() As Vector3 = mvecMeshPoints(X)

            If lSideIdx(X) < 0 Then
                lSideIdx(X) += 1
                vecTemp(0) = New Vector3(0, 0, 0)
            End If
            ReDim Preserve vecTemp(lSideIdx(X))
            mvecMeshPoints(X) = vecTemp
        Next X
        oMesh.VertexBuffer.Unlock()
    End Sub

    Public Function GetRandomMeshPoint(ByVal ySide As Byte) As Vector3
        If ySide > 3 Then ySide = 3
        Dim lUB As Int32 = mvecMeshPoints(ySide).GetUpperBound(0)
        If lUB = -1 Then Return New Vector3(0, 0, 0)
        Dim lIdx As Int32 = CInt(Rnd() * lUB)
        If lIdx < 0 Then lIdx = 0
        If lIdx > lUB Then lIdx = lUB
        Return mvecMeshPoints(ySide)(lIdx)
    End Function
#End Region

    Public Sub CleanMe()
        ' BaseMesh.oMesh, oSharedTex(), oNormalMap, oIllumMap, oShieldMesh, Textures(), oTurretMesh
        oMesh = Nothing
        ReDim oSharedTex(-1)
        oNormalMap = Nothing
        oIllumMap = Nothing
        oShieldMesh = Nothing
        ReDim Textures(-1)
        oTurretMesh = Nothing
    End Sub

    Public Sub New()
        'Some test data now...
        ReDim muFireLocs(3)
    End Sub
End Class

Public Structure BurnFXData
    Public lSide As Int32
    Public lOffsetX As Int32
    Public lOffsetY As Int32
    Public lOffsetZ As Int32
    Public lPCnt As Int32

    Private mlModVal As Int32
    Private mlBaseVal As Int32

    Public ReadOnly Property ModValue() As Int32
        Get
            If (mlEmitterType And BurnFX.ParticleEngine.EmitterType.eFireEmitter) <> 0 Then
                mlBaseVal = BurnFX.ParticleEngine.EmitterType.eFireEmitter
                Return mlEmitterType - mlBaseVal
            ElseIf (mlEmitterType And BurnFX.ParticleEngine.EmitterType.eSmokeyFireEmitter) <> 0 Then
                mlBaseVal = BurnFX.ParticleEngine.EmitterType.eSmokeyFireEmitter
                Return mlEmitterType - mlBaseVal
            End If
            Return 0
            'mlBaseVal = mlEmitterType
            'mlModVal = 0
            'Return mlModVal
        End Get
    End Property

    Public ReadOnly Property BaseValue() As Int32
        Get
            Return mlBaseVal
        End Get
    End Property

    Private mlEmitterType As Int32
    Public Property lEmitterType() As Int32
        Get
            Return mlEmitterType
        End Get
        Set(ByVal value As Int32)
            mlEmitterType = value
            If (mlEmitterType And BurnFX.ParticleEngine.EmitterType.eFireEmitter) <> 0 Then
                mlBaseVal = BurnFX.ParticleEngine.EmitterType.eFireEmitter
                mlModVal = mlEmitterType - mlBaseVal
            ElseIf (mlEmitterType And BurnFX.ParticleEngine.EmitterType.eSmokeyFireEmitter) <> 0 Then
                mlBaseVal = BurnFX.ParticleEngine.EmitterType.eSmokeyFireEmitter
                mlModVal = mlEmitterType - mlBaseVal
            End If
            mlBaseVal = mlEmitterType
            mlModVal = 0
        End Set
    End Property
End Structure