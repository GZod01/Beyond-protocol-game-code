Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

Public Class EpicaMesh
    Public lModelID As Int32
    Public oMesh As Mesh

    Public oShieldMesh As Mesh

    Public NumOfMaterials As Int32 = -1
    Public Materials() As Material
    Public Textures() As Texture
    Public sTextures() As String        'for getting back to the res mgr

    Public YMidPoint As Int32
    Public PlanetYAdjust As Int32

    Public EngineFXCnt As Int32
    Public EngineFXOffset() As Vector3
    Public EngineFXPCnt() As Int32

    Public ShieldXZRadius As Single

    Public RangeOffset As Int32         'a value added to all ranges...

    Public bLandBased As Boolean = False

    Public bTurretMesh As Boolean = False
    Public oTurretMesh As Mesh
    Public lTurretZOffset As Int32 

    'For Death Sequencing
    Public vecDeathSeqSize As Vector3
    Public vecHalfDeathSeqSize As Vector3       'to reduce a lot of calculations
    Public lDeathSeqExpCnt As Int32
    Public bDeathSeqFinale As Boolean
    Public sDeathSeqWav() As String
 
End Class