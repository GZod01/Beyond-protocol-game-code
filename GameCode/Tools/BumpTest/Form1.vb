Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

Public Class Form1
    Inherits System.Windows.Forms.Form

	Public Const gdPi As Single = 3.14159265358979
	Public Const gdHalfPie As Single = gdPi / 2.0F
	Public Const gdPieAndAHalf As Single = gdPi * 1.5F
	Public Const gdTwoPie As Single = gdPi * 2.0F
	Public Const gdDegreePerRad As Single = 180.0F / gdPi
	Friend WithEvents btnLoadMesh As System.Windows.Forms.Button
	Friend WithEvents chkGlowFX As System.Windows.Forms.CheckBox
	Friend WithEvents hscrGlowVal As System.Windows.Forms.HScrollBar
	Friend WithEvents Button1 As System.Windows.Forms.Button
	Friend WithEvents chkRenderCompare As System.Windows.Forms.CheckBox
	Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents btnPlayerRel As System.Windows.Forms.Button
    Friend WithEvents btnReloadTextures As System.Windows.Forms.Button
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Private Const gdRadPerDegree As Double = Math.PI / 180.0#

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents tmrRender As System.Windows.Forms.Timer
    Friend WithEvents opnDlg As System.Windows.Forms.OpenFileDialog
    Friend WithEvents savDlg As System.Windows.Forms.SaveFileDialog
    Friend WithEvents picMain As System.Windows.Forms.PictureBox
    Friend WithEvents cldDlg As System.Windows.Forms.ColorDialog
    Friend WithEvents hscrYaw As System.Windows.Forms.HScrollBar
    Friend WithEvents Label33 As System.Windows.Forms.Label
    Friend WithEvents Label34 As System.Windows.Forms.Label
    Friend WithEvents hscrRotate As System.Windows.Forms.HScrollBar
    Friend WithEvents btnSetDiffuseTexture As System.Windows.Forms.Button
    Friend WithEvents btnSetBumpTexture As System.Windows.Forms.Button
    Friend WithEvents btnSetIllumTexture As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents hscrSpecPwr As System.Windows.Forms.HScrollBar
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Me.tmrRender = New System.Windows.Forms.Timer(Me.components)
        Me.opnDlg = New System.Windows.Forms.OpenFileDialog
        Me.savDlg = New System.Windows.Forms.SaveFileDialog
        Me.picMain = New System.Windows.Forms.PictureBox
        Me.cldDlg = New System.Windows.Forms.ColorDialog
        Me.hscrYaw = New System.Windows.Forms.HScrollBar
        Me.hscrRotate = New System.Windows.Forms.HScrollBar
        Me.Label33 = New System.Windows.Forms.Label
        Me.Label34 = New System.Windows.Forms.Label
        Me.btnSetDiffuseTexture = New System.Windows.Forms.Button
        Me.btnSetBumpTexture = New System.Windows.Forms.Button
        Me.btnSetIllumTexture = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.hscrSpecPwr = New System.Windows.Forms.HScrollBar
        Me.btnLoadMesh = New System.Windows.Forms.Button
        Me.chkGlowFX = New System.Windows.Forms.CheckBox
        Me.hscrGlowVal = New System.Windows.Forms.HScrollBar
        Me.Button1 = New System.Windows.Forms.Button
        Me.chkRenderCompare = New System.Windows.Forms.CheckBox
        Me.Button2 = New System.Windows.Forms.Button
        Me.btnPlayerRel = New System.Windows.Forms.Button
        Me.btnReloadTextures = New System.Windows.Forms.Button
        Me.Button3 = New System.Windows.Forms.Button
        CType(Me.picMain, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'tmrRender
        '
        Me.tmrRender.Interval = 10
        '
        'savDlg
        '
        Me.savDlg.FileName = "doc1"
        '
        'picMain
        '
        Me.picMain.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.picMain.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.picMain.Location = New System.Drawing.Point(12, 67)
        Me.picMain.Name = "picMain"
        Me.picMain.Size = New System.Drawing.Size(200, 200)
        Me.picMain.TabIndex = 0
        Me.picMain.TabStop = False
        '
        'hscrYaw
        '
        Me.hscrYaw.Location = New System.Drawing.Point(274, 47)
        Me.hscrYaw.Maximum = 45
        Me.hscrYaw.Minimum = -45
        Me.hscrYaw.Name = "hscrYaw"
        Me.hscrYaw.Size = New System.Drawing.Size(80, 16)
        Me.hscrYaw.TabIndex = 9
        '
        'hscrRotate
        '
        Me.hscrRotate.Location = New System.Drawing.Point(447, 47)
        Me.hscrRotate.Maximum = 3599
        Me.hscrRotate.Name = "hscrRotate"
        Me.hscrRotate.Size = New System.Drawing.Size(80, 16)
        Me.hscrRotate.TabIndex = 10
        '
        'Label33
        '
        Me.Label33.Location = New System.Drawing.Point(236, 47)
        Me.Label33.Name = "Label33"
        Me.Label33.Size = New System.Drawing.Size(34, 14)
        Me.Label33.TabIndex = 62
        Me.Label33.Text = "YAW:"
        '
        'Label34
        '
        Me.Label34.Location = New System.Drawing.Point(389, 47)
        Me.Label34.Name = "Label34"
        Me.Label34.Size = New System.Drawing.Size(54, 14)
        Me.Label34.TabIndex = 63
        Me.Label34.Text = "ROTATE:"
        '
        'btnSetDiffuseTexture
        '
        Me.btnSetDiffuseTexture.Location = New System.Drawing.Point(114, 12)
        Me.btnSetDiffuseTexture.Name = "btnSetDiffuseTexture"
        Me.btnSetDiffuseTexture.Size = New System.Drawing.Size(110, 23)
        Me.btnSetDiffuseTexture.TabIndex = 2
        Me.btnSetDiffuseTexture.Text = "Set Diffuse Texture"
        Me.btnSetDiffuseTexture.UseVisualStyleBackColor = True
        '
        'btnSetBumpTexture
        '
        Me.btnSetBumpTexture.Location = New System.Drawing.Point(230, 12)
        Me.btnSetBumpTexture.Name = "btnSetBumpTexture"
        Me.btnSetBumpTexture.Size = New System.Drawing.Size(110, 23)
        Me.btnSetBumpTexture.TabIndex = 3
        Me.btnSetBumpTexture.Text = "Set Bump Texture"
        Me.btnSetBumpTexture.UseVisualStyleBackColor = True
        '
        'btnSetIllumTexture
        '
        Me.btnSetIllumTexture.Location = New System.Drawing.Point(346, 12)
        Me.btnSetIllumTexture.Name = "btnSetIllumTexture"
        Me.btnSetIllumTexture.Size = New System.Drawing.Size(110, 23)
        Me.btnSetIllumTexture.TabIndex = 4
        Me.btnSetIllumTexture.Text = "Set Illum Texture"
        Me.btnSetIllumTexture.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(12, 47)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(113, 17)
        Me.Label1.TabIndex = 76
        Me.Label1.Text = "Specular Power (30)"
        '
        'hscrSpecPwr
        '
        Me.hscrSpecPwr.Location = New System.Drawing.Point(128, 47)
        Me.hscrSpecPwr.Minimum = 1
        Me.hscrSpecPwr.Name = "hscrSpecPwr"
        Me.hscrSpecPwr.Size = New System.Drawing.Size(80, 16)
        Me.hscrSpecPwr.TabIndex = 8
        Me.hscrSpecPwr.Value = 30
        '
        'btnLoadMesh
        '
        Me.btnLoadMesh.Location = New System.Drawing.Point(12, 12)
        Me.btnLoadMesh.Name = "btnLoadMesh"
        Me.btnLoadMesh.Size = New System.Drawing.Size(96, 23)
        Me.btnLoadMesh.TabIndex = 1
        Me.btnLoadMesh.Text = "Load Mesh"
        '
        'chkGlowFX
        '
        Me.chkGlowFX.AutoSize = True
        Me.chkGlowFX.Location = New System.Drawing.Point(555, 47)
        Me.chkGlowFX.Name = "chkGlowFX"
        Me.chkGlowFX.Size = New System.Drawing.Size(90, 17)
        Me.chkGlowFX.TabIndex = 11
        Me.chkGlowFX.Text = "Glow Post FX"
        Me.chkGlowFX.UseVisualStyleBackColor = True
        '
        'hscrGlowVal
        '
        Me.hscrGlowVal.LargeChange = 1
        Me.hscrGlowVal.Location = New System.Drawing.Point(648, 47)
        Me.hscrGlowVal.Maximum = 10
        Me.hscrGlowVal.Name = "hscrGlowVal"
        Me.hscrGlowVal.Size = New System.Drawing.Size(80, 16)
        Me.hscrGlowVal.TabIndex = 12
        Me.hscrGlowVal.Value = 5
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(782, 12)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(96, 23)
        Me.Button1.TabIndex = 77
        Me.Button1.Text = "Load Compare"
        '
        'chkRenderCompare
        '
        Me.chkRenderCompare.AutoSize = True
        Me.chkRenderCompare.Location = New System.Drawing.Point(769, 46)
        Me.chkRenderCompare.Name = "chkRenderCompare"
        Me.chkRenderCompare.Size = New System.Drawing.Size(106, 17)
        Me.chkRenderCompare.TabIndex = 78
        Me.chkRenderCompare.Text = "Render Compare"
        Me.chkRenderCompare.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(680, 12)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(96, 23)
        Me.Button2.TabIndex = 79
        Me.Button2.Text = "Load Quad"
        '
        'btnPlayerRel
        '
        Me.btnPlayerRel.Location = New System.Drawing.Point(595, 12)
        Me.btnPlayerRel.Name = "btnPlayerRel"
        Me.btnPlayerRel.Size = New System.Drawing.Size(68, 23)
        Me.btnPlayerRel.TabIndex = 80
        Me.btnPlayerRel.Text = "PlayerRel"
        Me.btnPlayerRel.UseVisualStyleBackColor = True
        '
        'btnReloadTextures
        '
        Me.btnReloadTextures.Location = New System.Drawing.Point(462, 12)
        Me.btnReloadTextures.Name = "btnReloadTextures"
        Me.btnReloadTextures.Size = New System.Drawing.Size(110, 23)
        Me.btnReloadTextures.TabIndex = 5
        Me.btnReloadTextures.Text = "Set Spec Texture"
        Me.btnReloadTextures.UseVisualStyleBackColor = True
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(367, 219)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(75, 23)
        Me.Button3.TabIndex = 81
        Me.Button3.Text = "Button3"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(947, 679)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.btnPlayerRel)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.chkRenderCompare)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.hscrGlowVal)
        Me.Controls.Add(Me.chkGlowFX)
        Me.Controls.Add(Me.btnLoadMesh)
        Me.Controls.Add(Me.hscrSpecPwr)
        Me.Controls.Add(Me.btnReloadTextures)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btnSetIllumTexture)
        Me.Controls.Add(Me.btnSetBumpTexture)
        Me.Controls.Add(Me.btnSetDiffuseTexture)
        Me.Controls.Add(Me.Label34)
        Me.Controls.Add(Me.Label33)
        Me.Controls.Add(Me.hscrRotate)
        Me.Controls.Add(Me.hscrYaw)
        Me.Controls.Add(Me.picMain)
        Me.KeyPreview = True
        Me.Name = "Form1"
        Me.Text = "Form1"
        CType(Me.picMain, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region

    Private mlCameraX As Int32
    Private mlCameraY As Int32
    Private mlCameraZ As Int32
    Private mlCameraAtX As Int32
    Private mlCameraAtY As Int32
    Private mlCameraAtZ As Int32

    Private mbScrollLeft As Boolean
    Private mbScrollRight As Boolean
    Private mbScrollUp As Boolean
    Private mbScrollDown As Boolean
    Private mbRightDown As Boolean
    Private mbRightDrag As Boolean
    Private mlMouseX As Int32
    Private mlMouseY As Int32

    Private msCurrent As String

    Private moMesh As Mesh
    Private msMeshFile As String = ""
    Private matMesh As Material
    Private moDiffuse As Texture = Nothing
    Private msDiffuseTexture As String = ""
    Private moBump As Texture = Nothing
    Private msBumpTexture As String = ""
    Private moIllum As Texture = Nothing
    Private msIllumTexture As String = ""
    Private moSpec As Texture = Nothing
    Private msSpecTexture As String = ""

    Private moDevice As Device

    Private moShader As SimpleShader
    'Private moCloaker As CloakShader
    Private mfScaleVal As Single = 0.0F
    Private myScaleState As Byte = 0
    Private mfElapsedScale As Single = 0.013F
    Private moGlowFX As PostShader
    Private mbRenderGlowFX As Boolean = False

    Private matLightSrc As Material
    Private moLightMesh As Mesh

    Private mfTimeOfDay As Single = 0.0F

    Private Sub LoadMesh(ByVal sFile As String)
        Dim X As Int32
        Dim sTemp As String

        Dim sPostFix As String

        Dim sDetails As String = ""

        On Error Resume Next

        If sFile <> "" Then
            If Dir$(sFile) <> "" Then
                X = InStrRev(sFile, "\")
                msCurrent = Mid$(sFile, X + 1)

                'now set it up
                Dim mtrlBuffer() As ExtendedMaterial = Nothing
                moMesh = Mesh.FromFile(sFile, MeshFlags.Managed, moDevice, mtrlBuffer)

                'Now, because our engine uses Normals and lighting, let's get that out of the way
                If (moMesh.VertexFormat And VertexFormats.Normal) = 0 Then
                    Dim oTmpMesh As Mesh = moMesh.Clone(moMesh.Options.Value, moMesh.VertexFormat Or VertexFormats.Normal, moDevice)
                    oTmpMesh.ComputeNormals()
                    moMesh.Dispose()
                    moMesh = oTmpMesh
                    oTmpMesh = Nothing
                End If

                'DoTangentCalcs(moMesh)
                Dim elems(5) As VertexElement
                'position, normal, texture coords, tangent, binormal.

                elems(0) = New VertexElement(0, 0, DeclarationType.Float3, DeclarationMethod.Default, DeclarationUsage.Position, 0)
                elems(1) = New VertexElement(0, 12, DeclarationType.Float3, DeclarationMethod.Default, DeclarationUsage.Normal, 0)
                elems(2) = New VertexElement(0, 24, DeclarationType.Float2, DeclarationMethod.Default, DeclarationUsage.TextureCoordinate, 0)
                elems(3) = New VertexElement(0, 32, DeclarationType.Float3, DeclarationMethod.Default, DeclarationUsage.Tangent, 0)
                elems(4) = New VertexElement(0, 44, DeclarationType.Float3, DeclarationMethod.Default, DeclarationUsage.BiNormal, 0)
                elems(5) = VertexElement.VertexDeclarationEnd

                Dim clonedmesh As Mesh = moMesh.Clone(MeshFlags.Managed, elems, moDevice)
                moMesh.Dispose()
                moMesh = clonedmesh
                clonedmesh = Nothing

                moMesh = Geometry.ComputeTangentFrame(moMesh, DeclarationUsage.TextureCoordinate, 0, DeclarationUsage.Tangent, 0, DeclarationUsage.BiNormal, 0, _
                DeclarationUsage.Normal, 0, 0, Nothing, 0.0F, 0.0F, 0.0F, Nothing)

                sTemp = Mid$(sFile, InStrRev(sFile, "\") + 1)
                'txtFileName.Text = sTemp

                'sDetails = "File: " & sFile & vbCrLf
                'sDetails &= "Vertices: " & moMesh.NumberVertices & vbCrLf
                'sDetails &= "Faces: " & moMesh.NumberFaces & vbCrLf
                'sDetails &= "Material Count: " & mtrlBuffer.Length & vbCrLf & vbCrLf

                Dim lNumOfMaterials As Int32 = mtrlBuffer.Length
                'ReDim Materials(NumOfMaterials - 1)
                'ReDim Textures((NumOfMaterials * 4) - 1)
                'Now load our textures and materials
                For X = 0 To mtrlBuffer.Length - 1
                    matMesh = mtrlBuffer(X).Material3D
                    'Materials(X) = mtrlBuffer(X).Material3D
                    'Materials(X).Ambient = Materials(X).Diffuse
                    matMesh.Ambient = matMesh.Diffuse

                    'If mtrlBuffer(X).TextureFilename <> "" Then
                    '	sTemp = mtrlBuffer(X).TextureFilename

                    '	'ok, this is special now
                    '	'sTemp is the NEUTRAL version
                    '	sPostFix = Mid$(sTemp, Len(sTemp) - 3, 4)
                    '	sPostFix = ".bmp"
                    '	sTemp = Mid$(sTemp, 1, Len(sTemp) - 4)

                    '	'now, load our images
                    '	Textures(X) = TextureLoader.FromFile(moDevice, sTemp & sPostFix)
                    '	Textures((X * 4) + 1) = TextureLoader.FromFile(moDevice, sTemp & "_mine" & sPostFix)
                    '	Textures((X * 4) + 2) = TextureLoader.FromFile(moDevice, sTemp & "_ally" & sPostFix)
                    '	Textures((X * 4) + 3) = TextureLoader.FromFile(moDevice, sTemp & "_enemy" & sPostFix)
                    'End If
                Next X
            End If
        End If

    End Sub

    Public Function InitD3D(ByVal picBox As PictureBox) As Boolean
        Dim uParms As PresentParameters
        Dim uDispMode As DisplayMode
        Dim bWindowed As Boolean
        Dim bRes As Boolean


        Try
            uDispMode = Manager.Adapters.Default.CurrentDisplayMode
            uParms = New PresentParameters()
            With uParms
                bWindowed = True
                .Windowed = bWindowed
                .SwapEffect = SwapEffect.Discard
                .BackBufferCount = 1
                If bWindowed Then
                    .BackBufferFormat = uDispMode.Format
                    .BackBufferHeight = picBox.Height
                    .BackBufferWidth = picBox.Width
                Else
                    'TODO: Change our resolution to whatever the settings are
                    .BackBufferFormat = uDispMode.Format
                    .BackBufferHeight = uDispMode.Height
                    .BackBufferWidth = uDispMode.Width
                End If

                .PresentationInterval = PresentInterval.Immediate

                Dim bD32 As Boolean = Manager.CheckDepthStencilMatch(Manager.Adapters.Default.Adapter, DeviceType.Hardware, uDispMode.Format, uDispMode.Format, DepthFormat.D32, 0)
                Dim bD24X8 As Boolean = Manager.CheckDepthStencilMatch(Manager.Adapters.Default.Adapter, DeviceType.Hardware, uDispMode.Format, uDispMode.Format, DepthFormat.D24X8, 0)
                Dim bD24S8 As Boolean = Manager.CheckDepthStencilMatch(Manager.Adapters.Default.Adapter, DeviceType.Hardware, uDispMode.Format, uDispMode.Format, DepthFormat.D24S8, 0)
                Dim bD24X4S4 As Boolean = Manager.CheckDepthStencilMatch(Manager.Adapters.Default.Adapter, DeviceType.Hardware, uDispMode.Format, uDispMode.Format, DepthFormat.D24X4S4, 0)
                Dim bD16 As Boolean = Manager.CheckDepthStencilMatch(Manager.Adapters.Default.Adapter, DeviceType.Hardware, uDispMode.Format, uDispMode.Format, DepthFormat.D16, 0)
                Dim bD15S1 As Boolean = Manager.CheckDepthStencilMatch(Manager.Adapters.Default.Adapter, DeviceType.Hardware, uDispMode.Format, uDispMode.Format, DepthFormat.D15S1, 0)

                Dim lDpthFmt As Int32
                If bD32 = True Then
                    lDpthFmt = DepthFormat.D32
                ElseIf bD24X8 = True Then
                    lDpthFmt = DepthFormat.D24X8
                ElseIf bD24S8 = True Then
                    lDpthFmt = DepthFormat.D24S8
                ElseIf bD16 = True Then
                    lDpthFmt = DepthFormat.D16
                Else
                    MsgBox("Unable to determine Depth/Stencil Buffer format." & vbCrLf & "Please contact support for assistance!", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Error")
                    End
                End If

                .AutoDepthStencilFormat = CType(lDpthFmt, DepthFormat)
                .EnableAutoDepthStencil = True
            End With

            moDevice = New Device(0, DeviceType.Hardware, picBox.Handle, CreateFlags.HardwareVertexProcessing, uParms)
            bRes = Not moDevice Is Nothing
        Catch
            MsgBox(Err.Description, MsgBoxStyle.Critical Or MsgBoxStyle.OkOnly, "Error in initialize Direct3D")
            bRes = False
        Finally
            uParms = Nothing
            uDispMode = Nothing
        End Try

        Return bRes
    End Function

    Private Sub SetupMatrices()
        moDevice.Transform.View = Matrix.LookAtLH(New Vector3(mlCameraX, mlCameraY, mlCameraZ), _
          New Vector3(mlCameraAtX, mlCameraAtY, mlCameraAtZ), New Vector3(0.0#, 1.0#, 0.0#))    'up is always this
        moDevice.Transform.Projection = Matrix.PerspectiveFovLH(0.7853982F, CSng(moDevice.PresentationParameters.BackBufferWidth / moDevice.PresentationParameters.BackBufferHeight), 1, 2500000)
    End Sub

    Private Sub SetRenderStates()

        With moDevice.RenderState
            .Lighting = True
            .ZBufferEnable = True
            .DitherEnable = False
            .SpecularEnable = True
            .ShadeMode = ShadeMode.Gouraud
            .SourceBlend = Blend.SourceAlpha
            .DestinationBlend = Blend.InvSourceAlpha
            .AlphaBlendEnable = True
        End With

        moDevice.SetSamplerState(0, SamplerStageStates.MinFilter, TextureFilter.Linear)
        moDevice.SetSamplerState(0, SamplerStageStates.MagFilter, TextureFilter.Linear)
        moDevice.SetSamplerState(0, SamplerStageStates.MipFilter, TextureFilter.Linear)

    End Sub

    Private Sub ScrollCamera()
        'handle view scrolling... however, this comes with a twist... we need to determine our angle to the target
        ' deltaXx is the delta when scrolling camera's X along the X world axis
        ' deltaXz is the delta when scrolling camera's X along the Z world axis
        ' deltaZx is the delta when scrolling camera's Z along the X world axis
        ' deltaZz is the delta when scrolling camera's Z along the Z world axis
        ' which is to say, when changing X (scrolling horizontally), deltaXx is change to X and deltaXz is change to Z
        ' and when changing Z (scrolling vertically), deltaZx is change to X and deltaZz is change to Z

        Dim vecCameraLoc As Vector3 = New Vector3(mlCameraX, mlCameraY, mlCameraZ)
        Dim vecCameraAt As Vector3 = New Vector3(mlCameraAtX, mlCameraAtY, mlCameraAtZ)
        Dim vecTemp As Vector3 = Vector3.Subtract(vecCameraLoc, vecCameraAt)

        Dim deltaXx As Single
        Dim deltaXz As Single
        Dim deltaZx As Single
        Dim deltaZz As Single

        vecTemp.Normalize()
        Dim vecDot As Vector3 = Vector3.Cross(vecTemp, New Vector3(0, 1, 0))

        Dim mlScrollRate As Int32 = 10 '0

        deltaXx = vecDot.X * mlScrollRate
        deltaXz = vecDot.Z * mlScrollRate
        deltaZx = vecTemp.X * mlScrollRate
        deltaZz = vecTemp.Z * mlScrollRate

        If mbScrollLeft Then
            mlCameraX -= CInt(deltaXx)
            mlCameraAtX -= CInt(deltaXx)
            mlCameraZ -= CInt(deltaXz)
            mlCameraAtZ -= CInt(deltaXz)
        ElseIf mbScrollRight Then
            mlCameraX += CInt(deltaXx)
            mlCameraAtX += CInt(deltaXx)
            mlCameraZ += CInt(deltaXz)
            mlCameraAtZ += CInt(deltaXz)
        End If
        If mbScrollUp Then
            mlCameraZ -= CInt(deltaZz)
            mlCameraAtZ -= CInt(deltaZz)
            mlCameraX -= CInt(deltaZx)
            mlCameraAtX -= CInt(deltaZx)
        ElseIf mbScrollDown Then
            mlCameraZ += CInt(deltaZz)
            mlCameraAtZ += CInt(deltaZz)
            mlCameraX += CInt(deltaZx)
            mlCameraAtX += CInt(deltaZx)
        End If

        vecTemp = Nothing
        vecCameraAt = Nothing
        vecCameraLoc = Nothing
        vecDot = Nothing
    End Sub

    Public Sub RotateCamera(ByVal deltaX As Int32, ByVal deltaY As Int32)
        Dim fXYaw As Single
        Dim fZYaw As Single

        'X is easy
        RotatePoint(mlCameraAtX, mlCameraAtZ, mlCameraX, mlCameraZ, deltaX)

        'Ok, Y is the real pain...
        Dim vecAngle As Vector3 = New Vector3(mlCameraAtX - mlCameraX, mlCameraAtY - mlCameraY, mlCameraAtZ - mlCameraZ)
        vecAngle.Normalize()

        'deltaY = 0

        fXYaw = vecAngle.X * deltaY
        fZYaw = vecAngle.Z * deltaY

        RotatePoint(mlCameraAtX, mlCameraAtY, mlCameraX, mlCameraY, fXYaw)
        RotatePoint(mlCameraAtZ, mlCameraAtY, mlCameraZ, mlCameraY, fZYaw)
    End Sub

    Private Sub tmrRender_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles tmrRender.Tick
        DrawScene(False)
    End Sub

    Private Sub Form1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Select Case e.KeyCode
            Case Keys.W
                moDevice.RenderState.FillMode = FillMode.WireFrame
            Case Keys.S
                moDevice.RenderState.FillMode = FillMode.Solid
            Case Keys.Up
                mbScrollUp = True
            Case Keys.Down
                mbScrollDown = True
            Case Keys.Left
                mbScrollLeft = True
            Case Keys.Right
                mbScrollRight = True
            Case Keys.I
                mfScaleVal += mfElapsedScale
            Case Keys.K
                mfScaleVal -= mfElapsedScale
        End Select
    End Sub

    Private Sub Form1_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles picMain.MouseDown
        If e.Button = Windows.Forms.MouseButtons.Right Then
            mbRightDown = True
        End If
    End Sub

    Private Sub Form1_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles picMain.MouseUp
        mbRightDown = False
        mbRightDrag = False
    End Sub

    Private Sub Form1_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles picMain.MouseMove
        Dim deltaX As Int32
        Dim deltaY As Int32

        If mbRightDown Then
            mbRightDrag = True
            deltaX = e.X - mlMouseX
            deltaY = e.Y - mlMouseY
            RotateCamera(deltaX, deltaY)
        Else
            'mbScrollLeft = False : mbScrollRight = False : mbScrollUp = False : mbScrollDown = False
            'If e.X < 20 Then
            '	mbScrollLeft = True
            'ElseIf e.X > picMain.Width - 20 Then
            '	mbScrollRight = True
            'End If
            'If e.Y < 20 Then
            '	mbScrollUp = True
            'ElseIf e.Y > picMain.Height - 20 Then
            '	mbScrollDown = True
            'End If
        End If
        mlMouseX = e.X
        mlMouseY = e.Y
    End Sub

    Protected Overrides Sub OnMouseWheel(ByVal e As System.Windows.Forms.MouseEventArgs)
        If e.Delta < 0 Then MouseWheelDown() Else MouseWheelUp()
    End Sub

    Private Sub MouseWheelDown()
        ModifyZoom(50)
    End Sub

    Private Sub MouseWheelUp()
        ModifyZoom(-50)
    End Sub

    Private Sub ModifyZoom(ByVal lAmt As Int32)
        Dim oVec3 As Vector3 = New Vector3(mlCameraX - mlCameraAtX, mlCameraY - mlCameraAtY, mlCameraZ - mlCameraAtZ)
        oVec3.Normalize()
        mlCameraX += CInt(lAmt * oVec3.X)
        mlCameraY += CInt(lAmt * oVec3.Y)
        mlCameraZ += CInt(lAmt * oVec3.Z)
        oVec3 = Nothing
    End Sub

    'Private Sub cmdDiffuse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '	Dim oRes As DialogResult
    '	cldDlg.Color = matMesh.Diffuse
    '	oRes = cldDlg.ShowDialog(Me)
    '	If oRes = Windows.Forms.DialogResult.OK Then
    '		matMesh.Diffuse = cldDlg.Color
    '		cmdDiffuse.BackColor = cldDlg.Color
    '	End If
    'End Sub

    'Private Sub cmdEmissive_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '	Dim oRes As DialogResult
    '	cldDlg.Color = matMesh.Emissive
    '	oRes = cldDlg.ShowDialog(Me)
    '	If oRes = Windows.Forms.DialogResult.OK Then
    '		matMesh.Emissive = cldDlg.Color
    '		cmdEmissive.BackColor = cldDlg.Color
    '	End If
    'End Sub

    'Private Sub cmdSpecular_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '	Dim oRes As DialogResult
    '	cldDlg.Color = matMesh.Specular
    '	oRes = cldDlg.ShowDialog(Me)
    '	If oRes = Windows.Forms.DialogResult.OK Then
    '		matMesh.Specular = cldDlg.Color
    '		cmdSpecular.BackColor = cldDlg.Color
    '	End If
    Private mbPlayedSound As Boolean = False

    'End Sub
    Public Sub WriteRSToFile()
        Dim oFS As New IO.FileStream("C:\" & Now.ToString("MMddyyyyHHmmss") & ".txt", IO.FileMode.Create)
        Dim oWrite As New IO.StreamWriter(oFS)

        oWrite.Write(moDevice.SamplerState(0).ToString.Replace(vbLf, vbCrLf))
        oWrite.Write(moDevice.SamplerState(1).ToString.Replace(vbLf, vbCrLf))
        oWrite.Write(moDevice.SamplerState(2).ToString.Replace(vbLf, vbCrLf))
        oWrite.Write(moDevice.SamplerState(3).ToString.Replace(vbLf, vbCrLf))
        oWrite.WriteLine()
        oWrite.Write(moDevice.RenderState.ToString.Replace(vbLf, vbCrLf))
        oWrite.WriteLine()
        oWrite.WriteLine("DepthStencilFormat: " & moDevice.PresentationParameters.AutoDepthStencilFormat.ToString)
        oWrite.WriteLine()
        oWrite.WriteLine("Depth Size: " & moDevice.DepthStencilSurface.Description.Width & "x" & moDevice.DepthStencilSurface.Description.Height & "... MST: " & moDevice.DepthStencilSurface.Description.MultiSampleType.ToString & "... Quality: " & moDevice.DepthStencilSurface.Description.MultiSampleQuality)

        oWrite.Write(moDevice.Lights(0).ToString.Replace(vbLf, vbCrLf))
        oWrite.Write(moDevice.TextureState.TextureState(0).ToString.Replace(vbLf, vbCrLf))
        oWrite.Write(moDevice.TextureState.TextureState(1).ToString.Replace(vbLf, vbCrLf))
        oWrite.Write(moDevice.TextureState.TextureState(2).ToString.Replace(vbLf, vbCrLf))
        oWrite.Write(moDevice.TextureState.TextureState(3).ToString.Replace(vbLf, vbCrLf))
        oWrite.Write(moDevice.VertexDeclaration.ToString.Replace(vbLf, vbCrLf))
        oWrite.Write(moDevice.VertexFormat.ToString.Replace(vbLf, vbCrLf))

        oWrite.Close()
        oFS.Close()
        oWrite = Nothing
        oFS = Nothing
    End Sub

    Private mmtWorlds() As Matrix
    Private mlItemID() As Int32
    Private mlItemUB As Int32 = -1
    Private Function CreateMatrix(ByVal lX As Int32, ByVal lY As Int32, ByVal lZ As Int32, ByVal lYaw As Int32, ByVal lPitch As Int32, ByVal lRoll As Int32) As Matrix

        Dim fYaw As Single = lYaw * 0.0174532924F
        Dim fPitch As Single = lPitch * 0.0174532924F
        Dim fRoll As Single = lRoll * 0.0174532924F


        Dim matWorld As Matrix = Matrix.Identity
        matWorld.RotateYawPitchRoll(fYaw, fPitch, fRoll)
        Dim matTemp As Matrix = Matrix.Identity
        matTemp.Translate(lX, lY, lZ)
        matWorld.Multiply(matTemp)

        Return matWorld
    End Function

    Private moSW As Stopwatch
    Private Sub DrawScene(ByVal bAllWhite As Boolean)
        Dim matWorld As Matrix

        ''mlCameraAtX = -4905599 : mlCameraAtY = 0 : mlCameraAtZ = 2108452
        'If moOtherMesh Is Nothing Then
        '    ReDim moOtherMesh(3)
        '    LoadMesh("C:\Documents and Settings\Matthew Campbell\Desktop\TempCode\Packager\bin\FromCharles\Claw.x")
        '    moOtherMesh(0) = moMesh
        '    LoadMesh("C:\Documents and Settings\Matthew Campbell\Desktop\TempCode\Packager\bin\New Models\Models\ftr\Wasp.x")
        '    moOtherMesh(1) = moMesh
        '    LoadMesh("C:\Documents and Settings\Matthew Campbell\Desktop\TempCode\Packager\bin\FromCharles\Jackal.x")
        '    moOtherMesh(2) = moMesh
        '    LoadMesh("C:\Documents and Settings\Matthew Campbell\Desktop\TempCode\Packager\bin\FromCharles\Hawk.x")
        '    moOtherMesh(3) = moMesh

        '    '2 elephants
        '    '21 mantis
        '    '8 starblaster
        '    '8 medium 2
        '    mlItemUB = 16 '38
        '    ReDim mmtWorlds(mlItemUB)
        '    ReDim mlItemID(mlItemUB)

        '    'yaw = rotate in Y
        '    'pitch = rotate along X
        '    'roll = rotate along z
        '    mlItemID(0) = 0 : mmtWorlds(0) = CreateMatrix(0, -200, 0, 60, 0, -7)
        '    mlItemID(1) = 0 : mmtWorlds(1) = CreateMatrix(1000, -1500, 4000, 60, 0, 0)
        '    mlItemID(2) = 3 : mmtWorlds(2) = CreateMatrix(-500, 500, 500, 55, 10, -45)
        '    mlItemID(3) = 3 : mmtWorlds(3) = CreateMatrix(-700, 100, 200, 51, 14, -35)

        '    'mlItemID(4) = 2 : mmtWorlds(4) = CreateMatrix(1400, -200, 300, 60, 0, 0)
        '    mlItemID(4) = 2 : mmtWorlds(4) = CreateMatrix(1400, 20000, 300, 60, 0, 0)
        '    mlItemID(5) = 2 : mmtWorlds(5) = CreateMatrix(1100, -600, 200, 60, 0, 0)
        '    'mlItemID(5) = 2 : mmtWorlds(5) = CreateMatrix(1000, -600, 200, 60, 0, 0)
        '    'mlItemID(5) = 2 : mmtWorlds(5) = CreateMatrix(1000, 20000, 200, 60, 0, 0)

        '    'mlItemID(6) = 2 : mmtWorlds(6) = CreateMatrix(2400, -1700, 4300, 60, 0, 0)
        '    mlItemID(6) = 2 : mmtWorlds(6) = CreateMatrix(2400, 20000, 4300, 60, 0, 0)
        '    mlItemID(7) = 2 : mmtWorlds(7) = CreateMatrix(2000, -2100, 4200, 60, 0, 0)
        '    'mlItemID(7) = 2 : mmtWorlds(7) = CreateMatrix(2000, 20100, 4200, 60, 0, 0)

        '    mlItemID(8) = 1 : mmtWorlds(8) = CreateMatrix(0, 0, -900, 58, 0, 30)
        '    mlItemID(9) = 1 : mmtWorlds(9) = CreateMatrix(150, 50, -800, 58, 0, 5)
        '    mlItemID(10) = 1 : mmtWorlds(10) = CreateMatrix(230, 30, -930, 58, 0, -10)

        '    mlItemID(11) = 1 : mmtWorlds(11) = CreateMatrix(0, -250, 1000, 58, 0, -24)
        '    mlItemID(12) = 1 : mmtWorlds(12) = CreateMatrix(150, -300, 900, 58, 0, -5)
        '    mlItemID(13) = 1 : mmtWorlds(13) = CreateMatrix(230, -280, 1030, 58, 0, -10)

        '    mlItemID(14) = 1 : mmtWorlds(14) = CreateMatrix(1200, 100, 500, 58, 0, -10)
        '    mlItemID(15) = 1 : mmtWorlds(15) = CreateMatrix(1050, 50, 400, 57, 0, 10)
        '    mlItemID(16) = 1 : mmtWorlds(16) = CreateMatrix(970, 210, 530, 58, 0, 7)

        '    mlCameraX = -126 : mlCameraY = 1478 : mlCameraZ = -1709
        'End If

        moDevice.BeginScene()
        'moDevice.Clear(ClearFlags.Target Or ClearFlags.ZBuffer, System.Drawing.Color.FromArgb(255, 0, 64, 0), 1, 0)
        moDevice.Clear(ClearFlags.Target Or ClearFlags.ZBuffer, System.Drawing.Color.FromArgb(255, 0, 0, 0), 1, 0)
        ScrollCamera()
        SetRenderStates()
        moDevice.RenderState.Ambient = System.Drawing.Color.DarkGray

        Dim fElapsed As Single
        If moSW Is Nothing Then
            moSW = Stopwatch.StartNew
            fElapsed = 1.0F
        Else
            moSW.Stop()
            fElapsed = moSW.ElapsedMilliseconds / 30.0F
            moSW.Reset()
            moSW.Start()
        End If
        'fElapsed = 1.0F
        mfTimeOfDay += 0.0025F * fElapsed
        If mfTimeOfDay > 1.0F Then mfTimeOfDay -= 1.0F

        With moDevice.Lights(0)
            .Diffuse = System.Drawing.Color.White
            .Ambient = System.Drawing.Color.DarkGray
            .Type = LightType.Directional
            '.Type = LightType.Point

            Dim fX As Single = 1.0F
            Dim fY As Single = 0.0F
            RotatePoint(0, 0, fX, fY, mfTimeOfDay * 360.0F)

            Me.Text = mfTimeOfDay.ToString("0.000") & ", " & fX.ToString("0.000") & ", " & fY.ToString("0.000")

            .Direction = Vector3.Multiply(New Vector3(fX, fY, 0.0F), 5)
            '.Position = New Vector3(0, 0, 0)

            'New Vector3( 0, -500, -1)
            .Range = 10000000
            .Specular = System.Drawing.Color.White
            .Attenuation0 = 1
            .Attenuation1 = 0
            .Attenuation2 = 0
            .Falloff = 0.3

            .Enabled = True
            .Update()
        End With

        If moLightMesh Is Nothing Then
            moLightMesh = Mesh.Sphere(moDevice, 100, 32, 32)
            With matLightSrc
                .Ambient = Color.White
                .Diffuse = Color.White
                .Emissive = Color.White
                .Specular = Color.White
            End With
        End If

        Dim vecLightVec As Vector3 = moDevice.Lights(0).Direction
        vecLightVec.Scale(-250.0F)
        moDevice.Transform.World = Matrix.Translation(vecLightVec)
        moDevice.Material = matLightSrc
        moDevice.SetTexture(0, Nothing)
        'moLightMesh.DrawSubset(0)
        moDevice.Transform.World = Matrix.Identity

        matWorld = Matrix.Identity
        matWorld.Multiply(Matrix.RotationZ(CSng(hscrYaw.Value * gdRadPerDegree)))
        matWorld.Multiply(Matrix.RotationY(CSng((hscrRotate.Value / 10.0F) * gdRadPerDegree)))
        ' matWorld.Multiply(Matrix.Translation(mlCameraAtX, mlCameraAtY, mlCameraAtZ))
        moDevice.Transform.World = matWorld
        moDevice.RenderState.CullMode = Cull.None
        'Me.Text = mlCameraAtX & ", " & mlCameraAtY & ", " & mlCameraAtZ

        'do rendering here
        If moMesh Is Nothing = False OrElse mlItemUB > -1 Then
            'If moCloaker Is Nothing Then
            '    moCloaker = New CloakShader(moDevice)
            'End If
            If moShader Is Nothing Then
                moShader = New SimpleShader(moDevice)
            End If

            'If myScaleState = 0 Then
            '    mfScaleVal += (fElapsed * mfElapsedScale)
            '    If mfScaleVal > 1.0F Then
            '        mfScaleVal = 1.0F
            '        myScaleState = 1
            '    End If
            '    If mbPlayedSound = False Then
            '        mbPlayedSound = True

            '        Dim oSound As New System.Media.SoundPlayer()
            '        oSound.SoundLocation = "C:\Documents and Settings\Matthew Campbell\Desktop\AudioWork\Cloak\InitiateCloak.wav"
            '        oSound.Load()
            '        oSound.Play()
            '    End If

            '    moCloaker.PrepareToRender()
            '    moCloaker.RenderMesh(moMesh, mfScaleVal, moDiffuse, moIllum)
            '    moCloaker.EndRender()
            'ElseIf myScaleState = 128 Then
            '    mfScaleVal -= (fElapsed * mfElapsedScale)
            '    If mfScaleVal < 0 Then
            '        mfScaleVal = 0.0F
            '        myScaleState = 129
            '    End If
            '    If mbPlayedSound = False Then
            '        mbPlayedSound = True
            '        Dim oSound As New System.Media.SoundPlayer()
            '        oSound.SoundLocation = "C:\Documents and Settings\Matthew Campbell\Desktop\AudioWork\Cloak\DeactCloak.wav"
            '        oSound.Load()
            '        oSound.Play()

            '    End If

            '    moCloaker.PrepareToRender()
            '    moCloaker.RenderMesh(moMesh, mfScaleVal, moDiffuse, moIllum)
            '    moCloaker.EndRender()
            'Else
            '    mbPlayedSound = False
            '    Dim lNew As Int32 = myScaleState
            '    lNew += 1
            '    If lNew > 255 Then lNew = 0
            '    myScaleState = CByte(lNew)

            '    If myScaleState > 129 Then

            If bAllWhite = True Then

                moDevice.SetTexture(0, Nothing)
                Dim oMat As Material
                With oMat
                    .Ambient = Color.White
                    .Diffuse = Color.White
                    .Emissive = Color.White
                    .Specular = Color.White
                End With
                moDevice.Material = oMat

                moMesh.DrawSubset(0)
            Else
                moShader.PrepareToRender()
                moShader.RenderMesh(moMesh, mlCameraX, mlCameraY, mlCameraZ, moDevice.Lights(0).Direction, btnPlayerRel.BackColor)
                moShader.EndRender()
            End If

            
            '    End If
            'End If

            'moShader.PrepareToRender()
            'For X As Int32 = 0 To mlItemUB
            '    moDevice.Transform.World = mmtWorlds(X)
            '    moShader.RenderMesh(moOtherMesh(mlItemID(X)), mlCameraX, mlCameraY, mlCameraY, moDevice.Lights(0).Direction, btnPlayerRel.BackColor)
            'Next
            'moShader.EndRender()

        End If

        If chkRenderCompare.Checked = True Then
            If moCompareMesh Is Nothing = False Then
                matWorld = Matrix.Identity
                matWorld.Multiply(Matrix.Translation(0, 0, 0))
                moDevice.Transform.World = matWorld

                moDevice.SetTexture(0, Nothing)
                Dim oMat As Material
                With oMat
                    .Ambient = Color.White
                    .Diffuse = Color.White
                    .Emissive = Color.White
                    .Specular = Color.White
                End With
                moDevice.Material = oMat
                moCompareMesh.DrawSubset(0)

            End If
        End If

        If mbRenderGlowFX = True Then
            If moGlowFX Is Nothing Then moGlowFX = New PostShader(moDevice)
            moGlowFX.ExecutePostProcess()
        End If

        SetupMatrices()

        moDevice.EndScene()
        moDevice.Present()
    End Sub

    Private moOtherMesh() As Mesh


    Private Sub RefreshStatDisplay()
        ' Get the original mesh's vertex buffer.
        Dim ranks(0) As Integer
        ranks(0) = moMesh.NumberVertices
        Dim arr As System.Array = moMesh.VertexBuffer.Lock(0, (New CustomVertex.PositionNormalTextured()).GetType(), LockFlags.None, ranks)

        Dim i As Integer

        Dim lMinX As Int32 = Int32.MaxValue
        Dim lMinY As Int32 = Int32.MaxValue
        Dim lMinZ As Int32 = Int32.MaxValue
        Dim lMaxX As Int32 = Int32.MinValue
        Dim lMaxY As Int32 = Int32.MinValue
        Dim lMaxZ As Int32 = Int32.MinValue

        For i = 0 To arr.Length - 1

            Dim pn As Direct3D.CustomVertex.PositionNormalTextured = CType(arr.GetValue(i), CustomVertex.PositionNormalTextured)

            If pn.X < lMinX Then lMinX = CInt(pn.X)
            If pn.X > lMaxX Then lMaxX = CInt(pn.X)
            If pn.Y < lMinY Then lMinY = CInt(pn.Y)
            If pn.Y > lMaxY Then lMaxY = CInt(pn.Y)
            If pn.Z < lMinZ Then lMinZ = CInt(pn.Z)
            If pn.Z > lMaxZ Then lMaxZ = CInt(pn.Z)
        Next i

        moMesh.VertexBuffer.Unlock()
        arr = Nothing

        'cmdDiffuse.BackColor = matMesh.Diffuse
        'cmdSpecular.BackColor = matMesh.Specular
        'cmdEmissive.BackColor = matMesh.Emissive

    End Sub

    Public Sub RotatePoint(ByVal lAxisX As Int32, ByVal lAxisY As Int32, ByRef lEndX As Int32, ByRef lEndY As Int32, ByVal dDegree As Double)
        Dim dDX As Double
        Dim dDY As Double
        Dim dRads As Double

        dRads = dDegree * (Math.PI / 180)
        dDX = lEndX - lAxisX
        dDY = lEndY - lAxisY
        lEndX = lAxisX + CInt((dDX * Math.Cos(dRads)) + (dDY * Math.Sin(dRads)))
        lEndY = lAxisY + -CInt((dDX * Math.Sin(dRads)) - (dDY * Math.Cos(dRads)))
    End Sub

    Public Sub RotatePoint(ByVal lAxisX As Int32, ByVal lAxisY As Int32, ByRef lEndX As Single, ByRef lEndY As Single, ByVal dDegree As Double)
        Dim dDX As Double
        Dim dDY As Double
        Dim dRads As Double

        dRads = dDegree * (Math.PI / 180)
        dDX = lEndX - lAxisX
        dDY = lEndY - lAxisY
        lEndX = lAxisX + CSng((dDX * Math.Cos(dRads)) + (dDY * Math.Sin(dRads)))
        lEndY = lAxisY + -CSng((dDX * Math.Sin(dRads)) - (dDY * Math.Cos(dRads)))
    End Sub

    Private Sub hscrRotate_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles hscrRotate.ValueChanged
        Me.Text = "YAW: " & hscrYaw.Value & ", ROTATE: " & hscrRotate.Value
    End Sub

    Private Sub hscrYaw_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles hscrYaw.ValueChanged
        Me.Text = "YAW: " & hscrYaw.Value & ", ROTATE: " & hscrRotate.Value
    End Sub

    Public Function LineAngleDegrees(ByVal lX1 As Int32, ByVal lY1 As Int32, ByVal lX2 As Int32, ByVal lY2 As Int32) As Single
        Dim dDeltaX As Single
        Dim dDeltaY As Single
        Dim dAngle As Single

        dDeltaX = lX2 - lX1
        dDeltaY = lY2 - lY1

        If dDeltaX = 0 Then     'vertical
            If dDeltaY < 0 Then
                dAngle = gdHalfPie
            Else
                dAngle = gdPieAndAHalf
            End If
        ElseIf dDeltaY = 0 Then     'horizontal
            If dDeltaX < 0 Then
                dAngle = gdPi
            Else
                dAngle = 0
            End If
        Else    'angled
            dAngle = CSng(System.Math.Atan(System.Math.Abs(dDeltaY / dDeltaX)))

            'If dDeltaX > -1 AndAlso dDeltaY < 0 Then
            '    dAngle = gdTwoPie - dAngle
            'ElseIf dDeltaX < 0 AndAlso dDeltaY < 0 Then
            '    dAngle = gdPi + dAngle
            'ElseIf dDeltaX < 0 AndAlso dDeltaY > -1 Then
            '    dAngle = gdPi - dAngle
            'End If


            'Correct for VB's reversed Y... VB Upper Right is ok... but the other quads are not
            If dDeltaX > -1 And dDeltaY > -1 Then       'VB Lower Right
                dAngle = gdTwoPie - dAngle
            ElseIf dDeltaX < 0 And dDeltaY > -1 Then    'VB Lower Left
                dAngle = gdPi + dAngle
            ElseIf dDeltaX < 0 And dDeltaY < 0 Then     'VB Upper Left
                dAngle = gdPi - dAngle
            End If
        End If

        'Not sure this is suppose to be CINT
        'Return CInt(dAngle * gdDegreePerRad)
        Return (dAngle * gdDegreePerRad)

    End Function

    Private Sub btnLoadMesh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLoadMesh.Click
        Dim oRes As DialogResult

        opnDlg.Filter = "X Files *.x|*.x"
        opnDlg.FileName = msMeshFile
        oRes = opnDlg.ShowDialog(Me)

        If oRes = Windows.Forms.DialogResult.OK Then
            msMeshFile = opnDlg.FileName

            LoadMesh(msMeshFile)

            'mlCameraAtX = 0 : mlCameraAtY = 0 : mlCameraAtZ = 0
            'mlCameraX = 0 : mlCameraY = 1000 : mlCameraZ = -1000

            'Now, load our extended details...
            'LoadExtendedDetails(sFile)

            RefreshStatDisplay()
        End If
    End Sub

    'Private Sub btnReloadTextures_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReloadTextures.Click
    '    If moShader Is Nothing = False Then moShader.ClearTextures()
    '    If moDiffuse Is Nothing = False Then moDiffuse.Dispose()
    '    moDiffuse = Nothing
    '    If moBump Is Nothing = False Then moBump.Dispose()
    '    moBump = Nothing
    '    If moIllum Is Nothing = False Then moIllum.Dispose()
    '    moIllum = Nothing
    '    If moSpec Is Nothing = False Then moSpec.Dispose()
    '    moSpec = Nothing

    '    If msDiffuseTexture <> "" Then moDiffuse = TextureLoader.FromFile(moDevice, msDiffuseTexture)
    '    If msBumpTexture <> "" Then moBump = TextureLoader.FromFile(moDevice, msBumpTexture)
    '    If msIllumTexture <> "" Then moIllum = TextureLoader.FromFile(moDevice, msIllumTexture)
    '    If msSpecTexture <> "" Then moSpec = TextureLoader.FromFile(moDevice, msSpecTexture)

    '    If moShader Is Nothing = False Then
    '        moShader.SetDiffuseTexture(moDiffuse)
    '        moShader.SetBumpTexture(moBump)
    '        moShader.SetIllumTexture(moIllum)
    '        moShader.SetSpecTexture(moSpec)
    '    End If
    'End Sub

    Private Sub btnSetIllumTexture_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSetIllumTexture.Click
        Dim oRes As DialogResult

        opnDlg.Filter = "Image Files *.*|*.*"
        opnDlg.FileName = msIllumTexture
        oRes = opnDlg.ShowDialog(Me)

        If oRes = Windows.Forms.DialogResult.OK Then
            msIllumTexture = opnDlg.FileName

            If moIllum Is Nothing = False Then moIllum.Dispose()
            moIllum = Nothing
            moIllum = TextureLoader.FromFile(moDevice, msIllumTexture)

            If moShader Is Nothing = False Then moShader.SetIllumTexture(moIllum)
        End If
    End Sub

    Private Sub btnSetBumpTexture_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSetBumpTexture.Click
        Dim oRes As DialogResult

        opnDlg.Filter = "Image Files *.*|*.*"
        opnDlg.FileName = msBumpTexture
        oRes = opnDlg.ShowDialog(Me)

        If oRes = Windows.Forms.DialogResult.OK Then
            msBumpTexture = opnDlg.FileName

            If moBump Is Nothing = False Then moBump.Dispose()
            moBump = Nothing
            moBump = TextureLoader.FromFile(moDevice, msBumpTexture)

            If moShader Is Nothing = False Then moShader.SetBumpTexture(moBump)
        End If
    End Sub

    Private Sub btnSetDiffuseTexture_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSetDiffuseTexture.Click
        Dim oRes As DialogResult

        opnDlg.Filter = "Image Files *.*|*.*"
        opnDlg.FileName = msDiffuseTexture
        oRes = opnDlg.ShowDialog(Me)

        If oRes = Windows.Forms.DialogResult.OK Then
            msDiffuseTexture = opnDlg.FileName

            If moDiffuse Is Nothing = False Then moDiffuse.Dispose()
            moDiffuse = Nothing
            moDiffuse = TextureLoader.FromFile(moDevice, msDiffuseTexture)

            If moShader Is Nothing = False Then moShader.SetDiffuseTexture(moDiffuse)
        End If
    End Sub

    Private Sub hscrSpecPwr_Scroll(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ScrollEventArgs) Handles hscrSpecPwr.Scroll
        'If moCloaker Is Nothing = False Then
        moShader.SetSpecPower(hscrSpecPwr.Value)
        Label1.Text = "Specular Power (" & hscrSpecPwr.Value & ")"
        'End If
    End Sub

    Private Sub chkGlowFX_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkGlowFX.CheckedChanged
        mbRenderGlowFX = chkGlowFX.Checked
    End Sub

    Private Sub hscrGlowVal_Scroll(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ScrollEventArgs) Handles hscrGlowVal.Scroll
        If moGlowFX Is Nothing = False Then
            moGlowFX.SetGlowValue(hscrGlowVal.Value)
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim oRes As DialogResult

        opnDlg.Filter = "X Files *.x|*.x"
        opnDlg.FileName = msMeshFile
        oRes = opnDlg.ShowDialog(Me)

        If oRes = Windows.Forms.DialogResult.OK Then
            Dim sFile As String = opnDlg.FileName
            LoadCompareMesh(sFile)
            chkRenderCompare.Checked = True
        End If
    End Sub

    Private moCompareMesh As Mesh
    Private Sub LoadCompareMesh(ByVal sFile As String)
        On Error Resume Next

        If sFile <> "" Then
            If Dir$(sFile) <> "" Then

                'now set it up
                Dim mtrlBuffer() As ExtendedMaterial = Nothing
                moCompareMesh = Mesh.FromFile(sFile, MeshFlags.Managed, moDevice, mtrlBuffer)



            End If
        End If

    End Sub


    Public Function CreateTexturedBox(ByVal fWidth As Single, ByVal fHeight As Single, ByVal fDepth As Single) As Mesh

        Device.IsUsingEventHandlers = False

        Dim oTemp As Mesh = Mesh.Box(moDevice, fWidth, fHeight, fDepth)
        oTemp.ComputeNormals()

        Dim TexturedObject As Mesh = New Mesh(oTemp.NumberFaces, oTemp.NumberVertices, MeshFlags.Managed, CustomVertex.PositionNormalTextured.Format, moDevice)

        ' Get the original mesh's vertex buffer.
        Dim ranks(0) As Integer
        ranks(0) = oTemp.NumberVertices
        Dim arr As System.Array = oTemp.VertexBuffer.Lock(0, (New CustomVertex.PositionNormal()).GetType(), LockFlags.None, ranks)

        ' Set the vertex buffer
        Dim data As System.Array = TexturedObject.VertexBuffer.Lock(0, (New CustomVertex.PositionNormalTextured()).GetType(), LockFlags.None, ranks)

        'Dim u As Single
        Dim i As Integer

        Dim lFacing As Int32
        Dim bTopOrBottom As Boolean
        Dim fHalfHeight As Single = fHeight / 2

        For i = 0 To arr.Length - 1

            lFacing = CInt(Math.Floor(i / 4))
            bTopOrBottom = (lFacing = 1) OrElse (lFacing = 3)

            Dim pn As Direct3D.CustomVertex.PositionNormal = CType(arr.GetValue(i), CustomVertex.PositionNormal)
            Dim pnt As Direct3D.CustomVertex.PositionNormalTextured = CType(data.GetValue(i), CustomVertex.PositionNormalTextured)

            pnt.X = pn.X
            pnt.Y = pn.Y
            pnt.Z = pn.Z
            pnt.Nx = pn.Nx
            pnt.Ny = pn.Ny
            pnt.Nz = pn.Nz

            If bTopOrBottom = True Then
                If pnt.X < 0 Then
                    pnt.Tv = 0
                Else : pnt.Tv = 1
                End If

                If pnt.Z < 0 Then
                    pnt.Tu = 0
                Else : pnt.Tu = 1
                End If
            Else
                If pnt.Y < 0 Then
                    pnt.Tv = 0
                Else : pnt.Tv = 1
                End If

                If pnt.X < 0 Then
                    If pnt.Z < 0 Then
                        pnt.Tu = 0
                    Else : pnt.Tu = 1
                    End If
                Else
                    If pnt.Z < 0 Then
                        pnt.Tu = 1
                    Else : pnt.Tu = 0
                    End If
                End If
            End If

            'TODO: offset y here, but we may want to remove this later
            pnt.Y += fHalfHeight

            data.SetValue(pnt, i)
        Next i

        TexturedObject.VertexBuffer.Unlock()
        oTemp.VertexBuffer.Unlock()

        ' Set the index buffer. 
        ranks(0) = oTemp.NumberFaces * 3
        arr = oTemp.LockIndexBuffer((New Short()).GetType(), LockFlags.None, ranks)
        TexturedObject.IndexBuffer.SetData(arr, 0, LockFlags.None)

        ' Optimize the mesh for this graphics card's vertex cache 
        ' so when rendering the mesh's triangle list the vertices will 
        ' cache hit more often so it won't have to re-execute the vertex shader 
        ' on those vertices so it will improve perf.     
        Dim adj() As Int32 = TexturedObject.ConvertPointRepsToAdjacency(CType(Nothing, GraphicsStream))
        TexturedObject.OptimizeInPlace(MeshFlags.OptimizeVertexCache, adj)
        Erase adj

        Device.IsUsingEventHandlers = True

        oTemp = Nothing
        arr = Nothing
        data = Nothing
        Return TexturedObject

    End Function


    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Device.IsUsingEventHandlers = False

        moMesh = CreateTexturedBox(1000, 1, 1000)

        'Now, because our engine uses Normals and lighting, let's get that out of the way
        If (moMesh.VertexFormat And VertexFormats.Normal) = 0 Then
            Dim oTmpMesh As Mesh = moMesh.Clone(moMesh.Options.Value, moMesh.VertexFormat Or VertexFormats.Normal, moDevice)
            oTmpMesh.ComputeNormals()
            moMesh.Dispose()
            moMesh = oTmpMesh
            oTmpMesh = Nothing
        End If

        'DoTangentCalcs(moMesh)
        Dim elems(5) As VertexElement
        'position, normal, texture coords, tangent, binormal.

        elems(0) = New VertexElement(0, 0, DeclarationType.Float3, DeclarationMethod.Default, DeclarationUsage.Position, 0)
        elems(1) = New VertexElement(0, 12, DeclarationType.Float3, DeclarationMethod.Default, DeclarationUsage.Normal, 0)
        elems(2) = New VertexElement(0, 24, DeclarationType.Float2, DeclarationMethod.Default, DeclarationUsage.TextureCoordinate, 0)
        elems(3) = New VertexElement(0, 32, DeclarationType.Float3, DeclarationMethod.Default, DeclarationUsage.Tangent, 0)
        elems(4) = New VertexElement(0, 44, DeclarationType.Float3, DeclarationMethod.Default, DeclarationUsage.BiNormal, 0)
        elems(5) = VertexElement.VertexDeclarationEnd

        Dim clonedmesh As Mesh = moMesh.Clone(MeshFlags.Managed, elems, moDevice)
        moMesh.Dispose()
        moMesh = clonedmesh
        clonedmesh = Nothing

        moMesh = Geometry.ComputeTangentFrame(moMesh, DeclarationUsage.TextureCoordinate, 0, DeclarationUsage.Tangent, 0, DeclarationUsage.BiNormal, 0, _
        DeclarationUsage.Normal, 0, 0, Nothing, 0.0F, 0.0F, 0.0F, Nothing)


        Device.IsUsingEventHandlers = True
    End Sub

    Private Sub btnPlayerRel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPlayerRel.Click
        cldDlg.Color = btnPlayerRel.BackColor
        Dim vbRes As DialogResult = cldDlg.ShowDialog
        If vbRes = Windows.Forms.DialogResult.OK Then
            btnPlayerRel.BackColor = cldDlg.Color
        End If
    End Sub

    Private Sub btnSetSpecTex_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnReloadTextures.Click
        Dim oRes As DialogResult

        opnDlg.Filter = "Image Files *.*|*.*"
        opnDlg.FileName = msSpecTexture
        oRes = opnDlg.ShowDialog(Me)

        If oRes = Windows.Forms.DialogResult.OK Then
            msSpecTexture = opnDlg.FileName

            If moSpec Is Nothing = False Then moSpec.Dispose()
            moSpec = Nothing
            moSpec = TextureLoader.FromFile(moDevice, msSpecTexture)

            If moShader Is Nothing = False Then moShader.SetSpecTexture(moSpec)
        End If
    End Sub

    Private Sub Form1_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyUp
        Select Case e.KeyCode

            Case Keys.Up
                mbScrollUp = False
            Case Keys.Down
                mbScrollDown = False
            Case Keys.Left
                mbScrollLeft = False
            Case Keys.Right
                mbScrollRight = False

        End Select
    End Sub

    Private Sub Form1_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
        'Me.Height = CInt(Me.Width * (Screen.PrimaryScreen.WorkingArea.Height / Screen.PrimaryScreen.WorkingArea.Width))

        Me.Width = 1024
        Me.Height = 768
        Me.Left = 0
        Me.Top = 0
        If InitD3D(picMain) = False Then End

        mlCameraAtX = 0 : mlCameraAtY = 0 : mlCameraAtZ = 0
        mlCameraX = 0 : mlCameraY = 1000 : mlCameraZ = -1000

        'mlCameraAtX = -4905599 : mlCameraAtY = 0 : mlCameraAtZ = 2108452
        'mlCameraX = -4905599 : mlCameraY = 0 : mlCameraZ = 2106452

        tmrRender.Enabled = True

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim sFile As String = AppDomain.CurrentDomain.BaseDirectory
        If sFile.EndsWith("\") = False Then sFile &= "\"
        sFile &= "SS_" & Now.ToString("MM_dd_yyyy_HHmmss") '& ".bmp"

        'render the mesh normal
        DrawScene(False)
        DoCaptureScreenshot(sFile & ".bmp")

        'render the mesh white
        DrawScene(True)
        DoCaptureScreenshot(sFile & "_a.bmp")
    End Sub

    Private Sub DoCaptureScreenshot(ByVal sFile As String)
        Dim oSurf As Surface

        Dim ifmt As Microsoft.DirectX.Direct3D.ImageFileFormat = ImageFileFormat.Bmp


        oSurf = moDevice.GetBackBuffer(0, 0, BackBufferType.Mono)
        SurfaceLoader.Save(sFile, ifmt, oSurf)
        oSurf.Dispose()
        oSurf = Nothing
    End Sub
End Class

