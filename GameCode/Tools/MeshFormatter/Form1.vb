Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

Public Class Form1
    Inherits System.Windows.Forms.Form

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
    Friend WithEvents mnuMain As System.Windows.Forms.MainMenu
    Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem5 As System.Windows.Forms.MenuItem
    Friend WithEvents tmrRender As System.Windows.Forms.Timer
    Friend WithEvents mnuOpen As System.Windows.Forms.MenuItem
    Friend WithEvents mnuExit As System.Windows.Forms.MenuItem
    Friend WithEvents opnDlg As System.Windows.Forms.OpenFileDialog
    Friend WithEvents savDlg As System.Windows.Forms.SaveFileDialog
    Friend WithEvents picMain As System.Windows.Forms.PictureBox
    Friend WithEvents txtSizeX As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtSizeY As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtSizeZ As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtMinX As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtMinY As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents txtMinZ As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents txtMaxX As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents txtMaxY As System.Windows.Forms.TextBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents txtMaxZ As System.Windows.Forms.TextBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents txtMidX As System.Windows.Forms.TextBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents txtMidY As System.Windows.Forms.TextBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents txtMidZ As System.Windows.Forms.TextBox
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents cmdDiffuse As System.Windows.Forms.Button
    Friend WithEvents cmdEmissive As System.Windows.Forms.Button
    Friend WithEvents cmdSpecular As System.Windows.Forms.Button
    Friend WithEvents cldDlg As System.Windows.Forms.ColorDialog
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents txtScaleZ As System.Windows.Forms.TextBox
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents txtScaleY As System.Windows.Forms.TextBox
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents txtScaleX As System.Windows.Forms.TextBox
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents txtShiftZ As System.Windows.Forms.TextBox
    Friend WithEvents Label20 As System.Windows.Forms.Label
    Friend WithEvents txtShiftY As System.Windows.Forms.TextBox
    Friend WithEvents Label21 As System.Windows.Forms.Label
    Friend WithEvents txtShiftX As System.Windows.Forms.TextBox
    Friend WithEvents cmdScale As System.Windows.Forms.Button
    Friend WithEvents cmdShift As System.Windows.Forms.Button
    Friend WithEvents cmdReloadTexture As System.Windows.Forms.Button
    Friend WithEvents optMine As System.Windows.Forms.RadioButton
    Friend WithEvents optAlly As System.Windows.Forms.RadioButton
    Friend WithEvents optEnemy As System.Windows.Forms.RadioButton
    Friend WithEvents optNeutral As System.Windows.Forms.RadioButton
    Friend WithEvents tbctrlMain As System.Windows.Forms.TabControl
    Friend WithEvents tbp1 As System.Windows.Forms.TabPage
    Friend WithEvents tbp2 As System.Windows.Forms.TabPage
    Friend WithEvents txtDetails As System.Windows.Forms.TextBox
    Friend WithEvents Label22 As System.Windows.Forms.Label
    Friend WithEvents txtFileName As System.Windows.Forms.TextBox
    Friend WithEvents Label23 As System.Windows.Forms.Label
    Friend WithEvents txtModelID As System.Windows.Forms.TextBox
    Friend WithEvents Label24 As System.Windows.Forms.Label
    Friend WithEvents txtPlanetYAdjust As System.Windows.Forms.TextBox
    Friend WithEvents fraShieldSphere As System.Windows.Forms.GroupBox
    Friend WithEvents Label25 As System.Windows.Forms.Label
    Friend WithEvents txtShieldSphereSize As System.Windows.Forms.TextBox
    Friend WithEvents Label26 As System.Windows.Forms.Label
    Friend WithEvents txtShieldStretchX As System.Windows.Forms.TextBox
    Friend WithEvents txtShieldStretchY As System.Windows.Forms.TextBox
    Friend WithEvents txtShieldStretchZ As System.Windows.Forms.TextBox
    Friend WithEvents fraEngineFX As System.Windows.Forms.GroupBox
    Friend WithEvents lstEngineFX As System.Windows.Forms.ListBox
    Friend WithEvents Label27 As System.Windows.Forms.Label
    Friend WithEvents btnAddEngineFX As System.Windows.Forms.Button
    Friend WithEvents btnDeleteEngineFX As System.Windows.Forms.Button
    Friend WithEvents btnResetShieldSphere As System.Windows.Forms.Button
    Friend WithEvents txtEngineOffZ As System.Windows.Forms.TextBox
    Friend WithEvents txtEngineOffY As System.Windows.Forms.TextBox
    Friend WithEvents txtEngineOffX As System.Windows.Forms.TextBox
    Friend WithEvents mnuRenderOptions As System.Windows.Forms.MenuItem
    Friend WithEvents mnuWireFrame As System.Windows.Forms.MenuItem
    Friend WithEvents mnuRenderPlane As System.Windows.Forms.MenuItem
    Friend WithEvents mnuRenderShield As System.Windows.Forms.MenuItem
    Friend WithEvents mnuRenderEngineFX As System.Windows.Forms.MenuItem
    Friend WithEvents mnuLockCamera As System.Windows.Forms.MenuItem
    Friend WithEvents Label28 As System.Windows.Forms.Label
    Friend WithEvents txtShieldShiftZ As System.Windows.Forms.TextBox
    Friend WithEvents txtShieldShiftY As System.Windows.Forms.TextBox
    Friend WithEvents txtShieldShiftX As System.Windows.Forms.TextBox
    Friend WithEvents txtPCnt As System.Windows.Forms.TextBox
    Friend WithEvents Label29 As System.Windows.Forms.Label
    Friend WithEvents btnInvertZ As System.Windows.Forms.Button
    Friend WithEvents mnuOpenX As System.Windows.Forms.MenuItem
    Friend WithEvents mnuOpenCompareX As System.Windows.Forms.MenuItem
    Friend WithEvents mnuRenderCompare As System.Windows.Forms.MenuItem
    Friend WithEvents Label30 As System.Windows.Forms.Label
    Friend WithEvents txtTurretOffZ As System.Windows.Forms.TextBox
    Friend WithEvents mnuOpenTank As System.Windows.Forms.MenuItem
    Friend WithEvents mnuSaveXFileOnly As System.Windows.Forms.MenuItem
    Friend WithEvents btnReverseWind As System.Windows.Forms.Button
    Friend WithEvents lblRngOff As System.Windows.Forms.Label
    Friend WithEvents tbp3 As System.Windows.Forms.TabPage
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Label31 As System.Windows.Forms.Label
    Friend WithEvents txtBurnPCnt As System.Windows.Forms.TextBox
    Friend WithEvents btnDeleteBurnFX As System.Windows.Forms.Button
    Friend WithEvents btnAddBurnFX As System.Windows.Forms.Button
    Friend WithEvents txtBurnOffZ As System.Windows.Forms.TextBox
    Friend WithEvents txtBurnOffY As System.Windows.Forms.TextBox
    Friend WithEvents Label32 As System.Windows.Forms.Label
    Friend WithEvents txtBurnOffX As System.Windows.Forms.TextBox
    Friend WithEvents lstBurn As System.Windows.Forms.ListBox
    Friend WithEvents mnuShowDeath As System.Windows.Forms.MenuItem
    Friend WithEvents optFrontBurn As System.Windows.Forms.RadioButton
    Friend WithEvents optLeftBurn As System.Windows.Forms.RadioButton
    Friend WithEvents optRearBurn As System.Windows.Forms.RadioButton
    Friend WithEvents optRightBurn As System.Windows.Forms.RadioButton
    Friend WithEvents chkLandBased As System.Windows.Forms.CheckBox
    Friend WithEvents hscrYaw As System.Windows.Forms.HScrollBar
    Friend WithEvents Label33 As System.Windows.Forms.Label
    Friend WithEvents Label34 As System.Windows.Forms.Label
    Friend WithEvents hscrRotate As System.Windows.Forms.HScrollBar
    Friend WithEvents mnuAtmosBurn As System.Windows.Forms.MenuItem
    Friend WithEvents mnuSaveParent As System.Windows.Forms.MenuItem
    Friend WithEvents mnuSave As System.Windows.Forms.MenuItem
    Friend WithEvents mnuSaveDetailsOnly As System.Windows.Forms.MenuItem
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents Label36 As System.Windows.Forms.Label
    Friend WithEvents txtFireFromOffZ As System.Windows.Forms.TextBox
    Friend WithEvents txtFireFromOffY As System.Windows.Forms.TextBox
    Friend WithEvents txtFireFromOffX As System.Windows.Forms.TextBox
    Friend WithEvents lstFireFroms As System.Windows.Forms.ListBox
    Friend WithEvents btnDeleteFireFrom As System.Windows.Forms.Button
    Friend WithEvents btnAddFireFrom As System.Windows.Forms.Button
    Friend WithEvents optFireFromRight As System.Windows.Forms.RadioButton
    Friend WithEvents optFireFromRear As System.Windows.Forms.RadioButton
    Friend WithEvents optFireFromLeft As System.Windows.Forms.RadioButton
    Friend WithEvents optFireFromFront As System.Windows.Forms.RadioButton
    Friend WithEvents tbp4 As System.Windows.Forms.TabPage
    Friend WithEvents dgvBL As System.Windows.Forms.DataGridView
    Friend WithEvents mnuBL As System.Windows.Forms.MenuItem
    Friend WithEvents colOffsetX As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colOffsetY As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colOffsetZ As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colColorR As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colColorG As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colColorB As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colDecayRate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents btnInvertNormals As System.Windows.Forms.Button
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents optMinusZ As System.Windows.Forms.RadioButton
    Friend WithEvents optMinusY As System.Windows.Forms.RadioButton
    Friend WithEvents optMinusX As System.Windows.Forms.RadioButton
    Friend WithEvents optPosZ As System.Windows.Forms.RadioButton
    Friend WithEvents optPosY As System.Windows.Forms.RadioButton
    Friend WithEvents optPosX As System.Windows.Forms.RadioButton
    Friend WithEvents btnResetCamera As System.Windows.Forms.Button
    Friend WithEvents tbp5 As System.Windows.Forms.TabPage
    Friend WithEvents Label39 As System.Windows.Forms.Label
    Friend WithEvents txtDS_ExpCnt As System.Windows.Forms.TextBox
    Friend WithEvents chkFinale As System.Windows.Forms.CheckBox
    Friend WithEvents Label35 As System.Windows.Forms.Label
    Friend WithEvents txtDS_SizeZ As System.Windows.Forms.TextBox
    Friend WithEvents Label37 As System.Windows.Forms.Label
    Friend WithEvents txtDS_SizeY As System.Windows.Forms.TextBox
    Friend WithEvents Label38 As System.Windows.Forms.Label
    Friend WithEvents txtDS_SizeX As System.Windows.Forms.TextBox
    Friend WithEvents tbp6 As System.Windows.Forms.TabPage
    Friend WithEvents dgvVerts As System.Windows.Forms.DataGridView
    Friend WithEvents DataGridViewTextBoxColumn1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn3 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn4 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn5 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents hscrPMesh As System.Windows.Forms.HScrollBar
    Friend WithEvents btnPMesh As System.Windows.Forms.Button
    Friend WithEvents Label40 As System.Windows.Forms.Label
    Friend WithEvents btnShieldColor As System.Windows.Forms.Button
    Friend WithEvents btnEngineColor As System.Windows.Forms.Button
    Friend WithEvents MenuItem2 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuXY As System.Windows.Forms.MenuItem
    Friend WithEvents mnuXZ As System.Windows.Forms.MenuItem
    Friend WithEvents mnuYX As System.Windows.Forms.MenuItem
    Friend WithEvents mnuYZ As System.Windows.Forms.MenuItem
    Friend WithEvents mnuZX As System.Windows.Forms.MenuItem
    Friend WithEvents mnuZY As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem3 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuShaders As System.Windows.Forms.MenuItem
    Friend WithEvents mnuGlow As System.Windows.Forms.MenuItem
    Friend WithEvents txtSpecSharp As System.Windows.Forms.TextBox
	Friend WithEvents mnuComputeNormals As System.Windows.Forms.MenuItem
	Friend WithEvents Label41 As System.Windows.Forms.Label
	Friend WithEvents btnRemoveSFX As System.Windows.Forms.Button
	Friend WithEvents btnAddSFX As System.Windows.Forms.Button
	Friend WithEvents lstSFX As System.Windows.Forms.ListBox
    Friend WithEvents txtSoundEffect As System.Windows.Forms.TextBox
    Friend WithEvents cboBurn As System.Windows.Forms.ComboBox
    Friend WithEvents cboDS As System.Windows.Forms.ComboBox
	Friend WithEvents mnuFireFrom As System.Windows.Forms.MenuItem
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Me.mnuMain = New System.Windows.Forms.MainMenu(Me.components)
        Me.MenuItem1 = New System.Windows.Forms.MenuItem
        Me.mnuOpen = New System.Windows.Forms.MenuItem
        Me.mnuOpenX = New System.Windows.Forms.MenuItem
        Me.mnuOpenTank = New System.Windows.Forms.MenuItem
        Me.mnuOpenCompareX = New System.Windows.Forms.MenuItem
        Me.mnuSaveParent = New System.Windows.Forms.MenuItem
        Me.mnuSave = New System.Windows.Forms.MenuItem
        Me.mnuSaveXFileOnly = New System.Windows.Forms.MenuItem
        Me.mnuSaveDetailsOnly = New System.Windows.Forms.MenuItem
        Me.MenuItem5 = New System.Windows.Forms.MenuItem
        Me.mnuExit = New System.Windows.Forms.MenuItem
        Me.mnuRenderOptions = New System.Windows.Forms.MenuItem
        Me.mnuWireFrame = New System.Windows.Forms.MenuItem
        Me.mnuRenderPlane = New System.Windows.Forms.MenuItem
        Me.mnuRenderShield = New System.Windows.Forms.MenuItem
        Me.mnuRenderEngineFX = New System.Windows.Forms.MenuItem
        Me.mnuLockCamera = New System.Windows.Forms.MenuItem
        Me.mnuRenderCompare = New System.Windows.Forms.MenuItem
        Me.mnuShowDeath = New System.Windows.Forms.MenuItem
        Me.mnuAtmosBurn = New System.Windows.Forms.MenuItem
        Me.mnuFireFrom = New System.Windows.Forms.MenuItem
        Me.mnuBL = New System.Windows.Forms.MenuItem
        Me.MenuItem2 = New System.Windows.Forms.MenuItem
        Me.mnuXY = New System.Windows.Forms.MenuItem
        Me.mnuXZ = New System.Windows.Forms.MenuItem
        Me.mnuYX = New System.Windows.Forms.MenuItem
        Me.mnuYZ = New System.Windows.Forms.MenuItem
        Me.mnuZX = New System.Windows.Forms.MenuItem
        Me.mnuZY = New System.Windows.Forms.MenuItem
        Me.mnuComputeNormals = New System.Windows.Forms.MenuItem
        Me.MenuItem3 = New System.Windows.Forms.MenuItem
        Me.mnuShaders = New System.Windows.Forms.MenuItem
        Me.mnuGlow = New System.Windows.Forms.MenuItem
        Me.tmrRender = New System.Windows.Forms.Timer(Me.components)
        Me.opnDlg = New System.Windows.Forms.OpenFileDialog
        Me.savDlg = New System.Windows.Forms.SaveFileDialog
        Me.picMain = New System.Windows.Forms.PictureBox
        Me.txtSizeX = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.txtSizeY = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.txtSizeZ = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.txtMinX = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.txtMinY = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.txtMinZ = New System.Windows.Forms.TextBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.txtMaxX = New System.Windows.Forms.TextBox
        Me.Label8 = New System.Windows.Forms.Label
        Me.txtMaxY = New System.Windows.Forms.TextBox
        Me.Label9 = New System.Windows.Forms.Label
        Me.txtMaxZ = New System.Windows.Forms.TextBox
        Me.Label10 = New System.Windows.Forms.Label
        Me.txtMidX = New System.Windows.Forms.TextBox
        Me.Label11 = New System.Windows.Forms.Label
        Me.txtMidY = New System.Windows.Forms.TextBox
        Me.Label12 = New System.Windows.Forms.Label
        Me.txtMidZ = New System.Windows.Forms.TextBox
        Me.Label13 = New System.Windows.Forms.Label
        Me.Label14 = New System.Windows.Forms.Label
        Me.Label15 = New System.Windows.Forms.Label
        Me.cmdDiffuse = New System.Windows.Forms.Button
        Me.cmdEmissive = New System.Windows.Forms.Button
        Me.cmdSpecular = New System.Windows.Forms.Button
        Me.cldDlg = New System.Windows.Forms.ColorDialog
        Me.Label16 = New System.Windows.Forms.Label
        Me.txtScaleZ = New System.Windows.Forms.TextBox
        Me.Label17 = New System.Windows.Forms.Label
        Me.txtScaleY = New System.Windows.Forms.TextBox
        Me.Label18 = New System.Windows.Forms.Label
        Me.txtScaleX = New System.Windows.Forms.TextBox
        Me.Label19 = New System.Windows.Forms.Label
        Me.txtShiftZ = New System.Windows.Forms.TextBox
        Me.Label20 = New System.Windows.Forms.Label
        Me.txtShiftY = New System.Windows.Forms.TextBox
        Me.Label21 = New System.Windows.Forms.Label
        Me.txtShiftX = New System.Windows.Forms.TextBox
        Me.cmdScale = New System.Windows.Forms.Button
        Me.cmdShift = New System.Windows.Forms.Button
        Me.cmdReloadTexture = New System.Windows.Forms.Button
        Me.optNeutral = New System.Windows.Forms.RadioButton
        Me.optMine = New System.Windows.Forms.RadioButton
        Me.optAlly = New System.Windows.Forms.RadioButton
        Me.optEnemy = New System.Windows.Forms.RadioButton
        Me.tbctrlMain = New System.Windows.Forms.TabControl
        Me.tbp1 = New System.Windows.Forms.TabPage
        Me.txtDetails = New System.Windows.Forms.TextBox
        Me.tbp2 = New System.Windows.Forms.TabPage
        Me.chkLandBased = New System.Windows.Forms.CheckBox
        Me.Label30 = New System.Windows.Forms.Label
        Me.txtTurretOffZ = New System.Windows.Forms.TextBox
        Me.fraEngineFX = New System.Windows.Forms.GroupBox
        Me.Label29 = New System.Windows.Forms.Label
        Me.txtPCnt = New System.Windows.Forms.TextBox
        Me.btnDeleteEngineFX = New System.Windows.Forms.Button
        Me.btnAddEngineFX = New System.Windows.Forms.Button
        Me.txtEngineOffZ = New System.Windows.Forms.TextBox
        Me.txtEngineOffY = New System.Windows.Forms.TextBox
        Me.Label27 = New System.Windows.Forms.Label
        Me.txtEngineOffX = New System.Windows.Forms.TextBox
        Me.lstEngineFX = New System.Windows.Forms.ListBox
        Me.fraShieldSphere = New System.Windows.Forms.GroupBox
        Me.txtShieldShiftZ = New System.Windows.Forms.TextBox
        Me.txtShieldShiftY = New System.Windows.Forms.TextBox
        Me.Label28 = New System.Windows.Forms.Label
        Me.txtShieldShiftX = New System.Windows.Forms.TextBox
        Me.btnResetShieldSphere = New System.Windows.Forms.Button
        Me.txtShieldStretchZ = New System.Windows.Forms.TextBox
        Me.txtShieldStretchY = New System.Windows.Forms.TextBox
        Me.Label26 = New System.Windows.Forms.Label
        Me.txtShieldStretchX = New System.Windows.Forms.TextBox
        Me.Label25 = New System.Windows.Forms.Label
        Me.txtShieldSphereSize = New System.Windows.Forms.TextBox
        Me.Label24 = New System.Windows.Forms.Label
        Me.txtPlanetYAdjust = New System.Windows.Forms.TextBox
        Me.Label23 = New System.Windows.Forms.Label
        Me.txtModelID = New System.Windows.Forms.TextBox
        Me.Label22 = New System.Windows.Forms.Label
        Me.txtFileName = New System.Windows.Forms.TextBox
        Me.tbp3 = New System.Windows.Forms.TabPage
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.optFireFromRight = New System.Windows.Forms.RadioButton
        Me.optFireFromRear = New System.Windows.Forms.RadioButton
        Me.optFireFromLeft = New System.Windows.Forms.RadioButton
        Me.optFireFromFront = New System.Windows.Forms.RadioButton
        Me.btnDeleteFireFrom = New System.Windows.Forms.Button
        Me.btnAddFireFrom = New System.Windows.Forms.Button
        Me.Label36 = New System.Windows.Forms.Label
        Me.lstFireFroms = New System.Windows.Forms.ListBox
        Me.txtFireFromOffX = New System.Windows.Forms.TextBox
        Me.txtFireFromOffY = New System.Windows.Forms.TextBox
        Me.txtFireFromOffZ = New System.Windows.Forms.TextBox
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.optMinusZ = New System.Windows.Forms.RadioButton
        Me.optMinusY = New System.Windows.Forms.RadioButton
        Me.optMinusX = New System.Windows.Forms.RadioButton
        Me.optPosZ = New System.Windows.Forms.RadioButton
        Me.optPosY = New System.Windows.Forms.RadioButton
        Me.optPosX = New System.Windows.Forms.RadioButton
        Me.optRightBurn = New System.Windows.Forms.RadioButton
        Me.optRearBurn = New System.Windows.Forms.RadioButton
        Me.optLeftBurn = New System.Windows.Forms.RadioButton
        Me.optFrontBurn = New System.Windows.Forms.RadioButton
        Me.Label31 = New System.Windows.Forms.Label
        Me.txtBurnPCnt = New System.Windows.Forms.TextBox
        Me.btnDeleteBurnFX = New System.Windows.Forms.Button
        Me.btnAddBurnFX = New System.Windows.Forms.Button
        Me.txtBurnOffZ = New System.Windows.Forms.TextBox
        Me.txtBurnOffY = New System.Windows.Forms.TextBox
        Me.Label32 = New System.Windows.Forms.Label
        Me.txtBurnOffX = New System.Windows.Forms.TextBox
        Me.lstBurn = New System.Windows.Forms.ListBox
        Me.tbp4 = New System.Windows.Forms.TabPage
        Me.dgvBL = New System.Windows.Forms.DataGridView
        Me.colOffsetX = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colOffsetY = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colOffsetZ = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colColorR = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colColorG = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colColorB = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colDecayRate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.tbp5 = New System.Windows.Forms.TabPage
        Me.Label41 = New System.Windows.Forms.Label
        Me.btnRemoveSFX = New System.Windows.Forms.Button
        Me.btnAddSFX = New System.Windows.Forms.Button
        Me.lstSFX = New System.Windows.Forms.ListBox
        Me.txtSoundEffect = New System.Windows.Forms.TextBox
        Me.Label39 = New System.Windows.Forms.Label
        Me.txtDS_ExpCnt = New System.Windows.Forms.TextBox
        Me.chkFinale = New System.Windows.Forms.CheckBox
        Me.Label35 = New System.Windows.Forms.Label
        Me.txtDS_SizeZ = New System.Windows.Forms.TextBox
        Me.Label37 = New System.Windows.Forms.Label
        Me.txtDS_SizeY = New System.Windows.Forms.TextBox
        Me.Label38 = New System.Windows.Forms.Label
        Me.txtDS_SizeX = New System.Windows.Forms.TextBox
        Me.tbp6 = New System.Windows.Forms.TabPage
        Me.dgvVerts = New System.Windows.Forms.DataGridView
        Me.DataGridViewTextBoxColumn1 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DataGridViewTextBoxColumn2 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DataGridViewTextBoxColumn3 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DataGridViewTextBoxColumn4 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DataGridViewTextBoxColumn5 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.btnInvertZ = New System.Windows.Forms.Button
        Me.btnReverseWind = New System.Windows.Forms.Button
        Me.lblRngOff = New System.Windows.Forms.Label
        Me.hscrYaw = New System.Windows.Forms.HScrollBar
        Me.hscrRotate = New System.Windows.Forms.HScrollBar
        Me.Label33 = New System.Windows.Forms.Label
        Me.Label34 = New System.Windows.Forms.Label
        Me.btnInvertNormals = New System.Windows.Forms.Button
        Me.Button1 = New System.Windows.Forms.Button
        Me.btnResetCamera = New System.Windows.Forms.Button
        Me.hscrPMesh = New System.Windows.Forms.HScrollBar
        Me.btnPMesh = New System.Windows.Forms.Button
        Me.Label40 = New System.Windows.Forms.Label
        Me.btnShieldColor = New System.Windows.Forms.Button
        Me.btnEngineColor = New System.Windows.Forms.Button
        Me.txtSpecSharp = New System.Windows.Forms.TextBox
        Me.cboBurn = New System.Windows.Forms.ComboBox
        Me.cboDS = New System.Windows.Forms.ComboBox
        CType(Me.picMain, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.tbctrlMain.SuspendLayout()
        Me.tbp1.SuspendLayout()
        Me.tbp2.SuspendLayout()
        Me.fraEngineFX.SuspendLayout()
        Me.fraShieldSphere.SuspendLayout()
        Me.tbp3.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.tbp4.SuspendLayout()
        CType(Me.dgvBL, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.tbp5.SuspendLayout()
        Me.tbp6.SuspendLayout()
        CType(Me.dgvVerts, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'mnuMain
        '
        Me.mnuMain.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem1, Me.mnuRenderOptions, Me.MenuItem2, Me.MenuItem3})
        '
        'MenuItem1
        '
        Me.MenuItem1.Index = 0
        Me.MenuItem1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuOpen, Me.mnuSaveParent, Me.MenuItem5, Me.mnuExit})
        Me.MenuItem1.Text = "File"
        '
        'mnuOpen
        '
        Me.mnuOpen.Index = 0
        Me.mnuOpen.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuOpenX, Me.mnuOpenTank, Me.mnuOpenCompareX})
        Me.mnuOpen.Text = "Open"
        '
        'mnuOpenX
        '
        Me.mnuOpenX.Index = 0
        Me.mnuOpenX.Text = "From .X File..."
        '
        'mnuOpenTank
        '
        Me.mnuOpenTank.Index = 1
        Me.mnuOpenTank.Text = "From Tank .X File..."
        '
        'mnuOpenCompareX
        '
        Me.mnuOpenCompareX.Index = 2
        Me.mnuOpenCompareX.Text = "Compare Model .X File..."
        '
        'mnuSaveParent
        '
        Me.mnuSaveParent.Index = 1
        Me.mnuSaveParent.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuSave, Me.mnuSaveXFileOnly, Me.mnuSaveDetailsOnly})
        Me.mnuSaveParent.Text = "Save"
        '
        'mnuSave
        '
        Me.mnuSave.Index = 0
        Me.mnuSave.Text = "Save .X and Details"
        '
        'mnuSaveXFileOnly
        '
        Me.mnuSaveXFileOnly.Index = 1
        Me.mnuSaveXFileOnly.Text = "Save .X File Only..."
        '
        'mnuSaveDetailsOnly
        '
        Me.mnuSaveDetailsOnly.Index = 2
        Me.mnuSaveDetailsOnly.Text = "Save Details Only..."
        '
        'MenuItem5
        '
        Me.MenuItem5.Index = 2
        Me.MenuItem5.Text = "-"
        '
        'mnuExit
        '
        Me.mnuExit.Index = 3
        Me.mnuExit.Text = "Exit"
        '
        'mnuRenderOptions
        '
        Me.mnuRenderOptions.Index = 1
        Me.mnuRenderOptions.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuWireFrame, Me.mnuRenderPlane, Me.mnuRenderShield, Me.mnuRenderEngineFX, Me.mnuLockCamera, Me.mnuRenderCompare, Me.mnuShowDeath, Me.mnuAtmosBurn, Me.mnuFireFrom, Me.mnuBL})
        Me.mnuRenderOptions.Text = "Render Options"
        '
        'mnuWireFrame
        '
        Me.mnuWireFrame.Index = 0
        Me.mnuWireFrame.Text = "Wireframe"
        '
        'mnuRenderPlane
        '
        Me.mnuRenderPlane.Index = 1
        Me.mnuRenderPlane.Text = "Render 0 Plane"
        '
        'mnuRenderShield
        '
        Me.mnuRenderShield.Index = 2
        Me.mnuRenderShield.Text = "Render Shield Sphere"
        '
        'mnuRenderEngineFX
        '
        Me.mnuRenderEngineFX.Index = 3
        Me.mnuRenderEngineFX.Text = "Render Fire FX"
        '
        'mnuLockCamera
        '
        Me.mnuLockCamera.Checked = True
        Me.mnuLockCamera.Index = 4
        Me.mnuLockCamera.Shortcut = System.Windows.Forms.Shortcut.CtrlL
        Me.mnuLockCamera.Text = "Lock Camera"
        '
        'mnuRenderCompare
        '
        Me.mnuRenderCompare.Index = 5
        Me.mnuRenderCompare.Text = "Render Compare Mesh"
        '
        'mnuShowDeath
        '
        Me.mnuShowDeath.Index = 6
        Me.mnuShowDeath.Text = "Show Explosion Sequence"
        '
        'mnuAtmosBurn
        '
        Me.mnuAtmosBurn.Index = 7
        Me.mnuAtmosBurn.Text = "Atmospheric Burn"
        '
        'mnuFireFrom
        '
        Me.mnuFireFrom.Index = 8
        Me.mnuFireFrom.Text = "Render Fire From Locs"
        '
        'mnuBL
        '
        Me.mnuBL.Index = 9
        Me.mnuBL.Text = "Render Blinking Lights"
        '
        'MenuItem2
        '
        Me.MenuItem2.Index = 2
        Me.MenuItem2.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuXY, Me.mnuXZ, Me.mnuYX, Me.mnuYZ, Me.mnuZX, Me.mnuZY, Me.mnuComputeNormals})
        Me.MenuItem2.Text = "Transpose Mesh"
        '
        'mnuXY
        '
        Me.mnuXY.Index = 0
        Me.mnuXY.Text = "X -> Y"
        '
        'mnuXZ
        '
        Me.mnuXZ.Index = 1
        Me.mnuXZ.Text = "X -> Z"
        '
        'mnuYX
        '
        Me.mnuYX.Index = 2
        Me.mnuYX.Text = "Y -> X"
        '
        'mnuYZ
        '
        Me.mnuYZ.Index = 3
        Me.mnuYZ.Text = "Y -> Z"
        '
        'mnuZX
        '
        Me.mnuZX.Index = 4
        Me.mnuZX.Text = "Z -> X"
        '
        'mnuZY
        '
        Me.mnuZY.Index = 5
        Me.mnuZY.Text = "Z -> Y"
        '
        'mnuComputeNormals
        '
        Me.mnuComputeNormals.Index = 6
        Me.mnuComputeNormals.Text = "Compute Normals"
        '
        'MenuItem3
        '
        Me.MenuItem3.Index = 3
        Me.MenuItem3.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuShaders, Me.mnuGlow})
        Me.MenuItem3.Text = "Shaders"
        '
        'mnuShaders
        '
        Me.mnuShaders.Index = 0
        Me.mnuShaders.Text = "Do Shaders"
        '
        'mnuGlow
        '
        Me.mnuGlow.Index = 1
        Me.mnuGlow.Text = "Do Glow FX"
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
        Me.picMain.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.picMain.Location = New System.Drawing.Point(8, 8)
        Me.picMain.Name = "picMain"
        Me.picMain.Size = New System.Drawing.Size(672, 644)
        Me.picMain.TabIndex = 0
        Me.picMain.TabStop = False
        '
        'txtSizeX
        '
        Me.txtSizeX.Location = New System.Drawing.Point(768, 8)
        Me.txtSizeX.Name = "txtSizeX"
        Me.txtSizeX.ReadOnly = True
        Me.txtSizeX.Size = New System.Drawing.Size(100, 20)
        Me.txtSizeX.TabIndex = 1
        Me.txtSizeX.Text = "0"
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(696, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(56, 24)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Size X"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(696, 32)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(56, 24)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "Size Y"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtSizeY
        '
        Me.txtSizeY.Location = New System.Drawing.Point(768, 32)
        Me.txtSizeY.Name = "txtSizeY"
        Me.txtSizeY.ReadOnly = True
        Me.txtSizeY.Size = New System.Drawing.Size(100, 20)
        Me.txtSizeY.TabIndex = 3
        Me.txtSizeY.Text = "0"
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(696, 56)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(56, 24)
        Me.Label3.TabIndex = 6
        Me.Label3.Text = "Size Z"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtSizeZ
        '
        Me.txtSizeZ.Location = New System.Drawing.Point(768, 56)
        Me.txtSizeZ.Name = "txtSizeZ"
        Me.txtSizeZ.ReadOnly = True
        Me.txtSizeZ.Size = New System.Drawing.Size(100, 20)
        Me.txtSizeZ.TabIndex = 5
        Me.txtSizeZ.Text = "0"
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(696, 96)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(56, 24)
        Me.Label4.TabIndex = 8
        Me.Label4.Text = "Min X"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtMinX
        '
        Me.txtMinX.Location = New System.Drawing.Point(768, 96)
        Me.txtMinX.Name = "txtMinX"
        Me.txtMinX.ReadOnly = True
        Me.txtMinX.Size = New System.Drawing.Size(100, 20)
        Me.txtMinX.TabIndex = 7
        Me.txtMinX.Text = "0"
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(696, 120)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(56, 24)
        Me.Label5.TabIndex = 10
        Me.Label5.Text = "Min Y"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtMinY
        '
        Me.txtMinY.Location = New System.Drawing.Point(768, 120)
        Me.txtMinY.Name = "txtMinY"
        Me.txtMinY.ReadOnly = True
        Me.txtMinY.Size = New System.Drawing.Size(100, 20)
        Me.txtMinY.TabIndex = 9
        Me.txtMinY.Text = "0"
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(696, 144)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(56, 24)
        Me.Label6.TabIndex = 12
        Me.Label6.Text = "Min Z"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtMinZ
        '
        Me.txtMinZ.Location = New System.Drawing.Point(768, 144)
        Me.txtMinZ.Name = "txtMinZ"
        Me.txtMinZ.ReadOnly = True
        Me.txtMinZ.Size = New System.Drawing.Size(100, 20)
        Me.txtMinZ.TabIndex = 11
        Me.txtMinZ.Text = "0"
        '
        'Label7
        '
        Me.Label7.Location = New System.Drawing.Point(696, 176)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(56, 24)
        Me.Label7.TabIndex = 14
        Me.Label7.Text = "Max X"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtMaxX
        '
        Me.txtMaxX.Location = New System.Drawing.Point(768, 176)
        Me.txtMaxX.Name = "txtMaxX"
        Me.txtMaxX.ReadOnly = True
        Me.txtMaxX.Size = New System.Drawing.Size(100, 20)
        Me.txtMaxX.TabIndex = 13
        Me.txtMaxX.Text = "0"
        '
        'Label8
        '
        Me.Label8.Location = New System.Drawing.Point(696, 200)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(56, 24)
        Me.Label8.TabIndex = 16
        Me.Label8.Text = "Max Y"
        Me.Label8.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtMaxY
        '
        Me.txtMaxY.Location = New System.Drawing.Point(768, 200)
        Me.txtMaxY.Name = "txtMaxY"
        Me.txtMaxY.ReadOnly = True
        Me.txtMaxY.Size = New System.Drawing.Size(100, 20)
        Me.txtMaxY.TabIndex = 15
        Me.txtMaxY.Text = "0"
        '
        'Label9
        '
        Me.Label9.Location = New System.Drawing.Point(696, 224)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(56, 24)
        Me.Label9.TabIndex = 18
        Me.Label9.Text = "Max Z"
        Me.Label9.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtMaxZ
        '
        Me.txtMaxZ.Location = New System.Drawing.Point(768, 224)
        Me.txtMaxZ.Name = "txtMaxZ"
        Me.txtMaxZ.ReadOnly = True
        Me.txtMaxZ.Size = New System.Drawing.Size(100, 20)
        Me.txtMaxZ.TabIndex = 17
        Me.txtMaxZ.Text = "0"
        '
        'Label10
        '
        Me.Label10.Location = New System.Drawing.Point(696, 256)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(56, 24)
        Me.Label10.TabIndex = 20
        Me.Label10.Text = "Mid X"
        Me.Label10.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtMidX
        '
        Me.txtMidX.Location = New System.Drawing.Point(768, 256)
        Me.txtMidX.Name = "txtMidX"
        Me.txtMidX.ReadOnly = True
        Me.txtMidX.Size = New System.Drawing.Size(100, 20)
        Me.txtMidX.TabIndex = 19
        Me.txtMidX.Text = "0"
        '
        'Label11
        '
        Me.Label11.Location = New System.Drawing.Point(696, 280)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(56, 24)
        Me.Label11.TabIndex = 22
        Me.Label11.Text = "Mid Y"
        Me.Label11.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtMidY
        '
        Me.txtMidY.Location = New System.Drawing.Point(768, 280)
        Me.txtMidY.Name = "txtMidY"
        Me.txtMidY.ReadOnly = True
        Me.txtMidY.Size = New System.Drawing.Size(100, 20)
        Me.txtMidY.TabIndex = 21
        Me.txtMidY.Text = "0"
        '
        'Label12
        '
        Me.Label12.Location = New System.Drawing.Point(696, 304)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(56, 24)
        Me.Label12.TabIndex = 24
        Me.Label12.Text = "Mid Z"
        Me.Label12.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtMidZ
        '
        Me.txtMidZ.Location = New System.Drawing.Point(768, 304)
        Me.txtMidZ.Name = "txtMidZ"
        Me.txtMidZ.ReadOnly = True
        Me.txtMidZ.Size = New System.Drawing.Size(100, 20)
        Me.txtMidZ.TabIndex = 23
        Me.txtMidZ.Text = "0"
        '
        'Label13
        '
        Me.Label13.Location = New System.Drawing.Point(878, 10)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(56, 24)
        Me.Label13.TabIndex = 26
        Me.Label13.Text = "Diffuse"
        Me.Label13.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label14
        '
        Me.Label14.Location = New System.Drawing.Point(878, 34)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(56, 24)
        Me.Label14.TabIndex = 28
        Me.Label14.Text = "Emissive"
        Me.Label14.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label15
        '
        Me.Label15.Location = New System.Drawing.Point(878, 58)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(56, 24)
        Me.Label15.TabIndex = 30
        Me.Label15.Text = "Specular"
        Me.Label15.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'cmdDiffuse
        '
        Me.cmdDiffuse.Location = New System.Drawing.Point(950, 10)
        Me.cmdDiffuse.Name = "cmdDiffuse"
        Me.cmdDiffuse.Size = New System.Drawing.Size(96, 23)
        Me.cmdDiffuse.TabIndex = 31
        '
        'cmdEmissive
        '
        Me.cmdEmissive.Location = New System.Drawing.Point(950, 34)
        Me.cmdEmissive.Name = "cmdEmissive"
        Me.cmdEmissive.Size = New System.Drawing.Size(96, 23)
        Me.cmdEmissive.TabIndex = 32
        '
        'cmdSpecular
        '
        Me.cmdSpecular.Location = New System.Drawing.Point(950, 58)
        Me.cmdSpecular.Name = "cmdSpecular"
        Me.cmdSpecular.Size = New System.Drawing.Size(49, 23)
        Me.cmdSpecular.TabIndex = 33
        '
        'Label16
        '
        Me.Label16.Location = New System.Drawing.Point(878, 162)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(56, 24)
        Me.Label16.TabIndex = 39
        Me.Label16.Text = "Scale Z"
        Me.Label16.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtScaleZ
        '
        Me.txtScaleZ.Location = New System.Drawing.Point(950, 162)
        Me.txtScaleZ.Name = "txtScaleZ"
        Me.txtScaleZ.Size = New System.Drawing.Size(100, 20)
        Me.txtScaleZ.TabIndex = 38
        Me.txtScaleZ.Text = "1"
        '
        'Label17
        '
        Me.Label17.Location = New System.Drawing.Point(878, 138)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(56, 24)
        Me.Label17.TabIndex = 37
        Me.Label17.Text = "Scale Y"
        Me.Label17.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtScaleY
        '
        Me.txtScaleY.Location = New System.Drawing.Point(950, 138)
        Me.txtScaleY.Name = "txtScaleY"
        Me.txtScaleY.Size = New System.Drawing.Size(100, 20)
        Me.txtScaleY.TabIndex = 36
        Me.txtScaleY.Text = "1"
        '
        'Label18
        '
        Me.Label18.Location = New System.Drawing.Point(878, 114)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(56, 24)
        Me.Label18.TabIndex = 35
        Me.Label18.Text = "Scale X"
        Me.Label18.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtScaleX
        '
        Me.txtScaleX.Location = New System.Drawing.Point(950, 114)
        Me.txtScaleX.Name = "txtScaleX"
        Me.txtScaleX.Size = New System.Drawing.Size(100, 20)
        Me.txtScaleX.TabIndex = 34
        Me.txtScaleX.Text = "1"
        '
        'Label19
        '
        Me.Label19.Location = New System.Drawing.Point(878, 262)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(56, 24)
        Me.Label19.TabIndex = 45
        Me.Label19.Text = "Shift Z"
        Me.Label19.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtShiftZ
        '
        Me.txtShiftZ.Location = New System.Drawing.Point(950, 262)
        Me.txtShiftZ.Name = "txtShiftZ"
        Me.txtShiftZ.Size = New System.Drawing.Size(100, 20)
        Me.txtShiftZ.TabIndex = 44
        Me.txtShiftZ.Text = "0"
        '
        'Label20
        '
        Me.Label20.Location = New System.Drawing.Point(878, 238)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(56, 24)
        Me.Label20.TabIndex = 43
        Me.Label20.Text = "Shift Y"
        Me.Label20.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtShiftY
        '
        Me.txtShiftY.Location = New System.Drawing.Point(950, 238)
        Me.txtShiftY.Name = "txtShiftY"
        Me.txtShiftY.Size = New System.Drawing.Size(100, 20)
        Me.txtShiftY.TabIndex = 42
        Me.txtShiftY.Text = "0"
        '
        'Label21
        '
        Me.Label21.Location = New System.Drawing.Point(878, 214)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(56, 24)
        Me.Label21.TabIndex = 41
        Me.Label21.Text = "Shift X"
        Me.Label21.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtShiftX
        '
        Me.txtShiftX.Location = New System.Drawing.Point(950, 214)
        Me.txtShiftX.Name = "txtShiftX"
        Me.txtShiftX.Size = New System.Drawing.Size(100, 20)
        Me.txtShiftX.TabIndex = 40
        Me.txtShiftX.Text = "0"
        '
        'cmdScale
        '
        Me.cmdScale.Location = New System.Drawing.Point(966, 186)
        Me.cmdScale.Name = "cmdScale"
        Me.cmdScale.Size = New System.Drawing.Size(75, 23)
        Me.cmdScale.TabIndex = 39
        Me.cmdScale.Text = "Scale"
        '
        'cmdShift
        '
        Me.cmdShift.Location = New System.Drawing.Point(966, 286)
        Me.cmdShift.Name = "cmdShift"
        Me.cmdShift.Size = New System.Drawing.Size(75, 23)
        Me.cmdShift.TabIndex = 47
        Me.cmdShift.Text = "Shift"
        '
        'cmdReloadTexture
        '
        Me.cmdReloadTexture.Location = New System.Drawing.Point(692, 659)
        Me.cmdReloadTexture.Name = "cmdReloadTexture"
        Me.cmdReloadTexture.Size = New System.Drawing.Size(184, 23)
        Me.cmdReloadTexture.TabIndex = 48
        Me.cmdReloadTexture.Text = "Reload Texture"
        '
        'optNeutral
        '
        Me.optNeutral.Checked = True
        Me.optNeutral.Location = New System.Drawing.Point(8, 688)
        Me.optNeutral.Name = "optNeutral"
        Me.optNeutral.Size = New System.Drawing.Size(64, 16)
        Me.optNeutral.TabIndex = 49
        Me.optNeutral.TabStop = True
        Me.optNeutral.Text = "Neutral"
        '
        'optMine
        '
        Me.optMine.Location = New System.Drawing.Point(80, 688)
        Me.optMine.Name = "optMine"
        Me.optMine.Size = New System.Drawing.Size(64, 16)
        Me.optMine.TabIndex = 50
        Me.optMine.Text = "Player"
        '
        'optAlly
        '
        Me.optAlly.Location = New System.Drawing.Point(152, 688)
        Me.optAlly.Name = "optAlly"
        Me.optAlly.Size = New System.Drawing.Size(64, 16)
        Me.optAlly.TabIndex = 51
        Me.optAlly.Text = "Ally"
        '
        'optEnemy
        '
        Me.optEnemy.Location = New System.Drawing.Point(224, 688)
        Me.optEnemy.Name = "optEnemy"
        Me.optEnemy.Size = New System.Drawing.Size(64, 16)
        Me.optEnemy.TabIndex = 52
        Me.optEnemy.Text = "Enemy"
        '
        'tbctrlMain
        '
        Me.tbctrlMain.Controls.Add(Me.tbp1)
        Me.tbctrlMain.Controls.Add(Me.tbp2)
        Me.tbctrlMain.Controls.Add(Me.tbp3)
        Me.tbctrlMain.Controls.Add(Me.tbp4)
        Me.tbctrlMain.Controls.Add(Me.tbp5)
        Me.tbctrlMain.Controls.Add(Me.tbp6)
        Me.tbctrlMain.Location = New System.Drawing.Point(692, 340)
        Me.tbctrlMain.Name = "tbctrlMain"
        Me.tbctrlMain.SelectedIndex = 0
        Me.tbctrlMain.Size = New System.Drawing.Size(356, 312)
        Me.tbctrlMain.TabIndex = 55
        '
        'tbp1
        '
        Me.tbp1.Controls.Add(Me.txtDetails)
        Me.tbp1.Location = New System.Drawing.Point(4, 22)
        Me.tbp1.Name = "tbp1"
        Me.tbp1.Size = New System.Drawing.Size(348, 286)
        Me.tbp1.TabIndex = 0
        Me.tbp1.Text = "Mesh Details"
        Me.tbp1.UseVisualStyleBackColor = True
        '
        'txtDetails
        '
        Me.txtDetails.Location = New System.Drawing.Point(6, 6)
        Me.txtDetails.Multiline = True
        Me.txtDetails.Name = "txtDetails"
        Me.txtDetails.ReadOnly = True
        Me.txtDetails.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtDetails.Size = New System.Drawing.Size(336, 276)
        Me.txtDetails.TabIndex = 54
        '
        'tbp2
        '
        Me.tbp2.Controls.Add(Me.chkLandBased)
        Me.tbp2.Controls.Add(Me.Label30)
        Me.tbp2.Controls.Add(Me.txtTurretOffZ)
        Me.tbp2.Controls.Add(Me.fraEngineFX)
        Me.tbp2.Controls.Add(Me.fraShieldSphere)
        Me.tbp2.Controls.Add(Me.Label24)
        Me.tbp2.Controls.Add(Me.txtPlanetYAdjust)
        Me.tbp2.Controls.Add(Me.Label23)
        Me.tbp2.Controls.Add(Me.txtModelID)
        Me.tbp2.Controls.Add(Me.Label22)
        Me.tbp2.Controls.Add(Me.txtFileName)
        Me.tbp2.Location = New System.Drawing.Point(4, 22)
        Me.tbp2.Name = "tbp2"
        Me.tbp2.Size = New System.Drawing.Size(348, 286)
        Me.tbp2.TabIndex = 1
        Me.tbp2.Text = "Extended Properties"
        Me.tbp2.UseVisualStyleBackColor = True
        '
        'chkLandBased
        '
        Me.chkLandBased.Location = New System.Drawing.Point(254, 38)
        Me.chkLandBased.Name = "chkLandBased"
        Me.chkLandBased.Size = New System.Drawing.Size(86, 16)
        Me.chkLandBased.TabIndex = 58
        Me.chkLandBased.Text = "Land Based"
        '
        'Label30
        '
        Me.Label30.Location = New System.Drawing.Point(6, 244)
        Me.Label30.Name = "Label30"
        Me.Label30.Size = New System.Drawing.Size(82, 24)
        Me.Label30.TabIndex = 57
        Me.Label30.Text = "Turret Offset Z:"
        Me.Label30.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtTurretOffZ
        '
        Me.txtTurretOffZ.Location = New System.Drawing.Point(92, 246)
        Me.txtTurretOffZ.Name = "txtTurretOffZ"
        Me.txtTurretOffZ.Size = New System.Drawing.Size(40, 20)
        Me.txtTurretOffZ.TabIndex = 56
        Me.txtTurretOffZ.Text = "0"
        '
        'fraEngineFX
        '
        Me.fraEngineFX.Controls.Add(Me.Label29)
        Me.fraEngineFX.Controls.Add(Me.txtPCnt)
        Me.fraEngineFX.Controls.Add(Me.btnDeleteEngineFX)
        Me.fraEngineFX.Controls.Add(Me.btnAddEngineFX)
        Me.fraEngineFX.Controls.Add(Me.txtEngineOffZ)
        Me.fraEngineFX.Controls.Add(Me.txtEngineOffY)
        Me.fraEngineFX.Controls.Add(Me.Label27)
        Me.fraEngineFX.Controls.Add(Me.txtEngineOffX)
        Me.fraEngineFX.Controls.Add(Me.lstEngineFX)
        Me.fraEngineFX.Location = New System.Drawing.Point(4, 144)
        Me.fraEngineFX.Name = "fraEngineFX"
        Me.fraEngineFX.Size = New System.Drawing.Size(332, 94)
        Me.fraEngineFX.TabIndex = 55
        Me.fraEngineFX.TabStop = False
        Me.fraEngineFX.Text = "Engine FX Details"
        '
        'Label29
        '
        Me.Label29.Location = New System.Drawing.Point(264, 18)
        Me.Label29.Name = "Label29"
        Me.Label29.Size = New System.Drawing.Size(53, 24)
        Me.Label29.TabIndex = 63
        Me.Label29.Text = "Particles:"
        Me.Label29.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtPCnt
        '
        Me.txtPCnt.Location = New System.Drawing.Point(274, 42)
        Me.txtPCnt.Name = "txtPCnt"
        Me.txtPCnt.Size = New System.Drawing.Size(40, 20)
        Me.txtPCnt.TabIndex = 62
        Me.txtPCnt.Text = "30"
        '
        'btnDeleteEngineFX
        '
        Me.btnDeleteEngineFX.Location = New System.Drawing.Point(220, 66)
        Me.btnDeleteEngineFX.Name = "btnDeleteEngineFX"
        Me.btnDeleteEngineFX.Size = New System.Drawing.Size(90, 22)
        Me.btnDeleteEngineFX.TabIndex = 61
        Me.btnDeleteEngineFX.Text = "Delete"
        '
        'btnAddEngineFX
        '
        Me.btnAddEngineFX.Location = New System.Drawing.Point(116, 66)
        Me.btnAddEngineFX.Name = "btnAddEngineFX"
        Me.btnAddEngineFX.Size = New System.Drawing.Size(90, 22)
        Me.btnAddEngineFX.TabIndex = 60
        Me.btnAddEngineFX.Text = "Add"
        '
        'txtEngineOffZ
        '
        Me.txtEngineOffZ.Location = New System.Drawing.Point(212, 42)
        Me.txtEngineOffZ.Name = "txtEngineOffZ"
        Me.txtEngineOffZ.Size = New System.Drawing.Size(40, 20)
        Me.txtEngineOffZ.TabIndex = 59
        Me.txtEngineOffZ.Text = "0"
        '
        'txtEngineOffY
        '
        Me.txtEngineOffY.Location = New System.Drawing.Point(168, 42)
        Me.txtEngineOffY.Name = "txtEngineOffY"
        Me.txtEngineOffY.Size = New System.Drawing.Size(40, 20)
        Me.txtEngineOffY.TabIndex = 58
        Me.txtEngineOffY.Text = "0"
        '
        'Label27
        '
        Me.Label27.Location = New System.Drawing.Point(114, 18)
        Me.Label27.Name = "Label27"
        Me.Label27.Size = New System.Drawing.Size(122, 24)
        Me.Label27.TabIndex = 57
        Me.Label27.Text = "Model Offset (X, Y, Z):"
        Me.Label27.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtEngineOffX
        '
        Me.txtEngineOffX.Location = New System.Drawing.Point(124, 42)
        Me.txtEngineOffX.Name = "txtEngineOffX"
        Me.txtEngineOffX.Size = New System.Drawing.Size(40, 20)
        Me.txtEngineOffX.TabIndex = 56
        Me.txtEngineOffX.Text = "0"
        '
        'lstEngineFX
        '
        Me.lstEngineFX.Location = New System.Drawing.Point(8, 18)
        Me.lstEngineFX.Name = "lstEngineFX"
        Me.lstEngineFX.Size = New System.Drawing.Size(98, 69)
        Me.lstEngineFX.TabIndex = 0
        '
        'fraShieldSphere
        '
        Me.fraShieldSphere.Controls.Add(Me.txtShieldShiftZ)
        Me.fraShieldSphere.Controls.Add(Me.txtShieldShiftY)
        Me.fraShieldSphere.Controls.Add(Me.Label28)
        Me.fraShieldSphere.Controls.Add(Me.txtShieldShiftX)
        Me.fraShieldSphere.Controls.Add(Me.btnResetShieldSphere)
        Me.fraShieldSphere.Controls.Add(Me.txtShieldStretchZ)
        Me.fraShieldSphere.Controls.Add(Me.txtShieldStretchY)
        Me.fraShieldSphere.Controls.Add(Me.Label26)
        Me.fraShieldSphere.Controls.Add(Me.txtShieldStretchX)
        Me.fraShieldSphere.Controls.Add(Me.Label25)
        Me.fraShieldSphere.Controls.Add(Me.txtShieldSphereSize)
        Me.fraShieldSphere.Location = New System.Drawing.Point(4, 64)
        Me.fraShieldSphere.Name = "fraShieldSphere"
        Me.fraShieldSphere.Size = New System.Drawing.Size(332, 74)
        Me.fraShieldSphere.TabIndex = 54
        Me.fraShieldSphere.TabStop = False
        Me.fraShieldSphere.Text = "Shield Sphere Details"
        '
        'txtShieldShiftZ
        '
        Me.txtShieldShiftZ.Location = New System.Drawing.Point(182, 44)
        Me.txtShieldShiftZ.Name = "txtShieldShiftZ"
        Me.txtShieldShiftZ.Size = New System.Drawing.Size(40, 20)
        Me.txtShieldShiftZ.TabIndex = 65
        Me.txtShieldShiftZ.Text = "0"
        '
        'txtShieldShiftY
        '
        Me.txtShieldShiftY.Location = New System.Drawing.Point(138, 44)
        Me.txtShieldShiftY.Name = "txtShieldShiftY"
        Me.txtShieldShiftY.Size = New System.Drawing.Size(40, 20)
        Me.txtShieldShiftY.TabIndex = 64
        Me.txtShieldShiftY.Text = "0"
        '
        'Label28
        '
        Me.Label28.Location = New System.Drawing.Point(6, 43)
        Me.Label28.Name = "Label28"
        Me.Label28.Size = New System.Drawing.Size(88, 24)
        Me.Label28.TabIndex = 63
        Me.Label28.Text = "Shift (X, Y, Z):"
        Me.Label28.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtShieldShiftX
        '
        Me.txtShieldShiftX.Location = New System.Drawing.Point(94, 44)
        Me.txtShieldShiftX.Name = "txtShieldShiftX"
        Me.txtShieldShiftX.Size = New System.Drawing.Size(40, 20)
        Me.txtShieldShiftX.TabIndex = 62
        Me.txtShieldShiftX.Text = "0"
        '
        'btnResetShieldSphere
        '
        Me.btnResetShieldSphere.Location = New System.Drawing.Point(232, 46)
        Me.btnResetShieldSphere.Name = "btnResetShieldSphere"
        Me.btnResetShieldSphere.Size = New System.Drawing.Size(84, 22)
        Me.btnResetShieldSphere.TabIndex = 61
        Me.btnResetShieldSphere.Text = "Reset"
        '
        'txtShieldStretchZ
        '
        Me.txtShieldStretchZ.Location = New System.Drawing.Point(260, 20)
        Me.txtShieldStretchZ.Name = "txtShieldStretchZ"
        Me.txtShieldStretchZ.Size = New System.Drawing.Size(40, 20)
        Me.txtShieldStretchZ.TabIndex = 55
        Me.txtShieldStretchZ.Text = "0"
        '
        'txtShieldStretchY
        '
        Me.txtShieldStretchY.Location = New System.Drawing.Point(216, 20)
        Me.txtShieldStretchY.Name = "txtShieldStretchY"
        Me.txtShieldStretchY.Size = New System.Drawing.Size(40, 20)
        Me.txtShieldStretchY.TabIndex = 54
        Me.txtShieldStretchY.Text = "0"
        '
        'Label26
        '
        Me.Label26.Location = New System.Drawing.Point(84, 18)
        Me.Label26.Name = "Label26"
        Me.Label26.Size = New System.Drawing.Size(88, 24)
        Me.Label26.TabIndex = 53
        Me.Label26.Text = "Stretch (X, Y, Z):"
        Me.Label26.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtShieldStretchX
        '
        Me.txtShieldStretchX.Location = New System.Drawing.Point(172, 20)
        Me.txtShieldStretchX.Name = "txtShieldStretchX"
        Me.txtShieldStretchX.Size = New System.Drawing.Size(40, 20)
        Me.txtShieldStretchX.TabIndex = 52
        Me.txtShieldStretchX.Text = "0"
        '
        'Label25
        '
        Me.Label25.Location = New System.Drawing.Point(6, 18)
        Me.Label25.Name = "Label25"
        Me.Label25.Size = New System.Drawing.Size(34, 24)
        Me.Label25.TabIndex = 51
        Me.Label25.Text = "Size:"
        Me.Label25.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtShieldSphereSize
        '
        Me.txtShieldSphereSize.Location = New System.Drawing.Point(42, 20)
        Me.txtShieldSphereSize.Name = "txtShieldSphereSize"
        Me.txtShieldSphereSize.Size = New System.Drawing.Size(40, 20)
        Me.txtShieldSphereSize.TabIndex = 50
        Me.txtShieldSphereSize.Text = "0"
        '
        'Label24
        '
        Me.Label24.Location = New System.Drawing.Point(122, 34)
        Me.Label24.Name = "Label24"
        Me.Label24.Size = New System.Drawing.Size(88, 24)
        Me.Label24.TabIndex = 51
        Me.Label24.Text = "Planet Y Adjust:"
        Me.Label24.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtPlanetYAdjust
        '
        Me.txtPlanetYAdjust.Location = New System.Drawing.Point(210, 36)
        Me.txtPlanetYAdjust.Name = "txtPlanetYAdjust"
        Me.txtPlanetYAdjust.Size = New System.Drawing.Size(40, 20)
        Me.txtPlanetYAdjust.TabIndex = 50
        Me.txtPlanetYAdjust.Text = "0"
        '
        'Label23
        '
        Me.Label23.Location = New System.Drawing.Point(4, 34)
        Me.Label23.Name = "Label23"
        Me.Label23.Size = New System.Drawing.Size(66, 24)
        Me.Label23.TabIndex = 49
        Me.Label23.Text = "Model ID:"
        Me.Label23.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtModelID
        '
        Me.txtModelID.Location = New System.Drawing.Point(76, 34)
        Me.txtModelID.Name = "txtModelID"
        Me.txtModelID.Size = New System.Drawing.Size(40, 20)
        Me.txtModelID.TabIndex = 48
        Me.txtModelID.Text = "0"
        '
        'Label22
        '
        Me.Label22.Location = New System.Drawing.Point(4, 6)
        Me.Label22.Name = "Label22"
        Me.Label22.Size = New System.Drawing.Size(66, 24)
        Me.Label22.TabIndex = 47
        Me.Label22.Text = "File Name:"
        Me.Label22.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtFileName
        '
        Me.txtFileName.Location = New System.Drawing.Point(76, 6)
        Me.txtFileName.Name = "txtFileName"
        Me.txtFileName.ReadOnly = True
        Me.txtFileName.Size = New System.Drawing.Size(264, 20)
        Me.txtFileName.TabIndex = 46
        '
        'tbp3
        '
        Me.tbp3.Controls.Add(Me.GroupBox2)
        Me.tbp3.Controls.Add(Me.GroupBox1)
        Me.tbp3.Location = New System.Drawing.Point(4, 22)
        Me.tbp3.Name = "tbp3"
        Me.tbp3.Size = New System.Drawing.Size(348, 286)
        Me.tbp3.TabIndex = 2
        Me.tbp3.Text = "Fire FX"
        Me.tbp3.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.optFireFromRight)
        Me.GroupBox2.Controls.Add(Me.optFireFromRear)
        Me.GroupBox2.Controls.Add(Me.optFireFromLeft)
        Me.GroupBox2.Controls.Add(Me.optFireFromFront)
        Me.GroupBox2.Controls.Add(Me.btnDeleteFireFrom)
        Me.GroupBox2.Controls.Add(Me.btnAddFireFrom)
        Me.GroupBox2.Controls.Add(Me.Label36)
        Me.GroupBox2.Controls.Add(Me.lstFireFroms)
        Me.GroupBox2.Controls.Add(Me.txtFireFromOffX)
        Me.GroupBox2.Controls.Add(Me.txtFireFromOffY)
        Me.GroupBox2.Controls.Add(Me.txtFireFromOffZ)
        Me.GroupBox2.Location = New System.Drawing.Point(8, 163)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(328, 120)
        Me.GroupBox2.TabIndex = 57
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Fire From Locations"
        '
        'optFireFromRight
        '
        Me.optFireFromRight.Location = New System.Drawing.Point(272, 20)
        Me.optFireFromRight.Name = "optFireFromRight"
        Me.optFireFromRight.Size = New System.Drawing.Size(52, 16)
        Me.optFireFromRight.TabIndex = 71
        Me.optFireFromRight.Text = "Right"
        '
        'optFireFromRear
        '
        Me.optFireFromRear.Location = New System.Drawing.Point(220, 20)
        Me.optFireFromRear.Name = "optFireFromRear"
        Me.optFireFromRear.Size = New System.Drawing.Size(52, 16)
        Me.optFireFromRear.TabIndex = 70
        Me.optFireFromRear.Text = "Rear"
        '
        'optFireFromLeft
        '
        Me.optFireFromLeft.Location = New System.Drawing.Point(164, 20)
        Me.optFireFromLeft.Name = "optFireFromLeft"
        Me.optFireFromLeft.Size = New System.Drawing.Size(52, 16)
        Me.optFireFromLeft.TabIndex = 69
        Me.optFireFromLeft.Text = "Left"
        '
        'optFireFromFront
        '
        Me.optFireFromFront.Checked = True
        Me.optFireFromFront.Location = New System.Drawing.Point(112, 20)
        Me.optFireFromFront.Name = "optFireFromFront"
        Me.optFireFromFront.Size = New System.Drawing.Size(52, 16)
        Me.optFireFromFront.TabIndex = 68
        Me.optFireFromFront.TabStop = True
        Me.optFireFromFront.Text = "Front"
        '
        'btnDeleteFireFrom
        '
        Me.btnDeleteFireFrom.Location = New System.Drawing.Point(220, 84)
        Me.btnDeleteFireFrom.Name = "btnDeleteFireFrom"
        Me.btnDeleteFireFrom.Size = New System.Drawing.Size(90, 22)
        Me.btnDeleteFireFrom.TabIndex = 63
        Me.btnDeleteFireFrom.Text = "Delete"
        '
        'btnAddFireFrom
        '
        Me.btnAddFireFrom.Location = New System.Drawing.Point(112, 84)
        Me.btnAddFireFrom.Name = "btnAddFireFrom"
        Me.btnAddFireFrom.Size = New System.Drawing.Size(90, 22)
        Me.btnAddFireFrom.TabIndex = 62
        Me.btnAddFireFrom.Text = "Add"
        '
        'Label36
        '
        Me.Label36.Location = New System.Drawing.Point(112, 44)
        Me.Label36.Name = "Label36"
        Me.Label36.Size = New System.Drawing.Size(116, 16)
        Me.Label36.TabIndex = 3
        Me.Label36.Text = "Model Offset (X, Y, Z):"
        '
        'lstFireFroms
        '
        Me.lstFireFroms.Location = New System.Drawing.Point(8, 16)
        Me.lstFireFroms.Name = "lstFireFroms"
        Me.lstFireFroms.Size = New System.Drawing.Size(96, 95)
        Me.lstFireFroms.TabIndex = 0
        '
        'txtFireFromOffX
        '
        Me.txtFireFromOffX.Location = New System.Drawing.Point(124, 60)
        Me.txtFireFromOffX.Name = "txtFireFromOffX"
        Me.txtFireFromOffX.Size = New System.Drawing.Size(40, 20)
        Me.txtFireFromOffX.TabIndex = 56
        Me.txtFireFromOffX.Text = "0"
        '
        'txtFireFromOffY
        '
        Me.txtFireFromOffY.Location = New System.Drawing.Point(168, 60)
        Me.txtFireFromOffY.Name = "txtFireFromOffY"
        Me.txtFireFromOffY.Size = New System.Drawing.Size(40, 20)
        Me.txtFireFromOffY.TabIndex = 58
        Me.txtFireFromOffY.Text = "0"
        '
        'txtFireFromOffZ
        '
        Me.txtFireFromOffZ.Location = New System.Drawing.Point(212, 60)
        Me.txtFireFromOffZ.Name = "txtFireFromOffZ"
        Me.txtFireFromOffZ.Size = New System.Drawing.Size(40, 20)
        Me.txtFireFromOffZ.TabIndex = 59
        Me.txtFireFromOffZ.Text = "0"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.GroupBox3)
        Me.GroupBox1.Controls.Add(Me.optRightBurn)
        Me.GroupBox1.Controls.Add(Me.optRearBurn)
        Me.GroupBox1.Controls.Add(Me.optLeftBurn)
        Me.GroupBox1.Controls.Add(Me.optFrontBurn)
        Me.GroupBox1.Controls.Add(Me.Label31)
        Me.GroupBox1.Controls.Add(Me.txtBurnPCnt)
        Me.GroupBox1.Controls.Add(Me.btnDeleteBurnFX)
        Me.GroupBox1.Controls.Add(Me.btnAddBurnFX)
        Me.GroupBox1.Controls.Add(Me.txtBurnOffZ)
        Me.GroupBox1.Controls.Add(Me.txtBurnOffY)
        Me.GroupBox1.Controls.Add(Me.Label32)
        Me.GroupBox1.Controls.Add(Me.txtBurnOffX)
        Me.GroupBox1.Controls.Add(Me.lstBurn)
        Me.GroupBox1.Location = New System.Drawing.Point(8, 10)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(332, 147)
        Me.GroupBox1.TabIndex = 56
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Burn FX Details"
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.optMinusZ)
        Me.GroupBox3.Controls.Add(Me.optMinusY)
        Me.GroupBox3.Controls.Add(Me.optMinusX)
        Me.GroupBox3.Controls.Add(Me.optPosZ)
        Me.GroupBox3.Controls.Add(Me.optPosY)
        Me.GroupBox3.Controls.Add(Me.optPosX)
        Me.GroupBox3.Location = New System.Drawing.Point(112, 104)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(204, 42)
        Me.GroupBox3.TabIndex = 74
        Me.GroupBox3.TabStop = False
        '
        'optMinusZ
        '
        Me.optMinusZ.AutoSize = True
        Me.optMinusZ.BackColor = System.Drawing.Color.Transparent
        Me.optMinusZ.Location = New System.Drawing.Point(99, 25)
        Me.optMinusZ.Name = "optMinusZ"
        Me.optMinusZ.Size = New System.Drawing.Size(35, 17)
        Me.optMinusZ.TabIndex = 79
        Me.optMinusZ.TabStop = True
        Me.optMinusZ.Text = "-Z"
        Me.optMinusZ.UseVisualStyleBackColor = False
        '
        'optMinusY
        '
        Me.optMinusY.AutoSize = True
        Me.optMinusY.BackColor = System.Drawing.Color.Transparent
        Me.optMinusY.Location = New System.Drawing.Point(55, 25)
        Me.optMinusY.Name = "optMinusY"
        Me.optMinusY.Size = New System.Drawing.Size(35, 17)
        Me.optMinusY.TabIndex = 78
        Me.optMinusY.TabStop = True
        Me.optMinusY.Text = "-Y"
        Me.optMinusY.UseVisualStyleBackColor = False
        '
        'optMinusX
        '
        Me.optMinusX.AutoSize = True
        Me.optMinusX.BackColor = System.Drawing.Color.Transparent
        Me.optMinusX.Location = New System.Drawing.Point(12, 25)
        Me.optMinusX.Name = "optMinusX"
        Me.optMinusX.Size = New System.Drawing.Size(35, 17)
        Me.optMinusX.TabIndex = 77
        Me.optMinusX.TabStop = True
        Me.optMinusX.Text = "-X"
        Me.optMinusX.UseVisualStyleBackColor = False
        '
        'optPosZ
        '
        Me.optPosZ.AutoSize = True
        Me.optPosZ.BackColor = System.Drawing.Color.Transparent
        Me.optPosZ.Location = New System.Drawing.Point(99, 8)
        Me.optPosZ.Name = "optPosZ"
        Me.optPosZ.Size = New System.Drawing.Size(38, 17)
        Me.optPosZ.TabIndex = 76
        Me.optPosZ.TabStop = True
        Me.optPosZ.Text = "+Z"
        Me.optPosZ.UseVisualStyleBackColor = False
        '
        'optPosY
        '
        Me.optPosY.AutoSize = True
        Me.optPosY.BackColor = System.Drawing.Color.Transparent
        Me.optPosY.Location = New System.Drawing.Point(55, 8)
        Me.optPosY.Name = "optPosY"
        Me.optPosY.Size = New System.Drawing.Size(38, 17)
        Me.optPosY.TabIndex = 75
        Me.optPosY.TabStop = True
        Me.optPosY.Text = "+Y"
        Me.optPosY.UseVisualStyleBackColor = False
        '
        'optPosX
        '
        Me.optPosX.AutoSize = True
        Me.optPosX.BackColor = System.Drawing.Color.Transparent
        Me.optPosX.Location = New System.Drawing.Point(12, 8)
        Me.optPosX.Name = "optPosX"
        Me.optPosX.Size = New System.Drawing.Size(38, 17)
        Me.optPosX.TabIndex = 74
        Me.optPosX.TabStop = True
        Me.optPosX.Text = "+X"
        Me.optPosX.UseVisualStyleBackColor = False
        '
        'optRightBurn
        '
        Me.optRightBurn.Location = New System.Drawing.Point(275, 90)
        Me.optRightBurn.Name = "optRightBurn"
        Me.optRightBurn.Size = New System.Drawing.Size(52, 16)
        Me.optRightBurn.TabIndex = 67
        Me.optRightBurn.Text = "Right"
        '
        'optRearBurn
        '
        Me.optRearBurn.Location = New System.Drawing.Point(221, 90)
        Me.optRearBurn.Name = "optRearBurn"
        Me.optRearBurn.Size = New System.Drawing.Size(52, 16)
        Me.optRearBurn.TabIndex = 66
        Me.optRearBurn.Text = "Rear"
        '
        'optLeftBurn
        '
        Me.optLeftBurn.Location = New System.Drawing.Point(167, 90)
        Me.optLeftBurn.Name = "optLeftBurn"
        Me.optLeftBurn.Size = New System.Drawing.Size(52, 16)
        Me.optLeftBurn.TabIndex = 65
        Me.optLeftBurn.Text = "Left"
        '
        'optFrontBurn
        '
        Me.optFrontBurn.Checked = True
        Me.optFrontBurn.Location = New System.Drawing.Point(114, 90)
        Me.optFrontBurn.Name = "optFrontBurn"
        Me.optFrontBurn.Size = New System.Drawing.Size(52, 16)
        Me.optFrontBurn.TabIndex = 64
        Me.optFrontBurn.TabStop = True
        Me.optFrontBurn.Text = "Front"
        '
        'Label31
        '
        Me.Label31.Location = New System.Drawing.Point(264, 16)
        Me.Label31.Name = "Label31"
        Me.Label31.Size = New System.Drawing.Size(53, 24)
        Me.Label31.TabIndex = 63
        Me.Label31.Text = "Particles:"
        Me.Label31.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtBurnPCnt
        '
        Me.txtBurnPCnt.Location = New System.Drawing.Point(274, 40)
        Me.txtBurnPCnt.Name = "txtBurnPCnt"
        Me.txtBurnPCnt.Size = New System.Drawing.Size(40, 20)
        Me.txtBurnPCnt.TabIndex = 62
        Me.txtBurnPCnt.Text = "30"
        '
        'btnDeleteBurnFX
        '
        Me.btnDeleteBurnFX.Location = New System.Drawing.Point(220, 64)
        Me.btnDeleteBurnFX.Name = "btnDeleteBurnFX"
        Me.btnDeleteBurnFX.Size = New System.Drawing.Size(90, 22)
        Me.btnDeleteBurnFX.TabIndex = 61
        Me.btnDeleteBurnFX.Text = "Delete"
        '
        'btnAddBurnFX
        '
        Me.btnAddBurnFX.Location = New System.Drawing.Point(116, 64)
        Me.btnAddBurnFX.Name = "btnAddBurnFX"
        Me.btnAddBurnFX.Size = New System.Drawing.Size(90, 22)
        Me.btnAddBurnFX.TabIndex = 60
        Me.btnAddBurnFX.Text = "Add"
        '
        'txtBurnOffZ
        '
        Me.txtBurnOffZ.Location = New System.Drawing.Point(212, 40)
        Me.txtBurnOffZ.Name = "txtBurnOffZ"
        Me.txtBurnOffZ.Size = New System.Drawing.Size(40, 20)
        Me.txtBurnOffZ.TabIndex = 59
        Me.txtBurnOffZ.Text = "0"
        '
        'txtBurnOffY
        '
        Me.txtBurnOffY.Location = New System.Drawing.Point(168, 40)
        Me.txtBurnOffY.Name = "txtBurnOffY"
        Me.txtBurnOffY.Size = New System.Drawing.Size(40, 20)
        Me.txtBurnOffY.TabIndex = 58
        Me.txtBurnOffY.Text = "0"
        '
        'Label32
        '
        Me.Label32.Location = New System.Drawing.Point(114, 16)
        Me.Label32.Name = "Label32"
        Me.Label32.Size = New System.Drawing.Size(122, 24)
        Me.Label32.TabIndex = 57
        Me.Label32.Text = "Model Offset (X, Y, Z):"
        Me.Label32.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtBurnOffX
        '
        Me.txtBurnOffX.Location = New System.Drawing.Point(124, 40)
        Me.txtBurnOffX.Name = "txtBurnOffX"
        Me.txtBurnOffX.Size = New System.Drawing.Size(40, 20)
        Me.txtBurnOffX.TabIndex = 56
        Me.txtBurnOffX.Text = "0"
        '
        'lstBurn
        '
        Me.lstBurn.Location = New System.Drawing.Point(8, 18)
        Me.lstBurn.Name = "lstBurn"
        Me.lstBurn.Size = New System.Drawing.Size(98, 108)
        Me.lstBurn.TabIndex = 0
        '
        'tbp4
        '
        Me.tbp4.Controls.Add(Me.dgvBL)
        Me.tbp4.Location = New System.Drawing.Point(4, 22)
        Me.tbp4.Name = "tbp4"
        Me.tbp4.Padding = New System.Windows.Forms.Padding(3)
        Me.tbp4.Size = New System.Drawing.Size(348, 286)
        Me.tbp4.TabIndex = 3
        Me.tbp4.Text = "Blinking Lights"
        Me.tbp4.UseVisualStyleBackColor = True
        '
        'dgvBL
        '
        Me.dgvBL.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvBL.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.colOffsetX, Me.colOffsetY, Me.colOffsetZ, Me.colColorR, Me.colColorG, Me.colColorB, Me.colDecayRate})
        Me.dgvBL.Location = New System.Drawing.Point(6, 6)
        Me.dgvBL.Name = "dgvBL"
        Me.dgvBL.Size = New System.Drawing.Size(336, 274)
        Me.dgvBL.TabIndex = 0
        '
        'colOffsetX
        '
        Me.colOffsetX.HeaderText = "OffX"
        Me.colOffsetX.Name = "colOffsetX"
        Me.colOffsetX.Width = 50
        '
        'colOffsetY
        '
        Me.colOffsetY.HeaderText = "OffY"
        Me.colOffsetY.Name = "colOffsetY"
        Me.colOffsetY.Width = 50
        '
        'colOffsetZ
        '
        Me.colOffsetZ.HeaderText = "OffZ"
        Me.colOffsetZ.Name = "colOffsetZ"
        Me.colOffsetZ.Width = 50
        '
        'colColorR
        '
        Me.colColorR.HeaderText = "Red"
        Me.colColorR.Name = "colColorR"
        Me.colColorR.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.colColorR.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.colColorR.Width = 50
        '
        'colColorG
        '
        Me.colColorG.HeaderText = "Green"
        Me.colColorG.Name = "colColorG"
        Me.colColorG.Width = 50
        '
        'colColorB
        '
        Me.colColorB.HeaderText = "Blue"
        Me.colColorB.Name = "colColorB"
        Me.colColorB.Width = 50
        '
        'colDecayRate
        '
        Me.colDecayRate.HeaderText = "Decay Rate"
        Me.colDecayRate.Name = "colDecayRate"
        Me.colDecayRate.Width = 50
        '
        'tbp5
        '
        Me.tbp5.Controls.Add(Me.Label41)
        Me.tbp5.Controls.Add(Me.btnRemoveSFX)
        Me.tbp5.Controls.Add(Me.btnAddSFX)
        Me.tbp5.Controls.Add(Me.lstSFX)
        Me.tbp5.Controls.Add(Me.txtSoundEffect)
        Me.tbp5.Controls.Add(Me.Label39)
        Me.tbp5.Controls.Add(Me.txtDS_ExpCnt)
        Me.tbp5.Controls.Add(Me.chkFinale)
        Me.tbp5.Controls.Add(Me.Label35)
        Me.tbp5.Controls.Add(Me.txtDS_SizeZ)
        Me.tbp5.Controls.Add(Me.Label37)
        Me.tbp5.Controls.Add(Me.txtDS_SizeY)
        Me.tbp5.Controls.Add(Me.Label38)
        Me.tbp5.Controls.Add(Me.txtDS_SizeX)
        Me.tbp5.Location = New System.Drawing.Point(4, 22)
        Me.tbp5.Name = "tbp5"
        Me.tbp5.Padding = New System.Windows.Forms.Padding(3)
        Me.tbp5.Size = New System.Drawing.Size(348, 286)
        Me.tbp5.TabIndex = 4
        Me.tbp5.Text = "Death Sequence"
        Me.tbp5.UseVisualStyleBackColor = True
        '
        'Label41
        '
        Me.Label41.Location = New System.Drawing.Point(17, 126)
        Me.Label41.Name = "Label41"
        Me.Label41.Size = New System.Drawing.Size(56, 24)
        Me.Label41.TabIndex = 20
        Me.Label41.Text = "Sound FX"
        Me.Label41.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'btnRemoveSFX
        '
        Me.btnRemoveSFX.Location = New System.Drawing.Point(266, 184)
        Me.btnRemoveSFX.Name = "btnRemoveSFX"
        Me.btnRemoveSFX.Size = New System.Drawing.Size(73, 23)
        Me.btnRemoveSFX.TabIndex = 19
        Me.btnRemoveSFX.Text = "Remove"
        Me.btnRemoveSFX.UseVisualStyleBackColor = True
        '
        'btnAddSFX
        '
        Me.btnAddSFX.Location = New System.Drawing.Point(266, 155)
        Me.btnAddSFX.Name = "btnAddSFX"
        Me.btnAddSFX.Size = New System.Drawing.Size(73, 23)
        Me.btnAddSFX.TabIndex = 18
        Me.btnAddSFX.Text = "Add"
        Me.btnAddSFX.UseVisualStyleBackColor = True
        '
        'lstSFX
        '
        Me.lstSFX.FormattingEnabled = True
        Me.lstSFX.Location = New System.Drawing.Point(20, 155)
        Me.lstSFX.Name = "lstSFX"
        Me.lstSFX.Size = New System.Drawing.Size(240, 121)
        Me.lstSFX.TabIndex = 17
        '
        'txtSoundEffect
        '
        Me.txtSoundEffect.Location = New System.Drawing.Point(89, 129)
        Me.txtSoundEffect.Name = "txtSoundEffect"
        Me.txtSoundEffect.Size = New System.Drawing.Size(253, 20)
        Me.txtSoundEffect.TabIndex = 16
        '
        'Label39
        '
        Me.Label39.Location = New System.Drawing.Point(17, 103)
        Me.Label39.Name = "Label39"
        Me.Label39.Size = New System.Drawing.Size(56, 24)
        Me.Label39.TabIndex = 15
        Me.Label39.Text = "Expl Cnt:"
        Me.Label39.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtDS_ExpCnt
        '
        Me.txtDS_ExpCnt.Location = New System.Drawing.Point(89, 103)
        Me.txtDS_ExpCnt.Name = "txtDS_ExpCnt"
        Me.txtDS_ExpCnt.Size = New System.Drawing.Size(100, 20)
        Me.txtDS_ExpCnt.TabIndex = 14
        Me.txtDS_ExpCnt.Text = "0"
        '
        'chkFinale
        '
        Me.chkFinale.AutoSize = True
        Me.chkFinale.Location = New System.Drawing.Point(89, 80)
        Me.chkFinale.Name = "chkFinale"
        Me.chkFinale.Size = New System.Drawing.Size(72, 17)
        Me.chkFinale.TabIndex = 13
        Me.chkFinale.Text = "Big Finale"
        Me.chkFinale.UseVisualStyleBackColor = True
        '
        'Label35
        '
        Me.Label35.Location = New System.Drawing.Point(17, 54)
        Me.Label35.Name = "Label35"
        Me.Label35.Size = New System.Drawing.Size(56, 24)
        Me.Label35.TabIndex = 12
        Me.Label35.Text = "Size Z"
        Me.Label35.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtDS_SizeZ
        '
        Me.txtDS_SizeZ.Location = New System.Drawing.Point(89, 54)
        Me.txtDS_SizeZ.Name = "txtDS_SizeZ"
        Me.txtDS_SizeZ.Size = New System.Drawing.Size(100, 20)
        Me.txtDS_SizeZ.TabIndex = 11
        Me.txtDS_SizeZ.Text = "0"
        '
        'Label37
        '
        Me.Label37.Location = New System.Drawing.Point(17, 30)
        Me.Label37.Name = "Label37"
        Me.Label37.Size = New System.Drawing.Size(56, 24)
        Me.Label37.TabIndex = 10
        Me.Label37.Text = "Size Y"
        Me.Label37.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtDS_SizeY
        '
        Me.txtDS_SizeY.Location = New System.Drawing.Point(89, 30)
        Me.txtDS_SizeY.Name = "txtDS_SizeY"
        Me.txtDS_SizeY.Size = New System.Drawing.Size(100, 20)
        Me.txtDS_SizeY.TabIndex = 9
        Me.txtDS_SizeY.Text = "0"
        '
        'Label38
        '
        Me.Label38.Location = New System.Drawing.Point(17, 6)
        Me.Label38.Name = "Label38"
        Me.Label38.Size = New System.Drawing.Size(56, 24)
        Me.Label38.TabIndex = 8
        Me.Label38.Text = "Size X"
        Me.Label38.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtDS_SizeX
        '
        Me.txtDS_SizeX.Location = New System.Drawing.Point(89, 6)
        Me.txtDS_SizeX.Name = "txtDS_SizeX"
        Me.txtDS_SizeX.Size = New System.Drawing.Size(100, 20)
        Me.txtDS_SizeX.TabIndex = 7
        Me.txtDS_SizeX.Text = "0"
        '
        'tbp6
        '
        Me.tbp6.Controls.Add(Me.dgvVerts)
        Me.tbp6.Location = New System.Drawing.Point(4, 22)
        Me.tbp6.Name = "tbp6"
        Me.tbp6.Padding = New System.Windows.Forms.Padding(3)
        Me.tbp6.Size = New System.Drawing.Size(348, 286)
        Me.tbp6.TabIndex = 5
        Me.tbp6.Text = "Vertices"
        Me.tbp6.UseVisualStyleBackColor = True
        '
        'dgvVerts
        '
        Me.dgvVerts.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvVerts.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.DataGridViewTextBoxColumn1, Me.DataGridViewTextBoxColumn2, Me.DataGridViewTextBoxColumn3, Me.DataGridViewTextBoxColumn4, Me.DataGridViewTextBoxColumn5})
        Me.dgvVerts.Location = New System.Drawing.Point(6, 6)
        Me.dgvVerts.Name = "dgvVerts"
        Me.dgvVerts.Size = New System.Drawing.Size(336, 274)
        Me.dgvVerts.TabIndex = 1
        '
        'DataGridViewTextBoxColumn1
        '
        Me.DataGridViewTextBoxColumn1.HeaderText = "LocX"
        Me.DataGridViewTextBoxColumn1.Name = "DataGridViewTextBoxColumn1"
        Me.DataGridViewTextBoxColumn1.Width = 50
        '
        'DataGridViewTextBoxColumn2
        '
        Me.DataGridViewTextBoxColumn2.HeaderText = "LocY"
        Me.DataGridViewTextBoxColumn2.Name = "DataGridViewTextBoxColumn2"
        Me.DataGridViewTextBoxColumn2.Width = 50
        '
        'DataGridViewTextBoxColumn3
        '
        Me.DataGridViewTextBoxColumn3.HeaderText = "LocZ"
        Me.DataGridViewTextBoxColumn3.Name = "DataGridViewTextBoxColumn3"
        Me.DataGridViewTextBoxColumn3.Width = 50
        '
        'DataGridViewTextBoxColumn4
        '
        Me.DataGridViewTextBoxColumn4.HeaderText = "tu"
        Me.DataGridViewTextBoxColumn4.Name = "DataGridViewTextBoxColumn4"
        Me.DataGridViewTextBoxColumn4.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DataGridViewTextBoxColumn4.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.DataGridViewTextBoxColumn4.Width = 50
        '
        'DataGridViewTextBoxColumn5
        '
        Me.DataGridViewTextBoxColumn5.HeaderText = "tv"
        Me.DataGridViewTextBoxColumn5.Name = "DataGridViewTextBoxColumn5"
        Me.DataGridViewTextBoxColumn5.Width = 50
        '
        'btnInvertZ
        '
        Me.btnInvertZ.Location = New System.Drawing.Point(898, 314)
        Me.btnInvertZ.Name = "btnInvertZ"
        Me.btnInvertZ.Size = New System.Drawing.Size(144, 23)
        Me.btnInvertZ.TabIndex = 56
        Me.btnInvertZ.Text = "Invert Z Axis"
        '
        'btnReverseWind
        '
        Me.btnReverseWind.Location = New System.Drawing.Point(910, 659)
        Me.btnReverseWind.Name = "btnReverseWind"
        Me.btnReverseWind.Size = New System.Drawing.Size(136, 23)
        Me.btnReverseWind.TabIndex = 57
        Me.btnReverseWind.Text = "Reverse Wind"
        '
        'lblRngOff
        '
        Me.lblRngOff.Location = New System.Drawing.Point(752, 79)
        Me.lblRngOff.Name = "lblRngOff"
        Me.lblRngOff.Size = New System.Drawing.Size(116, 15)
        Me.lblRngOff.TabIndex = 59
        Me.lblRngOff.Text = "Range offset:"
        '
        'hscrYaw
        '
        Me.hscrYaw.Location = New System.Drawing.Point(352, 688)
        Me.hscrYaw.Maximum = 45
        Me.hscrYaw.Minimum = -45
        Me.hscrYaw.Name = "hscrYaw"
        Me.hscrYaw.Size = New System.Drawing.Size(80, 16)
        Me.hscrYaw.TabIndex = 60
        '
        'hscrRotate
        '
        Me.hscrRotate.Location = New System.Drawing.Point(520, 688)
        Me.hscrRotate.Maximum = 3599
        Me.hscrRotate.Name = "hscrRotate"
        Me.hscrRotate.Size = New System.Drawing.Size(80, 16)
        Me.hscrRotate.TabIndex = 61
        '
        'Label33
        '
        Me.Label33.Location = New System.Drawing.Point(314, 688)
        Me.Label33.Name = "Label33"
        Me.Label33.Size = New System.Drawing.Size(34, 14)
        Me.Label33.TabIndex = 62
        Me.Label33.Text = "YAW:"
        '
        'Label34
        '
        Me.Label34.Location = New System.Drawing.Point(462, 688)
        Me.Label34.Name = "Label34"
        Me.Label34.Size = New System.Drawing.Size(54, 14)
        Me.Label34.TabIndex = 63
        Me.Label34.Text = "ROTATE:"
        '
        'btnInvertNormals
        '
        Me.btnInvertNormals.Location = New System.Drawing.Point(750, 685)
        Me.btnInvertNormals.Name = "btnInvertNormals"
        Me.btnInvertNormals.Size = New System.Drawing.Size(93, 22)
        Me.btnInvertNormals.TabIndex = 64
        Me.btnInvertNormals.Text = "Invert Normals"
        Me.btnInvertNormals.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(849, 685)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(95, 22)
        Me.Button1.TabIndex = 65
        Me.Button1.Text = "Hull Size"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'btnResetCamera
        '
        Me.btnResetCamera.Location = New System.Drawing.Point(950, 685)
        Me.btnResetCamera.Name = "btnResetCamera"
        Me.btnResetCamera.Size = New System.Drawing.Size(96, 22)
        Me.btnResetCamera.TabIndex = 66
        Me.btnResetCamera.Text = "Reset Camera"
        Me.btnResetCamera.UseVisualStyleBackColor = True
        '
        'hscrPMesh
        '
        Me.hscrPMesh.Location = New System.Drawing.Point(609, 688)
        Me.hscrPMesh.Name = "hscrPMesh"
        Me.hscrPMesh.Size = New System.Drawing.Size(71, 16)
        Me.hscrPMesh.TabIndex = 67
        '
        'btnPMesh
        '
        Me.btnPMesh.Location = New System.Drawing.Point(683, 685)
        Me.btnPMesh.Name = "btnPMesh"
        Me.btnPMesh.Size = New System.Drawing.Size(61, 22)
        Me.btnPMesh.TabIndex = 68
        Me.btnPMesh.Text = "PMesh"
        Me.btnPMesh.UseVisualStyleBackColor = True
        '
        'Label40
        '
        Me.Label40.Location = New System.Drawing.Point(878, 82)
        Me.Label40.Name = "Label40"
        Me.Label40.Size = New System.Drawing.Size(78, 29)
        Me.Label40.TabIndex = 69
        Me.Label40.Text = "Shields / Engines"
        Me.Label40.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'btnShieldColor
        '
        Me.btnShieldColor.Location = New System.Drawing.Point(950, 82)
        Me.btnShieldColor.Name = "btnShieldColor"
        Me.btnShieldColor.Size = New System.Drawing.Size(49, 23)
        Me.btnShieldColor.TabIndex = 70
        '
        'btnEngineColor
        '
        Me.btnEngineColor.Location = New System.Drawing.Point(1005, 82)
        Me.btnEngineColor.Name = "btnEngineColor"
        Me.btnEngineColor.Size = New System.Drawing.Size(41, 23)
        Me.btnEngineColor.TabIndex = 71
        '
        'txtSpecSharp
        '
        Me.txtSpecSharp.Location = New System.Drawing.Point(1010, 59)
        Me.txtSpecSharp.Name = "txtSpecSharp"
        Me.txtSpecSharp.Size = New System.Drawing.Size(36, 20)
        Me.txtSpecSharp.TabIndex = 72
        Me.txtSpecSharp.Text = "10"
        '
        'cboBurn
        '
        Me.cboBurn.FormattingEnabled = True
        Me.cboBurn.Location = New System.Drawing.Point(62, 658)
        Me.cboBurn.Name = "cboBurn"
        Me.cboBurn.Size = New System.Drawing.Size(121, 21)
        Me.cboBurn.TabIndex = 73
        '
        'cboDS
        '
        Me.cboDS.FormattingEnabled = True
        Me.cboDS.Location = New System.Drawing.Point(189, 658)
        Me.cboDS.Name = "cboDS"
        Me.cboDS.Size = New System.Drawing.Size(121, 21)
        Me.cboDS.TabIndex = 74
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(1060, 710)
        Me.Controls.Add(Me.cboDS)
        Me.Controls.Add(Me.cboBurn)
        Me.Controls.Add(Me.txtSpecSharp)
        Me.Controls.Add(Me.btnEngineColor)
        Me.Controls.Add(Me.btnShieldColor)
        Me.Controls.Add(Me.Label40)
        Me.Controls.Add(Me.btnPMesh)
        Me.Controls.Add(Me.hscrPMesh)
        Me.Controls.Add(Me.btnResetCamera)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.btnInvertNormals)
        Me.Controls.Add(Me.Label34)
        Me.Controls.Add(Me.Label33)
        Me.Controls.Add(Me.hscrRotate)
        Me.Controls.Add(Me.hscrYaw)
        Me.Controls.Add(Me.lblRngOff)
        Me.Controls.Add(Me.btnReverseWind)
        Me.Controls.Add(Me.btnInvertZ)
        Me.Controls.Add(Me.tbctrlMain)
        Me.Controls.Add(Me.optEnemy)
        Me.Controls.Add(Me.optAlly)
        Me.Controls.Add(Me.optMine)
        Me.Controls.Add(Me.optNeutral)
        Me.Controls.Add(Me.cmdReloadTexture)
        Me.Controls.Add(Me.cmdShift)
        Me.Controls.Add(Me.cmdScale)
        Me.Controls.Add(Me.Label19)
        Me.Controls.Add(Me.txtShiftZ)
        Me.Controls.Add(Me.Label20)
        Me.Controls.Add(Me.txtShiftY)
        Me.Controls.Add(Me.Label21)
        Me.Controls.Add(Me.txtShiftX)
        Me.Controls.Add(Me.Label16)
        Me.Controls.Add(Me.txtScaleZ)
        Me.Controls.Add(Me.Label17)
        Me.Controls.Add(Me.txtScaleY)
        Me.Controls.Add(Me.Label18)
        Me.Controls.Add(Me.txtScaleX)
        Me.Controls.Add(Me.cmdSpecular)
        Me.Controls.Add(Me.cmdEmissive)
        Me.Controls.Add(Me.cmdDiffuse)
        Me.Controls.Add(Me.Label15)
        Me.Controls.Add(Me.Label14)
        Me.Controls.Add(Me.Label13)
        Me.Controls.Add(Me.Label12)
        Me.Controls.Add(Me.txtMidZ)
        Me.Controls.Add(Me.Label11)
        Me.Controls.Add(Me.txtMidY)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.txtMidX)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.txtMaxZ)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.txtMaxY)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.txtMaxX)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.txtMinZ)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.txtMinY)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.txtMinX)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.txtSizeZ)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.txtSizeY)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.txtSizeX)
        Me.Controls.Add(Me.picMain)
        Me.KeyPreview = True
        Me.Menu = Me.mnuMain
        Me.Name = "Form1"
        Me.Text = "Form1"
        CType(Me.picMain, System.ComponentModel.ISupportInitialize).EndInit()
        Me.tbctrlMain.ResumeLayout(False)
        Me.tbp1.ResumeLayout(False)
        Me.tbp1.PerformLayout()
        Me.tbp2.ResumeLayout(False)
        Me.tbp2.PerformLayout()
        Me.fraEngineFX.ResumeLayout(False)
        Me.fraEngineFX.PerformLayout()
        Me.fraShieldSphere.ResumeLayout(False)
        Me.fraShieldSphere.PerformLayout()
        Me.tbp3.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.tbp4.ResumeLayout(False)
        CType(Me.dgvBL, System.ComponentModel.ISupportInitialize).EndInit()
        Me.tbp5.ResumeLayout(False)
        Me.tbp5.PerformLayout()
        Me.tbp6.ResumeLayout(False)
        CType(Me.dgvVerts, System.ComponentModel.ISupportInitialize).EndInit()
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

	Private Const gdRadPerDegree As Double = Math.PI / 180.0#

	Private moMesh As Mesh
	Private moShieldMesh As Mesh
	Private moShieldTex(3) As Texture
	Private mlCurrShieldTex As Int32 = 0
	Private moShieldMat As Material

	Private mbRenderOldMesh As Boolean = False

	'Tanks share materials/textures across both meshes...
	Private NumOfMaterials As Int32
	Private Materials() As Material
	Private Textures() As Texture
	Private mtrlBuffer() As ExtendedMaterial

	Private moTurretMesh As Mesh
	Private mfTurretRot As Single

	Private moDevice As Device

	Private moPEngine As ParticleFX.ParticleEngine
	Private mbIgnoreEngineFXEvents As Boolean = False
	Private mbIgnoreBurnFXEvents As Boolean = False
	Private mbIgnoreFireFromEvents As Boolean = False

	Private moCompareMesh As Mesh
	Private mlCompareNumMat As Int32
	Private moCompareMat() As Material
	Private moCompareTex() As Texture

	Private mbTankMesh As Boolean = False	'indicates whether we need to load the tank and the turret (and render)

	'BURN FX LOCATIONS =====
	Private moBurnFX() As BurnFXData
	Private mlBurnFXUB As Int32 = -1
	Private myBurnFXUsed() As Byte
	'=======================

	'FIRE FROM LOCATIONS =====
	Private mcolFireFromLocs(3) As Collection
	Private mlFireFromSeq As Int32 = 0
	Private moFireFromIndicator As Mesh
	Private matFireFrom(3) As Material
	'=========================

	'BLINKING LIGHT LOCATIONS ====
	Private mvecBL_Loc() As Vector3
	Private mcolBL() As System.Drawing.Color
	Private mlBL_LocUB As Int32 = -1
	Private moBLTex As Texture
	Private mfAlpha() As Single
	Private mfAlphaChg() As Single
	'=============================

	Private mbShowMesh As Boolean = True

	Private moDeathSeq As DeathSequence

	Private mvecHitLoc As Vector3
	Private mbRenderHitloc As Boolean = False

	Private moFireTo() As Vector3

	Private mbRenderShader As Boolean = False
	Private moShader As SimpleShader
	Private moGlowFX As PostShader
	Private mbRenderGlowFX As Boolean = False

	Private mbCaptureScreenshot As Boolean = False

	Private Sub DoCaptureScreenshot()
		Dim oSurf As Surface
		Dim sFile As String = AppDomain.CurrentDomain.BaseDirectory

		If sFile.EndsWith("\") = False Then sFile = sFile & "\"

		oSurf = moDevice.GetBackBuffer(0, 0, BackBufferType.Mono)
		SurfaceLoader.Save(sFile & "SS_" & Now.ToString("MM_dd_yyyy_hhmmss") & ".bmp", ImageFileFormat.Bmp, oSurf)
		oSurf.Dispose()
		oSurf = Nothing
	End Sub

#Region " Resource Creation "
	Private Function CreateShieldSphere(ByVal fRadius As Single, ByVal lSlices As Int32, ByVal lStacks As Int32, ByVal lWrapCount As Int32, ByVal fStretchX As Single, ByVal fStretchY As Single, ByVal fStretchZ As Single, ByVal fShiftX As Single, ByVal fShiftY As Single, ByVal fShiftZ As Single) As Mesh

		Device.IsUsingEventHandlers = False

		Dim oTemp As Mesh = Mesh.Sphere(moDevice, fRadius, lSlices, lStacks)
		oTemp.ComputeNormals()
		Dim TexturedObject As Mesh = New Mesh(oTemp.NumberFaces, oTemp.NumberVertices, MeshFlags.Managed, CustomVertex.PositionNormalTextured.Format, moDevice)
		'Dim TexturedObject As Mesh = oTemp.Clone(MeshFlags.Managed, VertexFormats.Position Or VertexFormats.Normal Or VertexFormats.Texture0, moDevice)

		' Get the original mesh's vertex buffer.
		Dim ranks(0) As Integer
		ranks(0) = oTemp.NumberVertices
		Dim arr As System.Array = oTemp.VertexBuffer.Lock(0, (New CustomVertex.PositionNormal()).GetType(), LockFlags.None, ranks)

		' Set the vertex buffer
		Dim data As System.Array = TexturedObject.VertexBuffer.Lock(0, (New CustomVertex.PositionNormalTextured()).GetType(), LockFlags.None, ranks)

		Dim phi As Single
		Dim u As Single
		Dim i As Integer

		For i = 0 To arr.Length - 1
			Dim pn As Direct3D.CustomVertex.PositionNormal = CType(arr.GetValue(i), CustomVertex.PositionNormal)
			Dim pnt As Direct3D.CustomVertex.PositionNormalTextured = CType(data.GetValue(i), CustomVertex.PositionNormalTextured)
			pnt.X = pn.X + fShiftX
			pnt.Y = pn.Y + fShiftY
			pnt.Z = pn.Z + fShiftZ
			pnt.Nx = pn.Nx
			pnt.Ny = pn.Ny
			pnt.Nz = pn.Nz

			If lWrapCount = 0 Then
				phi = CSng(Math.Acos(pn.Nz))
				pnt.Tv = CSng(phi / Math.PI)
				u = CSng(Math.Acos(Math.Max(Math.Min(pnt.Ny / Math.Sin(phi), 1.0), -1.0)) / (2.0 * Math.PI))
				If pnt.Nx > 0 Then
					pnt.Tu = u
				Else
					pnt.Tu = 1 - u
				End If
			Else
				'TODO: At the moment, lWrapCount only determines if we wrap or not... I will want it to
				'  actually do more wraps with higher counts....
				pnt.Tu = CSng(Math.Asin(pnt.Nx) / Math.PI + 0.5F)
				pnt.Tv = CSng(Math.Asin(pnt.Ny) / Math.PI + 0.5F)
			End If

			'Now, do our stretch
			pnt.X *= fStretchX
			pnt.Y *= fStretchY
			pnt.Z *= fStretchZ


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
				moMesh = Mesh.FromFile(sFile, MeshFlags.Managed, moDevice, mtrlBuffer)


				'Now, because our engine uses Normals and lighting, let's get that out of the way
				If (moMesh.VertexFormat And VertexFormats.Normal) = 0 Then
					Dim oTmpMesh As Mesh = moMesh.Clone(moMesh.Options.Value, moMesh.VertexFormat Or VertexFormats.Normal, moDevice)
					oTmpMesh.ComputeNormals()
					moMesh.Dispose()
					moMesh = oTmpMesh
					oTmpMesh = Nothing
				End If

				sTemp = Mid$(sFile, InStrRev(sFile, "\") + 1)
				txtFileName.Text = sTemp

				sDetails = "File: " & sFile & vbCrLf
				sDetails &= "Vertices: " & moMesh.NumberVertices & vbCrLf
				sDetails &= "Faces: " & moMesh.NumberFaces & vbCrLf
				sDetails &= "Material Count: " & mtrlBuffer.Length & vbCrLf & vbCrLf

				NumOfMaterials = mtrlBuffer.Length
				ReDim Materials(NumOfMaterials - 1)
				ReDim Textures((NumOfMaterials * 4) - 1)
				'Now load our textures and materials
				For X = 0 To mtrlBuffer.Length - 1
					Materials(X) = mtrlBuffer(X).Material3D
					Materials(X).Ambient = Materials(X).Diffuse

					sDetails &= "Material " & (X + 1) & ":" & vbCrLf
					sDetails &= "Texture: " & mtrlBuffer(X).TextureFilename & vbCrLf
					sDetails &= "Diffuse: (" & Materials(X).Diffuse.A & ", " & Materials(X).Diffuse.R & ", " & Materials(X).Diffuse.G & ", " & Materials(X).Diffuse.B & ")" & vbCrLf
					sDetails &= "Emissive: (" & Materials(X).Emissive.A & ", " & Materials(X).Emissive.R & ", " & Materials(X).Emissive.G & ", " & Materials(X).Emissive.B & ")" & vbCrLf
					sDetails &= "Specular: (" & Materials(X).Specular.A & ", " & Materials(X).Specular.R & ", " & Materials(X).Specular.G & ", " & Materials(X).Specular.B & ")" & vbCrLf & vbCrLf
					sDetails &= "Specular Strength: " & Materials(X).SpecularSharpness & vbCrLf

					txtSpecSharp.Text = Materials(X).SpecularSharpness.ToString()

					If mtrlBuffer(X).TextureFilename <> "" Then
						sTemp = mtrlBuffer(X).TextureFilename

						'ok, this is special now
						'sTemp is the NEUTRAL version
						sPostFix = Mid$(sTemp, Len(sTemp) - 3, 4)
						'sPostFix = ".bmp"
						sTemp = Mid$(sTemp, 1, Len(sTemp) - 4)

						'now, load our images
						Textures(X) = TextureLoader.FromFile(moDevice, sTemp & sPostFix)
						Textures((X * 4) + 1) = TextureLoader.FromFile(moDevice, sTemp & "_mine" & sPostFix)
						Textures((X * 4) + 2) = TextureLoader.FromFile(moDevice, sTemp & "_ally" & sPostFix)
						Textures((X * 4) + 3) = TextureLoader.FromFile(moDevice, sTemp & "_enemy" & sPostFix)


						'Textures(X) = TextureLoader.FromFile(moDevice, sTemp)
						'Textures(X) = GetTexture(mtrlBuffer(X).TextureFilename)
					End If
				Next X
			End If
		End If

		txtDetails.Text = sDetails


	End Sub

	Private Sub LoadCompareMesh(ByVal sFile As String)
		Dim X As Int32
		Dim sTemp As String
		Dim omatBuff(-1) As ExtendedMaterial

		If sFile <> "" Then
			If Dir$(sFile) <> "" Then
				'now set it up
				moCompareMesh = Mesh.FromFile(sFile, MeshFlags.Managed, moDevice, omatBuff)

				mlCompareNumMat = omatBuff.Length
				ReDim moCompareMat(mlCompareNumMat - 1)
				ReDim moCompareTex(mlCompareNumMat - 1)
				'Now load our textures and materials
				For X = 0 To omatBuff.Length - 1
					moCompareMat(X) = omatBuff(X).Material3D
					moCompareMat(X).Ambient = moCompareMat(X).Diffuse

					If omatBuff(X).TextureFilename <> "" Then
						sTemp = omatBuff(X).TextureFilename

						'now, load our images
						moCompareTex(X) = TextureLoader.FromFile(moDevice, sTemp)
					End If
				Next X
			End If
		End If
	End Sub

#End Region

#Region " Direct3D-Specific "
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
				.AutoDepthStencilFormat = DepthFormat.D16
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

		moMesh = CreateTexturedSphere(1000, 26, 26, 0, False)

		Return bRes
	End Function

	Private Sub SetupMatrices()
		moDevice.Transform.View = Matrix.LookAtLH(New Vector3(mlCameraX, mlCameraY, mlCameraZ), _
		  New Vector3(mlCameraAtX, mlCameraAtY, mlCameraAtZ), New Vector3(0.0#, 1.0#, 0.0#))	'up is always this
		moDevice.Transform.Projection = Matrix.PerspectiveFovLH(0.7853981633974475, 1, 100, 2500000)
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
#End Region

#Region " Camera-Specific "
	Private Sub ScrollCamera()
		'handle view scrolling... however, this comes with a twist... we need to determine our angle to the target
		' deltaXx is the delta when scrolling camera's X along the X world axis
		' deltaXz is the delta when scrolling camera's X along the Z world axis
		' deltaZx is the delta when scrolling camera's Z along the X world axis
		' deltaZz is the delta when scrolling camera's Z along the Z world axis
		' which is to say, when changing X (scrolling horizontally), deltaXx is change to X and deltaXz is change to Z
		' and when changing Z (scrolling vertically), deltaZx is change to X and deltaZz is change to Z

		If mnuLockCamera.Checked = True Then Exit Sub

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

#End Region

#Region " Windows Controls Related (events) "
	Private Sub tmrRender_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles tmrRender.Tick
		DrawScene()
	End Sub

	Private Sub Form1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
		If InitD3D(picMain) = False Then End

		'Intervene()
		'Return

		mlCameraAtX = 0 : mlCameraAtY = 0 : mlCameraAtZ = 0
		mlCameraX = 0 : mlCameraY = 1000 : mlCameraZ = -1000

        tmrRender.Enabled = True

        For X As Int32 = 1 To 26
            cboBurn.Items.Add(CType(X, TextureOperation).ToString)
            cboDS.Items.Add(CType(X, TextureOperation).ToString)
            If X = TextureOperation.Modulate Then
                cboBurn.SelectedIndex = X - 1
                cboDS.SelectedIndex = X - 1
            End If
        Next

	End Sub

	Private Sub Intervene()
		Dim oTex As Texture = TextureLoader.FromFile(moDevice, "C:\Documents and Settings\Matthew Campbell\Desktop\Models\Reticle\reticle\reticle.dds", 0, 0, 1, Usage.Dynamic, Format.Unknown, Pool.Default, Filter.Linear, Filter.Linear, 0)
		Dim pitch As Int32
		Dim yDest As Byte() = CType(oTex.LockRectangle(GetType(Byte), 0, LockFlags.None, pitch, 512 * 512 * 4), Byte())
		'442, 62
		For Y As Int32 = 0 To 511
			For X As Int32 = 0 To 511
				yDest((Y * 512 + X) * 4 + 3) = CByte(yDest((Y * 512 + X) * 4 + 3) / 3)			'a
				'yDest((Y * 512 + X) * 4 + 2) = 0			'r
				'yDest((Y * 512 + X) * 4 + 1) = 0			'g
				'yDest((Y * 512 + X) * 4 + 0) = 0			'b
			Next X
		Next Y
		oTex.UnlockRectangle(0)
		Erase yDest

		SurfaceLoader.Save("C:\temp.dds", ImageFileFormat.Dds, oTex.GetSurfaceLevel(0))
	End Sub

	Private Sub Form1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
		Select Case e.KeyCode
			Case Keys.W
				moDevice.RenderState.FillMode = FillMode.WireFrame
				mnuWireFrame.Checked = True
			Case Keys.S
				moDevice.RenderState.FillMode = FillMode.Solid
				mnuWireFrame.Checked = False
			Case Keys.Escape
				mnuExit_Click(Nothing, Nothing)
			Case Keys.C
				mbCaptureScreenshot = True
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

	Private Sub optFrontBurn_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optFrontBurn.CheckedChanged
		UpdateBurnFX()
	End Sub

	Private Sub optRightBurn_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optRightBurn.CheckedChanged
		UpdateBurnFX()
	End Sub

	Private Sub optLeftBurn_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optLeftBurn.CheckedChanged
		UpdateBurnFX()
	End Sub

	Private Sub optRearBurn_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optRearBurn.CheckedChanged
		UpdateBurnFX()
	End Sub

	Private Sub optBurnDir_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optPosX.CheckedChanged, optMinusX.CheckedChanged, optPosY.CheckedChanged, optMinusY.CheckedChanged, optPosZ.CheckedChanged, optPosZ.CheckedChanged, optMinusZ.CheckedChanged
		UpdateBurnFX()
	End Sub

#Region " General User Interaction Events "
	Private Sub Form1_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles picMain.MouseMove
		Dim deltaX As Int32
		Dim deltaY As Int32

		If mbRightDown Then
			mbRightDrag = True
			deltaX = e.X - mlMouseX
			deltaY = e.Y - mlMouseY
			RotateCamera(deltaX, deltaY)
		Else
			mbScrollLeft = False : mbScrollRight = False : mbScrollUp = False : mbScrollDown = False
			If e.X < 20 Then
				mbScrollLeft = True
			ElseIf e.X > picMain.Width - 20 Then
				mbScrollRight = True
			End If
			If e.Y < 20 Then
				mbScrollUp = True
			ElseIf e.Y > picMain.Height - 20 Then
				mbScrollDown = True
			End If
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

#End Region

#Region " Menu Item Events "
	Private Sub mnuSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSave.Click
		SaveData(True, True)
	End Sub

	Private Sub mnuSaveXFileOnly_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSaveXFileOnly.Click
		'Dim oRes As DialogResult
		'Dim sFile As String
		'Dim X As Int32
		'Dim lAdj() As Int32
		'Dim sTemp As String

		'sFile = AppDomain.CurrentDomain.BaseDirectory
		'If sFile.EndsWith("\") = False Then sFile &= "\"

		''Get our file to save to
		'savDlg.Filter = "X Files *.x|*.x"
		'oRes = savDlg.ShowDialog(Me)
		'If oRes = DialogResult.OK Then
		'    'Check if the ModelID is in use already
		'    sFile = savDlg.FileName

		'    ReDim lAdj(3 * moMesh.NumberFaces)

		'    'Now, generate our adjacency
		'    moMesh.GenerateAdjacency(0, lAdj)

		'    'Now, go back and set our materials
		'    For X = 0 To mtrlBuffer.Length - 1
		'        mtrlBuffer(X).Material3D = Materials(X)
		'    Next X

		'    'save it
		'    moMesh.Save(sFile, lAdj, mtrlBuffer, XFileFormat.Binary)

		'    If mbTankMesh = True AndAlso moTurretMesh Is Nothing = False Then
		'        Dim sTurrFile As String
		'        If InStr(sTemp, "_B.X", CompareMethod.Text) <> 0 Then
		'            sTemp = Replace$(sTemp, "_B.X", "_T.X", , , CompareMethod.Text)
		'            sTurrFile = Replace$(sFile, "_B.X", "_T.X", , , CompareMethod.Text)
		'        Else
		'            sTemp = Replace$(sTemp, ".X", "_T.X", , , CompareMethod.Text)
		'            sTurrFile = Replace$(sFile, ".X", "_T.X", , , CompareMethod.Text)
		'        End If

		'        ReDim lAdj(3 * moTurretMesh.NumberFaces)
		'        moTurretMesh.GenerateAdjacency(0, lAdj)
		'        moTurretMesh.Save(sTurrFile, lAdj, mtrlBuffer, XFileFormat.Binary)
		'    End If
		'End If
		SaveData(True, False)
	End Sub

	Private Sub mnuExit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuExit.Click
		Me.Close()
	End Sub

	Private Sub mnuWireFrame_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuWireFrame.Click
		mnuWireFrame.Checked = Not mnuWireFrame.Checked
		If mnuWireFrame.Checked = True Then
			moDevice.RenderState.FillMode = FillMode.WireFrame
		Else
			moDevice.RenderState.FillMode = FillMode.Solid
		End If
	End Sub

	Private Sub mnuRenderPlane_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuRenderPlane.Click
		mnuRenderPlane.Checked = Not mnuRenderPlane.Checked
	End Sub

	Private Sub mnuRenderShield_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuRenderShield.Click
		mnuRenderShield.Checked = Not mnuRenderShield.Checked
		If mnuRenderShield.Checked = True Then btnResetShieldSphere_Click(sender, e)
	End Sub

	Private Sub mnuRenderEngineFX_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuRenderEngineFX.Click
		mnuRenderEngineFX.Checked = Not mnuRenderEngineFX.Checked
	End Sub

	Private Sub mnuLockCamera_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuLockCamera.Click
		mnuLockCamera.Checked = Not mnuLockCamera.Checked
	End Sub

	Private Sub mnuOpenCompareX_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuOpenCompareX.Click
		Dim oRes As DialogResult
		Dim sFile As String

		opnDlg.Filter = "X Files *.x|*.x"
		oRes = opnDlg.ShowDialog(Me)

		If oRes = Windows.Forms.DialogResult.OK Then
			sFile = opnDlg.FileName
			LoadCompareMesh(sFile)
		End If
	End Sub

	Private Sub mnuOpenX_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuOpenX.Click
		Dim oRes As DialogResult
		Dim sFile As String

		opnDlg.Filter = "X Files *.x|*.x"
		oRes = opnDlg.ShowDialog(Me)

		If oRes = Windows.Forms.DialogResult.OK Then
			sFile = opnDlg.FileName
			LoadMesh(sFile)
			mlCameraAtX = 0 : mlCameraAtY = 0 : mlCameraAtZ = 0
			mlCameraX = 0 : mlCameraY = 1000 : mlCameraZ = -1000

			'Now, load our extended details...
			LoadExtendedDetails(sFile)

			RefreshStatDisplay()
		End If
		mbTankMesh = False
	End Sub

	Private Sub mnuOpenTank_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuOpenTank.Click
		'Just like OpenX...
		Dim oRes As DialogResult
		Dim sFile As String = ""

		opnDlg.Filter = "X Files *.x|*.x"
		oRes = opnDlg.ShowDialog(Me)

		If oRes = Windows.Forms.DialogResult.OK Then
			sFile = opnDlg.FileName
			LoadMesh(sFile)
			mlCameraAtX = 0 : mlCameraAtY = 0 : mlCameraAtZ = 0
			mlCameraX = 0 : mlCameraY = 1000 : mlCameraZ = -1000

			RefreshStatDisplay()
		End If

		Dim sTurrFile As String = Replace$(sFile, "_B.X", "_T.X", , , CompareMethod.Text)
		'now, load that mesh independantly
		moTurretMesh = Mesh.FromFile(sTurrFile, MeshFlags.Managed, moDevice, mtrlBuffer)
		mbTankMesh = True
	End Sub

	Private Sub mnuRenderCompare_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuRenderCompare.Click
		mnuRenderCompare.Checked = Not mnuRenderCompare.Checked
	End Sub

	Private Sub mnuAtmosBurn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuAtmosBurn.Click
		mnuAtmosBurn.Checked = Not mnuAtmosBurn.Checked
		If moPEngine Is Nothing Then Exit Sub

		Dim lType As Int32
		Dim sItems() As String
		Dim X As Int32

		ReDim sItems(lstBurn.Items.Count - 1)
		For X = 0 To lstBurn.Items.Count - 1
			sItems(X) = lstBurn.Items(X).ToString
		Next X
		'Now, clear our list
		lstBurn.Items.Clear()

		If mnuAtmosBurn.Checked = True Then
			lType = ParticleFX.ParticleEngine.EmitterType.eSmokeyFireEmitter
		Else : lType = ParticleFX.ParticleEngine.EmitterType.eFireEmitter
		End If

		For X = 0 To sItems.Length - 1
			Dim lVal As Int32 = CInt(Val(Mid$(sItems(X), 5)))

			Dim vecLoc As Vector3 = moPEngine.GetEmitterLoc(lVal)
			Dim lCnt As Int32 = moPEngine.GetEmitterPCnt(lVal)
			moPEngine.StopEmitter(lVal)

			lVal = moPEngine.AddEmitter(CType(lType, ParticleFX.ParticleEngine.EmitterType), vecLoc, lCnt)
			lstBurn.Items.Add("Item " & lVal)
		Next X

	End Sub

	Private Sub mnuSaveDetailsOnly_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuSaveDetailsOnly.Click
		SaveData(False, True)
	End Sub

	Private Sub mnuFireFrom_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuFireFrom.Click
		mnuFireFrom.Checked = Not mnuFireFrom.Checked
	End Sub

	Private Sub mnuBL_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuBL.Click
		mnuBL.Checked = Not mnuBL.Checked
	End Sub

	Private Sub mnuShowDeath_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuShowDeath.Click
		mnuShowDeath.Checked = Not mnuShowDeath.Checked
		If mnuShowDeath.Checked = False Then mbShowMesh = True
	End Sub

	Private Sub mnuZX_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuZX.Click
		moMesh.ComputeNormals()

		'Get the original mesh's vertex buffer.
		Dim ranks(0) As Integer
		ranks(0) = moMesh.NumberVertices

		Dim arr As System.Array = moMesh.VertexBuffer.Lock(0, (New CustomVertex.PositionNormalTextured()).GetType(), LockFlags.None, ranks)

		Dim i As Integer

		For i = 0 To arr.Length - 1

			Dim pn As Direct3D.CustomVertex.PositionNormalTextured = CType(arr.GetValue(i), CustomVertex.PositionNormalTextured)

			Dim fTemp As Single = pn.X
			pn.X = pn.Z
			pn.Z = fTemp

			arr.SetValue(pn, i)
		Next i

		moMesh.VertexBuffer.Unlock()
		arr = Nothing
	End Sub
	Private Sub mnuYZ_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuYZ.Click
		moMesh.ComputeNormals()

		'Get the original mesh's vertex buffer.
		Dim ranks(0) As Integer
		ranks(0) = moMesh.NumberVertices

		Dim arr As System.Array = moMesh.VertexBuffer.Lock(0, (New CustomVertex.PositionNormalTextured()).GetType(), LockFlags.None, ranks)

		Dim i As Integer

		For i = 0 To arr.Length - 1

			Dim pn As Direct3D.CustomVertex.PositionNormalTextured = CType(arr.GetValue(i), CustomVertex.PositionNormalTextured)

			Dim fTemp As Single = pn.Z
			pn.Z = pn.Y
			pn.Y = fTemp

			arr.SetValue(pn, i)
		Next i

		moMesh.VertexBuffer.Unlock()
		arr = Nothing
	End Sub
	Private Sub mnuYX_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuYX.Click
		moMesh.ComputeNormals()

		'Get the original mesh's vertex buffer.
		Dim ranks(0) As Integer
		ranks(0) = moMesh.NumberVertices

		Dim arr As System.Array = moMesh.VertexBuffer.Lock(0, (New CustomVertex.PositionNormalTextured()).GetType(), LockFlags.None, ranks)

		Dim i As Integer

		For i = 0 To arr.Length - 1

			Dim pn As Direct3D.CustomVertex.PositionNormalTextured = CType(arr.GetValue(i), CustomVertex.PositionNormalTextured)

			Dim fTemp As Single = pn.X
			pn.X = pn.Y
			pn.Y = fTemp

			arr.SetValue(pn, i)
		Next i

		moMesh.VertexBuffer.Unlock()
		arr = Nothing
	End Sub
	Private Sub mnuXZ_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuXZ.Click
		moMesh.ComputeNormals()

		'Get the original mesh's vertex buffer.
		Dim ranks(0) As Integer
		ranks(0) = moMesh.NumberVertices

		Dim arr As System.Array = moMesh.VertexBuffer.Lock(0, (New CustomVertex.PositionNormalTextured()).GetType(), LockFlags.None, ranks)

		Dim i As Integer

		For i = 0 To arr.Length - 1

			Dim pn As Direct3D.CustomVertex.PositionNormalTextured = CType(arr.GetValue(i), CustomVertex.PositionNormalTextured)

			Dim fTemp As Single = pn.Z
			pn.Z = pn.X
			pn.X = fTemp

			arr.SetValue(pn, i)
		Next i

		moMesh.VertexBuffer.Unlock()
		arr = Nothing
	End Sub
	Private Sub mnuXY_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuXY.Click
		moMesh.ComputeNormals()

		'Get the original mesh's vertex buffer.
		Dim ranks(0) As Integer
		ranks(0) = moMesh.NumberVertices

		Dim arr As System.Array = moMesh.VertexBuffer.Lock(0, (New CustomVertex.PositionNormalTextured()).GetType(), LockFlags.None, ranks)

		Dim i As Integer

		For i = 0 To arr.Length - 1

			Dim pn As Direct3D.CustomVertex.PositionNormalTextured = CType(arr.GetValue(i), CustomVertex.PositionNormalTextured)

			Dim fTemp As Single = pn.Y
			pn.Y = pn.X
			pn.X = fTemp

			arr.SetValue(pn, i)
		Next i

		moMesh.VertexBuffer.Unlock()
		arr = Nothing
	End Sub
	Private Sub mnuZY_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuZY.Click
		moMesh.ComputeNormals()

		'Get the original mesh's vertex buffer.
		Dim ranks(0) As Integer
		ranks(0) = moMesh.NumberVertices

		Dim arr As System.Array = moMesh.VertexBuffer.Lock(0, (New CustomVertex.PositionNormalTextured()).GetType(), LockFlags.None, ranks)

		Dim i As Integer

		For i = 0 To arr.Length - 1

			Dim pn As Direct3D.CustomVertex.PositionNormalTextured = CType(arr.GetValue(i), CustomVertex.PositionNormalTextured)

			Dim fTemp As Single = pn.Y
			pn.Y = pn.Z
			pn.Z = fTemp

			arr.SetValue(pn, i)
		Next i

		moMesh.VertexBuffer.Unlock()
		arr = Nothing
	End Sub
#End Region

#Region " Button Clicks "
	Private Sub cmdDiffuse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDiffuse.Click
		Dim oRes As DialogResult
		cldDlg.Color = Materials(0).Diffuse
		oRes = cldDlg.ShowDialog(Me)
		If oRes = Windows.Forms.DialogResult.OK Then
			For X As Int32 = 0 To Materials.GetUpperBound(0)
				Materials(X).Diffuse = cldDlg.Color
			Next X
			'Materials(0).Diffuse = cldDlg.Color
			cmdDiffuse.BackColor = cldDlg.Color
		End If
	End Sub

	Private Sub cmdEmissive_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdEmissive.Click
		Dim oRes As DialogResult
		cldDlg.Color = Materials(0).Emissive
		oRes = cldDlg.ShowDialog(Me)
		If oRes = Windows.Forms.DialogResult.OK Then
			Materials(0).Emissive = cldDlg.Color
			cmdEmissive.BackColor = cldDlg.Color
		End If
	End Sub

	Private Sub cmdSpecular_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSpecular.Click
		Dim oRes As DialogResult
		cldDlg.Color = Materials(0).Specular
		oRes = cldDlg.ShowDialog(Me)
		If oRes = Windows.Forms.DialogResult.OK Then
			Materials(0).Specular = cldDlg.Color
			cmdSpecular.BackColor = cldDlg.Color
		End If
	End Sub

	Private Sub btnShieldColor_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnShieldColor.Click
		Dim oRes As DialogResult
		cldDlg.Color = moShieldMat.Emissive
		oRes = cldDlg.ShowDialog(Me)
		If oRes = Windows.Forms.DialogResult.OK Then
			moShieldMat.Emissive = cldDlg.Color
			btnShieldColor.BackColor = cldDlg.Color
		End If
	End Sub

	Private Sub btnEngineColor_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEngineColor.Click
		Dim oRes As DialogResult
		cldDlg.Color = btnEngineColor.BackColor
		oRes = cldDlg.ShowDialog(Me)
		If oRes = Windows.Forms.DialogResult.OK Then

			For X As Int32 = 0 To lstEngineFX.Items.Count - 1
				Dim lVal As Int32 = CInt(Val(Mid$(lstEngineFX.Items(X).ToString, 5)))
				moPEngine.ChangeEngineFXColor(lVal, cldDlg.Color.R, cldDlg.Color.G, cldDlg.Color.B)
			Next X

			btnEngineColor.BackColor = cldDlg.Color
		End If
	End Sub

	Private Sub cmdScale_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdScale.Click
		ApplyChange(CSng(Val(txtScaleX.Text)), CSng(Val(txtScaleY.Text)), CSng(Val(txtScaleZ.Text)), 0, 0, 0)
		RefreshStatDisplay()
	End Sub

	Private Sub cmdShift_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdShift.Click
		ApplyChange(0, 0, 0, CSng(Val(txtShiftX.Text)), CSng(Val(txtShiftY.Text)), CSng(Val(txtShiftZ.Text)))
		RefreshStatDisplay()
	End Sub

	Private Sub cmdReloadTexture_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdReloadTexture.Click
		'Dim sTemp As String
		'Dim sPostFix As String
		'Dim X As Int32
		'If mtrlBuffer(0).TextureFilename <> "" Then
		'    sTemp = mtrlBuffer(0).TextureFilename

		'    sPostFix = Mid$(sTemp, Len(sTemp) - 3, 4)
		'    sTemp = Mid$(sTemp, 1, Len(sTemp) - 4)

		'    For X = 0 To 3
		'        Textures(X).Dispose()
		'    Next X

		'    Textures(0) = TextureLoader.FromFile(moDevice, sTemp & sPostFix)
		'    Textures(1) = TextureLoader.FromFile(moDevice, sTemp & "_mine" & sPostFix)
		'    Textures(2) = TextureLoader.FromFile(moDevice, sTemp & "_ally" & sPostFix)
		'    Textures(3) = TextureLoader.FromFile(moDevice, sTemp & "_enemy" & sPostFix)
		'End If
	End Sub

	Private Sub btnResetShieldSphere_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnResetShieldSphere.Click
		Dim sFile As String

		mbRenderOldMesh = Not mbRenderOldMesh
		If mbRenderOldMesh = True Then

			Device.IsUsingEventHandlers = False
			Try
				If moShieldMesh Is Nothing = False Then
					moShieldMesh.Dispose()
					moShieldMesh = Nothing
				End If
				moShieldMesh = CreateShieldSphere(CSng(Val(txtShieldSphereSize.Text)), 32, 32, 1, CSng(Val(txtShieldStretchX.Text)), _
				  CSng(Val(txtShieldStretchY.Text)), CSng(Val(txtShieldStretchZ.Text)), CSng(Val(txtShieldShiftX.Text)), _
				  CSng(Val(txtShieldShiftY.Text)), CSng(Val(txtShieldShiftZ.Text)))

				'If moShieldTex Is Nothing = False Then
				'moShieldTex.Dispose()
				'moShieldTex = Nothing
				'End If
				sFile = AppDomain.CurrentDomain.BaseDirectory
				If sFile.EndsWith("\") = False Then sFile &= "\"
				'sFile &= "ShieldTex.dds"
				Dim X As Int32
				For X = 0 To 3
					If moShieldTex(X) Is Nothing = False Then moShieldTex(X).Dispose()
					moShieldTex(X) = Nothing
					moShieldTex(X) = TextureLoader.FromFile(moDevice, sFile & "Shield" & (X + 1) & ".dds")
				Next
				'moShieldTex = TextureLoader.FromFile(moDevice, sFile, 64, 64, 0, Usage.None, Format.Unknown, Pool.Default, Filter.Box, Filter.Box, 0)

				With moShieldMat
					.Ambient = System.Drawing.Color.FromArgb(64, 0, 0, 0)
					.Diffuse = System.Drawing.Color.FromArgb(64, 0, 0, 0)
					.Emissive = System.Drawing.Color.FromArgb(64, 0, 255, 255)
				End With
			Catch
				MsgBox(Err.Description, MsgBoxStyle.OkOnly)
				moShieldTex = Nothing
				moShieldMesh = Nothing
			Finally
				Device.IsUsingEventHandlers = True
			End Try
		End If
	End Sub

	Private Sub btnAddEngineFX_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddEngineFX.Click
		Dim lIdx As Int32
		If moPEngine Is Nothing Then moPEngine = New ParticleFX.ParticleEngine(moDevice, 32)

		lIdx = moPEngine.AddEmitter(ParticleFX.ParticleEngine.EmitterType.eEngineEmitter, New Vector3( _
		  CSng(Val(txtEngineOffX.Text)), CSng(Val(txtEngineOffY.Text)), CSng(Val(txtEngineOffZ.Text))), CInt(Val(txtPCnt.Text)))
		If lstEngineFX.Items.Count > lIdx Then
			lstEngineFX.Items.Insert(lIdx, "Item " & lIdx)
		Else : lstEngineFX.Items.Add("Item " & lIdx)
		End If

	End Sub

	Private Sub btnDeleteEngineFX_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDeleteEngineFX.Click
		If lstEngineFX.SelectedIndex > -1 Then
			Dim lVal As Int32 = CInt(Val(Mid$(lstEngineFX.SelectedItem.ToString, 5)))
			If moPEngine Is Nothing = False Then moPEngine.StopEmitter(lVal)
			lstEngineFX.Items.RemoveAt(lstEngineFX.SelectedIndex)
		End If
	End Sub

	Private Sub btnAddBurnFX_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddBurnFX.Click
		Dim lIdx As Int32 = -1

		If moPEngine Is Nothing Then moPEngine = New ParticleFX.ParticleEngine(moDevice, 32)

		For X As Int32 = 0 To mlBurnFXUB
			If myBurnFXUsed(X) = 0 Then
				lIdx = X
				Exit For
			End If
		Next X

		If lIdx = -1 Then
			mlBurnFXUB += 1
			ReDim Preserve moBurnFX(mlBurnFXUB)
			ReDim Preserve myBurnFXUsed(mlBurnFXUB)
			lIdx = mlBurnFXUB
		End If

		moBurnFX(lIdx) = New BurnFXData
		myBurnFXUsed(lIdx) = 255

		Dim lVal As Int32

		If mnuAtmosBurn.Checked = True Then
			lVal = ParticleFX.ParticleEngine.EmitterType.eSmokeyFireEmitter
		Else : lVal = ParticleFX.ParticleEngine.EmitterType.eFireEmitter
		End If

		If optPosX.Checked = True Then
			lVal += ParticleFX.ParticleEngine.EmitterType.eFireEmitterMod_PlusX
		ElseIf optMinusX.Checked = True Then
			lVal += ParticleFX.ParticleEngine.EmitterType.eFireEmitterMod_MinusX
		ElseIf optMinusY.Checked = True Then
			lVal += ParticleFX.ParticleEngine.EmitterType.eFireEmitterMod_MinusY
		ElseIf optPosZ.Checked = True Then
			lVal += ParticleFX.ParticleEngine.EmitterType.eFireEmitterMod_PlusZ
		ElseIf optMinusZ.Checked = True Then
			lVal += ParticleFX.ParticleEngine.EmitterType.eFireEmitterMod_MinusZ
		End If

		Dim lSide As Int32
		If optFrontBurn.Checked = True Then
			lSide = 0
		ElseIf optLeftBurn.Checked = True Then
			lSide = 1
		ElseIf optRearBurn.Checked = True Then
			lSide = 2
		Else : lSide = 3
		End If

		With moBurnFX(lIdx)
			.LocX = CInt(Val(txtBurnOffX.Text))
			.LocY = CInt(Val(txtBurnOffY.Text))
			.LocZ = CInt(Val(txtBurnOffZ.Text))
			.lPCnt = CInt(Val(txtBurnPCnt.Text))
			.lParticleFXID = moPEngine.AddEmitter(CType(lVal, ParticleFX.ParticleEngine.EmitterType), New Vector3(.LocX, .LocY, .LocZ), .lPCnt)
			.lEmitterType = CType(lVal, ParticleFX.ParticleEngine.EmitterType)
			.lSide = lSide
			.sDisplay = "Side: " & lSide & ", ID: " & .lParticleFXID
		End With

		lstBurn.Items.Add(moBurnFX(lIdx))
		lstBurn.DisplayMember = "sDisplay"
	End Sub

	Private Sub btnDeleteBurnFX_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDeleteBurnFX.Click
		If lstBurn.SelectedItem Is Nothing = False Then
			With CType(lstBurn.SelectedItem, BurnFXData)
				For X As Int32 = 0 To mlBurnFXUB
					If myBurnFXUsed(X) <> 0 Then
						If .lParticleFXID = moBurnFX(X).lParticleFXID Then
							myBurnFXUsed(X) = 0
							moBurnFX(X) = Nothing
							moPEngine.StopEmitter(.lParticleFXID)
						End If
					End If
				Next X
			End With
			lstBurn.Items.Remove(lstBurn.SelectedItem)
		End If
	End Sub

	Private Sub btnInvertZ_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnInvertZ.Click
		' Get the original mesh's vertex buffer.
		Dim ranks(0) As Integer
		ranks(0) = moMesh.NumberVertices

		Dim arr As System.Array = moMesh.VertexBuffer.Lock(0, (New CustomVertex.PositionNormalTextured()).GetType(), LockFlags.None, ranks)

		Dim i As Integer

		For i = 0 To arr.Length - 1

			Dim pn As Direct3D.CustomVertex.PositionNormalTextured = CType(arr.GetValue(i), CustomVertex.PositionNormalTextured)

			pn.Z *= -1

			arr.SetValue(pn, i)
		Next i

		moMesh.VertexBuffer.Unlock()
		arr = Nothing

		'Now, if we are in Tank View mode, we need to invert the turret too...
		If mbTankMesh = True Then
			ranks(0) = moTurretMesh.NumberVertices
			arr = moTurretMesh.VertexBuffer.Lock(0, (New CustomVertex.PositionNormalTextured()).GetType(), LockFlags.None, ranks)
			For i = 0 To arr.Length - 1
				Dim pn As Direct3D.CustomVertex.PositionNormalTextured = CType(arr.GetValue(i), CustomVertex.PositionNormalTextured)
				pn.Z *= -1
				arr.SetValue(pn, i)
			Next i
			moTurretMesh.VertexBuffer.Unlock()
			arr = Nothing
		End If
	End Sub

	Private Sub btnReverseWind_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReverseWind.Click
		' Get the original mesh's index buffer.
		Dim ranks(0) As Int32
		Dim arr As System.Array

		ranks(0) = moMesh.NumberFaces * 3
		arr = moMesh.LockIndexBuffer((New Short()).GetType(), LockFlags.None, ranks)
		Array.Reverse(arr)
		moMesh.IndexBuffer.SetData(arr, 0, LockFlags.None)

		arr = Nothing
	End Sub

	Private Sub btnDeleteFireFrom_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDeleteFireFrom.Click
		If lstFireFroms.SelectedIndex > -1 Then
			Dim sTemp As String = lstFireFroms.SelectedItem.ToString
			Dim ySide As Byte
			If InStr(sTemp, "(F)", CompareMethod.Binary) <> 0 Then
				ySide = 0
			ElseIf InStr(sTemp, "(L)", CompareMethod.Binary) <> 0 Then
				ySide = 1
			ElseIf InStr(sTemp, "(A)", CompareMethod.Binary) <> 0 Then
				ySide = 2
			Else
				ySide = 3
			End If

			sTemp = Mid$(sTemp, 1, sTemp.Length - 4)
			On Error Resume Next
			Err.Clear()
			mcolFireFromLocs(ySide).Remove(sTemp)
			'If Err.Number = 0 Then
			lstFireFroms.Items.RemoveAt(lstFireFroms.SelectedIndex)
			'End 'If
			On Error GoTo 0
		End If
	End Sub

	Private Sub btnAddFireFrom_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddFireFrom.Click
		Dim ySide As Byte
		Dim sAppend As String

		If optFireFromFront.Checked = True Then
			ySide = 0 : sAppend = " (F)"
		ElseIf optFireFromLeft.Checked = True Then
			ySide = 1 : sAppend = " (L)"
		ElseIf optFireFromRear.Checked = True Then
			ySide = 2 : sAppend = " (A)"
		Else : ySide = 3 : sAppend = " (R)"
		End If

		If mcolFireFromLocs(ySide) Is Nothing Then mcolFireFromLocs(ySide) = New Collection()

		'Now, create our loc
		Dim vecTemp As Vector3 = New Vector3(CSng(Val(txtFireFromOffX.Text)), CSng(Val(txtFireFromOffY.Text)), CSng(Val(txtFireFromOffZ.Text)))

		mcolFireFromLocs(ySide).Add(vecTemp, "FF" & mlFireFromSeq)
		lstFireFroms.Items.Add("FF" & mlFireFromSeq & sAppend)
		mlFireFromSeq += 1
	End Sub

	Private mbRenderBuildProgress As Boolean = False
	Private mfProgress As Single = 0.0F
	Private matProgress As Material
	Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
		'Dim lVal As Int32

		'If mbRenderBuildProgress = False Then
		'    mbRenderBuildProgress = True
		'    mfProgress = 0.0F
		'Else
		'    mfProgress += 0.1F
		'    If mfProgress > 100.0F Then mfProgress = 100.0F
		'End If
		'lVal = CInt(mfProgress / 100.0F * 255)
		'If lVal < 0 Then lVal = 0
		'If lVal > 255 Then lVal = 255

		'If lVal = 255 Then
		'    matProgress = Materials(0)
		'Else
		'    With matProgress
		'        .Ambient = System.Drawing.Color.FromArgb(lVal, 255, 255, 255)
		'        .Diffuse = System.Drawing.Color.FromArgb(lVal, 255, 255, 255)
		'        .Emissive = System.Drawing.Color.FromArgb(255 - lVal, 255, 255, 255)
		'        .Specular = System.Drawing.Color.FromArgb(0, 0, 0, 0)
		'        .SpecularSharpness = 0
		'    End With
		'End If

		'Me.Text = mfProgress.ToString("###.#")

		'Dim oIntLoc As IntersectInformation

		'Dim oFromLoc As Vector3 = New Vector3(mlCameraX, mlCameraY, mlCameraZ)
		'Dim oDir As Vector3 = New Vector3(-mlCameraX, -mlCameraY, -mlCameraZ)
		'oDir.Normalize()
		'moMesh.Intersect(oFromLoc, oDir, oIntLoc)

		'oDir.Scale(oIntLoc.Dist)
		'mvecHitLoc = Vector3.Add(oFromLoc, oDir)
		''mvecHitLoc = Vector3.Scale(oFromLoc, oIntLoc.Dist)

		'mbRenderHitloc = True

		'moMesh.ComputeNormals()

		'Materials(0).SpecularSharpness = 17

		' Get the original mesh's vertex buffer.
		'Dim ranks(0) As Integer
		'ranks(0) = moMesh.NumberVertices

		'Dim arr As System.Array = moMesh.VertexBuffer.Lock(0, (New CustomVertex.PositionNormalTextured()).GetType(), LockFlags.None, ranks)

		'Dim i As Integer

		'For i = 0 To arr.Length - 1

		'    Dim pn As Direct3D.CustomVertex.PositionNormalTextured = CType(arr.GetValue(i), CustomVertex.PositionNormalTextured)

		'    Dim fTemp As Single = pn.Y
		'    pn.Y = pn.Z
		'    pn.Z = fTemp

		'    arr.SetValue(pn, i)
		'Next i

		'moMesh.VertexBuffer.Unlock()
		'arr = Nothing

		'FillVertList()

		MsgBox(GetEstimatedHull)
	End Sub

	Private Sub btnResetCamera_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnResetCamera.Click
		mlCameraAtX = 0
		mlCameraAtY = 0
		mlCameraAtZ = 0
	End Sub
#End Region

#Region " Text Box Events "
	Private Sub txtEngineOffX_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtEngineOffX.TextChanged
		UpdateEngineFX()
	End Sub

	Private Sub txtEngineOffY_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtEngineOffY.TextChanged
		UpdateEngineFX()
	End Sub

	Private Sub txtEngineOffZ_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtEngineOffZ.TextChanged
		UpdateEngineFX()
	End Sub

	Private Sub txtBurnOffZ_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtBurnOffZ.TextChanged
		UpdateBurnFX()
	End Sub

	Private Sub txtBurnOffY_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtBurnOffY.TextChanged
		UpdateBurnFX()
	End Sub

	Private Sub txtBurnOffX_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtBurnOffX.TextChanged
		UpdateBurnFX()
	End Sub

	Private Sub txtFireFromOffY_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtFireFromOffY.TextChanged
		UpdateFireFromFX()
	End Sub

	Private Sub txtFireFromOffX_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtFireFromOffX.TextChanged
		UpdateFireFromFX()
	End Sub

	Private Sub txtFireFromOffZ_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtFireFromOffZ.TextChanged
		UpdateFireFromFX()
	End Sub
#End Region

#Region " List Box Events "

	Private Sub lstEngineFX_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lstEngineFX.SelectedIndexChanged
		Dim vecLoc As Vector3
		If moPEngine Is Nothing Then Exit Sub
		If lstEngineFX.SelectedIndex = -1 Then Exit Sub

		Dim lVal As Int32 = CInt(Val(Mid$(lstEngineFX.SelectedItem.ToString, 5)))

		mbIgnoreEngineFXEvents = True
		vecLoc = moPEngine.GetEmitterLoc(lVal)
		txtEngineOffX.Text = vecLoc.X.ToString
		txtEngineOffY.Text = vecLoc.Y.ToString
		txtEngineOffZ.Text = vecLoc.Z.ToString

		mbIgnoreEngineFXEvents = False
	End Sub

	Private Sub lstBurn_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lstBurn.SelectedIndexChanged
		If lstBurn.SelectedItem Is Nothing Then Return

		mbIgnoreBurnFXEvents = True

		Dim oFXData As BurnFXData = CType(lstBurn.SelectedItem, BurnFXData)
		If oFXData Is Nothing = False Then
			With oFXData
				txtBurnOffX.Text = .LocX.ToString
				txtBurnOffY.Text = .LocY.ToString
				txtBurnOffZ.Text = .LocZ.ToString

				optFrontBurn.Checked = False : optLeftBurn.Checked = False : optRightBurn.Checked = False : optRearBurn.Checked = False
				Select Case oFXData.lSide
					Case 0 : optFrontBurn.Checked = True
					Case 1 : optLeftBurn.Checked = True
					Case 2 : optRearBurn.Checked = True
					Case Else : optRightBurn.Checked = True
				End Select

				txtPCnt.Text = oFXData.lPCnt.ToString

				Dim lVal As Int32 = CInt(.lEmitterType)
				If (lVal And ParticleFX.ParticleEngine.EmitterType.eFireEmitter) <> 0 Then
					lVal -= ParticleFX.ParticleEngine.EmitterType.eFireEmitter
				ElseIf (lVal And ParticleFX.ParticleEngine.EmitterType.eSmokeyFireEmitter) <> 0 Then
					lVal -= ParticleFX.ParticleEngine.EmitterType.eSmokeyFireEmitter
				End If

				optPosX.Checked = False : optPosY.Checked = False : optPosZ.Checked = False
				optMinusX.Checked = False : optMinusY.Checked = False : optMinusZ.Checked = False

				Select Case CType(lVal, ParticleFX.ParticleEngine.EmitterType)
					Case ParticleFX.ParticleEngine.EmitterType.eFireEmitterMod_MinusX
						optMinusX.Checked = True
					Case ParticleFX.ParticleEngine.EmitterType.eFireEmitterMod_MinusY
						optMinusY.Checked = True
					Case ParticleFX.ParticleEngine.EmitterType.eFireEmitterMod_MinusZ
						optMinusZ.Checked = True
					Case ParticleFX.ParticleEngine.EmitterType.eFireEmitterMod_PlusX
						optPosX.Checked = True
					Case ParticleFX.ParticleEngine.EmitterType.eFireEmitterMod_PlusZ
						optPosZ.Checked = True
					Case Else
						optPosY.Checked = True
				End Select
			End With
		End If

		mbIgnoreBurnFXEvents = False
	End Sub

	Private Sub lstFireFroms_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lstFireFroms.SelectedIndexChanged

		mbIgnoreFireFromEvents = True

		Dim sTemp As String = lstFireFroms.SelectedItem.ToString
		Dim ySide As Byte

		optFireFromFront.Checked = False : optFireFromLeft.Checked = False : optFireFromRear.Checked = False : optFireFromRight.Checked = False
		If InStr(sTemp, "(F)", CompareMethod.Binary) <> 0 Then
			ySide = 0
			optFireFromFront.Checked = True
		ElseIf InStr(sTemp, "(L)", CompareMethod.Binary) <> 0 Then
			ySide = 1
			optFireFromLeft.Checked = True
		ElseIf InStr(sTemp, "(A)", CompareMethod.Binary) <> 0 Then
			ySide = 2
			optFireFromRear.Checked = True
		Else
			ySide = 3
			optFireFromRight.Checked = True
		End If

		sTemp = Mid$(sTemp, 1, sTemp.Length - 4)
		On Error Resume Next
		Err.Clear()
		Dim vecTemp As Vector3 = CType(mcolFireFromLocs(ySide)(sTemp), Vector3)
		If Err.Number = 0 Then
			txtFireFromOffX.Text = vecTemp.X.ToString
			txtFireFromOffY.Text = vecTemp.Y.ToString
			txtFireFromOffZ.Text = vecTemp.Z.ToString
		End If
		On Error GoTo 0

		mbIgnoreFireFromEvents = False
	End Sub

#End Region

#End Region

    Public Shared l_Burn_Op As TextureOperation = TextureOperation.Modulate
    Public Shared l_DS_Op As TextureOperation = TextureOperation.Modulate

	Private mfTimeOfDay As Single = 0.0F
	Private moSW As Stopwatch
	Private Sub DrawScene()
		Dim matWorld As Matrix
		Static xlLastTime As Int32
        Dim matTurretWorld As Matrix

        If cboBurn.SelectedIndex > -1 Then l_Burn_Op = CType(cboBurn.SelectedIndex + 1, TextureOperation)
        If cboDS.SelectedIndex > -1 Then l_DS_Op = CType(cboDS.SelectedIndex + 1, TextureOperation)

        moDevice.BeginScene()
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

        mfTimeOfDay += 0.0025F * fElapsed
        If mfTimeOfDay > 1.0F Then mfTimeOfDay -= 1.0F

        With moDevice.Lights(0)
            .Diffuse = System.Drawing.Color.White
            .Ambient = System.Drawing.Color.Black
            .Type = LightType.Directional
            '.Position = New Vector3(mlSunX, 1500, 1500)
            Dim fX As Single = 1.0F
            Dim fY As Single = 0.0F
            RotatePoint(0, 0, fX, fY, mfTimeOfDay * 360.0F)

            'Me.Text = mfTimeOfDay.ToString("0.000") & ", " & fX.ToString("0.000") & ", " & fY.ToString("0.000")

            .Direction = Vector3.Multiply(New Vector3(fX, fY, 0.0F), 5)

            '.Direction = New Vector3(-0.94F, -0.32F, -0.12F)

            'New Vector3( 0, -500, -1)
            .Range = 100000
            .Specular = System.Drawing.Color.White
            .Attenuation0 = 1
            .Attenuation1 = 0
            .Attenuation2 = 0
            .Falloff = 0.3
            '.InnerConeAngle = 0.0523598776
            '.OuterConeAngle = 0.698131701

            .Enabled = True
            .Update()
        End With

        If moFireFromIndicator Is Nothing Then
            Device.IsUsingEventHandlers = False
            moFireFromIndicator = Mesh.Sphere(moDevice, 32, 4, 4)
            Device.IsUsingEventHandlers = True
            SetFireFromMaterials()
        End If
        Dim vecLightVec As Vector3 = moDevice.Lights(0).Direction
        vecLightVec.Scale(-250.0F)
        moDevice.Transform.World = Matrix.Translation(vecLightVec)
        moDevice.Material = matFireFrom(0)
        moDevice.SetTexture(0, Nothing)
        moFireFromIndicator.DrawSubset(0)
        moDevice.Transform.World = Matrix.Identity

        matWorld = Matrix.Identity
        matWorld.Multiply(Matrix.RotationZ(CSng(hscrYaw.Value * gdRadPerDegree)))
        matWorld.Multiply(Matrix.RotationY(CSng((hscrRotate.Value / 10.0F) * gdRadPerDegree)))
        matWorld.Multiply(Matrix.Translation(0, CSng(Val(txtPlanetYAdjust.Text)), 0))
        moDevice.Transform.World = matWorld
        moDevice.RenderState.CullMode = Cull.None



        'do rendering here
        Dim X As Int32
        Dim lTexMod As Int32

        If optNeutral.Checked = True Then
            lTexMod = 0
        ElseIf optMine.Checked = True Then
            lTexMod = 1
        ElseIf optAlly.Checked = True Then
            lTexMod = 2
        ElseIf optEnemy.Checked = True Then
            lTexMod = 3
        End If

        If mbShowMesh = True Then

            If moMesh Is Nothing = False Then

                'moDevice.RenderState.AlphaBlendEnable = False
                If mbRenderBuildProgress = True Then

                    moDevice.SetTextureStageState(0, TextureStageStates.AlphaArgument0, TextureArgument.TextureColor)
                    moDevice.SetTextureStageState(0, TextureStageStates.AlphaArgument1, TextureArgument.Diffuse)
                    moDevice.SetTextureStageState(0, TextureStageStates.AlphaOperation, TextureOperation.Modulate2X)

                    Static xlPrevTicks As Long
                    If Now.Ticks - xlPrevTicks > 1000 Then
                        mfProgress += 0.1F
                        xlPrevTicks = Now.Ticks

                        If mfProgress > 100.0F Then mfProgress = 100

                        Dim lVal As Int32
                        If mfProgress < 50 Then
                            lVal = CInt(mfProgress / 50.0F * 255)
                        Else
                            lVal = CInt((mfProgress - 50) / 50.0F * 255)
                        End If
                        lVal = CInt(mfProgress / 100.0F * 255)

                        If lVal < 0 Then lVal = 0
                        If lVal > 255 Then lVal = 255

                        Me.Text = mfProgress.ToString("###.0") & " lVal: " & lVal

                        If lVal = 255 Then
                            matProgress = Materials(0)
                        Else
                            With Materials(0)
                                matProgress.Ambient = System.Drawing.Color.FromArgb(lVal, .Ambient.R, .Ambient.G, .Ambient.B)
                                matProgress.Diffuse = System.Drawing.Color.FromArgb(lVal, .Diffuse.R, .Diffuse.G, .Diffuse.B)
                                matProgress.Specular = System.Drawing.Color.FromArgb(lVal, .Specular.R, .Specular.G, .Specular.B)
                                matProgress.SpecularSharpness = .SpecularSharpness
                            End With
                        End If
                    End If

                End If

                If mbRenderShader = True Then
                    If moShader Is Nothing Then
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

                        moShader = New SimpleShader(moDevice)
                    End If
                    'moDevice.RenderState.CullMode = Cull.CounterClockwise
                    moShader.RenderMesh(moMesh, mlCameraX, mlCameraY, mlCameraZ, moDevice.Lights(0).Direction)
                Else
                    'moDevice.RenderState.ZBufferEnable = False 
                    For X = 0 To NumOfMaterials - 1
                        If mbRenderBuildProgress = True Then
                            moDevice.Material = matProgress
                        Else : moDevice.Material = Materials(X)
                        End If
                        moDevice.SetTexture(0, Textures((X * 4) + lTexMod))
                        moMesh.DrawSubset(X)
                    Next X
                    'moDevice.RenderState.ZBufferEnable = True
                End If

                If mbTankMesh = True AndAlso moTurretMesh Is Nothing = False Then
                    mfTurretRot += 0.01F
                    If mfTurretRot > 2 * Math.PI Then mfTurretRot -= CSng(2.0F * Math.PI)
                    matTurretWorld = Matrix.Identity
                    matTurretWorld.RotateY(mfTurretRot)

                    matTurretWorld.Multiply(Matrix.Translation(0, CSng(Val(txtPlanetYAdjust.Text)), CSng(Val(txtTurretOffZ.Text))))
                    moDevice.Transform.World = matTurretWorld

                    For X = 0 To NumOfMaterials - 1
                        moDevice.Material = Materials(X)
                        moDevice.SetTexture(0, Textures((X * 4) + lTexMod))
                        moTurretMesh.DrawSubset(X)
                    Next X

                    moDevice.Transform.World = matWorld
                End If

                If mbRenderBuildProgress = True Then
                    moDevice.SetTextureStageState(0, TextureStageStates.AlphaOperation, TextureOperation.Disable)

                    Dim lPrevSrcBlnd As Blend = moDevice.RenderState.SourceBlend
                    Dim lPrevDestBlnd As Blend = moDevice.RenderState.DestinationBlend
                    Dim lPrevBlndOp As BlendOperation = moDevice.RenderState.BlendOperation

                    moDevice.RenderState.SourceBlend = Blend.InvSourceAlpha
                    moDevice.RenderState.DestinationBlend = Blend.One
                    moDevice.RenderState.BlendOperation = BlendOperation.Add

                    For X = 0 To NumOfMaterials - 1
                        If mbRenderBuildProgress = True Then
                            moDevice.Material = matProgress
                        Else : moDevice.Material = Materials(X)
                        End If
                        moDevice.SetTexture(0, Textures((X * 4) + lTexMod))
                        moMesh.DrawSubset(X)
                    Next X

                    moDevice.RenderState.SourceBlend = lPrevSrcBlnd
                    moDevice.RenderState.DestinationBlend = lPrevDestBlnd
                    moDevice.RenderState.BlendOperation = lPrevBlndOp

                    If mfProgress = 100.0F Then mbRenderBuildProgress = False
                End If

                'moDevice.RenderState.AlphaBlendEnable = True
            End If

            'Blinking Lights
            If mnuBL.Checked = True Then
                DrawBlinkingLights()
            End If

            'Now, check for rendering engine fx
            If mnuRenderEngineFX.Checked = True Then
                If moPEngine Is Nothing Then moPEngine = New ParticleFX.ParticleEngine(moDevice, 32)
                moPEngine.Render(False)
            End If

            'check if we need to render our shield
            If mnuRenderShield.Checked = True Then
                'If moShieldMesh Is Nothing Then btnResetShieldSphere_Click(Nothing, Nothing)
                'If ParticleFX.Globals.timeGetTime - xlLastTime > 30 Then
                '    mlCurrShieldTex += 1
                '    xlLastTime = ParticleFX.Globals.timeGetTime
                'End If
                'If mlCurrShieldTex > moShieldTex.Length - 1 Then mlCurrShieldTex = 0
                'moDevice.SetTexture(0, moShieldTex(mlCurrShieldTex))
                'moDevice.Material = moShieldMat
                'moDevice.RenderState.AlphaBlendEnable = True
                'moShieldMesh.DrawSubset(0)

                'Dim omatNew As Matrix = Matrix.Identity
                'omatNew.Scale(1.05, 1.05, 1.05)
                'moDevice.Transform.World = omatNew

                If ParticleFX.Globals.timeGetTime - xlLastTime > 30 Then
                    mlCurrShieldTex += 1
                    xlLastTime = ParticleFX.Globals.timeGetTime
                End If
                If mlCurrShieldTex > moShieldTex.GetUpperBound(0) Then mlCurrShieldTex = 0

                moDevice.SetTexture(0, moShieldTex(mlCurrShieldTex))
                moDevice.Material = moShieldMat
                moDevice.RenderState.AlphaBlendEnable = True

                For X = 0 To NumOfMaterials - 1
                    'moDevice.Material = Materials(X)
                    'moDevice.SetTexture(0, Textures((X * 4) + lTexMod))
                    moMesh.DrawSubset(X)
                Next X

                If mbTankMesh = True AndAlso moTurretMesh Is Nothing = False Then
                    mfTurretRot += 0.01F
                    If mfTurretRot > 2 * Math.PI Then mfTurretRot -= CSng(2.0F * Math.PI)
                    matTurretWorld = Matrix.Identity
                    matTurretWorld.RotateY(mfTurretRot)

                    matTurretWorld.Multiply(Matrix.Translation(0, CSng(Val(txtPlanetYAdjust.Text)), CSng(Val(txtTurretOffZ.Text))))
                    moDevice.Transform.World = matTurretWorld

                    For X = 0 To NumOfMaterials - 1
                        'moDevice.Material = Materials(X)
                        'moDevice.SetTexture(0, Textures((X * 4) + lTexMod))
                        moTurretMesh.DrawSubset(X)
                    Next X

                    moDevice.Transform.World = matWorld
                End If
            End If

            If mnuRenderCompare.Checked = True Then
                If moCompareMesh Is Nothing = False Then
                    moDevice.Transform.World = Matrix.Identity
                    moDevice.RenderState.FillMode = FillMode.WireFrame


                    For X = 0 To mlCompareNumMat - 1
                        moDevice.Material = moCompareMat(X)
                        moDevice.SetTexture(0, moCompareTex(X))
                        moCompareMesh.DrawSubset(X)
                    Next X
                    If mnuWireFrame.Checked = False Then moDevice.RenderState.FillMode = FillMode.Solid
                End If
            End If

            If mnuFireFrom.Checked = True Then
                If moFireFromIndicator Is Nothing Then
                    Device.IsUsingEventHandlers = False
                    moFireFromIndicator = Mesh.Sphere(moDevice, 32, 4, 4)
                    Device.IsUsingEventHandlers = True
                    SetFireFromMaterials()
                End If

                moDevice.SetTexture(0, Nothing)

                Dim vecTemp As Vector3
                For X = 0 To 3
                    moDevice.Material = matFireFrom(X)
                    If mcolFireFromLocs(X) Is Nothing = False Then
                        For Each vecTemp In mcolFireFromLocs(X)
                            Dim matTemp As Matrix
                            matWorld = Matrix.Identity
                            matTemp = Matrix.Identity
                            matTemp.Translate(vecTemp)
                            matWorld.Multiply(matTemp)

                            moDevice.SetTransform(TransformType.World, matWorld)
                            moFireFromIndicator.DrawSubset(0)
                        Next
                    End If
                Next

                If moFireTo Is Nothing = False Then
                    For X = 0 To 3
                        moDevice.Material = matFireFrom(X)
                        Dim matTemp As Matrix
                        matWorld = Matrix.Identity
                        matTemp = Matrix.Identity
                        matTemp.Translate(moFireTo(X))
                        matWorld.Multiply(matTemp)

                        moDevice.SetTransform(TransformType.World, matWorld)
                        moFireFromIndicator.DrawSubset(0)
                    Next
                End If
            End If
        End If

        If mnuShowDeath.Checked = True Then
            If moDeathSeq Is Nothing Then
                Dim lExpCnt As Int32 = CInt(Val(txtDS_ExpCnt.Text))
                Dim lSizeX As Int32 = CInt(Val(txtDS_SizeX.Text))
                Dim lSizeY As Int32 = CInt(Val(txtDS_SizeY.Text))
                Dim lSizeZ As Int32 = CInt(Val(txtDS_SizeZ.Text))

                moDeathSeq = New DeathSequence(moDevice, New Vector3(0, CInt(Val(txtMidY.Text)) \ 2, 0), lExpCnt, New Vector3(lSizeX, lSizeY, lSizeZ), chkFinale.Checked, -1)
            End If

            If moDeathSeq.bSequenceEnded = True Then
                moDeathSeq = Nothing
                mbShowMesh = True
                'mnuShowDeath.Checked = False
            Else
                moDeathSeq.Render(False)
                mbShowMesh = moDeathSeq.bRenderMesh
            End If
        End If

        moDevice.RenderState.AlphaBlendEnable = True

        'do we render our plane?
        'moDevice.RenderState.ZBufferEnable = False
        If mnuRenderPlane.Checked = True Then
            'yes, okay, we do this with a User Quad
            Dim uVerts(3) As CustomVertex.PositionColored
            uVerts(0) = New CustomVertex.PositionColored(-250000, 0, -250000, System.Drawing.Color.FromArgb(255, 32, 64, 92).ToArgb) ' System.Drawing.Color.Gray.ToArgb)
            uVerts(1) = New CustomVertex.PositionColored(-250000, 0, 250000, System.Drawing.Color.FromArgb(255, 32, 64, 92).ToArgb) 'System.Drawing.Color.Gray.ToArgb)
            uVerts(2) = New CustomVertex.PositionColored(250000, 0, -250000, System.Drawing.Color.FromArgb(255, 32, 64, 92).ToArgb) '
            uVerts(3) = New CustomVertex.PositionColored(250000, 0, 250000, System.Drawing.Color.FromArgb(255, 32, 128, 255).ToArgb) 'System.Drawing.Color.Gray.ToArgb)
            moDevice.Transform.World = Matrix.Identity
            moDevice.VertexFormat = CustomVertex.PositionColored.Format
            moDevice.RenderState.Lighting = False
            moDevice.SetTexture(0, Nothing)
            moDevice.DrawUserPrimitives(PrimitiveType.TriangleStrip, 2, uVerts)
            moDevice.RenderState.Lighting = True
            Erase uVerts
        End If
        'moDevice.RenderState.ZBufferEnable = True

        If mbRenderGlowFX = True Then
            If moGlowFX Is Nothing Then moGlowFX = New PostShader(moDevice)
            moGlowFX.ExecutePostProcess()
        End If

        If mbRenderHitloc = True Then
            If moFireFromIndicator Is Nothing Then
                Device.IsUsingEventHandlers = False
                moFireFromIndicator = Mesh.Sphere(moDevice, 32, 4, 4)
                Device.IsUsingEventHandlers = True
                SetFireFromMaterials()
            End If

            Dim matTemp As Matrix
            matWorld = Matrix.Identity
            matTemp = Matrix.Identity
            matTemp.Translate(mvecHitLoc)
            matWorld.Multiply(matTemp)

            moDevice.Material = matFireFrom(0)
            moDevice.SetTexture(0, Nothing)

            moDevice.SetTransform(TransformType.World, matWorld)
            moFireFromIndicator.DrawSubset(0)
        End If

        If mbRenderOldMesh = True Then
            If moShieldMesh Is Nothing Then btnResetShieldSphere_Click(Nothing, Nothing)
            mbRenderOldMesh = True
            If ParticleFX.Globals.timeGetTime - xlLastTime > 30 Then
                mlCurrShieldTex += 1
                xlLastTime = ParticleFX.Globals.timeGetTime
            End If
            If mlCurrShieldTex > moShieldTex.Length - 1 Then mlCurrShieldTex = 0
            moDevice.SetTexture(0, moShieldTex(mlCurrShieldTex))
            moDevice.Material = moShieldMat
            moDevice.RenderState.AlphaBlendEnable = True
            moShieldMesh.DrawSubset(0)
        End If


        SetupMatrices()

        moDevice.EndScene()

        If mbCaptureScreenshot = True Then
            DoCaptureScreenshot()
            mbCaptureScreenshot = False
        End If
        moDevice.Present()
    End Sub

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

		'Now, set my values
		txtSizeX.Text = (lMaxX - lMinX).ToString
		txtSizeY.Text = (lMaxY - lMinY).ToString
		txtSizeZ.Text = (lMaxZ - lMinZ).ToString

		txtMinX.Text = lMinX.ToString
		txtMinY.Text = lMinY.ToString
		txtMinZ.Text = lMinZ.ToString
		txtMaxX.Text = lMaxX.ToString
		txtMaxY.Text = lMaxY.ToString
		txtMaxZ.Text = lMaxZ.ToString

		txtMidX.Text = (lMinX + ((lMaxX - lMinX) / 2)).ToString
		txtMidY.Text = (lMinY + ((lMaxY - lMinY) / 2)).ToString
		txtMidZ.Text = (lMinZ + ((lMaxZ - lMinZ) / 2)).ToString

		cmdDiffuse.BackColor = Materials(0).Diffuse
		cmdSpecular.BackColor = Materials(0).Specular
		cmdEmissive.BackColor = Materials(0).Emissive
		btnShieldColor.BackColor = moShieldMat.Emissive
		btnEngineColor.BackColor = System.Drawing.Color.FromArgb(255, 64, 192, 255)

		lblRngOff.Text = ((((lMaxX - lMinX) + (lMaxY - lMinY) + (lMaxZ + lMinZ)) / 3) / 25).ToString

	End Sub

	Private Sub LoadExtendedDetails(ByVal sFile As String)
		'sfile is the whole path, so get just the ending name
		Dim lTemp As Int32 = InStrRev(sFile, "\")
		sFile = UCase$(Mid$(sFile, lTemp + 1))

		'Now, go through our Dat file...
		Dim X As Int32
		Dim sTemp As String = AppDomain.CurrentDomain.BaseDirectory
		If sTemp.EndsWith("\") = False Then sTemp &= "\"
		Dim oINI As InitFile = New InitFile(sTemp & "Models.dat")
		Dim lModelID As Int32
		Dim sModelHdr As String = ""

		Dim lIdx As Int32

		mlBurnFXUB = -1
		If mcolFireFromLocs Is Nothing = False Then
			For X = 0 To mcolFireFromLocs.GetUpperBound(0)
				If mcolFireFromLocs(X) Is Nothing = False Then mcolFireFromLocs(X).Clear()
			Next X
		End If
		mlBL_LocUB = -1
		lstSFX.Items.Clear()

		For X = 0 To 1000		'NOTE: if more than 1000 models, than we will need to change this
			sTemp = UCase$(oINI.GetString("MODEL_" & X, "FileName", ""))
			If sTemp = sFile Then
				'ok, found it...
				lModelID = X
				txtModelID.Text = lModelID.ToString
				sModelHdr = "MODEL_" & lModelID
				Exit For
			End If
		Next X

		'Ok, load the rest... clear out our particle engine
		If moPEngine Is Nothing = False Then
			moPEngine = Nothing
		End If
		moPEngine = New ParticleFX.ParticleEngine(moDevice, 32)

		'Now, set our attrs that weren't set from loading
		txtPlanetYAdjust.Text = oINI.GetString(sModelHdr, "PlanetYAdjust", "0")
		txtShieldSphereSize.Text = (oINI.GetString(sModelHdr, "ShieldSphereSize", "0"))
		txtShieldStretchX.Text = Val(oINI.GetString(sModelHdr, "ShieldSphereStretchX", "0")).ToString("###.##")
		txtShieldStretchY.Text = Val(oINI.GetString(sModelHdr, "ShieldSphereStretchY", "0")).ToString("###.##")
		txtShieldStretchZ.Text = Val(oINI.GetString(sModelHdr, "ShieldSphereStretchZ", "0")).ToString("###.##")
		txtShieldShiftX.Text = (oINI.GetString(sModelHdr, "ShieldSphereShiftX", "0"))
		txtShieldShiftY.Text = (oINI.GetString(sModelHdr, "ShieldSphereShiftY", "0"))
		txtShieldShiftZ.Text = (oINI.GetString(sModelHdr, "ShieldSphereShiftZ", "0"))
		chkLandBased.Checked = (Val(oINI.GetString(sModelHdr, "LandBased", "0")) <> 0)

		lTemp = CInt(Val(oINI.GetString(sModelHdr, "EngineFXCnt", "0"))) - 1I
		lstEngineFX.Items.Clear()
		mbIgnoreEngineFXEvents = True

		For X = 0 To lTemp
			txtEngineOffX.Text = (oINI.GetString(sModelHdr, "Engine" & X & "_OffsetX", "0"))
			txtEngineOffY.Text = (oINI.GetString(sModelHdr, "Engine" & X & "_OffsetY", "0"))
			txtEngineOffZ.Text = (oINI.GetString(sModelHdr, "Engine" & X & "_OffsetZ", "0"))
			txtPCnt.Text = (oINI.GetString(sModelHdr, "Engine" & X & "_PCnt", "0"))
			btnAddEngineFX_Click(Nothing, Nothing)
		Next X
		mbIgnoreEngineFXEvents = False

		lTemp = CInt(Val(oINI.GetString(sModelHdr, "BurnFXCnt", "0"))) - 1
		lstBurn.Items.Clear()
		mbIgnoreBurnFXEvents = True
		For X = 0 To lTemp
			txtBurnOffX.Text = (oINI.GetString(sModelHdr, "Burn" & X & "_X", "0"))
			txtBurnOffY.Text = (oINI.GetString(sModelHdr, "Burn" & X & "_Y", "0"))
			txtBurnOffZ.Text = (oINI.GetString(sModelHdr, "Burn" & X & "_Z", "0"))
			txtBurnPCnt.Text = (oINI.GetString(sModelHdr, "Burn" & X & "_PCnt", "0"))
			lIdx = CInt(Val(oINI.GetString(sModelHdr, "BurnSide_" & X, "0")))

			optFrontBurn.Checked = False : optLeftBurn.Checked = False : optRightBurn.Checked = False : optRearBurn.Checked = False
			Select Case lIdx
				Case 0 : optFrontBurn.Checked = True
				Case 1 : optLeftBurn.Checked = True
				Case 2 : optRearBurn.Checked = True
				Case Else : optRightBurn.Checked = True
			End Select

			btnAddBurnFX_Click(Nothing, Nothing)
		Next X
		mbIgnoreBurnFXEvents = False

		oINI = Nothing

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

	Private Sub ApplyChange(ByVal fScaleX As Single, ByVal fScaleY As Single, ByVal fScaleZ As Single, ByVal fShiftX As Single, ByVal fShiftY As Single, ByVal fShiftZ As Single)
		' Get the original mesh's vertex buffer.
		Dim ranks(0) As Integer
		ranks(0) = moMesh.NumberVertices

		Dim arr As System.Array = moMesh.VertexBuffer.Lock(0, (New CustomVertex.PositionNormalTextured()).GetType(), LockFlags.None, ranks)

		Dim i As Integer

		For i = 0 To arr.Length - 1

			Dim pn As Direct3D.CustomVertex.PositionNormalTextured = CType(arr.GetValue(i), CustomVertex.PositionNormalTextured)

			If fScaleX <> 0 Then pn.X *= fScaleX
			If fScaleY <> 0 Then pn.Y *= fScaleY
			If fScaleZ <> 0 Then pn.Z *= fScaleZ

			pn.X += fShiftX
			pn.Y += fShiftY
			pn.Z += fShiftZ

			arr.SetValue(pn, i)
		Next i

		moMesh.VertexBuffer.Unlock()
		arr = Nothing

		'And now for tank turrets
		If mbTankMesh = True AndAlso moTurretMesh Is Nothing = False Then
			ranks(0) = moTurretMesh.NumberVertices
			arr = moTurretMesh.VertexBuffer.Lock(0, (New CustomVertex.PositionNormalTextured()).GetType(), LockFlags.None, ranks)
			For i = 0 To arr.Length - 1
				Dim pn As Direct3D.CustomVertex.PositionNormalTextured = CType(arr.GetValue(i), CustomVertex.PositionNormalTextured)

				If fScaleX <> 0 Then pn.X *= fScaleX
				If fScaleY <> 0 Then pn.Y *= fScaleY
				If fScaleZ <> 0 Then pn.Z *= fScaleZ

				pn.X += fShiftX
				pn.Y += fShiftY
				pn.Z += fShiftZ

				arr.SetValue(pn, i)
			Next i
			moTurretMesh.VertexBuffer.Unlock()
			arr = Nothing
		End If

		txtShiftX.Text = "0"
		txtShiftY.Text = "0"
		txtShiftZ.Text = "0"
		txtScaleX.Text = "1"
		txtScaleY.Text = "1"
		txtScaleZ.Text = "1"
	End Sub

	Private Sub UpdateEngineFX()
		If moPEngine Is Nothing Then Exit Sub
		If lstEngineFX.SelectedIndex = -1 Then Exit Sub

		If mbIgnoreEngineFXEvents = False Then
			Dim lVal As Int32 = CInt(Val(Mid$(lstEngineFX.SelectedItem.ToString, 5)))
			moPEngine.SetEmitterLoc(lVal, CSng(Val(txtEngineOffX.Text)), CSng(Val(txtEngineOffY.Text)), CSng(Val(txtEngineOffZ.Text)))
		End If
	End Sub

	Private Sub UpdateBurnFX()
		If moPEngine Is Nothing Then Return
		If lstBurn.SelectedItem Is Nothing Then Return

		If mbIgnoreBurnFXEvents = False Then

			With CType(lstBurn.SelectedItem, BurnFXData)
				.LocX = CInt(Val(txtBurnOffX.Text))
				.LocY = CInt(Val(txtBurnOffY.Text))
				.LocZ = CInt(Val(txtBurnOffZ.Text))

				If optFrontBurn.Checked = True Then
					.lSide = 0
				ElseIf optLeftBurn.Checked = True Then
					.lSide = 1
				ElseIf optRearBurn.Checked = True Then
					.lSide = 2
				Else : .lSide = 3
				End If

				Dim lType As Int32
				Dim lModVal As Int32
				If mnuAtmosBurn.Checked = True Then
					lType = ParticleFX.ParticleEngine.EmitterType.eSmokeyFireEmitter
				Else : lType = ParticleFX.ParticleEngine.EmitterType.eFireEmitter
				End If

				If optPosX.Checked = True Then
					lModVal = ParticleFX.ParticleEngine.EmitterType.eFireEmitterMod_PlusX
				ElseIf optPosZ.Checked = True Then
					lModVal = ParticleFX.ParticleEngine.EmitterType.eFireEmitterMod_PlusZ
				ElseIf optMinusX.Checked = True Then
					lModVal = ParticleFX.ParticleEngine.EmitterType.eFireEmitterMod_MinusX
				ElseIf optMinusY.Checked = True Then
					lModVal = ParticleFX.ParticleEngine.EmitterType.eFireEmitterMod_MinusY
				ElseIf optMinusZ.Checked = True Then
					lModVal = ParticleFX.ParticleEngine.EmitterType.eFireEmitterMod_MinusZ
				End If

				.lEmitterType = CType(lType Or lModVal, ParticleFX.ParticleEngine.EmitterType)
				moPEngine.SetEmitterModType(.lParticleFXID, lModVal)
				moPEngine.SetEmitterLoc(.lParticleFXID, .LocX, .LocY, .LocZ)
			End With

			' Dim lVal As Int32 = CInt(Val(Mid$(lstBurn.SelectedItem.ToString, 5)))
			'moPEngine.SetEmitterLoc(lVal, CSng(Val(txtBurnOffX.Text)), CSng(Val(txtBurnOffY.Text)), CSng(Val(txtBurnOffZ.Text)))

		End If
	End Sub

	Private Sub UpdateFireFromFX()
		If lstFireFroms.SelectedIndex = -1 Then Return

		If mbIgnoreFireFromEvents = False Then
			Dim sTemp As String = lstFireFroms.SelectedItem.ToString
			Dim ySide As Byte

			optFireFromFront.Checked = False : optFireFromLeft.Checked = False : optFireFromRear.Checked = False : optFireFromRight.Checked = False
			If InStr(sTemp, "(F)", CompareMethod.Binary) <> 0 Then
				ySide = 0
			ElseIf InStr(sTemp, "(L)", CompareMethod.Binary) <> 0 Then
				ySide = 1
			ElseIf InStr(sTemp, "(A)", CompareMethod.Binary) <> 0 Then
				ySide = 2
			Else
				ySide = 3
			End If

			sTemp = Mid$(sTemp, 1, sTemp.Length - 4)
			On Error Resume Next
			Dim vecTemp As Vector3 = CType(mcolFireFromLocs(ySide)(sTemp), Vector3)
			mcolFireFromLocs(ySide).Remove(sTemp)
			vecTemp.X = CSng(Val(txtFireFromOffX.Text))
			vecTemp.Y = CSng(Val(txtFireFromOffY.Text))
			vecTemp.Z = CSng(Val(txtFireFromOffZ.Text))
			mcolFireFromLocs(ySide).Add(vecTemp, sTemp)
			On Error GoTo 0
		End If
	End Sub

	Private Sub hscrRotate_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles hscrRotate.ValueChanged
		Me.Text = "YAW: " & hscrYaw.Value & ", ROTATE: " & hscrRotate.Value
	End Sub

	Private Sub hscrYaw_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles hscrYaw.ValueChanged
		Me.Text = "YAW: " & hscrYaw.Value & ", ROTATE: " & hscrRotate.Value
	End Sub

	Private Sub SaveData(ByVal bXFile As Boolean, ByVal bDetails As Boolean)
		Dim oRes As DialogResult
		Dim sFile As String
		Dim X As Int32
		Dim lAdj() As Int32

		sFile = AppDomain.CurrentDomain.BaseDirectory
		If sFile.EndsWith("\") = False Then sFile &= "\"

		Dim oINI As InitFile = New InitFile(sFile & "Models.dat")
		Dim sModelHdr As String = "MODEL_" & Val(txtModelID.Text)

		Dim sTemp As String = ""
		Dim vecLoc As Vector3

		Dim lVal As Int32

		'Get our file to save to
		If bXFile = True Then
			savDlg.Filter = "X Files *.x|*.x"
			oRes = savDlg.ShowDialog(Me)
			sFile = savDlg.FileName
		Else
			sFile = msCurrent
		End If

		If (oRes = Windows.Forms.DialogResult.OK) Or bXFile = False Then

			'Check if the ModelID is in use already
			If bDetails = True Then
				sTemp = oINI.GetString(sModelHdr, "FileName", "")
				If sTemp <> "" AndAlso UCase$(sTemp) <> UCase$(Mid$(sFile, InStrRev(sFile, "\") + 1)) Then
					If MsgBox("Model ID " & Val(txtModelID.Text) & " already has data in" & vbCrLf & _
					  "it for File: '" & sTemp & "'." & vbCrLf & vbCrLf & "Do you want to overwrite the current data?", MsgBoxStyle.YesNo, "Overwrite Confirmation") = MsgBoxResult.No Then
						oINI = Nothing
						Exit Sub
					End If
				End If
			End If

			If bXFile = True Then

				'Array of three Int32 values per face to be filled with adjacent face indices. 
				'The size of this array must be at least 3 * BaseMesh.NumberFaces.
				ReDim lAdj(3 * moMesh.NumberFaces)

				'Now, generate our adjacency
				moMesh.GenerateAdjacency(0, lAdj)

				'Dim sErrors As String = ""
				'moMesh.Validate(lAdj, sErrors)

				'moMesh = Mesh.Clean(CleanType.BowTies, moMesh, lAdj, lAdj)


				'Now, go back and set our materials
				For X = 0 To mtrlBuffer.Length - 1
					mtrlBuffer(X).Material3D = Materials(X)
				Next X

				'save it
				moMesh.Save(sFile, lAdj, mtrlBuffer, XFileFormat.Binary)
			End If

			'Ok, now... save off our extended data...
			If bDetails = True Then
				sTemp = Mid$(sFile, InStrRev(sFile, "\") + 1)
				oINI.WriteString(sModelHdr, "FileName", sTemp)
				oINI.WriteString(sModelHdr, "PlanetYAdjust", txtPlanetYAdjust.Text)
				oINI.WriteString(sModelHdr, "YMidPoint", txtMidY.Text)
				oINI.WriteString(sModelHdr, "ShieldSphereSize", txtShieldSphereSize.Text)
				oINI.WriteString(sModelHdr, "ShieldSphereStretchX", txtShieldStretchX.Text)
				oINI.WriteString(sModelHdr, "ShieldSphereStretchY", txtShieldStretchY.Text)
				oINI.WriteString(sModelHdr, "ShieldSphereStretchZ", txtShieldStretchZ.Text)
				oINI.WriteString(sModelHdr, "ShieldSphereShiftX", txtShieldShiftX.Text)
				oINI.WriteString(sModelHdr, "ShieldSphereShiftY", txtShieldShiftY.Text)
				oINI.WriteString(sModelHdr, "ShieldSphereShiftZ", txtShieldShiftZ.Text)
				oINI.WriteString(sModelHdr, "EngineFXCnt", lstEngineFX.Items.Count.ToString)
				For X = 0 To lstEngineFX.Items.Count - 1
					lVal = CInt(Val(Mid$(lstEngineFX.Items(X).ToString, 5)))
					vecLoc = moPEngine.GetEmitterLoc(lVal)
					oINI.WriteString(sModelHdr, "Engine" & X & "_OffsetX", vecLoc.X.ToString)
					oINI.WriteString(sModelHdr, "Engine" & X & "_OffsetY", vecLoc.Y.ToString)
					oINI.WriteString(sModelHdr, "Engine" & X & "_OffsetZ", vecLoc.Z.ToString)
					oINI.WriteString(sModelHdr, "Engine" & X & "_PCnt", moPEngine.GetEmitterPCnt(lVal).ToString)
				Next X



				Dim lBurnFXCnt As Int32 = 0
				For X = 0 To mlBurnFXUB
					If myBurnFXUsed(X) <> 0 Then lBurnFXCnt += 1
				Next X

				oINI.WriteString(sModelHdr, "BurnFXCnt", lBurnFXCnt.ToString)
				For X = 0 To mlBurnFXUB
					With moBurnFX(X)
						oINI.WriteString(sModelHdr, "BurnSide_" & X, .lSide.ToString)
						oINI.WriteString(sModelHdr, "Burn" & X & "_X", .LocX.ToString)
						oINI.WriteString(sModelHdr, "Burn" & X & "_Y", .LocY.ToString)
						oINI.WriteString(sModelHdr, "Burn" & X & "_Z", .LocZ.ToString)
						oINI.WriteString(sModelHdr, "Burn" & X & "_PCnt", .lPCnt.ToString)
						oINI.WriteString(sModelHdr, "Burn" & X & "_Type", CInt(.lEmitterType).ToString)
					End With
				Next X


				If mcolFireFromLocs Is Nothing = False Then
					For X = 0 To 3
						If mcolFireFromLocs(X) Is Nothing = False Then
							oINI.WriteString(sModelHdr, "FireFromCnt_" & X, mcolFireFromLocs(X).Count.ToString)
							Dim lFFIdx As Int32 = 0
							For Each vecLoc In mcolFireFromLocs(X)
								oINI.WriteString(sModelHdr, "FireFromLoc_" & X & "_" & lFFIdx & "_X", vecLoc.X.ToString)
								oINI.WriteString(sModelHdr, "FireFromLoc_" & X & "_" & lFFIdx & "_Y", vecLoc.Y.ToString)
								oINI.WriteString(sModelHdr, "FireFromLoc_" & X & "_" & lFFIdx & "_Z", vecLoc.Z.ToString)
								lFFIdx += 1
							Next
						End If
					Next X
				End If

				oINI.WriteString(sModelHdr, "BlinkCnt", (mlBL_LocUB + 1).ToString)
				For X = 0 To mlBL_LocUB
					oINI.WriteString(sModelHdr, "Blink_" & X & "_X", mvecBL_Loc(X).X.ToString)
					oINI.WriteString(sModelHdr, "Blink_" & X & "_Y", mvecBL_Loc(X).Y.ToString)
					oINI.WriteString(sModelHdr, "Blink_" & X & "_Z", mvecBL_Loc(X).Z.ToString)
					oINI.WriteString(sModelHdr, "Blink_" & X & "_R", mcolBL(X).R.ToString)
					oINI.WriteString(sModelHdr, "Blink_" & X & "_G", mcolBL(X).G.ToString)
					oINI.WriteString(sModelHdr, "Blink_" & X & "_B", mcolBL(X).B.ToString)
					oINI.WriteString(sModelHdr, "Blink_" & X & "_AChg", mfAlphaChg(X).ToString)
				Next X

				oINI.WriteString(sModelHdr, "DeathSeqSize_X", txtDS_SizeX.Text)
				oINI.WriteString(sModelHdr, "DeathSeqSize_Y", txtDS_SizeY.Text)
				oINI.WriteString(sModelHdr, "DeathSeqSize_Z", txtDS_SizeZ.Text)
				oINI.WriteString(sModelHdr, "DeathSeqExpCnt", txtDS_ExpCnt.Text)
				If chkFinale.Checked = True Then oINI.WriteString(sModelHdr, "DeathSeqFinale", "1") Else oINI.WriteString(sModelHdr, "DeathSeqFinale", "0")
				oINI.WriteString(sModelHdr, "DeathSeqWAV_Cnt", lstSFX.Items.Count.ToString)
				For X = 0 To lstSFX.Items.Count - 1
					Dim sItem As String = CType(lstSFX.Items(X), String)
					oINI.WriteString(sModelHdr, "DeathSeqWav_" & X.ToString, sItem)
				Next X


			End If

			'Now, for tanks...
			If mbTankMesh = True AndAlso moTurretMesh Is Nothing = False Then
				Dim sTurrFile As String
				If InStr(sTemp, "_B.X", CompareMethod.Text) <> 0 Then
					sTemp = Replace$(sTemp, "_B.X", "_T.X", , , CompareMethod.Text)
					sTurrFile = Replace$(sFile, "_B.X", "_T.X", , , CompareMethod.Text)
				Else
					sTemp = Replace$(sTemp, ".X", "_T.X", , , CompareMethod.Text)
					sTurrFile = Replace$(sFile, ".X", "_T.X", , , CompareMethod.Text)
				End If
				If bDetails = True Then
					oINI.WriteString(sModelHdr, "TurretFileName", sTemp)
					oINI.WriteString(sModelHdr, "TurretOffsetZ", txtTurretOffZ.Text)
				End If

				If bXFile = True Then
					ReDim lAdj(3 * moTurretMesh.NumberFaces)
					moTurretMesh.GenerateAdjacency(0, lAdj)
					moTurretMesh.Save(sTurrFile, lAdj, mtrlBuffer, XFileFormat.Binary)
				End If

			End If
		End If

		oINI = Nothing

		MsgBox("File Saved Successfully!", MsgBoxStyle.OkOnly Or MsgBoxStyle.Information, "Success")
	End Sub

	Private Sub SetFireFromMaterials()
		With matFireFrom(0)		'Front
			.Emissive = System.Drawing.Color.FromArgb(255, 255, 0, 255)
			.Diffuse = .Emissive : .Specular = .Emissive : .Ambient = .Emissive
		End With
		With matFireFrom(1)		'Left
			.Emissive = System.Drawing.Color.FromArgb(255, 255, 255, 0)
			.Diffuse = .Emissive : .Specular = .Emissive : .Ambient = .Emissive
		End With
		With matFireFrom(2)		'Rear
			.Emissive = System.Drawing.Color.FromArgb(255, 0, 255, 255)
			.Diffuse = .Emissive : .Specular = .Emissive : .Ambient = .Emissive
		End With
		With matFireFrom(3)		'Right
			.Emissive = System.Drawing.Color.FromArgb(255, 255, 0, 0)
			.Diffuse = .Emissive : .Specular = .Emissive : .Ambient = .Emissive
		End With
	End Sub

	Private Sub DrawBlinkingLights()
		Dim colTemp As System.Drawing.Color

		Device.IsUsingEventHandlers = False

		If moBLTex Is Nothing Then
			Dim sFile As String
			sFile = AppDomain.CurrentDomain.BaseDirectory
			If sFile.EndsWith("\") = False Then sFile &= "\"
			sFile &= "Particle.dds"
			moBLTex = TextureLoader.FromFile(moDevice, sFile)
		End If

		With moDevice
			.Transform.World = Matrix.Identity
			.RenderState.ZBufferWriteEnable = False
			.RenderState.Lighting = False
		End With

		Using oTSprite As Sprite = New Sprite(moDevice)
			oTSprite.SetWorldViewLH(Matrix.Identity, moDevice.Transform.View)
			oTSprite.Begin(SpriteFlags.Billboard Or SpriteFlags.AlphaBlend Or SpriteFlags.ObjectSpace)

			Dim lPrevBlnd As Blend = moDevice.RenderState.DestinationBlend
			moDevice.RenderState.DestinationBlend = Blend.One
			For X As Int32 = 0 To mlBL_LocUB
				mfAlpha(X) -= mfAlphaChg(X)
				If mfAlpha(X) < 0 Then mfAlpha(X) = 255.0F
				colTemp = System.Drawing.Color.FromArgb(CInt(mfAlpha(X)), mcolBL(X).R, mcolBL(X).G, mcolBL(X).B)

				oTSprite.Draw(moBLTex, System.Drawing.Rectangle.Empty, New Vector3(16, 0, 16), mvecBL_Loc(X), colTemp)
			Next X

			oTSprite.End()
			oTSprite.Dispose()
			moDevice.RenderState.DestinationBlend = lPrevBlnd
		End Using

		With moDevice
			'Then, reset our device...
			.RenderState.ZBufferWriteEnable = True
			.RenderState.Lighting = True
		End With

		Device.IsUsingEventHandlers = True
	End Sub

	Private Sub dgvBL_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvBL.CellEndEdit
		mlBL_LocUB = dgvBL.RowCount - 2
		ReDim mvecBL_Loc(mlBL_LocUB)
		ReDim mcolBL(mlBL_LocUB)
		ReDim mfAlpha(mlBL_LocUB)
		ReDim mfAlphaChg(mlBL_LocUB)

		On Error Resume Next

		For iRow As Int32 = 0 To mlBL_LocUB
			mvecBL_Loc(iRow).X = CSng(dgvBL(0, iRow).Value.ToString)
			mvecBL_Loc(iRow).Y = CSng(dgvBL(1, iRow).Value.ToString)
			mvecBL_Loc(iRow).Z = CSng(dgvBL(2, iRow).Value.ToString)
			mcolBL(iRow) = System.Drawing.Color.FromArgb(255, CInt(dgvBL(3, iRow).Value.ToString), CInt(dgvBL(4, iRow).Value.ToString), CInt(dgvBL(5, iRow).Value.ToString))
			mfAlpha(iRow) = 255
			mfAlphaChg(iRow) = CSng(dgvBL(6, iRow).Value.ToString)
		Next iRow
	End Sub

	Private Sub btnInvertNormals_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnInvertNormals.Click
		moMesh.ComputeNormals()
		If moTurretMesh Is Nothing = False Then moTurretMesh.ComputeNormals()

		Materials(0).SpecularSharpness = 17

		' Get the original mesh's vertex buffer.
		Dim ranks(0) As Integer
		ranks(0) = moMesh.NumberVertices

		Dim arr As System.Array = moMesh.VertexBuffer.Lock(0, (New CustomVertex.PositionNormalTextured()).GetType(), LockFlags.None, ranks)

		Dim i As Integer

		For i = 0 To arr.Length - 1

			Dim pn As Direct3D.CustomVertex.PositionNormalTextured = CType(arr.GetValue(i), CustomVertex.PositionNormalTextured)
			'pn.Normal = Vector3.Multiply(pn.Normal, -1)
			Dim vecTemp As Vector3 = pn.Normal
			vecTemp.Y *= -1
			If vecTemp.Y < 0 Then vecTemp.Y = Math.Abs(vecTemp.Y) Else vecTemp.Y = vecTemp.Y * -1
			pn.Normal = vecTemp

			arr.SetValue(pn, i)
		Next i

		moMesh.VertexBuffer.Unlock()
		arr = Nothing

		'Now, if we are in Tank View mode, we need to invert the turret too...
		If mbTankMesh = True Then
			ranks(0) = moTurretMesh.NumberVertices
			arr = moTurretMesh.VertexBuffer.Lock(0, (New CustomVertex.PositionNormalTextured()).GetType(), LockFlags.None, ranks)
			For i = 0 To arr.Length - 1
				Dim pn As Direct3D.CustomVertex.PositionNormalTextured = CType(arr.GetValue(i), CustomVertex.PositionNormalTextured)
				pn.Normal = Vector3.Multiply(pn.Normal, -1)
				arr.SetValue(pn, i)
			Next i
			moTurretMesh.VertexBuffer.Unlock()
			arr = Nothing
		End If

	End Sub

	Private Function GetEstimatedHull() As Int32
		Dim fLowX As Single = 0.0F
		Dim lLowXCnt As Int32 = 0
		Dim fLowY As Single = 0.0F
		Dim lLowYCnt As Int32 = 0
		Dim fLowZ As Single = 0.0F
		Dim lLowZCnt As Int32 = 0
		Dim fHighX As Single = 0.0F
		Dim lHighXCnt As Int32 = 0
		Dim fHighY As Single = 0.0F
		Dim lHighYCnt As Int32 = 0
		Dim fHighZ As Single = 0.0F
		Dim lHighZCnt As Int32 = 0

		' Get the original mesh's vertex buffer.
		Dim ranks(0) As Integer
		ranks(0) = moMesh.NumberVertices

		Dim arr As System.Array = moMesh.VertexBuffer.Lock(0, (New CustomVertex.PositionNormalTextured()).GetType(), LockFlags.None, ranks)

		Dim i As Integer

		For i = 0 To arr.Length - 1

			Dim pn As Direct3D.CustomVertex.PositionNormalTextured = CType(arr.GetValue(i), CustomVertex.PositionNormalTextured)

			If pn.X < 0 Then
				fLowX += pn.X
				lLowXCnt += 1
			Else
				fHighX += pn.X
				lHighXCnt += 1
			End If

			If pn.Y < 0 Then
				fLowY += pn.Y
				lLowYCnt += 1
			Else
				fHighY += pn.Y
				lHighYCnt += 1
			End If

			If pn.Z < 0 Then
				fLowZ += pn.Z
				lLowZCnt += 1
			Else
				fHighZ += pn.Z
				lHighZCnt += 1
			End If
		Next i

		moMesh.VertexBuffer.Unlock()
		arr = Nothing

		''Now, if we are in Tank View mode, we need to invert the turret too...
		'If mbTankMesh = True Then
		'    ranks(0) = moTurretMesh.NumberVertices
		'    arr = moTurretMesh.VertexBuffer.Lock(0, (New CustomVertex.PositionNormalTextured()).GetType(), LockFlags.None, ranks)
		'    For i = 0 To arr.Length - 1
		'        Dim pn As Direct3D.CustomVertex.PositionNormalTextured = CType(arr.GetValue(i), CustomVertex.PositionNormalTextured)
		'        pn.Z *= -1
		'        arr.SetValue(pn, i)
		'    Next i
		'    moTurretMesh.VertexBuffer.Unlock()
		'    arr = Nothing
		'End If


		If lLowXCnt <> 0 Then fLowX = (fLowX / lLowXCnt)
		fHighX = (fHighX / lHighXCnt)
		If lLowYCnt <> 0 Then fLowY = (fLowY / lLowYCnt)
		fHighY = (fHighY / lHighYCnt)
		If lLowZCnt <> 0 Then fLowZ = (fLowZ / lLowZCnt)
		fHighZ = (fHighZ / lHighZCnt)

		Dim fResult As Single = (Math.Abs(fLowX) + Math.Abs(fHighX)) * (Math.Abs(fLowY) + Math.Abs(fHighY)) * (Math.Abs(fLowZ) + Math.Abs(fHighZ)) / 100.0F

		Return CInt(fResult)
	End Function

	Private Sub FillVertList()

		dgvVerts.RowCount = moMesh.NumberVertices

		' Get the original mesh's vertex buffer.
		Dim ranks(0) As Integer
		ranks(0) = moMesh.NumberVertices

		Dim arr As System.Array = moMesh.VertexBuffer.Lock(0, (New CustomVertex.PositionNormalTextured()).GetType(), LockFlags.None, ranks)

		Dim i As Integer

		For i = 0 To arr.Length - 1

			Dim pn As Direct3D.CustomVertex.PositionNormalTextured = CType(arr.GetValue(i), CustomVertex.PositionNormalTextured)

			dgvVerts(0, i).Value = pn.X.ToString
			dgvVerts(1, i).Value = pn.Y.ToString
			dgvVerts(2, i).Value = pn.Z.ToString
			dgvVerts(3, i).Value = pn.Tu.ToString
			dgvVerts(4, i).Value = pn.Tv.ToString
		Next i

		moMesh.VertexBuffer.Unlock()
		arr = Nothing

	End Sub

	Private Sub dgvVerts_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvVerts.CellEndEdit
		Dim lVar As Int32 = e.RowIndex

		Dim ranks(0) As Integer
		ranks(0) = moMesh.NumberVertices
		Dim arr As System.Array = moMesh.VertexBuffer.Lock(0, (New CustomVertex.PositionNormalTextured()).GetType(), LockFlags.None, ranks)

		Dim pn As Direct3D.CustomVertex.PositionNormalTextured = CType(arr.GetValue(lVar), CustomVertex.PositionNormalTextured)
		pn.X = CSng(dgvVerts(0, lVar).Value)
		pn.Y = CSng(dgvVerts(1, lVar).Value)
		pn.Z = CSng(dgvVerts(2, lVar).Value)
		pn.Tu = CSng(dgvVerts(3, lVar).Value)
		pn.Tv = CSng(dgvVerts(4, lVar).Value)
		arr.SetValue(pn, lVar)

		moMesh.VertexBuffer.Unlock()
		arr = Nothing
	End Sub

	Private Sub dgvVerts_RowEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvVerts.RowEnter
		'Private mvecHitLoc As Vector3
		'Private mbRenderHitloc As Boolean = False
		mvecHitLoc.X = CSng(dgvVerts(0, e.RowIndex).Value)
		mvecHitLoc.Y = CSng(dgvVerts(1, e.RowIndex).Value)
		mvecHitLoc.Z = CSng(dgvVerts(2, e.RowIndex).Value)
		mbRenderHitloc = True
	End Sub

	Private Sub btnPMesh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPMesh.Click
		Static xlPrevFireTo As Int32 = 0

		If moFireTo Is Nothing Then
			ReDim moFireTo(3)
			moFireTo(0).Z = 1000
			moFireTo(1).X = -1000
			moFireTo(2).Z = -1000
			moFireTo(3).X = 1000
		Else
			xlPrevFireTo += 1
			If xlPrevFireTo = 4 Then xlPrevFireTo = 0
			Dim fAngle As Single = LineAngleDegrees(0, 0, CInt(moFireTo(xlPrevFireTo).X), CInt(moFireTo(xlPrevFireTo).Z))


			Dim fMyAngle As Single = (Me.hscrRotate.Value / 10.0F) - 90.0F
			fAngle -= fMyAngle
			If fAngle > 360 Then fAngle -= 360
			If fAngle < 0 Then fAngle += 360




			Dim ySide As Byte = AngleToQuadrant(CInt(fAngle))
			Me.Text = "Fire From Side " & ySide & ", xlPrevFireTo: " & xlPrevFireTo

		End If
	End Sub
	Public Const gdPi As Single = 3.14159265358979
	Public Const gdHalfPie As Single = gdPi / 2.0F
	Public Const gdPieAndAHalf As Single = gdPi * 1.5F
	Public Const gdTwoPie As Single = gdPi * 2.0F
	Public Const gdDegreePerRad As Single = 180.0F / gdPi

	Public Function LineAngleDegrees(ByVal lX1 As Int32, ByVal lY1 As Int32, ByVal lX2 As Int32, ByVal lY2 As Int32) As Single
		Dim dDeltaX As Single
		Dim dDeltaY As Single
		Dim dAngle As Single

		dDeltaX = lX2 - lX1
		dDeltaY = lY2 - lY1

		If dDeltaX = 0 Then		'vertical
			If dDeltaY < 0 Then
				dAngle = gdHalfPie
			Else
				dAngle = gdPieAndAHalf
			End If
		ElseIf dDeltaY = 0 Then		'horizontal
			If dDeltaX < 0 Then
				dAngle = gdPi
			Else
				dAngle = 0
			End If
		Else	'angled
			dAngle = CSng(System.Math.Atan(System.Math.Abs(dDeltaY / dDeltaX)))

			'If dDeltaX > -1 AndAlso dDeltaY < 0 Then
			'    dAngle = gdTwoPie - dAngle
			'ElseIf dDeltaX < 0 AndAlso dDeltaY < 0 Then
			'    dAngle = gdPi + dAngle
			'ElseIf dDeltaX < 0 AndAlso dDeltaY > -1 Then
			'    dAngle = gdPi - dAngle
			'End If


			'Correct for VB's reversed Y... VB Upper Right is ok... but the other quads are not
			If dDeltaX > -1 And dDeltaY > -1 Then		'VB Lower Right
				dAngle = gdTwoPie - dAngle
			ElseIf dDeltaX < 0 And dDeltaY > -1 Then	'VB Lower Left
				dAngle = gdPi + dAngle
			ElseIf dDeltaX < 0 And dDeltaY < 0 Then		'VB Upper Left
				dAngle = gdPi - dAngle
			End If
		End If

		'Not sure this is suppose to be CINT
		'Return CInt(dAngle * gdDegreePerRad)
		Return (dAngle * gdDegreePerRad)

	End Function

	Public Function AngleToQuadrant(ByVal lAngle As Int32) As Byte
		'here, we will return the quadrant from the angle
		Select Case lAngle
			Case Is < 45, Is > 315
				Return 0
			Case Is < 135
				'Return UnitArcs.eLeftArc
				Return 3
			Case Is < 225
				Return 2
			Case Else
				'Return UnitArcs.eRightArc
				Return 1
		End Select
	End Function

	Private Sub mnuShaders_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuShaders.Click
		mbRenderShader = Not mbRenderShader
		mnuShaders.Checked = mbRenderShader

	End Sub
	Private Sub mnuGlow_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuGlow.Click
		mbRenderGlowFX = Not mbRenderGlowFX
		mnuGlow.Checked = mbRenderGlowFX
	End Sub

	Private Sub txtSpecSharp_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSpecSharp.TextChanged
		Dim fValue As Single = CSng(Val(txtSpecSharp.Text))
		If Materials Is Nothing = False Then
			For X As Int32 = 0 To Materials.GetUpperBound(0)
				Materials(X).SpecularSharpness = fValue
			Next X
		End If
	End Sub

	Private Sub mnuComputeNormals_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuComputeNormals.Click
		If moMesh Is Nothing = False Then moMesh.ComputeNormals()
		If moTurretMesh Is Nothing = False Then moTurretMesh.ComputeNormals()
	End Sub


	Private Structure TangentVert
		Public X As Single
		Public Y As Single
		Public Z As Single
		Public NX As Single
		Public NY As Single
		Public NZ As Single
		Public tu As Single
		Public tv As Single
		Public TX As Single
		Public TY As Single
		Public TZ As Single
		Public BX As Single
		Public BY As Single
		Public BZ As Single
	End Structure

	Private Structure CarrierVert
		Public X As Single
		Public Y As Single
		Public Z As Single
		Public NX As Single
		Public NY As Single
		Public NZ As Single
		Public tu As Single
		Public tv As Single
		Public tu2 As Single
		Public tv2 As Single
	End Structure

	Private Sub DoTangentCalcs(ByRef oMesh As Mesh)
		Dim elems(5) As VertexElement
		'position, normal, texture coords, tangent, binormal.

		elems(0) = New VertexElement(0, 0, DeclarationType.Float3, DeclarationMethod.Default, DeclarationUsage.Position, 0)
		elems(1) = New VertexElement(0, 12, DeclarationType.Float3, DeclarationMethod.Default, DeclarationUsage.Normal, 0)
		elems(2) = New VertexElement(0, 24, DeclarationType.Float2, DeclarationMethod.Default, DeclarationUsage.TextureCoordinate, 0)
		elems(3) = New VertexElement(0, 32, DeclarationType.Float3, DeclarationMethod.Default, DeclarationUsage.Tangent, 0)
		elems(4) = New VertexElement(0, 44, DeclarationType.Float3, DeclarationMethod.Default, DeclarationUsage.BiNormal, 0)
		elems(5) = VertexElement.VertexDeclarationEnd

		oMesh.ComputeNormals()

		'Dim clonedmesh As Mesh = oMesh.Clone(MeshFlags.Managed, elems, moDevice)
		Dim clonedmesh As Mesh = New Mesh(oMesh.NumberFaces, oMesh.NumberVertices, MeshFlags.Managed, elems, moDevice)

		Dim ranks(0) As Int32
		ranks(0) = oMesh.NumberVertices

		Dim arr As System.Array = oMesh.VertexBuffer.Lock(0, (New CustomVertex.PositionNormalTextured).GetType(), LockFlags.None, ranks)
		'Dim arr As System.Array = oMesh.VertexBuffer.Lock(0, (New CarrierVert).GetType(), LockFlags.None, ranks)
		Dim oDest As System.Array = clonedmesh.VertexBuffer.Lock(0, (New TangentVert).GetType(), LockFlags.None, ranks)

		Dim uVerts(arr.Length - 1) As TangentVert

		'Load our entire vertex buffer into an array
		For i As Int32 = 0 To arr.Length - 1
			Dim pnSrc As CustomVertex.PositionNormalTextured = CType(arr.GetValue(i), CustomVertex.PositionNormalTextured)
			'Dim pnSrc As CarrierVert = CType(arr.GetValue(i), CarrierVert)
			With uVerts(i)
				.X = pnSrc.X
				.Y = pnSrc.Y
				.Z = pnSrc.Z
				.NX = pnSrc.Nx
				.NY = pnSrc.Ny
				.NZ = pnSrc.Nz
				.tu = pnSrc.Tu
				.tv = pnSrc.Tv
				.BX = 0 : .BY = 0 : .BZ = 0
				.TX = 0 : .TY = 0 : .TZ = 0
			End With
		Next i

		'omesh.IndexBuffer.Description.Is16BitIndices - tells me if it is 16 bit or 32 bit
		Dim oIndexArr As Array
		If oMesh.IndexBuffer.Description.Is16BitIndices = True Then
			Dim lRanks2(0) As Int32
			lRanks2(0) = oMesh.NumberFaces * 3
			oIndexArr = oMesh.LockIndexBuffer(New Short().GetType, LockFlags.None, lRanks2)
		Else
			Dim lRanks2(0) As Int32
			lRanks2(0) = oMesh.NumberFaces * 3
			oIndexArr = oMesh.LockIndexBuffer(New Int32().GetType, LockFlags.None, lRanks2)
		End If

		'Assume it is a triangle list?
		For i As Int32 = 0 To oIndexArr.Length - 3 'Step 3
			Dim iIndex1 As Int16 = CShort(oIndexArr.GetValue(i))
			Dim iIndex2 As Int16 = CShort(oIndexArr.GetValue(i + 1))
			Dim iIndex3 As Int16 = CShort(oIndexArr.GetValue(i + 2))

			'Now, get our vecs and UVs
			Dim vA, vB, vC As Vector3
			Dim uva, uvb, uvc As Vector2
			With uVerts(iIndex1)
				vA.X = .X : vA.Y = .Y : vA.Z = .Z
				uva.X = .tu : uva.Y = .tv
			End With
			With uVerts(iIndex2)
				vB.X = .X : vB.Y = .Y : vB.Z = .Z
				uvb.X = .tu : uvb.Y = .tv
			End With
			With uVerts(iIndex3)
				vC.X = .X : vC.Y = .Y : vC.Z = .Z
				uvc.X = .tu : uvc.Y = .tv
			End With

			Dim fX1 As Single = vB.X - vA.X
			Dim fX2 As Single = vC.X - vA.X
			Dim fY1 As Single = vB.Y - vA.Y
			Dim fY2 As Single = vC.Y - vA.Y
			Dim fZ1 As Single = vB.Z - vA.Z
			Dim fZ2 As Single = vC.Z - vA.Z
			Dim fS1 As Single = uvb.X - uva.X
			Dim fS2 As Single = uvc.X - uva.X
			Dim fT1 As Single = uvb.Y - uva.Y
			Dim fT2 As Single = uvc.Y - uva.Y

			Dim fDenom As Single = fS1 * fT2 - fS2 * fT1
			If Math.Abs(fDenom) < 0.0001 Then Continue For

			Dim fR As Single = 1.0F / fDenom
			Dim vecS As Vector3
			vecS.X = (fT2 * fX1 - fT1 * fX2) * fR
			vecS.Y = (fT2 * fY1 - fT1 * fY2) * fR
			vecS.Z = (fT2 * fZ1 - fT1 * fZ2) * fR

			Dim vecT As Vector3
			vecT.X = (fS1 * fX2 - fS2 * fX1) * fR
			vecT.Y = (fS1 * fY2 - fS2 * fY1) * fR
			vecT.Z = (fS1 * fZ2 - fS2 * fZ1) * fR

			With uVerts(iIndex1)
				.TX += vecS.X : .TY += vecS.Y : .TZ += vecS.Z
				.BX += vecT.X : .BY += vecT.Y : .BZ += vecT.Z
			End With
			With uVerts(iIndex2)
				.TX += vecS.X : .TY += vecS.Y : .TZ += vecS.Z
				.BX += vecT.X : .BY += vecT.Y : .BZ += vecT.Z
			End With
			With uVerts(iIndex3)
				.TX += vecS.X : .TY += vecS.Y : .TZ += vecS.Z
				.BX += vecT.X : .BY += vecT.Y : .BZ += vecT.Z
			End With
		Next i

		For i As Int32 = 0 To arr.Length - 1
			With uVerts(i)
				Dim vecTemp As Vector3 = New Vector3(.TX, .TY, .TZ)
				vecTemp.Normalize()
				.TX = vecTemp.X : .TY = vecTemp.Y : .TZ = vecTemp.Z
				vecTemp = New Vector3(.BX, .BY, .BZ)
				vecTemp.Normalize()
				.BX = vecTemp.X : .BY = vecTemp.Y : .BZ = vecTemp.Z
			End With

			'v.N = v.normal;
			'v.T = v.T - v.normal * mDot(v.normal, v.T);
			'v.T.normalize();

			'Point3F b;
			'mCross(v.normal, v.T, &b);
			'b *= (mDot(b, v.B) < 0.0F) ? -1.0F : 1.0F;
			'v.B = b;

			oDest.SetValue(uVerts(i), i)
		Next i

		ranks(0) = oMesh.NumberFaces * 3
		clonedmesh.IndexBuffer.SetData(oIndexArr, 0, LockFlags.None)

		oMesh.UnlockIndexBuffer()
		oMesh.VertexBuffer.Unlock()
		clonedmesh.VertexBuffer.Unlock()

		'oMesh.Dispose()
		oMesh = Nothing
		oMesh = clonedmesh
		clonedmesh = Nothing


		' Optimize the mesh for this graphics card's vertex cache 
		' so when rendering the mesh's triangle list the vertices will 
		' cache hit more often so it won't have to re-execute the vertex shader 
		' on those vertices so it will improve perf.     
		Dim adj() As Int32 = oMesh.ConvertPointRepsToAdjacency(CType(Nothing, GraphicsStream))
		oMesh.OptimizeInPlace(MeshFlags.OptimizeVertexCache, adj)
		Erase adj
	End Sub

	'TODO: Added the bOffset flag to this function as an optional. We want to remove this once the models all default to 0
	Public Function CreateTexturedSphere(ByVal fRadius As Single, ByVal lSlices As Int32, ByVal lStacks As Int32, ByVal lWrapCount As Int32, Optional ByVal bOffsetY As Boolean = False) As Mesh

		Device.IsUsingEventHandlers = False

		Dim oTemp As Mesh = Mesh.Sphere(moDevice, fRadius, lSlices, lStacks)
		oTemp.ComputeNormals()
		Dim TexturedObject As Mesh = New Mesh(oTemp.NumberFaces, oTemp.NumberVertices, MeshFlags.Managed, CustomVertex.PositionNormalTextured.Format, moDevice)
		'Dim TexturedObject As Mesh = oTemp.Clone(MeshFlags.Managed, VertexFormats.Position Or VertexFormats.Normal Or VertexFormats.Texture0, moDevice)

		' Get the original mesh's vertex buffer.
		Dim ranks(0) As Integer
		ranks(0) = oTemp.NumberVertices
		Dim arr As System.Array = oTemp.VertexBuffer.Lock(0, (New CustomVertex.PositionNormal()).GetType(), LockFlags.None, ranks)

		' Set the vertex buffer
		Dim data As System.Array = TexturedObject.VertexBuffer.Lock(0, (New CustomVertex.PositionNormalTextured()).GetType(), LockFlags.None, ranks)

		Dim phi As Single
		Dim u As Single
		Dim i As Integer

		For i = 0 To arr.Length - 1
			Dim pn As Direct3D.CustomVertex.PositionNormal = CType(arr.GetValue(i), CustomVertex.PositionNormal)
			Dim pnt As Direct3D.CustomVertex.PositionNormalTextured = CType(data.GetValue(i), CustomVertex.PositionNormalTextured)
			pnt.X = pn.X
			pnt.Y = pn.Y
			pnt.Z = pn.Z
			pnt.Nx = pn.Nx
			pnt.Ny = pn.Ny
			pnt.Nz = pn.Nz

			If lWrapCount = 0 Then
				phi = CSng(Math.Acos(pn.Nz))
				pnt.Tv = CSng(phi / Math.PI)
				u = CSng(Math.Acos(Math.Max(Math.Min(pnt.Ny / Math.Sin(phi), 1.0), -1.0)) / (2.0 * Math.PI))
				If pnt.Nx > 0 Then
					pnt.Tu = u
				Else
					pnt.Tu = 1 - u
				End If
			Else
				'TODO: At the moment, lWrapCount only determines if we wrap or not... I will want it to
				'  actually do more wraps with higher counts....
				pnt.Tu = CSng(Math.Asin(pnt.Nx) / gdPi + 0.5)
				pnt.Tv = CSng(Math.Asin(pnt.Ny) / gdPi + 0.5)
			End If

			If bOffsetY = True Then
				pnt.Y += fRadius
			End If

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

	Private Sub btnAddSFX_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddSFX.Click
		If txtSoundEffect.Text Is Nothing = False AndAlso txtSoundEffect.Text.Trim.Length > 0 Then
			lstSFX.Items.Add(txtSoundEffect.Text)
		End If
	End Sub

	Private Sub btnRemoveSFX_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRemoveSFX.Click
		If lstSFX.SelectedItem Is Nothing = False Then
			lstSFX.Items.Remove(lstSFX.SelectedItem)
		End If
	End Sub
End Class

Namespace ParticleFX
    Public Module Globals
        Public Declare Function timeGetTime Lib "winmm.dll" Alias "timeGetTime" () As Int32
    End Module

    Public Class Particle
        Public vecLoc As Vector3
        Public vecAcc As Vector3
        Public vecSpeed As Vector3

        Public mfA As Single
        Public mfR As Single
        Public mfG As Single
        Public mfB As Single

        Public fAChg As Single
        Public fRChg As Single
        Public fGChg As Single
        Public fBChg As Single

        Public ParticleColor As System.Drawing.Color

        Public ParticleActive As Boolean

        Public Sub Update(ByVal fElapsedTime As Single)
            vecLoc.Add(Vector3.Multiply(vecSpeed, fElapsedTime))

            vecSpeed.Add(Vector3.Multiply(vecAcc, fElapsedTime))

            mfA += (fAChg * fElapsedTime)
            mfR += (fRChg * fElapsedTime)
            mfG += (fGChg * fElapsedTime)
            mfB += (fBChg * fElapsedTime)

            If mfA < 0 Then mfA = 0
            If mfA > 255 Then mfA = 255
            If mfR < 0 Then mfR = 0
            If mfR > 255 Then mfR = 255
            If mfG < 0 Then mfG = 0
            If mfG > 255 Then mfG = 255
            If mfB < 0 Then mfB = 0
            If mfB > 255 Then mfB = 255

            ParticleColor = System.Drawing.Color.FromArgb(CInt(mfA), CInt(mfR), CInt(mfG), CInt(mfB))
        End Sub

        Public Sub Reset(ByVal fX As Single, ByVal fY As Single, ByVal fZ As Single, ByVal fXSpeed As Single, ByVal fYSpeed As Single, ByVal fZSpeed As Single, ByVal fXAcc As Single, ByVal fYAcc As Single, ByVal fZAcc As Single, ByVal fR As Single, ByVal fG As Single, ByVal fB As Single, ByVal fA As Single)
            vecLoc.X = fX : vecLoc.Y = fY : vecLoc.Z = fZ
            vecAcc.X = fXAcc : vecAcc.Y = fYAcc : vecAcc.Z = fZAcc
            vecSpeed.X = fXSpeed : vecSpeed.Y = fYSpeed : vecSpeed.Z = fZSpeed
            mfA = fA : mfR = fR : mfG = fG : mfB = fB

            ParticleActive = True
        End Sub

    End Class

    Public Class ParticleEngine

        Public Enum EmitterType As Integer
            eEngineEmitter = 0
            eSmokeyFireEmitter
            eFireEmitter                                'Defaults to horizontal (x-z) going positive Y

            eFireEmitterMod_MinusY = 33554432           'horizontal (x-z) emitter going Negative in the Y column
            eFireEmitterMod_PlusX = 67108864            'vertical (y-z) emitter going positive in the X
            eFireEmitterMod_MinusX = 134217728          'vertical (y-z) emitter going Negative in the X Column
            eFireEmitterMod_PlusZ = 268435456           'vertical (x-y) emitter going positive z
            eFireEmitterMod_MinusZ = 536870912          'vertical (x-y) emitter going negative z
        End Enum

        Private moChildren() As FXEmitter
        Private myChildUsed() As Byte
        Public mlChildrenUB As Int32 = -1

        'The standard texture to use for emitters using this engine object
        Private moTex As Texture
        Private mfParticleSize As Single
        Private moDevice As Device

        Public Sub New(ByRef oDevice As Device, ByVal fParticleSize As Single)
			moDevice = oDevice 
            mfParticleSize = fParticleSize
        End Sub

        Public Sub Render(ByVal bUpdateNoRender As Boolean)
            Dim X As Int32

            If moTex Is Nothing Then
                'moTex = goResMgr.GetTexture("Particle.dds", EpicaResourceManager.eGetTextureType.NoSpecifics)
                Dim sFile As String
                sFile = AppDomain.CurrentDomain.BaseDirectory
                If sFile.EndsWith("\") = False Then sFile &= "\"
                sFile &= "Particle.dds"
                moTex = TextureLoader.FromFile(moDevice, sFile)
            End If

            If mlChildrenUB = -1 Then Exit Sub

            For X = 0 To mlChildrenUB
                If myChildUsed(X) > 0 Then
                    moChildren(X).Update()
                    If moChildren(X).EmitterStopped = True Then myChildUsed(X) = 0
                End If

            Next X

            If bUpdateNoRender = False Then
                'And now render them... first set up our device for renders...
                With moDevice
                    '.Transform.World = Matrix.Identity

                    .SetTextureStageState(0, TextureStageStates.AlphaArgument1, TextureArgument.TextureColor)
                    .SetTextureStageState(0, TextureStageStates.AlphaArgument2, TextureArgument.Diffuse)
                    .SetTextureStageState(0, TextureStageStates.AlphaOperation, Form1.l_Burn_Op)

                    .RenderState.PointSpriteEnable = True
                    .RenderState.PointScaleEnable = True
                    .RenderState.PointSize = mfParticleSize
                    .RenderState.PointScaleA = 0
                    .RenderState.PointScaleB = 0
                    .RenderState.PointScaleC = 0.5F
                    .RenderState.SourceBlend = Blend.SourceAlpha
                    .RenderState.DestinationBlend = Blend.One
                    .RenderState.AlphaBlendEnable = True

                    .RenderState.ZBufferWriteEnable = False

                    .RenderState.Lighting = False

                    .VertexFormat = CustomVertex.PositionColoredTextured.Format
                    '.VertexFormat = moPoints(0).Format
                    .SetTexture(0, moTex)

                    'Now render everything 
                    For X = 0 To mlChildrenUB
                        If myChildUsed(X) > 0 Then
                            .DrawUserPrimitives(PrimitiveType.PointList, moChildren(X).mlParticleUB + 1, moChildren(X).moPoints)
                        End If
                    Next X

                    'Then, reset our device...
                    .RenderState.ZBufferWriteEnable = True
                    .RenderState.Lighting = True

                    .RenderState.SourceBlend = Blend.SourceAlpha
                    .RenderState.DestinationBlend = Blend.InvSourceAlpha
                    .RenderState.AlphaBlendEnable = True

                    .SetTextureStageState(0, TextureStageStates.AlphaArgument1, TextureArgument.TextureColor)
                    .SetTextureStageState(0, TextureStageStates.AlphaArgument2, TextureArgument.Current)
                    .SetTextureStageState(0, TextureStageStates.AlphaOperation, TextureOperation.SelectArg1)
                End With
            End If
        End Sub

        'Public Function AddEmitter(ByVal iType As EmitterType, ByVal oLoc As Vector3, ByVal lPCnt As Int32, ByVal oAttachedObject As RenderObject) As Int32
        '    Dim lRes As Int32

        '    lRes = AddEmitter(iType, oLoc, lPCnt)
        '    'moChildren(lRes).AttachedObject = oAttachedObject
        'End Function

        Public Function AddEmitter(ByVal iType As EmitterType, ByVal oLoc As Vector3, ByVal lPCnt As Int32) As Int32
            Dim X As Int32
            Dim lIdx As Int32 = -1


            For X = 0 To mlChildrenUB
                If myChildUsed(X) = 0 Then
                    lIdx = X
                    Exit For
                End If
            Next X

            If lIdx = -1 Then
                mlChildrenUB += 1
                ReDim Preserve moChildren(mlChildrenUB)
                ReDim Preserve myChildUsed(mlChildrenUB)
                lIdx = mlChildrenUB
            End If

            Dim lBaseType As Int32
            Dim lModValue As Int32

            If (iType And EmitterType.eEngineEmitter) <> 0 Then
                lBaseType = EmitterType.eEngineEmitter
            ElseIf (iType And EmitterType.eFireEmitter) <> 0 Then
                lBaseType = EmitterType.eFireEmitter
            ElseIf (iType And EmitterType.eSmokeyFireEmitter) <> 0 Then
                lBaseType = EmitterType.eSmokeyFireEmitter
            End If
            lModValue = iType - lBaseType

            Select Case lBaseType
                Case EmitterType.eEngineEmitter
                    If lPCnt <= 0 Then lPCnt = 30
                    moChildren(lIdx) = New EngineFX(oLoc, Me, lPCnt)
                Case EmitterType.eFireEmitter
                    If lPCnt <= 0 Then lPCnt = 50
                    moChildren(lIdx) = New FireFX(oLoc, Me, lPCnt)
                    CType(moChildren(lIdx), FireFX).lEmitterMod = CType(lModValue, ParticleEngine.EmitterType)
                Case EmitterType.eSmokeyFireEmitter
                    If lPCnt <= 0 Then lPCnt = 100
                    moChildren(lIdx) = New SmokeyFireFX(oLoc, Me, lPCnt)
            End Select
            myChildUsed(lIdx) = 255


            Return lIdx
        End Function

        Public Sub StopEmitter(ByVal lIndex As Int32)
            If lIndex > -1 AndAlso lIndex <= mlChildrenUB Then moChildren(lIndex).StopEmitter()
        End Sub

        Public Sub MoveEmitter(ByVal lIndex As Int32, ByVal fVecX As Single, ByVal fVecY As Single, ByVal fVecZ As Single)
            If lIndex > -1 AndAlso lIndex <= mlChildrenUB Then moChildren(lIndex).MoveEmitter(fVecX, fVecY, fVecZ)
        End Sub

        Public Function GetEmitterLoc(ByVal lIndex As Int32) As Vector3
            Return moChildren(lIndex).GetAbsolutePosition
        End Function

        Public Function GetEmitterPCnt(ByVal lIndex As Int32) As Int32
            Return moChildren(lIndex).mlParticleUB + 1
        End Function

        Public Sub SetEmitterLoc(ByVal lIndex As Int32, ByVal fLocX As Single, ByVal fLocY As Single, ByVal fLocZ As Single)
            moChildren(lIndex).mvecEmitter.X = fLocX
            moChildren(lIndex).mvecEmitter.Y = fLocY
            moChildren(lIndex).mvecEmitter.Z = fLocZ
        End Sub

        Public Sub SetEmitterModType(ByVal lIndex As Int32, ByVal lModVal As Int32)
            CType(moChildren(lIndex), FireFX).lEmitterMod = CType(lModVal, EmitterType)
        End Sub

        Public Sub ChangeEngineFXColor(ByVal lIndex As Int32, ByVal lR As Int32, ByVal lG As Int32, ByVal lB As Int32)
            Try
                With CType(moChildren(lIndex), EngineFX)
                    .lBaseR = lR
                    .lBaseG = lG
                    .lBaseB = lB
                End With
            Catch ex As Exception

            End Try
        End Sub
    End Class

    Public MustInherit Class FXEmitter
        Public mvecEmitter As Vector3

        Protected Shared moInvisColor As System.Drawing.Color = System.Drawing.Color.FromArgb(0, 0, 0, 0)

        Public moParticles() As Particle
        Protected moParentEngine As ParticleEngine
        Protected mlPrevFrame As Int32
        Protected mbEmitterStopping As Boolean = False

        Public moPoints() As CustomVertex.PositionColoredTextured
        Public mlParticleUB As Int32
        Public EmitterStopped As Boolean = False
        'Public AttachedObject As RenderObject

        Protected MustOverride Sub ResetParticle(ByVal lIndex As Int32)
        Public MustOverride Sub Update()

        Public Sub MoveEmitter(ByVal fVecX As Single, ByVal fVecY As Single, ByVal fVecZ As Single)
            mvecEmitter.Add(New Vector3(fVecX, fVecY, fVecZ))
        End Sub

        Public Function GetAbsolutePosition() As Vector3
            'If AttachedObject Is Nothing Then
            Return mvecEmitter
            'Else : Return New Vector3(mvecEmitter.X + AttachedObject.LocX, mvecEmitter.Y + AttachedObject.LocY, mvecEmitter.Z + AttachedObject.LocZ)
            'End If
        End Function

        Public ReadOnly Property ParticleCount() As Int32
            Get
                Return mlParticleUB + 1
            End Get
        End Property

        Public Sub New(ByVal oVecLoc As Vector3, ByRef oParentEngine As ParticleEngine, ByVal lParticleCnt As Int32)
            Dim X As Int32

            moParentEngine = oParentEngine
            mvecEmitter = oVecLoc
            mlParticleUB = lParticleCnt - 1
            ReDim moParticles(mlParticleUB)
            ReDim moPoints(mlParticleUB)

            For X = 0 To mlParticleUB
                moParticles(X) = New Particle()
                ResetParticle(X)
            Next X
        End Sub

        Public Sub StopEmitter()
            mbEmitterStopping = True
        End Sub

        Public Sub StartEmitter()
            Dim X As Int32
            mbEmitterStopping = False
            For X = 0 To mlParticleUB
                moParticles(X).ParticleActive = True
                ResetParticle(X)
            Next X
        End Sub
    End Class

    Public Class EngineFX
        Inherits FXEmitter

        Private mfAlphaChgMult As Single

        Public lBaseR As Int32 = 64
        Public lBaseG As Int32 = 192
        Public lBaseB As Int32 = 255

        Public Sub New(ByVal oVecLoc As Vector3, ByVal oParentEngine As ParticleEngine, ByVal lParticleCnt As Int32)
            MyBase.New(oVecLoc, oParentEngine, lParticleCnt)
        End Sub

        Protected Overrides Sub ResetParticle(ByVal lIndex As Int32)
            Dim fX As Single
            Dim fY As Single
            Dim fZ As Single
            Dim fXS As Single
            Dim fYS As Single
            Dim fZS As Single
            Dim fXA As Single
            Dim fYA As Single
            Dim fZA As Single
            Dim fOffsetX As Single
            Dim fOffsetY As Single
            Dim fTemp As Single

            Dim lX As Int32

            If mbEmitterStopping = True Then
                MyBase.moParticles(lIndex).mfA = 0
                MyBase.moParticles(lIndex).ParticleActive = False
                MyBase.moPoints(lIndex).Color = System.Drawing.Color.FromArgb(0, 0, 0, 0).ToArgb

                MyBase.EmitterStopped = True
                For lX = 0 To MyBase.mlParticleUB
                    If MyBase.moParticles(lX).ParticleActive = True Then
                        MyBase.EmitterStopped = False
                        Exit For
                    End If
                Next lX
            Else
                'Speed
                fXS = Rnd() * 1 - 0.1F
                fYS = Rnd() * 1 - 0.1F
                fZS = 10.0F
                'Accel
                If fXS < 0 Then fXA = Rnd() * -0.01F Else fXA = Rnd() * 0.01F
                If fYS < 0 Then fYA = Rnd() * -0.01F Else fYA = Rnd() * 0.01F
                fZA = 0

                'Now, offset our starting X and Y based on the particle count... at 50, we use the X and Y
                If MyBase.ParticleCount > 50 Then
                    fTemp = MyBase.ParticleCount / 10.0F
                    fOffsetX = ((fTemp * 2) * Rnd()) - fTemp
                    fOffsetY = ((fTemp * 2) * Rnd()) - fTemp
                End If

                Dim vecEmitter As Vector3 = MyBase.GetAbsolutePosition()

                fX = vecEmitter.X + fOffsetX
                fY = vecEmitter.Y + fOffsetY
                fZ = vecEmitter.Z

                'Now, rotate everyone
                RotateVals(fX, fZ, fXS, fZS, fXA, fZA)

                moParticles(lIndex).Reset(fX, fY, fZ, fXS, fYS, fZS, fXA, fYA, fZA, lBaseR, lBaseG, lBaseB, 255)
                If mfAlphaChgMult = 0 Then
                    If MyBase.ParticleCount < 50 Then
                        mfAlphaChgMult = 0.1F * CSng(100 / (MyBase.ParticleCount * 2))
                    Else
                        mfAlphaChgMult = 0.1
                    End If
                End If
                moParticles(lIndex).fAChg = -((mfAlphaChgMult + (mfAlphaChgMult * Rnd())) * 255)
            End If


        End Sub

        Public Overrides Sub Update()
            Dim fElapsed As Single
            Dim X As Int32
            Dim vecCenter As Vector3 = New Vector3(16, 16, 0)
            
            If MyBase.EmitterStopped = True Then Exit Sub

            If mlPrevFrame = 0 Then fElapsed = 1 Else fElapsed = (timeGetTime - mlPrevFrame) / 30.0F
            mlPrevFrame = timeGetTime

            'Update the particles
            For X = 0 To mlParticleUB
                With moParticles(X)
                    .Update(fElapsed)
                    If .mfA <= 0 Then
                        ResetParticle(X)
                    End If

                    'MyBase.moPoints(X).Color = FXEmitter.moInvisColor.ToArgb
                    'If MyBase.AttachedObject Is Nothing = False Then
                    'If MyBase.AttachedObject.yVisibility = eVisibilityType.Visible Then
                    MyBase.moPoints(X).Color = .ParticleColor.ToArgb
                    'End If
                    'End If
                    moPoints(X).Position = .vecLoc
                End With
            Next X

        End Sub

        Private Sub RotateVals(ByRef fXLoc As Single, ByRef fZLoc As Single, ByRef fXSpeed As Single, ByRef fZSpeed As Single, ByRef fXAcc As Single, ByRef fZAcc As Single)
            
            Exit Sub

            'Const gdRadPerDegree As Single = Math.PI / 180.0#

            'Dim fRads As Single = (AttachedObject.LocAngle / 10.0F) * gdRadPerDegree
            'Dim fCosR As Single = Math.Cos(fRads)
            'Dim fSinR As Single = Math.Sin(fRads)

            ''Loc
            'fDX = fXLoc
            'fDZ = fZLoc
            'fXLoc = +((fDX * fCosR) + (fDZ * fSinR))
            'fZLoc = -((fDX * fSinR) - (fDZ * fCosR))

            ''Speed
            'fDX = fXSpeed
            'fDZ = fZSpeed
            'fXSpeed = +((fDX * fCosR) + (fDZ * fSinR))
            'fZSpeed = -((fDX * fSinR) - (fDZ * fCosR))

            ''Acc
            'fDX = fXAcc
            'fDZ = fZAcc
            'fXAcc = +((fDX * fCosR) + (fDZ * fSinR))
            'fZAcc = -((fDX * fSinR) - (fDZ * fCosR))
        End Sub

    End Class

    Public Class SmokeyFireFX
        Inherits FXEmitter

        Public Sub New(ByVal oVecLoc As Vector3, ByVal oParentEngine As ParticleEngine, ByVal lParticleCnt As Int32)
            MyBase.New(oVecLoc, oParentEngine, lParticleCnt)
        End Sub

        Protected Overrides Sub ResetParticle(ByVal lIndex As Int32)
            Dim lX As Int32

            If mbEmitterStopping = True Then
                MyBase.moParticles(lIndex).mfA = 0
                MyBase.moParticles(lIndex).ParticleActive = False
                MyBase.moPoints(lIndex).Color = System.Drawing.Color.FromArgb(0, 0, 0, 0).ToArgb

                MyBase.EmitterStopped = True
                For lX = 0 To MyBase.mlParticleUB
                    If MyBase.moParticles(lX).ParticleActive = True Then
                        MyBase.EmitterStopped = False
                        Exit For
                    End If
                Next lX
            Else
                Dim vecEmitter As Vector3 = MyBase.GetAbsolutePosition()

                'moParticles(lIndex).Reset(vecEmitter.X + ((Rnd() * 10) - 5), vecEmitter.Y, vecEmitter.Z + ((Rnd() * 10) - 5), 0, Rnd(), 0, 0, -0.001, 0, 255, 128, Rnd() * 50, (0.8F + (0.2 * Rnd())) * 255)
                moParticles(lIndex).Reset(vecEmitter.X + ((Rnd() * 10) - 5), vecEmitter.Y, vecEmitter.Z + ((Rnd() * 10) - 5), Rnd() * 0.5F - 0.25F, Rnd() + 1, Rnd() * 0.5F - 0.25F, 0, -0.001, 0, 255, 128, Rnd() * 50, CSng(0.8F + (0.2 * Rnd())) * 255)
                moParticles(lIndex).fAChg = (0.01F - (Rnd() * 0.005F)) * -255.0F
                moParticles(lIndex).fRChg = (0.01F + (Rnd() * 0.01F)) * -255.0F
                moParticles(lIndex).fGChg = moParticles(lIndex).fRChg / 2
                moParticles(lIndex).fBChg = Math.Abs((64 - moParticles(lIndex).mfB) / ((moParticles(lIndex).mfR - 64) / moParticles(lIndex).fRChg))
            End If


        End Sub

        Public Overrides Sub Update()
            Dim fElapsed As Single
            Dim X As Int32
            Dim vecCenter As Vector3 = New Vector3(16, 16, 0)
           
            If MyBase.EmitterStopped = True Then Exit Sub

            If mlPrevFrame = 0 Then fElapsed = 1 Else fElapsed = (timeGetTime - mlPrevFrame) / 30.0F
            mlPrevFrame = timeGetTime

            'Update the particles
            For X = 0 To mlParticleUB
                With moParticles(X)
                    .Update(fElapsed)
                    If .mfA <= 0 Then
                        ResetParticle(X)
                    End If

                    'check our limits
                    If .mfR < 64 Then .mfR = 64
                    If .mfG < 64 Then .mfG = 64
                    If .mfB > 64 Then .mfB = 64

                    'MyBase.moPoints(X).Color = FXEmitter.moInvisColor.ToArgb
                    'If MyBase.AttachedObject Is Nothing = False Then
                    'If MyBase.AttachedObject.yVisibility = eVisibilityType.Visible Then
                    MyBase.moPoints(X).Color = .ParticleColor.ToArgb
                    '    End If
                    'End If
                    MyBase.moPoints(X).Position = .vecLoc
                End With
            Next X

        End Sub

    End Class

    Public Class FireFX
        Inherits FXEmitter

        Private mfXLocMin As Single
        Private mfYLocMin As Single
        Private mfZLocMin As Single

        Private mfXLocMax As Single
        Private mfYLocMax As Single
        Private mfZLocMax As Single

        Private mfXSpeedMin As Single
        Private mfYSpeedMin As Single
        Private mfZSpeedMin As Single

        Private mfXSpeedMax As Single
        Private mfYSpeedMax As Single
        Private mfZSpeedMax As Single

        Private mfXAccMin As Single
        Private mfXAccMax As Single

        Private mfYAccMin As Single
        Private mfYAccMax As Single

        Private mfZAccMin As Single
        Private mfZAccMax As Single

        Private mlEmitterMod As ParticleEngine.EmitterType
        Public Property lEmitterMod() As ParticleEngine.EmitterType
            Get
                Return mlEmitterMod
            End Get
            Set(ByVal value As ParticleEngine.EmitterType)
                mlEmitterMod = value

                Dim lRndRng As Int32 = mlParticleUB \ 3
                Dim lHlfRndRng As Int32 = lRndRng \ 2

                Select Case mlEmitterMod
                    Case ParticleEngine.EmitterType.eFireEmitterMod_MinusX  'vertical (y-z) emitter going negative X
                        mfXLocMax = 0 : mfXLocMin = 0
                        mfYLocMax = lRndRng : mfYLocMin = lHlfRndRng
                        mfZLocMax = lRndRng : mfZLocMin = lHlfRndRng
                        mfXSpeedMax = -1.0F : mfXSpeedMin = 0.0F
                        mfYSpeedMax = 0.8F : mfYSpeedMin = 0.4F
                        mfZSpeedMax = 0.8F : mfZSpeedMin = 0.4F
                        mfXAccMax = -0.03F : mfXAccMin = 0.0F
                        mfYAccMax = 0 : mfYAccMin = 0
                        mfZAccMax = 0 : mfZAccMin = 0
                    Case ParticleEngine.EmitterType.eFireEmitterMod_MinusY  'horizontal (x-z) emitter going negative Y
                        mfXLocMax = lRndRng : mfXLocMin = lHlfRndRng
                        mfYLocMax = 0 : mfYLocMin = 0
                        mfZLocMax = lRndRng : mfZLocMin = lHlfRndRng
                        mfXSpeedMax = 0.8F : mfXSpeedMin = 0.4F
                        mfYSpeedMax = -1.0F : mfYSpeedMin = 0.0F
                        mfZSpeedMax = 0.8F : mfZSpeedMin = 0.4F
                        mfXAccMax = 0 : mfXAccMin = 0
                        mfYAccMax = -0.03F : mfYAccMin = 0.0F
                        mfZAccMax = 0 : mfZAccMin = 0
                    Case ParticleEngine.EmitterType.eFireEmitterMod_MinusZ  'vertical (x-y) emitter going negative Z
                        mfXLocMax = lRndRng : mfXLocMin = lHlfRndRng
                        mfYLocMax = lRndRng : mfYLocMin = lHlfRndRng
                        mfZLocMax = 0 : mfZLocMin = 0
                        mfXSpeedMax = 0.8F : mfXSpeedMin = 0.4F
                        mfYSpeedMax = 0.8F : mfYSpeedMin = 0.4F
                        mfZSpeedMax = -1.0F : mfZSpeedMin = 0.0F
                        mfXAccMax = 0 : mfXAccMin = 0
                        mfYAccMax = 0 : mfYAccMin = 0
                        mfZAccMax = -0.03F : mfZAccMin = 0
                    Case ParticleEngine.EmitterType.eFireEmitterMod_PlusX   'vertical (y-z) emitter going positive X
                        mfXLocMax = 0 : mfXLocMin = 0
                        mfYLocMax = lRndRng : mfYLocMin = lHlfRndRng
                        mfZLocMax = lRndRng : mfZLocMin = lHlfRndRng
                        mfXSpeedMax = 1.0F : mfXSpeedMin = 0.0F
                        mfYSpeedMax = 0.8F : mfYSpeedMin = 0.4F
                        mfZSpeedMax = 0.8F : mfZSpeedMin = 0.4F
                        mfXAccMax = 0.03F : mfXAccMin = 0.0F
                        mfYAccMax = 0 : mfYAccMin = 0
                        mfZAccMax = 0 : mfZAccMin = 0
                    Case ParticleEngine.EmitterType.eFireEmitterMod_PlusZ   'vertical (x-y) emitter going positive Z
                        mfXLocMax = lRndRng : mfXLocMin = lHlfRndRng
                        mfYLocMax = lRndRng : mfYLocMin = lHlfRndRng
                        mfZLocMax = 0 : mfZLocMin = 0
                        mfXSpeedMax = 0.8F : mfXSpeedMin = 0.4F
                        mfYSpeedMax = 0.8F : mfYSpeedMin = 0.4F
                        mfZSpeedMax = 1.0F : mfZSpeedMin = 0.0F
                        mfXAccMax = 0 : mfXAccMin = 0
                        mfYAccMax = 0 : mfYAccMin = 0
                        mfZAccMax = 0.03F : mfZAccMin = 0
                    Case Else                                               'horizontal (x-z) emitter going positive Y
                        mfXLocMax = lRndRng : mfXLocMin = lHlfRndRng
                        mfYLocMax = 0 : mfYLocMin = 0
                        mfZLocMax = lRndRng : mfZLocMin = lHlfRndRng
                        mfXSpeedMax = 0.8F : mfXSpeedMin = 0.4F
                        mfYSpeedMax = 1.0F : mfYSpeedMin = 0.0F
                        mfZSpeedMax = 0.8F : mfZSpeedMin = 0.4F
                        mfXAccMax = 0 : mfXAccMin = 0
                        mfYAccMax = 0.03F : mfYAccMin = 0.0F
                        mfZAccMax = 0 : mfZAccMin = 0
                End Select
            End Set
        End Property

        Public Sub New(ByVal oVecLoc As Vector3, ByVal oParentEngine As ParticleEngine, ByVal lParticleCnt As Int32)
            MyBase.New(oVecLoc, oParentEngine, lParticleCnt)
        End Sub

        Protected Overrides Sub ResetParticle(ByVal lIndex As Int32)
            Dim lX As Int32

            If mbEmitterStopping = True Then
                MyBase.moParticles(lIndex).mfA = 0
                MyBase.moParticles(lIndex).ParticleActive = False
                MyBase.moPoints(lIndex).Color = System.Drawing.Color.FromArgb(0, 0, 0, 0).ToArgb

                MyBase.EmitterStopped = True
                For lX = 0 To MyBase.mlParticleUB
                    If MyBase.moParticles(lX).ParticleActive = True Then
                        MyBase.EmitterStopped = False
                        Exit For
                    End If
                Next lX
            Else
                Dim vecEmitter As Vector3 = MyBase.GetAbsolutePosition()

                Dim fX As Single = (Rnd() * mfXLocMax) - mfXLocMin
                Dim fY As Single = (Rnd() * mfYLocMax) - mfYLocMin
                Dim fZ As Single = (Rnd() * mfZLocMax) - mfZLocMin
                Dim fXS As Single = (Rnd() * mfXSpeedMax) - mfXSpeedMin
                Dim fYS As Single = (Rnd() * mfYSpeedMax) - mfYSpeedMin
                Dim fZS As Single = (Rnd() * mfZSpeedMax) - mfZSpeedMin
                Dim fXA As Single = (Rnd() * mfXAccMax) - mfXAccMin
                Dim fYA As Single = (Rnd() * mfYAccMax) - mfYAccMin
                Dim fZA As Single = (Rnd() * mfZAccMax) - mfZAccMin

                RotateVals(fX, fY, fZ, fXS, fZS, fXA, fZA)

                moParticles(lIndex).Reset(vecEmitter.X + fX, vecEmitter.Y + fY, vecEmitter.Z + fZ, fXS, fYS, fZS, fXA, fYA, fZA, 255, 128, 50, (0.6F + (0.2F * Rnd())) * 255)
                moParticles(lIndex).fAChg = -((0.01F + Rnd() * 0.05F) * 255)
            End If

        End Sub

        Public Overrides Sub Update()
            Dim fElapsed As Single
            Dim X As Int32
            Dim vecCenter As Vector3 = New Vector3(16, 16, 0)
            'Dim matWorld As Matrix

            If MyBase.EmitterStopped = True Then Exit Sub

            If mlPrevFrame = 0 Then fElapsed = 1 Else fElapsed = (timeGetTime - mlPrevFrame) / 30.0F
            mlPrevFrame = timeGetTime

            'Update the particles
            For X = 0 To mlParticleUB
                With moParticles(X)
                    .Update(fElapsed)
                    If .mfA <= 0 Then
                        ResetParticle(X)
                    End If
                    'MyBase.moPoints(X).Color = FXEmitter.moInvisColor.ToArgb
                    'If MyBase.AttachedObject Is Nothing = False Then
                    '    If MyBase.AttachedObject.yVisibility = eVisibilityType.Visible Then
                    MyBase.moPoints(X).Color = .ParticleColor.ToArgb
                    '    End If
                    'End If
                    MyBase.moPoints(X).Position = .vecLoc
                End With
            Next X

        End Sub

        Private Sub RotateVals(ByRef fXLoc As Single, ByRef fYLoc As Single, ByRef fZLoc As Single, ByRef fXSpeed As Single, ByRef fZSpeed As Single, ByRef fXAcc As Single, ByRef fZAcc As Single)
            'Dim fDX As Single
            'Dim fDZ As Single

            ''If MyBase.AttachedObject Is Nothing Then Return
            ''Dim fRads As Single = -((AttachedObject.LocYaw / 10.0F)) * gdRadPerDegree
            'Dim fRads As Single = -((giLocYaw / 10.0F)) * gdRadPerDegree

            'Dim fCosR As Single = CSng(Math.Cos(fRads))
            'Dim fSinR As Single = CSng(Math.Sin(fRads))

            ''Yaw...
            'fDX = fXLoc
            'fDZ = fYLoc
            'fXLoc = +((fDX * fCosR) + (fDZ * fSinR))
            'fYLoc = -((fDX * fSinR) - (fDZ * fCosR))

            ''Now set up for standard rotation...
            ''fRads = ((AttachedObject.LocAngle / 10.0F) - 90) * gdRadPerDegree
            'fRads = ((gilocangle / 10.0F) - 90) * gdRadPerDegree
            'fCosR = CSng(Math.Cos(fRads))
            'fSinR = CSng(Math.Sin(fRads))

            ''Loc
            'fDX = fXLoc
            'fDZ = fZLoc
            'fXLoc = +((fDX * fCosR) + (fDZ * fSinR))
            'fZLoc = -((fDX * fSinR) - (fDZ * fCosR))

            ''Speed
            'fDX = fXSpeed
            'fDZ = fZSpeed
            'fXSpeed = +((fDX * fCosR) + (fDZ * fSinR))
            'fZSpeed = -((fDX * fSinR) - (fDZ * fCosR))

            ''Acc
            'fDX = fXAcc
            'fDZ = fZAcc
            'fXAcc = +((fDX * fCosR) + (fDZ * fSinR))
            'fZAcc = -((fDX * fSinR) - (fDZ * fCosR))
        End Sub

    End Class

End Namespace

Public Class BurnFXData
    Public lParticleFXID As Int32 = -1
    Public lSide As Int32
    Public lEmitterType As ParticleFX.ParticleEngine.EmitterType

    Public LocX As Int32
    Public LocY As Int32
    Public LocZ As Int32

    Public lPCnt As Int32

    Private msDisplayVal As String

    Public Property sDisplay() As String
        Get
            Return msDisplayVal
        End Get
        Set(ByVal value As String)
            msDisplayVal = value
        End Set
    End Property
End Class