Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

Public Class frmMain
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
    Friend WithEvents mnuFile As System.Windows.Forms.MenuItem
    Friend WithEvents picRender As System.Windows.Forms.PictureBox
    Friend WithEvents mnuNewForm As System.Windows.Forms.MenuItem
    Friend WithEvents mnuOpenForm As System.Windows.Forms.MenuItem
    Friend WithEvents mnuSaveForm As System.Windows.Forms.MenuItem
    Friend WithEvents mnuExit As System.Windows.Forms.MenuItem
    Friend WithEvents mnuInsert As System.Windows.Forms.MenuItem
    Friend WithEvents mnuInterface As System.Windows.Forms.MenuItem
    Friend WithEvents mnuClearAll As System.Windows.Forms.MenuItem
    Friend WithEvents mnuCompile As System.Windows.Forms.MenuItem
    Friend WithEvents tmrRender As System.Windows.Forms.Timer
    Friend WithEvents mnuSep1 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuSep2 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuNewFullScreen As System.Windows.Forms.MenuItem
    Friend WithEvents mnuNewNormal As System.Windows.Forms.MenuItem
    Friend WithEvents mnuTestMode As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuUndo As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem4 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuCopy As System.Windows.Forms.MenuItem
    Friend WithEvents mnuPaste As System.Windows.Forms.MenuItem
    Friend WithEvents mnuDelete As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem2 As System.Windows.Forms.MenuItem
    Friend WithEvents dlgOpen As System.Windows.Forms.OpenFileDialog
    Friend WithEvents dlgSave As System.Windows.Forms.SaveFileDialog
    Friend WithEvents MenuItem3 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem5 As System.Windows.Forms.MenuItem
    Friend WithEvents cboControls As System.Windows.Forms.ComboBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Me.mnuMain = New System.Windows.Forms.MainMenu(Me.components)
        Me.mnuFile = New System.Windows.Forms.MenuItem
        Me.mnuNewForm = New System.Windows.Forms.MenuItem
        Me.mnuNewFullScreen = New System.Windows.Forms.MenuItem
        Me.mnuNewNormal = New System.Windows.Forms.MenuItem
        Me.mnuOpenForm = New System.Windows.Forms.MenuItem
        Me.mnuSaveForm = New System.Windows.Forms.MenuItem
        Me.mnuSep1 = New System.Windows.Forms.MenuItem
        Me.mnuExit = New System.Windows.Forms.MenuItem
        Me.MenuItem1 = New System.Windows.Forms.MenuItem
        Me.mnuUndo = New System.Windows.Forms.MenuItem
        Me.MenuItem4 = New System.Windows.Forms.MenuItem
        Me.mnuCopy = New System.Windows.Forms.MenuItem
        Me.mnuPaste = New System.Windows.Forms.MenuItem
        Me.mnuDelete = New System.Windows.Forms.MenuItem
        Me.MenuItem2 = New System.Windows.Forms.MenuItem
        Me.mnuInsert = New System.Windows.Forms.MenuItem
        Me.mnuInterface = New System.Windows.Forms.MenuItem
        Me.mnuClearAll = New System.Windows.Forms.MenuItem
        Me.mnuSep2 = New System.Windows.Forms.MenuItem
        Me.mnuCompile = New System.Windows.Forms.MenuItem
        Me.mnuTestMode = New System.Windows.Forms.MenuItem
        Me.MenuItem3 = New System.Windows.Forms.MenuItem
        Me.picRender = New System.Windows.Forms.PictureBox
        Me.tmrRender = New System.Windows.Forms.Timer(Me.components)
        Me.dlgOpen = New System.Windows.Forms.OpenFileDialog
        Me.dlgSave = New System.Windows.Forms.SaveFileDialog
        Me.cboControls = New System.Windows.Forms.ComboBox
        Me.MenuItem5 = New System.Windows.Forms.MenuItem
        CType(Me.picRender, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'mnuMain
        '
        Me.mnuMain.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuFile, Me.MenuItem1, Me.mnuInsert, Me.mnuInterface, Me.MenuItem3, Me.MenuItem5})
        '
        'mnuFile
        '
        Me.mnuFile.Index = 0
        Me.mnuFile.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuNewForm, Me.mnuOpenForm, Me.mnuSaveForm, Me.mnuSep1, Me.mnuExit})
        Me.mnuFile.Text = "File"
        '
        'mnuNewForm
        '
        Me.mnuNewForm.Index = 0
        Me.mnuNewForm.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuNewFullScreen, Me.mnuNewNormal})
        Me.mnuNewForm.Text = "New Window"
        '
        'mnuNewFullScreen
        '
        Me.mnuNewFullScreen.Index = 0
        Me.mnuNewFullScreen.Shortcut = System.Windows.Forms.Shortcut.CtrlF
        Me.mnuNewFullScreen.Text = "Full Screen..."
        '
        'mnuNewNormal
        '
        Me.mnuNewNormal.Index = 1
        Me.mnuNewNormal.Shortcut = System.Windows.Forms.Shortcut.CtrlN
        Me.mnuNewNormal.Text = "Normal..."
        '
        'mnuOpenForm
        '
        Me.mnuOpenForm.Index = 1
        Me.mnuOpenForm.Shortcut = System.Windows.Forms.Shortcut.CtrlO
        Me.mnuOpenForm.Text = "Open..."
        '
        'mnuSaveForm
        '
        Me.mnuSaveForm.Index = 2
        Me.mnuSaveForm.Shortcut = System.Windows.Forms.Shortcut.CtrlS
        Me.mnuSaveForm.Text = "Save..."
        '
        'mnuSep1
        '
        Me.mnuSep1.Index = 3
        Me.mnuSep1.Text = "-"
        '
        'mnuExit
        '
        Me.mnuExit.Index = 4
        Me.mnuExit.Shortcut = System.Windows.Forms.Shortcut.AltF4
        Me.mnuExit.Text = "Exit"
        '
        'MenuItem1
        '
        Me.MenuItem1.Index = 1
        Me.MenuItem1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuUndo, Me.MenuItem4, Me.mnuCopy, Me.mnuPaste, Me.mnuDelete, Me.MenuItem2})
        Me.MenuItem1.Text = "Edit"
        '
        'mnuUndo
        '
        Me.mnuUndo.Index = 0
        Me.mnuUndo.Shortcut = System.Windows.Forms.Shortcut.CtrlZ
        Me.mnuUndo.Text = "Undo"
        '
        'MenuItem4
        '
        Me.MenuItem4.Index = 1
        Me.MenuItem4.Text = "-"
        '
        'mnuCopy
        '
        Me.mnuCopy.Index = 2
        Me.mnuCopy.Shortcut = System.Windows.Forms.Shortcut.CtrlC
        Me.mnuCopy.Text = "Copy"
        '
        'mnuPaste
        '
        Me.mnuPaste.Index = 3
        Me.mnuPaste.Shortcut = System.Windows.Forms.Shortcut.CtrlV
        Me.mnuPaste.Text = "Paste"
        '
        'mnuDelete
        '
        Me.mnuDelete.Index = 4
        Me.mnuDelete.Shortcut = System.Windows.Forms.Shortcut.Del
        Me.mnuDelete.Text = "Delete"
        '
        'MenuItem2
        '
        Me.MenuItem2.Index = 5
        Me.MenuItem2.Text = "Transparency Hack"
        '
        'mnuInsert
        '
        Me.mnuInsert.Index = 2
        Me.mnuInsert.Text = "Insert"
        '
        'mnuInterface
        '
        Me.mnuInterface.Index = 3
        Me.mnuInterface.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuClearAll, Me.mnuSep2, Me.mnuCompile, Me.mnuTestMode})
        Me.mnuInterface.Text = "Interface"
        '
        'mnuClearAll
        '
        Me.mnuClearAll.Index = 0
        Me.mnuClearAll.Text = "Clear All..."
        '
        'mnuSep2
        '
        Me.mnuSep2.Index = 1
        Me.mnuSep2.Text = "-"
        '
        'mnuCompile
        '
        Me.mnuCompile.Index = 2
        Me.mnuCompile.Text = "Create Code Text..."
        '
        'mnuTestMode
        '
        Me.mnuTestMode.Index = 3
        Me.mnuTestMode.Text = "Enter Test Mode..."
        '
        'MenuItem3
        '
        Me.MenuItem3.Index = 4
        Me.MenuItem3.Text = "HullSlots"
        '
        'picRender
        '
        Me.picRender.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.picRender.Location = New System.Drawing.Point(8, 8)
        Me.picRender.Name = "picRender"
        Me.picRender.Size = New System.Drawing.Size(1024, 768)
        Me.picRender.TabIndex = 0
        Me.picRender.TabStop = False
        '
        'tmrRender
        '
        Me.tmrRender.Interval = 20
        '
        'dlgSave
        '
        Me.dlgSave.FileName = "doc1"
        '
        'cboControls
        '
        Me.cboControls.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboControls.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboControls.Location = New System.Drawing.Point(808, 784)
        Me.cboControls.Name = "cboControls"
        Me.cboControls.Size = New System.Drawing.Size(224, 21)
        Me.cboControls.TabIndex = 1
        '
        'MenuItem5
        '
        Me.MenuItem5.Index = 5
        Me.MenuItem5.Text = "LabelScroll"
        '
        'frmMain
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(1040, 809)
        Me.Controls.Add(Me.cboControls)
        Me.Controls.Add(Me.picRender)
        Me.Menu = Me.mnuMain
        Me.Name = "frmMain"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Interface Builder"
        CType(Me.picRender, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private moDevice As Device

    Private moUILib As UILib

    Private moCurrentWindow As UIWindow

    Public moSelectedControl As UIControl
    Private mlSelectedControlIdx As Int32 = -1
    Private mbMouseDown As Boolean = False

    Private mbResizeLeft As Boolean
    Private mbResizeRight As Boolean
    Private mbResizeTop As Boolean
    Private mbResizeBottom As Boolean

    Private WithEvents mfrmProps As frmProps

    Private mfrmCompiled As frmCompiled

    Private mbTestMode As Boolean

    Private Enum eLastActionType As Integer
        eNoAction_LAT = 0
        eChangeLocation_LAT
        eResize_LAT
        eVisible_LAT
        eEnabled_LAT
        eDelete_LAT
    End Enum
    Private ml_LastActionType(9) As eLastActionType
    Private ml_LastActionVal1(9) As Int32
    Private ml_LastActionVal2(9) As Int32
    Private mo_LastActionObj(9) As Object
    Private Const ml_UNDOABLE_ACTION_QUEUE_SIZE As Int32 = 10

    Private mlControlX As Int32
    Private mlControlY As Int32
    Private mbControlMoved As Boolean = False

    Private moCopyObj As Object

    Private Sub AddUndoableAction(ByVal lActionType As eLastActionType, ByVal lVal1 As Int32, ByVal lVal2 As Int32, ByVal oObj As Object)
        Dim X As Int32

        'shift all values
        For X = ml_UNDOABLE_ACTION_QUEUE_SIZE - 1 To 1 Step -1
            ml_LastActionType(X) = ml_LastActionType(X - 1)
            ml_LastActionVal1(X) = ml_LastActionVal1(X - 1)
            ml_LastActionVal2(X) = ml_LastActionVal2(X - 1)
            mo_LastActionObj(X) = mo_LastActionObj(X - 1)
        Next X

        'now, store this one
        ml_LastActionType(0) = lActionType
        ml_LastActionVal1(0) = lVal1
        ml_LastActionVal2(0) = lVal2
        mo_LastActionObj(0) = oObj
    End Sub

    Public Function InitD3D(ByVal picBox As PictureBox) As Boolean
        Dim uParms As PresentParameters
        Dim uDispMode As DisplayMode
        Dim bRes As Boolean

        Try
            uDispMode = Manager.Adapters.Default.CurrentDisplayMode
            uParms = New PresentParameters()
            With uParms
                .Windowed = True
                .SwapEffect = SwapEffect.Discard
                .BackBufferCount = 1

                .BackBufferFormat = uDispMode.Format
                .BackBufferHeight = picBox.Height
                .BackBufferWidth = picBox.Width

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

        Return bRes
    End Function

    Private Sub frmMain_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If InitD3D(picRender) = False Then
            End
        Else
            moUILib = New UILib(moDevice)
            LoadInsertObjects()
            tmrRender.Enabled = True
        End If
    End Sub

    Private Sub LoadInsertObjects()
        mnuInsert.MenuItems.Add("Label", AddressOf mnuInsertObj_Click)
        mnuInsert.MenuItems.Add("Text Box", AddressOf mnuInsertObj_Click)
        mnuInsert.MenuItems.Add("Button", AddressOf mnuInsertObj_Click)
        mnuInsert.MenuItems.Add("Option Button", AddressOf mnuInsertObj_Click)
        mnuInsert.MenuItems.Add("Checkbox Button", AddressOf mnuInsertObj_Click)
        mnuInsert.MenuItems.Add("Horizontal Scroll Bar", AddressOf mnuInsertObj_Click)
        mnuInsert.MenuItems.Add("Vertical Scroll Bar", AddressOf mnuInsertObj_Click)

        mnuInsert.MenuItems.Add("List Box", AddressOf mnuInsertObj_Click)
        mnuInsert.MenuItems.Add("Combo Box", AddressOf mnuInsertObj_Click)

        mnuInsert.MenuItems.Add("Frame", AddressOf mnuInsertObj_Click)

        mnuInsert.MenuItems.Add("Line", AddressOf mnuInsertObj_Click)
    End Sub

    Private Sub CreateNewControl(ByVal sControlTypeName As String, ByVal sControlName As String)
        Select Case sControlTypeName
            Case "Label"
                Dim oLbl As UILabel = New UILabel(moUILib)
                oLbl.ControlName = sControlName
                oLbl.Width = 100
                oLbl.Height = 32
                oLbl.Left = 0
                oLbl.Top = 0
                oLbl.Visible = True
                oLbl.Enabled = True
                oLbl.Caption = "Label Caption"

                moCurrentWindow.AddChild(CType(oLbl, UIControl))
                oLbl = Nothing

            Case "Text Box"
                Dim oTxt As UITextBox = New UITextBox(moUILib)
                With oTxt
                    .ControlName = sControlName
                    .Width = 100
                    .Height = 32
                    .Top = 0
                    .Left = 0
                    .Visible = True
                    .Enabled = True
                    .Caption = "Text Caption"
                End With
                moCurrentWindow.AddChild(CType(oTxt, UIControl))
                oTxt = Nothing
            Case "Button"
                Dim oBtn As UIButton = New UIButton(moUILib)
                With oBtn
                    .ControlName = sControlName
                    .Width = 100
                    .Height = 32
                    .Top = 0
                    .Left = 0
                    .Visible = True
                    .Enabled = True
                    .Caption = "Button Caption"
                End With
                moCurrentWindow.AddChild(CType(oBtn, UIControl))
                oBtn = Nothing
            Case "Option Button"
                Dim oOpt As UIOption = New UIOption(moUILib)
                With oOpt
                    .ControlName = sControlName
                    .Width = 100
                    .Height = 32
                    .Top = 0
                    .Left = 0
                    .Visible = True
                    .Enabled = True
                    .Caption = "Option Caption"
                End With
                moCurrentWindow.AddChild(CType(oOpt, UIControl))
                oOpt = Nothing
            Case "Checkbox Button"
                Dim oChk As UICheckBox = New UICheckBox(moUILib)
                With oChk
                    .ControlName = sControlName
                    .Width = 100
                    .Height = 32
                    .Top = 0
                    .Left = 0
                    .Visible = True
                    .Enabled = True
                    .Caption = "Check Caption"
                End With
                moCurrentWindow.AddChild(CType(oChk, UIControl))
                oChk = Nothing
            Case "Horizontal Scroll Bar"
                Dim oHScr As UIScrollBar = New UIScrollBar(moUILib, False)
                With oHScr
                    .ControlName = sControlName
                    .Left = 0
                    .Top = 0
                    .Width = 100
                    .Height = 32
                    .Value = 0
                    .MaxValue = 100
                    .Visible = True
                    .Enabled = True
                End With
                moCurrentWindow.AddChild(CType(oHScr, UIControl))
                oHScr = Nothing
            Case "Vertical Scroll Bar"
                Dim oVScr As UIScrollBar = New UIScrollBar(moUILib, True)
                With oVScr
                    .ControlName = sControlName
                    .Left = 0
                    .Top = 0
                    .Width = 32
                    .Height = 100
                    .Value = 0
                    .MaxValue = 100
                    .Visible = True
                    .Enabled = True
                End With
                moCurrentWindow.AddChild(CType(oVScr, UIControl))
                oVScr = Nothing
            Case "List Box"
                Dim oLst As UIListBox = New UIListBox(moUILib)
                With oLst
                    .ControlName = sControlName
                    .Left = 0
                    .Top = 0
                    .Width = 100
                    .Height = 100
                    .Visible = True
                    .Enabled = True
                End With
                moCurrentWindow.AddChild(CType(oLst, UIControl))
                oLst = Nothing
            Case "Combo Box"
                Dim oCbo As UIComboBox = New UIComboBox(moUILib)
                With oCbo
                    .ControlName = sControlName
                    .Left = 0
                    .Top = 0
                    .Width = 100
                    .Height = 24
                    .Visible = True
                    .Enabled = True
                End With

                'TODO: For testing purposes...
                oCbo.AddItem("Item 1")
                oCbo.AddItem("Item 2")
                oCbo.AddItem("Item 3")
                oCbo.AddItem("Item 4")
                oCbo.AddItem("Item 5")

                moCurrentWindow.AddChild(CType(oCbo, UIControl))
                oCbo = Nothing
            Case "Frame"
                Dim oFra As UIWindow = New UIWindow(moUILib)
                With oFra
                    .ControlName = sControlName
                    .Left = 0
                    .Top = 0
                    .Width = 100
                    .Height = 100
                    .Visible = True
                    .Enabled = True
                End With
                moCurrentWindow.AddChild(CType(oFra, UIControl))
                oFra = Nothing
            Case "Line"
                Dim oUILn As UILine = New UILine(moUILib)
                With oUILn
                    .ControlName = sControlName
                    .Left = 0
                    .Top = 0
                    .Width = 100
                    .Height = 100
                    .Visible = True
                    .Enabled = True
                End With
                moCurrentWindow.AddChild(CType(oUILn, UIControl))
                oUILn = Nothing
        End Select

        RefreshComboList()
    End Sub

    Private Sub RefreshComboList()
        If moCurrentWindow Is Nothing = False Then
            'Now, refresh our list...
            cboControls.Items.Clear()
            cboControls.Items.Add(moCurrentWindow.ControlName & " (UIWindow)")

            Dim X As Int32
            Dim sTemp As String

            For X = 0 To moCurrentWindow.ChildrenUB
                If moCurrentWindow.moChildren(X) Is Nothing = False Then
                    sTemp = Replace$(moCurrentWindow.moChildren(X).GetType.ToString, "InterfaceBuilder.", "")
                    cboControls.Items.Add(moCurrentWindow.moChildren(X).ControlName & " (" & sTemp & ")")
                End If
            Next X

            cboControls.Sorted = True

            If moSelectedControl Is Nothing = False Then
                cboControls.FindString(moSelectedControl.ControlName)
            End If


        End If
    End Sub

    Private Sub mnuInsertObj_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Static xlControlCnt As Int32

        If moCurrentWindow Is Nothing Then Exit Sub
        xlControlCnt += 1
        CreateNewControl(CType(sender, MenuItem).Text, "NewControl" & xlControlCnt)

        moSelectedControl = moCurrentWindow.moChildren(moCurrentWindow.ChildrenUB)
        If mfrmProps Is Nothing Then mfrmProps = New frmProps()
        If mfrmProps.Visible = False Then mfrmProps.Show(Me)
        mfrmProps.Left = Me.Left + Me.Width
        mfrmProps.Top = Me.Top
        mfrmProps.ConfigProps(moSelectedControl)
    End Sub

    Private Sub tmrRender_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tmrRender.Tick
        'Dim oWindow As UIWindow

        'Here, we render our scene...
        moDevice.Clear(ClearFlags.Target Or ClearFlags.ZBuffer, System.Drawing.Color.Black, 1.0F, 0)
        moDevice.BeginScene()

        'do our render here...
        'For Each oWindow In moUILib.WindowList
        '    oWindow.DrawControl()
        'Next
        If moCurrentWindow Is Nothing = False Then moCurrentWindow.DrawControl()

        moDevice.EndScene()
        moDevice.Present()
    End Sub

    Private Sub frmMain_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        moCurrentWindow = Nothing
        moUILib.Pen.Dispose()
        moUILib.Pen = Nothing
        moUILib.oDevice = Nothing
        moUILib = Nothing
        moDevice.Dispose()
        moDevice = Nothing
        Application.Exit()
    End Sub

    Private Sub mnuExit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuExit.Click
        Me.Close()
    End Sub

    Private Sub frmMain_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Resize
        If moDevice Is Nothing = True Then Exit Sub

        Dim uParms As PresentParameters
        Dim bWindowed As Boolean = moDevice.PresentationParameters.Windowed
        Dim uDispMode As DisplayMode = moDevice.DisplayMode
        uParms = New PresentParameters()

        With uParms
            bWindowed = True
            .Windowed = bWindowed
            .SwapEffect = SwapEffect.Discard
            .BackBufferCount = 1
            If bWindowed Then
                .BackBufferFormat = uDispMode.Format
                .BackBufferHeight = Me.ClientSize.Height
                .BackBufferWidth = Me.ClientSize.Width
            Else
                'TODO: Eventually, I will want to set this up to change the resolution of the screen
                '  but right now, I want to get a working version, so we just use the current resolution
                .BackBufferFormat = uDispMode.Format
                .BackBufferHeight = uDispMode.Height
                .BackBufferWidth = uDispMode.Width
            End If

            .PresentationInterval = PresentInterval.Immediate
            .AutoDepthStencilFormat = DepthFormat.D16
            .EnableAutoDepthStencil = True
        End With

        moDevice.Reset(uParms)
        uParms = Nothing
        uDispMode = Nothing

        'before we can proceed... need to make sure the UILib has a proper RenderTarget
        moUILib.oRenderTarget = Nothing
        moUILib.oRenderTarget = moDevice.GetRenderTarget(0)


        'TODO: Replace this with intelligent texture loading
        'oInterfaceTexture = oRes.GetTexture("Interface_Main_Texture.bmp")
        Dim sFile As String = AppDomain.CurrentDomain.BaseDirectory
        If sFile.EndsWith("\") = False Then sFile &= "\"
        sFile &= "Textures\Interface_Main_Texture.bmp"
        moUILib.oInterfaceTexture = TextureLoader.FromFile(moDevice, sFile, 0, 0, 0, Usage.None, Format.Unknown, Pool.Default, Filter.None, Filter.None, System.Drawing.Color.FromArgb(255, 255, 0, 255).ToArgb)
    End Sub

#Region " Events from Properties Window "
    Private Sub mfrmProps_UpdateBackColorDisabled(ByVal cldNew As System.Drawing.Color) Handles mfrmProps.UpdateBackColorDisabled
        Dim oObj As Object = moSelectedControl
        CallByName(oObj, "BackColorDisabled", CallType.Let, cldNew)
    End Sub

    Private Sub mfrmProps_UpdateBackColorEnabled(ByVal clrNew As System.Drawing.Color) Handles mfrmProps.UpdateBackColorEnabled
        Dim oObj As Object = moSelectedControl
        CallByName(oObj, "BackColorEnabled", CallType.Let, clrNew)
    End Sub

    Private Sub mfrmProps_UpdateBorderColor(ByVal clrNew As System.Drawing.Color) Handles mfrmProps.UpdateBorderColor
        Dim oObj As Object = moSelectedControl
        CallByName(oObj, "BorderColor", CallType.Let, clrNew)
    End Sub

    Private Sub mfrmProps_UpdateCaption(ByVal sCaption As String) Handles mfrmProps.UpdateCaption
        Dim oObj As Object = moSelectedControl
        CallByName(oObj, "Caption", CallType.Let, sCaption)
    End Sub

    Private Sub mfrmProps_UpdateDrawBackImage(ByVal bNewVal As Boolean) Handles mfrmProps.UpdateDrawBackImage
        Dim oObj As Object = moSelectedControl
        CallByName(oObj, "DrawBackImage", CallType.Let, bNewVal)
    End Sub

    Private Sub mfrmProps_UpdateEnabled(ByVal bEnabled As Boolean) Handles mfrmProps.UpdateEnabled
        AddUndoableAction(eLastActionType.eEnabled_LAT, CInt(moSelectedControl.Enabled), 0, moSelectedControl)
        moSelectedControl.Enabled = bEnabled
    End Sub

    Private Sub mfrmProps_UpdateFillColor(ByVal clrNew As System.Drawing.Color) Handles mfrmProps.UpdateFillColor
        Dim oObj As Object = moSelectedControl
        CallByName(oObj, "FillColor", CallType.Let, clrNew)
    End Sub

    Private Sub mfrmProps_UpdateFont(ByVal oFont As System.Drawing.Font) Handles mfrmProps.UpdateFont
        Dim oObj As Object = moSelectedControl
        CallByName(oObj, "SetFont", CallType.Method, oFont)
    End Sub

    Private Sub mfrmProps_UpdateForeColor(ByVal clrNew As System.Drawing.Color) Handles mfrmProps.UpdateForeColor
        Dim oObj As Object = moSelectedControl
        CallByName(oObj, "ForeColor", CallType.Let, clrNew)
    End Sub

    Private Sub mfrmProps_UpdateHeight(ByVal lHeight As Integer) Handles mfrmProps.UpdateHeight
        If lHeight = 0 AndAlso moSelectedControl.GetType.ToString.ToUpper <> "INTERFACEBUILDER.UILINE" Then Return
        AddUndoableAction(eLastActionType.eResize_LAT, moSelectedControl.Width, moSelectedControl.Height, moSelectedControl)
        moSelectedControl.Height = lHeight
    End Sub

    Private Sub mfrmProps_UpdateLargeChange(ByVal lLargeChange As Integer) Handles mfrmProps.UpdateLargeChange
        Dim oObj As Object = moSelectedControl
        CallByName(oObj, "LargeChange", CallType.Let, lLargeChange)
    End Sub

    Private Sub mfrmProps_UpdateLeft(ByVal lLeft As Integer) Handles mfrmProps.UpdateLeft
        AddUndoableAction(eLastActionType.eChangeLocation_LAT, moSelectedControl.Left, moSelectedControl.Top, moSelectedControl)
        moSelectedControl.Left = lLeft
    End Sub

    Private Sub mfrmProps_UpdateMaxLength(ByVal lMaxLength As Integer) Handles mfrmProps.UpdateMaxLength
        Dim oObj As Object = moSelectedControl
        CallByName(oObj, "MaxLength", CallType.Let, lMaxLength)
    End Sub

    Private Sub mfrmProps_UpdateMaxValue(ByVal lMaxValue As Integer) Handles mfrmProps.UpdateMaxValue
        Dim oObj As Object = moSelectedControl
        CallByName(oObj, "MaxValue", CallType.Let, lMaxValue)
    End Sub

    Private Sub mfrmProps_UpdateMinValue(ByVal lMinValue As Integer) Handles mfrmProps.UpdateMinValue
        Dim oObj As Object = moSelectedControl
        CallByName(oObj, "MinValue", CallType.Let, lMinValue)
    End Sub

    Private Sub mfrmProps_UpdateName(ByVal sName As String) Handles mfrmProps.UpdateName
        moSelectedControl.ControlName = sName
    End Sub

    Private Sub mfrmProps_UpdateReverseDirection(ByVal bNewVal As Boolean) Handles mfrmProps.UpdateReverseDirection
        Dim oObj As Object = moSelectedControl
        CallByName(oObj, "ReverseDirection", CallType.Let, bNewVal)
    End Sub

    Private Sub mfrmProps_UpdateSmallChange(ByVal lSmallChange As Integer) Handles mfrmProps.UpdateSmallChange
        Dim oObj As Object = moSelectedControl
        CallByName(oObj, "SmallChange", CallType.Let, lSmallChange)
    End Sub

    Private Sub mfrmProps_UpdateTop(ByVal lTop As Integer) Handles mfrmProps.UpdateTop
        AddUndoableAction(eLastActionType.eChangeLocation_LAT, moSelectedControl.Left, moSelectedControl.Top, moSelectedControl)
        moSelectedControl.Top = lTop
    End Sub

    Private Sub mfrmProps_UpdateValue(ByVal lValue As Integer) Handles mfrmProps.UpdateValue
        Dim oObj As Object = moSelectedControl
        CallByName(oObj, "Value", CallType.Let, lValue)
    End Sub

    Private Sub mfrmProps_UpdateVisible(ByVal bVisible As Boolean) Handles mfrmProps.UpdateVisible
        AddUndoableAction(eLastActionType.eVisible_LAT, CInt(moSelectedControl.Visible), 0, moSelectedControl)
        moSelectedControl.Visible = bVisible
    End Sub

    Private Sub mfrmProps_UpdateWidth(ByVal lWidth As Integer) Handles mfrmProps.UpdateWidth
        If lWidth = 0 AndAlso moSelectedControl.GetType.ToString.ToUpper <> "INTERFACEBUILDER.UILINE" Then Return
        AddUndoableAction(eLastActionType.eResize_LAT, moSelectedControl.Width, moSelectedControl.Height, moSelectedControl)
        moSelectedControl.Width = lWidth
    End Sub

    Private Sub mfrmProps_UpdateDropDownBorderColor(ByVal clrNew As System.Drawing.Color) Handles mfrmProps.UpdateDropDownBorderColor
        Dim oObj As Object = moSelectedControl
        CallByName(oObj, "DropDownListBorderColor", CallType.Let, clrNew)
    End Sub

    Private Sub mfrmProps_UpdateHighlightColor(ByVal clrNew As System.Drawing.Color) Handles mfrmProps.UpdateHighlightColor
        Dim oObj As Object = moSelectedControl
        CallByName(oObj, "HighlightColor", CallType.Let, clrNew)
    End Sub

    Private Sub mfrmProps_UpdateLocRect(ByVal lX As Integer, ByVal lY As Integer, ByVal lWidth As Int32, ByVal lHeight As Int32) Handles mfrmProps.UpdateLocRect
        If lWidth = 0 OrElse lHeight = 0 Then Return
        moSelectedControl.ControlImageRect.Location = New Point(lX, lY)
        moSelectedControl.ControlImageRect.Width = lWidth
        moSelectedControl.ControlImageRect.Height = lHeight
    End Sub
#End Region

#Region " User Interface Events "
    Private Sub frmMain_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyUp
        If mbTestMode = True Then
            If moCurrentWindow Is Nothing = False Then
                Call moCurrentWindow.PostMessage(UILibMsgCode.eKeyUpCode, e)
            End If
        End If
    End Sub

    Private Sub frmMain_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        If mbTestMode = True Then
            If moCurrentWindow Is Nothing = False Then
                moCurrentWindow.PostMessage(UILibMsgCode.eKeyDownCode, e)
            End If
        End If
    End Sub

    Private Sub frmMain_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        If mbTestMode = True Then
            If moCurrentWindow Is Nothing = False Then
                moCurrentWindow.PostMessage(UILibMsgCode.eKeyPressCode, e)
            End If
        End If
    End Sub

    Private Sub picRender_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles picRender.MouseDown
        mbMouseDown = True

        Dim oTmp As UIControl = Nothing
        Dim X As Int32

        If mbTestMode = True Then
            If moCurrentWindow Is Nothing = False Then
                moCurrentWindow.PostMessage(UILibMsgCode.eMouseDownMsgCode, e.X, e.Y, e.Button)
            End If
        Else

            If moCurrentWindow Is Nothing = False Then
                'test for edge click
                If moSelectedControl Is Nothing = False Then
                    Dim oPos As Point

                    If mfrmProps Is Nothing Then mfrmProps = New frmProps()
                    If mfrmProps.Visible = False Then mfrmProps.Show(Me)
                    mfrmProps.Left = Me.Left + Me.Width
                    mfrmProps.Top = Me.Top
                    mfrmProps.ConfigProps(moSelectedControl)

                    oPos = moSelectedControl.GetAbsolutePosition()
                    mbResizeLeft = (e.X = oPos.X)
                    mbResizeRight = (e.X = (oPos.X + moSelectedControl.Width))
                    mbResizeTop = (e.Y = oPos.Y)
                    mbResizeBottom = (e.Y = (oPos.Y + moSelectedControl.Height))
                End If

                If mbResizeLeft = False And mbResizeRight = False And mbResizeTop = False And mbResizeBottom = False Then
                    mlSelectedControlIdx = -1
                    If moCurrentWindow.TestRegion(e.X, e.Y) Then
                        oTmp = moCurrentWindow

                        For X = 0 To oTmp.ChildrenUB
                            If oTmp.moChildren(X).TestRegion(e.X, e.Y) Then
                                oTmp = oTmp.moChildren(X)
                                mlSelectedControlIdx = X
                                Exit For
                            End If
                        Next X
                    End If
                    moSelectedControl = oTmp
                    If moSelectedControl Is Nothing = False Then
                        mlControlX = moSelectedControl.Left
                        mlControlY = moSelectedControl.Top
                        mbControlMoved = False
                    End If
                Else
                    AddUndoableAction(eLastActionType.eResize_LAT, moSelectedControl.Width, moSelectedControl.Height, moSelectedControl)
                End If
            End If
        End If
    End Sub

    Private Sub picRender_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles picRender.MouseMove
        Static xlLastX As Int32
        Static xlLastY As Int32

        Dim bMove As Boolean

        If mbTestMode = True Then
            If moCurrentWindow Is Nothing = False Then
                moCurrentWindow.PostMessage(UILibMsgCode.eMouseMoveMsgCode, e.X, e.Y, e.Button)
            End If
        Else

            If moSelectedControl Is Nothing = False AndAlso mbMouseDown Then
                bMove = Not (mbResizeLeft Or mbResizeRight Or mbResizeTop Or mbResizeBottom)
                If bMove Then
                    mbControlMoved = True
                    moSelectedControl.Left += e.X - xlLastX
                    moSelectedControl.Top += e.Y - xlLastY
                Else
                    If mbResizeLeft Then
                        moSelectedControl.Left += e.X - xlLastX
                        moSelectedControl.Width -= e.X - xlLastX
                    ElseIf mbResizeRight Then
                        moSelectedControl.Width += e.X - xlLastX
                    End If

                    If mbResizeTop Then
                        moSelectedControl.Top += e.Y - xlLastY
                        moSelectedControl.Height -= e.Y - xlLastY
                    ElseIf mbResizeBottom Then
                        moSelectedControl.Height += e.Y - xlLastY
                    End If
                End If
            ElseIf moSelectedControl Is Nothing = False Then
                Dim oPos As Point
                Dim bNS As Boolean
                Dim bEW As Boolean

                oPos = moSelectedControl.GetAbsolutePosition
                If e.X = oPos.X Or e.X = (oPos.X + moSelectedControl.Width) Then
                    bEW = True
                End If
                If e.Y = oPos.Y Or e.Y = (oPos.Y + moSelectedControl.Height) Then
                    bNS = True
                End If
                If bNS And bEW Then
                    System.Windows.Forms.Cursor.Current = Cursors.SizeAll
                ElseIf bNS Then
                    System.Windows.Forms.Cursor.Current = Cursors.SizeNS
                ElseIf bEW Then
                    System.Windows.Forms.Cursor.Current = Cursors.SizeWE
                Else
                    System.Windows.Forms.Cursor.Current = Cursors.Default
                End If
            End If

            xlLastX = e.X
            xlLastY = e.Y
        End If
    End Sub

    Private Sub picRender_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles picRender.MouseUp

        If mbTestMode = True Then
            If moCurrentWindow Is Nothing = False Then
                moCurrentWindow.PostMessage(UILibMsgCode.eMouseUpMsgCode, e.X, e.Y, e.Button)
            End If
        Else
            If moSelectedControl Is Nothing = False Then
                If mfrmProps Is Nothing Then mfrmProps = New frmProps()
                If mfrmProps.Visible = False Then mfrmProps.Show(Me)
                mfrmProps.Left = Me.Left + Me.Width
                mfrmProps.Top = Me.Top
                mfrmProps.ConfigProps(moSelectedControl)

                If mbControlMoved = True Then
                    AddUndoableAction(eLastActionType.eChangeLocation_LAT, mlControlX, mlControlY, moSelectedControl)
                    mlControlX = moSelectedControl.Left
                    mlControlY = moSelectedControl.Top
                End If
            Else
                If mfrmProps Is Nothing = False Then mfrmProps.Hide()
            End If

            mbControlMoved = False
            mbMouseDown = False
            mbResizeLeft = False
            mbResizeRight = False
            mbResizeTop = False
            mbResizeBottom = False
        End If

    End Sub

#End Region

    Private Sub mnuClearAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuClearAll.Click
        If MsgBox("Are you sure you wish to delete this window and all children?" & vbCrLf & "(This action cannot be undone later!)", MsgBoxStyle.YesNo Or MsgBoxStyle.Question, "Confirmation") = MsgBoxResult.Yes Then
            If moCurrentWindow Is Nothing = False Then
                moCurrentWindow.RemoveAllChildren()
                moSelectedControl = Nothing
                moCurrentWindow = Nothing
            End If

            'now, clear our undo buffer
            ReDim mo_LastActionObj(ml_UNDOABLE_ACTION_QUEUE_SIZE - 1)
            ReDim ml_LastActionType(ml_UNDOABLE_ACTION_QUEUE_SIZE - 1)
            ReDim ml_LastActionVal1(ml_UNDOABLE_ACTION_QUEUE_SIZE - 1)
            ReDim ml_LastActionVal2(ml_UNDOABLE_ACTION_QUEUE_SIZE - 1)
        End If
    End Sub

    Private Sub mnuCompile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCompile.Click
        'Ok, compile it...
        Dim sFinal As String
        Dim X As Int32

        sFinal = "Option Strict On" & vbCrLf & vbCrLf & "Imports Microsoft.DirectX" & vbCrLf & "Imports Microsoft.DirectX.Direct3D" & vbCrLf & vbCrLf
        sFinal &= "'Interface created from Interface Builder" & vbCrLf

        'Ok, add our Class for the window...
        If moCurrentWindow Is Nothing = False Then
            sFinal &= "Public Class " & moCurrentWindow.ControlName & vbCrLf

            sFinal &= vbTab & "Inherits UIWindow" & vbCrLf
            sFinal &= vbCrLf

            For X = 0 To moCurrentWindow.ChildrenUB
                Dim sType As String = Replace(moCurrentWindow.moChildren(X).GetType.ToString, "InterfaceBuilder.", "").Trim
                Dim sWE As String = "WithEvents "
                If sType.ToUpper = "UILABEL" OrElse sType.ToUpper = "UILINE" Then sWE = ""

                sFinal &= vbTab & "Private " & sWE & moCurrentWindow.moChildren(X).ControlName & " As " & sType & vbCrLf
            Next X

            sFinal &= vbTab & "Public Sub New(ByRef oUILib As UILib)" & vbCrLf
            sFinal &= vbTab & vbTab & "MyBase.New(oUILib)" & vbCrLf

            sFinal &= GetControlPropsString(moCurrentWindow, 2, True)

            'Now, the control properties...
            For X = 0 To moCurrentWindow.ChildrenUB
                sFinal &= GetControlPropsString(moCurrentWindow.moChildren(X), 2, False)
                sFinal &= vbTab & vbTab & "Me.AddChild(CType(" & moCurrentWindow.moChildren(X).ControlName & ", UIControl))" & vbCrLf
            Next X

            sFinal &= vbTab & "End Sub" & vbCrLf

            sFinal &= "End Class"
        End If

        If mfrmCompiled Is Nothing Then mfrmCompiled = New frmCompiled()
        mfrmCompiled.SetText(sFinal)
        mfrmCompiled.ShowDialog(Me)

    End Sub

    Private Sub mnuOpenForm_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuOpenForm.Click
        Dim sFinal As String
        Dim X As Int32
        Dim lTemp As Int32

        Dim oFileStream As System.IO.FileStream
        Dim oFileReader As System.IO.StreamReader

        Dim sLines() As String
        Dim sTemp As String
        Dim sType As String
        Dim sVals() As String '= Split(sTemp, ",")

        dlgOpen.DefaultExt = "txt"
        dlgOpen.Filter = "Text Files *.txt|*.txt|All Files *.*|*.*"
        dlgOpen.InitialDirectory = AppDomain.CurrentDomain.BaseDirectory
        dlgOpen.Title = "Open File..."
        dlgOpen.ValidateNames = True
        If dlgOpen.ShowDialog(Me) = Windows.Forms.DialogResult.OK Then
            oFileStream = CType(dlgOpen.OpenFile(), IO.FileStream)
            oFileReader = New System.IO.StreamReader(oFileStream)

            sFinal = oFileReader.ReadToEnd()

            oFileReader.Close()
            oFileReader = Nothing
            oFileStream.Close()
            oFileStream = Nothing

            'remove tabs and Public keywords
            sFinal = Replace$(sFinal, vbTab, "")

            'Now, parse the contents of sFinal
            'first, divide the string into substrings
            sLines = Split(sFinal, vbCrLf)

            'is the file empty?
            If sLines.Length = 0 Then
                MsgBox("Invalid Interface File.", MsgBoxStyle.Critical Or MsgBoxStyle.OkOnly, "Invalid")
                Exit Sub
            End If

            'release and recreate our window
            If moCurrentWindow Is Nothing = False Then
                moCurrentWindow.RemoveAllChildren()
            End If
            moCurrentWindow = Nothing
            moCurrentWindow = New UIWindow(moUILib)

            'Now, go thru each line
            For X = 0 To sLines.Length - 1
                sTemp = Trim$(sLines(X))

                If Mid$(sTemp, 1, 1) <> "'" AndAlso Len(sTemp) <> 0 Then
                    'Ok, not a commented line and not an empty line
                    'now, if Public Class is in the string
                    If InStr(1, sTemp, "Public Class", CompareMethod.Text) <> 0 Then
                        'this is our window definition
                        sTemp = Replace$(sTemp, "Public Class ", "", , , CompareMethod.Text)
                        moCurrentWindow.ControlName = Trim$(sTemp)
                    ElseIf InStr(1, sTemp, "End Class", CompareMethod.Text) <> 0 Then
                        'this is the end of our window defintion (and consequently our file)
                        Exit For
                    ElseIf InStr(1, sTemp, "Private WithEvents", CompareMethod.Text) <> 0 Then
                        'Ok, this is a control definition
                        sTemp = Trim$(Replace$(sTemp, "Private WithEvents", "", , , CompareMethod.Text))
                        sType = Split(sTemp, " As ", , CompareMethod.Text)(1)
                        sTemp = Split(sTemp, " As ", , CompareMethod.Text)(0)

                        Select Case UCase$(sType)
                            Case "UILABEL" : CreateNewControl("Label", sTemp)
                            Case "UITEXTBOX" : CreateNewControl("Text Box", sTemp)
                            Case "UIBUTTON" : CreateNewControl("Button", sTemp)
                            Case "UIOPTION" : CreateNewControl("Option Button", sTemp)
                            Case "UICHECKBOX" : CreateNewControl("Checkbox Button", sTemp)
                            Case "UISCROLLBAR" : CreateNewControl("Vertical Scroll Bar", sTemp)
                            Case "UILISTBOX" : CreateNewControl("List Box", sTemp)
                            Case "UICOMBOBOX" : CreateNewControl("Combo Box", sTemp)
                            Case "UIWINDOW" : CreateNewControl("Frame", sTemp)
                            Case "UILINE" : CreateNewControl("Line", sTemp)
                        End Select
                    ElseIf InStr(1, sTemp, "Public Sub New", CompareMethod.Text) <> 0 OrElse _
                           InStr(1, sTemp, "End Sub", CompareMethod.Text) <> 0 OrElse _
                           InStr(1, sTemp, "MyBase.New", CompareMethod.Text) <> 0 OrElse _
                           InStr(1, sTemp, "End With", CompareMethod.Text) <> 0 Then
                        'do nothing
                    ElseIf InStr(1, sTemp, "= New", CompareMethod.Text) <> 0 Then
                        'ok, new control already done with the Private WithEvents
                        ' however, if the control is a ScrollBar, we need to take more action
                        If InStr(1, sTemp, "UIScrollBar", CompareMethod.Text) <> 0 Then
                            'Ok, find the True or False
                            Dim bVal As Boolean = (InStr(1, sTemp, "True", CompareMethod.Text) <> 0)
                            'Now, get the scrollbar's name
                            sType = UCase$(Trim$(Mid$(sTemp, 1, InStr(1, sTemp, "=", CompareMethod.Text) - 1)))
                            For lTemp = 0 To moCurrentWindow.ChildrenUB
                                If moCurrentWindow.moChildren(lTemp) Is Nothing = False Then
                                    If UCase$(moCurrentWindow.moChildren(lTemp).ControlName) = sType Then
                                        CallByName(moCurrentWindow.moChildren(lTemp), "IsVertical", CallType.Let, bVal)
                                        Exit For
                                    End If
                                End If
                            Next lTemp
                        End If
                    ElseIf InStr(1, sTemp, "With Me", CompareMethod.Text) <> 0 Then
                        'ok, with our current window
                        moSelectedControl = moCurrentWindow
                    ElseIf InStr(1, sTemp, "With", CompareMethod.Text) <> 0 Then
                        'ok, find the control with that name
                        sTemp = UCase$(Trim$(Replace$(sTemp, "With ", "")))
                        For lTemp = 0 To moCurrentWindow.ChildrenUB
                            If moCurrentWindow.moChildren(lTemp) Is Nothing = False Then
                                If sTemp = UCase$(moCurrentWindow.moChildren(lTemp).ControlName) Then
                                    moSelectedControl = moCurrentWindow.moChildren(lTemp)
                                    Exit For
                                End If
                            End If
                        Next lTemp
                    ElseIf InStr(1, sTemp, "AddChild", CompareMethod.Text) <> 0 Then
                        'Do nothing, already handled for us
                    ElseIf Mid$(Trim$(sTemp), 1, 1) = "." Then
                        Dim sValName As String
                        Dim sValue As String

                        sTemp = Trim$(sTemp)

                        If Mid$(sTemp, 1, 8) = ".SetFont" Then
                            'Ok, got a setfont call
                            '.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10, 0, 3, 0))
                            sTemp = Replace$(sTemp, ".SetFont(New System.Drawing.Font", "", , , CompareMethod.Text)
                            'now, sTemp =
                            '("Microsoft Sans Serif", 10, 0, 3, 0))
                            sTemp = Replace$(Replace$(Replace$(sTemp, "(", ""), ")", ""), """", "")
                            'now, sTemp = 
                            'Microsoft Sans Serif, 10, 0, 3, 0
                            sVals = Split(sTemp, ",")
                            If sVals.Length >= 5 Then

                                Dim sFontStyle As String = Trim$(sVals(2))
                                Dim lFontStyle As FontStyle = FontStyle.Regular
                                If sFontStyle.ToLower.Contains("bold") = True Then
                                    lFontStyle = FontStyle.Bold
                                End If

                                Dim oFont As New System.Drawing.Font(Trim$(sVals(0)), CSng(Val(Trim$(sVals(1)))), lFontStyle, GraphicsUnit.Point, CByte(Val(Trim$(sVals(4)))))
                                CallByName(moSelectedControl, "SetFont", CallType.Method, oFont)
                            End If
                        ElseIf sTemp.ToUpper.Contains("FONTFORMAT") = True Then
                            lTemp = InStr(1, sTemp, "=", CompareMethod.Text)
                            sValue = Mid$(sTemp, lTemp + 1)
                            sValue = sValue.ToUpper.Replace("CTYPE(", "").Replace(", DRAWTEXTFORMAT)", "")
                            lTemp = CInt(Val(sValue))
                            CallByName(moSelectedControl, "FontFormat", CallType.Let, lTemp)
                        ElseIf UCase$(Mid$(sTemp, 1, 17)) = ".CONTROLIMAGERECT" Then
                            'Ok, the property is a rectangle
                            Dim rcTemp As Rectangle

                            lTemp = InStr(1, sTemp, "=", CompareMethod.Text)
                            sValName = Mid$(sTemp, 2, lTemp - 2)
                            sValue = Mid$(sTemp, lTemp + 1)

                            'sValName will be "ControlImageRect"
                            'sValue will be "Rectangle.FromLTRB(Left, Top, Width, Height)"
                            sValue = Replace$(Replace$(sValue, "Rectangle.FROMLTRB(", "", , , CompareMethod.Text), ")", "")
                            'now, sValue = "Left, Top, Width, Height"
                            sVals = Split(sValue, ",", , CompareMethod.Text)
                            If sVals.Length >= 4 Then
                                rcTemp = Rectangle.FromLTRB(CInt(Val(Trim$(sVals(0)))), CInt(Val(Trim$(sVals(1)))), CInt(Val(Trim$(sVals(2)))), CInt(Val(Trim$(sVals(3)))))
                                CallByName(moSelectedControl, sValName, CallType.Let, rcTemp)
                            End If
                        Else
                            lTemp = InStr(1, sTemp, "=", CompareMethod.Text)
							sValName = Mid$(sTemp, 2, lTemp - 2).Trim
                            sValue = Mid$(sTemp, lTemp + 1)

                            'Now, a few special cases for sValue...
                            If InStr(1, sValue, "System.Drawing.Color.FromARGB(", CompareMethod.Text) <> 0 Then
                                'ok, get the color result
                                sValue = Replace$(sValue, "System.Drawing.Color.FromARGB(", "", , , CompareMethod.Text)
                                sValue = Replace$(sValue, ")", "")

                                'Now, check for a comma
                                If InStr(sValue, ",", CompareMethod.Binary) = 0 Then
                                    'No comma, so use it straight up
                                    CallByName(moSelectedControl, sValName, CallType.Let, System.Drawing.Color.FromArgb(CInt(Val(sValue))))
                                Else
                                    'Ok, comma, set up our values
                                    Dim sColorVals() As String = Split(sValue, ",")
                                    Dim lAVal As Int32 = CInt(Val(Trim$(sColorVals(0))))
                                    Dim lRVal As Int32 = CInt(Val(Trim$(sColorVals(1))))
                                    Dim lGVal As Int32 = CInt(Val(Trim$(sColorVals(2))))
                                    Dim lBVal As Int32 = CInt(Val(Trim$(sColorVals(3))))

                                    CallByName(moSelectedControl, sValName, CallType.Let, System.Drawing.Color.FromArgb(lAVal, lRVal, lGVal, lBVal))
                                End If

                            ElseIf InStr(1, sValue, """", CompareMethod.Text) <> 0 Then
                                'ok, its a string value
                                sValue = Replace$(sValue, """", "")
                                CallByName(moSelectedControl, sValName, CallType.Let, sValue)
                            ElseIf UCase$(sValue) = "TRUE" Then
                                CallByName(moSelectedControl, sValName, CallType.Let, True)
                            ElseIf UCase$(sValue) = "FALSE" Then
                                CallByName(moSelectedControl, sValName, CallType.Let, False)
                            Else
                                Try
                                    CallByName(moSelectedControl, sValName, CallType.Let, sValue)
                                Catch
                                End Try
                            End If
                        End If


                    End If
                End If
            Next X

        End If

    End Sub

    Private Sub mnuSaveForm_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSaveForm.Click
        'just like compile
        Dim sFinal As String
        Dim X As Int32

        Dim oFileStream As System.IO.FileStream
        Dim oFileWriter As System.IO.StreamWriter

        dlgSave.DefaultExt = "txt"
        dlgSave.Filter = "Text Files *.txt|*.txt|All Files *.*|*.*"
        dlgSave.InitialDirectory = AppDomain.CurrentDomain.BaseDirectory
        dlgSave.Title = "Save File As..."
        dlgSave.ValidateNames = True
        If dlgSave.ShowDialog(Me) = Windows.Forms.DialogResult.OK Then
            oFileStream = CType(dlgSave.OpenFile(), IO.FileStream)
            oFileWriter = New System.IO.StreamWriter(oFileStream)
            oFileWriter.AutoFlush = True

            sFinal = "Option Strict On" & vbCrLf & vbCrLf & "Imports Microsoft.DirectX" & vbCrLf & "Imports Microsoft.DirectX.Direct3D" & vbCrLf & vbCrLf
            sFinal &= "'Interface created from Interface Builder" & vbCrLf

            'Ok, add our Class for the window...
            If moCurrentWindow Is Nothing = False Then
                sFinal &= "Public Class " & moCurrentWindow.ControlName & vbCrLf

                sFinal &= vbTab & "Inherits UIWindow" & vbCrLf
                sFinal &= vbCrLf

                For X = 0 To moCurrentWindow.ChildrenUB
                    Dim sType As String = Replace(moCurrentWindow.moChildren(X).GetType.ToString, "InterfaceBuilder.", "").Trim
                    Dim sWE As String = "WithEvents "
                    If sType.ToUpper = "UILABEL" OrElse sType.ToUpper = "UILINE" Then sWE = ""

                    sFinal &= vbTab & "Private " & sWE & moCurrentWindow.moChildren(X).ControlName & " As " & sType & vbCrLf
                Next X

                sFinal &= vbTab & "Public Sub New(ByRef oUILib As UILib)" & vbCrLf
                sFinal &= vbTab & vbTab & "MyBase.New(oUILib)" & vbCrLf

                sFinal &= GetControlPropsString(moCurrentWindow, 2, True)

                'Now, the control properties...
                For X = 0 To moCurrentWindow.ChildrenUB
                    sFinal &= GetControlPropsString(moCurrentWindow.moChildren(X), 2, False)
                    sFinal &= vbTab & vbTab & "Me.AddChild(CType(" & moCurrentWindow.moChildren(X).ControlName & ", UIControl))" & vbCrLf
                Next X

                sFinal &= vbTab & "End Sub" & vbCrLf

                sFinal &= "End Class"
            End If

            oFileWriter.Write(sFinal)
            oFileWriter.Flush()
            oFileWriter.Close()
            oFileWriter = Nothing
            oFileStream.Close()
            oFileStream = Nothing
        End If

    End Sub

    Private Function GetControlPropsString(ByVal oControl As Object, ByVal lTabCnt As Int32, ByVal bUseMe As Boolean) As String
        Dim sRes As String = ""

        With CType(oControl, UIControl)
            sRes &= vbCrLf & StrDup(lTabCnt, vbTab) & "'" & .ControlName & " initial props" & vbCrLf
            If bUseMe = False Then
                sRes &= StrDup(lTabCnt, vbTab) & .ControlName & " = New " & _
                  Replace(oControl.GetType.ToString, "InterfaceBuilder.", "") & "(oUILib"
                If oControl.GetType.ToString = "InterfaceBuilder.UIScrollBar" Then
                    sRes &= ", " & CStr(CallByName(oControl, "IsVertical", CallType.Get))
                End If
                sRes &= ")" & vbCrLf
                sRes &= StrDup(lTabCnt, vbTab) & "With " & .ControlName & vbCrLf
            Else
                sRes &= StrDup(lTabCnt, vbTab) & "With Me" & vbCrLf
            End If

            lTabCnt += 1

            sRes &= StrDup(lTabCnt, vbTab) & ".ControlName=""" & .ControlName & """" & vbCrLf
            sRes &= StrDup(lTabCnt, vbTab) & ".Left=" & .Left & vbCrLf
            sRes &= StrDup(lTabCnt, vbTab) & ".Top=" & .Top & vbCrLf
            sRes &= StrDup(lTabCnt, vbTab) & ".Width=" & .Width & vbCrLf
            sRes &= StrDup(lTabCnt, vbTab) & ".Height=" & .Height & vbCrLf
            sRes &= StrDup(lTabCnt, vbTab) & ".Enabled=" & .Enabled & vbCrLf
            sRes &= StrDup(lTabCnt, vbTab) & ".Visible=" & .Visible & vbCrLf
        End With


        Select Case oControl.GetType.ToString
            Case "InterfaceBuilder.UILine"
                With CType(oControl, UILine)
                    sRes &= StrDup$(lTabCnt, vbTab) & ".BorderColor = muSettings.InterfaceBorderColor" & vbCrLf
                End With
            Case "InterfaceBuilder.UIWindow"
                With CType(oControl, UIWindow)
                    sRes &= StrDup(lTabCnt, vbTab) & ".BorderColor = muSettings.InterfaceBorderColor" & vbCrLf
                    sRes &= StrDup(lTabCnt, vbTab) & ".FillColor = muSettings.InterfaceFillColor" & vbCrLf
                    sRes &= StrDup(lTabCnt, vbTab) & ".FullScreen=" & .FullScreen & vbCrLf
                End With
            Case "InterfaceBuilder.UILabel", "InterfaceBuilder.UIButton", "InterfaceBuilder.UITextBox", "InterfaceBuilder.UIOption", "InterfaceBuilder.UICheckBox"
                With CType(oControl, UILabel)
                    sRes &= StrDup(lTabCnt, vbTab) & ".Caption=""" & .Caption & """" & vbCrLf
                    If oControl.GetType.ToString = "InterfaceBuilder.UITextBox" Then
                        sRes &= StrDup(lTabCnt, vbTab) & ".ForeColor = muSettings.InterfaceTextBoxForeColor" & vbCrLf
                    Else
                        sRes &= StrDup(lTabCnt, vbTab) & ".ForeColor = muSettings.InterfaceBorderColor" & vbCrLf
                    End If
                    sRes &= StrDup(lTabCnt, vbTab) & ".SetFont(New System.Drawing.Font(""" & .GetFont.Name & """, " & .GetFont.Size & "F, FontStyle." & .GetFont.Style.ToString & ", GraphicsUnit.Point, 0))" & vbCrLf
                    sRes &= StrDup(lTabCnt, vbTab) & ".DrawBackImage=" & .DrawBackImage & vbCrLf
                    sRes &= StrDup(lTabCnt, vbTab) & ".FontFormat=CType(" & .FontFormat & ", DrawTextFormat)" & vbCrLf
                    If .DrawBackImage = True Then sRes &= StrDup(lTabCnt, vbTab) & ".ControlImageRect = New Rectangle(" & .ControlImageRect.X & ", " & .ControlImageRect.Y & ", " & .ControlImageRect.Width & ", " & .ControlImageRect.Height & ")" & vbCrLf
                End With
                If oControl.GetType.ToString = "InterfaceBuilder.UITextBox" Then
                    With CType(oControl, UITextBox)
                        sRes &= StrDup(lTabCnt, vbTab) & ".BackColorEnabled = muSettings.InterfaceTextBoxFillColor" & vbCrLf
                        sRes &= StrDup(lTabCnt, vbTab) & ".BackColorDisabled=System.Drawing.Color.FromArgb(" & .BackColorDisabled.A & ", " & .BackColorDisabled.R & ", " & .BackColorDisabled.G & ", " & .BackColorDisabled.B & ")" & vbCrLf
                        sRes &= StrDup(lTabCnt, vbTab) & ".MaxLength=" & .MaxLength & vbCrLf
                        sRes &= StrDup(lTabCnt, vbTab) & ".BorderColor = muSettings.InterfaceBorderColor" & vbCrLf
                    End With
                ElseIf oControl.GetType.ToString = "InterfaceBuilder.UIOption" Or oControl.GetType.ToString = "InterfaceBuilder.UICheckBox" Then
                    sRes &= StrDup(lTabCnt, vbTab) & ".Value=" & CBool(CallByName(oControl, "Value", CallType.Get)) & vbCrLf
                End If
            Case "InterfaceBuilder.UIScrollBar"
                With CType(oControl, UIScrollBar)
                    sRes &= StrDup(lTabCnt, vbTab) & ".Value=" & .Value & vbCrLf
                    sRes &= StrDup(lTabCnt, vbTab) & ".MaxValue=" & .MaxValue & vbCrLf
                    sRes &= StrDup(lTabCnt, vbTab) & ".MinValue=" & .MinValue & vbCrLf
                    sRes &= StrDup(lTabCnt, vbTab) & ".SmallChange=" & .SmallChange & vbCrLf
                    sRes &= StrDup(lTabCnt, vbTab) & ".LargeChange=" & .LargeChange & vbCrLf
                    sRes &= StrDup(lTabCnt, vbTab) & ".ReverseDirection=" & .ReverseDirection & vbCrLf
                End With
            Case "InterfaceBuilder.UIComboBox"
                With CType(oControl, UIComboBox)
                    sRes &= StrDup(lTabCnt, vbTab) & ".BorderColor = muSettings.InterfaceBorderColor" & vbCrLf
                    sRes &= StrDup(lTabCnt, vbTab) & ".FillColor=muSettings.InterfaceTextBoxFillColor" & vbCrLf 'System.Drawing.Color.FromArgb(" & .FillColor.A & ", " & .FillColor.R & ", " & .FillColor.G & ", " & .FillColor.B & ")" & vbCrLf
                    sRes &= StrDup(lTabCnt, vbTab) & ".ForeColor=muSettings.InterfaceBorderColor" & vbCrLf 'System.Drawing.Color.FromArgb(" & .ForeColor.A & ", " & .ForeColor.R & ", " & .ForeColor.G & ", " & .ForeColor.B & ")" & vbCrLf
                    sRes &= StrDup(lTabCnt, vbTab) & ".SetFont(New System.Drawing.Font(""" & .GetFont.Name & """, " & .GetFont.Size & "F, FontStyle." & .GetFont.Style.ToString & ", GraphicsUnit.Point, 0))" & vbCrLf
                    sRes &= StrDup(lTabCnt, vbTab) & ".HighlightColor=System.Drawing.Color.FromArgb(" & .HighlightColor.A & ", " & .HighlightColor.R & ", " & .HighlightColor.G & ", " & .HighlightColor.B & ")" & vbCrLf
                    sRes &= StrDup(lTabCnt, vbTab) & ".DropDownListBorderColor=muSettings.InterfaceBorderColor" & vbCrLf 'System.Drawing.Color.FromArgb(" & .DropDownListBorderColor.A & ", " & .DropDownListBorderColor.R & ", " & .DropDownListBorderColor.G & ", " & .DropDownListBorderColor.B & ")" & vbCrLf
                End With
            Case "InterfaceBuilder.UIListBox"
                With CType(oControl, UIListBox)
                    sRes &= StrDup(lTabCnt, vbTab) & ".BorderColor = muSettings.InterfaceBorderColor" & vbCrLf
                    sRes &= StrDup(lTabCnt, vbTab) & ".FillColor = muSettings.InterfaceFillColor" & vbCrLf
                    sRes &= StrDup(lTabCnt, vbTab) & ".ForeColor = muSettings.InterfaceBorderColor" & vbCrLf
                    sRes &= StrDup(lTabCnt, vbTab) & ".SetFont(New System.Drawing.Font(""" & .GetFont.Name & """, " & .GetFont.Size & "F, FontStyle." & .GetFont.Style.ToString & ", GraphicsUnit.Point, 0))" & vbCrLf
                    sRes &= StrDup(lTabCnt, vbTab) & ".HighlightColor=System.Drawing.Color.FromArgb(" & .HighlightColor.A & ", " & .HighlightColor.R & ", " & .HighlightColor.G & ", " & .HighlightColor.B & ")" & vbCrLf
                End With
        End Select

        lTabCnt -= 1
        sRes &= StrDup(lTabCnt, vbTab) & "End With" & vbCrLf

        Return sRes

    End Function

    Private Sub mnuNewFullScreen_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuNewFullScreen.Click
        'TODO: Check here if they need to save the window
        moCurrentWindow = Nothing

        moCurrentWindow = New UIWindow(moUILib)
        moCurrentWindow.Left = 0
        moCurrentWindow.Top = 0
        moCurrentWindow.Height = 600
        moCurrentWindow.Width = 800
        moCurrentWindow.Enabled = True
        moCurrentWindow.Visible = True
        moCurrentWindow.FullScreen = True

        'moUILib.WindowList.Add(moCurrentWindow)
    End Sub

    Private Sub mnuNewNormal_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuNewNormal.Click
        'TODO: Check here if they need to save the window
        moCurrentWindow = Nothing

        moCurrentWindow = New UIWindow(moUILib)
        moCurrentWindow.Left = 0
        moCurrentWindow.Top = 0
        moCurrentWindow.Height = 100
        moCurrentWindow.Width = 100
        moCurrentWindow.Enabled = True
        moCurrentWindow.Visible = True
        moCurrentWindow.FullScreen = False

        'moUILib.WindowList.Add(moCurrentWindow)
    End Sub

    Private Sub mnuTestMode_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuTestMode.Click
        mbTestMode = Not mbTestMode
        If mbTestMode Then
            mnuTestMode.Text = "Exit Test Mode..."
        Else
            mnuTestMode.Text = "Enter Test Mode..."
        End If

        If mbTestMode = True Then
            Dim X As Int32
            Dim Y As Int32

            For X = 0 To moCurrentWindow.ChildrenUB
                If moCurrentWindow.moChildren(X).GetType.ToString = "InterfaceBuilder.UIListBox" Then
                    With CType(moCurrentWindow.moChildren(X), UIListBox)
                        .Clear()
                        For Y = 0 To 10
                            .AddItem("Item " & Y)
                        Next Y
                    End With
                End If
            Next X

        End If

        'and clear all flags
        mbResizeLeft = False
        mbResizeRight = False
        mbResizeTop = False
        mbResizeBottom = False
        moSelectedControl = Nothing
        If mfrmProps Is Nothing = False Then mfrmProps.Hide()
    End Sub

    Private Sub mfrmProps_UpdateFontFormat(ByVal oNewFmt As Microsoft.DirectX.Direct3D.DrawTextFormat) Handles mfrmProps.UpdateFontFormat
        Dim oObj As Object = moSelectedControl
        CallByName(oObj, "FontFormat", CallType.Let, oNewFmt)
    End Sub

    Private Sub mnuUndo_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuUndo.Click
        'ok, store are vals
        Dim lAction As eLastActionType = ml_LastActionType(0)
        Dim lVal1 As Int32 = ml_LastActionVal1(0)
        Dim lVal2 As Int32 = ml_LastActionVal2(0)
        Dim oObj As Object = mo_LastActionObj(0)

        Dim X As Int32

        'now, shift our values
        For X = 0 To ml_UNDOABLE_ACTION_QUEUE_SIZE - 2
            ml_LastActionType(X) = ml_LastActionType(X + 1)
            ml_LastActionVal1(X) = ml_LastActionVal1(X + 1)
            ml_LastActionVal2(X) = ml_LastActionVal2(X + 1)
            mo_LastActionObj(X) = mo_LastActionObj(X + 1)
        Next X

        'set the last value to nothing
        ml_LastActionType(ml_UNDOABLE_ACTION_QUEUE_SIZE - 1) = eLastActionType.eNoAction_LAT
        ml_LastActionVal1(ml_UNDOABLE_ACTION_QUEUE_SIZE - 1) = 0
        ml_LastActionVal2(ml_UNDOABLE_ACTION_QUEUE_SIZE - 1) = 0
        mo_LastActionObj(ml_UNDOABLE_ACTION_QUEUE_SIZE - 1) = Nothing

        If oObj Is Nothing = False Then
            Select Case lAction
                Case eLastActionType.eChangeLocation_LAT
                    CallByName(oObj, "Left", CallType.Let, lVal1)
                    CallByName(oObj, "Top", CallType.Let, lVal2)
                Case eLastActionType.eEnabled_LAT
                    CallByName(oObj, "Enabled", CallType.Let, CBool(lVal1))
                Case eLastActionType.eResize_LAT
                    CallByName(oObj, "Width", CallType.Let, lVal1)
                    CallByName(oObj, "Height", CallType.Let, lVal2)
                Case eLastActionType.eVisible_LAT
                    CallByName(oObj, "Visible", CallType.Let, CBool(lVal1))
                Case eLastActionType.eDelete_LAT
                    If moCurrentWindow Is Nothing = False Then
                        moCurrentWindow.AddChild(CType(oObj, UIControl))
                        moSelectedControl = CType(oObj, UIControl)
                    End If
            End Select
        End If

        If moSelectedControl Is Nothing = False Then mfrmProps.ConfigProps(moSelectedControl)

    End Sub

    Private Sub mnuCopy_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuCopy.Click
        moCopyObj = moSelectedControl
    End Sub

    Private Sub mnuPaste_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuPaste.Click
        Dim oObj As Object

        If moCopyObj Is Nothing = False Then
            Select Case moCopyObj.GetType.ToString
                Case "InterfaceBuilder.UIWindow" : oObj = New UIWindow(moUILib)
                Case "InterfaceBuilder.UILabel" : oObj = New UILabel(moUILib)
                Case "InterfaceBuilder.UIButton" : oObj = New UIButton(moUILib)
                Case "InterfaceBuilder.UITextBox" : oObj = New UITextBox(moUILib)
                Case "InterfaceBuilder.UIOption" : oObj = New UIOption(moUILib)
                Case "InterfaceBuilder.UICheckBox" : oObj = New UICheckBox(moUILib)
                Case "InterfaceBuilder.UIScrollBar" : oObj = New UIScrollBar(moUILib, CType(moCopyObj, UIScrollBar).IsVertical)
                Case "InterfaceBuilder.UIListBox" : oObj = New UIListBox(moUILib)
                Case "InterfaceBuilder.UIComboBox" : oObj = New UIComboBox(moUILib)
                Case "InterfaceBuilder.UILine" : oObj = New UILine(moUILib)
                Case Else
                    Exit Sub
            End Select

            With CType(oObj, UIControl)
                .ControlName = "New Control"
                .Enabled = CType(moCopyObj, UIControl).Enabled
                .Height = CType(moCopyObj, UIControl).Height
                .Width = CType(moCopyObj, UIControl).Width
                .Left = 0
                .Top = 0
                .Visible = CType(moCopyObj, UIControl).Visible
            End With

            'Now, do our specifics...
            Select Case moCopyObj.GetType.ToString
                Case "InterfaceBuilder.UILine"
                    CType(oObj, UILine).BorderColor = CType(moCopyObj, UILine).BorderColor
                Case "InterfaceBuilder.UIWindow"
                    CType(oObj, UIWindow).BorderColor = CType(moCopyObj, UIWindow).BorderColor
                    CType(oObj, UIWindow).FillColor = CType(moCopyObj, UIWindow).FillColor

                Case "InterfaceBuilder.UILabel", "InterfaceBuilder.UIButton", "InterfaceBuilder.UITextBox", "InterfaceBuilder.UIOption", "InterfaceBuilder.UICheckBox"
                    With CType(oObj, UILabel)
                        .Caption = CType(moCopyObj, UILabel).Caption
                        .ForeColor = CType(moCopyObj, UILabel).ForeColor
                        .SetFont(CType(moCopyObj, UILabel).GetFont)
                        .FontFormat = CType(moCopyObj, UILabel).FontFormat
                        .DrawBackImage = CType(moCopyObj, UILabel).DrawBackImage
                    End With

                    If moCopyObj.GetType.ToString = "InterfaceBuilder.UITextBox" Then
                        With CType(oObj, UITextBox)
                            .BackColorEnabled = CType(moCopyObj, UITextBox).BackColorEnabled
                            .BackColorDisabled = CType(moCopyObj, UITextBox).BackColorDisabled
                            .MaxLength = CType(moCopyObj, UITextBox).MaxLength
                            .BorderColor = CType(moCopyObj, UITextBox).BorderColor
                        End With
                    ElseIf moCopyObj.GetType.ToString = "InterfaceBuilder.UIOption" Or moCopyObj.GetType.ToString = "InterfaceBuilder.UICheckBox" Then
                        CallByName(oObj, "Value", CallType.Let, CallByName(moCopyObj, "Value", CallType.Get))
                    End If
                Case "InterfaceBuilder.UIScrollBar"
                    With CType(oObj, UIScrollBar)
                        .MaxValue = CType(moCopyObj, UIScrollBar).MaxValue
                        .Value = CType(moCopyObj, UIScrollBar).Value
                        .MinValue = CType(moCopyObj, UIScrollBar).MinValue
                        .SmallChange = CType(moCopyObj, UIScrollBar).SmallChange
                        .LargeChange = CType(moCopyObj, UIScrollBar).LargeChange
                        .ReverseDirection = CType(moCopyObj, UIScrollBar).ReverseDirection
                    End With
                Case "InterfaceBuilder.UIListBox"
                    With CType(oObj, UIListBox)
                        .BorderColor = CType(moCopyObj, UIListBox).BorderColor
                        .FillColor = CType(moCopyObj, UIListBox).FillColor
                        .ForeColor = CType(moCopyObj, UIListBox).ForeColor
                        .SetFont(CType(moCopyObj, UIListBox).GetFont)
                        .HighlightColor = CType(moCopyObj, UIListBox).HighlightColor
                    End With
                Case "InterfaceBuilder.UIComboBox"
                    With CType(oObj, UIComboBox)
                        .BorderColor = CType(moCopyObj, UIComboBox).BorderColor
                        .FillColor = CType(moCopyObj, UIComboBox).FillColor
                        .ForeColor = CType(moCopyObj, UIComboBox).ForeColor
                        .SetFont(CType(moCopyObj, UIComboBox).GetFont)
                        .HighlightColor = CType(moCopyObj, UIComboBox).HighlightColor
                        .DropDownListBorderColor = CType(moCopyObj, UIComboBox).DropDownListBorderColor
                    End With
            End Select

            If moCurrentWindow Is Nothing = False Then
                moCurrentWindow.AddChild(CType(oObj, UIControl))
            End If

            If oObj Is Nothing = False Then
                moSelectedControl = CType(oObj, UIControl)
            End If

            RefreshComboList()

            oObj = Nothing

        End If

    End Sub

    Private Sub mnuDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuDelete.Click
        DeleteSelected()
    End Sub

    Private Sub DeleteSelected()
        If moSelectedControl Is Nothing = False Then
            If mlSelectedControlIdx <> -1 Then
                AddUndoableAction(eLastActionType.eDelete_LAT, 0, 0, moSelectedControl)

                moSelectedControl = Nothing
                moCurrentWindow.RemoveChild(mlSelectedControlIdx)
                mlSelectedControlIdx = -1
            Else
                AddUndoableAction(eLastActionType.eDelete_LAT, 0, 0, moCurrentWindow)
                'remove the window
                moCurrentWindow.RemoveAllChildren()
                moSelectedControl = Nothing
                moCurrentWindow = Nothing
            End If
        End If
    End Sub

    Private Sub MenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem2.Click
        Dim lR As Int32
        Dim lG As Int32
        Dim lB As Int32
        Dim lA As Int32

        Dim oColor As System.Drawing.Color
        Dim sProperty As String = ""

        oColor = System.Drawing.Color.FromArgb(lA, lR, lG, lB)

        CallByName(moSelectedControl, sProperty, CallType.Let, oColor)
    End Sub

    Private Sub cboControls_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboControls.SelectedIndexChanged
        Dim sCtrlName As String
        Dim sCtrlType As String
        Dim sClick As String = cboControls.Text
        Dim lTemp As Int32 = InStr(1, sClick, "(", CompareMethod.Text)
        'ControlName (type)

        sCtrlName = Mid$(sClick, 1, lTemp - 2)
        sCtrlType = "InterfaceBuilder." & Mid$(sClick, lTemp + 1, Len(sClick) - lTemp - 1)

        If moCurrentWindow Is Nothing = False Then
            If sCtrlName = moCurrentWindow.ControlName Then
                If sCtrlType = moCurrentWindow.GetType.ToString Then
                    moSelectedControl = moCurrentWindow
                    If mfrmProps Is Nothing Then mfrmProps = New frmProps()
                    mfrmProps.Show(Me)
                    mfrmProps.Left = Me.Left + Me.Width
                    mfrmProps.Top = Me.Top
                    mfrmProps.ConfigProps(moSelectedControl)
                    Exit Sub
                End If
            End If

            For lTemp = 0 To moCurrentWindow.ChildrenUB
                If moCurrentWindow.moChildren(lTemp) Is Nothing = False Then
                    If moCurrentWindow.moChildren(lTemp).ControlName = sCtrlName Then
                        If moCurrentWindow.moChildren(lTemp).GetType.ToString = sCtrlType Then
                            'ok, found it
                            moSelectedControl = moCurrentWindow.moChildren(lTemp)
                            If mfrmProps Is Nothing Then mfrmProps = New frmProps()
                            mfrmProps.Show(Me)
                            mfrmProps.Left = Me.Left + Me.Width
                            mfrmProps.Top = Me.Top
                            mfrmProps.ConfigProps(moSelectedControl)
                            Exit For
                        End If
                    End If
                End If
            Next lTemp
        End If
    End Sub

    Private Sub MenuItem3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem3.Click
        Dim oTmp As UIHullSlots = New UIHullSlots(moUILib)

        With oTmp
            .Visible = True
            .Enabled = True
        End With

        moCurrentWindow.AddChild(CType(oTmp, UIControl))
        oTmp = Nothing
    End Sub

    Private Sub MenuItem5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem5.Click
        Dim oTmp As UILabelScroller = New UILabelScroller(moUILib)

        With oTmp
            .Visible = True
            .Enabled = True
        End With

        moCurrentWindow.AddChild(CType(oTmp, UIControl))
        oTmp = Nothing
    End Sub
End Class