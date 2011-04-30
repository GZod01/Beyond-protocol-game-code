Public Class frmProps
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
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtName As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtLeft As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtTop As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtWidth As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtHeight As System.Windows.Forms.TextBox
    Friend WithEvents chkVisible As System.Windows.Forms.CheckBox
    Friend WithEvents chkEnabled As System.Windows.Forms.CheckBox
    Friend WithEvents clrdlg1 As System.Windows.Forms.ColorDialog
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents btnBorderClr As System.Windows.Forms.Button
    Friend WithEvents btnFillClr As System.Windows.Forms.Button
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents txtCaption As System.Windows.Forms.TextBox
    Friend WithEvents btnForeClr As System.Windows.Forms.Button
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents btnFont As System.Windows.Forms.Button
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents chkDrawBackImage As System.Windows.Forms.CheckBox
    Friend WithEvents btnBackClrEnabled As System.Windows.Forms.Button
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents btnBackClrDisabled As System.Windows.Forms.Button
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents txtMaxLen As System.Windows.Forms.TextBox
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents txtValue As System.Windows.Forms.TextBox
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents txtMaxValue As System.Windows.Forms.TextBox
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents txtMinValue As System.Windows.Forms.TextBox
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents txtSmallChng As System.Windows.Forms.TextBox
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents txtLargeChng As System.Windows.Forms.TextBox
    Friend WithEvents chkReverseDirection As System.Windows.Forms.CheckBox
    Friend WithEvents FontDlg As System.Windows.Forms.FontDialog
    Friend WithEvents btnHighlight As System.Windows.Forms.Button
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents btnDropDownBorder As System.Windows.Forms.Button
    Friend WithEvents Label20 As System.Windows.Forms.Label
    Friend WithEvents btnSetRect As System.Windows.Forms.Button
    Friend WithEvents btnAlign As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.txtName = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.txtLeft = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.txtTop = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.txtWidth = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.txtHeight = New System.Windows.Forms.TextBox
        Me.chkVisible = New System.Windows.Forms.CheckBox
        Me.chkEnabled = New System.Windows.Forms.CheckBox
        Me.clrdlg1 = New System.Windows.Forms.ColorDialog
        Me.Label7 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.btnBorderClr = New System.Windows.Forms.Button
        Me.btnFillClr = New System.Windows.Forms.Button
        Me.Label8 = New System.Windows.Forms.Label
        Me.txtCaption = New System.Windows.Forms.TextBox
        Me.btnForeClr = New System.Windows.Forms.Button
        Me.Label9 = New System.Windows.Forms.Label
        Me.btnFont = New System.Windows.Forms.Button
        Me.Label10 = New System.Windows.Forms.Label
        Me.chkDrawBackImage = New System.Windows.Forms.CheckBox
        Me.btnBackClrEnabled = New System.Windows.Forms.Button
        Me.Label11 = New System.Windows.Forms.Label
        Me.btnBackClrDisabled = New System.Windows.Forms.Button
        Me.Label12 = New System.Windows.Forms.Label
        Me.Label13 = New System.Windows.Forms.Label
        Me.txtMaxLen = New System.Windows.Forms.TextBox
        Me.Label14 = New System.Windows.Forms.Label
        Me.txtValue = New System.Windows.Forms.TextBox
        Me.Label15 = New System.Windows.Forms.Label
        Me.txtMaxValue = New System.Windows.Forms.TextBox
        Me.Label16 = New System.Windows.Forms.Label
        Me.txtMinValue = New System.Windows.Forms.TextBox
        Me.Label17 = New System.Windows.Forms.Label
        Me.txtSmallChng = New System.Windows.Forms.TextBox
        Me.Label18 = New System.Windows.Forms.Label
        Me.txtLargeChng = New System.Windows.Forms.TextBox
        Me.chkReverseDirection = New System.Windows.Forms.CheckBox
        Me.FontDlg = New System.Windows.Forms.FontDialog
        Me.btnHighlight = New System.Windows.Forms.Button
        Me.Label19 = New System.Windows.Forms.Label
        Me.btnDropDownBorder = New System.Windows.Forms.Button
        Me.Label20 = New System.Windows.Forms.Label
        Me.btnAlign = New System.Windows.Forms.Button
        Me.btnSetRect = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'txtName
        '
        Me.txtName.Location = New System.Drawing.Point(64, 8)
        Me.txtName.Name = "txtName"
        Me.txtName.Size = New System.Drawing.Size(120, 20)
        Me.txtName.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(8, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(40, 24)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Name:"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(8, 32)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(40, 24)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Left:"
        '
        'txtLeft
        '
        Me.txtLeft.Location = New System.Drawing.Point(64, 32)
        Me.txtLeft.Name = "txtLeft"
        Me.txtLeft.Size = New System.Drawing.Size(120, 20)
        Me.txtLeft.TabIndex = 2
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(8, 56)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(40, 24)
        Me.Label3.TabIndex = 5
        Me.Label3.Text = "Top:"
        '
        'txtTop
        '
        Me.txtTop.Location = New System.Drawing.Point(64, 56)
        Me.txtTop.Name = "txtTop"
        Me.txtTop.Size = New System.Drawing.Size(120, 20)
        Me.txtTop.TabIndex = 4
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(8, 80)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(40, 24)
        Me.Label4.TabIndex = 7
        Me.Label4.Text = "Width:"
        '
        'txtWidth
        '
        Me.txtWidth.Location = New System.Drawing.Point(64, 80)
        Me.txtWidth.Name = "txtWidth"
        Me.txtWidth.Size = New System.Drawing.Size(120, 20)
        Me.txtWidth.TabIndex = 6
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(8, 104)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(40, 24)
        Me.Label5.TabIndex = 9
        Me.Label5.Text = "Height:"
        '
        'txtHeight
        '
        Me.txtHeight.Location = New System.Drawing.Point(64, 104)
        Me.txtHeight.Name = "txtHeight"
        Me.txtHeight.Size = New System.Drawing.Size(120, 20)
        Me.txtHeight.TabIndex = 8
        '
        'chkVisible
        '
        Me.chkVisible.Location = New System.Drawing.Point(8, 128)
        Me.chkVisible.Name = "chkVisible"
        Me.chkVisible.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.chkVisible.Size = New System.Drawing.Size(72, 24)
        Me.chkVisible.TabIndex = 10
        Me.chkVisible.Text = "Visible"
        '
        'chkEnabled
        '
        Me.chkEnabled.Location = New System.Drawing.Point(112, 128)
        Me.chkEnabled.Name = "chkEnabled"
        Me.chkEnabled.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.chkEnabled.Size = New System.Drawing.Size(72, 24)
        Me.chkEnabled.TabIndex = 11
        Me.chkEnabled.Text = "Enabled"
        '
        'clrdlg1
        '
        Me.clrdlg1.AnyColor = True
        '
        'Label7
        '
        Me.Label7.Location = New System.Drawing.Point(8, 183)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(48, 24)
        Me.Label7.TabIndex = 20
        Me.Label7.Text = "Fill:"
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(8, 152)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(48, 24)
        Me.Label6.TabIndex = 18
        Me.Label6.Text = "Border:"
        '
        'btnBorderClr
        '
        Me.btnBorderClr.BackColor = System.Drawing.SystemColors.Control
        Me.btnBorderClr.Location = New System.Drawing.Point(64, 152)
        Me.btnBorderClr.Name = "btnBorderClr"
        Me.btnBorderClr.Size = New System.Drawing.Size(120, 23)
        Me.btnBorderClr.TabIndex = 21
        Me.btnBorderClr.UseVisualStyleBackColor = False
        '
        'btnFillClr
        '
        Me.btnFillClr.Location = New System.Drawing.Point(64, 180)
        Me.btnFillClr.Name = "btnFillClr"
        Me.btnFillClr.Size = New System.Drawing.Size(120, 23)
        Me.btnFillClr.TabIndex = 22
        '
        'Label8
        '
        Me.Label8.Location = New System.Drawing.Point(8, 360)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(48, 24)
        Me.Label8.TabIndex = 24
        Me.Label8.Text = "Caption:"
        '
        'txtCaption
        '
        Me.txtCaption.Location = New System.Drawing.Point(64, 360)
        Me.txtCaption.Name = "txtCaption"
        Me.txtCaption.Size = New System.Drawing.Size(120, 20)
        Me.txtCaption.TabIndex = 23
        '
        'btnForeClr
        '
        Me.btnForeClr.Location = New System.Drawing.Point(64, 208)
        Me.btnForeClr.Name = "btnForeClr"
        Me.btnForeClr.Size = New System.Drawing.Size(120, 23)
        Me.btnForeClr.TabIndex = 26
        '
        'Label9
        '
        Me.Label9.Location = New System.Drawing.Point(8, 208)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(48, 24)
        Me.Label9.TabIndex = 25
        Me.Label9.Text = "Fore:"
        '
        'btnFont
        '
        Me.btnFont.Location = New System.Drawing.Point(64, 384)
        Me.btnFont.Name = "btnFont"
        Me.btnFont.Size = New System.Drawing.Size(64, 23)
        Me.btnFont.TabIndex = 28
        Me.btnFont.Text = "Set Font"
        '
        'Label10
        '
        Me.Label10.Location = New System.Drawing.Point(8, 384)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(48, 24)
        Me.Label10.TabIndex = 27
        Me.Label10.Text = "Font:"
        '
        'chkDrawBackImage
        '
        Me.chkDrawBackImage.Location = New System.Drawing.Point(8, 408)
        Me.chkDrawBackImage.Name = "chkDrawBackImage"
        Me.chkDrawBackImage.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.chkDrawBackImage.Size = New System.Drawing.Size(176, 24)
        Me.chkDrawBackImage.TabIndex = 29
        Me.chkDrawBackImage.Text = "Draw Back Image"
        '
        'btnBackClrEnabled
        '
        Me.btnBackClrEnabled.Location = New System.Drawing.Point(64, 236)
        Me.btnBackClrEnabled.Name = "btnBackClrEnabled"
        Me.btnBackClrEnabled.Size = New System.Drawing.Size(120, 23)
        Me.btnBackClrEnabled.TabIndex = 31
        '
        'Label11
        '
        Me.Label11.Location = New System.Drawing.Point(8, 236)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(48, 24)
        Me.Label11.TabIndex = 30
        Me.Label11.Text = "Back (E)"
        '
        'btnBackClrDisabled
        '
        Me.btnBackClrDisabled.Location = New System.Drawing.Point(64, 264)
        Me.btnBackClrDisabled.Name = "btnBackClrDisabled"
        Me.btnBackClrDisabled.Size = New System.Drawing.Size(120, 23)
        Me.btnBackClrDisabled.TabIndex = 33
        '
        'Label12
        '
        Me.Label12.Location = New System.Drawing.Point(8, 264)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(48, 24)
        Me.Label12.TabIndex = 32
        Me.Label12.Text = "Back (D)"
        '
        'Label13
        '
        Me.Label13.Location = New System.Drawing.Point(8, 432)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(72, 24)
        Me.Label13.TabIndex = 35
        Me.Label13.Text = "Max Length:"
        '
        'txtMaxLen
        '
        Me.txtMaxLen.Location = New System.Drawing.Point(88, 432)
        Me.txtMaxLen.Name = "txtMaxLen"
        Me.txtMaxLen.Size = New System.Drawing.Size(96, 20)
        Me.txtMaxLen.TabIndex = 34
        '
        'Label14
        '
        Me.Label14.Location = New System.Drawing.Point(8, 456)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(48, 24)
        Me.Label14.TabIndex = 37
        Me.Label14.Text = "Value:"
        '
        'txtValue
        '
        Me.txtValue.Location = New System.Drawing.Point(64, 456)
        Me.txtValue.Name = "txtValue"
        Me.txtValue.Size = New System.Drawing.Size(120, 20)
        Me.txtValue.TabIndex = 36
        '
        'Label15
        '
        Me.Label15.Location = New System.Drawing.Point(8, 480)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(72, 24)
        Me.Label15.TabIndex = 39
        Me.Label15.Text = "Max Value:"
        '
        'txtMaxValue
        '
        Me.txtMaxValue.Location = New System.Drawing.Point(88, 480)
        Me.txtMaxValue.Name = "txtMaxValue"
        Me.txtMaxValue.Size = New System.Drawing.Size(96, 20)
        Me.txtMaxValue.TabIndex = 38
        '
        'Label16
        '
        Me.Label16.Location = New System.Drawing.Point(8, 504)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(72, 24)
        Me.Label16.TabIndex = 41
        Me.Label16.Text = "Min Value:"
        '
        'txtMinValue
        '
        Me.txtMinValue.Location = New System.Drawing.Point(88, 504)
        Me.txtMinValue.Name = "txtMinValue"
        Me.txtMinValue.Size = New System.Drawing.Size(96, 20)
        Me.txtMinValue.TabIndex = 40
        '
        'Label17
        '
        Me.Label17.Location = New System.Drawing.Point(8, 528)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(72, 24)
        Me.Label17.TabIndex = 43
        Me.Label17.Text = "Small Chng:"
        '
        'txtSmallChng
        '
        Me.txtSmallChng.Location = New System.Drawing.Point(88, 528)
        Me.txtSmallChng.Name = "txtSmallChng"
        Me.txtSmallChng.Size = New System.Drawing.Size(96, 20)
        Me.txtSmallChng.TabIndex = 42
        '
        'Label18
        '
        Me.Label18.Location = New System.Drawing.Point(8, 552)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(72, 24)
        Me.Label18.TabIndex = 45
        Me.Label18.Text = "Large Chng:"
        '
        'txtLargeChng
        '
        Me.txtLargeChng.Location = New System.Drawing.Point(88, 552)
        Me.txtLargeChng.Name = "txtLargeChng"
        Me.txtLargeChng.Size = New System.Drawing.Size(96, 20)
        Me.txtLargeChng.TabIndex = 44
        '
        'chkReverseDirection
        '
        Me.chkReverseDirection.Location = New System.Drawing.Point(8, 576)
        Me.chkReverseDirection.Name = "chkReverseDirection"
        Me.chkReverseDirection.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.chkReverseDirection.Size = New System.Drawing.Size(176, 24)
        Me.chkReverseDirection.TabIndex = 46
        Me.chkReverseDirection.Text = "Reverse Direction"
        '
        'btnHighlight
        '
        Me.btnHighlight.Location = New System.Drawing.Point(65, 293)
        Me.btnHighlight.Name = "btnHighlight"
        Me.btnHighlight.Size = New System.Drawing.Size(120, 23)
        Me.btnHighlight.TabIndex = 48
        '
        'Label19
        '
        Me.Label19.Location = New System.Drawing.Point(9, 293)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(48, 24)
        Me.Label19.TabIndex = 47
        Me.Label19.Text = "Highlight"
        '
        'btnDropDownBorder
        '
        Me.btnDropDownBorder.Location = New System.Drawing.Point(64, 328)
        Me.btnDropDownBorder.Name = "btnDropDownBorder"
        Me.btnDropDownBorder.Size = New System.Drawing.Size(120, 23)
        Me.btnDropDownBorder.TabIndex = 50
        '
        'Label20
        '
        Me.Label20.Location = New System.Drawing.Point(8, 328)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(64, 24)
        Me.Label20.TabIndex = 49
        Me.Label20.Text = "Drop Down Border"
        '
        'btnAlign
        '
        Me.btnAlign.Location = New System.Drawing.Point(136, 384)
        Me.btnAlign.Name = "btnAlign"
        Me.btnAlign.Size = New System.Drawing.Size(48, 23)
        Me.btnAlign.TabIndex = 51
        Me.btnAlign.Text = "Align"
        '
        'btnSetRect
        '
        Me.btnSetRect.Location = New System.Drawing.Point(8, 408)
        Me.btnSetRect.Name = "btnSetRect"
        Me.btnSetRect.Size = New System.Drawing.Size(61, 23)
        Me.btnSetRect.TabIndex = 52
        Me.btnSetRect.Text = "Set Rect"
        '
        'frmProps
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(194, 607)
        Me.Controls.Add(Me.btnSetRect)
        Me.Controls.Add(Me.btnAlign)
        Me.Controls.Add(Me.btnDropDownBorder)
        Me.Controls.Add(Me.Label20)
        Me.Controls.Add(Me.btnHighlight)
        Me.Controls.Add(Me.Label19)
        Me.Controls.Add(Me.chkReverseDirection)
        Me.Controls.Add(Me.Label18)
        Me.Controls.Add(Me.txtLargeChng)
        Me.Controls.Add(Me.Label17)
        Me.Controls.Add(Me.txtSmallChng)
        Me.Controls.Add(Me.Label16)
        Me.Controls.Add(Me.txtMinValue)
        Me.Controls.Add(Me.Label15)
        Me.Controls.Add(Me.txtMaxValue)
        Me.Controls.Add(Me.Label14)
        Me.Controls.Add(Me.txtValue)
        Me.Controls.Add(Me.Label13)
        Me.Controls.Add(Me.txtMaxLen)
        Me.Controls.Add(Me.btnBackClrDisabled)
        Me.Controls.Add(Me.Label12)
        Me.Controls.Add(Me.btnBackClrEnabled)
        Me.Controls.Add(Me.Label11)
        Me.Controls.Add(Me.chkDrawBackImage)
        Me.Controls.Add(Me.btnFont)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.btnForeClr)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.txtCaption)
        Me.Controls.Add(Me.btnFillClr)
        Me.Controls.Add(Me.btnBorderClr)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.chkEnabled)
        Me.Controls.Add(Me.chkVisible)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.txtHeight)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.txtWidth)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.txtTop)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.txtLeft)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.txtName)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.Name = "frmProps"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Properties"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region

    Public Event UpdateName(ByVal sName As String)
    Public Event UpdateEnabled(ByVal bEnabled As Boolean)
    Public Event UpdateHeight(ByVal lHeight As Int32)
    Public Event UpdateLeft(ByVal lLeft As Int32)
    Public Event UpdateTop(ByVal lTop As Int32)
    Public Event UpdateVisible(ByVal bVisible As Boolean)
    Public Event UpdateWidth(ByVal lWidth As Int32)
    Public Event UpdateBorderColor(ByVal clrNew As System.Drawing.Color)
    Public Event UpdateFillColor(ByVal clrNew As System.Drawing.Color)
    Public Event UpdateForeColor(ByVal clrNew As System.Drawing.Color)
    Public Event UpdateBackColorEnabled(ByVal clrNew As System.Drawing.Color)
    Public Event UpdateBackColorDisabled(ByVal cldNew As System.Drawing.Color)
    Public Event UpdateCaption(ByVal sCaption As String)
    Public Event UpdateFont(ByVal oFont As System.Drawing.Font)
    Public Event UpdateDrawBackImage(ByVal bNewVal As Boolean)
    Public Event UpdateMaxLength(ByVal lMaxLength As Int32)
    Public Event UpdateValue(ByVal lValue As Int32)
    Public Event UpdateMaxValue(ByVal lMaxValue As Int32)
    Public Event UpdateMinValue(ByVal lMinValue As Int32)
    Public Event UpdateSmallChange(ByVal lSmallChange As Int32)
    Public Event UpdateLargeChange(ByVal lLargeChange As Int32)
    Public Event UpdateReverseDirection(ByVal bNewVal As Boolean)
    Public Event UpdateHighlightColor(ByVal clrNew As System.Drawing.Color)
    Public Event UpdateDropDownBorderColor(ByVal clrNew As System.Drawing.Color)
    Public Event UpdateFontFormat(ByVal oNewFmt As Microsoft.DirectX.Direct3D.DrawTextFormat)
    Public Event UpdateLocRect(ByVal lX As Int32, ByVal lY As Int32, ByVal lWidth As Int32, ByVal lHeight As Int32)

    Private mbNoRaiseEvents As Boolean = False

    Private moFontFormat As Microsoft.DirectX.Direct3D.DrawTextFormat

    Private mrectDraw As Rectangle

    Public Sub ConfigProps(ByVal oControl As Object)

        mbNoRaiseEvents = True

        With CType(oControl, UIControl)
            txtName.Text = .ControlName
            chkEnabled.Checked = .Enabled
            txtHeight.Text = .Height.ToString
            txtLeft.Text = .Left.ToString
            txtTop.Text = .Top.ToString
            chkVisible.Checked = .Visible
            txtWidth.Text = .Width.ToString

            mrectDraw = .ControlImageRect
        End With

        DisableControls()

        Select Case oControl.GetType.ToString
            Case "InterfaceBuilder.UILine"
                btnBorderClr.Visible = True
                btnBorderClr.BackColor = CType(oControl, UILine).BorderColor
            Case "InterfaceBuilder.UIWindow"
                btnBorderClr.Visible = True
                btnBorderClr.BackColor = CType(oControl, UIWindow).BorderColor
                btnFillClr.Visible = True
                btnFillClr.BackColor = CType(oControl, UIWindow).FillColor

                txtCaption.Enabled = True
                txtCaption.Text = CType(oControl, UIWindow).Caption
            Case "InterfaceBuilder.UILabel", "InterfaceBuilder.UIButton", "InterfaceBuilder.UITextBox", "InterfaceBuilder.UIOption", "InterfaceBuilder.UICheckBox"
                With CType(oControl, UILabel)
                    txtCaption.Enabled = True
                    txtCaption.Text = .Caption
                    btnForeClr.Visible = True
                    btnForeClr.BackColor = .ForeColor
                    btnFont.Visible = True
                    btnAlign.Visible = True
                    SetBtnAlign(.FontFormat)
                    moFontFormat = .FontFormat
                    btnFont.Font = .GetFont()
                    chkDrawBackImage.Checked = .DrawBackImage
                    chkDrawBackImage.Enabled = True
                End With

                If oControl.GetType.ToString = "InterfaceBuilder.UITextBox" Then
                    With CType(oControl, UITextBox)
                        btnBackClrEnabled.Visible = True
                        btnBackClrEnabled.BackColor = .BackColorEnabled
                        btnBackClrDisabled.Visible = True
                        btnBackClrDisabled.BackColor = .BackColorDisabled
                        txtMaxLen.Enabled = True
                        txtMaxLen.Text = .MaxLength.ToString
                        btnBorderClr.Visible = True
                        btnBorderClr.BackColor = .BorderColor
                    End With
                ElseIf oControl.GetType.ToString = "InterfaceBuilder.UIOption" Or oControl.GetType.ToString = "InterfaceBuilder.UICheckBox" Then
                    txtValue.Text = CStr(CallByName(oControl, "Value", CallType.Get))
                    txtValue.Enabled = True
                End If
            Case "InterfaceBuilder.UIScrollBar"
                With CType(oControl, UIScrollBar)
                    txtValue.Enabled = True
                    txtValue.Text = .Value.ToString
                    txtMaxValue.Enabled = True
                    txtMaxValue.Text = .MaxValue.ToString
                    txtMinValue.Enabled = True
                    txtMinValue.Text = .MinValue.ToString
                    txtSmallChng.Enabled = True
                    txtSmallChng.Text = .SmallChange.ToString
                    txtLargeChng.Enabled = True
                    txtLargeChng.Text = .LargeChange.ToString
                    chkReverseDirection.Enabled = True
                    chkReverseDirection.Checked = .ReverseDirection
                End With
            Case "InterfaceBuilder.UIListBox"
                With CType(oControl, UIListBox)
                    btnBorderClr.Visible = True
                    btnBorderClr.BackColor = .BorderColor
                    btnFillClr.Visible = True
                    btnFillClr.BackColor = .FillColor
                    btnForeClr.Visible = True
                    btnForeClr.BackColor = .ForeColor
                    btnFont.Visible = True
                    btnFont.Font = .GetFont
                    btnHighlight.Visible = True
                    btnHighlight.BackColor = .HighlightColor
                End With
            Case "InterfaceBuilder.UIComboBox"
                With CType(oControl, UIComboBox)
                    btnBorderClr.Visible = True
                    btnBorderClr.BackColor = .BorderColor
                    btnFillClr.Visible = True
                    btnFillClr.BackColor = .FillColor
                    btnForeClr.Visible = True
                    btnForeClr.BackColor = .ForeColor
                    btnFont.Visible = True
                    btnFont.Font = .GetFont
                    btnHighlight.Visible = True
                    btnHighlight.BackColor = .HighlightColor
                    btnDropDownBorder.Visible = True
                    btnDropDownBorder.BackColor = .DropDownListBorderColor
                End With
        End Select

        mbNoRaiseEvents = False
    End Sub

    Private Sub DisableControls()
        btnBorderClr.Visible = False
        btnFillClr.Visible = False
        btnForeClr.Visible = False
        btnBackClrEnabled.Visible = False
        btnBackClrDisabled.Visible = False
        btnHighlight.Visible = False
        btnDropDownBorder.Visible = False
        txtCaption.Enabled = False
        btnFont.Visible = False
        btnAlign.Visible = False
        chkDrawBackImage.Enabled = False
        txtMaxLen.Enabled = False
        txtValue.Enabled = False
        txtMaxValue.Enabled = False
        txtMinValue.Enabled = False
        txtSmallChng.Enabled = False
        txtLargeChng.Enabled = False
        chkReverseDirection.Enabled = False
    End Sub

    Private Sub txtName_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtName.TextChanged
        If mbNoRaiseEvents = False Then RaiseEvent UpdateName(txtName.Text)
    End Sub

    Private Sub txtLeft_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtLeft.TextChanged
        If mbNoRaiseEvents = False Then RaiseEvent UpdateLeft(CInt(Val(txtLeft.Text)))
    End Sub

    Private Sub txtTop_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtTop.TextChanged
        If mbNoRaiseEvents = False Then RaiseEvent UpdateTop(CInt(Val(txtTop.Text)))
    End Sub

    Private Sub txtWidth_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtWidth.TextChanged
        If mbNoRaiseEvents = False Then RaiseEvent UpdateWidth(CInt(Val(txtWidth.Text)))
    End Sub

    Private Sub txtHeight_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtHeight.TextChanged
        If mbNoRaiseEvents = False Then RaiseEvent UpdateHeight(CInt(Val(txtHeight.Text)))
    End Sub

    Private Sub chkVisible_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkVisible.CheckedChanged
        If mbNoRaiseEvents = False Then RaiseEvent UpdateVisible(chkVisible.Checked)
    End Sub

    Private Sub chkEnabled_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkEnabled.CheckedChanged
        If mbNoRaiseEvents = False Then RaiseEvent UpdateEnabled(chkEnabled.Checked)
    End Sub

    Private Sub btnBorderClr_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBorderClr.Click
        Dim clrRes As System.Drawing.Color
        Dim dlgRes As DialogResult

        If mbNoRaiseEvents = True Then Exit Sub

        clrdlg1.Color = btnBorderClr.BackColor

        dlgRes = clrdlg1.ShowDialog(Me)
        If dlgRes = Windows.Forms.DialogResult.OK Then
            clrRes = clrdlg1.Color
            btnBorderClr.BackColor = clrRes
            RaiseEvent UpdateBorderColor(clrRes)
        End If
    End Sub

    Private Sub btnFillClr_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFillClr.Click
        Dim clrRes As System.Drawing.Color
        Dim dlgRes As DialogResult

        If mbNoRaiseEvents = True Then Exit Sub

        clrdlg1.Color = btnFillClr.BackColor

        dlgRes = clrdlg1.ShowDialog(Me)
        If dlgRes = Windows.Forms.DialogResult.OK Then
            clrRes = clrdlg1.Color
            btnFillClr.BackColor = clrRes
            RaiseEvent UpdateFillColor(clrRes)
        End If
    End Sub

    Private Sub btnForeClr_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnForeClr.Click
        Dim clrRes As System.Drawing.Color
        Dim dlgRes As DialogResult

        If mbNoRaiseEvents = True Then Exit Sub

        clrdlg1.Color = btnForeClr.BackColor

        dlgRes = clrdlg1.ShowDialog(Me)
        If dlgRes = Windows.Forms.DialogResult.OK Then
            clrRes = clrdlg1.Color
            btnForeClr.BackColor = clrRes
            RaiseEvent UpdateForeColor(clrRes)
        End If
    End Sub

    Private Sub btnBackClrEnabled_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBackClrEnabled.Click
        Dim clrRes As System.Drawing.Color
        Dim dlgRes As DialogResult

        If mbNoRaiseEvents = True Then Exit Sub

        clrdlg1.Color = btnBackClrEnabled.BackColor

        dlgRes = clrdlg1.ShowDialog(Me)
        If dlgRes = Windows.Forms.DialogResult.OK Then
            clrRes = clrdlg1.Color
            btnBackClrEnabled.BackColor = clrRes
            RaiseEvent UpdateBackColorEnabled(clrRes)
        End If
    End Sub

    Private Sub btnBackClrDisabled_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBackClrDisabled.Click
        Dim clrRes As System.Drawing.Color
        Dim dlgRes As DialogResult

        If mbNoRaiseEvents = True Then Exit Sub

        clrdlg1.Color = btnBackClrDisabled.BackColor

        dlgRes = clrdlg1.ShowDialog(Me)
        If dlgRes = Windows.Forms.DialogResult.OK Then
            clrRes = clrdlg1.Color
            btnBackClrDisabled.BackColor = clrRes
            RaiseEvent UpdateBackColorDisabled(clrRes)
        End If
    End Sub

    Private Sub txtCaption_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCaption.TextChanged
        If mbNoRaiseEvents = False Then RaiseEvent UpdateCaption(txtCaption.Text)
    End Sub

    Private Sub btnFont_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFont.Click
        Dim oRes As System.Drawing.Font
        Dim dlgRes As DialogResult

        If mbNoRaiseEvents = True Then Exit Sub

        FontDlg.Font = btnFont.Font
        dlgRes = FontDlg.ShowDialog(Me)
        If dlgRes = Windows.Forms.DialogResult.OK Then
            oRes = FontDlg.Font
            btnFont.Font = oRes
            RaiseEvent UpdateFont(oRes)
        End If
    End Sub

    Private Sub chkDrawBackImage_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkDrawBackImage.CheckedChanged
        If mbNoRaiseEvents = False Then RaiseEvent UpdateDrawBackImage(chkDrawBackImage.Checked)
    End Sub

    Private Sub txtMaxLen_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtMaxLen.TextChanged
        If mbNoRaiseEvents = False Then RaiseEvent UpdateMaxLength(CInt(Val(txtMaxLen.Text)))
    End Sub

    Private Sub txtValue_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtValue.TextChanged
        If mbNoRaiseEvents = False Then RaiseEvent UpdateValue(CInt(Val(txtValue.Text)))
    End Sub

    Private Sub txtMaxValue_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtMaxValue.TextChanged
        If mbNoRaiseEvents = False Then RaiseEvent UpdateMaxValue(CInt(Val(txtMaxValue.Text)))
    End Sub

    Private Sub txtMinValue_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtMinValue.TextChanged
        If mbNoRaiseEvents = False Then RaiseEvent UpdateMinValue(CInt(Val(txtMinValue.Text)))
    End Sub

    Private Sub txtSmallChng_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSmallChng.TextChanged
        If mbNoRaiseEvents = False Then RaiseEvent UpdateSmallChange(CInt(Val(txtSmallChng.Text)))
    End Sub

    Private Sub txtLargeChng_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtLargeChng.TextChanged
        If mbNoRaiseEvents = False Then RaiseEvent UpdateLargeChange(CInt(Val(txtLargeChng.Text)))
    End Sub

    Private Sub chkReverseDirection_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkReverseDirection.CheckedChanged
        If mbNoRaiseEvents = False Then RaiseEvent UpdateReverseDirection(chkReverseDirection.Checked)
    End Sub

    Private Sub btnHighlight_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnHighlight.Click
        Dim clrRes As System.Drawing.Color
        Dim dlgRes As DialogResult

        If mbNoRaiseEvents = True Then Exit Sub

        clrdlg1.Color = btnHighlight.BackColor

        dlgRes = clrdlg1.ShowDialog(Me)
        If dlgRes = Windows.Forms.DialogResult.OK Then
            clrRes = clrdlg1.Color
            btnHighlight.BackColor = clrRes
            RaiseEvent UpdateHighlightColor(clrRes)
        End If
    End Sub

    Private Sub btnDropDownBorder_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDropDownBorder.Click
        Dim clrRes As System.Drawing.Color
        Dim dlgRes As DialogResult

        If mbNoRaiseEvents = True Then Exit Sub

        clrdlg1.Color = btnDropDownBorder.BackColor

        dlgRes = clrdlg1.ShowDialog(Me)
        If dlgRes = Windows.Forms.DialogResult.OK Then
            clrRes = clrdlg1.Color
            btnDropDownBorder.BackColor = clrRes
            RaiseEvent UpdateDropDownBorderColor(clrRes)
        End If
    End Sub

    Private Sub btnAlign_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAlign.Click
        Dim lVal As Microsoft.DirectX.Direct3D.DrawTextFormat

        Dim yHorizontal As Byte         '0 = left, 1 = middle, 2 = right
        Dim yVertical As Byte           '0 = top, 1 = middle, 2 = right

        'ok, gotta do this the hard way
        lVal = moFontFormat

        If (lVal And Microsoft.DirectX.Direct3D.DrawTextFormat.Bottom) <> 0 Then
            yVertical = 2
        ElseIf (lVal And Microsoft.DirectX.Direct3D.DrawTextFormat.VerticalCenter) <> 0 Then
            yVertical = 1
        Else
            yVertical = 0
        End If

        If (lVal And Microsoft.DirectX.Direct3D.DrawTextFormat.Right) <> 0 Then
            yHorizontal = 2
        ElseIf (lVal And Microsoft.DirectX.Direct3D.DrawTextFormat.Center) <> 0 Then
            yHorizontal = 1
        Else
            yHorizontal = 0
        End If

        'now, increment them
        yHorizontal += CByte(1)
        If yHorizontal = 3 Then
            yHorizontal = 0
            yVertical += CByte(1)
            If yVertical = 3 Then yVertical = 0
        End If

        'Now, determine our values, set horizontal first
        Select Case yHorizontal
            Case 0 : lVal = Microsoft.DirectX.Direct3D.DrawTextFormat.Left
            Case 2 : lVal = Microsoft.DirectX.Direct3D.DrawTextFormat.Right
            Case Else : lVal = Microsoft.DirectX.Direct3D.DrawTextFormat.Center
        End Select

        Select Case yVertical
            Case 0 : lVal = lVal Or Microsoft.DirectX.Direct3D.DrawTextFormat.Top
            Case 2 : lVal = lVal Or Microsoft.DirectX.Direct3D.DrawTextFormat.Bottom
            Case Else : lVal = lVal Or Microsoft.DirectX.Direct3D.DrawTextFormat.VerticalCenter
        End Select

        moFontFormat = lVal
        RaiseEvent UpdateFontFormat(moFontFormat)
        SetBtnAlign(lVal)
    End Sub

    Private Sub SetBtnAlign(ByVal lVal As Microsoft.DirectX.Direct3D.DrawTextFormat)
        Dim yHorizontal As Byte
        Dim yVertical As Byte

        If (lVal And Microsoft.DirectX.Direct3D.DrawTextFormat.Bottom) <> 0 Then
            yVertical = 2
        ElseIf (lVal And Microsoft.DirectX.Direct3D.DrawTextFormat.VerticalCenter) <> 0 Then
            yVertical = 1
        Else
            yVertical = 0
        End If

        If (lVal And Microsoft.DirectX.Direct3D.DrawTextFormat.Right) <> 0 Then
            yHorizontal = 2
        ElseIf (lVal And Microsoft.DirectX.Direct3D.DrawTextFormat.Center) <> 0 Then
            yHorizontal = 1
        Else
            yHorizontal = 0
        End If

        Select Case yHorizontal
            Case 0
                'left
                Select Case yVertical
                    Case 0 : btnAlign.TextAlign = ContentAlignment.TopLeft
                    Case 1 : btnAlign.TextAlign = ContentAlignment.MiddleLeft
                    Case 2 : btnAlign.TextAlign = ContentAlignment.BottomLeft
                End Select
            Case 1
                'middle
                Select Case yVertical
                    Case 0 : btnAlign.TextAlign = ContentAlignment.TopCenter
                    Case 1 : btnAlign.TextAlign = ContentAlignment.MiddleCenter
                    Case 2 : btnAlign.TextAlign = ContentAlignment.BottomCenter
                End Select
            Case 2
                'right
                Select Case yVertical
                    Case 0 : btnAlign.TextAlign = ContentAlignment.TopRight
                    Case 1 : btnAlign.TextAlign = ContentAlignment.MiddleRight
                    Case 2 : btnAlign.TextAlign = ContentAlignment.BottomRight
                End Select
        End Select

    End Sub

    Private Sub btnSetRect_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSetRect.Click
        Dim frmTempSetRect As frmSetRect = New frmSetRect()
        frmTempSetRect.Show()
        frmTempSetRect.SetByRect(mrectDraw)
        AddHandler frmTempSetRect.SetRectLoc, AddressOf frmSetRect_SetRect
    End Sub

    Private Sub frmSetRect_SetRect(ByVal lX As Int32, ByVal lY As Int32, ByVal lWidth As Int32, ByVal lHeight As Int32)
        RaiseEvent UpdateLocRect(lX, lY, lWidth, lHeight)
    End Sub
End Class
