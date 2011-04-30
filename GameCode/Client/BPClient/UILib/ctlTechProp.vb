Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Interface created from Interface Builder
Public Class ctlTechProp
    Inherits UIControl

    Private lblPropName As UILabel
    Private WithEvents txtValue As UITextBox

    'these are generally only used for production time and research time
    Private txtDays As UITextBox
    Private txtHours As UITextBox
    Private txtMinutes As UITextBox
    Private txtSeconds As UITextBox

    Private WithEvents lblLock As UILabel
    Private WithEvents lblPercOfPayment As UILabel

    Private WithEvents shpBack As UIWindow
    Private WithEvents shpFore As UIWindow
    Private shpNub As UIWindow

    Private mbIgnoreTextChange As Boolean = False

    Public bNoMaxValue As Boolean = False
    Public blAbsoluteMaximum As Int64 = Int64.MaxValue

    Public Event PropertyValueChanged(ByVal sPropName As String, ByRef ctl As ctlTechProp)
    Public Event LockBoxDoubleClick(ByVal sPropName As String, ByRef ctl As ctlTechProp)

    Public Property TextMaxLength() As Int32
        Get
            Return txtValue.MaxLength
        End Get
        Set(ByVal value As Int32)
            txtValue.MaxLength = value
        End Set
    End Property

    Public Sub UpdateTextDisplay()
        'Ok, for now, we will only do the time based, when we change the rof stuff, we'll need to revisit
        If mbDisplayedAsTimeIndicator = True Then
            'ok, check our divisor
            Dim blSeconds As Int64 = mblValue
            If mlValueDivisor <> 1 Then
                blSeconds = blSeconds \ CLng(mlValueDivisor)
            End If
            'now, determine our values...
            Dim blDays As Int64 = blSeconds \ 86400L
            blSeconds -= (blDays * 86400L)
            Dim blHours As Int64 = blSeconds \ 3600L
            blSeconds -= (blHours * 3600L)
            Dim blMinutes As Int64 = blSeconds \ 60L
            blSeconds -= (blMinutes * 60L)
            If txtSeconds Is Nothing = True OrElse txtSeconds.Visible = False Then
                If blSeconds > 30 Then
                    blMinutes += 1
                    If blMinutes = 60 Then
                        blMinutes = 0
                        blHours += 1
                        If blHours = 24 Then
                            blHours = 0
                            blDays += 1
                        End If
                    End If
                End If
            End If

            txtDays.Caption = blDays.ToString("###0")
            txtHours.Caption = blHours.ToString("#0")
            txtMinutes.Caption = blMinutes.ToString("#0")
            If txtSeconds Is Nothing = False Then txtSeconds.Caption = blSeconds.ToString("#0")
        ElseIf mlValueDivisor > 1 Then
            Dim decTemp As Decimal = CDec(mblValue) / CDec(mlValueDivisor)
            Dim bContainsDot As Boolean = txtValue.Caption.Contains(".")
            'txtValue.Caption = decTemp.ToString("###0.##")
            If txtValue.HasFocus = False Then txtValue.Caption = decTemp.ToString("#,##0.##") Else txtValue.Caption = decTemp.ToString("###0.##")
            If bContainsDot = True Then
                If txtValue.Caption.Contains(".") = False Then
                    txtValue.Caption &= "."
                    txtValue.CursorPos = txtValue.Caption.Length + 1
                    txtValue.SelStart = txtValue.CursorPos
                End If
            End If
            Dim clrTmp As Color = muSettings.InterfaceTextBoxForeColor
            If txtValue.ForeColor <> clrTmp Then txtValue.ForeColor = clrTmp
        Else
            If txtValue.HasFocus = False Then txtValue.Caption = mblValue.ToString("#,##0.##") Else txtValue.Caption = mblValue.ToString
            Dim clrTmp As Color = muSettings.InterfaceTextBoxForeColor
            If txtValue.ForeColor <> clrTmp Then txtValue.ForeColor = clrTmp
        End If
    End Sub

    Private mblMaxValue As Int64 = 255
    Public Property MaxValue() As Int64
        Get
            Return mblMaxValue
        End Get
        Set(ByVal value As Int64)
            mblMaxValue = value
            If Me.mblValue > mblMaxValue Then
                Me.mblValue = mblMaxValue
            End If
            Me.IsDirty = True
        End Set
    End Property
    Private mblMinValue As Int64 = 0
    Public Property MinValue() As Int64
        Get
            Return mblMinValue
        End Get
        Set(ByVal value As Int64)
            mblMinValue = value
            If Me.mblValue < mblMinValue Then Me.mblValue = mblMinValue
            Me.IsDirty = True
        End Set
    End Property
    Public Property PropertyName() As String
        Get
            Return lblPropName.Caption
        End Get
        Set(ByVal value As String)
            lblPropName.Caption = value
        End Set
    End Property
    Private mbLocked As Boolean = False
    Public Property PropertyLocked() As Boolean
        Get
            Return mbLocked
        End Get
        Set(ByVal value As Boolean)
            mbLocked = value
            Me.IsDirty = True
        End Set
    End Property
    Private mblValue As Int64 = 0
    Public Property PropertyValue() As Int64
        Get
            Return mblValue
        End Get
        Set(ByVal value As Int64)
            If value > mblMaxValue Then
                If bNoMaxValue = False Then
                    value = mblMaxValue
                Else
                    mblMaxValue = value
                End If
            End If

            If value < mblMinValue Then value = mblMinValue
            mblValue = value
            mbIgnoreTextChange = True

            UpdateTextDisplay()

            mbIgnoreTextChange = False
            RaiseEvent PropertyValueChanged(Me.PropertyName, Me)
            Me.IsDirty = True
        End Set
    End Property

    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)

        'ctlTechProp initial props
        With Me
            .ControlName = "ctlTechProp"
            .Left = 200
            .Top = 370
            .Width = 456
            .Height = 18
            .Enabled = True
            .Visible = True
        End With

        'lblPropName initial props
        lblPropName = New UILabel(oUILib)
        With lblPropName
            .ControlName = "lblPropName"
            .Left = 0
            .Top = 0
            .Width = 150
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Property Name Here:"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblPropName, UIControl))

        'lblPercOfPayment initial props
        lblPercOfPayment = New UILabel(oUILib)
        With lblPercOfPayment
            .ControlName = "lblPercOfPayment"
            .Left = 0
            .Top = 0
            .Width = 40
            .Height = 18
            .Enabled = True
            .Visible = False
            .Caption = ""
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
            .bAcceptEvents = True
        End With
        Me.AddChild(CType(lblPercOfPayment, UIControl))

        'txtValue initial props
        txtValue = New UITextBox(oUILib)
        With txtValue
            .ControlName = "txtValue"
            .Left = 359
            .Top = 0
            .Width = 90
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = ""
            .ForeColor = muSettings.InterfaceTextBoxForeColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
            .BackColorEnabled = muSettings.InterfaceTextBoxFillColor
            .BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
            .BorderColor = muSettings.InterfaceBorderColor
            .bNumericOnly = True
            .MaxLength = 15
        End With
        Me.AddChild(CType(txtValue, UIControl))

        'lblLock initial props
        lblLock = New UILabel(oUILib)
        With lblLock
            .ControlName = "lblLock"
            .Left = 432
            .Top = 2
            .Width = 16
            .Height = 16
            .Enabled = True
            .Visible = True
            .Caption = ""
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = grc_UI(elInterfaceRectangle.eLock)
            .bAcceptEvents = True
            .ToolTipText = "If this icon is red (Locked), the value will not change" & vbCrLf & _
                           "automatically but will retain the current value between" & vbCrLf & _
                           "estimate recalculation. Differences in costs of between" & vbCrLf & _
                           "the estimated value and the current value are applied to" & vbCrLf & _
                           "other properties and costs of the component."
        End With
        Me.AddChild(CType(lblLock, UIControl))

        shpBack = New UIWindow(oUILib)
        With shpBack
            .BorderColor = muSettings.InterfaceBorderColor
            .BorderLineWidth = 3
            .bRoundedBorder = False
            .ControlName = "shpBack"
            .FillColor = System.Drawing.Color.FromArgb(255, 0, 0, 0)
            .Height = 16
            .Left = 155
            .Moveable = False
            .Resizable = False
            .Top = 2
            .Visible = True
            .Enabled = True
            .Width = 200
        End With
        Me.AddChild(CType(shpBack, UIControl))

        shpFore = New UIWindow(oUILib)
        With shpFore
            .BorderColor = muSettings.InterfaceBorderColor
            .BorderLineWidth = 2
            .bRoundedBorder = False
            .DrawBorder = False
            .ControlName = "shpFore"
            .FillColor = System.Drawing.Color.FromArgb(255, 0, 0, 255)
            .Height = 13
            .Left = 157
            .Moveable = False
            .Resizable = False
            .Top = 4
            .Visible = True
            .Enabled = True
            .Width = 20
        End With
        Me.AddChild(CType(shpFore, UIControl))

        shpNub = New UIWindow(oUILib)
        With shpNub
            .BorderColor = muSettings.InterfaceBorderColor
            .BorderLineWidth = 2
            .bRoundedBorder = False
            .DrawBorder = False
            .ControlName = "shpNub"
            .FillColor = System.Drawing.Color.FromArgb(255, 0, 0, 255)
            .Height = Me.Height - 2
            .Left = 157
            .Moveable = False
            .Resizable = False
            .Top = 2
            .Visible = True
            .Enabled = True
            .Width = 5
            .bAcceptEvents = False
        End With
        Me.AddChild(CType(shpNub, UIControl))
    End Sub

    Private Sub ctlTechProp_OnGotFocus() Handles Me.OnGotFocus
        If mbDisplayedAsTimeIndicator = True Then
            If txtDays.HasFocus = False AndAlso txtHours.HasFocus = False AndAlso txtMinutes.HasFocus = False AndAlso (txtSeconds Is Nothing OrElse txtSeconds.HasFocus = False) Then
                If MyBase.moUILib.FocusedControl Is Nothing = False Then
                    MyBase.moUILib.FocusedControl.HasFocus = False
                End If
                txtDays.HasFocus = True
                MyBase.moUILib.FocusedControl = txtDays
            End If
        Else
            If MyBase.moUILib.FocusedControl Is Nothing OrElse MyBase.moUILib.FocusedControl Is txtValue = False Then
                If MyBase.moUILib.FocusedControl Is Nothing = False Then
                    MyBase.moUILib.FocusedControl.HasFocus = False
                End If
                Me.HasFocus = False
                txtValue.HasFocus = True
                MyBase.moUILib.FocusedControl = txtValue
            End If
        End If
    End Sub

    Public Sub ctlTechProp_OnMouseMove(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles Me.OnMouseMove
        If lButton = MouseButtons.Left Then
            Try
                Dim lAbsPosX As Int32 = shpBack.GetAbsolutePosition.X
                lMouseX -= lAbsPosX

                Dim decValue As Decimal = 0
                If lMouseX < 10 Then
                    decValue = 0D
                ElseIf lMouseX > shpBack.Width + 10 Then
                    decValue = 1D
                Else
                    decValue = CDec(lMouseX / shpBack.Width)
                End If
                'If lMouseX > shpBack.Width + 10 Then Return

                decValue *= (mblMaxValue - mblMinValue)
                decValue += mblMinValue

                mbLocked = True
                If decValue < Int64.MinValue Then
                    Me.PropertyValue = Int64.MinValue
                ElseIf decValue > Int64.MaxValue Then
                    Me.PropertyValue = Int64.MaxValue
                Else
                    Me.PropertyValue = CLng(decValue)
                End If
            Catch
            End Try
        End If
    End Sub

    Private Sub ctlTechProp_OnRender() Handles Me.OnRender
        If mbLocked = True Then
            If lblLock.BackImageColor <> System.Drawing.Color.FromArgb(255, 255, 0, 0) Then
                lblLock.BackImageColor = System.Drawing.Color.FromArgb(255, 255, 0, 0)
            End If
        ElseIf lblLock.BackImageColor <> System.Drawing.Color.FromArgb(128, 255, 255, 255) Then
            lblLock.BackImageColor = System.Drawing.Color.FromArgb(128, 255, 255, 255)
        End If

        Dim fWidth As Single = CSng((mblValue - mblMinValue) / Math.Max(1, (mblMaxValue - mblMinValue)))
        Dim fClrLookup As Single = fWidth

        Try
            fWidth *= (shpBack.Width - 4)
            Dim lValue As Int32 = Math.Max(1, Math.Min(shpBack.Width - 4, CInt(fWidth)))
            If shpFore.Width <> lValue Then shpFore.Width = lValue
        

            Dim lNewLeft As Int32 = shpFore.Left + shpFore.Width
            If shpNub.Left <> lNewLeft Then shpNub.Left = lNewLeft

            ' 0 to .5 = B = 1 to 0
            Dim lB As Int32 = Math.Min(255, Math.Max(0, CInt(255 * (1.0F - (fClrLookup / 0.5F)))))
            ' .25 to .5 = G = 0 to 1
            ' .5 to .75 = G = 1 to 0
            Dim lG As Int32 = 0

            If fClrLookup < 0.5F Then
                lG = Math.Min(255, Math.Max(0, CInt(255 * (fClrLookup - 0.25F) / 0.25F)))
            Else
                lG = Math.Min(255, Math.Max(0, CInt(255 * (1.0F - ((fClrLookup - 0.5F) / 0.25F)))))
            End If

            ' .5 to 1 = R = 0 to 1
            Dim lR As Int32 = Math.Min(255, Math.Max(0, CInt(255 * ((fClrLookup - 0.5F) / 0.5F))))

            shpFore.FillColor = System.Drawing.Color.FromArgb(255, lR, lG, lB)
            If shpNub.FillColor <> muSettings.InterfaceBorderColor Then shpNub.FillColor = muSettings.InterfaceBorderColor
        Catch
        End Try

    End Sub

    Private Sub lblLock_OnMouseDown(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles lblLock.OnMouseDown
        mbLocked = Not mbLocked
        Me.IsDirty = True
        ctlTechProp_OnGotFocus()
        RaiseEvent PropertyValueChanged(Me.PropertyName, Me)
    End Sub

    Private Sub txtValue_OnGotFocus() Handles txtValue.OnGotFocus
        mbIgnoreTextChange = True
        txtValue.Caption = txtValue.Caption.Replace(",", "")
        mbIgnoreTextChange = False
    End Sub

    Private Sub txtValue_OnLostFocus() Handles txtValue.OnLostFocus
        If mbIgnoreTextChange = True Then Return
        If IsNumeric(txtValue.Caption) = True OrElse txtValue.Caption = "" Then

            If mlValueDivisor > 1 Then
                mblValue = CLng(Val(txtValue.Caption) * CLng(mlValueDivisor))
            ElseIf txtValue.Caption = "" Then
                mblValue = 0
            Else
                mblValue = CLng(txtValue.Caption)
            End If

            If mblValue > mblMaxValue Then
                If bNoMaxValue = False Then
                    mblValue = mblMaxValue
                Else
                    If mblValue > blAbsoluteMaximum Then mblValue = blAbsoluteMaximum
                    mblMaxValue = mblValue
                End If
            End If

            If mblValue < mblMinValue Then mblValue = mblMinValue
            Dim clrTmp As Color = muSettings.InterfaceTextBoxForeColor
            If txtValue.ForeColor <> clrTmp Then txtValue.ForeColor = clrTmp
            'If (mlValueDivisor > 1 AndAlso mblValue.ToString("#,##0.##") <> txtValue.Caption) OrElse (mlValueDivisor < 2 AndAlso mblValue.ToString("#,##0") <> txtValue.Caption) Then
            If mblValue.ToString("#,##0.##") <> txtValue.Caption Then
                mbIgnoreTextChange = True
                If mlValueDivisor > 1 Then
                    Dim decTemp As Decimal = CDec(mblValue) / CDec(mlValueDivisor)
                    'txtValue.Caption = decTemp.ToString("###0.##")
                    If txtValue.HasFocus = False Then txtValue.Caption = decTemp.ToString("#,##0.##") Else txtValue.Caption = decTemp.ToString
                Else
                    'txtValue.Caption = mblValue.ToString
                    If txtValue.HasFocus = False Then txtValue.Caption = mblValue.ToString("#,##0") Else txtValue.Caption = mblValue.ToString
                End If
                mbIgnoreTextChange = False
                RaiseEvent PropertyValueChanged(Me.PropertyName, Me)
            End If
            'mbLocked = True

        End If
    End Sub

    Private Sub txtValue_TextChanged() Handles txtValue.TextChanged
        If mbIgnoreTextChange = True Then Return
        If IsNumeric(txtValue.Caption) = True OrElse txtValue.Caption = "" Then

            If mlValueDivisor > 1 Then
                mblValue = CLng(Val(txtValue.Caption) * CLng(mlValueDivisor))
            ElseIf txtValue.Caption = "" Then
                mblValue = 0
            Else
                mblValue = CLng(txtValue.Caption)
            End If
            Dim blTmpVal As Int64 = mblValue

            If mblValue > mblMaxValue Then
                If bNoMaxValue = False Then
                    mblValue = mblMaxValue
                Else
                    If mblValue > blAbsoluteMaximum Then mblValue = blAbsoluteMaximum
                    mblMaxValue = mblValue
                End If
            End If
            'If mblValue > mblMaxValue Then mblValue = mblMaxValue
            If mblValue < mblMinValue Then mblValue = mblMinValue

            Dim clrTmp As Color
            If blTmpVal < mblMinValue OrElse (blTmpVal > mblMaxValue AndAlso (bNoMaxValue = False OrElse blTmpVal > blAbsoluteMaximum)) Then
                clrTmp = Color.FromArgb(255, 255, 0, 0)
            Else:clrTmp = muSettings.InterfaceTextBoxForeColor
            End If
            If txtValue.ForeColor <> clrTmp Then txtValue.ForeColor = clrTmp
            If blTmpVal.ToString <> txtValue.Caption AndAlso txtValue.Caption <> "" Then
                mbIgnoreTextChange = True
                If mlValueDivisor > 1 Then
                    Dim decTemp As Decimal = CDec(mblValue) / CDec(mlValueDivisor)
                    Dim bContainsDot As Boolean = txtValue.Caption.Contains(".")
                    txtValue.Caption = decTemp.ToString("###0.##")
                    If bContainsDot = True Then
                        If txtValue.Caption.Contains(".") = False Then
                            txtValue.Caption &= "."
                            txtValue.CursorPos = txtValue.Caption.Length + 1
                            txtValue.SelStart = txtValue.CursorPos
                        End If
                    End If
                Else
                    txtValue.Caption = blTmpVal.ToString
                End If
                mbIgnoreTextChange = False
            End If
            mbLocked = True
            'If mblValue > blAbsoluteMaximum Then
            '    mblValue = blAbsoluteMaximum
            '    mbIgnoreTextChange = True
            '    UpdateTextDisplay()
            '    mbIgnoreTextChange = False
            'End If
            'If mblValue >= mblMinValue AndAlso (bNoMaxValue = True OrElse mblValue <= mblMaxValue) Then
            RaiseEvent PropertyValueChanged(Me.PropertyName, Me)
            'End If
        End If
    End Sub

    Private Sub shpBack_OnMouseDown(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles shpBack.OnMouseDown
        'we switch focus for our controls if not already at it
        ctlTechProp_OnGotFocus
        'ok, adjust x
        lMouseX -= shpBack.GetAbsolutePosition.X
        Dim decValue As Decimal = CDec(lMouseX / shpBack.Width)
        decValue *= (mblMaxValue - mblMinValue)
        decValue += mblMinValue
        mbLocked = True
        MyBase.moUILib.lUISelectState = UILib.eSelectState.eScrollBarDrag
        MyBase.moUILib.oTechProp = Me
        Me.PropertyValue = CLng(decValue)
    End Sub

    Private Sub shpBack_OnMouseMove(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles shpBack.OnMouseMove
        If lButton = MouseButtons.Left Then
            'we switch focus for our controls if not already at it
            ctlTechProp_OnGotFocus

            lMouseX -= shpBack.GetAbsolutePosition.X
            Dim decValue As Decimal = CDec(lMouseX / shpBack.Width)
            decValue *= (mblMaxValue - mblMinValue)
            decValue += mblMinValue
            mbLocked = True
            MyBase.moUILib.lUISelectState = UILib.eSelectState.eScrollBarDrag
            MyBase.moUILib.oTechProp = Me
            Me.PropertyValue = CLng(decValue)
        End If
    End Sub

    Private Sub shpFore_OnMouseDown(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles shpFore.OnMouseDown
        'we switch focus for our controls if not already at it
        ctlTechProp_OnGotFocus

        lMouseX -= shpBack.GetAbsolutePosition.X
        Dim decValue As Decimal = CDec(lMouseX / shpBack.Width)
        decValue *= (mblMaxValue - mblMinValue)
        decValue += mblMinValue
        mbLocked = True
        MyBase.moUILib.lUISelectState = UILib.eSelectState.eScrollBarDrag
        MyBase.moUILib.oTechProp = Me
        Me.PropertyValue = CLng(decValue)
    End Sub

    Private Sub shpFore_OnMouseMove(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles shpFore.OnMouseMove
        If lButton = MouseButtons.Left Then
            'we switch focus for our controls if not already at it
            ctlTechProp_OnGotFocus

            lMouseX -= shpBack.GetAbsolutePosition.X
            Dim decValue As Decimal = CDec(lMouseX / shpBack.Width)
            decValue *= (mblMaxValue - mblMinValue)
            decValue += mblMinValue
            mbLocked = True
            MyBase.moUILib.lUISelectState = UILib.eSelectState.eScrollBarDrag
            MyBase.moUILib.oTechProp = Me
            Me.PropertyValue = CLng(decValue)
        End If
    End Sub

    Public Sub SetValue(ByVal blVal As Int64, ByVal bRaiseEvent As Boolean)
        If bRaiseEvent = True Then
            PropertyValue = blVal
        Else
            mblValue = blVal
            If mblValue > mblMaxValue Then
                If bNoMaxValue = False Then
                    mblValue = mblMaxValue
                Else
                    If mblValue > blAbsoluteMaximum Then mblValue = blAbsoluteMaximum
                    mblMaxValue = mblValue
                End If
            End If
            If mblValue < mblMinValue Then mblValue = mblMinValue
            Me.IsDirty = True
        End If
    End Sub

    Private mbNoPropDisplay As Boolean = False
    Public Sub SetNoPropDisplay(ByVal bPropDisplay As Boolean)
        mbNoPropDisplay = bPropDisplay
        ctlTechProp_OnResize()
    End Sub

    Public Sub SetExtendedProps(ByVal bNoPropDisplay As Boolean, ByVal lPropWidth As Int32, ByVal lValueWidth As Int32, ByVal bShowLockSymbol As Boolean)
        mbNoPropDisplay = bNoPropDisplay
        txtValue.Width = lValueWidth
        If mbDisplayedAsTimeIndicator = True Then
            If txtSeconds Is Nothing = False AndAlso txtSeconds.Visible = True Then
                txtSeconds.Left = txtValue.Left + txtValue.Width - txtSeconds.Width
                txtMinutes.Left = txtSeconds.Left - txtMinutes.Width - 2
            Else
                txtMinutes.Left = txtValue.Left + txtValue.Width - txtMinutes.Width
            End If
            txtHours.Left = txtMinutes.Left - txtHours.Width - 2
            txtDays.Left = txtValue.Left
            txtDays.Width = txtHours.Left - txtDays.Left - 2
        End If
        lblPropName.Width = lPropWidth
        If bShowLockSymbol = False Then
            lblLock.Visible = False
        End If
        ctlTechProp_OnResize()
    End Sub

    Private mbDisplayedAsTimeIndicator As Boolean = False
    Private mlValueDivisor As Int32 = 1
    Public Sub SetAsTimeIndicator(ByVal lDivisor As Int32, ByVal bShowSeconds As Boolean)
        For X As Int32 = 0 To Me.ChildrenUB
            If Me.moChildren(X).ControlName = txtValue.ControlName Then
                Me.RemoveChild(X)
                Exit For
            End If
        Next X
        txtValue.Visible = False
        txtValue.Enabled = False

        mbDisplayedAsTimeIndicator = True
        mlValueDivisor = lDivisor

        If txtDays Is Nothing Then
            txtDays = New UITextBox(MyBase.moUILib)
            With txtDays
                .ControlName = "txtDays"
                .Left = 359
                .Top = 0
                .Width = 37
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = ""
                .ForeColor = muSettings.InterfaceTextBoxForeColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
                .BackColorEnabled = muSettings.InterfaceTextBoxFillColor
                .BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
                .BorderColor = muSettings.InterfaceBorderColor
                .bNumericOnly = True
                .ToolTipText = "Number of Days"
            End With
            Me.AddChild(CType(txtDays, UIControl))
            AddHandler txtDays.TextChanged, AddressOf TimeText_TextChanged
        End If
        If txtHours Is Nothing Then
            txtHours = New UITextBox(MyBase.moUILib)
            With txtHours
                .ControlName = "txtHours"
                .Left = 399
                .Top = 0
                .Width = 23
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = ""
                .ForeColor = muSettings.InterfaceTextBoxForeColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
                .BackColorEnabled = muSettings.InterfaceTextBoxFillColor
                .BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
                .BorderColor = muSettings.InterfaceBorderColor
                .bNumericOnly = True
                .ToolTipText = "Number of Hours in addition to the days"
            End With
            Me.AddChild(CType(txtHours, UIControl))
            AddHandler txtHours.TextChanged, AddressOf TimeText_TextChanged
        End If
        If txtMinutes Is Nothing Then
            txtMinutes = New UITextBox(MyBase.moUILib)
            With txtMinutes
                .ControlName = "txtMinutes"
                .Left = 425
                .Top = 0
                .Width = 23
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = ""
                .ForeColor = muSettings.InterfaceTextBoxForeColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
                .BackColorEnabled = muSettings.InterfaceTextBoxFillColor
                .BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
                .BorderColor = muSettings.InterfaceBorderColor
                .bNumericOnly = True
                .ToolTipText = "Number of Minutes in addition to the hours and days"
            End With
            Me.AddChild(CType(txtMinutes, UIControl))
            AddHandler txtMinutes.TextChanged, AddressOf TimeText_TextChanged
        End If
        If txtSeconds Is Nothing Then
            txtSeconds = New UITextBox(MyBase.moUILib)
            With txtSeconds
                .ControlName = "txtSeconds"
                .Left = 425
                .Top = 0
                .Width = 23
                .Height = 18
                .Enabled = True
                .Visible = False
                .Caption = ""
                .ForeColor = muSettings.InterfaceTextBoxForeColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
                .BackColorEnabled = muSettings.InterfaceTextBoxFillColor
                .BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
                .BorderColor = muSettings.InterfaceBorderColor
                .bNumericOnly = True
                .ToolTipText = "Number of Seconds in addition to the hours and days"
            End With
            Me.AddChild(CType(txtSeconds, UIControl))
            AddHandler txtSeconds.TextChanged, AddressOf TimeText_TextChanged
        End If

        txtDays.Visible = True
        txtHours.Visible = True
        txtMinutes.Visible = True
        txtSeconds.Visible = bShowSeconds

    End Sub

    'Public Sub SetAsArmorTimeIndicator(ByVal lDivisor As Int32)
    '    For X As Int32 = 0 To Me.ChildrenUB
    '        If Me.moChildren(X).ControlName = txtValue.ControlName Then
    '            Me.RemoveChild(X)
    '            Exit For
    '        End If
    '    Next X
    '    txtValue.Visible = False
    '    txtValue.Enabled = False

    '    mbDisplayedAsTimeIndicator = True
    '    mlValueDivisor = lDivisor

    '    Dim bAdded As Boolean = False
    '    If txtDays Is Nothing Then
    '        txtDays = New UITextBox(MyBase.moUILib)
    '        bAdded = True
    '    End If
    '    With txtDays
    '        .ControlName = "txtDays"
    '        .Left = 359
    '        .Top = 0
    '        .Width = 37
    '        .Height = 18
    '        .Enabled = True
    '        .Visible = True
    '        .Caption = ""
    '        .ForeColor = muSettings.InterfaceTextBoxForeColor
    '        .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
    '        .DrawBackImage = False
    '        .FontFormat = CType(4, DrawTextFormat)
    '        .BackColorEnabled = muSettings.InterfaceTextBoxFillColor
    '        .BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
    '        .BorderColor = muSettings.InterfaceBorderColor
    '        .bNumericOnly = True
    '        .ToolTipText = "Number of hours"
    '    End With
    '    If bAdded = True Then
    '        Me.AddChild(CType(txtDays, UIControl))
    '        AddHandler txtDays.TextChanged, AddressOf TimeText_TextChanged
    '    End If

    '    bAdded = False
    '    If txtHours Is Nothing Then
    '        txtHours = New UITextBox(MyBase.moUILib)
    '        bAdded = True
    '    End If

    '    With txtHours
    '        .ControlName = "txtHours"
    '        .Left = 399
    '        .Top = 0
    '        .Width = 23
    '        .Height = 18
    '        .Enabled = True
    '        .Visible = True
    '        .Caption = ""
    '        .ForeColor = muSettings.InterfaceTextBoxForeColor
    '        .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
    '        .DrawBackImage = False
    '        .FontFormat = CType(4, DrawTextFormat)
    '        .BackColorEnabled = muSettings.InterfaceTextBoxFillColor
    '        .BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
    '        .BorderColor = muSettings.InterfaceBorderColor
    '        .bNumericOnly = True
    '        .ToolTipText = "Number of minutes in addition to the hours"
    '    End With
    '    If bAdded = True Then
    '        Me.AddChild(CType(txtHours, UIControl))
    '        AddHandler txtHours.TextChanged, AddressOf TimeText_TextChanged
    '    End If

    '    bAdded = False
    '    If txtMinutes Is Nothing Then
    '        txtMinutes = New UITextBox(MyBase.moUILib)
    '        bAdded = True
    '    End If

    '    With txtMinutes
    '        .ControlName = "txtMinutes"
    '        .Left = 425
    '        .Top = 0
    '        .Width = 23
    '        .Height = 18
    '        .Enabled = True
    '        .Visible = True
    '        .Caption = ""
    '        .ForeColor = muSettings.InterfaceTextBoxForeColor
    '        .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
    '        .DrawBackImage = False
    '        .FontFormat = CType(4, DrawTextFormat)
    '        .BackColorEnabled = muSettings.InterfaceTextBoxFillColor
    '        .BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
    '        .BorderColor = muSettings.InterfaceBorderColor
    '        .bNumericOnly = True
    '        .ToolTipText = "Number of seconds in addition to the hours and minutes"
    '    End With
    '    If bAdded = True Then
    '        Me.AddChild(CType(txtMinutes, UIControl))
    '        AddHandler txtMinutes.TextChanged, AddressOf TimeText_TextChanged
    '    End If
    '    txtDays.Visible = True
    '    txtHours.Visible = True
    '    txtMinutes.Visible = True
    'End Sub

    Public Sub SetDivisor(ByVal lDivisor As Int32)
        mlValueDivisor = lDivisor
    End Sub

    Private Sub TimeText_TextChanged()
        If mbIgnoreTextChange = True Then Return
        Try
            If IsNumeric(txtDays.Caption) = True AndAlso IsNumeric(txtHours.Caption) = True AndAlso IsNumeric(txtMinutes.Caption) = True AndAlso (txtSeconds Is Nothing OrElse txtSeconds.Visible = False OrElse IsNumeric(txtSeconds.Caption)) Then

                Dim lDays As Int32 = CInt(Val(txtDays.Caption))
                Dim lHours As Int32 = CInt(Val(txtHours.Caption))
                Dim lMinutes As Int32 = CInt(Val(txtMinutes.Caption))

                'ok, determine number of seconds...
                Dim blSeconds As Int64 = (CLng(lMinutes) * 60L) + (CLng(lHours) * 3600L) + (CLng(lDays) * 86400L)

                If txtSeconds Is Nothing = False AndAlso txtSeconds.Visible = True Then
                    blSeconds += CLng(Val(txtSeconds.Caption))
                End If

                'Now, check our divisor... we use the divisor to determine number of seconds by dividing, so determine
                '  the real result, we must multiply by the divisor
                mblValue = blSeconds * CLng(mlValueDivisor)

                If mblValue > mblMaxValue Then
                    If bNoMaxValue = False Then
                        mblValue = mblMaxValue
                    Else
                        If mblValue > blAbsoluteMaximum Then mblValue = blAbsoluteMaximum
                        mblMaxValue = mblValue
                    End If
                End If
                If mblValue < mblMinValue Then mblValue = mblMinValue
                If mblValue.ToString <> txtValue.Caption Then
                    mbIgnoreTextChange = True
                    txtValue.Caption = mblValue.ToString
                    mbIgnoreTextChange = False
                End If
                mbLocked = True
                RaiseEvent PropertyValueChanged(Me.PropertyName, Me)
            End If
        Catch
        End Try
    End Sub

    Private Sub ctlTechProp_OnResize() Handles Me.OnResize
        Dim lRemainingWidth As Int32 = Me.Width
        Dim lCurrLeft As Int32 = 0
        If lblPropName Is Nothing = False Then
            lblPropName.Left = 0
            lblPropName.Top = 0
            lblPropName.Visible = Not mbNoPropDisplay
            lblPropName.Enabled = Not mbNoPropDisplay
            If mbNoPropDisplay = False Then
                lRemainingWidth -= lblPropName.Width - 5
                lCurrLeft += lblPropName.Width + 5
            End If
        End If

        'in all cases, the property caption is left-aligned to the control
        'the shape bar is left-aligned and on the right of the property caption
        'the textbox is left-aligned to the shape bar
        'the lock box is left-aligned to the textbox
        If lblLock Is Nothing = False AndAlso lblLock.Visible = True Then
            lRemainingWidth -= lblLock.Width
            lRemainingWidth -= 5
        End If
        If txtValue Is Nothing = False Then
            lRemainingWidth -= txtValue.Width
            lRemainingWidth -= 5
        End If
        If lblPercOfPayment Is Nothing = False AndAlso lblPercOfPayment.Visible = True Then
            lblPercOfPayment.Left = txtValue.Left - lblPercOfPayment.Width
            lRemainingWidth -= lblPercOfPayment.Width
        End If

        If shpFore Is Nothing = False Then
            shpFore.Left = lCurrLeft + 2
        End If
        If shpBack Is Nothing = False Then
            shpBack.Width = lRemainingWidth
            shpBack.Left = lCurrLeft

            lCurrLeft += shpBack.Width + 5
            If lblPercOfPayment Is Nothing = False AndAlso lblPercOfPayment.Visible = True Then
                lCurrLeft = lblPercOfPayment.Left + lblPercOfPayment.Width
            End If
        End If

        If txtValue Is Nothing = False Then
            txtValue.Left = lCurrLeft
            'change width?
            lCurrLeft += txtValue.Width + 5

            If txtSeconds Is Nothing = False AndAlso txtSeconds.Visible = True Then
                txtSeconds.Left = txtValue.Left + txtValue.Width - txtSeconds.Width
                If txtMinutes Is Nothing = False Then txtMinutes.Left = txtSeconds.Left - txtMinutes.Width - 2
            ElseIf txtMinutes Is Nothing = False Then
                txtMinutes.Left = txtValue.Left + txtValue.Width - txtMinutes.Width
            End If
            If txtHours Is Nothing = False Then
                If txtMinutes Is Nothing = False Then txtHours.Left = txtMinutes.Left - txtHours.Width - 2
            End If
            If txtDays Is Nothing = False Then
                txtDays.Left = txtValue.Left
                If txtHours Is Nothing = False Then txtDays.Width = txtHours.Left - txtDays.Left - 2
            End If
            'If txtDays Is Nothing = False Then
            '    txtDays.Left = txtValue.Left
            '    If txtHours Is Nothing = False Then
            '        txtHours.Left = txtDays.Left + txtDays.Width + 2
            '        If txtMinutes Is Nothing = False Then
            '            txtMinutes.Left = txtHours.Left + txtHours.Width + 2
            '        End If
            '    End If
            'End If
        End If

        If lblLock Is Nothing = False AndAlso lblLock.Visible = True Then
            lblLock.Left = lCurrLeft
        End If
    End Sub 

    Public Shadows Property ToolTipText() As String
        Get
            Return MyBase.ToolTipText
        End Get
        Set(ByVal value As String)
            MyBase.ToolTipText = value
            For X As Int32 = 0 To Me.ChildrenUB
                If moChildren(X) Is Nothing = False Then
                    moChildren(X).ToolTipText = value
                End If
            Next X
        End Set
    End Property

    Public Sub SetToInitialDefault()
        Dim blVal As Int64 = CLng(MaxValue \ 10) '   ((MaxValue - MinValue) \ 2) + MinValue
        If blVal < MinValue Then blVal = MinValue
        mblValue = blVal
        mbIgnoreTextChange = True
        UpdateTextDisplay()
        mbIgnoreTextChange = False
    End Sub

    Public Sub SetPercOfPaymentVisibility(ByVal bVisible As Boolean)
        lblPercOfPayment.Visible = bVisible
        ctlTechProp_OnResize()
    End Sub
    Public Sub SetPercOfPayment(ByVal sValue As String)
        lblPercOfPayment.Caption = sValue 'yValue.ToString & "%"
    End Sub

    Public Sub SetPercError(ByVal clrVal As System.Drawing.Color, ByVal sToolTip As String)
        lblPercOfPayment.ForeColor = clrVal

        If Me.ParentControl Is Nothing = False Then
            If Me.ParentControl.ControlName.ToUpper = "FRMTECHBUILDERCOST" Then
                If sToolTip <> "" Then sToolTip &= vbCrLf
                sToolTip &= "Double-click the percentage indicator to move the next whole percentage."
            End If
        End If
        
        lblPercOfPayment.ToolTipText = sToolTip
    End Sub

    Public Function GetPercOfPayment() As String
        Return Replace(lblPercOfPayment.Caption, "%", "")
    End Function

    Private mlLastMouseDown As Int32 = 0
    Private Sub lblPercOfPayment_OnMouseDown(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles lblPercOfPayment.OnMouseDown
        If glCurrentCycle - mlLastMouseDown < 15 Then
            'double click
            ctlTechProp_OnGotFocus()

            RaiseEvent LockBoxDoubleClick(Me.PropertyName, Me)
            RaiseEvent PropertyValueChanged(Me.PropertyName, Me)
        End If
        mlLastMouseDown = glCurrentCycle

    End Sub
End Class