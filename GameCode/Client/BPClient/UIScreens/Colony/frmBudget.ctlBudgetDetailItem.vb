Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

Partial Class frmBudget
    'Interface created from Interface Builder
    Public Class ctlBudgetDetailItem
        Inherits UIWindow

        Public Const ml_ITEM_HEIGHT As Int32 = 24 '34

        Private lblType As UILabel
        Private lblAlert As UILabel

        Private lblEnvir As UILabel
        Private lblColony As UILabel
        Private lblRevenue As UILabel
        Private lblExpense As UILabel
        Private lblControl As UILabel
        Private lblColonyCnt As UILabel

        Private WithEvents txtTaxRate As UITextBox
        Private WithEvents btnSetTaxRate As UIButton
        Private WithEvents btnExpand As UIButton

        Private mlEnvirID As Int32
        Private miEnvirTypeID As Int16
        Private mlColonyID As Int32 = -1
        Private mbMouseDown As Boolean = False

        Private mlLastClickEvent As Int32

        Private mbEditingTaxRate As Boolean = False
        Private mbInSetValues As Boolean = True

        Public Event DetailItemClick(ByVal lEnvirID As Int32, ByVal iEnvirTypeID As Int16)
        Public Event DetailItemDoubleClick(ByVal lEnvirID As Int32, ByVal iEnvirTypeID As Int16)
        Public Event AlertButtonClick(ByVal lEnvirID As Int32, ByVal iEnvirTypeID As Int16)

        Public Event ExpandChanged(ByVal lEnvirID As Int32, ByVal iEnvirTypeID As Int16, ByVal bNewValue As Boolean)

        Private mbFlashState As Boolean = False
        Private mbWarInProgress As Boolean = False

        Private mbExpanded As Boolean = False
        Public Property bExpanded() As Boolean
            Get
                Return mbExpanded
            End Get
            Set(ByVal value As Boolean)
                mbExpanded = value
                If btnExpand Is Nothing = False Then
                    With btnExpand
                        If mbExpanded = True Then
                            .ControlImageRect_Disabled = grc_UI(elInterfaceRectangle.eDownArrow_Button_Disabled)
                            .ControlImageRect_Normal = grc_UI(elInterfaceRectangle.eDownArrow_Button_Normal)
                            .ControlImageRect_Pressed = grc_UI(elInterfaceRectangle.eDownArrow_Button_Down)
                        Else
                            .ControlImageRect_Disabled = grc_UI(elInterfaceRectangle.eRightArrow_Button_Disabled)
                            .ControlImageRect_Normal = grc_UI(elInterfaceRectangle.eRightArrow_Button_Normal)
                            .ControlImageRect_Pressed = grc_UI(elInterfaceRectangle.eRightArrow_Button_Down)
                        End If
                        .ControlImageRect = .ControlImageRect_Normal
                    End With
                End If
            End Set
        End Property

        Public Sub New(ByRef oUILib As UILib)
            MyBase.New(oUILib)

            'ctlBudgetDetailItem initial props
            With Me
                .ControlName = "ctlBudgetDetailItem"
                .Left = 70
                .Top = 183
                .Width = 765
                .Height = ml_ITEM_HEIGHT
                .Enabled = True
                .Visible = True
                .BorderColor = muSettings.InterfaceBorderColor
                .FillColor = muSettings.InterfaceFillColor
                .FullScreen = False
                .BorderLineWidth = 1
                .Moveable = False
            End With

            'lblType initial props
            lblType = New UILabel(oUILib)
            With lblType
                .ControlName = "lblType"
                .Left = 1
                .Top = 1
                .Height = Me.Height - 2
                .Width = .Height
                .Enabled = True
                .Visible = True
                .Caption = ""
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 11.0F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = True
                .FontFormat = CType(5, DrawTextFormat)
                .ControlImageRect = grc_UI(elInterfaceRectangle.ePlanetIcon)
            End With
            Me.AddChild(CType(lblType, UIControl))

            'lblEnvir initial props
            lblEnvir = New UILabel(oUILib)
            With lblEnvir
                .ControlName = "lblEnvir"
                .Left = 45
                .Top = 0 '8
                .Width = 150
                .Height = ml_item_height'18
                .Enabled = True
                .Visible = True
                .Caption = ""
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat) Or DrawTextFormat.VerticalCenter
            End With
            Me.AddChild(CType(lblEnvir, UIControl))

            'lblColony initial props
            lblColony = New UILabel(oUILib)
            With lblColony
                .ControlName = "lblColony"
                .Left = 200
                .Top = 0
                .Width = 150
                .Height = ml_ITEM_HEIGHT '18
                .Enabled = True
                .Visible = True
                .Caption = ""
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat) Or DrawTextFormat.VerticalCenter
            End With
            Me.AddChild(CType(lblColony, UIControl))

            'lblColonyCnt initial props
            lblColonyCnt = New UILabel(oUILib)
            With lblColonyCnt
                .ControlName = "lblColonyCnt"
                .Left = 310
                .Top = 0
                .Width = 50 '150
                .Height = ml_ITEM_HEIGHT '18
                .Enabled = True
                .Visible = True
                .Caption = ""
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = DrawTextFormat.Center Or DrawTextFormat.VerticalCenter
            End With
            Me.AddChild(CType(lblColonyCnt, UIControl))

            'lblRevenue initial props
            lblRevenue = New UILabel(oUILib)
            With lblRevenue
                .ControlName = "lblRevenue"
                .Left = 350
                .Top = 0
                .Width = 135 '150
                .Height = ml_ITEM_HEIGHT '18
                .Enabled = True
                .Visible = True
                .Caption = ""
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = DrawTextFormat.Center Or DrawTextFormat.VerticalCenter
            End With
            Me.AddChild(CType(lblRevenue, UIControl))

            'lblExpense initial props
            lblExpense = New UILabel(oUILib)
            With lblExpense
                .ControlName = "lblExpense"
                .Left = 495 '510
                .Top = 0
                .Width = 135 '150
                .Height = ml_ITEM_HEIGHT '18
                .Enabled = True
                .Visible = True
                .Caption = ""
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = DrawTextFormat.Center Or DrawTextFormat.VerticalCenter
            End With
            Me.AddChild(CType(lblExpense, UIControl))

            'lblcontrol initial props
            lblControl = New UILabel(oUILib)
            With lblControl
                .ControlName = "lblControl"
                .Left = 640
                .Top = 0
                .Width = 30
                .Height = ml_ITEM_HEIGHT '18
                .Enabled = True
                .Visible = True
                .Caption = ""
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = DrawTextFormat.Center Or DrawTextFormat.VerticalCenter
            End With
            Me.AddChild(CType(lblControl, UIControl))

            'txtTaxRate initial props
            txtTaxRate = New UITextBox(oUILib)
            With txtTaxRate
                .ControlName = "txtTaxRate"
                .Left = 675
                .Top = 0
                .Width = 40
                .Height = ml_ITEM_HEIGHT '18
                .Enabled = True
                .Visible = True
                .Caption = ""
                .ForeColor = muSettings.InterfaceTextBoxForeColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat) Or DrawTextFormat.VerticalCenter
                .BackColorEnabled = muSettings.InterfaceTextBoxFillColor
                .BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
                .MaxLength = 0
                .BorderColor = muSettings.InterfaceBorderColor
            End With
            Me.AddChild(CType(txtTaxRate, UIControl))

            'btnSetTaxRate initial props
            btnSetTaxRate = New UIButton(oUILib)
            With btnSetTaxRate
                .ControlName = "btnSetTaxRate"
                .Left = 720
                .Top = 1
                .Width = 40
                .Height = 22
                .Enabled = True
                .Visible = True
                .Caption = "Set"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = True
                .FontFormat = CType(5, DrawTextFormat) Or DrawTextFormat.VerticalCenter
                .ControlImageRect = New Rectangle(0, 0, 120, 32)
            End With
            Me.AddChild(CType(btnSetTaxRate, UIControl))

            'btnExpand initial props
            btnExpand = New UIButton(oUILib)
            With btnExpand
                .ControlName = "btnExpand"
                .Left = 1
                .Top = 4
                .Width = 16
                .Height = 16 '22
                .Enabled = True
                .Visible = True
                .Caption = ""
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = True
                .FontFormat = CType(5, DrawTextFormat) Or DrawTextFormat.VerticalCenter
                .ControlImageRect_Normal = grc_UI(elInterfaceRectangle.eRightArrow_Button_Normal)
                .ControlImageRect = grc_UI(elInterfaceRectangle.eRightArrow_Button_Normal)
                '.ControlImageRect_Normal = grc_UI(elInterfaceRectangle.eDownArrow_Button_Normal)
            End With
            Me.AddChild(CType(btnExpand, UIControl))

            'lblAlert initial props
            lblAlert = New UILabel(oUILib)
            With lblAlert
                .ControlName = "lblAlert"
                .Left = 18
                .Top = 0 '5
                .Width = 26
                .Height = 24
                .Enabled = True
                .Visible = False
                .Caption = ""
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = True
                .FontFormat = CType(5, DrawTextFormat)
                .ControlImageRect = New Rectangle(126, 142, 20, 20)
                .ToolTipText = "Indicates conflict in this environment. Any units with Dock During Battle" & vbCrLf & _
                               "as an AI setting will remain docked until you click this icon to disable the alert."
                .bAcceptEvents = True
            End With
            Me.AddChild(CType(lblAlert, UIControl))
            AddHandler lblAlert.OnMouseUp, AddressOf AlertLabelClicked

            mbInSetValues = False
        End Sub

        Public Sub SetValues(ByVal lEnvirID As Int32, ByVal iEnvirTypeID As Int16, ByVal sEnvir As String, ByVal sColony As String, ByVal blRev As Int64, ByVal blExp As Int64, ByVal yTaxRate As Byte, ByVal lColonyID As Int32, ByVal bHasCC As Boolean, ByVal bHasConflict As Boolean, ByVal yGovScore As Byte, ByVal lColonyCnt As Int32, ByVal lColonyCap As Int32)
            Dim bPlanet As Boolean = iEnvirTypeID = ObjectType.ePlanet
            mbInSetValues = True

            mlEnvirID = lEnvirID
            miEnvirTypeID = iEnvirTypeID

            If miEnvirTypeID = ObjectType.eSolarSystem Then
                If Me.Left <> 1 Then Me.Left = 1
                If Me.Width <> 773 Then Me.Width = 773
                If btnExpand.Visible = False Then btnExpand.Visible = True
            Else
                If Me.Width <> 765 Then Me.Width = 765
                If Me.Left <> 9 Then Me.Left = 9
                If btnExpand.Visible = True Then btnExpand.Visible = False
            End If

            mbWarInProgress = bHasConflict

            If lblType.Visible <> bPlanet Then
                lblType.Visible = bPlanet
                If bPlanet = True Then
                    lblType.ToolTipText = "Planetside"
                Else : lblType.ToolTipText = "Space"
                End If
            End If
            If iEnvirTypeID = ObjectType.ePlanet Then
                Dim sScore As String
                If lColonyID > 0 Then
                    sScore = yGovScore.ToString & "%"
                Else : sScore = ""
                End If
                If sScore <> "" Then
                    If lblControl.Caption <> sScore Then lblControl.Caption = sScore
                    If lblControl.Visible <> True Then lblControl.Visible = True
                ElseIf lblControl.Visible <> False Then
                    lblControl.Visible = False
                End If
                Dim sColonyCnt As String = lColonyCnt.ToString & "/" & lColonyCap.ToString
                If lblColonyCnt.Caption <> sColonyCnt Then
                    Dim clrTmp As Color
                    If lColonyCnt < lColonyCap Then
                        clrTmp = muSettings.InterfaceBorderColor
                    ElseIf lColonyCnt = lColonyCap Then
                        clrTmp = Color.FromArgb(255, 255, 255, 0)
                    Else : clrTmp = Color.FromArgb(255, 255, 0, 0)
                    End If
                    If lblColonyCnt.ForeColor <> clrTmp Then lblColonyCnt.ForeColor = clrTmp
                    lblColonyCnt.Caption = sColonyCnt
                End If
                If lblColonyCnt.Visible <> True Then lblColonyCnt.Visible = True
            Else
                If lblControl.Visible <> False Then lblControl.Visible = False
                If lblColonyCnt.Visible <> False Then lblColonyCnt.Visible = False
            End If

            If lblEnvir.Caption <> sEnvir Then lblEnvir.Caption = sEnvir
            If lblColony.Caption <> sColony Then lblColony.Caption = sColony
            Dim sTemp As String = blRev.ToString("#,###")
            If lblRevenue.Caption <> sTemp Then lblRevenue.Caption = sTemp
            sTemp = blExp.ToString("#,###")
            If lblExpense.Caption <> sTemp Then lblExpense.Caption = sTemp

            If mbEditingTaxRate = False Then
                sTemp = yTaxRate & "%"
                If txtTaxRate.Caption <> sTemp Then txtTaxRate.Caption = sTemp
            End If

            mlColonyID = lColonyID
            If lColonyID < 1 Then
                If txtTaxRate.Visible = True Then txtTaxRate.Visible = False
                If btnSetTaxRate.Visible = True Then btnSetTaxRate.Visible = False
                If lblColony.Caption <> "No Colony" Then lblColony.Caption = "No Colony"
            Else
                If txtTaxRate.Visible = False Then txtTaxRate.Visible = True
                If btnSetTaxRate.Visible = False Then btnSetTaxRate.Visible = True
            End If

            If bHasCC = True Then
                If lblColony.ForeColor <> lblEnvir.ForeColor Then lblColony.ForeColor = lblEnvir.ForeColor
            ElseIf lblColony.ForeColor <> Color.FromArgb(255, 255, 0, 0) Then
                lblColony.ForeColor = Color.FromArgb(255, 255, 0, 0)
            End If

            mbInSetValues = False
        End Sub

		Private Sub btnSetTaxRate_Click(ByVal sName As String) Handles btnSetTaxRate.Click
			Dim yMsg(8) As Byte
			Dim yNewRate As Byte

			If mlColonyID = -1 Then Return

			If HasAliasedRights(AliasingRights.eAlterColonyStats) = False Then
				MyBase.moUILib.AddNotification("You lack the rights to alter tax rates.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
				If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eALWAYS_PLAY, Nothing, Nothing)
				Return
			End If

			Dim sTemp As String = Replace$(txtTaxRate.Caption, "%", "")

			If (IsNumeric(sTemp) = False) OrElse Val(sTemp) < 0 OrElse Val(sTemp) > 100 Then
				MyBase.moUILib.AddNotification("Tax Rate must be a numerical value from 0 to 100!", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
				Return
            End If

            Dim lTaxRateMax As Int32 = 80

            If Val(sTemp) >= lTaxRateMax Then
                Dim ofrm As New frmMsgBox(goUILib, "Setting the tax rate to values greater than " & lTaxRateMax & " could destroy the colony. Are you sure?", MsgBoxStyle.YesNo, "Confirm Tax Rate")
                ofrm.Visible = True
                AddHandler ofrm.DialogClosed, AddressOf MsgBoxResult
                Return
            End If

			yNewRate = CByte(Val(sTemp))

			System.BitConverter.GetBytes(GlobalMessageCode.eSetColonyTaxRate).CopyTo(yMsg, 0)
			System.BitConverter.GetBytes(mlColonyID).CopyTo(yMsg, 2)
			yMsg(6) = yNewRate
			yMsg(7) = 255
			yMsg(8) = 255

			MyBase.moUILib.SendMsgToPrimary(yMsg)
			mbEditingTaxRate = False
			Me.IsDirty = True
        End Sub

        Private Sub MsgBoxResult(ByVal lResult As MsgBoxResult)
            If lResult = Microsoft.VisualBasic.MsgBoxResult.Yes Then
                Dim yMsg(8) As Byte
                Dim yNewRate As Byte

                If mlColonyID = -1 Then Return

                If HasAliasedRights(AliasingRights.eAlterColonyStats) = False Then
                    MyBase.moUILib.AddNotification("You lack the rights to alter tax rates.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                    If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eALWAYS_PLAY, Nothing, Nothing)
                    Return
                End If

                Dim sTemp As String = Replace$(txtTaxRate.Caption, "%", "")

                If (IsNumeric(sTemp) = False) OrElse Val(sTemp) < 0 OrElse Val(sTemp) > 100 Then
                    MyBase.moUILib.AddNotification("Tax Rate must be a numerical value from 0 to 100!", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                    Return
                End If

                yNewRate = CByte(Val(sTemp))

                System.BitConverter.GetBytes(GlobalMessageCode.eSetColonyTaxRate).CopyTo(yMsg, 0)
                System.BitConverter.GetBytes(mlColonyID).CopyTo(yMsg, 2)
                yMsg(6) = yNewRate
                yMsg(7) = 255
                yMsg(8) = 255

                MyBase.moUILib.SendMsgToPrimary(yMsg)
            End If
            mbEditingTaxRate = False
            Me.IsDirty = True
        End Sub

        Private Sub ctlBudgetDetailItem_OnMouseDown(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles Me.OnMouseDown
            mbMouseDown = True
        End Sub

        Private Sub ctlBudgetDetailItem_OnMouseUp(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles Me.OnMouseUp
            If mbMouseDown = True Then
                If glCurrentCycle - mlLastClickEvent < 15 Then
                    RaiseEvent DetailItemDoubleClick(mlEnvirID, miEnvirTypeID)
                Else
                    RaiseEvent DetailItemClick(mlEnvirID, miEnvirTypeID)
                    mlLastClickEvent = glCurrentCycle
                End If
            End If
            mbMouseDown = False
        End Sub

        Private mbSelected As Boolean = False

        Public Property Selected() As Boolean
            Get
                Return mbSelected
            End Get
            Set(ByVal value As Boolean)
                If mbSelected <> value Then
                    Dim lFontClr As Color
                    If value = True Then
                        Me.FillColor = System.Drawing.Color.FromArgb(255, muSettings.InterfaceFillColor.R, muSettings.InterfaceFillColor.G, muSettings.InterfaceFillColor.B)
                        lFontClr = Color.FromArgb(255, 255 - FillColor.R, 255 - FillColor.G, 255 - FillColor.B)
                    Else
                        Me.FillColor = muSettings.InterfaceFillColor
                        lFontClr = muSettings.InterfaceBorderColor
                    End If

                    lblType.ForeColor = lFontClr
                    lblEnvir.ForeColor = lFontClr
                    lblColony.ForeColor = lFontClr
                    lblRevenue.ForeColor = lFontClr
                    lblExpense.ForeColor = lFontClr
                End If
                mbSelected = value
            End Set
        End Property

        Private Sub txtTaxRate_TextChanged() Handles txtTaxRate.TextChanged
            If mbInSetValues = True Then Return
            mbEditingTaxRate = True
        End Sub

        Public Function SetFlashState(ByVal bState As Boolean) As Boolean
            mbFlashState = bState
            If mbWarInProgress = True Then
                lblAlert.Visible = True
                If mbFlashState = True Then
                    lblAlert.BackImageColor = System.Drawing.Color.FromArgb(255, 255, 0, 0)
                Else
                    lblAlert.BackImageColor = System.Drawing.Color.FromArgb(255, 64, 0, 0)
                End If
                Me.IsDirty = True
            Else
                If lblAlert.Visible = True Then lblAlert.Visible = False
            End If
            Return mbWarInProgress
        End Function

        Private Sub AlertLabelClicked(ByVal lMouseX As Int32, ByVal lMouseY As Int32, ByVal lButtons As MouseButtons)
            If mbWarInProgress = True Then
                mbWarInProgress = False
                RaiseEvent AlertButtonClick(mlEnvirID, miEnvirTypeID)
            End If
        End Sub

        Public Sub ClearEditingTaxRate()
            mbEditingTaxRate = False
        End Sub

        Private Sub btnExpand_Click(ByVal sName As String) Handles btnExpand.Click
            bExpanded = Not bExpanded
            RaiseEvent ExpandChanged(mlEnvirID, miEnvirTypeID, mbExpanded)
        End Sub

        Public Sub SetExpandButtonVisible(ByVal bVal As Boolean)
            If btnExpand Is Nothing = False Then
                If btnExpand.Visible <> bVal Then btnExpand.Visible = bVal
            End If
        End Sub
    End Class
End Class
