Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Interface created from Interface Builder
Public Class fraBuyOrderRequires 
    Inherits UIWindow

    Private lblMinimum As UILabel
    Private lblMaximum As UILabel

    'Private lblProp() As UILabel
    Private msPropName() As String
    Private txtMin() As UITextBox
    Private txtMax() As UITextBox
	Private mlPropID() As Int32
	Private mlMinValue() As Int32
	Private mlMaxValue() As Int32

    Private myItemType As MarketListType

    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)

        'fraBuyOrderDetails initial props
        With Me
            .lWindowMetricID = BPMetrics.eWindow.eCreateBuyOrderReqs
            .ControlName = "fraBuyOrderDetails"
            .Left = 329
            .Top = 57
            .Width = 256
            .Height = 512
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = False
            .Moveable = False
            .BorderLineWidth = 2
        End With

        'lblMinimum initial props
        lblMinimum = New UILabel(oUILib)
        With lblMinimum
            .ControlName = "lblMinimum"
            .Left = 145
            .Top = 5
            .Width = 50
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Min"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(5, DrawTextFormat)
        End With
        Me.AddChild(CType(lblMinimum, UIControl))

        'lblMaximum initial props
        lblMaximum = New UILabel(oUILib)
        With lblMaximum
            .ControlName = "lblMaximum"
            .Left = 200
            .Top = 5
            .Width = 50
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Max"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(5, DrawTextFormat)
        End With
        Me.AddChild(CType(lblMaximum, UIControl))

    End Sub

    Public Sub ClearList()
        'First, remove all of the items
        Dim bDone As Boolean = False
        While bDone = False
            bDone = True
            For X As Int32 = 0 To Me.ChildrenUB
                Dim sName As String = Me.moChildren(X).ControlName.ToLower
                If sName.Contains("lblprop") OrElse sName.Contains("txtmin") OrElse sName.Contains("txtmax") Then
                    Me.RemoveChild(X)
                    bDone = False
                    Exit For
                End If
            Next X
        End While
        Erase msPropName
        Erase txtMin
        Erase txtMax
        Erase mlPropID
    End Sub

    Public Sub SetFromItemType(ByVal yType As MarketListType)
        myItemType = yType

        ClearList()

        'Only select items can be requested via a Buy Order...
        Select Case yType
            Case MarketListType.eArmorComponent
                LoadArmorComponentList()
            Case MarketListType.eBeamPulseComponent
                LoadPulseComponentList()
            Case MarketListType.eBeamSolidComponent
                LoadSolidComponentList()
            Case MarketListType.eEngineComponent
                LoadEngineComponentList()
            Case MarketListType.eMissileComponent
                LoadMissileComponentList()
            Case MarketListType.eProjectileComponent
                LoadProjectileComponentList()
            Case MarketListType.eRadarComponent
                LoadRadarComponentList()
            Case MarketListType.eShieldComponent
                LoadShieldComponentList()
            Case MarketListType.eMinerals, MarketListType.eAlloys
                LoadMaterialItemList()
        End Select
    End Sub

	Private Sub AddRow(ByVal lIndex As Int32, ByVal sPropName As String, ByVal bMinVis As Boolean, ByVal bMaxVis As Boolean, ByVal lValue As Int32, ByVal lMinValue As Int32, ByVal lMaxValue As Int32, ByVal lEntryChars As Int32)
		mlPropID(lIndex) = lValue
		msPropName(lIndex) = sPropName
		mlMinValue(lIndex) = lMinValue
		mlMaxValue(lIndex) = lMaxValue

		'txtMin initial props
		txtMin(lIndex) = New UITextBox(MyBase.moUILib)
		With txtMin(lIndex)
			.ControlName = "txtMin(" & lIndex & ")"
			.Left = 145
			.Top = 25 + (lIndex * 20)
			.Width = 50
			.Height = 18
			.Enabled = True
			.Visible = bMinVis
			.Caption = "0"
			.ForeColor = muSettings.InterfaceTextBoxForeColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
			.BackColorEnabled = muSettings.InterfaceTextBoxFillColor
			.BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
			.MaxLength = lEntryChars
			.BorderColor = muSettings.InterfaceBorderColor
			.DoNotRender = UITextBox.DoNotRenderSetting.eDoNotRenderControl
			If lMaxValue <> Int32.MaxValue Then
				.ToolTipText = "Any value between " & lMinValue & " and " & lMaxValue & "."
			Else : .ToolTipText = ""
			End If
		End With
		Me.AddChild(CType(txtMin(lIndex), UIControl))

		'txtMax initial props
		txtMax(lIndex) = New UITextBox(MyBase.moUILib)
		With txtMax(lIndex)
			.ControlName = "txtMax(" & lIndex & ")"
			.Left = 200
			.Top = 25 + (lIndex * 20)
			.Width = 50
			.Height = 18
			.Enabled = True
			.Visible = bMaxVis
			.Caption = "0"
			.ForeColor = muSettings.InterfaceTextBoxForeColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
			.BackColorEnabled = muSettings.InterfaceTextBoxFillColor
			.BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
			.MaxLength = lEntryChars
			.BorderColor = muSettings.InterfaceBorderColor
			.DoNotRender = UITextBox.DoNotRenderSetting.eDoNotRenderControl
			If lMaxValue <> Int32.MaxValue Then
				.ToolTipText = "Any value between " & lMinValue & " and " & lMaxValue & "."
			Else : .ToolTipText = ""
			End If
		End With
		Me.AddChild(CType(txtMax(lIndex), UIControl))
	End Sub

    Public Function ValidateData() As Boolean
        If msPropName Is Nothing Then Return False
        Dim lCnt As Int32 = 0
        For X As Int32 = 0 To msPropName.GetUpperBound(0)
			If CInt(Val(txtMin(X).Caption)) > 0 OrElse CInt(Val(txtMax(X).Caption)) > 0 Then
				lCnt += 1

				Dim lMin As Int32 = CInt(Val(txtMin(X).Caption))
				Dim lMax As Int32 = CInt(Val(txtMax(X).Caption))
				If lMin < mlMinValue(X) OrElse lMin > mlMaxValue(X) Then Return False
				If lMax > mlMaxValue(X) OrElse lMax < mlMinValue(X) Then Return False
			End If
		Next X
        Return lCnt <> 0
    End Function

    Public Function GetPropertyList() As Byte()
        Dim lCnt As Int32 = 0
        For X As Int32 = 0 To msPropName.GetUpperBound(0)
            If CInt(Val(txtMin(X).Caption)) > 0 OrElse CInt(Val(txtMax(X).Caption)) > 0 Then lCnt += 1
        Next X

        Dim yResp(lCnt * 9) As Byte
        Dim lPos As Int32 = 1
        yResp(0) = CByte(lCnt)
        For X As Int32 = 0 To msPropName.GetUpperBound(0)
            If CInt(Val(txtMin(X).Caption)) > 0 OrElse CInt(Val(txtMax(X).Caption)) > 0 Then
                System.BitConverter.GetBytes(mlPropID(X)).CopyTo(yResp, lPos) : lPos += 4

                Dim bTime As Boolean = mlPropID(X) = elBuyOrderPropID.ePulseROF OrElse mlPropID(X) = elBuyOrderPropID.eSolidROF OrElse _
                    mlPropID(X) = elBuyOrderPropID.eMissileROF OrElse mlPropID(X) = elBuyOrderPropID.eMissileFlightTime OrElse _
                    mlPropID(X) = elBuyOrderPropID.eProjectileROF OrElse mlPropID(X) = elBuyOrderPropID.eShieldRechargeFreq

				Dim lMinVal As Int32 = CInt(Val(txtMin(X).Caption))
				Dim lMaxVal As Int32 = CInt(Val(txtMax(X).Caption))

				If lMinVal = lMaxVal Then
					Dim lFinalVal As Int32 = lMinVal
					If bTime = True Then lFinalVal *= 30
					System.BitConverter.GetBytes(lFinalVal).CopyTo(yResp, lPos) : lPos += 4
					yResp(lPos) = BuyOrderCompareTypes.eEqualTo
				ElseIf lMinVal > 0 AndAlso (lMaxVal = 0 OrElse lMaxVal < lMinVal) Then
					Dim lFinalVal As Int32 = lMinVal
					If bTime = True Then lFinalVal *= 30
					System.BitConverter.GetBytes(lFinalVal).CopyTo(yResp, lPos) : lPos += 4
					yResp(lPos) = BuyOrderCompareTypes.eGreaterThanEqualTo
				Else
					Dim lFinalVal As Int32 = CInt(Val(txtMax(X).Caption))
					If bTime = True Then lFinalVal *= 30
					System.BitConverter.GetBytes(lFinalVal).CopyTo(yResp, lPos) : lPos += 4
					yResp(lPos) = BuyOrderCompareTypes.eLessThanEqualTo : lPos += 1
				End If
			End If
		Next X
        Return yResp

    End Function

#Region "  Load List Routines  "
    Private Sub LoadArmorComponentList()
        'eArmorBeamResist to eArmorIntegrity
        Dim lUB As Int32 = Math.Abs(Math.Abs(elBuyOrderPropID.eArmorIntegrity) - Math.Abs(elBuyOrderPropID.eArmorBeamResist))
        ReDim msPropName(lUB)
        ReDim txtMin(lUB)
        ReDim txtMax(lUB)
		ReDim mlPropID(lUB)
		ReDim mlMaxValue(lUB)
		ReDim mlMinValue(lUB)

        For X As Int32 = 0 To lUB
			Dim lValue As Int32 = elBuyOrderPropID.eArmorBeamResist - X

			AddRow(X, BuyOrder.GetBuyOrderPropertyText(lValue), lValue <> elBuyOrderPropID.eArmorHullUsagePerPlate, lValue = elBuyOrderPropID.eArmorHullUsagePerPlate, lValue, 0, 100, 3)
        Next X
    End Sub
    Private Sub LoadPulseComponentList()
        'ePulseAccuracy to ePulsePower
        Dim lUB As Int32 = Math.Abs(Math.Abs(elBuyOrderPropID.ePulseAccuracy) - Math.Abs(elBuyOrderPropID.ePulsePower))
        ReDim msPropName(lUB)
        ReDim txtMin(lUB)
        ReDim txtMax(lUB)
		ReDim mlPropID(lUB)
		ReDim mlMaxValue(lUB)
		ReDim mlMinValue(lUB)

        For X As Int32 = 0 To lUB
            Dim lValue As Int32 = elBuyOrderPropID.ePulseAccuracy - X
			Dim bMinVis As Boolean = lValue > elBuyOrderPropID.ePulseROF
 
			Dim lMax As Int32
			Dim lChars As Int32

			If lValue = elBuyOrderPropID.ePulseAccuracy OrElse lValue = elBuyOrderPropID.ePulseAOE OrElse lValue = elBuyOrderPropID.ePulseRange Then
				lMax = 255
				lChars = 3
			Else
				lMax = Int32.MaxValue
				lChars = 9
			End If

			AddRow(X, BuyOrder.GetBuyOrderPropertyText(lValue), bMinVis, Not bMinVis, lValue, 0, lMax, lChars)
        Next X
    End Sub
    Private Sub LoadSolidComponentList()
        'eSolidAccuracy to eSolidPower
        Dim lUB As Int32 = Math.Abs(Math.Abs(elBuyOrderPropID.eSolidPower) - Math.Abs(elBuyOrderPropID.eSolidAccuracy))
        ReDim msPropName(lUB)
        ReDim txtMin(lUB)
        ReDim txtMax(lUB)
		ReDim mlPropID(lUB)
		ReDim mlMaxValue(lUB)
		ReDim mlMinValue(lUB)

        For X As Int32 = 0 To lUB
            Dim lValue As Int32 = elBuyOrderPropID.eSolidAccuracy - X
			Dim bMaxVis As Boolean = lValue = elBuyOrderPropID.eSolidROF OrElse lValue = elBuyOrderPropID.eSolidDmgType OrElse lValue = elBuyOrderPropID.eSolidPower OrElse lValue = elBuyOrderPropID.eSolidHull

 
			Dim lMax As Int32
			Dim lChars As Int32
			If lValue = elBuyOrderPropID.eSolidAccuracy OrElse lValue = elBuyOrderPropID.eSolidRange Then
				lMax = 255
				lChars = 3
			ElseIf lValue = elBuyOrderPropID.eSolidDmgType Then
				lChars = 1
				lMax = 1
			Else
				lMax = Int32.MaxValue
				lChars = 9
			End If

			AddRow(X, BuyOrder.GetBuyOrderPropertyText(lValue), Not bMaxVis, bMaxVis, lValue, 0, lMax, lChars)
        Next X
    End Sub
    Private Sub LoadEngineComponentList()
        'eEngineThrust to eEngineHull
        Dim lUB As Int32 = Math.Abs(Math.Abs(elBuyOrderPropID.eEngineHull) - Math.Abs(elBuyOrderPropID.eEngineThrust))
        ReDim msPropName(lUB)
        ReDim txtMin(lUB)
        ReDim txtMax(lUB)
		ReDim mlPropID(lUB)
		ReDim mlMaxValue(lUB)
		ReDim mlMinValue(lUB)

        For X As Int32 = 0 To lUB
            Dim lValue As Int32 = elBuyOrderPropID.eEngineThrust - X
			Dim bMinVis As Boolean = lValue <> elBuyOrderPropID.eEngineHull

			Dim lMax As Int32
			Dim lChars As Int32
			If lValue = elBuyOrderPropID.eEngineManeuver OrElse lValue = elBuyOrderPropID.eEngineMaxSpeed Then
				lMax = 255
				lChars = 3
			Else
				lMax = Int32.MaxValue
				lChars = 9
			End If

			AddRow(X, BuyOrder.GetBuyOrderPropertyText(lValue), bMinVis, Not bMinVis, lValue, 0, lMax, lChars)
        Next X
    End Sub
    Private Sub LoadMissileComponentList()
        'eMissileROF to eMissilePower
        Dim lUB As Int32 = Math.Abs(Math.Abs(elBuyOrderPropID.eMissilePower) - Math.Abs(elBuyOrderPropID.eMissileROF))
        ReDim msPropName(lUB)
        ReDim txtMin(lUB)
        ReDim txtMax(lUB)
		ReDim mlPropID(lUB)
		ReDim mlMaxValue(lUB)
		ReDim mlMinValue(lUB)

        For X As Int32 = 0 To lUB
            Dim lValue As Int32 = elBuyOrderPropID.eMissileROF - X
			Dim bMaxVis As Boolean = lValue = elBuyOrderPropID.eMissileROF OrElse lValue = elBuyOrderPropID.eMissileMissileHullSize OrElse _
			  lValue = elBuyOrderPropID.eMissileHull OrElse lValue = elBuyOrderPropID.eMissilePower

			Dim lMax As Int32
			Dim lChars As Int32
			If lValue = elBuyOrderPropID.eMissileAccuracy OrElse lValue = elBuyOrderPropID.eMissileManeuver OrElse lValue = elBuyOrderPropID.eMissileMaxSpeed OrElse _
			  lValue = elBuyOrderPropID.eMissileAOE OrElse lValue = elBuyOrderPropID.eMissilePayload Then
				lMax = 255
				lChars = 3
			Else
				lMax = Int32.MaxValue
				lChars = 9
			End If

			AddRow(X, BuyOrder.GetBuyOrderPropertyText(lValue), Not bMaxVis, bMaxVis, lValue, 0, lMax, lChars)
		Next X
    End Sub
    Private Sub LoadProjectileComponentList()
        'eProjectileROF to eProjectilePower
        Dim lUB As Int32 = Math.Abs(Math.Abs(elBuyOrderPropID.eProjectilePower) - Math.Abs(elBuyOrderPropID.eProjectileROF))
        ReDim msPropName(lUB)
        ReDim txtMin(lUB)
        ReDim txtMax(lUB)
		ReDim mlPropID(lUB)
		ReDim mlMaxValue(lUB)
		ReDim mlMinValue(lUB)

        For X As Int32 = 0 To lUB
            Dim lValue As Int32 = elBuyOrderPropID.eProjectileROF - X
			Dim bMaxVis As Boolean = lValue = elBuyOrderPropID.eProjectileROF OrElse lValue = elBuyOrderPropID.eProjectileAmmoSize OrElse _
			  lValue = elBuyOrderPropID.eProjectileHull OrElse lValue = elBuyOrderPropID.eProjectilePower

			Dim lMax As Int32
			Dim lChars As Int32
			If lValue = elBuyOrderPropID.eProjectileMaxRng OrElse lValue = elBuyOrderPropID.eProjectilePayload OrElse _
			 lValue = elBuyOrderPropID.eProjectileAOE OrElse lValue = elBuyOrderPropID.eProjectilePierceRatio Then
				lMax = 255
				lChars = 3
			Else
				lMax = Int32.MaxValue
				lChars = 9
			End If


			AddRow(X, BuyOrder.GetBuyOrderPropertyText(lValue), Not bMaxVis, bMaxVis, lValue, 0, lMax, lChars)
        Next X
    End Sub
    Private Sub LoadRadarComponentList()
        'eRadarAccuracy to eRadarHull
        Dim lUB As Int32 = Math.Abs(Math.Abs(elBuyOrderPropID.eRadarHull) - Math.Abs(elBuyOrderPropID.eRadarAccuracy))
        ReDim msPropName(lUB)
        ReDim txtMin(lUB)
        ReDim txtMax(lUB)
		ReDim mlPropID(lUB)
		ReDim mlMaxValue(lUB)
		ReDim mlMinValue(lUB)

        For X As Int32 = 0 To lUB
            Dim lValue As Int32 = elBuyOrderPropID.eRadarAccuracy - X
			Dim bMaxVis As Boolean = lValue = elBuyOrderPropID.eRadarPower OrElse lValue = elBuyOrderPropID.eRadarHull

			Dim lMax As Int32
			Dim lChars As Int32

			If lValue = elBuyOrderPropID.eRadarPower OrElse lValue = elBuyOrderPropID.eRadarHull Then
				lMax = Int32.MaxValue
				lChars = 9
			Else
				lMax = 255
				lChars = 3
			End If

            If lValue = elBuyOrderPropID.eRadarJamImm OrElse lValue = elBuyOrderPropID.eRadarJamTargets OrElse lValue = elBuyOrderPropID.eRadarJamEffect Then
				AddRow(X, BuyOrder.GetBuyOrderPropertyText(lValue), False, False, lValue, 0, lMax, lChars)
			Else : AddRow(X, BuyOrder.GetBuyOrderPropertyText(lValue), Not bMaxVis, bMaxVis, lValue, 0, lMax, lChars)
            End If
        Next X
    End Sub
    Private Sub LoadShieldComponentList()
        'eShieldMaxHP to eShieldHull
        Dim lUB As Int32 = Math.Abs(Math.Abs(elBuyOrderPropID.eShieldHull) - Math.Abs(elBuyOrderPropID.eShieldMaxHP))
        ReDim msPropName(lUB)
        ReDim txtMin(lUB)
        ReDim txtMax(lUB)
		ReDim mlPropID(lUB)
		ReDim mlMaxValue(lUB)
		ReDim mlMinValue(lUB)

        For X As Int32 = 0 To lUB
			Dim lValue As Int32 = elBuyOrderPropID.eShieldMaxHP - X

			Dim lMax As Int32 = Int32.MaxValue
			Dim lChars As Int32 = 9

            Select Case CType(lValue, elBuyOrderPropID)
                Case elBuyOrderPropID.eShieldHull, elBuyOrderPropID.eShieldPower, elBuyOrderPropID.eShieldRechargeFreq
					AddRow(X, BuyOrder.GetBuyOrderPropertyText(lValue), False, True, lValue, 0, lMax, lChars)
                Case Else
					AddRow(X, BuyOrder.GetBuyOrderPropertyText(lValue), True, False, lValue, 0, lMax, lChars)
            End Select
        Next X
    End Sub
    Private Sub LoadMaterialItemList()
        Dim lUB As Int32 = glMineralPropertyUB
        ReDim msPropName(lUB)
        ReDim txtMin(lUB)
        ReDim txtMax(lUB)
		ReDim mlPropID(lUB)
		ReDim mlMaxValue(lUB)
		ReDim mlMinValue(lUB)

        For X As Int32 = 0 To lUB
			If glMineralPropertyIdx(X) <> -1 Then AddRow(X, goMineralProperty(X).MineralPropertyName, True, True, goMineralProperty(X).ObjectID, 0, 10, 2)
        Next X
    End Sub
#End Region

    Private Sub fraBuyOrderRequires_OnRenderEnd() Handles Me.OnRenderEnd
        If GFXEngine.gbPaused = True OrElse GFXEngine.gbDeviceLost = True Then Return
        If msPropName Is Nothing = False Then
            'Do our multi-color blends first
            Using oSprite As New Sprite(MyBase.moUILib.oDevice)
                oSprite.Begin(SpriteFlags.AlphaBlend)
                Try
                    For X As Int32 = 0 To msPropName.GetUpperBound(0)
                        'Next, do our multi-color fills for the min/max box
                        If txtMax(X).Visible = True Then
                            With txtMax(X)
                                DoMultiColorFill(New Rectangle(.GetAbsolutePosition, New Size(.Width, .Height)), muSettings.InterfaceTextBoxFillColor, .GetAbsolutePosition, oSprite)
                            End With
                        End If
                        If txtMin(X).Visible = True Then
                            With txtMin(X)
                                DoMultiColorFill(New Rectangle(.GetAbsolutePosition, New Size(.Width, .Height)), muSettings.InterfaceTextBoxFillColor, .GetAbsolutePosition, oSprite)
                            End With
                        End If
                    Next X
                Catch ex As Exception

                End Try
                oSprite.End()
                oSprite.Dispose()
            End Using

            'Set up our line to draw
            'ValidateBorderLine()
            'moBorderLine.Antialias = True
            'moBorderLine.Width = BorderLineWidth
            'moBorderLine.Begin()
            BPLine.PrepareMultiDraw(BorderLineWidth, True)

            For X As Int32 = 0 To msPropName.GetUpperBound(0)

                If txtMax(X).Visible = True Then
                    With txtMax(X)
                        Dim v2Lines(4) As Vector2
                        Dim oLoc As Point = .GetAbsolutePosition()
                        v2Lines(0).X = oLoc.X : v2Lines(0).Y = oLoc.Y
                        v2Lines(1).X = oLoc.X + .Width : v2Lines(1).Y = oLoc.Y
                        v2Lines(2).X = oLoc.X + .Width : v2Lines(2).Y = oLoc.Y + .Height
                        v2Lines(3).X = oLoc.X : v2Lines(3).Y = oLoc.Y + .Height
                        v2Lines(4).X = oLoc.X : v2Lines(4).Y = oLoc.Y
                        'moBorderLine.Draw(v2Lines, .BorderColor)
                        BPLine.MultiDrawLine(v2Lines, .BorderColor)

                        If .HasFocus = True Then
                            ReDim v2Lines(1)
                            Dim sTemp As String = Trim$(Mid$(.Caption, 1, .SelStart))
                            Dim oRect As Rectangle = .GetTextDimensions(Mid$(.Caption, 1, .SelStart))
                            If sTemp.Length < .SelStart Then
                                oRect.Width += (5 * (.SelStart - sTemp.Length))
                            End If
                            If .Caption = "" Then
                                v2Lines(0).X = oLoc.X + 5
                                v2Lines(0).Y = oLoc.Y + 2
                                v2Lines(1).X = v2Lines(0).X
                                v2Lines(1).Y = oLoc.Y + .Height - 2
                            Else
                                v2Lines(0).X = oRect.Left + oRect.Width + oLoc.X + 5
                                v2Lines(0).Y = oLoc.Y + 2
                                v2Lines(1).X = v2Lines(0).X
                                v2Lines(1).Y = oLoc.Y + .Height - 2
                            End If
                            'moBorderLine.Draw(v2Lines, .BorderColor)
                            BPLine.MultiDrawLine(v2Lines, .BorderColor)
                        End If
                    End With
                End If
                If txtMin(X).Visible = True Then
                    With txtMin(X)
                        Dim v2Lines(4) As Vector2
                        Dim oLoc As Point = .GetAbsolutePosition()
                        v2Lines(0).X = oLoc.X : v2Lines(0).Y = oLoc.Y
                        v2Lines(1).X = oLoc.X + .Width : v2Lines(1).Y = oLoc.Y
                        v2Lines(2).X = oLoc.X + .Width : v2Lines(2).Y = oLoc.Y + .Height
                        v2Lines(3).X = oLoc.X : v2Lines(3).Y = oLoc.Y + .Height
                        v2Lines(4).X = oLoc.X : v2Lines(4).Y = oLoc.Y
                        'moBorderLine.Draw(v2Lines, .BorderColor)
                        BPLine.MultiDrawLine(v2Lines, .BorderColor)

                        If .HasFocus = True Then
                            ReDim v2Lines(1)
                            Dim sTemp As String = Trim$(Mid$(.Caption, 1, .SelStart))
                            Dim oRect As Rectangle = .GetTextDimensions(Mid$(.Caption, 1, .SelStart))
                            If sTemp.Length < .SelStart Then
                                oRect.Width += (5 * (.SelStart - sTemp.Length))
                            End If
                            If .Caption = "" Then
                                v2Lines(0).X = oLoc.X + 5
                                v2Lines(0).Y = oLoc.Y + 2
                                v2Lines(1).X = v2Lines(0).X
                                v2Lines(1).Y = oLoc.Y + .Height - 2
                            Else
                                v2Lines(0).X = oRect.Left + oRect.Width + oLoc.X + 5
                                v2Lines(0).Y = oLoc.Y + 2
                                v2Lines(1).X = v2Lines(0).X
                                v2Lines(1).Y = oLoc.Y + .Height - 2
                            End If
                            'moBorderLine.Draw(v2Lines, .BorderColor)
                            BPLine.MultiDrawLine(v2Lines, .BorderColor)
                        End If
                    End With
                End If
            Next X
            'moBorderLine.End()
            BPLine.EndMultiDraw()

            'Finally, render our text
            Using oFont As New Font(MyBase.moUILib.oDevice, New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                Using oSprite As New Sprite(MyBase.moUILib.oDevice)
                    oSprite.Begin(SpriteFlags.AlphaBlend)
                    Try
                        Dim lTopAdj As Int32 = Me.GetAbsolutePosition().Y + 25
                        Dim lLeft As Int32 = Me.GetAbsolutePosition.X + 5
                        For X As Int32 = 0 To msPropName.GetUpperBound(0)
                            oFont.DrawText(oSprite, msPropName(X), lLeft, lTopAdj + (X * 20), muSettings.InterfaceBorderColor)

                            If txtMax(X).Visible = True Then
                                With txtMax(X)
                                    Dim oLoc As Point = .GetAbsolutePosition()
                                    oFont.DrawText(oSprite, .Caption, oLoc.X + 5, oLoc.Y, .ForeColor)
                                End With
                            End If
                            If txtMin(X).Visible = True Then
                                With txtMin(X)
                                    Dim oLoc As Point = .GetAbsolutePosition()
                                    oFont.DrawText(oSprite, .Caption, oLoc.X + 5, oLoc.Y, .ForeColor)
                                End With
                            End If
                        Next X
                    Catch
                        'Do nothing
                    End Try

                    oSprite.End()
                    oSprite.Dispose()
                End Using
                oFont.Dispose()
            End Using

        End If
    End Sub

    '    sRenderedText = sCaption
    '    If lSelStart > sRenderedText.Length Then lSelStart = sRenderedText.Length
    '        RenderBorder(0, 0)

    '        'Draw the caption, but we need to scoot the rect over a bit
    '        oRect.X += 5
    '        oRect.Width -= 5

    '        'Ok, not multiline... see if our text scrolls off the side
    '        Dim rcTemp As Rectangle = moFont.MeasureString(Nothing, sRenderedText, FontFormat, ForeColor)
    '        While rcTemp.Width > oRect.Width
    '            If (FontFormat And DrawTextFormat.Right) <> 0 Then
    '                'ok, right aligned so we want the left side of the string
    '                If sRenderedText.Length > 0 Then
    '                    sRenderedText = sRenderedText.Substring(0, sRenderedText.Length - 2)
    '                Else : Exit While
    '                End If
    '            Else
    '                'ok, left aligned or Center aligned, so we want the right side of the string
    '                If sRenderedText.Length > 0 Then
    '                    sRenderedText = sRenderedText.Substring(1)
    '                Else : Exit While
    '                End If
    '            End If
    '            rcTemp = moFont.MeasureString(Nothing, sRenderedText, FontFormat, ForeColor)
    '        End While
    '        moFont.DrawText(Nothing, sRenderedText, oRect, FontFormat, ForeColor)

    '        If HasFocus = True Then
    '            Dim sTemp As String = Trim$(Mid$(sRenderedText, 1, SelStart))
    '            oRect = moFont.MeasureString(Nothing, Mid$(sRenderedText, 1, SelStart), FontFormat, ForeColor)
    '            If sTemp.Length < SelStart Then
    '                oRect.Width += (5 * (SelStart - sTemp.Length))
    '            End If
    '            If (FontFormat And DrawTextFormat.Right) <> 0 Then
    '                'ok, now, draw a line there
    '                If Caption = "" Then
    '                    moCursorVerts(0).X = oLoc.X + Me.Width - 5
    '                    moCursorVerts(0).Y = oLoc.Y + 2
    '                    moCursorVerts(1).X = moCursorVerts(0).X
    '                    moCursorVerts(1).Y = oLoc.Y + Me.Height - 2
    '                Else
    '                    moCursorVerts(0).X = (oLoc.X + Me.Width) - (oRect.Width + 5)
    '                    moCursorVerts(0).Y = oLoc.Y + 2
    '                    moCursorVerts(1).X = moCursorVerts(0).X
    '                    moCursorVerts(1).Y = oLoc.Y + Me.Height - 2
    '                End If
    '            Else
    '                'ok, now, draw a line there
    '                If Caption = "" Then
    '                    moCursorVerts(0).X = oLoc.X + 5
    '                    moCursorVerts(0).Y = oLoc.Y + 2
    '                    moCursorVerts(1).X = moCursorVerts(0).X
    '                    moCursorVerts(1).Y = oLoc.Y + Me.Height - 2
    '                Else
    '                    moCursorVerts(0).X = oRect.Left + oRect.Width + oLoc.X + 5
    '                    moCursorVerts(0).Y = oLoc.Y + 2
    '                    moCursorVerts(1).X = moCursorVerts(0).X
    '                    moCursorVerts(1).Y = oLoc.Y + Me.Height - 2
    '                End If
    '            End If

    '            moCursorLine.Begin()
    '            moCursorLine.Draw(moCursorVerts, ForeColor)
    '            moCursorLine.End()
    '        End If 

    Private Sub DoMultiColorFill(ByVal rcDest As Rectangle, ByVal clrFill As System.Drawing.Color, ByVal ptLoc As Point, ByRef oSpr As Sprite)
        Dim rcSrc As Rectangle

        Dim fX As Single
        Dim fY As Single

        If rcDest.Width = 0 OrElse rcDest.Height = 0 Then Exit Sub

        rcSrc.Location = New Point(192, 0)
        rcSrc.Width = 62
        rcSrc.Height = 64

        'Now, draw it...
        With oSpr
            fX = ptLoc.X * (rcSrc.Width / CSng(rcDest.Width))
            fY = ptLoc.Y * (rcSrc.Height / CSng(rcDest.Height))
            .Draw2D(MyBase.moUILib.oInterfaceTexture, rcSrc, rcDest, System.Drawing.Point.Empty, 0, New Point(CInt(fX), CInt(fY)), clrFill)
        End With
    End Sub
End Class