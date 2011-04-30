Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Interface created from Interface Builder
Public Class frmFleetOrders
    Inherits UIWindow

    Private lblTitle As UILabel
    Private lnDiv1 As UILine
    Private WithEvents btnSetDest As UIButton
    Private lblOrders As UILabel
    Private lblNewOrders As UILabel
    Private WithEvents btnConfirm As UIButton
    Private lnDiv2 As UILine
    Private WithEvents btnClose As UIButton

    Private mlFleetIndex As Int32 = -1

    Private mlDestSystemIdx As Int32 = -1

    Private mlPreviousView As Int32
    Private mlPrevCX As Int32
    Private mlPrevCY As Int32
    Private mlPrevCZ As Int32
    Private mlPrevCAX As Int32
    Private mlPrevCAY As Int32
    Private mlPrevCAZ As Int32

    Private mbMoving As Boolean = False

    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)

        'frmFleetOrders initial props
        With Me
            .lWindowMetricID = BPMetrics.eWindow.eFleetOrders
            .ControlName = "frmFleetOrders"
            .Left = 350
            .Top = 223
            .Width = 170
            .Height = 180
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = False
            .Moveable = True
            .BorderLineWidth = 1
        End With

        'btnClose initial props
        btnClose = New UIButton(oUILib)
        With btnClose
            .ControlName = "btnClose"
            .Left = 145
            .Top = 2
            .Width = 22
            .Height = 22
            .Enabled = True
            .Visible = True
            .Caption = "X"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
            .ToolTipText = "Click to Close"
        End With
        Me.AddChild(CType(btnClose, UIControl))

        'lblTitle initial props
        lblTitle = New UILabel(oUILib)
        With lblTitle
            .ControlName = "lblTitle"
            .Left = 5
            .Top = 5
            .Width = 160
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Battlegroup Name Goes Here"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblTitle, UIControl))

        'lnDiv1 initial props
        lnDiv1 = New UILine(oUILib)
        With lnDiv1
            .ControlName = "lnDiv1"
            .Left = 0
            .Top = 25
            .Width = 170
            .Height = 0
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(lnDiv1, UIControl))

        'btnSetDest initial props
        btnSetDest = New UIButton(oUILib)
        With btnSetDest
            .ControlName = "btnSetDest"
            .Left = 25
            .Top = 75
            .Width = 120
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Set Destination"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnSetDest, UIControl))

        'lblOrders initial props
        lblOrders = New UILabel(oUILib)
        With lblOrders
            .ControlName = "lblOrders"
            .Left = 5
            .Top = 30
            .Width = 154
            .Height = 36
            .Enabled = True
            .Visible = True
            .Caption = "Multiline Current Orders"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(0, DrawTextFormat)
        End With
        Me.AddChild(CType(lblOrders, UIControl))

        'lblNewOrders initial props
        lblNewOrders = New UILabel(oUILib)
        With lblNewOrders
            .ControlName = "lblNewOrders"
            .Left = 5
            .Top = 105
            .Width = 154
            .Height = 36
            .Enabled = True
            .Visible = True
            .Caption = "Multiline New Orders"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(0, DrawTextFormat)
        End With
        Me.AddChild(CType(lblNewOrders, UIControl))

        'btnConfirm initial props
        btnConfirm = New UIButton(oUILib)
        With btnConfirm
            .ControlName = "btnConfirm"
            .Left = 25
            .Top = 150
            .Width = 120
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Confirm"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnConfirm, UIControl))

        'lnDiv2 initial props
        lnDiv2 = New UILine(oUILib)
        With lnDiv2
            .ControlName = "lnDiv2"
            .Left = 0
            .Top = 102
            .Width = 170
            .Height = 0
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(lnDiv2, UIControl))

        MyBase.moUILib.RemoveWindow(Me.ControlName)
        MyBase.moUILib.AddWindow(Me)
    End Sub

    Public Sub SetFromFleet(ByVal lFleetID As Int32)
        mbMoving = False
        If lFleetID = -1 Then
            MyBase.moUILib.RemoveWindow(Me.ControlName)
            Return
        End If

        lblTitle.Caption = ""
        lblOrders.Caption = ""
        lblNewOrders.Caption = "None"
        If btnSetDest.Enabled = False Then btnSetDest.Enabled = True
        If btnConfirm.Enabled = True Then btnConfirm.Enabled = False

        For X As Int32 = 0 To goCurrentPlayer.mlUnitGroupUB
            If goCurrentPlayer.mlUnitGroupIdx(X) = lFleetID Then
                mlFleetIndex = X
                With goCurrentPlayer.moUnitGroups(X)

                    lblTitle.Caption = .sName

                    'Now, our current orders
                    If .lInterSystemOriginID <> -1 AndAlso .lInterSystemTargetID <> -1 Then
                        'Moving
                        lblOrders.Caption = "Moving to " & goGalaxy.GetSystemName(.lInterSystemTargetID) & vbCrLf & "  ETA: " & .GetInterSystemMovementETA
                        mbMoving = True
                    Else
                        'TODO: Put more here as more becomes determined
                        'Deployed
                        If .iParentTypeID = ObjectType.eSolarSystem Then
                            lblOrders.Caption = "Deployed" & vbCrLf & "  " & goGalaxy.GetSystemName(.lParentID) & " (system)"
                        Else
                            Dim sTemp As String = goGalaxy.GetPlanetName(.lParentID)
                            If sTemp = "" Then sTemp = GetCacheObjectValue(.lParentID, .iParentTypeID)
                            If sTemp = "Unknown" Then sTemp = "  On A Planet"
                            lblOrders.Caption = "Deployed" & vbCrLf & sTemp
                        End If
                    End If
                End With

                Exit For
            End If
        Next X
    End Sub

	Private Sub btnSetDest_Click(ByVal sName As String) Handles btnSetDest.Click
        btnSetDest.Enabled = False
        If goUILib.lUISelectState = UILib.eSelectState.eSetFleetDest Then Return
        '      glCurrentEnvirView = mlPreviousView
        '      With goCamera
        '          .mlCameraX = mlPrevCX : .mlCameraY = mlPrevCY : .mlCameraZ = mlPrevCZ
        '          .mlCameraAtX = mlPrevCAX : .mlCameraAtY = mlPrevCAY : .mlCameraAtZ = mlPrevCAZ
        '      End With
        '      goUILib.lUISelectState = UILib.eSelectState.eNoSelectState
        '      goUILib.AddNotification("Set Destination Cancelled", Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
        'Else

        mlPreviousView = glCurrentEnvirView
        With goCamera
            mlPrevCX = .mlCameraX : mlPrevCY = .mlCameraY : mlPrevCZ = .mlCameraZ
            mlPrevCAX = .mlCameraAtX : mlPrevCAY = .mlCameraAtY : mlPrevCAZ = .mlCameraAtZ

            .mlCameraAtX = 0 : .mlCameraAtY = 0 : .mlCameraAtZ = 0
            .mlCameraX = 0 : .mlCameraY = 1000 : .mlCameraZ = -1000
        End With
        glCurrentEnvirView = CurrentView.eGalaxyMapView
        goUILib.AddNotification("Left-Click Destination System...", Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
        goUILib.AddNotification("Right-Click to Cancel", Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
        goUILib.lUISelectState = UILib.eSelectState.eSetFleetDest
        'End If
	End Sub

	Private Sub btnConfirm_Click(ByVal sName As String) Handles btnConfirm.Click

		If mlDestSystemIdx <> -1 AndAlso mlFleetIndex <> -1 Then
			'ok, here we go
			Dim yMsg(9) As Byte
			System.BitConverter.GetBytes(GlobalMessageCode.eSetFleetDest).CopyTo(yMsg, 0)
			System.BitConverter.GetBytes(goCurrentPlayer.moUnitGroups(mlFleetIndex).ObjectID).CopyTo(yMsg, 2)
			System.BitConverter.GetBytes(goGalaxy.moSystems(mlDestSystemIdx).ObjectID).CopyTo(yMsg, 6)
			MyBase.moUILib.SendMsgToPrimary(yMsg)
		End If

		goUILib.lUISelectState = UILib.eSelectState.eNoSelectState
		glCurrentEnvirView = mlPreviousView
		With goCamera
			.mlCameraX = mlPrevCX : .mlCameraY = mlPrevCY : .mlCameraZ = mlPrevCZ
			.mlCameraAtX = mlPrevCAX : .mlCameraAtY = mlPrevCAY : .mlCameraAtZ = mlPrevCAZ
		End With
		MyBase.moUILib.RemoveWindow(Me.ControlName)
	End Sub

	Public Sub SetDestNewSystem(ByVal lSystemIdx As Int32)
		Try
			lblNewOrders.Caption = "None"
			mlDestSystemIdx = lSystemIdx

			'Ok, make sure the fleet index is set
			If mlFleetIndex = -1 Then Return
			If mlDestSystemIdx = -1 Then Return

			Dim vecLoc As Vector3
			'Now, get the fleet's LOC in system coordinates
			With goCurrentPlayer.moUnitGroups(mlFleetIndex)
				'Is the fleet moving?
				If .lInterSystemTargetID <> -1 AndAlso .lInterSystemOriginID <> -1 AndAlso .iParentTypeID = ObjectType.eGalaxy Then
					'Yes, the fleet is moving, so we need to determine where it is located on its current path
					Dim vecFrom As Vector3
					Dim vecTo As Vector3
					Dim lOriginID As Int32 = .lInterSystemOriginID
					Dim lTargetID As Int32 = .lInterSystemTargetID

					For lSysIdx As Int32 = 0 To goGalaxy.mlSystemUB
						With goGalaxy.moSystems(lSysIdx)
							If .ObjectID = lOriginID Then
								lOriginID = -1
								vecFrom.X = .LocX : vecFrom.Y = .LocY : vecFrom.Z = .LocZ
							ElseIf .ObjectID = lTargetID Then
								lTargetID = -1
								vecTo.X = .LocX : vecTo.Y = .LocY : vecTo.Z = .LocZ
							End If
						End With
						If lOriginID = -1 AndAlso lTargetID = -1 Then Exit For
					Next lSysIdx

					'Now... find out the mult
					Dim fMult As Single = CSng(.lInterSystemMoveCyclesRemaining / .InterSystemTotalCycles)
					vecLoc = Vector3.Subtract(vecFrom, vecTo)
					vecLoc.Multiply(fMult)
					vecLoc.Add(vecFrom)
				Else
					'No, the fleet is not moving, check it's parent...
					If .iParentTypeID = ObjectType.eSolarSystem Then
						For lSysIdx As Int32 = 0 To goGalaxy.mlSystemUB
							If goGalaxy.moSystems(lSysIdx).ObjectID = .lParentID Then
								vecLoc.X = goGalaxy.moSystems(lSysIdx).LocX
								vecLoc.Y = goGalaxy.moSystems(lSysIdx).LocY
								vecLoc.Z = goGalaxy.moSystems(lSysIdx).LocZ
								Exit For
							End If
						Next lSysIdx
					ElseIf .iParentTypeID = ObjectType.ePlanet Then
						'Not guaranteed that our planet will be loaded
						Dim bFound As Boolean = False
						For lSysIdx As Int32 = 0 To goGalaxy.mlSystemUB
							For lPlntIdx As Int32 = 0 To goGalaxy.moSystems(lSysIdx).PlanetUB
								If goGalaxy.moSystems(lSysIdx).moPlanets(lPlntIdx).ObjectID = .lParentID Then
									vecLoc.X = goGalaxy.moSystems(lSysIdx).LocX
									vecLoc.Y = goGalaxy.moSystems(lSysIdx).LocY
									vecLoc.Z = goGalaxy.moSystems(lSysIdx).LocZ
									bFound = True
									Exit For
								End If
							Next lPlntIdx
							If bFound = True Then Exit For
						Next lSysIdx

						If bFound = False Then
							lblNewOrders.Caption = "Move to " & goGalaxy.moSystems(mlDestSystemIdx).SystemName
							Return
						End If
					End If
				End If
			End With

			'Now, get the destination
			With goGalaxy.moSystems(mlDestSystemIdx)
				Dim fDX As Single = .LocX - vecLoc.X
				Dim fDY As Single = .LocY - vecLoc.Y
				Dim fDZ As Single = .LocZ - vecLoc.Z
				fDX *= fDX
				fDY *= fDY
				fDZ *= fDZ

                Dim fDist As Single = CSng(Math.Sqrt(fDX + fDY + fDZ))
                fDist *= 0.5F

				'Multiply by 10 million for system size
				fDist *= 10000000

				'TODO: Put player NonGravityWellSpeed Multiplier here
				Dim fPlayerMult As Single = 1.0F
				fDist /= (goCurrentPlayer.moUnitGroups(mlFleetIndex).yInterSystemMovementSpeed * fPlayerMult)

				Dim lCycles As Int32
				If fDist > Int32.MaxValue Then lCycles = Int32.MaxValue Else lCycles = CInt(fDist)

				Dim lSeconds As Int32 = CInt(Math.Ceiling(lCycles / 33.3333F))
				Dim lMinutes As Int32 = lSeconds \ 60
				lSeconds -= (lMinutes * 60)
				Dim lHours As Int32 = lMinutes \ 60
				lMinutes -= (lHours * 60)
				Dim lDays As Int32 = lHours \ 24
				lHours -= (lDays * 24)

				lblNewOrders.Caption = "Move to " & .SystemName & vbCrLf & _
				   "  ETA: " & lDays.ToString("0#") & ":" & lHours.ToString("0#") & ":" & lMinutes.ToString("0#") & ":" & lSeconds.ToString("0#")
                If btnConfirm.Enabled = False Then btnConfirm.Enabled = True
            End With
		Catch
			goUILib.AddNotification("That route is impossible!", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
			lblNewOrders.Caption = "None"
			mlDestSystemIdx = -1
		End Try
	End Sub

	Private Sub frmFleetOrders_OnNewFrame() Handles Me.OnNewFrame

		If mbMoving = True AndAlso mlFleetIndex <> -1 Then
			With goCurrentPlayer.moUnitGroups(mlFleetIndex)
				If lblTitle.Caption <> .sName Then lblTitle.Caption = .sName
				'Moving
				Dim sValue As String = "Moving to " & goGalaxy.GetSystemName(.lInterSystemTargetID) & vbCrLf & "  ETA: " & .GetInterSystemMovementETA
				If sValue <> lblOrders.Caption Then lblOrders.Caption = sValue
			End With
		End If
	End Sub

    Public Sub RemoveMe()
        MyBase.moUILib.RemoveWindow(Me.ControlName)
        If MyBase.moUILib.lUISelectState = UILib.eSelectState.eSetFleetDest Then
            glCurrentEnvirView = mlPreviousView
            With goCamera
                .mlCameraX = mlPrevCX : .mlCameraY = mlPrevCY : .mlCameraZ = mlPrevCZ
                .mlCameraAtX = mlPrevCAX : .mlCameraAtY = mlPrevCAY : .mlCameraAtZ = mlPrevCAZ
            End With
            goUILib.lUISelectState = UILib.eSelectState.eNoSelectState
            goUILib.AddNotification("Set Destination Cancelled", Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
        End If
    End Sub

	Private Sub btnClose_Click(ByVal sName As String) Handles btnClose.Click
        RemoveMe()
    End Sub

    Private Sub frmFleetOrders_WindowMoved() Handles Me.WindowMoved
        Dim ofrm As UIWindow = MyBase.moUILib.GetWindow("frmFleet")
        If ofrm Is Nothing = False AndAlso ofrm.Visible = True Then
            ofrm.Left = Me.Left + Me.Width + 2
            ofrm.Top = Me.Top
        End If
        ofrm = Nothing
    End Sub
End Class