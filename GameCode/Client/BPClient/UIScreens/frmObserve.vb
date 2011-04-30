Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Interface created from Interface Builder
Public Class frmObserve
    Inherits UIWindow

    Private WithEvents lblObsTitle As UILabel
    Private WithEvents lblDetails As UILabel
    Private WithEvents btnClose As UIButton

    Private mlID As Int32 = -1
    Private miTypeID As Int16 = -1

    Private mlEntityIndex As Int32 = -1

    Private msw_FPS As Stopwatch
    Private mlFPS As Int32 = 0
    Private mbLoading As Boolean = True

    Public Sub New(ByRef oUILib As UILib, ByVal lEntityID As Int32, ByVal iTypeId As Int16)
        MyBase.New(oUILib)

        'frmObserve initial props
        With Me
            .lWindowMetricID = BPMetrics.eWindow.eObserve
            .ControlName = "frmObserve_" & lEntityID.ToString
            .Width = 175
            .Height = 512
            .DoWindowInitialPosition(100, 100, 175, 512, muSettings.ObserveX, muSettings.ObserveY, -1, -1, False)

            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = True
        End With

        'lblObsTitle initial props
        lblObsTitle = New UILabel(oUILib)
        With lblObsTitle
            .ControlName = "lblObsTitle"
            .Left = 5
            .Top = 0
            .Width = 149
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Observing Unit ID: 99999"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblObsTitle, UIControl))

        'lblDetails initial props
        lblDetails = New UILabel(oUILib)
        With lblDetails
            .ControlName = "lblDetails"
            .Left = 5
            .Top = 20
            .Width = 165
            .Height = 492
            .Enabled = True
            .Visible = True
            .Caption = "Details"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(0, DrawTextFormat)
        End With
        Me.AddChild(CType(lblDetails, UIControl))

        'btnClose initial props
        btnClose = New UIButton(oUILib)
        With btnClose
            .ControlName = "btnClose"
            .Left = 150
            .Top = 1
            .Width = 24
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "X"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnClose, UIControl))

        mlID = lEntityID
        miTypeID = iTypeId

        If mlID = -254 Then
            If gsw_Camera Is Nothing = False Then gsw_Camera = Nothing
            gsw_Camera = New Stopwatch()
            If gsw_Movement Is Nothing = False Then gsw_Movement = Nothing
            gsw_Movement = New Stopwatch()
            If gsw_Geography Is Nothing = False Then gsw_Geography = Nothing
            gsw_Geography = New Stopwatch()
            If gsw_Models Is Nothing = False Then gsw_Models = Nothing
            gsw_Models = New Stopwatch()
            If gsw_Caches Is Nothing = False Then gsw_Caches = Nothing
            gsw_Caches = New Stopwatch()
            If gsw_Minimap Is Nothing = False Then gsw_Minimap = Nothing
            gsw_Minimap = New Stopwatch()
            If gsw_WpnFX Is Nothing = False Then gsw_WpnFX = Nothing
            gsw_WpnFX = New Stopwatch()
            If gsw_BombFX Is Nothing = False Then gsw_BombFX = Nothing
            gsw_BombFX = New Stopwatch()
            If gsw_ShieldFX Is Nothing = False Then gsw_ShieldFX = Nothing
            gsw_ShieldFX = New Stopwatch()
            If gsw_BurnFX Is Nothing = False Then gsw_BurnFX = Nothing
            gsw_BurnFX = New Stopwatch()
            If gsw_FireworksFX Is Nothing = False Then gsw_FireworksFX = Nothing
            gsw_FireworksFX = New Stopwatch()
            If gsw_Explosions Is Nothing = False Then gsw_Explosions = Nothing
            gsw_Explosions = New Stopwatch
            If gsw_AOEExplosions Is Nothing = False Then gsw_AOEExplosions = Nothing
            gsw_AOEExplosions = New Stopwatch
            If gsw_CommitEntitySoundChanges Is Nothing = False Then gsw_CommitEntitySoundChanges = Nothing
            gsw_CommitEntitySoundChanges = New Stopwatch
            If gsw_DeathFX Is Nothing = False Then gsw_DeathFX = Nothing
            gsw_DeathFX = New Stopwatch()
            If gsw_HPBars Is Nothing = False Then gsw_HPBars = Nothing
            gsw_HPBars = New Stopwatch()
            If gsw_UI Is Nothing = False Then gsw_UI = Nothing
            gsw_UI = New Stopwatch()
            If gsw_PostEffects Is Nothing = False Then gsw_PostEffects = Nothing
            gsw_PostEffects = New Stopwatch()
            If gsw_Present Is Nothing = False Then gsw_Present = Nothing
            gsw_Present = New Stopwatch()

            msw_FPS = Stopwatch.StartNew()

            lblObsTitle.Caption = "PERFORMANCE"

            gbMonitorPerformance = True
        End If

        MyBase.moUILib.RemoveWindow(Me.ControlName)
        MyBase.moUILib.AddWindow(Me)

        mbLoading = False
    End Sub

	Private Sub btnClose_Click(ByVal sName As String) Handles btnClose.Click
		MyBase.moUILib.RemoveWindow(Me.ControlName)

		If mlID = -254 Then gbMonitorPerformance = False
	End Sub

    Private Sub frmObserve_OnNewFrame() Handles Me.OnNewFrame
        Dim oSB As System.Text.StringBuilder = New System.Text.StringBuilder

        If mlID = -1 OrElse (mlID <> -254 AndAlso goCurrentEnvir Is Nothing) Then
            btnClose_Click(btnClose.ControlName)
            Return
        End If

        If mlID = -255 Then
            If goSound Is Nothing Then
				btnClose_Click(btnClose.ControlName)
            Else
                Dim sTemp As String = ""
                sTemp = goSound.mlExcitementLevel.ToString
                If lblDetails.Caption <> sTemp Then lblDetails.Caption = sTemp
            End If
            Return
        ElseIf mlID = -254 Then
            mlFPS += 1
            If msw_FPS.ElapsedMilliseconds > 1000 Then
                msw_FPS.Reset()

                'Update our stats
                oSB.AppendLine("FPS: " & mlFPS)

                Dim fValue As Single
                fValue = CSng(gsw_Camera.ElapsedTicks / Stopwatch.Frequency) * 100
                oSB.AppendLine("Camera: " & fValue.ToString("0.0#"))
                fValue = CSng(gsw_Movement.ElapsedTicks / Stopwatch.Frequency) * 100
                oSB.AppendLine("Movement: " & fValue.ToString("0.0#"))
                fValue = CSng(gsw_Geography.ElapsedTicks / Stopwatch.Frequency) * 100
                oSB.AppendLine("Geography: " & fValue.ToString("0.0#") & " for " & glQuadsRendered)
                fValue = CSng(gsw_Models.ElapsedTicks / Stopwatch.Frequency) * 100
                oSB.AppendLine("Models: " & fValue.ToString("0.0#") & " for " & glModelsRendered)
                fValue = CSng(gsw_Caches.ElapsedTicks / Stopwatch.Frequency) * 100
                oSB.AppendLine("Caches: " & fValue.ToString("0.0#") & " for " & glCachesRendered)
                fValue = CSng(gsw_Minimap.ElapsedTicks / Stopwatch.Frequency) * 100
                oSB.AppendLine("Minimap: " & fValue.ToString("0.0#") & " for " & glMinimapItemsRendered)
                fValue = CSng(gsw_WpnFX.ElapsedTicks / Stopwatch.Frequency) * 100
                oSB.AppendLine("WeaponFX: " & fValue.ToString("0.0#") & " for " & glWpnFXRendered)
                fValue = CSng(gsw_BombFX.ElapsedTicks / Stopwatch.Frequency) * 100
                oSB.AppendLine("BombFX: " & fValue.ToString("0.0#") & " for " & glBombFXRendered)
                fValue = CSng(gsw_ShieldFX.ElapsedTicks / Stopwatch.Frequency) * 100
                oSB.AppendLine("ShieldFX: " & fValue.ToString("0.0#") & " for " & glShieldFXRendered)
                fValue = CSng(gsw_CommitEntitySoundChanges.ElapsedTicks / Stopwatch.Frequency) * 100
                oSB.AppendLine("CommitSoundChg: " & fValue.ToString("0.0#"))
                fValue = CSng(gsw_Explosions.ElapsedTicks / Stopwatch.Frequency) * 100
                oSB.AppendLine("Explosions: " & fValue.ToString("0.0#") & " for " & glExplosionRendered)
                fValue = CSng(gsw_AOEExplosions.ElapsedTicks / Stopwatch.Frequency) * 100
                oSB.AppendLine("AOEExplosions: " & fValue.ToString("0.0#") & " for " & glAOEExplRendered)
                fValue = CSng(gsw_BurnFX.ElapsedTicks / Stopwatch.Frequency) * 100
                oSB.AppendLine("BurnFX: " & fValue.ToString("0.0#") & " for " & glBurnFXRendered)
                fValue = CSng(gsw_FireworksFX.ElapsedTicks / Stopwatch.Frequency) * 100
                oSB.AppendLine("Fireworks: " & fValue.ToString("0.0#") & " for " & glFireworksRendered)
                fValue = CSng(gsw_DeathFX.ElapsedTicks / Stopwatch.Frequency) * 100
                oSB.AppendLine("DeathFX: " & fValue.ToString("0.0#"))
                fValue = CSng(gsw_HPBars.ElapsedTicks / Stopwatch.Frequency) * 100
                oSB.AppendLine("HPBars/WPInds: " & fValue.ToString("0.0#"))
                fValue = CSng(gsw_UI.ElapsedTicks / Stopwatch.Frequency) * 100
                oSB.AppendLine("UI: " & fValue.ToString("0.0#") & " for " & glUI_Rendered.ToString)
                fValue = CSng(gsw_PostEffects.ElapsedTicks / Stopwatch.Frequency) * 100
                oSB.AppendLine("PostFX: " & fValue.ToString("0.0#"))
                fValue = CSng(gsw_Present.ElapsedTicks / Stopwatch.Frequency) * 100
                oSB.AppendLine("Present: " & fValue.ToString("0.0#"))
                oSB.AppendLine("Particles: " & BurnFX.ParticleEngine.l_PARTICLE_CNT.ToString)
                'oSB.AppendLine("R: " & MsgSystem.l_REGION_MSGS.ToString)
                'oSB.AppendLine("P: " & MsgSystem.l_PRIMARY_MSGS.ToString)
                If goSound Is Nothing = False Then oSB.AppendLine("New Sounds: " & goSound.lSoundStarts.ToString)
                oSB.AppendLine("C:" & goCamera.mlCameraX.ToString & "," & goCamera.mlCameraY.ToString & "," & goCamera.mlCameraZ.ToString)
                oSB.AppendLine("A:" & goCamera.mlCameraAtX.ToString & "," & goCamera.mlCameraAtY.ToString & "," & goCamera.mlCameraAtZ.ToString)
                oSB.AppendLine("Cache:" & GetCacheSize.ToString)
                If lblDetails.Caption <> oSB.ToString Then lblDetails.Caption = oSB.ToString

                'Now, reset all of our stuff...
                mlFPS = 0
                gsw_Camera.Reset()
                gsw_Movement.Reset()
                gsw_Geography.Reset()
                gsw_Models.Reset()
                gsw_Caches.Reset()
                gsw_Minimap.Reset()
                gsw_WpnFX.Reset()
                gsw_BombFX.Reset()
                gsw_CommitEntitySoundChanges.Reset()
                gsw_ShieldFX.Reset()
                gsw_BurnFX.Reset()
                gsw_FireworksFX.Reset()
                gsw_Explosions.Reset()
                gsw_AOEExplosions.Reset()
                gsw_DeathFX.Reset()
                gsw_HPBars.Reset()
                gsw_UI.Reset()
                gsw_PostEffects.Reset()
                gsw_Present.Reset()

                'and start our timer
                msw_FPS.Start()
            End If
            Return
        ElseIf mlID = -253 Then
            oSB.AppendLine("Camera Loc:")
            With goCamera
                oSB.AppendLine("  (" & .mlCameraX & ", " & .mlCameraY & ", " & .mlCameraZ & ")")
                oSB.AppendLine("  (" & .mlCameraAtX & ", " & .mlCameraAtY & ", " & .mlCameraAtZ & ")")
            End With
            If lblDetails.Caption <> oSB.ToString Then lblDetails.Caption = oSB.ToString
            Return
        End If

        If mlEntityIndex = -1 Then
            For X As Int32 = 0 To goCurrentEnvir.lEntityUB
                If goCurrentEnvir.lEntityIdx(X) = mlID AndAlso goCurrentEnvir.oEntity(X).ObjTypeID = miTypeID Then
                    mlEntityIndex = X
                    Exit For
                End If
            Next X
        End If

        '     If mlEntityIndex = -1 OrElse goCurrentEnvir.lEntityIdx(mlEntityIndex) = -1 Then
        'btnClose_Click(btnClose.ControlName)
        '         Return
        '     End If

        'Now, lets fill our details
        Try
            Dim oTmpEnvir As BaseEnvironment = goCurrentEnvir
            If oTmpEnvir Is Nothing = False Then
                If oTmpEnvir.lEntityUB >= mlEntityIndex Then
                    Dim oEntity As BaseEntity = oTmpEnvir.oEntity(mlEntityIndex)
                    With oEntity

                        If mlEntityID <> -1 AndAlso mlEntityID <> .ObjectID Then
                            For X As Int32 = 0 To oTmpEnvir.lEntityUB
                                If oTmpEnvir.lEntityIdx(X) = mlEntityID Then
                                    mlEntityIndex = X
                                    Return
                                End If
                            Next
                        End If

                        mlEntityID = .ObjectID

                        oSB.AppendLine("LocX: " & CInt(.LocX))
                        oSB.AppendLine("LocZ: " & CInt(.LocZ))
                        oSB.AppendLine("LocA: " & .LocAngle)
                        oSB.AppendLine("Yaw: " & .LocYaw)
                        oSB.AppendLine("DestX: " & .DestX)
                        oSB.AppendLine("DestZ: " & .DestZ)
                        oSB.AppendLine("VelX: " & CInt(.VelX))
                        oSB.AppendLine("velZ: " & CInt(.VelZ))
                        oSB.AppendLine("MaxSpeed: " & .oUnitDef.MaxSpeed.ToString)
                        oSB.AppendLine("Maneuver: " & .oUnitDef.Maneuver.ToString)
                        oSB.AppendLine("LocY: " & .LocY)
                        oSB.AppendLine("DestY: " & .mlTargetY)
                        If (.CurrentStatus And elUnitStatus.eUnitMoving) <> 0 Then oSB.AppendLine("Moving")

                        For X As Int32 = 0 To 3
                            If .lTargetIdx(X) > -1 Then
                                Dim oTmpEnt As BaseEntity = goCurrentEnvir.oEntity(.lTargetIdx(X))
                                If oTmpEnt Is Nothing = False Then
                                    oSB.AppendLine("T" & X & ": " & oTmpEnt.ObjectID)
                                End If
                            End If
                        Next X
                        oSB.AppendLine("TC: " & .lTargetMsg.ToString("#,##0"))

                        Dim sFinal As String = oSB.ToString

                        If lblDetails.Caption <> sFinal Then lblDetails.Caption = sFinal
                        sFinal = "ID: " & .ObjectID
                        If lblObsTitle.Caption <> sFinal Then lblObsTitle.Caption = sFinal
                    End With
                End If
            End If
        Catch
            MyBase.moUILib.RemoveWindow(Me.ControlName)
        End Try
    End Sub
    Private mlEntityID As Int32

    Private Sub frmObserve_WindowMoved() Handles Me.WindowMoved
        If mbLoading = False Then
            muSettings.ObserveY = Me.Top
            muSettings.ObserveX = Me.Left
        End If
    End Sub
End Class
