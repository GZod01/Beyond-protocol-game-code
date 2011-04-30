Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Interface created from Interface Builder
Public Class frmPlanetDetails
    Inherits UIWindow

    Private lblName As UILabel
    Private lnDiv1 As UILine

    'In order of appearance
    Private lblClass As UILabel
    Private lblSize As UILabel
    Private lblVeg As UILabel
    Private lblAtmosphere As UILabel
    Private lblHydrosphere As UILabel
    Private lblGravity As UILabel
    Private lblTemp As UILabel

    Private lblOwner As UILabel

    Private lblRingMin As UILabel
    Private lblRingMinConc As UILabel

    Private lblColonyLimits As UILabel

    Private mfSizeMult As Single = 0.0F

    Public PlanetID As Int32

    Private moPlanet As Planet = Nothing
    Private moSystem As SolarSystem = Nothing

    Private mbWormhole As Boolean = False
    Private mbSolarsystem As Boolean = False

    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)

        'frmPlanetDetails initial props
        With Me
            .lWindowMetricID = BPMetrics.eWindow.ePlanetDetails
            .ControlName = "frmPlanetDetails"
            .Left = 665
            .Top = 233
            .Width = 160
            .Height = 195
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = False
            .Moveable = True
            .BorderLineWidth = 1
            .ClickThru = True
        End With

        'lblName initial props
        lblName = New UILabel(oUILib)
        With lblName
            .ControlName = "lblName"
            .Left = 5
            .Top = 5
            .Width = 150
            .Height = 16
            .Enabled = True
            .Visible = True
            .Caption = "Gunarus Prime"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblName, UIControl))

        'lnDiv1 initial props
        lnDiv1 = New UILine(oUILib)
        With lnDiv1
            .ControlName = "lnDiv1"
            .Left = 0
            .Top = 25
            .Width = 160
            .Height = 0
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(lnDiv1, UIControl))

        'lblClass initial props
        lblClass = New UILabel(oUILib)
        With lblClass
            .ControlName = "lblClass"
            .Left = 5
            .Top = 30
            .Width = 150
            .Height = 16
            .Enabled = True
            .Visible = True
            .Caption = "Classification: Oceanic"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblClass, UIControl))

        'lblSize initial props
        lblSize = New UILabel(oUILib)
        With lblSize
            .ControlName = "lblSize"
            .Left = 5
            .Top = 50
            .Width = 150
            .Height = 16
            .Enabled = True
            .Visible = True
            .Caption = "Size Class: Medium"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblSize, UIControl))

        'lblVeg initial props
        lblVeg = New UILabel(oUILib)
        With lblVeg
            .ControlName = "lblVeg"
            .Left = 5
            .Top = 70
            .Width = 150
            .Height = 16
            .Enabled = True
            .Visible = True
            .Caption = "Vegetation: Lush"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblVeg, UIControl))

        'lblAtmosphere initial props
        lblAtmosphere = New UILabel(oUILib)
        With lblAtmosphere
            .ControlName = "lblAtmosphere"
            .Left = 5
            .Top = 90
            .Width = 150
            .Height = 16
            .Enabled = True
            .Visible = True
            .Caption = "Atmosphere: Breathable"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblAtmosphere, UIControl))

        'lblHydrosphere initial props
        lblHydrosphere = New UILabel(oUILib)
        With lblHydrosphere
            .ControlName = "lblHydrosphere"
            .Left = 5
            .Top = 110
            .Width = 150
            .Height = 16
            .Enabled = True
            .Visible = True
            .Caption = "Hydrosphere: Lava"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblHydrosphere, UIControl))

        'lblGravity initial props
        lblGravity = New UILabel(oUILib)
        With lblGravity
            .ControlName = "lblGravity"
            .Left = 5
            .Top = 130
            .Width = 150
            .Height = 16
            .Enabled = True
            .Visible = True
            .Caption = "Gravity: 9.8 m/s"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblGravity, UIControl))

        'lblTemp initial props
        lblTemp = New UILabel(oUILib)
        With lblTemp
            .ControlName = "lblTemp"
            .Left = 5
            .Top = 150
            .Width = 150
            .Height = 16
            .Enabled = True
            .Visible = True
            .Caption = "Avg. Temperature: 300K"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblTemp, UIControl))

        'lblOwner initial props
        lblOwner = New UILabel(oUILib)
        With lblOwner
            .ControlName = "lblOwner"
            .Left = 5
            .Top = 170
            .Width = 150
            .Height = 16
            .Enabled = True
            .Visible = True
            .Caption = "Owner: None"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblOwner, UIControl))

        Dim bRings As Boolean = False
        If goCurrentPlayer Is Nothing = False Then
            'indicates they have orbital mining platforms
            Dim bVis As Boolean = False
            bVis = (goCurrentPlayer.GetSpecialTechPropertyValue(PlayerSpecialAttributeID.eSuperSpecials) And 16) <> 0

            Me.Height = 230
            'lblRingMin initial props
            lblRingMin = New UILabel(oUILib)
            With lblRingMin
                .ControlName = "lblRingMin"
                .Left = 5
                .Top = 190
                .Width = 150
                .Height = 16
                .Enabled = True
                .Visible = bVis
                .Caption = "Rings: None"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblRingMin, UIControl))

            'lblRingMinConc initial props
            lblRingMinConc = New UILabel(oUILib)
            With lblRingMinConc
                .ControlName = "lblRingMinConc"
                .Left = 5
                .Top = 210
                .Width = 150
                .Height = 16
                .Enabled = True
                .Visible = bVis
                .Caption = "Mining Rate: None"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblRingMinConc, UIControl))
            bRings = True
        End If

        'lblColonyLimits initial props
        lblColonyLimits = New UILabel(oUILib)
        With lblColonyLimits
            .ControlName = "lblColonyLimits"
            .Left = 5
            If bRings Then
                .Top = lblRingMinConc.Top + 20
            Else
                .Top = lblOwner.Top + 20
            End If
            .Width = 150
            .Height = 16
            .Enabled = True
            .Visible = True
            .Caption = "Colony Count : 0 / 0"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblColonyLimits, UIControl))


        MyBase.moUILib.RemoveWindow(Me.ControlName)
        MyBase.moUILib.AddWindow(Me)
    End Sub

    Public Sub SetFromPlanet(ByRef oPlanet As Planet, ByVal lX As Int32, ByVal lY As Int32)
        mfSizeMult = 0.0F
        Me.Left = lX
        Me.Top = lY
        mbWormhole = False
        Me.PlanetID = oPlanet.ObjectID

        moPlanet = oPlanet
        Me.Height = 210

        With oPlanet
            If .Atmosphere > 0 Then Me.lblAtmosphere.Caption = "Atmosphere: Breathable" Else Me.lblAtmosphere.Caption = "Atmosphere: None"
            Select Case .MapTypeID
                Case PlanetType.eAcidic
                    lblClass.Caption = "Classification: Acidic"
                    lblHydrosphere.Caption = "Hydrosphere: Acid"
                    Me.lblVeg.Caption = "Vegetation: Hostile"
                Case PlanetType.eAdaptable
                    lblClass.Caption = "Classification: Adaptable"
                    lblHydrosphere.Caption = "Hydrosphere: Water"
                    Me.lblVeg.Caption = "Vegetation: Fungi"
                Case PlanetType.eBarren
                    lblClass.Caption = "Classification: Barren"
                    lblHydrosphere.Caption = "Hydrosphere: None"
                    Me.lblVeg.Caption = "Vegetation: None"
                Case PlanetType.eDesert
                    lblClass.Caption = "Classification: Desert"
                    lblHydrosphere.Caption = "Hydrosphere: Water"
                    Me.lblVeg.Caption = "Vegetation: Sparse"
                Case PlanetType.eGasGiant
                    lblClass.Caption = "Classification: Gas Giant"
                    lblHydrosphere.Caption = "Hydrosphere: Gas"
                    Me.lblVeg.Caption = "Vegetation: None"
                Case PlanetType.eGeoPlastic
                    lblClass.Caption = "Classification: Lava"
                    lblHydrosphere.Caption = "Hydrosphere: Lava"
                    Me.lblVeg.Caption = "Vegetation: None"
                Case PlanetType.eSuperGiant
                    lblClass.Caption = "Classification: Super Giant"
                    lblHydrosphere.Caption = "Hydrosphere: Gas"
                    Me.lblVeg.Caption = "Vegetation: None"
                Case PlanetType.eTerran
                    lblClass.Caption = "Classification: Terran"
                    lblHydrosphere.Caption = "Hydrosphere: Water"
                    Me.lblVeg.Caption = "Vegetation: Lush"
                Case PlanetType.eTundra
                    lblClass.Caption = "Classification: Frozen"
                    lblHydrosphere.Caption = "Hydrosphere: Water"
                    Me.lblVeg.Caption = "Vegetation: None"
                Case PlanetType.eWaterWorld
                    lblClass.Caption = "Classification: Oceanic"
                    lblHydrosphere.Caption = "Hydrosphere: Water"
                    Me.lblVeg.Caption = "Vegetation: Tropical"
            End Select

            Me.lblGravity.Caption = "Gravity: " & (CInt(.Gravity) / 10.0F).ToString("###.0") & " m/s"
            Me.lblName.Caption = .PlanetName
            Select Case .PlanetSizeID
                Case 0
                    Me.lblSize.Caption = "Size Class: Tiny"
                Case 1
                    Me.lblSize.Caption = "Size Class: Small"
                Case 2
                    Me.lblSize.Caption = "Size Class: Medium"
                Case 3
                    Me.lblSize.Caption = "Size Class: Large"
                Case 4
                    Me.lblSize.Caption = "Size Class: Huge"
            End Select

            Me.lblTemp.Caption = "Avg. Temperature: " & (CInt(.SurfaceTemperature) * 10) & "K"

            If goCurrentPlayer Is Nothing = False AndAlso (goCurrentPlayer.GetSpecialTechPropertyValue(PlayerSpecialAttributeID.eSuperSpecials) And 16) <> 0 Then
                If lblRingMin Is Nothing = False Then
                    Dim sTemp As String = "Rings: None"
                    If .RingMineralID > 0 Then
                        sTemp = "Rings: Unknown Mineral"
                        For X As Int32 = 0 To glMineralUB
                            If glMineralIdx(X) = .RingMineralID Then
                                sTemp = "Rings: " & goMinerals(X).MineralName
                                Exit For
                            End If
                        Next X
                        Me.Height = 230
                    Else
                        Me.Height = 210
                        Me.lblColonyLimits.Top = Me.lblRingMinConc.Top
                    End If
                    lblRingMin.Caption = sTemp
                End If
                If lblRingMinConc Is Nothing = False Then
                    Dim sTemp As String = ""
                    If .RingMineralConcentration > 0 Then
                        sTemp = "Mining Rate: " & .RingMineralConcentration
                    End If
                    lblRingMinConc.Caption = sTemp

                End If
            Else
                Me.Height = 190
                Me.lblColonyLimits.Top = Me.lblRingMin.Top
            End If

        End With
    End Sub

    Public Sub SetFromWormhole(ByVal oWormhole As Wormhole, ByVal oCurSys As SolarSystem, ByVal lX As Int32, ByVal lY As Int32)
        mfSizeMult = 0.0F
        Me.Left = lX
        Me.Top = lY
        Me.Height = 70
        Me.PlanetID = oWormhole.ObjectID
        mbWormhole = True

        With oWormhole

            If (goCurrentPlayer.GetSpecialTechPropertyValue(PlayerSpecialAttributeID.eSuperSpecials) And 4) <> 0 Then
                lblName.Caption = "Wormhole"
                'lblSize.Caption = "Stability: " & ((1.0F - (.StartCycle / Math.Max(1, .EndCycle))) * 100).ToString("###.##") & "%"
                Select Case (.WormholeFlags And (elWormholeFlag.eSystem1Jumpable Or elWormholeFlag.eSystem2Jumpable))
                    Case elWormholeFlag.eSystem1Jumpable
                        If .System1 Is Nothing = False AndAlso oCurSys.ObjectID = .System1.ObjectID Then
                            lblClass.Caption = "One-way jump from this side"
                        Else
                            lblClass.Caption = "Stability is questionable"
                        End If
                    Case elWormholeFlag.eSystem2Jumpable
                        If .System2 Is Nothing = False AndAlso oCurSys.ObjectID = .System2.ObjectID Then
                            lblClass.Caption = "One-way jump from this side"
                        Else
                            lblClass.Caption = "Stability is questionable"
                        End If
                    Case 0
                        lblClass.Caption = "Stability is questionable"
                    Case Else
                        lblClass.Caption = "Two-way traversable"
                End Select

                Dim oOtherSys As SolarSystem = Nothing
                If .System1 Is Nothing = False Then
                    If .System1.ObjectID = oCurSys.ObjectID Then
                        oOtherSys = .System2
                    Else : oOtherSys = .System1
                    End If
                ElseIf .System2 Is Nothing = False Then
                    If .System2.ObjectID = oCurSys.ObjectID Then
                        oOtherSys = .System1
                    Else : oOtherSys = .System2
                    End If
                End If

                If oOtherSys Is Nothing = False Then
                    lblSize.Caption = "Destination: " & oOtherSys.SystemName
                Else
                    lblSize.Caption = "Destination: Unknown"
                End If
                lblVeg.Caption = ""
                lblAtmosphere.Caption = ""
                lblHydrosphere.Caption = ""
                lblGravity.Caption = ""
                lblTemp.Caption = ""
            Else
                lblClass.Caption = ""
                lblClass.Caption = "Spatial Anomaly"
                lblName.Caption = ""
                lblSize.Caption = "Our scientists have"
                lblVeg.Caption = "to study this anomaly"
                lblAtmosphere.Caption = "further to determine"
                lblHydrosphere.Caption = "any details related to it"
                lblGravity.Caption = ""
                lblTemp.Caption = ""
            End If

        End With
    End Sub

    Public Sub SetFromStar(ByVal oOtherSys As SolarSystem, ByVal lX As Int32, ByVal lY As Int32)
        If oOtherSys Is Nothing Then Return

        mfSizeMult = 0.0F
        Me.Left = lX
        Me.Top = lY

        Me.PlanetID = oOtherSys.ObjectID
        mbSolarsystem = True

        'Clear things up
        lblClass.Caption = ""
        lblSize.Caption = ""
        lblVeg.Caption = ""
        lblAtmosphere.Caption = ""
        lblHydrosphere.Caption = ""
        lblGravity.Caption = ""
        lblTemp.Caption = ""
        'Me.Height = 50
        lblOwner.Caption = ""
        lblRingMin.Caption = ""
        lblRingMinConc.Caption = ""

        lblName.Caption = oOtherSys.SystemName
        moSystem = oOtherSys
        oOtherSys.RequestGalMapViewDetails()
        'If goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).moPlanets Is Nothing Then
        '    lblClass.Caption = "Planets : 0"
        'Else
        '    lblClass.Caption = "Planets : " & UBound(goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).moPlanets).ToString
        'End If
        'If goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).moWormholes Is Nothing Then
        '    lblSize.Caption = "Known Wormholes : 0"
        'Else
        '    lblSize.Caption = "Known Wormholes : " & (UBound(goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).moWormholes) + 1).ToString
        'End If
        If glCurrentEnvirView = CurrentView.eGalaxyMapView Then
            If goGalaxy.GalaxySelectionIdx = -1 Then Return
            Dim oCurSys As SolarSystem = goGalaxy.moSystems(goGalaxy.GalaxySelectionIdx)
            If oCurSys Is Nothing Then Return
            If oCurSys.ObjectID <> oOtherSys.ObjectID Then
                Dim fDistX As Single = oOtherSys.LocX - oCurSys.LocX
                Dim fDistY As Single = oOtherSys.LocY - oCurSys.LocY
                Dim fDistZ As Single = oOtherSys.LocZ - oCurSys.LocZ
                Dim dDist As Double
                fDistX *= fDistX
                fDistY *= fDistY
                fDistZ *= fDistZ
                dDist = Math.Sqrt(fDistX + fDistY + fDistZ) * 53.7889544
                dDist *= 0.5F
                dDist *= 100

                lblClass.Caption = "Distance : " & dDist.ToString("#,###.00") & " AU"
            End If
        End If

    End Sub

    Private Sub frmPlanetDetails_OnNewFrame() Handles Me.OnNewFrame
        If mfSizeMult <> 1.0F Then
            mfSizeMult += 0.05F
            If mfSizeMult > 1.0F Then mfSizeMult = 1.0F

            Dim lVal As Int32 = CInt((Me.Width - 10) * mfSizeMult)


            lblName.Width = lVal
            lblSize.Width = lVal
            lblVeg.Width = lVal
            lblAtmosphere.Width = lVal
            lblHydrosphere.Width = lVal
            lblGravity.Width = lVal
            lblTemp.Width = lVal
            lblClass.Width = lVal
            lblOwner.Width = lVal
            lblColonyLimits.Width = lVal

            If lblRingMin Is Nothing = False Then lblRingMin.Width = lVal
            If lblRingMinConc Is Nothing = False Then lblRingMinConc.Width = lVal
        End If

        If glCurrentEnvirView <> CurrentView.eSystemMapView1 AndAlso glCurrentEnvirView <> CurrentView.eSystemMapView2 AndAlso (glCurrentEnvirView <> CurrentView.eSystemView OrElse mbWormhole = False) AndAlso (glCurrentEnvirView <> CurrentView.eGalaxyMapView OrElse mbSolarsystem = False) Then
            MyBase.moUILib.RemoveWindow(Me.ControlName)
        End If

        If moPlanet Is Nothing = False AndAlso mbWormhole = False Then
            Dim lID As Int32 = moPlanet.OwnerID
            Dim sTemp As String
            If lID > 0 Then
                sTemp = "Owner: " & GetCacheObjectValue(lID, ObjectType.ePlayer)
            Else
                sTemp = "Owner: None"
            End If
            If lblOwner.Caption <> sTemp Then lblOwner.Caption = sTemp

            If goCurrentPlayer Is Nothing = False AndAlso (goCurrentPlayer.GetSpecialTechPropertyValue(PlayerSpecialAttributeID.eSuperSpecials) And 16) <> 0 Then
                If lblRingMin Is Nothing = False Then
                    sTemp = "Rings: None"
                    If moPlanet.RingMineralID > 0 Then
                        sTemp = "Rings: Unknown Mineral"
                        For X As Int32 = 0 To glMineralUB
                            If glMineralIdx(X) = moPlanet.RingMineralID Then
                                sTemp = "Rings: " & goMinerals(X).MineralName
                                Exit For
                            End If
                        Next X
                    End If
                    If lblRingMin.Caption <> sTemp Then lblRingMin.Caption = sTemp
                End If
                If lblRingMinConc Is Nothing = False Then
                    sTemp = ""
                    If moPlanet.RingMineralConcentration > 0 Then
                        sTemp = "Mining Rate: " & moPlanet.RingMineralConcentration
                    End If
                    If lblRingMinConc.Caption <> sTemp Then lblRingMinConc.Caption = sTemp
                End If
            End If
            Dim lColonies As Int32 = moPlanet.ColonyCount
            If lColonies > -1 Then
                sTemp = "Colony Count : " & lColonies.ToString & " of " & CInt(moPlanet.PlanetSizeID) + 2
                If Me.Height < Me.lblColonyLimits.Top + 20 Then
                    Me.Height = Me.Height + 20
                End If
            Else
                sTemp = ""
            End If
            If Me.lblColonyLimits.Caption <> sTemp Then Me.lblColonyLimits.Caption = sTemp

        ElseIf mbSolarsystem = True AndAlso moSystem Is Nothing = False Then
            Dim sTemp As String = "Planets: " & moSystem.yPlanetCnt.ToString("#,##0")
            If sTemp <> lblSize.Caption Then lblSize.Caption = sTemp

            If lblVeg.Caption <> "VOTERS:" Then lblVeg.Caption = "VOTERS:"
            Dim lHt As Int32 = lblVeg.Top + lblVeg.Height + 5
            If moSystem.lVoterUB > -1 Then
                For X As Int32 = 0 To moSystem.lVoterUB
                    sTemp = GetCacheObjectValue(moSystem.lVoterIDs(X), ObjectType.ePlayer) & ": " & moSystem.yVoterCnts(X).ToString("#,##0")

                    Select Case X
                        Case 0
                            If lblAtmosphere.Caption <> sTemp Then lblAtmosphere.Caption = sTemp
                            lHt = lblAtmosphere.Top + lblAtmosphere.Height + 5
                        Case 1
                            If lblHydrosphere.Caption <> sTemp Then lblHydrosphere.Caption = sTemp
                            lHt = lblHydrosphere.Top + lblHydrosphere.Height + 5
                        Case 2
                            If lblGravity.Caption <> sTemp Then lblGravity.Caption = sTemp
                            lHt = lblGravity.Top + lblGravity.Height + 5
                        Case 3
                            If lblTemp.Caption <> sTemp Then lblTemp.Caption = sTemp
                            lHt = lblTemp.Top + lblTemp.Height + 5
                        Case 4
                            If lblOwner.Caption <> sTemp Then lblOwner.Caption = sTemp
                            lHt = lblOwner.Top + lblOwner.Height + 5
                        Case 5
                            If lblRingMin.Caption <> sTemp Then lblRingMin.Caption = sTemp
                            lHt = lblRingMin.Top + 5
                        Case Else
                            lHt = lblRingMinConc.Top + lblRingMinConc.Height + 5
                            If moSystem.lVoterUB = X Then
                                If lblRingMinConc.Caption <> sTemp Then lblRingMinConc.Caption = sTemp
                            Else
                                sTemp = (moSystem.lVoterUB - 5).ToString & " more..."
                                If lblRingMinConc.Caption <> sTemp Then lblRingMinConc.Caption = sTemp
                                Exit For
                            End If
                    End Select

                Next X
            End If
            If lHt <> Me.Height Then lHt = Me.Height
        Else
            Dim sTemp As String = ""
            If lblOwner.Caption <> sTemp Then lblOwner.Caption = sTemp
        End If

        If mbWormhole = True Then
            If lblRingMin Is Nothing = False Then
                If lblRingMin.Visible <> False Then lblRingMin.Visible = False
            End If
            If lblRingMinConc Is Nothing = False Then
                If lblRingMinConc.Visible <> False Then lblRingMinConc.Visible = False
            End If
        End If
    End Sub

    Public Sub UpdateLoc(ByVal lX As Int32, ByVal lY As Int32)
        Me.Left = lX
        Me.Top = lY
    End Sub
End Class