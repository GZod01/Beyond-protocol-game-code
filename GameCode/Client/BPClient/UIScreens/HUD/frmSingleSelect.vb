Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

''Selection of a single entity display
'Public Class frmSingleSelect_Old
'    Inherits UIWindow

'    Private lblEntityName As UILabel
'    Private lblEngineStatus As UILabel
'    Private lblRadarStatus As UILabel
'    Private lblShieldStatus As UILabel
'    Private lblHangarStatus As UILabel
'    Private lblCargoStatus As UILabel
'    Private lblForeWpn As UILabel
'    Private lblLeftWpns As UILabel
'    Private lblRearWpns As UILabel
'    Private lblRightWpns As UILabel
'    Private lblForeWpn1Stat As UILabel
'    Private lblForeWpn2Stat As UILabel
'    Private lblLeftWpn1Stat As UILabel
'    Private lblLeftWpn2Stat As UILabel
'    Private lblRearWpn1Stat As UILabel
'    Private lblRearWpn2Stat As UILabel
'    Private lblRightWpn1Stat As UILabel
'    Private lblRightWpn2Stat As UILabel

'    Private lblFuelStatus As UILabel
'    Private lblFuelAmmo As UILabel
'    Private lblFuelCap As UILabel
'    Private lblAmmoCap As UILabel

'    Private lblETA As UILabel

'    'Private shpProgressBack As UIWindow
'    'Private shpProgressFore As UIWindow
'    Private hln1 As UILine
'    Private lblExpLevel As UILabel

'    Private mlEntityIndex As Int32 = -1
'    Private mbForceUpdate As Boolean

'    Private mlLastHPUpdateChange As Int32 = Int32.MinValue

'    Private mlLastMovementTest As Int32

'    Public Sub New(ByRef oUILib As UILib)
'        MyBase.New(oUILib)

'        'frmSingleSelect initial props
'        With Me
'            .ControlName = "frmSingleSelect"
'            .Left = 0
'            .Top = moUILib.oDevice.PresentationParameters.BackBufferHeight - 86
'            .Width = 190
'            .Height = 86
'            .Enabled = True
'            .Visible = True
'            '.BorderColor = System.Drawing.Color.FromArgb(-16711681)
'            .BorderColor = muSettings.InterfaceBorderColor
'            '.FillColor = System.Drawing.Color.FromArgb(-16768960)
'            '.FillColor = System.Drawing.Color.FromArgb(128, 64, 128, 192)
'            .FillColor = muSettings.InterfaceFillColor
'            .FullScreen = False
'            .BorderLineWidth = 2
'        End With
'        Debug.Write(Me.ControlName & " Newed" & vbCrLf)

'        'lblEntityName initial props
'        lblEntityName = New UILabel(oUILib)
'        With lblEntityName
'            .ControlName = "lblEntityName"
'            .Left = 5
'            .Top = 2
'            .Width = 150
'            .Height = 16
'            .Enabled = True
'            .Visible = True
'            .Caption = "ENTITY NAME"
'            '.ForeColor = System.Drawing.Color.FromArgb(-16711681)
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Arial", 9, FontStyle.Bold, GraphicsUnit.Point, 0))
'            .DrawBackImage = False
'            .FontFormat = DrawTextFormat.VerticalCenter
'        End With
'        Me.AddChild(CType(lblEntityName, UIControl))

'        'lblEngineStatus initial props
'        lblEngineStatus = New UILabel(oUILib)
'        With lblEngineStatus
'            .ControlName = "lblEngineStatus"
'            .Left = 6
'            .Top = 24
'            .Width = 60
'            .Height = 10
'            .Enabled = True
'            .Visible = True
'            .Caption = "ENGINES"
'            .ForeColor = System.Drawing.Color.FromArgb(-16711936)
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 6, FontStyle.Bold, GraphicsUnit.Point, 0))
'            .DrawBackImage = False
'            .FontFormat = DrawTextFormat.VerticalCenter
'        End With
'        Me.AddChild(CType(lblEngineStatus, UIControl))

'        'lblRadarStatus initial props
'        lblRadarStatus = New UILabel(oUILib)
'        With lblRadarStatus
'            .ControlName = "lblRadarStatus"
'            .Left = 6
'            .Top = 34
'            .Width = 60
'            .Height = 10
'            .Enabled = True
'            .Visible = True
'            .Caption = "RADAR"
'            .ForeColor = System.Drawing.Color.FromArgb(-16711936)
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 6, FontStyle.Bold, GraphicsUnit.Point, 0))
'            .DrawBackImage = False
'            .FontFormat = DrawTextFormat.VerticalCenter
'        End With
'        Me.AddChild(CType(lblRadarStatus, UIControl))

'        'lblShieldStatus initial props
'        lblShieldStatus = New UILabel(oUILib)
'        With lblShieldStatus
'            .ControlName = "lblShieldStatus"
'            .Left = 6
'            .Top = 44
'            .Width = 60
'            .Height = 10
'            .Enabled = True
'            .Visible = True
'            .Caption = "SHIELDS"
'            .ForeColor = System.Drawing.Color.FromArgb(-16711936)
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 6, FontStyle.Bold, GraphicsUnit.Point, 0))
'            .DrawBackImage = False
'            .FontFormat = DrawTextFormat.VerticalCenter
'        End With
'        Me.AddChild(CType(lblShieldStatus, UIControl))

'        'lblHangarStatus initial props
'        lblHangarStatus = New UILabel(oUILib)
'        With lblHangarStatus
'            .ControlName = "lblHangarStatus"
'            .Left = 6
'            .Top = 54
'            .Width = 60
'            .Height = 10
'            .Enabled = True
'            .Visible = True
'            .Caption = "HANGAR"
'            .ForeColor = System.Drawing.Color.FromArgb(-8355712)
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 6, FontStyle.Bold, GraphicsUnit.Point, 0))
'            .DrawBackImage = False
'            .FontFormat = DrawTextFormat.VerticalCenter
'        End With
'        Me.AddChild(CType(lblHangarStatus, UIControl))

'        'lblCargoStatus initial props
'        lblCargoStatus = New UILabel(oUILib)
'        With lblCargoStatus
'            .ControlName = "lblCargoStatus"
'            .Left = 6
'            .Top = 64
'            .Width = 60
'            .Height = 10
'            .Enabled = True
'            .Visible = True
'            .Caption = "CARGO BAY"
'            .ForeColor = System.Drawing.Color.FromArgb(-8355712)
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 6, FontStyle.Bold, GraphicsUnit.Point, 0))
'            .DrawBackImage = False
'            .FontFormat = DrawTextFormat.VerticalCenter
'        End With
'        Me.AddChild(CType(lblCargoStatus, UIControl))

'        'lblFuelStatus initial props
'        lblFuelStatus = New UILabel(oUILib)
'        With lblFuelStatus
'            .ControlName = "lblFuelStatus"
'            .Left = 6
'            .Top = 74
'            .Width = 60
'            .Height = 10
'            .Enabled = True
'            .Visible = True
'            .Caption = "FUEL BAY"
'            .ForeColor = System.Drawing.Color.FromArgb(-16711936)
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 6, FontStyle.Bold, GraphicsUnit.Point, 0))
'            .DrawBackImage = False
'            .FontFormat = DrawTextFormat.VerticalCenter
'        End With
'        Me.AddChild(CType(lblFuelStatus, UIControl))

'        'lblForeWpn initial props
'        lblForeWpn = New UILabel(oUILib)
'        With lblForeWpn
'            .ControlName = "lblForeWpn"
'            .Left = 67
'            .Top = 24
'            .Width = 95
'            .Height = 10
'            .Enabled = True
'            .Visible = True
'            .Caption = "FORWARD WPNS    /"
'            '.ForeColor = System.Drawing.Color.FromArgb(-16711681)
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 6, FontStyle.Bold, GraphicsUnit.Point, 0))
'            .DrawBackImage = False
'            .FontFormat = DrawTextFormat.VerticalCenter
'        End With
'        Me.AddChild(CType(lblForeWpn, UIControl))

'        'lblLeftWpns initial props
'        lblLeftWpns = New UILabel(oUILib)
'        With lblLeftWpns
'            .ControlName = "lblLeftWpns"
'            .Left = 67
'            .Top = 34
'            .Width = 95
'            .Height = 10
'            .Enabled = True
'            .Visible = True
'            .Caption = "LEFT WPNS    /"
'            '.ForeColor = System.Drawing.Color.FromArgb(-16711681)
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 6, FontStyle.Bold, GraphicsUnit.Point, 0))
'            .DrawBackImage = False
'            .FontFormat = DrawTextFormat.VerticalCenter
'        End With
'        Me.AddChild(CType(lblLeftWpns, UIControl))

'        'lblRearWpns initial props
'        lblRearWpns = New UILabel(oUILib)
'        With lblRearWpns
'            .ControlName = "lblRearWpns"
'            .Left = 67
'            .Top = 44
'            .Width = 95
'            .Height = 10
'            .Enabled = True
'            .Visible = True
'            .Caption = "REAR WPNS    /"
'            '.ForeColor = System.Drawing.Color.FromArgb(-16711681)
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 6, FontStyle.Bold, GraphicsUnit.Point, 0))
'            .DrawBackImage = False
'            .FontFormat = DrawTextFormat.VerticalCenter
'        End With
'        Me.AddChild(CType(lblRearWpns, UIControl))

'        'lblRightWpns initial props
'        lblRightWpns = New UILabel(oUILib)
'        With lblRightWpns
'            .ControlName = "lblRightWpns"
'            .Left = 67
'            .Top = 54
'            .Width = 95
'            .Height = 10
'            .Enabled = True
'            .Visible = True
'            .Caption = "RIGHT WPNS    /"
'            '.ForeColor = System.Drawing.Color.FromArgb(-16711681)
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 6, FontStyle.Bold, GraphicsUnit.Point, 0))
'            .DrawBackImage = False
'            .FontFormat = DrawTextFormat.VerticalCenter
'        End With
'        Me.AddChild(CType(lblRightWpns, UIControl))

'        'lblFuelAmmo initial props
'        lblFuelAmmo = New UILabel(oUILib)
'        With lblFuelAmmo
'            .ControlName = "lblFuelAmmo"
'            .Left = 67
'            .Top = 64
'            .Width = 95
'            .Height = 10
'            .Enabled = True
'            .Visible = True
'            .Caption = "FUEL / AMMO          /"
'            '.ForeColor = System.Drawing.Color.FromArgb(-16711681)
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 6, FontStyle.Bold, GraphicsUnit.Point, 0))
'            .DrawBackImage = False
'            .FontFormat = DrawTextFormat.VerticalCenter
'        End With
'        Me.AddChild(CType(lblFuelAmmo, UIControl))

'        'lblETA initial props
'        lblETA = New UILabel(oUILib)
'        With lblETA
'            .ControlName = "lblETA"
'            .Left = 67
'            .Top = 74
'            .Width = 95
'            .Height = 10
'            .Enabled = True
'            .Visible = True
'            .Caption = "ETA 00:00:00:00"
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 6, FontStyle.Bold, GraphicsUnit.Point, 0))
'            .DrawBackImage = False
'            .FontFormat = DrawTextFormat.VerticalCenter
'            .ToolTipText = "The ETA to the current waypoint."
'        End With
'        Me.AddChild(CType(lblETA, UIControl))

'        'lblForeWpn1Stat initial props
'        lblForeWpn1Stat = New UILabel(oUILib)
'        With lblForeWpn1Stat
'            .ControlName = "lblForeWpn1Stat"
'            .Left = 152
'            .Top = 24
'            .Width = 10
'            .Height = 10
'            .Enabled = True
'            .Visible = True
'            .Caption = "1"
'            .ForeColor = System.Drawing.Color.FromArgb(-16711936)
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 6, FontStyle.Bold, GraphicsUnit.Point, 0))
'            .DrawBackImage = False
'            .FontFormat = DrawTextFormat.VerticalCenter
'        End With
'        Me.AddChild(CType(lblForeWpn1Stat, UIControl))

'        'lblForeWpn2Stat initial props
'        lblForeWpn2Stat = New UILabel(oUILib)
'        With lblForeWpn2Stat
'            .ControlName = "lblForeWpn2Stat"
'            .Left = 165
'            .Top = 24
'            .Width = 10
'            .Height = 10
'            .Enabled = True
'            .Visible = True
'            .Caption = "2"
'            .ForeColor = System.Drawing.Color.FromArgb(-8355712)
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 6, FontStyle.Bold, GraphicsUnit.Point, 0))
'            .DrawBackImage = False
'            .FontFormat = DrawTextFormat.VerticalCenter
'        End With
'        Me.AddChild(CType(lblForeWpn2Stat, UIControl))

'        'lblLeftWpn1Stat initial props
'        lblLeftWpn1Stat = New UILabel(oUILib)
'        With lblLeftWpn1Stat
'            .ControlName = "lblLeftWpn1Stat"
'            .Left = 127
'            .Top = 34
'            .Width = 10
'            .Height = 10
'            .Enabled = True
'            .Visible = True
'            .Caption = "1"
'            .ForeColor = System.Drawing.Color.FromArgb(-8355712)
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 6, FontStyle.Bold, GraphicsUnit.Point, 0))
'            .DrawBackImage = False
'            .FontFormat = DrawTextFormat.VerticalCenter
'        End With
'        Me.AddChild(CType(lblLeftWpn1Stat, UIControl))

'        'lblLeftWpn2Stat initial props
'        lblLeftWpn2Stat = New UILabel(oUILib)
'        With lblLeftWpn2Stat
'            .ControlName = "lblLeftWpn2Stat"
'            .Left = 139
'            .Top = 34
'            .Width = 10
'            .Height = 10
'            .Enabled = True
'            .Visible = True
'            .Caption = "2"
'            .ForeColor = System.Drawing.Color.FromArgb(-8355712)
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 6, FontStyle.Bold, GraphicsUnit.Point, 0))
'            .DrawBackImage = False
'            .FontFormat = DrawTextFormat.VerticalCenter
'        End With
'        Me.AddChild(CType(lblLeftWpn2Stat, UIControl))

'        'lblRearWpn1Stat initial props
'        lblRearWpn1Stat = New UILabel(oUILib)
'        With lblRearWpn1Stat
'            .ControlName = "lblRearWpn1Stat"
'            .Left = 130
'            .Top = 44
'            .Width = 10
'            .Height = 10
'            .Enabled = True
'            .Visible = True
'            .Caption = "1"
'            .ForeColor = System.Drawing.Color.FromArgb(-8355712)
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 6, FontStyle.Bold, GraphicsUnit.Point, 0))
'            .DrawBackImage = False
'            .FontFormat = DrawTextFormat.VerticalCenter
'        End With
'        Me.AddChild(CType(lblRearWpn1Stat, UIControl))

'        'lblRearWpn2Stat initial props
'        lblRearWpn2Stat = New UILabel(oUILib)
'        With lblRearWpn2Stat
'            .ControlName = "lblRearWpn2Stat"
'            .Left = 142
'            .Top = 44
'            .Width = 10
'            .Height = 10
'            .Enabled = True
'            .Visible = True
'            .Caption = "2"
'            .ForeColor = System.Drawing.Color.FromArgb(-8355712)
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 6, FontStyle.Bold, GraphicsUnit.Point, 0))
'            .DrawBackImage = False
'            .FontFormat = DrawTextFormat.VerticalCenter
'        End With
'        Me.AddChild(CType(lblRearWpn2Stat, UIControl))

'        'lblRightWpn1Stat initial props
'        lblRightWpn1Stat = New UILabel(oUILib)
'        With lblRightWpn1Stat
'            .ControlName = "lblRightWpn1Stat"
'            .Left = 133
'            .Top = 54
'            .Width = 10
'            .Height = 10
'            .Enabled = True
'            .Visible = True
'            .Caption = "1"
'            .ForeColor = System.Drawing.Color.FromArgb(-8355712)
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 6, FontStyle.Bold, GraphicsUnit.Point, 0))
'            .DrawBackImage = False
'            .FontFormat = DrawTextFormat.VerticalCenter
'        End With
'        Me.AddChild(CType(lblRightWpn1Stat, UIControl))

'        'lblRightWpn2Stat initial props
'        lblRightWpn2Stat = New UILabel(oUILib)
'        With lblRightWpn2Stat
'            .ControlName = "lblRightWpn2Stat"
'            .Left = 146
'            .Top = 54
'            .Width = 10
'            .Height = 10
'            .Enabled = True
'            .Visible = True
'            .Caption = "2"
'            .ForeColor = System.Drawing.Color.FromArgb(-8355712)
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 6, FontStyle.Bold, GraphicsUnit.Point, 0))
'            .DrawBackImage = False
'            .FontFormat = DrawTextFormat.VerticalCenter
'        End With
'        Me.AddChild(CType(lblRightWpn2Stat, UIControl))

'        'lblFuelCap initial props
'        lblFuelCap = New UILabel(oUILib)
'        With lblFuelCap
'            .ControlName = "lblFuelCap"
'            .Left = 134
'            .Top = 64
'            .Width = 22
'            .Height = 10
'            .Enabled = True
'            .Visible = True
'            .Caption = "100%"
'            .ForeColor = System.Drawing.Color.FromArgb(-16711936)
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 6, FontStyle.Bold, GraphicsUnit.Point, 0))
'            .DrawBackImage = False
'            .FontFormat = DrawTextFormat.VerticalCenter
'        End With
'        Me.AddChild(CType(lblFuelCap, UIControl))

'        'lblAmmoCap initial props
'        lblAmmoCap = New UILabel(oUILib)
'        With lblAmmoCap
'            .ControlName = "lblAmmoCap"
'            .Left = 164
'            .Top = 64
'            .Width = 22
'            .Height = 10
'            .Enabled = True
'            .Visible = True
'            .Caption = "100%"
'            .ForeColor = System.Drawing.Color.FromArgb(-16711936)
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 6, FontStyle.Bold, GraphicsUnit.Point, 0))
'            .DrawBackImage = False
'            .FontFormat = DrawTextFormat.VerticalCenter
'        End With
'        Me.AddChild(CType(lblAmmoCap, UIControl))

'        'vln2 initial props
'        hln1 = New UILine(oUILib)
'        With hln1
'            .ControlName = "hln1"
'            .Left = 1
'            .Top = 19
'            .Width = 189
'            .Height = 0
'            .Enabled = True
'            .Visible = True
'            '.BorderColor = System.Drawing.Color.FromArgb(-16711681)
'            .BorderColor = muSettings.InterfaceBorderColor
'        End With
'        Me.AddChild(CType(hln1, UIControl))

'        'lblExpLevel initial props
'        lblExpLevel = New UILabel(oUILib)
'        With lblExpLevel
'            .ControlName = "lblExpLevel"
'            .Left = Me.Width - 45
'            .Top = 3
'            .Width = 40
'            .Height = 16
'            .Enabled = True
'            .Visible = True
'            .Caption = "G"
'            .ForeColor = System.Drawing.Color.FromArgb(255, 0, 255, 0)
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 12, FontStyle.Bold, GraphicsUnit.Point, 0))
'            .DrawBackImage = False
'            .FontFormat = DrawTextFormat.Right Or DrawTextFormat.VerticalCenter
'        End With
'        Me.AddChild(CType(lblExpLevel, UIControl))
'    End Sub

'    Private Sub UpdateDisplay()
'        Dim oOpColor As System.Drawing.Color = System.Drawing.Color.FromArgb(255, 0, 255, 0)
'        Dim oDestColor As System.Drawing.Color = System.Drawing.Color.FromArgb(255, 255, 0, 0)
'        Dim oNAColor As System.Drawing.Color = System.Drawing.Color.FromArgb(255, 128, 128, 128)

'        'we should never have to refresh this unless HPchanged is true
'        If goCurrentEnvir.oEntity(mlEntityIndex).lLastHPUpdate = mlLastHPUpdateChange AndAlso mbForceUpdate = False Then
'            If (goCurrentEnvir.oEntity(mlEntityIndex).OwnerID = glPlayerID OrElse goCurrentEnvir.oEntity(mlEntityIndex).OwnerID = gl_HARDCODE_PIRATE_PLAYER_ID) _
'              AndAlso (goCurrentEnvir.oEntity(mlEntityIndex).CurrentStatus And elUnitStatus.eUnitMoving) <> 0 _
'              AndAlso (glCurrentCycle - mlLastMovementTest) > 15 Then
'                'Ok, just update the ETA..
'                UpdateETALabel()
'            End If

'            Return
'        End If
'        mlLastHPUpdateChange = goCurrentEnvir.oEntity(mlEntityIndex).lLastHPUpdate

'        With goCurrentEnvir.oEntity(mlEntityIndex)
'            If .EntityName Is Nothing OrElse .EntityName = "" Then
'                lblEntityName.Caption = "Unknown Unit"
'            Else : lblEntityName.Caption = .EntityName
'            End If

'            If .OwnerID <> glPlayerID AndAlso .OwnerID <> gl_HARDCODE_PIRATE_PLAYER_ID Then
'                lblEngineStatus.ForeColor = oNAColor
'                lblRadarStatus.ForeColor = oNAColor
'                lblShieldStatus.ForeColor = oNAColor
'                lblHangarStatus.ForeColor = oNAColor
'                lblCargoStatus.ForeColor = oNAColor
'                lblForeWpn1Stat.ForeColor = oNAColor
'                lblForeWpn2Stat.ForeColor = oNAColor
'                lblLeftWpn1Stat.ForeColor = oNAColor
'                lblLeftWpn2Stat.ForeColor = oNAColor
'                lblRearWpn1Stat.ForeColor = oNAColor
'                lblRearWpn2Stat.ForeColor = oNAColor
'                lblRightWpn1Stat.ForeColor = oNAColor
'                lblRightWpn2Stat.ForeColor = oNAColor
'                lblFuelCap.ForeColor = oNAColor
'                lblAmmoCap.ForeColor = oNAColor
'                lblFuelCap.Caption = "???"
'                lblAmmoCap.Caption = "???"
'                lblETA.Caption = "ETA ???"

'                lblExpLevel.Caption = "U"
'                lblExpLevel.ForeColor = System.Drawing.Color.FromArgb(255, 128, 128, 128)
'                lblExpLevel.ToolTipText = "Experience: Unknown"
'            Else
'                'NOTE: CritList works differently... if the bit is 1 then the item is destroyed
'                'Engine Status
'                If (.CritList And elUnitStatus.eEngineOperational) <> 0 Then
'                    lblEngineStatus.ForeColor = oDestColor
'                Else : lblEngineStatus.ForeColor = oOpColor
'                End If

'                'Radar Status
'                If (.CritList And elUnitStatus.eRadarOperational) <> 0 Then
'                    lblRadarStatus.ForeColor = oDestColor
'                Else : lblRadarStatus.ForeColor = oOpColor
'                End If

'                'Shield Status
'                If (.CritList And elUnitStatus.eShieldOperational) <> 0 Then
'                    lblShieldStatus.ForeColor = oDestColor
'                Else : lblShieldStatus.ForeColor = oOpColor
'                End If

'                'Hangar Status
'                If (.CritList And elUnitStatus.eHangarOperational) <> 0 Then
'                    lblHangarStatus.ForeColor = oDestColor
'                Else : lblHangarStatus.ForeColor = oOpColor
'                End If

'                'Cargo Bay Status
'                If (.CritList And elUnitStatus.eCargoBayOperational) <> 0 Then
'                    lblCargoStatus.ForeColor = oDestColor
'                Else : lblCargoStatus.ForeColor = oOpColor
'                End If

'                'Fuel Bay Status
'                If (.CritList And elUnitStatus.eFuelBayOperational) <> 0 Then
'                    lblFuelStatus.ForeColor = oDestColor
'                Else : lblFuelStatus.ForeColor = oOpColor
'                End If

'                'Fore Wpn 1 Status
'                If (.CritList And elUnitStatus.eForwardWeapon1) <> 0 Then
'                    lblForeWpn1Stat.ForeColor = oDestColor
'                Else : lblForeWpn1Stat.ForeColor = oOpColor
'                End If

'                'Fore Wpn 2 Status
'                If (.CritList And elUnitStatus.eForwardWeapon2) <> 0 Then
'                    lblForeWpn2Stat.ForeColor = oDestColor
'                Else : lblForeWpn2Stat.ForeColor = oOpColor
'                End If

'                'Left Wpn 1 Status
'                If (.CritList And elUnitStatus.eLeftWeapon1) <> 0 Then
'                    lblLeftWpn1Stat.ForeColor = oDestColor
'                Else : lblLeftWpn1Stat.ForeColor = oOpColor
'                End If

'                'Left Wpn 2 Status
'                If (.CritList And elUnitStatus.eLeftWeapon2) <> 0 Then
'                    lblLeftWpn2Stat.ForeColor = oDestColor
'                Else : lblLeftWpn2Stat.ForeColor = oOpColor
'                End If

'                'Rear Wpn 1 Status
'                If (.CritList And elUnitStatus.eAftWeapon1) <> 0 Then
'                    lblRearWpn1Stat.ForeColor = oDestColor
'                Else : lblRearWpn1Stat.ForeColor = oOpColor
'                End If

'                'Rear Wpn 2 Status
'                If (.CritList And elUnitStatus.eAftWeapon2) <> 0 Then
'                    lblRearWpn2Stat.ForeColor = oDestColor
'                Else : lblRearWpn2Stat.ForeColor = oOpColor
'                End If

'                'Right Wpn 1 Status
'                If (.CritList And elUnitStatus.eRightWeapon1) <> 0 Then
'                    lblRightWpn1Stat.ForeColor = oDestColor
'                Else : lblRightWpn1Stat.ForeColor = oOpColor
'                End If

'                'Right Wpn 2 Status
'                If (.CritList And elUnitStatus.eRightWeapon2) <> 0 Then
'                    lblRightWpn2Stat.ForeColor = oDestColor
'                Else : lblRightWpn2Stat.ForeColor = oOpColor
'                End If

'                Dim lXPRank As Int32 = Math.Abs((CInt(.Exp_Level) - 1) \ 25) '  CInt(Math.Floor(Math.Abs(.Exp_Level - 1) / 25))
'                Select Case lXPRank
'                    Case 0  'Green
'                        lblExpLevel.Caption = "G"
'                        lblExpLevel.ForeColor = System.Drawing.Color.FromArgb(255, 0, 255, 0)
'                        lblExpLevel.ToolTipText = "Experience: Green"
'                    Case 1  'Trained
'                        lblExpLevel.Caption = "T"
'                        lblExpLevel.ForeColor = System.Drawing.Color.FromArgb(255, 0, 255, 0)
'                        lblExpLevel.ToolTipText = "Experience: Trained" & vbCrLf & "+5 To-Hit Bonus"
'                    Case 2  'Experienced
'                        lblExpLevel.Caption = "E"
'                        lblExpLevel.ForeColor = System.Drawing.Color.FromArgb(255, 0, 255, 0)
'                        lblExpLevel.ToolTipText = "Experience: Experienced" & vbCrLf & "+8 To-Hit Bonus"
'                    Case 3  'Adept
'                        lblExpLevel.Caption = "A"
'                        lblExpLevel.ForeColor = System.Drawing.Color.FromArgb(255, 0, 255, 0)
'                        lblExpLevel.ToolTipText = "Experience: Adept" & vbCrLf & "+8 To-Hit Bonus" & vbCrLf & "+1 Speed"
'                    Case 4  'Veteran
'                        lblExpLevel.Caption = "V"
'                        lblExpLevel.ForeColor = System.Drawing.Color.FromArgb(255, 255, 0, 0)
'                        lblExpLevel.ToolTipText = "Experience: Veteran" & vbCrLf & "+8 To-Hit Bonus" & vbCrLf & "+1 Speed" & vbCrLf & "+1 Maneuver"
'                    Case 5  'Ace
'                        lblExpLevel.Caption = "A"
'                        lblExpLevel.ForeColor = System.Drawing.Color.FromArgb(255, 255, 0, 0)
'                        lblExpLevel.ToolTipText = "Experience: Ace" & vbCrLf & "+10 To-Hit Bonus" & vbCrLf & "+1 Speed" & vbCrLf & "+1 Maneuver"
'                    Case 6  'Top Ace
'                        lblExpLevel.Caption = "T"
'                        lblExpLevel.ForeColor = System.Drawing.Color.FromArgb(255, 255, 0, 0)
'                        lblExpLevel.ToolTipText = "Experience: Top Ace" & vbCrLf & "+10 To-Hit Bonus" & vbCrLf & "+1 Speed" & vbCrLf & "+1 Maneuver" & vbCrLf & "+5 Damage"
'                    Case 7  'Distinguished
'                        lblExpLevel.Caption = "D"
'                        lblExpLevel.ForeColor = System.Drawing.Color.FromArgb(255, 255, 0, 0)
'                        lblExpLevel.ToolTipText = "Experience: Distinguished" & vbCrLf & "+10 To-Hit Bonus" & vbCrLf & "+3 Speed" & vbCrLf & "+1 Maneuver" & vbCrLf & "+5 Damage"
'                    Case 8  'Revered
'                        lblExpLevel.Caption = "R"
'                        lblExpLevel.ForeColor = System.Drawing.Color.FromArgb(255, 255, 0, 0)
'                        lblExpLevel.ToolTipText = "Experience: Revered" & vbCrLf & "+10 To-Hit Bonus" & vbCrLf & "+4 Speed" & vbCrLf & "+2 Maneuver" & vbCrLf & "+5 Damage"
'                    Case Else   'Elite
'                        lblExpLevel.Caption = "E"
'                        lblExpLevel.ForeColor = System.Drawing.Color.FromArgb(255, 255, 255, 255)
'                        lblExpLevel.ToolTipText = "Experience: Elite" & vbCrLf & "+13 To-Hit Bonus" & vbCrLf & "+5 Speed" & vbCrLf & "+3 Maneuver" & vbCrLf & "+10 Damage"
'                End Select

'                If .ObjTypeID <> ObjectType.eFacility Then
'                    lblExpLevel.ToolTipText &= vbCrLf & "CP Usage: " & Math.Max((10 - lXPRank), 1)
'                End If


'                lblFuelCap.Caption = .Fuel_Cap & "%"
'                If .Fuel_Cap < 50 Then
'                    lblFuelCap.ForeColor = System.Drawing.Color.FromArgb(255, 255, 255, 0)
'                Else : lblFuelCap.ForeColor = oOpColor
'                End If

'                lblAmmoCap.Caption = .Ammo_Cap & "%"
'                If .Ammo_Cap < 25 Then
'                    lblAmmoCap.ForeColor = System.Drawing.Color.FromArgb(255, 255, 255, 0)
'                Else : lblAmmoCap.ForeColor = oOpColor
'                End If

'                UpdateETALabel()
'            End If
'        End With

'        mbForceUpdate = False
'    End Sub

'    Private Sub UpdateETALabel()
'        Dim sTemp As String = "ETA 00:00:00:00"
'        With goCurrentEnvir.oEntity(mlEntityIndex)
'            If (.CurrentStatus And elUnitStatus.eUnitMoving) <> 0 AndAlso .oUnitDef Is Nothing = False AndAlso .oUnitDef.MaxSpeed <> 0 Then
'                mlLastMovementTest = glCurrentCycle

'                'Ok, unit is moving, get the distance between it's loc and dest
'                Dim fDX As Single = .DestX - .fMapWrapLocX
'                Dim fDZ As Single = .DestZ - .LocZ
'                fDX *= fDX
'                fDZ *= fDZ
'                Dim fDist As Single = CSng(Math.Sqrt(fDX + fDZ))
'                fDist /= .oUnitDef.MaxSpeed
'                'fDist is number of cycles...
'                Dim lSeconds As Int32 = CInt(Math.Ceiling(fDist / 30.0F))
'                Dim lMinutes As Int32 = lSeconds \ 60
'                lSeconds -= (lMinutes * 60)
'                Dim lHours As Int32 = lMinutes \ 60
'                lMinutes -= (lHours * 60)
'                Dim lDays As Int32 = lHours \ 24
'                lHours -= (lDays * 24)

'                sTemp = "ETA " & lDays.ToString("00") & ":" & lHours.ToString("00") & ":" & lMinutes.ToString("00") & ":" & lSeconds.ToString("00")
'            End If
'        End With

'        If lblETA.Caption <> sTemp Then lblETA.Caption = sTemp
'    End Sub

'    Public Sub SetFromEntity(ByVal lEntityIndex As Int32)
'        'Verify our objects
'        If goCurrentEnvir Is Nothing Then Exit Sub
'        If lEntityIndex < 0 OrElse lEntityIndex > goCurrentEnvir.lEntityUB Then Exit Sub
'        If goCurrentEnvir.oEntity(lEntityIndex) Is Nothing Then Exit Sub
'        If goCurrentEnvir.oEntity(lEntityIndex).oUnitDef Is Nothing Then Exit Sub

'        mlEntityIndex = lEntityIndex
'        mbForceUpdate = True

'        If goCurrentEnvir.oEntity(lEntityIndex).ObjTypeID = ObjectType.eFacility AndAlso goCurrentEnvir.oEntity(lEntityIndex).OwnerID = glPlayerID Then
'            Dim oTmpWin As frmFacilityAdv = New frmFacilityAdv(MyBase.moUILib)
'            oTmpWin.SetFromEntity(mlEntityIndex)
'            oTmpWin = Nothing
'        Else
'            MyBase.moUILib.RemoveWindow("frmFacilityAdv")
'        End If

'        Dim oProdWin As frmProdStatus = New frmProdStatus(MyBase.moUILib)
'        oProdWin.SetFromEntity(mlEntityIndex)
'        oProdWin = Nothing


'    End Sub

'    Private Sub frmSingleSelect_OnNewFrame() Handles MyBase.OnNewFrame
'        If mlEntityIndex <> -1 AndAlso goCurrentEnvir Is Nothing = False AndAlso goCurrentEnvir.lEntityIdx(mlEntityIndex) <> -1 Then
'            UpdateDisplay()
'        End If
'    End Sub

'    Protected Overrides Sub Finalize()
'        Debug.Write(Me.ControlName & " Finalized" & vbCrLf)
'        MyBase.Finalize()
'    End Sub
'End Class



'Selection of a single entity display

Public Class frmSingleSelect
    Inherits UIWindow

    Private WithEvents txtEntityName As UITextBox
    Private lblOwner As UILabel
    Private lblETA As UILabel
    Private lblTether As UILabel

    Private hln1 As UILine
    Private hln2 As UILine

    'Private WithEvents chkAlertFinalDest As UICheckBox

    Private mlEntityIndex As Int32 = -1
    Private mbForceUpdate As Boolean

    Private mlLastHPUpdateChange As Int32 = Int32.MinValue
    Private mlPreviousStatus As Int32 = 0
    Private mlPreviousCrits As Int32 = 0

    Private mlLastMovementTest As Int32
    Private mlLastWeaponUpdate As Int32 = Int32.MinValue
    Private mlLastWeaponUpdateForced As Int32 = 0
    Private mlLastHPRequest As Int32 = -1000

    Private mlCurrentLeft As Int32 = 0
    Private mlCurrentTop As Int32 = 0

    Private mbNameChangeInProgress As Boolean = False

    Private rcAlertFinalDest As Rectangle
    Private rcGuildAsset As Rectangle
    Private mbAlertFinalDest As Boolean = False

    Private Structure uUnitAlert
        Public lObjID As Int32
        Public iObjTypeID As Int16
    End Structure
    Private Shared muUnitAlerts() As uUnitAlert
    Private Shared mlUnitAlertUB As Int32 = -1
    Public Shared Sub UnitAlertReceived(ByVal lID As Int32, ByVal iTypeID As Int16)
        For X As Int32 = 0 To mlUnitAlertUB
            If muUnitAlerts(X).lObjID = lID AndAlso muUnitAlerts(X).iObjTypeID = iTypeID Then
                For Y As Int32 = X To mlUnitAlertUB - 1
                    muUnitAlerts(Y) = muUnitAlerts(Y + 1)
                Next Y
                mlUnitAlertUB -= 1
                Exit For
            End If
        Next X
    End Sub
    Public Shared Sub AddUnitAlert(ByVal lID As Int32, ByVal iTypeID As Int16)
        For X As Int32 = 0 To mlUnitAlertUB
            If muUnitAlerts(X).lObjID = lID AndAlso muUnitAlerts(X).iObjTypeID = iTypeID Then
                Exit Sub
            End If
        Next X
        ReDim Preserve muUnitAlerts(mlUnitAlertUB + 1)
        mlUnitAlertUB += 1
        muUnitAlerts(mlUnitAlertUB).lObjID = lID
        muUnitAlerts(mlUnitAlertUB).iObjTypeID = iTypeID
    End Sub

    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)

        'frmSingleSelect initial props
        With Me
            .lWindowMetricID = BPMetrics.eWindow.eSingleSelect
            .ControlName = "frmSingleSelect"
            .Width = 128
            .Height = 128
            .Left = 0
            .Top = moUILib.oDevice.PresentationParameters.BackBufferHeight - .Height

            If muSettings.SelectLocX <> -1 AndAlso NewTutorialManager.TutorialOn = False Then .Left = muSettings.SelectLocX
            If muSettings.SelectLocY <> -1 AndAlso NewTutorialManager.TutorialOn = False Then .Top = muSettings.SelectLocY
            If .Left + .Width > oUILib.oDevice.PresentationParameters.BackBufferWidth Then .Left = oUILib.oDevice.PresentationParameters.BackBufferWidth - .Width
            If .Top + .Height > oUILib.oDevice.PresentationParameters.BackBufferHeight Then .Top = oUILib.oDevice.PresentationParameters.BackBufferHeight - .Height
            .Enabled = True
            .Visible = True
            '.BorderColor = System.Drawing.Color.FromArgb(-16711681)
            .BorderColor = muSettings.InterfaceBorderColor
            '.FillColor = System.Drawing.Color.FromArgb(-16768960)
            '.FillColor = System.Drawing.Color.FromArgb(128, 64, 128, 192)
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = False
            .BorderLineWidth = 2
        End With
        Debug.Write(Me.ControlName & " Newed" & vbCrLf)

        ''lblEntityName initial props
        'lblEntityName = New UILabel(oUILib)
        'With lblEntityName
        '    .ControlName = "lblEntityName"
        '    .Left = 3
        '    .Top = 2
        '    .Width = 126
        '    .Height = 16
        '    .Enabled = True
        '    .Visible = True
        '    .Caption = "ENTITY NAME"
        '    '.ForeColor = System.Drawing.Color.FromArgb(-16711681)
        '    .ForeColor = muSettings.InterfaceBorderColor
        '    .SetFont(New System.Drawing.Font("Arial", 9, FontStyle.Bold, GraphicsUnit.Point, 0))
        '    .DrawBackImage = False
        '    .FontFormat = DrawTextFormat.VerticalCenter
        'End With
        'Me.AddChild(CType(lblEntityName, UIControl))

        'txtEntityName initial props
        txtEntityName = New UITextBox(oUILib)
        With txtEntityName
            .ControlName = "txtEntityName"
            .Left = 0 '3
            .Top = 2
            .Width = 126
            .Height = 16
            .Enabled = True
            .Visible = True
            .Caption = "ENTITY NAME"
            .ForeColor = muSettings.InterfaceTextBoxForeColor
            .SetFont(New System.Drawing.Font("Arial", 9, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.VerticalCenter
            .BackColorEnabled = Me.FillColor
            .BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
            .MaxLength = 20
            .BorderColor = muSettings.InterfaceBorderColor
            .DoNotRender = UITextBox.DoNotRenderSetting.eBorder Or UITextBox.DoNotRenderSetting.eFillColor
        End With
        Me.AddChild(CType(txtEntityName, UIControl))

        'lblOwner initial props
        lblOwner = New UILabel(oUILib)
        With lblOwner
            .ControlName = "lblOwner"
            .Left = 3
            .Top = 20
            .Width = 126
            .Height = 16
            .Enabled = True
            .Visible = True
            .Caption = "OWNER NAME"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Arial", 9, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.VerticalCenter
        End With
        Me.AddChild(CType(lblOwner, UIControl))

        'lblTether initial props
        lblTether = New UILabel(oUILib)
        With lblTether
            .ControlName = "lblTether"
            .Left = Me.Width - 25   '3
            .Top = 20
            .Width = 16 ' 126
            .Height = 16
            .Enabled = True
            .Visible = False
            .Caption = "T"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Arial", 9, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.VerticalCenter
            .bAcceptEvents = True
            .ToolTipText = "This unit has a tether location. Click this T to go to that point." & vbCrLf & "Press Ctrl+T to set the unit's position as the tether point." & vbCrLf & _
               "Press Ctrl+Shift+T to clear the unit's tether point." & vbCrLf & "A unit will try to remain at its tether location when not engaged or pursuing."
        End With
        Me.AddChild(CType(lblTether, UIControl))

        'lblETA initial props
        lblETA = New UILabel(oUILib)
        With lblETA
            .ControlName = "lblETA"
            .Left = 2
            .Top = 115
            .Width = 126
            .Height = 10
            .Enabled = True
            .Visible = True
            .Caption = "ETA 00:00:00:00"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 6, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Center
            .ToolTipText = "The ETA to the current waypoint."
        End With
		Me.AddChild(CType(lblETA, UIControl))

        ''chkAlertFinalDest initial props
        'chkAlertFinalDest = New UICheckBox(oUILib)
        'With chkAlertFinalDest
        '	.ControlName = "chkAlertFinalDest"
        '	.Left = 15
        '	.Top = lblETA.Top - 20
        '	.Width = 80
        '	.Height = 18
        '	.Enabled = True
        '	.Visible = False
        '	.Caption = "Alert On Stop"
        '	.ForeColor = muSettings.InterfaceBorderColor
        '	.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 7, FontStyle.Bold, GraphicsUnit.Point, 0))
        '	.DrawBackImage = False
        '	.FontFormat = CType(6, DrawTextFormat)
        '	.Value = False
        '	.DisplayType = UICheckBox.eCheckTypes.eSmallCheck
        'End With
        'Me.AddChild(CType(chkAlertFinalDest, UIControl))
        rcAlertFinalDest = New Rectangle(15, lblETA.Top - 20, 16, 16)
        rcGuildAsset = New Rectangle(rcAlertFinalDest.Left + rcAlertFinalDest.Width + 5, rcAlertFinalDest.Top, 16, 16)

        'hln1 initial props
        hln1 = New UILine(oUILib)
        With hln1
            .ControlName = "hln1"
            .Left = 1
            .Top = 19
            .Width = 189
            .Height = 0
            .Enabled = True
            .Visible = True
            '.BorderColor = System.Drawing.Color.FromArgb(-16711681)
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(hln1, UIControl))

        'hln2 initial props
        hln2 = New UILine(oUILib)
        With hln2
            .ControlName = "hln2"
            .Left = 1
            .Top = 37
            .Width = 189
            .Height = 0
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(hln2, UIControl))


        AddHandler lblTether.OnMouseDown, AddressOf lblTetherClick
    End Sub

    Private Sub lblTetherClick(ByVal lMouseX As Int32, ByVal lMouseY As Int32, ByVal lButton As MouseButtons)
        If mlEntityIndex = -1 Then Return
        If goCurrentEnvir Is Nothing Then Return

        Try
            With goCurrentEnvir.oEntity(mlEntityIndex)
                If .lTetherPointX <> Int32.MinValue AndAlso .lTetherPointZ <> Int32.MinValue Then
                    If goCamera Is Nothing = False Then
                        goCamera.SimplyPlaceCamera(.lTetherPointX, goCamera.mlCameraY, .lTetherPointZ)
                    End If
                End If
            End With
        Catch
        End Try
    End Sub

    Private Sub UpdateETALabel()
        Try
            Dim sTemp As String = "ETA 00:00:00:00"
            With goCurrentEnvir.oEntity(mlEntityIndex)
                If (.CurrentStatus And elUnitStatus.eUnitMoving) <> 0 AndAlso .oUnitDef Is Nothing = False AndAlso .oUnitDef.MaxSpeed <> 0 Then
                    mlLastMovementTest = glCurrentCycle

                    'Ok, unit is moving, get the distance between it's loc and dest
                    Dim fDX As Single = .DestX - .LocX '.fMapWrapLocX
                    Dim fDZ As Single = .DestZ - .LocZ
                    fDX *= fDX
                    fDZ *= fDZ
                    Dim fDist As Single = CSng(Math.Sqrt(fDX + fDZ))
                    fDist /= .oUnitDef.MaxSpeed
                    'fDist is number of cycles...
                    Dim lSeconds As Int32 = CInt(Math.Ceiling(fDist / 30.0F))
                    Dim lMinutes As Int32 = lSeconds \ 60
                    lSeconds -= (lMinutes * 60)
                    Dim lHours As Int32 = lMinutes \ 60
                    lMinutes -= (lHours * 60)
                    Dim lDays As Int32 = lHours \ 24
                    lHours -= (lDays * 24)

                    sTemp = "ETA " & lDays.ToString("00") & ":" & lHours.ToString("00") & ":" & lMinutes.ToString("00") & ":" & lSeconds.ToString("00")
                End If
            End With

            If lblETA.Caption <> sTemp Then lblETA.Caption = sTemp
        Catch
            'do nothing
        End Try
    End Sub

    Public Sub SetFromEntity(ByVal lEntityIndex As Int32)
        'Verify our objects

        Try
            If goCurrentEnvir Is Nothing Then Return
            If lEntityIndex < 0 OrElse lEntityIndex > goCurrentEnvir.lEntityUB Then Return
            If goCurrentEnvir.oEntity(lEntityIndex) Is Nothing Then Return
            If goCurrentEnvir.oEntity(lEntityIndex).oUnitDef Is Nothing Then Return
            txtEntityName.Enabled = Not gbAliased
            mlEntityIndex = lEntityIndex
            mbForceUpdate = True

            'If goCurrentEnvir.oEntity(lEntityIndex).ObjTypeID = ObjectType.eFacility AndAlso goCurrentEnvir.oEntity(lEntityIndex).OwnerID = glPlayerID Then
            '    Dim oTmpWin As frmFacilityAdv = New frmFacilityAdv(MyBase.moUILib)
            '    oTmpWin.SetFromEntity(mlEntityIndex)
            '    oTmpWin = Nothing
            'Else
            '    MyBase.moUILib.RemoveWindow("frmFacilityAdv")
            'End If
            'chkAlertFinalDest.Visible = False
            'chkAlertFinalDest.Value = False
            mbAlertFinalDest = False

            Dim lEntityID As Int32 = goCurrentEnvir.oEntity(lEntityIndex).ObjectID
            Dim iEntityTypeID As Int16 = goCurrentEnvir.oEntity(lEntityIndex).ObjTypeID

            For X As Int32 = 0 To mlUnitAlertUB
                If muUnitAlerts(X).lObjID = lEntityID AndAlso muUnitAlerts(X).iObjTypeID = iEntityTypeID Then
                    mbAlertFinalDest = True
                    Exit For
                End If
            Next X
            Dim lProductionTypeID As Byte = goCurrentEnvir.oEntity(lEntityIndex).yProductionType
            If goCurrentEnvir.oEntity(lEntityIndex).OwnerID = glPlayerID Then
                If lProductionTypeID > 0 AndAlso lProductionTypeID <> ProductionType.eMining AndAlso lProductionTypeID <> ProductionType.eColonists AndAlso (lProductionTypeID And ProductionType.ePowerCenter) = 0 Then
                    Dim oProdWin As frmProdStatus = New frmProdStatus(MyBase.moUILib)
                    oProdWin.SetFromEntity(mlEntityIndex)
                    oProdWin = Nothing
                End If
                'Send a request for the entity's details
                Dim yMsg(13) As Byte
                System.BitConverter.GetBytes(GlobalMessageCode.eGetEntityDetails).CopyTo(yMsg, 0)
                goCurrentEnvir.oEntity(lEntityIndex).GetGUIDAsString.CopyTo(yMsg, 2)
                goCurrentEnvir.GetGUIDAsString.CopyTo(yMsg, 8)
                MyBase.moUILib.SendMsgToRegion(yMsg)
                goCurrentEnvir.oEntity(lEntityIndex).bRequestedDetails = True
                txtEntityName.Locked = False
                'chkAlertFinalDest.Visible = True

                Dim oEntity As BaseEntity = goCurrentEnvir.oEntity(lEntityIndex)
                lblETA.Visible = oEntity.oUnitDef Is Nothing = False AndAlso oEntity.oUnitDef.MaxSpeed > 0

            Else
                txtEntityName.Locked = True
                MyBase.moUILib.RemoveWindow("frmProdStatus")
            End If
        Catch
            'Do nothing
        End Try
        mbNameChangeInProgress = False
    End Sub

    Private Sub frmSingleSelect_OnMouseDown(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles Me.OnMouseDown
        Dim lX As Int32 = lMouseX - GetAbsolutePosition.X
        Dim lY As Int32 = lMouseY - GetAbsolutePosition.Y

        Dim oEnvir As BaseEnvironment = goCurrentEnvir
        If oEnvir Is Nothing Then Return
        Dim oEntity As BaseEntity = Nothing
        If oEnvir.lEntityUB >= mlEntityIndex AndAlso mlEntityIndex > -1 Then
            oEntity = oEnvir.oEntity(mlEntityIndex)
        End If
        If oEntity Is Nothing Then Return

        If oEntity.OwnerID = glPlayerID Then
            If rcAlertFinalDest.Contains(lX, lY) = True AndAlso oEntity.oUnitDef Is Nothing = False AndAlso oEntity.oUnitDef.MaxSpeed > 0 Then
                chkAlertFinalDest_Click()
            ElseIf rcGuildAsset.Contains(lX, lY) = True Then
                chkGuildAsset_Click()
            End If
        End If
    End Sub

    Private Sub frmSingleSelect_OnMouseMove(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles Me.OnMouseMove
        Dim lX As Int32 = lMouseX - GetAbsolutePosition.X
        Dim lY As Int32 = lMouseY - GetAbsolutePosition.Y
		Dim sTemp As String = ""

		MyBase.moUILib.SetToolTip(False)

		Dim oEnvir As BaseEnvironment = goCurrentEnvir
		If oEnvir Is Nothing Then Return
		Dim oEntity As BaseEntity = Nothing
		If oEnvir.lEntityUB >= mlEntityIndex AndAlso mlEntityIndex > -1 Then
			oEntity = oEnvir.oEntity(mlEntityIndex)
		End If
		If oEntity Is Nothing Then Return

        Dim bDisplay As Boolean = (oEntity.OwnerID = glPlayerID)

		If New Rectangle(2, 33, 50, 64).Contains(lX, lY) = True Then
			bDisplay = True
			With oEntity
				If .yArmorHP Is Nothing Then Return
				sTemp = "Shield: " & .yShieldHP & "%" & vbCrLf
				If .yArmorHP(0) = 255 Then sTemp &= "Forward: None" & vbCrLf Else sTemp &= "Forward: " & .yArmorHP(0) & "%" & vbCrLf
				If .yArmorHP(1) = 255 Then sTemp &= "Left: None" & vbCrLf Else sTemp &= "Left: " & .yArmorHP(1) & "%" & vbCrLf
				If .yArmorHP(2) = 255 Then sTemp &= "Rear: None" & vbCrLf Else sTemp &= "Rear: " & .yArmorHP(2) & "%" & vbCrLf
				If .yArmorHP(3) = 255 Then sTemp &= "Right: None" & vbCrLf Else sTemp &= "Right: " & .yArmorHP(3) & "%" & vbCrLf
				sTemp &= "Structure: " & .yStructureHP & "%" & vbCrLf
			End With
			MyBase.moUILib.SetToolTip(sTemp, lMouseX, lMouseY)
		ElseIf New Rectangle(53, 40, 16, 16).Contains(lX, lY) = True Then
			If (oEntity.CritList And elUnitStatus.eEngineOperational) <> 0 Then
                sTemp = "Engines are destroyed"
            ElseIf (oEntity.CurrentStatus And elUnitStatus.eEngineOperational) <> 0 Then : sTemp = "Engines are Operational"
            End If
            If oEntity.oUnitDef Is Nothing = False AndAlso oEntity.oUnitDef.Maneuver > 0 Then sTemp &= vbCrLf & "Maneuver: " & oEntity.oUnitDef.Maneuver.ToString
            If oEntity.oUnitDef Is Nothing = False AndAlso oEntity.oUnitDef.MaxSpeed > 0 Then sTemp &= vbCrLf & "Max Speed: " & oEntity.oUnitDef.MaxSpeed.ToString
            'If oEntity.oUnitDef Is Nothing = False AndAlso oEntity.oChild(0) is nothing = False andalso oentity.oChild(0). > 0 Then sTemp &= vbCrLf & "Max Speed: " & oEntity.oUnitDef.MaxSpeed.ToString
		ElseIf New Rectangle(53, 60, 16, 16).Contains(lX, lY) = True Then
			If (oEntity.CritList And elUnitStatus.eRadarOperational) <> 0 Then
                sTemp = "Radar is destroyed"
            ElseIf (oEntity.CurrentStatus And elUnitStatus.eRadarOperational) <> 0 Then : sTemp = "Radar is Operational"
            End If
            If oEntity.oUnitDef Is Nothing = False AndAlso oEntity.oUnitDef.OptRadarRange > 0 Then sTemp &= vbCrLf & "Optimum Radar Range: " & oEntity.oUnitDef.OptRadarRange.ToString
            If oEntity.oUnitDef Is Nothing = False AndAlso oEntity.oUnitDef.MaxRadarRange > 0 Then sTemp &= vbCrLf & "Maximum Radar Range: " & oEntity.oUnitDef.MaxRadarRange.ToString
		ElseIf New Rectangle(53, 80, 16, 16).Contains(lX, lY) = True Then
			If (oEntity.CritList And elUnitStatus.eShieldOperational) <> 0 Then
                sTemp = "Shields are destroyed"
            ElseIf (oEntity.CurrentStatus And elUnitStatus.eShieldOperational) <> 0 Then : sTemp = "Shields are working"
            End If
		ElseIf New Rectangle(72, 40, 16, 16).Contains(lX, lY) = True Then
			If (oEntity.CritList And elUnitStatus.eHangarOperational) <> 0 Then
                sTemp = "Hangar is destroyed"
            ElseIf (oEntity.CurrentStatus And elUnitStatus.eHangarOperational) <> 0 Then : sTemp = "Hangars are Operational"
            End If
		ElseIf New Rectangle(72, 60, 16, 16).Contains(lX, lY) = True Then
			If (oEntity.CritList And elUnitStatus.eCargoBayOperational) <> 0 Then
                sTemp = "Cargo Bay is destroyed"
            ElseIf (oEntity.CurrentStatus And elUnitStatus.eCargoBayOperational) <> 0 Then : sTemp = "Cargo Bay is Operational"
            End If
			'ElseIf New Rectangle(72, 80, 16, 16).Contains(lX, lY) = True Then
			'    If (goCurrentEnvir.oEntity(mlEntityIndex).CritList And elUnitStatus.eFuelBayOperational) <> 0 Then
			'        sTemp = "Fuel Bay has been destroyed"
			'    Else : sTemp = "Fuel Bay is working"
			'    End If
		ElseIf New Rectangle(97, 45, 10, 6).Contains(lX, lY) = True Then
			If (oEntity.CritList And elUnitStatus.eForwardWeapon1) <> 0 Then
                sTemp = "Forward Weapon Group 1 is destroyed"
            ElseIf (oEntity.CurrentStatus And elUnitStatus.eForwardWeapon1) <> 0 Then : sTemp = "Forward Weapon Group 1 is Operational"
            End If
		ElseIf New Rectangle(108, 45, 10, 6).Contains(lX, lY) = True Then
			If (oEntity.CritList And elUnitStatus.eForwardWeapon2) <> 0 Then
                sTemp = "Forward Weapon Group 2 is destroyed"
            ElseIf (oEntity.CurrentStatus And elUnitStatus.eForwardWeapon2) <> 0 Then : sTemp = "Forward Weapon Group 2 is Operational"
			End If
		ElseIf New Rectangle(92, 55, 6, 10).Contains(lX, lY) = True Then
			If (oEntity.CritList And elUnitStatus.eLeftWeapon1) <> 0 Then
                sTemp = "Left Weapon Group 1 is destroyed"
            ElseIf (oEntity.CurrentStatus And elUnitStatus.eLeftWeapon1) <> 0 Then : sTemp = "Left Weapon Group 1 is Operational"
			End If
		ElseIf New Rectangle(92, 71, 6, 10).Contains(lX, lY) = True Then
			If (oEntity.CritList And elUnitStatus.eLeftWeapon2) <> 0 Then
                sTemp = "Left Weapon Group 2 is destroyed"
            ElseIf (oEntity.CurrentStatus And elUnitStatus.eLeftWeapon2) <> 0 Then : sTemp = "Left Weapon Group 2 is Operational"
			End If
		ElseIf New Rectangle(97, 81, 10, 6).Contains(lX, lY) = True Then
			If (oEntity.CritList And elUnitStatus.eAftWeapon1) <> 0 Then
                sTemp = "Rear Weapon Group 1 is destroyed"
            ElseIf (oEntity.CurrentStatus And elUnitStatus.eAftWeapon1) <> 0 Then : sTemp = "Rear Weapon Group 1 is Operational"
			End If
		ElseIf New Rectangle(109, 85, 10, 6).Contains(lX, lY) = True Then
			If (oEntity.CritList And elUnitStatus.eAftWeapon2) <> 0 Then
                sTemp = "Rear Weapon Group 2 is destroyed"
            ElseIf (oEntity.CurrentStatus And elUnitStatus.eAftWeapon2) <> 0 Then : sTemp = "Rear Weapon Group 2 is Operational"
			End If
		ElseIf New Rectangle(118, 55, 6, 10).Contains(lX, lY) = True Then
			If (oEntity.CritList And elUnitStatus.eRightWeapon1) <> 0 Then
                sTemp = "Right Weapon Group 1 is destroyed"
            ElseIf (oEntity.CurrentStatus And elUnitStatus.eRightWeapon1) <> 0 Then : sTemp = "Right Weapon Group 1 is Operational"
			End If
		ElseIf New Rectangle(118, 71, 6, 10).Contains(lX, lY) = True Then
			If (oEntity.CritList And elUnitStatus.eRightWeapon2) <> 0 Then
                sTemp = "Right Weapon Group 2 is destroyed"
            ElseIf (oEntity.CurrentStatus And elUnitStatus.eRightWeapon2) <> 0 Then : sTemp = "Right Weapon Group 2 is Operational"
			End If
			'ElseIf New Rectangle(2, 99, 61, 10).Contains(lX, lY) = True Then
			'    sTemp = "Available Fuel: " & (goCurrentEnvir.oEntity(mlEntityIndex).Fuel_Cap).ToString("##0") & "%"
			'ElseIf New Rectangle(65, 99, 61, 10).Contains(lX, lY) = True Then
			'    sTemp = "Available Ammo: " & (goCurrentEnvir.oEntity(mlEntityIndex).Ammo_Cap).ToString("##0") & "%"
        ElseIf New Rectangle(Me.Width - 16, lblOwner.Top, 16, 16).Contains(lX, lY) = True AndAlso oEntity.ObjTypeID <> ObjectType.eFacility Then
            Dim lXPRank As Int32 = Math.Abs((CInt(oEntity.Exp_Level) - 1) \ 25) '  CInt(Math.Floor(Math.Abs(.Exp_Level - 1) / 25))
            sTemp = BaseEntity.GetExperienceLevelNameAndBenefits(oEntity.Exp_Level)
            If oEntity.ObjTypeID <> ObjectType.eFacility Then
                sTemp &= vbCrLf & "CP Usage: " & Math.Max((10 - lXPRank), 1)
                'Dim lWPUpkeep As Int32 = GetUnitWarpointUpkeep(oEntity.ObjectID)
                'If lWPUpkeep > -1 Then sTemp &= vbCrLf & "WP Upkeep Cost: " & lWPUpkeep.ToString("#,##0")
            End If
        ElseIf bDisplay = True AndAlso rcAlertFinalDest.Contains(lX, lY) = True AndAlso oEntity.oUnitDef Is Nothing = False AndAlso oEntity.oUnitDef.MaxSpeed > 0 Then
            sTemp = "If green, an alert will be sent when the unit reaches its destination." & vbCrLf & "Click to toggle between green and red."
        ElseIf bDisplay = True AndAlso rcGuildAsset.Contains(lX, lY) = True Then
            sTemp = "If green, guild mates can give orders to this asset." & vbCrLf & "If red, guild mates cannot use this asset." & vbCrLf & "Click the icon to toggle between green and red."
        End If

		If sTemp <> "" AndAlso bDisplay = True Then MyBase.moUILib.SetToolTip(sTemp, lMouseX, lMouseY)
    End Sub

	Private Sub frmSingleSelect_OnNewFrame() Handles MyBase.OnNewFrame
		Try
			Dim oEnvir As BaseEnvironment = goCurrentEnvir

			If mlEntityIndex <> -1 AndAlso oEnvir Is Nothing = False AndAlso mlEntityIndex <= oEnvir.lEntityUB Then
				'AndAlso goCurrentEnvir.lEntityIdx(mlEntityIndex) <> -1 Then
				Dim oEntity As BaseEntity = oEnvir.oEntity(mlEntityIndex)
				If oEntity Is Nothing Then
					MyBase.moUILib.RemoveSelection(mlEntityIndex)
					Return
				End If

				'Update anything that is not accurate due to time...
				UpdateETALabel()

                With oEntity

                    Dim bOn As Boolean = mbAlertFinalDest
                    Dim bStillOn As Boolean = False
                    For X As Int32 = 0 To mlUnitAlertUB
                        If muUnitAlerts(X).lObjID = .ObjectID AndAlso muUnitAlerts(X).iObjTypeID = .ObjTypeID Then
                            bStillOn = True
                            Exit For
                        End If
                    Next X
                    If bOn <> bStillOn Then
                        mbAlertFinalDest = bStillOn
                        Me.IsDirty = True
                    End If

                    If .lTetherPointX <> Int32.MinValue AndAlso .lTetherPointZ <> Int32.MinValue Then
                        If lblTether.Visible = False Then lblTether.Visible = True
                    ElseIf lblTether.Visible = True Then
                        lblTether.Visible = False
                    End If

                    If .lLastHPUpdate <> mlLastHPUpdateChange OrElse .CurrentStatus <> mlPreviousStatus OrElse .CritList <> mlPreviousCrits Then
                        Me.IsDirty = True
                    End If
                    If mbNameChangeInProgress = False AndAlso .EntityName Is Nothing = False AndAlso .EntityName <> txtEntityName.Caption Then txtEntityName.Caption = .EntityName
                    If .sOverrideOwnerName Is Nothing OrElse .sOverrideOwnerName = "" Then
                        Dim bUnknown As Boolean = False
                        .sOverrideOwnerName = GetCacheObjectValueCheckUnknowns(.OwnerID, ObjectType.ePlayer, bUnknown)
                        If bUnknown = True Then .sOverrideOwnerName = ""
                    End If
                    If .sOverrideOwnerName <> lblOwner.Caption Then lblOwner.Caption = .sOverrideOwnerName

                    If (.lLastWeaponUpdate > mlLastWeaponUpdate OrElse glCurrentCycle - mlLastWeaponUpdateForced > 300) AndAlso glCurrentCycle - mlLastHPRequest > 30 Then

                        If .lLastHPUpdate > 0 Then
                            Dim bCheckDmg As Boolean = False
                            For Y As Int32 = 0 To 3
                                If .yArmorHP(Y) <> 100 Then
                                    bCheckDmg = True
                                    Exit For
                                End If
                            Next Y
                            If bCheckDmg = False Then
                                bCheckDmg = .yStructureHP <> 100
                            End If
                            If bCheckDmg = True Then
                                If .ObjTypeID = ObjectType.eUnit Then
                                    TutorialAlert.CheckTrigger(TutorialAlert.eTriggerType.SelectDmgdUnit)
                                Else
                                    'TutorialAlert.CheckTrigger(TutorialAlert.eTriggerType.SelectDmgdFacility)
                                End If
                            End If
                        End If

                        MyBase.moUILib.GetMsgSys.RequestEntityHPUpdate(.ObjectID, .ObjTypeID)
                        mlLastWeaponUpdateForced = glCurrentCycle
                        mlLastWeaponUpdate = .lLastWeaponUpdate
                        mlLastHPRequest = glCurrentCycle
                    End If

                    Dim clrVal As System.Drawing.Color
                    If .OwnerID <> glPlayerID Then
                        Dim bMemberInGuild As Boolean = False
                        If goCurrentPlayer Is Nothing = False Then
                            Dim oGuild As Guild = goCurrentPlayer.oGuild
                            If oGuild Is Nothing = False Then
                                bMemberInGuild = oGuild.MemberInGuild(.OwnerID)
                            End If
                        End If
                        If bMemberInGuild = True Then
                            clrVal = System.Drawing.Color.FromArgb(255, 255, 0, 255)
                        Else
                            Dim yRel As Byte = .yRelID
                            If yRel = 0 OrElse yRel = Byte.MaxValue Then
                                .yRelID = goCurrentPlayer.GetPlayerRelScore(.OwnerID)
                                yRel = .yRelID
                            End If

                            If yRel <= elRelTypes.eWar Then
                                clrVal = System.Drawing.Color.FromArgb(255, 255, 0, 0)      'enemy
                            ElseIf yRel <= elRelTypes.eNeutral Then
                                clrVal = System.Drawing.Color.FromArgb(255, 192, 192, 192)   'neutral
                            Else
                                clrVal = System.Drawing.Color.FromArgb(255, 0, 255, 255)    'ally
                            End If
                        End If
                    Else
                        clrVal = System.Drawing.Color.FromArgb(255, 0, 255, 0)           'player's
                    End If
                    If lblOwner.ForeColor <> clrVal Then lblOwner.ForeColor = clrVal
                End With
			Else : MyBase.moUILib.RemoveSelection(mlEntityIndex)
			End If
			If Me.IsDirty = True Then
				mlCurrentLeft = Me.Left
				mlCurrentTop = Me.Top
			End If
		Catch
			'do nothing,we'llcatchit next pass
		End Try
	End Sub

    Protected Overrides Sub Finalize()
        Debug.Write(Me.ControlName & " Finalized" & vbCrLf)
        MyBase.Finalize()
    End Sub

    Private Sub frmSingleSelect_OnRender() Handles Me.OnRender
        If mlEntityIndex = -1 Then Return
        If goCurrentEnvir Is Nothing Then Return

        Try
            With goCurrentEnvir.oEntity(mlEntityIndex)
                Dim sTemp As String
                If .EntityName Is Nothing OrElse .EntityName = "" Then
                    Dim oModelDef As ModelDef = goModelDefs.GetModelDef(.oMesh.lModelID)
                    If oModelDef Is Nothing = False Then
                        sTemp = oModelDef.FrameName
                    Else
                        sTemp = "Unknown Unit"
                    End If
                Else : sTemp = .EntityName
                End If
                mlPreviousStatus = .CurrentStatus
                mlPreviousCrits = .CritList
                mlLastHPUpdateChange = .lLastHPUpdate

                Dim oTTLoc As Point = MyBase.moUILib.GetToolTipLoc()
                If oTTLoc <> Point.Empty Then
                    Dim oTmpLoc As Point = New Point(mlCurrentLeft, mlCurrentTop)
                    If oTTLoc.X > oTmpLoc.X AndAlso oTTLoc.X < oTmpLoc.X + Me.Width Then
                        If oTTLoc.Y > oTmpLoc.Y AndAlso oTTLoc.Y < oTmpLoc.Y + Me.Height Then
                            'Have to set these because UI Lib sets them to 0
                            Me.Left = mlCurrentLeft
                            Me.Top = mlCurrentTop
                            frmSingleSelect_OnMouseMove(oTTLoc.X, oTTLoc.Y, MouseButtons.None)
                            'reset our left and top
                            Me.Left = 0
                            Me.Top = 0
                        End If
                    End If
                End If

                If mbNameChangeInProgress = False AndAlso txtEntityName.Caption <> sTemp Then txtEntityName.Caption = sTemp
                If .sOverrideOwnerName Is Nothing OrElse .sOverrideOwnerName = "" Then
                    Dim bUnknown As Boolean = False
                    .sOverrideOwnerName = GetCacheObjectValueCheckUnknowns(.OwnerID, ObjectType.ePlayer, bUnknown)
                    If bUnknown = True Then .sOverrideOwnerName = ""
                End If
                If lblOwner.Caption <> .sOverrideOwnerName Then lblOwner.Caption = .sOverrideOwnerName


            End With
        Catch
        End Try
    End Sub

    Private Sub frmSingleSelect_OnRenderEnd() Handles Me.OnRenderEnd
        Dim oOpColor As System.Drawing.Color = System.Drawing.Color.FromArgb(255, 0, 255, 0)
        Dim oDestColor As System.Drawing.Color = System.Drawing.Color.FromArgb(255, 255, 0, 0)
        Dim oNoColor As System.Drawing.Color = System.Drawing.Color.FromArgb(255, 128, 128, 128)
        If GFXEngine.gbPaused = True OrElse GFXEngine.gbDeviceLost = True Then Return
        Try
            With goCurrentEnvir.oEntity(mlEntityIndex)
                mlLastHPUpdateChange = .lLastHPUpdate
                .RefreshHPTexture(MyBase.moUILib.oDevice)

                Using oImgSprite As New Sprite(MyBase.moUILib.oDevice)
                    oImgSprite.Begin(SpriteFlags.AlphaBlend)

                    'Draw the hp image
                    'oImgSprite.Draw2D(.HPTexture, Rectangle.Empty, New Rectangle(0, 0, 50, 64), Point.Empty, 0, New Point(CInt(1.28F), CInt(16.5F)), Color.White)
                    RenderHPBars(oImgSprite)

                    Dim bCheckGuild As Boolean = False
                    Dim oGuild As Guild = Nothing
                    If goCurrentPlayer Is Nothing = False Then
                        oGuild = goCurrentPlayer.oGuild
                        If oGuild Is Nothing = False Then
                            bCheckGuild = (.CurrentStatus And elUnitStatus.eGuildAsset) <> 0
                        End If
                    End If
                    If goCurrentPlayer.yPlayerPhase <> 255 Then bCheckGuild = False

                    'Now, check player ID
                    If .OwnerID = glPlayerID OrElse .OwnerID = gl_HARDCODE_PIRATE_PLAYER_ID OrElse (bCheckGuild = True AndAlso oGuild.MemberInGuild(.OwnerID) = True) Then
                        'Ok, render extended properties
                        Dim oTex As Texture = goResMgr.GetTexture("HullBuilderIcons.dds", GFXResourceManager.eGetTextureType.UserInterface, "Interface.pak")
                        Dim clrVal As System.Drawing.Color

                        'NOTE: .CritList works different, if the bit is 1, the item is destroyed

                        If (.CritList And elUnitStatus.eEngineOperational) <> 0 Then
                            clrVal = oDestColor
                        ElseIf (.CurrentStatus And elUnitStatus.eEngineOperational) <> 0 Then : clrVal = oOpColor
                        Else : clrVal = oNoColor
                        End If
                        oImgSprite.Draw2D(oTex, New Rectangle(0, 16, 16, 16), New Rectangle(0, 0, 16, 16), Point.Empty, 0, New Point(53, 40), clrVal)

                        If (.CritList And elUnitStatus.eRadarOperational) <> 0 Then
                            clrVal = oDestColor
                        ElseIf (.CurrentStatus And elUnitStatus.eRadarOperational) <> 0 Then : clrVal = oOpColor
                        Else : clrVal = oNoColor
                        End If
                        oImgSprite.Draw2D(oTex, New Rectangle(48, 0, 16, 16), New Rectangle(0, 0, 16, 16), Point.Empty, 0, New Point(53, 60), clrVal)

                        If (.CritList And elUnitStatus.eShieldOperational) <> 0 Then
                            clrVal = oDestColor
                        ElseIf (.CurrentStatus And elUnitStatus.eShieldOperational) <> 0 Then : clrVal = oOpColor
                        Else : clrVal = oNoColor
                        End If
                        oImgSprite.Draw2D(oTex, New Rectangle(16, 16, 16, 16), New Rectangle(0, 0, 16, 16), Point.Empty, 0, New Point(53, 80), clrVal)

                        If (.CritList And elUnitStatus.eHangarOperational) <> 0 Then
                            clrVal = oDestColor
                        ElseIf (.CurrentStatus And elUnitStatus.eHangarOperational) <> 0 Then : clrVal = oOpColor
                        Else : clrVal = oNoColor
                        End If
                        oImgSprite.Draw2D(oTex, New Rectangle(48, 16, 16, 16), New Rectangle(0, 0, 16, 16), Point.Empty, 0, New Point(72, 40), clrVal)

                        If (.CritList And elUnitStatus.eCargoBayOperational) <> 0 Then
                            clrVal = oDestColor
                        ElseIf (.CurrentStatus And elUnitStatus.eCargoBayOperational) <> 0 Then : clrVal = oOpColor
                        Else : clrVal = oNoColor
                        End If
                        oImgSprite.Draw2D(oTex, New Rectangle(32, 0, 16, 16), New Rectangle(0, 0, 16, 16), Point.Empty, 0, New Point(72, 60), clrVal)

                        'If (.CritList And elUnitStatus.eFuelBayOperational) <> 0 Then
                        '    clrVal = oDestColor
                        'Else : clrVal = oOpColor
                        'End If
                        'oImgSprite.Draw2D(oTex, New Rectangle(32, 16, 16, 16), New Rectangle(0, 0, 16, 16), Point.Empty, 0, New Point(72, 80), clrVal)

                        'Weapon Crits
                        If (.CritList And elUnitStatus.eForwardWeapon1) <> 0 Then
                            clrVal = oDestColor
                        ElseIf (.CurrentStatus And elUnitStatus.eForwardWeapon1) <> 0 Then : clrVal = oOpColor
                        Else : clrVal = oNoColor
                        End If
                        oImgSprite.Draw2D(MyBase.moUILib.oInterfaceTexture, New Rectangle(15, 238, 10, 6), New Rectangle(0, 0, 10, 6), Point.Empty, 0, New Point(97, 45), clrVal)

                        If (.CritList And elUnitStatus.eForwardWeapon2) <> 0 Then
                            clrVal = oDestColor
                        ElseIf (.CurrentStatus And elUnitStatus.eForwardWeapon2) <> 0 Then : clrVal = oOpColor
                        Else : clrVal = oNoColor
                        End If
                        oImgSprite.Draw2D(MyBase.moUILib.oInterfaceTexture, New Rectangle(15, 238, 10, 6), New Rectangle(0, 0, 10, 6), Point.Empty, 0, New Point(108, 45), clrVal)

                        If (.CritList And elUnitStatus.eLeftWeapon1) <> 0 Then
                            clrVal = oDestColor
                        ElseIf (.CurrentStatus And elUnitStatus.eLeftWeapon1) <> 0 Then : clrVal = oOpColor
                        Else : clrVal = oNoColor
                        End If
                        oImgSprite.Draw2D(MyBase.moUILib.oInterfaceTexture, New Rectangle(0, 236, 6, 10), New Rectangle(0, 0, 6, 10), Point.Empty, 0, New Point(92, 55), clrVal)

                        If (.CritList And elUnitStatus.eLeftWeapon2) <> 0 Then
                            clrVal = oDestColor
                        ElseIf (.CurrentStatus And elUnitStatus.eLeftWeapon2) <> 0 Then : clrVal = oOpColor
                        Else : clrVal = oNoColor
                        End If
                        oImgSprite.Draw2D(MyBase.moUILib.oInterfaceTexture, New Rectangle(0, 236, 6, 10), New Rectangle(0, 0, 6, 10), Point.Empty, 0, New Point(92, 71), clrVal)

                        If (.CritList And elUnitStatus.eAftWeapon1) <> 0 Then
                            clrVal = oDestColor
                        ElseIf (.CurrentStatus And elUnitStatus.eAftWeapon1) <> 0 Then : clrVal = oOpColor
                        Else : clrVal = oNoColor
                        End If
                        oImgSprite.Draw2D(MyBase.moUILib.oInterfaceTexture, New Rectangle(27, 238, 10, 6), New Rectangle(0, 0, 10, 6), Point.Empty, 0, New Point(97, 85), clrVal)

                        If (.CritList And elUnitStatus.eAftWeapon2) <> 0 Then
                            clrVal = oDestColor
                        ElseIf (.CurrentStatus And elUnitStatus.eAftWeapon2) <> 0 Then : clrVal = oOpColor
                        Else : clrVal = oNoColor
                        End If
                        oImgSprite.Draw2D(MyBase.moUILib.oInterfaceTexture, New Rectangle(27, 238, 10, 6), New Rectangle(0, 0, 10, 6), Point.Empty, 0, New Point(109, 85), clrVal)

                        If (.CritList And elUnitStatus.eRightWeapon1) <> 0 Then
                            clrVal = oDestColor
                        ElseIf (.CurrentStatus And elUnitStatus.eRightWeapon1) <> 0 Then : clrVal = oOpColor
                        Else : clrVal = oNoColor
                        End If
                        oImgSprite.Draw2D(MyBase.moUILib.oInterfaceTexture, New Rectangle(8, 236, 6, 10), New Rectangle(0, 0, 6, 10), Point.Empty, 0, New Point(118, 55), clrVal)

                        If (.CritList And elUnitStatus.eRightWeapon2) <> 0 Then
                            clrVal = oDestColor
                        ElseIf (.CurrentStatus And elUnitStatus.eRightWeapon2) <> 0 Then : clrVal = oOpColor
                        Else : clrVal = oNoColor
                        End If
                        oImgSprite.Draw2D(MyBase.moUILib.oInterfaceTexture, New Rectangle(8, 236, 6, 10), New Rectangle(0, 0, 6, 10), Point.Empty, 0, New Point(118, 71), clrVal)

                        ''Fuel Bar Fill Color
                        'Dim fFillX As Single
                        'Dim fFillY As Single
                        'Dim rcSrc As Rectangle
                        'rcSrc.Location = New Point(192, 0)
                        'rcSrc.Width = 62 : rcSrc.Height = 64

                        'fFillX = 2.0F * (rcSrc.Width / 61.0F)
                        'fFillY = 99.0F * (rcSrc.Height / 10.0F)
                        ''oImgSprite.Draw2D(MyBase.moUILib.oInterfaceTexture, rcSrc, New Rectangle(0, 0, 61, 10), System.Drawing.Point.Empty, 0, New Point(CInt(fFillX), CInt(fFillY)), System.Drawing.Color.FromArgb(255, 128, 128, 0))

                        ''Ammo Bar Fill Color
                        'fFillX = 65 * (rcSrc.Width / 61.0F)
                        'oImgSprite.Draw2D(MyBase.moUILib.oInterfaceTexture, rcSrc, New Rectangle(0, 0, 61, 10), System.Drawing.Point.Empty, 0, New Point(CInt(fFillX), CInt(fFillY)), System.Drawing.Color.FromArgb(255, 128, 0, 128))

                        'Dim fTemp As Single
                        ' ''Now, fuel Bar amount
                        ''Dim fTemp As Single = (.Fuel_Cap / 100.0F)
                        ''If fTemp > 1 Then fTemp = 1
                        ''If fTemp < 0 Then fTemp = 0
                        ''fTemp *= 61
                        ''If fTemp = 0 Then fFillX = 2 Else fFillX = 2 * (rcSrc.Width / fTemp)
                        ''oImgSprite.Draw2D(MyBase.moUILib.oInterfaceTexture, rcSrc, New Rectangle(0, 0, CInt(fTemp), 10), System.Drawing.Point.Empty, 0, New Point(CInt(fFillX), CInt(fFillY)), System.Drawing.Color.FromArgb(255, 255, 255, 0))

                        ''now, the ammo bar amount
                        'fTemp = (.Ammo_Cap / 100.0F)
                        'If fTemp > 1 Then fTemp = 1
                        'If fTemp < 0 Then fTemp = 0
                        'fTemp *= 61
                        'If fTemp = 0 Then fFillX = 65 Else fFillX = 65 * (rcSrc.Width / fTemp)
                        'oImgSprite.Draw2D(MyBase.moUILib.oInterfaceTexture, rcSrc, New Rectangle(0, 0, CInt(fTemp), 10), System.Drawing.Point.Empty, 0, New Point(CInt(fFillX), CInt(fFillY)), System.Drawing.Color.FromArgb(255, 255, 0, 255))

                        'Experience Insignia
                        If .ObjTypeID <> ObjectType.eFacility Then
                            Dim lXPRank As Int32 = Math.Abs((CInt(.Exp_Level) - 1) \ 25) '  CInt(Math.Floor(Math.Abs(.Exp_Level - 1) / 25))
                            Select Case lXPRank
                                Case 0  'Green
                                    oImgSprite.Draw2D(MyBase.moUILib.oInterfaceTexture, grc_UI(elInterfaceRectangle.eXPRank_EmptyDot), New Rectangle(0, 0, 16, 16), Point.Empty, 0, New Point(Me.Width - 17, lblOwner.Top), Color.White)
                                Case 1  'Trained
                                    oImgSprite.Draw2D(MyBase.moUILib.oInterfaceTexture, grc_UI(elInterfaceRectangle.eXPRank_SolidDot), New Rectangle(0, 0, 16, 16), Point.Empty, 0, New Point(Me.Width - 17, lblOwner.Top), Color.White)
                                Case 2  'Experienced
                                    oImgSprite.Draw2D(MyBase.moUILib.oInterfaceTexture, grc_UI(elInterfaceRectangle.eXPRank_SolidDot), New Rectangle(0, 0, 16, 16), Point.Empty, 0, New Point(Me.Width - 17, lblOwner.Top - 4), Color.White)
                                    oImgSprite.Draw2D(MyBase.moUILib.oInterfaceTexture, grc_UI(elInterfaceRectangle.eXPRank_SolidDot), New Rectangle(0, 0, 16, 16), Point.Empty, 0, New Point(Me.Width - 17, lblOwner.Top + 4), Color.White)
                                Case 3  'Adept
                                    oImgSprite.Draw2D(MyBase.moUILib.oInterfaceTexture, grc_UI(elInterfaceRectangle.eXPRank_SolidDot), New Rectangle(0, 0, 16, 16), Point.Empty, 0, New Point(Me.Width - 17, lblOwner.Top - 4), Color.White)
                                    oImgSprite.Draw2D(MyBase.moUILib.oInterfaceTexture, grc_UI(elInterfaceRectangle.eXPRank_SolidDot), New Rectangle(0, 0, 16, 16), Point.Empty, 0, New Point(Me.Width - 21, lblOwner.Top + 4), Color.White)
                                    oImgSprite.Draw2D(MyBase.moUILib.oInterfaceTexture, grc_UI(elInterfaceRectangle.eXPRank_SolidDot), New Rectangle(0, 0, 16, 16), Point.Empty, 0, New Point(Me.Width - 13, lblOwner.Top + 4), Color.White)
                                Case 4  'Veteran
                                    oImgSprite.Draw2D(MyBase.moUILib.oInterfaceTexture, grc_UI(elInterfaceRectangle.eXPRank_SolidDot), New Rectangle(0, 0, 16, 16), Point.Empty, 0, New Point(Me.Width - 17, lblOwner.Top - 4), Color.White)
                                    oImgSprite.Draw2D(MyBase.moUILib.oInterfaceTexture, grc_UI(elInterfaceRectangle.eXPRank_Arrow), New Rectangle(0, 0, 16, 16), Point.Empty, 0, New Point(Me.Width - 17, lblOwner.Top + 4), Color.White)
                                Case 5  'Ace
                                    oImgSprite.Draw2D(MyBase.moUILib.oInterfaceTexture, grc_UI(elInterfaceRectangle.eXPRank_Arrow), New Rectangle(0, 0, 16, 16), Point.Empty, 0, New Point(Me.Width - 17, lblOwner.Top - 4), Color.White)
                                    oImgSprite.Draw2D(MyBase.moUILib.oInterfaceTexture, grc_UI(elInterfaceRectangle.eXPRank_Arrow), New Rectangle(0, 0, 16, 16), Point.Empty, 0, New Point(Me.Width - 17, lblOwner.Top + 4), Color.White)
                                Case 6  'Top Ace
                                    oImgSprite.Draw2D(MyBase.moUILib.oInterfaceTexture, grc_UI(elInterfaceRectangle.eXPRank_Arrow), New Rectangle(0, 0, 16, 16), Point.Empty, 0, New Point(Me.Width - 17, lblOwner.Top + 6), Color.White)
                                    oImgSprite.Draw2D(MyBase.moUILib.oInterfaceTexture, grc_UI(elInterfaceRectangle.eXPRank_Arrow), New Rectangle(0, 0, 16, 16), Point.Empty, 0, New Point(Me.Width - 17, lblOwner.Top), Color.White)
                                    oImgSprite.Draw2D(MyBase.moUILib.oInterfaceTexture, grc_UI(elInterfaceRectangle.eXPRank_Arrow), New Rectangle(0, 0, 16, 16), Point.Empty, 0, New Point(Me.Width - 17, lblOwner.Top - 6), Color.White)
                                Case 7  'Distinguished
                                    oImgSprite.Draw2D(MyBase.moUILib.oInterfaceTexture, grc_UI(elInterfaceRectangle.eXPRank_SolidDot), New Rectangle(0, 0, 16, 16), Point.Empty, 0, New Point(Me.Width - 17, lblOwner.Top - 4), Color.White)
                                    oImgSprite.Draw2D(MyBase.moUILib.oInterfaceTexture, grc_UI(elInterfaceRectangle.eXPRank_Bar), New Rectangle(0, 0, 16, 16), Point.Empty, 0, New Point(Me.Width - 17, lblOwner.Top + 4), Color.White)
                                Case 8  'Revered
                                    oImgSprite.Draw2D(MyBase.moUILib.oInterfaceTexture, grc_UI(elInterfaceRectangle.eXPRank_Bar), New Rectangle(0, 0, 16, 16), Point.Empty, 0, New Point(Me.Width - 17, lblOwner.Top - 4), Color.White)
                                    oImgSprite.Draw2D(MyBase.moUILib.oInterfaceTexture, grc_UI(elInterfaceRectangle.eXPRank_Bar), New Rectangle(0, 0, 16, 16), Point.Empty, 0, New Point(Me.Width - 17, lblOwner.Top + 4), Color.White)
                                Case Else   'Elite
                                    oImgSprite.Draw2D(MyBase.moUILib.oInterfaceTexture, grc_UI(elInterfaceRectangle.eXPRank_Bar), New Rectangle(0, 0, 16, 16), Point.Empty, 0, New Point(Me.Width - 17, lblOwner.Top + 6), Color.White)
                                    oImgSprite.Draw2D(MyBase.moUILib.oInterfaceTexture, grc_UI(elInterfaceRectangle.eXPRank_Bar), New Rectangle(0, 0, 16, 16), Point.Empty, 0, New Point(Me.Width - 17, lblOwner.Top), Color.White)
                                    oImgSprite.Draw2D(MyBase.moUILib.oInterfaceTexture, grc_UI(elInterfaceRectangle.eXPRank_Bar), New Rectangle(0, 0, 16, 16), Point.Empty, 0, New Point(Me.Width - 17, lblOwner.Top - 6), Color.White)
                            End Select
                        End If

                        If .oUnitDef Is Nothing = False AndAlso .oUnitDef.MaxSpeed > 0 Then UpdateETALabel()

                        If goCurrentPlayer.oGuild Is Nothing = False Then
                            clrVal = System.Drawing.Color.FromArgb(255, 64, 64, 64)
                            If .ObjTypeID = ObjectType.eUnit AndAlso .OwnerID = glPlayerID AndAlso gbAliased = False Then
                                If (.CurrentStatus And elUnitStatus.eGuildAsset) <> 0 Then
                                    clrVal = System.Drawing.Color.FromArgb(255, 0, 255, 0)
                                Else
                                    clrVal = System.Drawing.Color.FromArgb(255, 255, 0, 0)
                                End If
                            End If
                            oImgSprite.Draw2D(MyBase.moUILib.oInterfaceTexture, grc_UI(elInterfaceRectangle.eKeyButton), rcGuildAsset, Point.Empty, 0, rcGuildAsset.Location, clrVal)
                        End If

                        If .OwnerID = glPlayerID AndAlso .oUnitDef Is Nothing = False AndAlso .oUnitDef.MaxSpeed > 0 Then
                            If mbAlertFinalDest = True Then
                                oImgSprite.Draw2D(MyBase.moUILib.oInterfaceTexture, grc_UI(elInterfaceRectangle.eAlarm), rcAlertFinalDest, Point.Empty, 0, rcAlertFinalDest.Location, System.Drawing.Color.FromArgb(255, 0, 255, 0))
                            Else
                                oImgSprite.Draw2D(MyBase.moUILib.oInterfaceTexture, grc_UI(elInterfaceRectangle.eAlarm), rcAlertFinalDest, Point.Empty, 0, rcAlertFinalDest.Location, System.Drawing.Color.FromArgb(255, 255, 0, 0))
                            End If
                        End If

                    End If

                    oImgSprite.End()

                    ''Fuel and Ammo Status BORDERs
                    'If .OwnerID = glPlayerID OrElse .OwnerID = gl_HARDCODE_PIRATE_PLAYER_ID Then
                    '    MyBase.moBorderLine.Width = 1
                    '    MyBase.moBorderLine.Antialias = True
                    '    MyBase.moBorderLine.Begin()
                    '    Dim oVerts(4) As Vector2
                    '    oVerts(0).X = 2 : oVerts(0).Y = 99
                    '    oVerts(1).X = 62 : oVerts(1).Y = 99
                    '    oVerts(2).X = 62 : oVerts(2).Y = 108
                    '    oVerts(3).X = 2 : oVerts(3).Y = 108
                    '    oVerts(4).X = 2 : oVerts(4).Y = 99
                    '    'MyBase.moBorderLine.Draw(oVerts, System.Drawing.Color.FromArgb(255, 255, 255, 0))
                    '    oVerts(0).X = 65
                    '    oVerts(1).X = 125
                    '    oVerts(2).X = 125
                    '    oVerts(3).X = 65
                    '    oVerts(4).X = 65
                    '    MyBase.moBorderLine.Draw(oVerts, System.Drawing.Color.FromArgb(255, 255, 0, 255))
                    '    MyBase.moBorderLine.End()
                    'End If
                End Using
            End With
        Catch

        End Try
    End Sub

    Private Sub RenderHPBars(ByRef oSpr As Sprite)
        Dim fPerc As Single

        With goCurrentEnvir.oEntity(mlEntityIndex)
            fPerc = .yShieldHP / 100.0F
            DoMultiColorFill(New Rectangle(0, 0, 47, 6), Color.Red, New Point(3, 39), oSpr)
            If fPerc <> 0 Then
                DoMultiColorFill(New Rectangle(0, 0, CInt(47 * fPerc), 6), System.Drawing.Color.FromArgb(255, 0, 128, 255), New Point(3, 39), oSpr)
            End If

			If .yArmorHP(0) = 255 Then fPerc = 0 Else fPerc = .yArmorHP(0) / 100.0F
            DoMultiColorFill(New Rectangle(0, 0, 47, 6), Color.Red, New Point(3, 46), oSpr)
            If fPerc <> 0 Then
                DoMultiColorFill(New Rectangle(0, 0, CInt(47 * fPerc), 6), System.Drawing.Color.FromArgb(255, 0, 255, 0), New Point(3, 46), oSpr)
            End If

			If .yArmorHP(1) = 255 Then fPerc = 0 Else fPerc = .yArmorHP(1) / 100.0F
            DoMultiColorFill(New Rectangle(0, 0, 6, 33), Color.Red, New Point(3, 53), oSpr)
            If fPerc <> 0 Then
                Dim lHeight As Int32 = CInt(33 * fPerc)
                Dim lTop As Int32 = 53 + (33 - lHeight)
                DoMultiColorFill(New Rectangle(0, 0, 6, lHeight), System.Drawing.Color.FromArgb(255, 0, 255, 0), New Point(3, lTop), oSpr)
            End If

			If .yArmorHP(2) = 255 Then fPerc = 0 Else fPerc = .yArmorHP(2) / 100.0F
            DoMultiColorFill(New Rectangle(0, 0, 47, 6), Color.Red, New Point(3, 87), oSpr)
            If fPerc <> 0 Then
                DoMultiColorFill(New Rectangle(0, 0, CInt(47 * fPerc), 6), System.Drawing.Color.FromArgb(255, 0, 255, 0), New Point(3, 87), oSpr)
            End If

			If .yArmorHP(3) = 255 Then fPerc = 0 Else fPerc = .yArmorHP(3) / 100.0F
            DoMultiColorFill(New Rectangle(0, 0, 6, 33), Color.Red, New Point(44, 53), oSpr)
            If fPerc <> 0 Then
                Dim lHeight As Int32 = CInt(33 * fPerc)
                Dim lTop As Int32 = 53 + (33 - lHeight)
                DoMultiColorFill(New Rectangle(0, 0, 6, lHeight), System.Drawing.Color.FromArgb(255, 0, 255, 0), New Point(44, lTop), oSpr)
            End If

            fPerc = .yStructureHP / 100.0F
            DoMultiColorFill(New Rectangle(0, 0, 33, 33), Color.Red, New Point(10, 53), oSpr)
            If fPerc <> 0 Then
                Dim lHeight As Int32 = CInt(33 * fPerc)
                Dim lTop As Int32 = 53 + (33 - lHeight)
                DoMultiColorFill(New Rectangle(0, 0, 33, lHeight), System.Drawing.Color.FromArgb(255, 0, 255, 0), New Point(10, lTop), oSpr)
            End If
        End With
    End Sub

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

    Private Sub txtEntityName_OnKeyDown(ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtEntityName.OnKeyDown
        mbNameChangeInProgress = True
        If e.KeyCode = Keys.Enter AndAlso txtEntityName.Locked = False Then
            If mlEntityIndex <> -1 AndAlso goCurrentEnvir Is Nothing = False AndAlso goCurrentEnvir.lEntityIdx(mlEntityIndex) <> -1 AndAlso goCurrentEnvir.oEntity(mlEntityIndex).OwnerID = glPlayerID Then
                Try
                    'Submit a name change...
                    Dim yMsg(27) As Byte
                    System.BitConverter.GetBytes(GlobalMessageCode.eSetEntityName).CopyTo(yMsg, 0)
                    goCurrentEnvir.oEntity(mlEntityIndex).GetGUIDAsString.CopyTo(yMsg, 2)
                    System.Text.ASCIIEncoding.ASCII.GetBytes(Mid$(txtEntityName.Caption, 1, 20)).CopyTo(yMsg, 8)
                    MyBase.moUILib.SendMsgToPrimary(yMsg)

                    SetCacheObjectValue(goCurrentEnvir.oEntity(mlEntityIndex).ObjectID, goCurrentEnvir.oEntity(mlEntityIndex).ObjTypeID, txtEntityName.Caption)

                    BPMetrics.MetricMgr.AddActivity(BPMetrics.eInterface.eEntityChangeName, goCurrentEnvir.oEntity(mlEntityIndex).ObjectID, goCurrentEnvir.oEntity(mlEntityIndex).ObjTypeID)

                    goCurrentEnvir.oEntity(mlEntityIndex).EntityName = txtEntityName.Caption
                    mbNameChangeInProgress = False
                Catch
                End Try
            End If
        End If
    End Sub

    Private Sub frmSingleSelect_WindowMoved() Handles Me.WindowMoved
        muSettings.SelectLocX = Me.Left
        muSettings.SelectLocY = Me.Top
    End Sub

    Private Sub chkAlertFinalDest_Click() 'Handles chkAlertFinalDest.Click
        Dim oEnvir As BaseEnvironment = goCurrentEnvir
        If oEnvir Is Nothing = False AndAlso mlEntityIndex > -1 AndAlso mlEntityIndex <= oEnvir.lEntityUB Then
            Dim oEntity As BaseEntity = oEnvir.oEntity(mlEntityIndex)
            If oEntity Is Nothing = False Then

                mbAlertFinalDest = Not mbAlertFinalDest

                If mbAlertFinalDest = False Then
                    UnitAlertReceived(oEntity.ObjectID, oEntity.ObjTypeID)
                Else
                    ReDim Preserve muUnitAlerts(mlUnitAlertUB + 1)
                    mlUnitAlertUB += 1
                    muUnitAlerts(mlUnitAlertUB).lObjID = oEntity.ObjectID
                    muUnitAlerts(mlUnitAlertUB).iObjTypeID = oEntity.ObjTypeID
                End If

                Me.IsDirty = True

                Dim yMsg(8) As Byte
                Dim lPos As Int32 = 0
                System.BitConverter.GetBytes(GlobalMessageCode.eAlertDestinationReached).CopyTo(yMsg, lPos) : lPos += 2
                oEntity.GetGUIDAsString.CopyTo(yMsg, lPos) : lPos += 6

                If mbAlertFinalDest = True Then yMsg(lPos) = 1 Else yMsg(lPos) = 0
                lPos += 1

                MyBase.moUILib.SendMsgToRegion(yMsg)
            End If
        End If
    End Sub
    Private Sub chkGuildAsset_Click()
        If gbAliased = True Then
            goUILib.AddNotification("Aliases may not toggle guild share status of units", Color.Red, -1, -1, -1, -1, -1, -1)
            If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eUserInterface, Nothing, Nothing)
            Return
        End If
        'Dim oFrm As New frmMsgBox(goUILib, "Activating guild share for this unit requires an immediate upkeep payment of warpoints, in addition to a payment every 24 hours. Disabling guild share status removes further payments but will require the payment again should guild share be re-enabled. " & vbCrLf & vbCrLf & "Do you wish to change the status?", MsgBoxStyle.YesNo, "Confirm Guild Share Status")
        'oFrm.Visible = True
        'AddHandler oFrm.DialogClosed, AddressOf ConfirmGuildAsset
        ConfirmGuildAsset(MsgBoxResult.Yes)
    End Sub

    Private Sub ConfirmGuildAsset(ByVal lResult As Microsoft.VisualBasic.MsgBoxResult)
        If lResult = MsgBoxResult.Yes Then
            Dim oEnvir As BaseEnvironment = goCurrentEnvir
            If goCurrentPlayer Is Nothing OrElse goCurrentPlayer.yPlayerPhase <> 255 OrElse goCurrentPlayer.lAccountStatus <> AccountStatusType.eActiveAccount Then Return

            If oEnvir Is Nothing = False AndAlso mlEntityIndex > -1 AndAlso mlEntityIndex <= oEnvir.lEntityUB Then
                Dim oEntity As BaseEntity = oEnvir.oEntity(mlEntityIndex)
                If oEntity Is Nothing = False Then
                    If oEntity.ObjTypeID = ObjectType.eFacility OrElse oEntity.OwnerID <> glPlayerID Then Return
                    'Dim lWPUpkeep As Int32 = GetUnitWarpointUpkeep(oEntity.ObjectID)
                    'If lWPUpkeep = -1 Then Return

                    Dim lStatusVal As Int32 = 0
                    If (oEntity.CurrentStatus And elUnitStatus.eGuildAsset) <> 0 Then
                        oEntity.CurrentStatus = oEntity.CurrentStatus Xor elUnitStatus.eGuildAsset
                        lStatusVal = -elUnitStatus.eGuildAsset
                    Else
                        'If CLng(lWPUpkeep) > goCurrentPlayer.blWarpoints Then
                        '    goUILib.AddNotification("Insufficient warpoints to enable guild sharing this unit!", Color.Red, -1, -1, -1, -1, -1, -1)
                        '    Return
                        'End If
                        oEntity.CurrentStatus = oEntity.CurrentStatus Or elUnitStatus.eGuildAsset
                        lStatusVal = elUnitStatus.eGuildAsset
                    End If
                    Me.IsDirty = True

                    Dim yMsg(11) As Byte
                    Dim lPos As Int32 = 0
                    System.BitConverter.GetBytes(GlobalMessageCode.eSetEntityStatus).CopyTo(yMsg, lPos) : lPos += 2
                    oEntity.GetGUIDAsString.CopyTo(yMsg, lPos) : lPos += 6
                    System.BitConverter.GetBytes(lStatusVal).CopyTo(yMsg, lPos) : lPos += 4

                    MyBase.moUILib.SendMsgToRegion(yMsg)
                End If
            End If
        End If
    End Sub

    Public Function GetSelectedItemsDistance(ByVal fX As Single, ByVal fZ As Single) As Single
        Try
            Dim oEnvir As BaseEnvironment = goCurrentEnvir
            If oEnvir Is Nothing = False AndAlso mlEntityIndex > -1 AndAlso oEnvir.lEntityIdx(mlEntityIndex) > -1 Then
                Dim oEntity As BaseEntity = oEnvir.oEntity(mlEntityIndex)
                If oEntity Is Nothing = False Then
                    Dim fDX As Single = fX - oEntity.LocX
                    Dim fDZ As Single = fZ - oEntity.LocZ
                    fDX *= fDX
                    fDZ *= fDZ
                    Dim fDist As Single = CSng(Math.Sqrt(fDX + fDZ)) / gl_FINAL_GRID_SQUARE_SIZE
                    If oEntity.oMesh Is Nothing = False Then
                        Return Math.Max(0, fDist - oEntity.oMesh.RangeOffset)
                    Else
                        Return fDist
                    End If
                End If
            End If
        Catch
        End Try
        Return Single.MinValue
    End Function
End Class
