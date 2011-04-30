Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Cache Selection Window... when a mineral cache is selected, this is shown... or even hover (a la tooltip style)
Public Class frmCacheSelect
    Inherits UIWindow

    Private lblTitle As UILabel
    Private lnDivider As UILine
    Private lblMinName As UILabel
    Private lblConcOverQuant As UILabel

    Private msw_Delay As Stopwatch

    Private mbComponentCache As Boolean = False
    Private mlCompID As Int32 = -1
    Private miCompTypeID As Int16 = -1

    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)

        TutorialAlert.CheckTrigger(TutorialAlert.eTriggerType.MineralCacheHover)

        'frmCacheSelect initial props
        With Me
            .lWindowMetricID = BPMetrics.eWindow.eCacheSelect
            .ControlName = "frmCacheSelect"
            .Left = 217
            .Top = 171
            .Width = 182
            .Height = 66
            .Enabled = True
            .Visible = True
            '.BorderColor = System.Drawing.Color.FromArgb(-16711681)
            .BorderColor = muSettings.InterfaceBorderColor
            '.FillColor = System.Drawing.Color.FromArgb(-12549952)
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = False
            .BorderLineWidth = 2
            .ClickThru = True
        End With
        Debug.Write(Me.ControlName & " Newed" & vbCrLf)
        'lblTitle initial props
        lblTitle = New UILabel(oUILib)
        With lblTitle
            .ControlName = "lblTitle"
            .Left = 0
            .Top = 3
            .Width = 182
            .Height = 16
            .Enabled = True
            .Visible = True
            .Caption = "Title Bar (cache type)"
            '.ForeColor = System.Drawing.Color.FromArgb(-16711681)
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.Center Or DrawTextFormat.VerticalCenter
        End With
        Me.AddChild(CType(lblTitle, UIControl))

        'lnDivider initial props
        lnDivider = New UILine(oUILib)
        With lnDivider
            .ControlName = "lnDivider"
            .Left = Me.BorderLineWidth
            .Top = 19
            .Width = Me.Width - Me.BorderLineWidth
            .Height = 0
            .Enabled = True
            .Visible = True
            '.BorderColor = System.Drawing.Color.FromArgb(-16711681)
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(lnDivider, UIControl))

        'lblMinName initial props
        lblMinName = New UILabel(oUILib)
        With lblMinName
            .ControlName = "lblMinName"
            .Left = 6
            .Top = 25
            .Width = 182
            .Height = 16
            .Enabled = True
            .Visible = True
            .Caption = "Mineral Name"
            '.ForeColor = System.Drawing.Color.FromArgb(-16711681)
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.VerticalCenter
        End With
        Me.AddChild(CType(lblMinName, UIControl))

        'lblConcOverQuant initial props
        lblConcOverQuant = New UILabel(oUILib)
        With lblConcOverQuant
            .ControlName = "lblConcOverQuant"
            .Left = 6
            .Top = 45
            .Width = 182
            .Height = 16
            .Enabled = True
            .Visible = True
            .Caption = "Concentration/Quantity"
            '.ForeColor = System.Drawing.Color.FromArgb(-16711681)
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.VerticalCenter
        End With
        Me.AddChild(CType(lblConcOverQuant, UIControl))

        oUILib.RemoveWindow(Me.ControlName)

        'Now, add me...
        oUILib.AddWindow(Me)

        msw_Delay = Stopwatch.StartNew

        If goTutorial Is Nothing = False AndAlso goTutorial.TutorialOn = True Then goTutorial.HandleTutorialTrigger(TutorialManager.TutorialTriggerType.MineralCacheHover)
    End Sub

    'returns whether the mineral cache details are available, if false... the caller must make a call to the 
    '  the message system to request the details
    Public Function SetProps(ByRef oCache As MineralCache) As Boolean
        Try
            If oCache Is Nothing Then Return True

            If oCache.ObjTypeID = ObjectType.eMineralCache Then
                If oCache.CacheTypeID = MineralCacheType.eMineable Then
                    lblTitle.Caption = "Mineable Mineral Deposit"
                Else
                    lblTitle.Caption = "Minerals"
                End If

                If oCache.oMineral Is Nothing Then
                    lblMinName.Caption = "Unknown Mineral"
                Else : lblMinName.Caption = oCache.oMineral.MineralName
                End If

                If NewTutorialManager.TutorialOn = True Then
                    NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eMineralDepositHover, -1, -1, -1, "")
                End If
            Else
                'Now... this is a component cache... which we display differently
                Dim iTempTypeID As Int16 = MineralCache.GetComponentTypeID(oCache.CacheTypeID)
                Select Case iTempTypeID
                    Case ObjectType.eArmorTech
                        lblTitle.Caption = "Armor Component"
                    Case ObjectType.eEngineTech
                        lblTitle.Caption = "Engine Component"
                        'Case ObjectType.eHangarTech
                        '    lblTitle.Caption = "Hangar Component"
                    Case ObjectType.eHullTech
                        lblTitle.Caption = "Hull Component"
                    Case ObjectType.eRadarTech
                        lblTitle.Caption = "Radar Component"
                    Case ObjectType.eShieldTech
                        lblTitle.Caption = "Shield Component"
                    Case ObjectType.eWeaponTech
                        lblTitle.Caption = "Weapon Component"
                    Case Else
                        lblTitle.Caption = "Unknown Component"
                End Select
                If iTempTypeID <> 0 Then lblMinName.Caption = GetCacheObjectValue(oCache.MineralID, iTempTypeID)
                mlCompID = oCache.MineralID
                miCompTypeID = iTempTypeID
                mbComponentCache = True
            End If

            If msw_Delay.ElapsedMilliseconds > 10000 Then
                lblConcOverQuant.Caption = "Acquiring Details..."
                msw_Delay.Reset()
                msw_Delay.Start()
                AddHandler oCache.DetailsArrived, AddressOf DetailsArrived

                'Caller is responsible for calling the get details call...
                Return False
            ElseIf oCache.ObjTypeID = ObjectType.eComponentCache AndAlso oCache.Quantity <> Int32.MinValue Then
                lblConcOverQuant.Caption = "Quantity: " & oCache.Quantity.ToString("#,###")
                Return True
            ElseIf oCache.Concentration <> Int32.MinValue AndAlso oCache.Quantity <> Int32.MinValue Then
                Dim lAddVal As Int32 = goCurrentPlayer.GetSpecialTechPropertyValue(PlayerSpecialAttributeID.eMineralConcentrationBonus)
                Dim sAddVal As String = ""
                If lAddVal <> 0 Then sAddVal = " (+" & lAddVal & ")"
                lblConcOverQuant.Caption = oCache.Concentration.ToString("#,###") & "/" & oCache.Quantity.ToString("#,###") & sAddVal
                Return True
            End If
        Catch
        End Try
    End Function

    Private Sub DetailsArrived(ByVal oCache As MineralCache)
        'mbRequestedDetails = False
        If mbComponentCache = True Then
            lblConcOverQuant.Caption = "Quantity: " & oCache.Quantity.ToString("#,###")
        Else
            Dim lAddVal As Int32 = goCurrentPlayer.GetSpecialTechPropertyValue(PlayerSpecialAttributeID.eMineralConcentrationBonus)
            Dim sAddVal As String = ""
            If lAddVal <> 0 Then sAddVal = " (+" & lAddVal & ")"
            lblConcOverQuant.Caption = oCache.Concentration.ToString("#,###") & "/" & oCache.Quantity.ToString("#,###") & sAddVal
        End If

        RemoveHandler oCache.DetailsArrived, AddressOf DetailsArrived
    End Sub

    Protected Overrides Sub Finalize()
        Debug.Write(Me.ControlName & " Finalized" & vbCrLf)
        msw_Delay.Stop()
        msw_Delay = Nothing
        MyBase.Finalize()
    End Sub

    Private Sub frmCacheSelect_OnNewFrame() Handles Me.OnNewFrame
        If mbComponentCache = True Then
            If lblMinName.Caption = "Unknown" Then
                lblMinName.Caption = GetCacheObjectValue(mlCompID, miCompTypeID)
            End If
        End If
    End Sub
End Class
