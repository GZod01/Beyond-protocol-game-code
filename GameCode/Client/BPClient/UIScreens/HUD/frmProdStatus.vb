Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Production Status window for displaying the status of an entity's production status
Public Class frmProdStatus
    Inherits UIWindow

    Private Const ml_UPDATE_INTERVAL As Int32 = 3000
    Private mlEntityIndex As Int32 = -1
    Private msw_Delay As Stopwatch

    Private WithEvents shpProdStatus As UIWindow
    Private WithEvents shpProdStatusBack As UIWindow
    Private lblStatusPerc As UILabel
    Private lblProducing As UILabel

    Private mbUseChild As Boolean = False
    Private mlChildID As Int32 = -1
    Private miChildTypeID As Int16 = -1

    Private mlLastProdID As Int32 = -1
    Private miLastProdTypeID As Int16 = -1

    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)

        'frmProdStatus initial props
        Dim lFormWidth As Int32 = 250
        Dim lshpProdStatusBack As Int32 = 200
        Dim lblStatusPercLeft As Int32 = 200
        Dim lblProducingWidth As Int32 = 198

        With Me
            .lWindowMetricID = BPMetrics.eWindow.eProdStatus
            .ControlName = "frmProdStatus"
            .Left = 120
            .Top = 0

            If muSettings.ProdStatusLocX <> -1 AndAlso NewTutorialManager.TutorialOn = False Then .Left = muSettings.ProdStatusLocX
            If muSettings.ProdStatusLocY <> -1 AndAlso NewTutorialManager.TutorialOn = False Then .Top = muSettings.ProdStatusLocY
            If muSettings.ProdStatusWidth <> -1 AndAlso NewTutorialManager.TutorialOn = False Then
                lFormWidth = muSettings.ProdStatusWidth + 50 '275 ' 250
                lshpProdStatusBack = muSettings.ProdStatusWidth  '225 ' 200
                lblStatusPercLeft = muSettings.ProdStatusWidth  '225 '200
                lblProducingWidth = muSettings.ProdStatusWidth - 2 '223 ' 198
            End If
            .Width = lFormWidth ' 275 ' 250
            .Height = 14
            If .Left + .Width > oUILib.oDevice.PresentationParameters.BackBufferWidth Then .Left = oUILib.oDevice.PresentationParameters.BackBufferWidth - .Width
            If .Top + .Height > oUILib.oDevice.PresentationParameters.BackBufferHeight Then .Top = oUILib.oDevice.PresentationParameters.BackBufferHeight - .Height
            .Enabled = True
            .Visible = True
            '.BorderColor = System.Drawing.Color.FromArgb(-1)
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = False
            .BorderLineWidth = 2
            .bRoundedBorder = False
        End With

        'shpProdStatusBack initial props
        shpProdStatusBack = New UIWindow(oUILib)
        With shpProdStatusBack
            .ControlName = "shpProdStatusBack"
            .Left = 1
            .Top = 1
            .Width = lshpProdStatusBack '225 ' 200
            .Height = 12
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = System.Drawing.Color.FromArgb(255, 255, 0, 0)
            .FullScreen = False
            .BorderLineWidth = 1
            .Moveable = False
            .bRoundedBorder = False
        End With
        Me.AddChild(CType(shpProdStatusBack, UIControl))

        'shpProdStatus initial props
        shpProdStatus = New UIWindow(oUILib)
        With shpProdStatus
            .ControlName = "shpProdStatus"
            .Left = 1
            .Top = 1
            .Width = 1
            .Height = 12
            .Enabled = True
            .Visible = True
            .BorderColor = System.Drawing.Color.FromArgb(-16711936)
            .FillColor = System.Drawing.Color.FromArgb(-16711936)
            .FullScreen = False
            .BorderLineWidth = 1
            .Moveable = False
            .bRoundedBorder = False
        End With
        Me.AddChild(CType(shpProdStatus, UIControl))

        'lblStatusPerc initial props
        lblStatusPerc = New UILabel(oUILib)
        With lblStatusPerc
            .ControlName = "lblStatusPerc"
            .Left = lblStatusPercLeft '225 '200
            .Top = 1
            .Width = 50
            .Height = 12
            .Enabled = True
            .Visible = True
            .Caption = "0.00%"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Courier New", 9.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Center
        End With
        Me.AddChild(CType(lblStatusPerc, UIControl))

        'lblProducing initial props
        lblProducing = New UILabel(oUILib)
        With lblProducing
            .ControlName = "lblProducing"
            .Left = 1
            .Top = 1
            .Width = lblProducingWidth '223 ' 198
            .Height = 12
            .Enabled = True
            .Visible = True
            .Caption = ""
            .ForeColor = System.Drawing.Color.FromArgb(255, 0, 0, 0)
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblProducing, UIControl))

        'Ensure that I am the only version of me in the UI Lib
        oUILib.RemoveWindow(Me.ControlName)
        oUILib.AddWindow(Me)

        msw_Delay = Stopwatch.StartNew
    End Sub

    Public Sub SetFromEntity(ByVal lEntityIndex As Int32, Optional ByVal lChildID As Int32 = -1, Optional ByVal iChildTypeID As Int16 = -1)
        'Verify our objects
        If goCurrentEnvir Is Nothing Then Exit Sub
        If lEntityIndex < 0 OrElse lEntityIndex > goCurrentEnvir.lEntityUB Then Exit Sub
        If goCurrentEnvir.oEntity(lEntityIndex) Is Nothing Then Exit Sub

        mlEntityIndex = lEntityIndex
        Me.Visible = True

        mbUseChild = False
        mlChildID = lChildID
        miChildTypeID = iChildTypeID

        mbUseChild = (goCurrentEnvir.oEntity(mlEntityIndex).GetChild(lChildID, iChildTypeID) Is Nothing = False)

        RequestProdStatus()
    End Sub

    Private Sub RequestProdStatus()
        'Send our request for data
        Dim yData(7) As Byte
        System.BitConverter.GetBytes(GlobalMessageCode.eEntityProductionStatus).CopyTo(yData, 0)

        If mbUseChild = True Then
            System.BitConverter.GetBytes(mlChildID).CopyTo(yData, 2)
            System.BitConverter.GetBytes(miChildTypeID).CopyTo(yData, 6)
        Else
            System.BitConverter.GetBytes(goCurrentEnvir.oEntity(mlEntityIndex).ObjectID).CopyTo(yData, 2)
            System.BitConverter.GetBytes(goCurrentEnvir.oEntity(mlEntityIndex).ObjTypeID).CopyTo(yData, 6)
        End If

        MyBase.moUILib.SendMsgToPrimary(yData)
        Erase yData
    End Sub

    Private Sub frmProdStatus_OnNewFrame() Handles Me.OnNewFrame
		Dim fPerc As Single

		Dim oEnvir As BaseEnvironment = goCurrentEnvir
		If oEnvir Is Nothing Then
			MyBase.moUILib.RemoveWindow(Me.ControlName)
			Return
		End If

		If mlEntityIndex = -1 OrElse oEnvir.lEntityIdx(mlEntityIndex) = -1 Then
			MyBase.moUILib.RemoveWindow(Me.ControlName)
			Return
		End If

		Dim oEntity As BaseEntity = oEnvir.oEntity(mlEntityIndex)
		If oEntity Is Nothing OrElse oEntity.bSelected = False Then
			MyBase.moUILib.RemoveWindow(Me.ControlName)
			Return
		End If

		If msw_Delay.ElapsedMilliseconds > ml_UPDATE_INTERVAL Then
			RequestProdStatus()
			msw_Delay.Reset()
			msw_Delay.Start()
		End If

        Dim lNewProdID As Int32 = mlLastProdID
        Dim iNewProdTypeID As Int16 = miLastProdTypeID

        If mbUseChild = False Then
			fPerc = oEntity.GetProductionStatus()
			lNewProdID = oEntity.lProducingID
			iNewProdTypeID = oEntity.iProducingTypeID
        Else
			Dim oChild As StationChild = oEntity.GetChild(mlChildID, miChildTypeID)
            If oChild Is Nothing = False Then
                fPerc = oChild.GetProductionStatus
                lNewProdID = oChild.lProdID
                iNewProdTypeID = oChild.iProdTypeID
            Else : fPerc = Single.NegativeInfinity
            End If
        End If

        If fPerc = Single.NegativeInfinity Then fPerc = 0.0F

        If fPerc <= 0.0F Then
            If lblStatusPerc.Caption <> "0.00%" Then lblStatusPerc.Caption = "0.00%"
            If lblProducing.Caption <> "" Then Me.lblProducing.Caption = ""
        Else
            If fPerc > 1.0F Then
                fPerc = 1.0F
            ElseIf fPerc < 0 Then
                fPerc = 0
            End If
            shpProdStatus.FillColor = System.Drawing.Color.FromArgb(255, 0, 255, 0)
            Dim lWidth As Int32 = CInt(fPerc * shpProdStatusBack.Width)
            If lWidth < 1 Then lWidth = 1
            shpProdStatus.Width = lWidth

            If lNewProdID <> mlLastProdID OrElse iNewProdTypeID <> miLastProdTypeID Then
                mlLastProdID = lNewProdID
                miLastProdTypeID = iNewProdTypeID

                If mlLastProdID = -1 OrElse miLastProdTypeID = -1 Then
                    Me.lblProducing.Caption = ""
				Else
					Try
						'Now... set our caption...
						If miLastProdTypeID = ObjectType.eFacilityDef OrElse miLastProdTypeID = ObjectType.eUnitDef OrElse miLastProdTypeID = ObjectType.eEnlisted OrElse miLastProdTypeID = ObjectType.eOfficers Then
							For X As Int32 = 0 To glEntityDefUB
								If glEntityDefIdx(X) = mlLastProdID AndAlso goEntityDefs(X).ObjTypeID = miLastProdTypeID Then
									Me.lblProducing.Caption = goEntityDefs(X).DefName
									Exit For
								End If
							Next X
						ElseIf miLastProdTypeID = ObjectType.eMineral Then
							For X As Int32 = 0 To goCurrentPlayer.mlTechUB
								If goCurrentPlayer.moTechs(X) Is Nothing = False AndAlso goCurrentPlayer.moTechs(X).ObjTypeID = ObjectType.eAlloyTech Then
									If CType(goCurrentPlayer.moTechs(X), AlloyTech).AlloyResultID = mlLastProdID Then
										Me.lblProducing.Caption = goCurrentPlayer.moTechs(X).GetComponentName
										Exit For
									End If
								End If
							Next X
						ElseIf miLastProdTypeID = ObjectType.eMineralTech Then
							For X As Int32 = 0 To glMineralUB
								If glMineralIdx(X) = mlLastProdID Then
									If goMinerals(X).bDiscovered = True Then
										lblProducing.Caption = "Study " & goMinerals(X).MineralName
									Else : lblProducing.Caption = "Discover " & goMinerals(X).MineralName
									End If
									Exit For
								End If
							Next X
						Else
							For X As Int32 = 0 To goCurrentPlayer.mlTechUB
								If goCurrentPlayer.moTechs(X) Is Nothing = False AndAlso goCurrentPlayer.moTechs(X).ObjectID = mlLastProdID AndAlso goCurrentPlayer.moTechs(X).ObjTypeID = miLastProdTypeID Then
									Me.lblProducing.Caption = goCurrentPlayer.moTechs(X).GetComponentName
									Exit For
								End If
							Next X
						End If
					Catch
						mlLastProdID = -1
						miLastProdTypeID = -1
					End Try
                End If
            End If
            lblStatusPerc.Caption = fPerc.ToString("#0.#0%")
        End If
    End Sub

    Protected Overrides Sub Finalize()
        msw_Delay.Stop()
        msw_Delay = Nothing
        MyBase.Finalize()
    End Sub

    Private Sub frmProdStatus_WindowMoved() Handles Me.WindowMoved
        muSettings.ProdStatusLocX = Me.Left
        muSettings.ProdStatusLocY = Me.Top
    End Sub
End Class
