Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Interface created from Interface Builder
Public Class frmCancelBuild
	Inherits UIWindow

	Public Enum eyOrderType As Byte
		SetProduction = 0
		SetPrimaryTarget = 1
		MoveRequest = 2
	End Enum

	Private lblTitle As UILabel
	Private txtDisplay As UITextBox
	Private WithEvents btnConfirm As UIButton
	Private WithEvents btnCancel As UIButton
	Private chkDoNotShow As UICheckBox

	Private myOrderType As eyOrderType
	Private mlEntityID As Int32
	Private miEntityTypeID As Int16
	Private mlTargetID As Int32
	Private miTargetTypeID As Int16
	Private mlLocX As Int32
	Private mlLocZ As Int32
	Private miLocA As Int16

	Public Sub New(ByRef oUILib As UILib)
		MyBase.New(oUILib)

		'frmCancelBuild initial props
        With Me
            .lWindowMetricID = BPMetrics.eWindow.eCancelBuild
            .ControlName = "frmCancelBuild"
            .Left = MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth \ 2 - 128
            .Top = MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight \ 2 - 64
            .Width = 256
            .Height = 128
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = False
            .Moveable = True
        End With

		'lblTitle initial props
		lblTitle = New UILabel(oUILib)
		With lblTitle
			.ControlName = "lblTitle"
			.Left = 0
			.Top = 0
			.Width = 256
			.Height = 24
			.Enabled = True
			.Visible = True
			.Caption = "Confirm Cancel Build Order"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(5, DrawTextFormat)
		End With
		Me.AddChild(CType(lblTitle, UIControl))

		'txtDisplay initial props
		txtDisplay = New UITextBox(oUILib)
		With txtDisplay
			.ControlName = "txtDisplay"
			.Left = 5
			.Top = 20
			.Width = 246
			.Height = 53
			.Enabled = True
			.Visible = True
			.Caption = "The selected engineer has already been ordered to build something. Issuing a new order will cancel the previous build order."
			.ForeColor = muSettings.InterfaceTextBoxForeColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(0, DrawTextFormat)
			.BackColorEnabled = muSettings.InterfaceTextBoxFillColor
			.BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
			.MaxLength = 0
			.MultiLine = True
			.Locked = True
			.BorderColor = muSettings.InterfaceBorderColor
		End With
		Me.AddChild(CType(txtDisplay, UIControl))

		'btnConfirm initial props
		btnConfirm = New UIButton(oUILib)
		With btnConfirm
			.ControlName = "btnConfirm"
			.Left = 5
			.Top = 100
			.Width = 100
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

		'btnCancel initial props
		btnCancel = New UIButton(oUILib)
		With btnCancel
			.ControlName = "btnCancel"
			.Left = 153
			.Top = 100
			.Width = 100
			.Height = 24
			.Enabled = True
			.Visible = True
			.Caption = "Cancel"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = True
			.FontFormat = CType(5, DrawTextFormat)
			.ControlImageRect = New Rectangle(0, 0, 120, 32)
		End With
		Me.AddChild(CType(btnCancel, UIControl))

		'chkDoNotShow initial props
		chkDoNotShow = New UICheckBox(oUILib)
		With chkDoNotShow
			.ControlName = "chkDoNotShow"
			.Left = 50
			.Top = 70
			.Width = 155
			.Height = 32
			.Enabled = True
			.Visible = True
			.Caption = "Never Show This Alert "
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(6, DrawTextFormat)
			.Value = False
		End With
		Me.AddChild(CType(chkDoNotShow, UIControl))

		MyBase.moUILib.RemoveWindow(Me.ControlName)
		MyBase.moUILib.AddWindow(Me)
	End Sub

	Public Sub SetOrder(ByVal yOrderType As eyOrderType, ByVal lEntityID As Int32, ByVal iEntityTypeID As Int16, ByVal lTargetID As Int32, ByVal iTargetTypeID As Int16, ByVal lX As Int32, ByVal lZ As Int32, ByVal iA As Int16)
		'here, we set up the order so that if Confirm is clicked, the order is sent from here
		myOrderType = yOrderType
		mlEntityID = lEntityID
		miEntityTypeID = iEntityTypeID
		mlTargetID = lTargetID
		miTargetTypeID = iTargetTypeID
		mlLocX = lX
		mlLocZ = lZ
		miLocA = iA
	End Sub

	Private Sub btnCancel_Click(ByVal sName As String) Handles btnCancel.Click
		muSettings.bDoNotShowEngineerCancelAlert = chkDoNotShow.Value
		MyBase.moUILib.RemoveWindow(Me.ControlName)
	End Sub

	Private Sub btnConfirm_Click(ByVal sName As String) Handles btnConfirm.Click
		muSettings.bDoNotShowEngineerCancelAlert = chkDoNotShow.Value
		Select Case myOrderType
			Case eyOrderType.SetProduction
				Dim yData(23) As Byte
				System.BitConverter.GetBytes(GlobalMessageCode.eSetEntityProduction).CopyTo(yData, 0)
				System.BitConverter.GetBytes(mlEntityID).CopyTo(yData, 2)
				System.BitConverter.GetBytes(miEntityTypeID).CopyTo(yData, 6)
				System.BitConverter.GetBytes(mlTargetID).CopyTo(yData, 8)
				System.BitConverter.GetBytes(miTargetTypeID).CopyTo(yData, 12)
				System.BitConverter.GetBytes(mlLocX).CopyTo(yData, 14)
				System.BitConverter.GetBytes(mlLocZ).CopyTo(yData, 18)
				System.BitConverter.GetBytes(miLocA).CopyTo(yData, 22)

				If miEntityTypeID = ObjectType.eUnit Then
					If goSound Is Nothing = False Then goSound.StartUnitSpeechSound(True, goSound.GetUnitSpeechTypeCnt(SoundMgr.SpeechType.eBuildingConstruction, 1))
				End If
				MyBase.moUILib.SendMsgToRegion(yData)
			Case eyOrderType.MoveRequest
				Dim bShiftDown As Boolean = mlEntityID = Int32.MinValue
				MyBase.moUILib.GetMsgSys.SendMoveRequestMsg(mlLocX, mlLocZ, miLocA, bShiftDown, True)
			Case eyOrderType.SetPrimaryTarget
				Dim oTarget As Base_GUID = Nothing
				Dim oEnvir As BaseEnvironment = goCurrentEnvir
				If oEnvir Is Nothing = False Then
					For X As Int32 = 0 To oEnvir.lEntityUB
						If oEnvir.lEntityIdx(X) = mlTargetID AndAlso oEnvir.oEntity(X).ObjTypeID = miTargetTypeID Then
							oTarget = oEnvir.oEntity(X)
							Exit For
						End If
					Next X
				End If
				If oTarget Is Nothing = False Then MyBase.moUILib.GetMsgSys.SendSetPrimaryTarget(oTarget, True)
		End Select
		MyBase.moUILib.RemoveWindow(Me.ControlName)
	End Sub
End Class