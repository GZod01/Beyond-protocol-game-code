Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Interface created from Interface Builder
Public Class frmGuildMain
	Inherits UIWindow

	Private Const ml_CONTENTS_TOP As Int32 = 5

	Private lblTitle As UILabel
	Private lnDiv1 As UILine
	Private btnClose As UIButton

	'Private btnMainTab As UIButton
	'Private btnAssetsTab As UIButton
	'Private btnEventsTab As UIButton
	'Private btnInternalAffairsTab As UIButton
	'Private btnMembershipTab As UIButton
	'Private btnPoliticsTab As UIButton
	'Private btnStructureTab As UIButton

	Protected Shared mrc_INT_AFF_INPROGRESS As Rectangle = New Rectangle(430, 53, 100, 24) '57
	Protected Shared mrc_INT_AFF_HISTORY As Rectangle = New Rectangle(530, 53, 100, 24)	'57
	Protected Shared ml_IntAffSelectedTab As Int32 = 0
    Private mbLoading As Boolean = True
	Private mfraFrame As guildframe

	Private Structure uTabData
		Public lID As Int32
		Public sTabText As String
		Public rcRect As Rectangle
		Public sButtonName As String
	End Structure
	Private Enum eTabProgCtrl As Int32
		MainTab = 0
        EventTab = 1
        InternalAffairsTab = 2
        MembershipTab = 3
        StructureTab = 4
        FinancesTab = 5
        SharedAssets = 6
	End Enum

	Private muTabs() As uTabData
	Private Sub PopulateTabData()
        ReDim muTabs(5)
		With muTabs(0)
			.lID = eTabProgCtrl.MainTab : .sTabText = "Main" : .rcRect = New Rectangle(5, 26, 110, 24) : .sButtonName = "BTNMAINTAB"
		End With
        With muTabs(1) '2)
            '.lID = eTabProgCtrl.EventTab : .sTabText = "Events" : .rcRect = New Rectangle(225, 26, 110, 24) : .sButtonName = "BTNEVENTSTAB"
            .lID = eTabProgCtrl.EventTab : .sTabText = "Events" : .rcRect = New Rectangle(115, 26, 110, 24) : .sButtonName = "BTNEVENTSTAB"
        End With
        With muTabs(2) '3)
            '.lID = eTabProgCtrl.InternalAffairsTab : .sTabText = "Internal Affairs" : .rcRect = New Rectangle(335, 26, 110, 24) : .sButtonName = "BTNINTERNALAFFAIRSTAB"
            .lID = eTabProgCtrl.InternalAffairsTab : .sTabText = "Internal Affairs" : .rcRect = New Rectangle(225, 26, 110, 24) : .sButtonName = "BTNINTERNALAFFAIRSTAB"
        End With
        With muTabs(3) '4)
            '.lID = eTabProgCtrl.MembershipTab : .sTabText = "Membership" : .rcRect = New Rectangle(445, 26, 110, 24) : .sButtonName = "BTNMEMBERSHIPTAB"
            .lID = eTabProgCtrl.MembershipTab : .sTabText = "Membership" : .rcRect = New Rectangle(335, 26, 110, 24) : .sButtonName = "BTNMEMBERSHIPTAB"
        End With
        With muTabs(4) '5)
            '.lID = eTabProgCtrl.PoliticsTab : .sTabText = "Politics" : .rcRect = New Rectangle(555, 26, 110, 24) : .sButtonName = "BTNPOLITICSTAB"
            '.lID = eTabProgCtrl.StructureTab : .sTabText = "Structure" : .rcRect = New Rectangle(555, 26, 110, 24) : .sButtonName = "BTNSTRUCTURETAB"
            .lID = eTabProgCtrl.StructureTab : .sTabText = "Structure" : .rcRect = New Rectangle(445, 26, 110, 24) : .sButtonName = "BTNSTRUCTURETAB"
        End With
        With muTabs(5)
            .lID = eTabProgCtrl.FinancesTab : .sTabText = "Finances" : .rcRect = New Rectangle(555, 26, 110, 24) : .sButtonName = "BTNFINANCESTAB"
        End With
        'With muTabs(6)
        '    .lID = eTabProgCtrl.SharedAssets : .sTabText = "Shared Assets" : .rcRect = New Rectangle(665, 26, 110, 24) : .sButtonName = "BTNFSHAREDASSETSTAB"
        'End With
	End Sub

	Public Sub New(ByRef oUILib As UILib)
		MyBase.New(oUILib)

		'frmGuildMain initial props
		With Me
			.ControlName = "frmGuildMain"
            '.Left = MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth \ 2 - 400
            '.Top = MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight \ 2 - 300
            Dim lLeft As Int32 = -1
            Dim lTop As Int32 = -1

            If NewTutorialManager.TutorialOn = False Then
                lLeft = muSettings.GuildMainX
                lTop = muSettings.GuildMainY
            End If
            If lLeft < 0 Then lLeft = MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth \ 2 - 400
            If lTop < 0 Then lTop = MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight \ 2 - 300
            If lLeft + 800 > MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth Then lLeft = MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth - 800
            If lTop + 600 > MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight Then lTop = MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight - 600

            .Left = lLeft
            .Top = lTop

			.Width = 800
			.Height = 600
			.Enabled = True
			.Visible = True
			.BorderColor = muSettings.InterfaceBorderColor
			.FillColor = muSettings.InterfaceFillColor
			.FullScreen = True
			.mbAcceptReprocessEvents = True
		End With

		'lblTitle initial props
		lblTitle = New UILabel(oUILib)
		With lblTitle
			.ControlName = "lblTitle"
			.Left = 1
			.Top = 1
			.Width = 798
			.Height = 26
			.Enabled = True
			.Visible = True
			.Caption = "No Guild Name"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 12.0F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(5, DrawTextFormat)
		End With
		Me.AddChild(CType(lblTitle, UIControl))

		'btnClose initial props
		btnClose = New UIButton(oUILib)
		With btnClose
			.ControlName = "btnClose"
			.Left = 773
			.Top = 3
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

		'lnDiv1 initial props
		lnDiv1 = New UILine(oUILib)
		With lnDiv1
			.ControlName = "lnDiv1"
			.Left = 1
			.Top = 24
			.Width = 798
			.Height = 0
			.Enabled = True
			.Visible = True
			.BorderColor = muSettings.InterfaceBorderColor
		End With
		Me.AddChild(CType(lnDiv1, UIControl))

		''btnMainTab initial props
		'btnMainTab = New UIButton(oUILib)
		'With btnMainTab
		'	.ControlName = "btnMainTab"
		'	.Left = 5
		'	.Top = 26
		'	.Width = 110
		'	.Height = 24
		'	.Enabled = True
		'	.Visible = True
		'	.Caption = "Main"
		'	.ForeColor = muSettings.InterfaceBorderColor
		'	.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
		'	.DrawBackImage = True
		'	.FontFormat = CType(5, DrawTextFormat)
		'	.ControlImageRect = New Rectangle(0, 0, 120, 32)
		'End With
		'Me.AddChild(CType(btnMainTab, UIControl))

		''btnAssetsTab initial props
		'btnAssetsTab = New UIButton(oUILib)
		'With btnAssetsTab
		'	.ControlName = "btnAssetsTab"
		'	.Left = 115
		'	.Top = 26
		'	.Width = 110
		'	.Height = 24
		'	.Enabled = True
		'	.Visible = True
		'	.Caption = "Assets"
		'	.ForeColor = muSettings.InterfaceBorderColor
		'	.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
		'	.DrawBackImage = True
		'	.FontFormat = CType(5, DrawTextFormat)
		'	.ControlImageRect = New Rectangle(0, 0, 120, 32)
		'End With
		'Me.AddChild(CType(btnAssetsTab, UIControl))

		''btnEventsTab initial props
		'btnEventsTab = New UIButton(oUILib)
		'With btnEventsTab
		'	.ControlName = "btnEventsTab"
		'	.Left = 225
		'	.Top = 26
		'	.Width = 110
		'	.Height = 24
		'	.Enabled = True
		'	.Visible = True
		'	.Caption = "Events"
		'	.ForeColor = muSettings.InterfaceBorderColor
		'	.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
		'	.DrawBackImage = True
		'	.FontFormat = CType(5, DrawTextFormat)
		'	.ControlImageRect = New Rectangle(0, 0, 120, 32)
		'End With
		'Me.AddChild(CType(btnEventsTab, UIControl))

		''btnInternalAffairsTab initial props
		'btnInternalAffairsTab = New UIButton(oUILib)
		'With btnInternalAffairsTab
		'	.ControlName = "btnInternalAffairsTab"
		'	.Left = 335
		'	.Top = 26
		'	.Width = 110
		'	.Height = 24
		'	.Enabled = True
		'	.Visible = True
		'	.Caption = "Internal Affairs"
		'	.ForeColor = muSettings.InterfaceBorderColor
		'	.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
		'	.DrawBackImage = True
		'	.FontFormat = CType(5, DrawTextFormat)
		'	.ControlImageRect = New Rectangle(0, 0, 120, 32)
		'End With
		'Me.AddChild(CType(btnInternalAffairsTab, UIControl))

		''btnMembershipTab initial props
		'btnMembershipTab = New UIButton(oUILib)
		'With btnMembershipTab
		'	.ControlName = "btnMembershipTab"
		'	.Left = 445
		'	.Top = 26
		'	.Width = 110
		'	.Height = 24
		'	.Enabled = True
		'	.Visible = True
		'	.Caption = "Membership"
		'	.ForeColor = muSettings.InterfaceBorderColor
		'	.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
		'	.DrawBackImage = True
		'	.FontFormat = CType(5, DrawTextFormat)
		'	.ControlImageRect = New Rectangle(0, 0, 120, 32)
		'End With
		'Me.AddChild(CType(btnMembershipTab, UIControl))

		''btnPoliticsTab initial props
		'btnPoliticsTab = New UIButton(oUILib)
		'With btnPoliticsTab
		'	.ControlName = "btnPoliticsTab"
		'	.Left = 555
		'	.Top = 26
		'	.Width = 110
		'	.Height = 24
		'	.Enabled = True
		'	.Visible = True
		'	.Caption = "Politics"
		'	.ForeColor = muSettings.InterfaceBorderColor
		'	.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
		'	.DrawBackImage = True
		'	.FontFormat = CType(5, DrawTextFormat)
		'	.ControlImageRect = New Rectangle(0, 0, 120, 32)
		'End With
		'Me.AddChild(CType(btnPoliticsTab, UIControl))

		''btnStructureTab initial props
		'btnStructureTab = New UIButton(oUILib)
		'With btnStructureTab
		'	.ControlName = "btnStructureTab"
		'	.Left = 665
		'	.Top = 26
		'	.Width = 110
		'	.Height = 24
		'	.Enabled = True
		'	.Visible = True
		'	.Caption = "Structure"
		'	.ForeColor = muSettings.InterfaceBorderColor
		'	.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
		'	.DrawBackImage = True
		'	.FontFormat = CType(5, DrawTextFormat)
		'	.ControlImageRect = New Rectangle(0, 0, 120, 32)
		'End With
		'Me.AddChild(CType(btnStructureTab, UIControl))

		'AddHandler btnAssetsTab.Click, AddressOf btnClick_Click
		AddHandler btnClose.Click, AddressOf btnClick_Click
		'AddHandler btnEventsTab.Click, AddressOf btnClick_Click
		'AddHandler btnInternalAffairsTab.Click, AddressOf btnClick_Click
		'AddHandler btnMainTab.Click, AddressOf btnClick_Click
		'AddHandler btnMembershipTab.Click, AddressOf btnClick_Click
		'AddHandler btnPoliticsTab.Click, AddressOf btnClick_Click
		'AddHandler btnStructureTab.Click, AddressOf btnClick_Click

		PopulateTabData()

		mfraFrame = New fra_Main(oUILib)
		With mfraFrame
			.Left = 0
			.Top = 60
			.Height = Me.Height - 60
			.Width = Me.Width
			.DrawBorder = False
			.FillWindow = False
			.mbAcceptReprocessEvents = True
		End With
		Me.AddChild(CType(mfraFrame, UIControl))

		Dim yMsg(1) As Byte
		System.BitConverter.GetBytes(GlobalMessageCode.eRequestGuildEvents).CopyTo(yMsg, 0)
		MyBase.moUILib.SendMsgToPrimary(yMsg)
        mbLoading = False

		MyBase.moUILib.RemoveWindow(Me.ControlName)
		MyBase.moUILib.AddWindow(Me)
	End Sub

	Private mlCurrTabVal As Int32 = 0
	Protected Property mlCurrentSelectedTab() As Int32
		Get
			Return mlCurrTabVal
		End Get
		Set(ByVal value As Int32)
			If value <> mlCurrTabVal Then
				mlCurrTabVal = value

				If mfraFrame Is Nothing = False Then
					For Y As Int32 = 0 To Me.ChildrenUB
						If Me.moChildren(Y) Is Nothing = False AndAlso Me.moChildren(Y).ControlName Is Nothing = False Then
							If Me.moChildren(Y).ControlName = mfraFrame.ControlName Then
								Me.RemoveChild(Y)
								Exit For
							End If
						End If
					Next Y
					mfraFrame = Nothing
				End If

				Select Case mlCurrTabVal
                    'Case eTabProgCtrl.AssetTab
                    '	mfraFrame = New fra_Assets(MyBase.moUILib)
					Case eTabProgCtrl.EventTab
						mfraFrame = New fra_Events(MyBase.moUILib)
					Case eTabProgCtrl.InternalAffairsTab
						mfraFrame = New fra_InternalAffairs(MyBase.moUILib)
					Case eTabProgCtrl.MainTab
						mfraFrame = New fra_Main(MyBase.moUILib)
					Case eTabProgCtrl.MembershipTab
						mfraFrame = New fra_Membership(MyBase.moUILib)
                        'Case eTabProgCtrl.PoliticsTab
                        '	mfraFrame = New fra_Politics(MyBase.moUILib)
					Case eTabProgCtrl.StructureTab
                        mfraFrame = New fra_Structure(MyBase.moUILib)
                    Case eTabProgCtrl.FinancesTab
                        mfraFrame = New fra_Finances(MyBase.moUILib)
                    Case eTabProgCtrl.SharedAssets
                        mfraFrame = New fra_SharedAssets(MyBase.moUILib)
                End Select
				If mfraFrame Is Nothing = False Then
					With mfraFrame
						.Left = 0
						.Top = 60
						.Height = Me.Height - 60
						.Width = Me.Width
						.DrawBorder = False
						.FillWindow = False
						.mbAcceptReprocessEvents = True
					End With
					Me.AddChild(CType(mfraFrame, UIControl))
				End If

			End If
		End Set
	End Property

	Private Sub frmGuildMain_OnMouseDown(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles Me.OnMouseDown
		Dim ptLoc As Point = Me.GetAbsolutePosition
		lMouseX -= ptLoc.X
		lMouseY -= ptLoc.Y

		For X As Int32 = 0 To muTabs.GetUpperBound(0)
			If muTabs(X).rcRect.Contains(lMouseX, lMouseY) = True Then

				If muTabs(X).lID = eTabProgCtrl.EventTab Then
					If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.oGuild Is Nothing = False AndAlso goCurrentPlayer.oGuild.HasPermission(RankPermissions.ViewEvents) = False Then
						MyBase.moUILib.AddNotification("You do not have rights to view the events tab.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
						If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eUserInterface, Nothing, Nothing)
						Return
					End If
				End If

				mlCurrentSelectedTab = CByte(muTabs(X).lID)
				If goSound Is Nothing = False Then goSound.StartSound("UserInterface\ButtonClick.wav", False, SoundMgr.SoundUsage.eUserInterface, Nothing, Nothing)
				Return
			End If
		Next X

		If mlCurrentSelectedTab = eTabProgCtrl.InternalAffairsTab Then
			If mrc_INT_AFF_HISTORY.Contains(lMouseX, lMouseY) = True Then
				ml_IntAffSelectedTab = 1
				Me.IsDirty = True
				If goSound Is Nothing = False Then goSound.StartSound("UserInterface\ButtonClick.wav", False, SoundMgr.SoundUsage.eUserInterface, Nothing, Nothing)
			ElseIf mrc_INT_AFF_INPROGRESS.Contains(lMouseX, lMouseY) = True Then
				ml_IntAffSelectedTab = 0
				Me.IsDirty = True
				If goSound Is Nothing = False Then goSound.StartSound("UserInterface\ButtonClick.wav", False, SoundMgr.SoundUsage.eUserInterface, Nothing, Nothing)
			End If
		End If
	End Sub

	Private Sub frmGuildMain_OnNewFrame() Handles Me.OnNewFrame
		If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.oGuild Is Nothing = False Then
			If lblTitle.Caption <> goCurrentPlayer.oGuild.sName Then lblTitle.Caption = goCurrentPlayer.oGuild.sName
		Else
			MyBase.moUILib.RemoveWindow(Me.ControlName)
			Return
		End If
		If mfraFrame Is Nothing = False Then mfraFrame.NewFrame()
	End Sub

	Private Sub frmGuildMain_OnRenderEnd() Handles Me.OnRenderEnd
		Dim rcRect As Rectangle

		rcRect = muTabs(mlCurrentSelectedTab).rcRect

		'render our selected item
		MyBase.moUILib.DoAlphaBlendColorFill(rcRect, System.Drawing.Color.FromArgb(255, 128, 128, 160), rcRect.Location)

		If mlCurrentSelectedTab = eTabProgCtrl.InternalAffairsTab Then
			rcRect = mrc_INT_AFF_INPROGRESS
			If ml_IntAffSelectedTab = 0 Then
				MyBase.moUILib.DoAlphaBlendColorFill(rcRect, System.Drawing.Color.FromArgb(255, 128, 128, 160), rcRect.Location)
			Else : MyBase.moUILib.DoAlphaBlendColorFill(rcRect, muSettings.InterfaceFillColor, rcRect.Location)
			End If
			rcRect = mrc_INT_AFF_HISTORY
			If ml_IntAffSelectedTab = 1 Then
				MyBase.moUILib.DoAlphaBlendColorFill(rcRect, System.Drawing.Color.FromArgb(255, 128, 128, 160), rcRect.Location)
			Else : MyBase.moUILib.DoAlphaBlendColorFill(rcRect, muSettings.InterfaceFillColor, rcRect.Location)
			End If
		End If

		Dim clrEvent As System.Drawing.Color = muSettings.InterfaceBorderColor
		If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.oGuild Is Nothing = False AndAlso goCurrentPlayer.oGuild.HasPermission(RankPermissions.ViewEvents) = False Then
			clrEvent = System.Drawing.Color.FromArgb(clrEvent.A \ 2, clrEvent.R, clrEvent.G, clrEvent.B)
		End If

		For X As Int32 = 0 To muTabs.GetUpperBound(0)
			rcRect = muTabs(X).rcRect

			If muTabs(X).lID = eTabProgCtrl.EventTab Then
				MyBase.RenderRoundedBorder(rcRect, 1, clrEvent)
			Else
				MyBase.RenderRoundedBorder(rcRect, 1, muSettings.InterfaceBorderColor)
			End If
		Next X

		If mlCurrentSelectedTab = eTabProgCtrl.InternalAffairsTab Then
			rcRect = mrc_INT_AFF_INPROGRESS
			MyBase.RenderRoundedBorder(rcRect, 1, muSettings.InterfaceBorderColor)
			rcRect = mrc_INT_AFF_HISTORY
			MyBase.RenderRoundedBorder(rcRect, 1, muSettings.InterfaceBorderColor)
		End If

        If GFXEngine.gbPaused = True OrElse GFXEngine.gbDeviceLost = True Then Return

		'Now, render our text...
		Using oFont As Font = New Font(MyBase.moUILib.oDevice, New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			Using oTextSpr As Sprite = New Sprite(MyBase.moUILib.oDevice)
				Dim bBegun As Boolean = False
				Try
					oTextSpr.Begin(SpriteFlags.AlphaBlend)
					bBegun = True

					For X As Int32 = 0 To muTabs.GetUpperBound(0)
						rcRect = muTabs(X).rcRect
						If muTabs(X).lID = eTabProgCtrl.EventTab Then
							oFont.DrawText(oTextSpr, muTabs(X).sTabText, rcRect, DrawTextFormat.Center Or DrawTextFormat.VerticalCenter, clrEvent)
						Else
							oFont.DrawText(oTextSpr, muTabs(X).sTabText, rcRect, DrawTextFormat.Center Or DrawTextFormat.VerticalCenter, muSettings.InterfaceBorderColor)
						End If
					Next X

					If mlCurrentSelectedTab = eTabProgCtrl.InternalAffairsTab Then
						rcRect = mrc_INT_AFF_INPROGRESS
						oFont.DrawText(oTextSpr, "In Progress", rcRect, DrawTextFormat.Center Or DrawTextFormat.VerticalCenter, muSettings.InterfaceBorderColor)
						rcRect = mrc_INT_AFF_HISTORY
						oFont.DrawText(oTextSpr, "History", rcRect, DrawTextFormat.Center Or DrawTextFormat.VerticalCenter, muSettings.InterfaceBorderColor)
					End If
				Catch
				End Try

				If bBegun = True Then oTextSpr.End()
				oTextSpr.Dispose()
			End Using
			oFont.Dispose()
		End Using

		If mfraFrame Is Nothing = False Then mfraFrame.RenderEnd()
	End Sub

	Private Sub btnClick_Click(ByVal sName As String)
		Me.RemoveAllChildren()

		Me.AddChild(CType(lblTitle, UIControl))
		Me.AddChild(CType(lnDiv1, UIControl))
		Me.AddChild(CType(btnClose, UIControl))

		Select Case sName.ToUpper
			Case "BTNCLOSE"
				MyBase.moUILib.RemoveWindow(Me.ControlName)
		End Select
	End Sub

	Protected Function GetFullReferencedControl(ByVal sFullName As String) As UIControl
		Dim sTree() As String = Split(sFullName, ".")
		Dim oParent As UIControl = Me
		If sTree Is Nothing = False Then
			For lTreeItem As Int32 = 0 To sTree.GetUpperBound(0)
				Dim sItem As String = sTree(lTreeItem).ToUpper.Trim

				If sItem = "" Then Return oParent

				'Now, find our item
				Dim oResult As UIControl = Nothing
				For X As Int32 = 0 To oParent.ChildrenUB
					If oParent.moChildren(X) Is Nothing = False AndAlso oParent.moChildren(X).ControlName Is Nothing = False Then
						If oParent.moChildren(X).ControlName.ToUpper = sItem Then
							oResult = oParent.moChildren(X)
							Exit For
						End If
					End If
				Next X

				If oResult Is Nothing Then Return Nothing
				oParent = oResult
			Next lTreeItem
			If Object.Equals(oParent, Me) = False Then Return oParent
		End If
		Return Nothing
	End Function

    Public MustInherit Class guildframe
        Inherits UIWindow

        Public MustOverride Sub NewFrame()
        Public MustOverride Sub RenderEnd()

        Public Sub New(ByRef oUILib As UILib)
            MyBase.New(oUILib)
        End Sub
    End Class

	Public Sub HandleGetTradepostList(ByRef yData() As Byte)
		If mfraFrame Is Nothing = False AndAlso mfraFrame.ControlName Is Nothing = False AndAlso mfraFrame.ControlName.ToUpper = "FRA_ASSETS" Then
            'CType(mfraFrame, fra_Assets).HandleGetTradePostList(yData)
		End If
	End Sub
 
    Private Sub frmGuildMain_WindowMoved() Handles Me.WindowMoved
        If mbLoading = False Then
            muSettings.GuildMainX = Me.Left
            muSettings.GuildMainY = Me.Top
        End If
    End Sub

    Public Sub HandleGetGuildShareAssets(ByRef yData() As Byte)
        If mfraFrame Is Nothing = False AndAlso mfraFrame.ControlName Is Nothing = False AndAlso mfraFrame.ControlName.ToUpper = "FRA_SHAREDASSETS" Then
            CType(mfraFrame, fra_SharedAssets).HandleGetGuildShareAssets(yData)
        End If
    End Sub

End Class