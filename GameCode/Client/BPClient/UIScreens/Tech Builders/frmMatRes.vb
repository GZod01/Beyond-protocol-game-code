Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Material Research
Public Class frmMatRes
    Inherits UIWindow

    'Private WithEvents btnDiscover As UIButton
    'Private WithEvents btnStudy As UIButton
    Private WithEvents lstMaterials As UIListBox
    Private WithEvents btnResearch As UIButton
    Private WithEvents btnCancel As UIButton
    Private WithEvents btnHelp As UIButton
    Private WithEvents btnClose As UIButton
    Private WithEvents btnExport As UIButton
    Private chkUntilFullyStudied As UICheckBox
    Private txtProdCosts As UITextBox
    Private fraMaterialDetails As UIWindow
    Private lnDivider As UILine

	Private WithEvents chkFilterArchived As UICheckBox
	Private WithEvents chkSetFilter As UICheckBox

	Private mbIgnoreSetFilter As Boolean = False

    'Private myClickedButton As Byte = 0		 '0 = none, 1 = discover, 2 = study

	Private mlLastUpdateMsg As Int32
	Private mlCurrentMineralID As Int32

	Private mlLastListCheck As Int32

    'Private lblKnown() As UILabel 
    Private lblProps() As UILabel
    Private lblValue() As UILabel

    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)

        'frmMatRes initial props
        With Me
            .lWindowMetricID = BPMetrics.eWindow.eMaterialResearch
            .ControlName = "frmMatRes"
            .Width = 512 '650
            .Height = 512 '444
            Dim ofrmResearchMain As frmResearchMain = CType(MyBase.moUILib.GetWindow("frmResearchMain"), frmResearchMain)
            If ofrmResearchMain Is Nothing = False Then
                If ofrmResearchMain.Left - .Width < 0 Then
                    .Left = ofrmResearchMain.Left + ofrmResearchMain.Width
                Else
                    .Left = ofrmResearchMain.Left - .Width
                End If
                .Top = ofrmResearchMain.Top
            Else
                .Left = CInt(oUILib.oDevice.PresentationParameters.BackBufferWidth / 2) - 325
                .Top = CInt(oUILib.oDevice.PresentationParameters.BackBufferHeight / 2) - 222
            End If


            .Enabled = True
            .Visible = True
            '.BorderColor = System.Drawing.Color.FromArgb(255, 0, 255, 255)
            .BorderColor = muSettings.InterfaceBorderColor
            '.FillColor = System.Drawing.Color.FromArgb(255, 64, 128, 192)
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = False
        End With
        'btnClose initial props
        btnClose = New UIButton(oUILib)
        With btnClose
            .ControlName = "btnClose"
            .Left = Me.Width - 24 - Me.BorderLineWidth
            .Top = Me.BorderLineWidth
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

        'btnHelp initial props
        btnHelp = New UIButton(oUILib)
        With btnHelp
            .ControlName = "btnHelp"
            .Left = btnClose.Left - 25
            .Top = btnClose.Top
            .Width = 24
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "?"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
            .ToolTipText = "Click to begin the tutorial for this window."
        End With
        Me.AddChild(CType(btnHelp, UIControl))

        'btnExport
        btnExport = New UIButton(oUILib)
        With btnExport
            .ControlName = "btnExport"
            .Left = btnHelp.Left - 25
            .Top = btnClose.Top
            .Width = 24
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "E"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
            .ToolTipText = "Click to export data." & vbCrLf & "Exported data will go into your GameDir\ExportedData"
        End With
        Me.AddChild(CType(btnExport, UIControl))

        'lnDivider initial props
        lnDivider = New UILine(oUILib)
        With lnDivider
            .ControlName = "lnDivider"
            .Left = Me.BorderLineWidth \ 2
            .Top = btnClose.Top + btnClose.Height + 2
            .Width = Me.Width - Me.BorderLineWidth
            .Height = 0
            .Enabled = True
            .Visible = True
            '.BorderColor = System.Drawing.Color.FromArgb(255, 0, 255, 255)
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(lnDivider, UIControl))

        ''btnDiscover initial props
        'btnDiscover = New UIButton(oUILib)
        'With btnDiscover
        '    .ControlName = "btnDiscover"
        '    .Left = 7
        '    .Top = 10
        '    .Width = 122
        '    .Height = 23
        '    .Enabled = True
        '    .Visible = True
        '    .Caption = "Discover Material"
        '    .ForeColor = muSettings.InterfaceBorderColor
        '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10, FontStyle.Regular, GraphicsUnit.Point, 0))
        '    .DrawBackImage = True
        '    .FontFormat = DrawTextFormat.Center Or DrawTextFormat.VerticalCenter
        'End With
        'Me.AddChild(CType(btnDiscover, UIControl))

        ''btnStudy initial props
        'btnStudy = New UIButton(oUILib)
        'With btnStudy
        '    .ControlName = "btnStudy"
        '    .Left = 135
        '    .Top = 10
        '    .Width = 122
        '    .Height = 23
        '    .Enabled = True
        '    .Visible = True
        '    .Caption = "Study Material"
        '    .ForeColor = muSettings.InterfaceBorderColor
        '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10, FontStyle.Regular, GraphicsUnit.Point, 0))
        '    .DrawBackImage = True
        '    .FontFormat = DrawTextFormat.Center Or DrawTextFormat.VerticalCenter
        'End With
        'Me.AddChild(CType(btnStudy, UIControl))

        'lstMaterials initial props
        lstMaterials = New UIListBox(oUILib)
        With lstMaterials
            .ControlName = "lstMaterials"
            .Left = 9
            .Top = 53
            .Width = 155 '184
            .Height = 240
            .Enabled = True
            .Visible = True
            '.BorderColor = System.Drawing.Color.FromArgb(255, 0, 255, 255)
            '.FillColor = System.Drawing.Color.FromArgb(255, 255, 255, 255)
            '.ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8, FontStyle.Regular, GraphicsUnit.Point, 0))
            '.HighlightColor = System.Drawing.Color.FromArgb(255, 0, 255, 255)
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .ForeColor = muSettings.InterfaceBorderColor
            .HighlightColor = System.Drawing.Color.FromArgb(255, muSettings.InterfaceFillColor.R, muSettings.InterfaceFillColor.G, muSettings.InterfaceFillColor.B)
        End With
        Me.AddChild(CType(lstMaterials, UIControl))

        ''btnHelp initial props
        'btnHelp = New UIButton(oUILib)
        'With btnHelp
        '    .ControlName = "btnHelp"
        '    .Left = Me.Width - 24 - Me.BorderLineWidth
        '    .Top = Me.BorderLineWidth
        '    .Width = 24
        '    .Height = 24
        '    .Enabled = True
        '    .Visible = True
        '    .Caption = "?"
        '    .ForeColor = muSettings.InterfaceBorderColor
        '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 14.0F, FontStyle.Bold, GraphicsUnit.Point, 0))
        '    .DrawBackImage = True
        '    .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Center
        '    .ToolTipText = "Click to begin the tutorial for this window."
        'End With
        'Me.AddChild(CType(btnHelp, UIControl))

        'btnCancel initial props
        btnCancel = New UIButton(oUILib)
        With btnCancel
            .ControlName = "btnCancel"
            .Left = Me.Width - 120 '530 
            .Top = Me.Height - 33
            .Width = 110
            .Height = 23
            .Enabled = True
            .Visible = True
            .Caption = "Cancel"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Center
        End With
        Me.AddChild(CType(btnCancel, UIControl))

        'btnResearch initial props
        btnResearch = New UIButton(oUILib)
        With btnResearch
            .ControlName = "btnResearch"
            .Left = Me.Width - 240 '415 
            .Top = btnCancel.Top
            .Width = 110
            .Height = 23
            .Enabled = True
            .Visible = True
            .Caption = "Research"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = DrawTextFormat.Center Or DrawTextFormat.VerticalCenter
        End With
        Me.AddChild(CType(btnResearch, UIControl))

        'txtMaterialDetails initial props
        fraMaterialDetails = New UIWindow(oUILib)
        With fraMaterialDetails
            .ControlName = "fraMaterialDetails"
            .Left = 180 '202
            .Top = 53
            .Width = 325  '440
            .Height = 415 '347
            .Enabled = False
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
        End With
        Me.AddChild(CType(fraMaterialDetails, UIControl))

        'txtProdCosts initial props
        txtProdCosts = New UITextBox(oUILib)
        With txtProdCosts
            .ControlName = "txtProdCosts"
            .Left = 9
            .Top = fraMaterialDetails.Top + fraMaterialDetails.Height - 100 ' 300
            .Width = lstMaterials.Width
            .Height = 100
            .Enabled = True
            .Visible = True
            .Caption = "RESEARCH COSTS:" & vbCrLf & vbCrLf & "*  10000 credits" & vbCrLf & "*  200 units of the material"
            .ForeColor = muSettings.InterfaceTextBoxForeColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(0, DrawTextFormat)
            .BackColorEnabled = muSettings.InterfaceTextBoxFillColor
            .BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
            .MaxLength = 0
            .BorderColor = muSettings.InterfaceBorderColor
            .MultiLine = True
            .Locked = True
        End With
        Me.AddChild(CType(txtProdCosts, UIControl))

        'chkUntilFullyStudied initial props
        chkUntilFullyStudied = New UICheckBox(oUILib)
        With chkUntilFullyStudied
            .ControlName = "chkUntilFullyStudied"
            .Left = btnResearch.Left - 150
            .Top = (btnResearch.Top + (btnResearch.Height \ 2)) - 9
            .Width = 130
            .Height = 18
            .Enabled = True
            .Visible = False
            .Caption = "Until Fully Studied"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = True
            .DisplayType = UICheckBox.eCheckTypes.eSmallCheck
            .ToolTipText = "Check to cause the research facility to research this material until it is fully understood."
        End With
		Me.AddChild(CType(chkUntilFullyStudied, UIControl))

		'chkFilterArchived initial props
		chkFilterArchived = New UICheckBox(oUILib)
		With chkFilterArchived
			.ControlName = "chkFilterArchived"
			.Left = lstMaterials.Left
			.Top = lstMaterials.Top + lstMaterials.Height + 10
			.Width = 100
			.Height = 18
			.Enabled = True
            .Visible = True
			.Caption = "Filter Archived"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(6, DrawTextFormat)
			.Value = True
			.DisplayType = UICheckBox.eCheckTypes.eSmallCheck
		End With
		Me.AddChild(CType(chkFilterArchived, UIControl))

		'chkSetFilter initial props
		chkSetFilter = New UICheckBox(oUILib)
		With chkSetFilter
			.ControlName = "chkSetFilter"
			.Left = chkFilterArchived.Left
			.Top = chkFilterArchived.Top + chkFilterArchived.Height + 5
			.Width = 100
			.Height = 18
            .Enabled = Not gbAliased
            .Visible = True
			.Caption = "Filter Selected"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(6, DrawTextFormat)
			.Value = True
			.DisplayType = UICheckBox.eCheckTypes.eSmallCheck
		End With
		Me.AddChild(CType(chkSetFilter, UIControl))

        ReDim lblProps(glMineralPropertyUB)
        ReDim lblValue(glMineralPropertyUB)
        'ReDim lblKnown(glMineralPropertyUB)

        For lIdx As Int32 = 0 To glMineralPropertyUB
            lblProps(lIdx) = New UILabel(MyBase.moUILib)
            With lblProps(lIdx)
                .ControlName = "lblProps(" & lIdx.ToString & ")"
                .Left = 185 ' 207
                .Top = 60 + (lIdx * 18)
                .Width = 175
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = goMineralProperty(lIdx).MineralPropertyName
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = DrawTextFormat.VerticalCenter
            End With
            Me.AddChild(CType(lblProps(lIdx), UIControl))

            lblValue(lIdx) = New UILabel(MyBase.moUILib)
            With lblValue(lIdx)
                .ControlName = "lblValue(" & lIdx.ToString & ")"
                .Left = 355 '385
                .Top = lblProps(lIdx).Top + 4
                .Width = 63
                .Height = 9
                .Enabled = True
                .Visible = False
                .Caption = ""
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = True
                .FontFormat = DrawTextFormat.VerticalCenter
            End With
            Me.AddChild(CType(lblValue(lIdx), UIControl))

            'lblKnown(lIdx) = New UILabel(MyBase.moUILib)
            'With lblKnown(lIdx)
            '    .ControlName = "lblKnown(" & lIdx.ToString & ")"
            '    .Left = 425 ' 475
            '    .Top = lblProps(lIdx).Top + 4
            '    .Width = 63
            '    .Height = 9
            '    .Enabled = True
            '    .Visible = False
            '    .Caption = ""
            '    .ForeColor = muSettings.InterfaceBorderColor
            '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10, FontStyle.Bold, GraphicsUnit.Point, 0))
            '    .DrawBackImage = True
            '    .FontFormat = DrawTextFormat.VerticalCenter
            'End With
            'Me.AddChild(CType(lblKnown(lIdx), UIControl)) 
        Next

        MyBase.moUILib.RemoveWindow(Me.ControlName)
		MyBase.moUILib.AddWindow(Me)

        If goTutorial Is Nothing = False AndAlso goTutorial.TutorialOn = True Then goTutorial.HandleTutorialTrigger(TutorialManager.TutorialTriggerType.MaterialResearchWindow)
    End Sub

    'Private Sub btnDiscover_Click(ByVal sName As String) Handles btnDiscover.Click
    '	If goTutorial Is Nothing = False AndAlso goTutorial.TutorialOn = True Then goTutorial.HandleTutorialTrigger(TutorialManager.TutorialTriggerType.MatResDiscoverButton)

    '	If NewTutorialManager.TutorialOn = True Then
    '		NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eDiscoverMineralButton, -1, -1, -1, "")
    '	End If

    '	lstMaterials.Clear()
    '	fraMaterialDetails.Visible = False
    '	chkUntilFullyStudied.Visible = False
    '	chkFilterArchived.Visible = False
    '	chkSetFilter.Visible = False

    '	For X As Int32 = 0 To lblProps.GetUpperBound(0)
    '		lblProps(X).Visible = False
    '		lblValue(X).Visible = False
    '           'lblKnown(X).Visible = False
    '	Next X

    '	myClickedButton = 1

    '	btnDiscover.ControlImageRect = btnDiscover.ControlImageRect_Pressed
    'End Sub

    Private Sub FillMinList()
        Dim bFilter As Boolean = chkFilterArchived.Value

        lstMaterials.Clear()

        Dim lSorted() As Int32 = GetSortedMineralIdxArray(True)
        If lSorted Is Nothing = False Then
            For X As Int32 = 0 To lSorted.GetUpperBound(0)
                If bFilter = False OrElse goMinerals(lSorted(X)).bArchived = False Then
                    lstMaterials.AddItem(goMinerals(lSorted(X)).MineralName)
                    lstMaterials.ItemData(lstMaterials.NewIndex) = goMinerals(lSorted(X)).ObjectID
                End If
            Next X
        End If
    End Sub

	Private Sub btnResearch_Click(ByVal sName As String) Handles btnResearch.Click
		If HasAliasedRights(AliasingRights.eAddResearch) = False Then
			MyBase.moUILib.AddNotification("You lack the rights to begin research.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
			Return
		End If
		'Ok research it
        'If myClickedButton = 0 Then
        ' MyBase.moUILib.AddNotification("You must select Discover or Study!", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
        If lstMaterials.ListIndex = -1 Then
            MyBase.moUILib.AddNotification("Select a material to research first!", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
        Else

            'If myClickedButton = 1 Then
            ' If goTutorial Is Nothing = False AndAlso goTutorial.TutorialOn = True Then goTutorial.HandleTutorialTrigger(TutorialManager.TutorialTriggerType.MatResDiscoverResearchStart)
            'End If

            Dim yData(13) As Byte

            For X As Int32 = 0 To goCurrentEnvir.lEntityUB
                If goCurrentEnvir.lEntityIdx(X) <> -1 AndAlso goCurrentEnvir.oEntity(X).bSelected = True AndAlso goCurrentEnvir.oEntity(X).OwnerID = glPlayerID Then
                    Dim oResearchGuid As Base_GUID = goCurrentEnvir.oEntity(X)
                    'ok, if the entity production is station...
                    If goCurrentEnvir.oEntity(X).yProductionType = ProductionType.eSpaceStationSpecial Then
                        'Ok, need to find research lab... try our lstFac...
                        Dim frmSelFac As frmSelectFac = CType(MyBase.moUILib.GetWindow("frmSelectFac"), frmSelectFac)
                        If frmSelFac Is Nothing = False Then
                            Dim oChild As StationChild = frmSelFac.GetCurrentChild()
                            If oChild Is Nothing = False Then
                                oResearchGuid = New Base_GUID
                                oResearchGuid.ObjectID = oChild.lChildID
                                oResearchGuid.ObjTypeID = oChild.iChildTypeID
                            End If
                        End If
                    End If

                    Dim lMineralID As Int32 = lstMaterials.ItemData(lstMaterials.ListIndex)
                    'If myClickedButton <> 1 Then
                    '    Dim bGood As Boolean = False
                    '    For Z As Int32 = 0 To glMineralUB
                    '        If glMineralIdx(Z) = lMineralID Then
                    '            For lProp As Int32 = 0 To glMineralPropertyUB
                    '                If glMineralPropertyIdx(lProp) <> -1 Then
                    '                    If goMinerals(Z).MinPropKnownScore(lProp) + 1 <> 10 Then
                    '                        bGood = True
                    '                        Exit For
                    '                    End If
                    '                End If
                    '            Next lProp
                    '            Exit For
                    '        End If
                    '    Next Z
                    '    If bGood = False Then
                    '        MyBase.moUILib.AddNotification("You already fully understand this material.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                    '        Return
                    '    End If
                    'End If

                    System.BitConverter.GetBytes(GlobalMessageCode.eSetEntityProduction).CopyTo(yData, 0)
                    oResearchGuid.GetGUIDAsString.CopyTo(yData, 2)
                    System.BitConverter.GetBytes(lstMaterials.ItemData(lstMaterials.ListIndex)).CopyTo(yData, 8)
                    'If chkUntilFullyStudied.Value = True AndAlso myClickedButton = 2 Then
                    '    System.BitConverter.GetBytes(-ObjectType.eMineralTech).CopyTo(yData, 12)
                    'Else : 
                    System.BitConverter.GetBytes(ObjectType.eMineralTech).CopyTo(yData, 12)
                    'End If


                    MyBase.moUILib.SendMsgToPrimary(yData)

                    If NewTutorialManager.TutorialOn = True Then
                        Dim iTmpTypeID As Int16 = ObjectType.eMineralTech
                        'If chkUntilFullyStudied.Value = True AndAlso myClickedButton = 2 Then
                        '    iTmpTypeID = -ObjectType.eMineralTech
                        'End If
                        NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eBuildOrderSent, lMineralID, iTmpTypeID, 1, "")
                    End If

                    MyBase.moUILib.RemoveWindow(Me.ControlName)

                    Exit For
                End If
            Next X
        End If
	End Sub

    'Private Sub btnStudy_Click(ByVal sName As String) Handles btnStudy.Click
    '	If goTutorial Is Nothing = False AndAlso goTutorial.TutorialOn = True Then goTutorial.HandleTutorialTrigger(TutorialManager.TutorialTriggerType.MatResStudyButton)

    '	If NewTutorialManager.TutorialOn = True Then
    '		NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eStudyMineralButton, -1, -1, -1, "")
    '	End If

    '	lstMaterials.Clear()
    '	fraMaterialDetails.Visible = True

    '	chkUntilFullyStudied.Visible = True

    '	chkFilterArchived.Visible = True
    '	chkSetFilter.Visible = True

    '	Dim bFilter As Boolean = chkFilterArchived.Value

    '       Dim lSorted() As Int32 = GetSortedMineralIdxArray(False)
    '	If lSorted Is Nothing = False Then
    '		For X As Int32 = 0 To lSorted.GetUpperBound(0)
    '			If bFilter = False OrElse goMinerals(lSorted(X)).bArchived = False Then
    '				lstMaterials.AddItem(goMinerals(lSorted(X)).MineralName)
    '				lstMaterials.ItemData(lstMaterials.NewIndex) = goMinerals(lSorted(X)).ObjectID
    '			End If
    '		Next X
    '	End If

    '	myClickedButton = 2

    '	btnStudy.ControlImageRect = btnStudy.ControlImageRect_Pressed
    'End Sub

	Private Sub btnCancel_Click(ByVal sName As String) Handles btnCancel.Click
		MyBase.moUILib.RemoveWindow(Me.ControlName)
	End Sub

    Private Sub btnClose_Click(ByVal sName As String) Handles btnClose.Click
        MyBase.moUILib.RemoveWindow(Me.ControlName)
    End Sub

	Private Sub lstMaterials_ItemClick(ByVal lIndex As Integer) Handles lstMaterials.ItemClick
        'If myClickedButton = 1 Then Return

		For X As Int32 = 0 To lblProps.GetUpperBound(0)
			lblProps(X).Visible = False
			lblValue(X).Visible = False
            'lblKnown(X).Visible = False
        Next X

        btnResearch.Visible = False

		If lstMaterials.ListIndex > -1 Then
			Dim lMineralID As Int32 = lstMaterials.ItemData(lIndex)

			If NewTutorialManager.TutorialOn = True Then
				NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eStudyMineralListSelected, lMineralID, ObjectType.eMineral, -1, "")
			End If

			For X As Int32 = 0 To glMineralUB
				If glMineralIdx(X) = lMineralID Then
					'ok, found it... now... populate our data
					Dim lTemp As Int32
					'Dim lLeftOffset As Int32

					mlLastUpdateMsg = goMinerals(X).lLastMsgUpdate
					mlCurrentMineralID = lMineralID

					mbIgnoreSetFilter = True
					chkSetFilter.Value = goMinerals(X).bArchived
					mbIgnoreSetFilter = False

					If goMinerals(X).bDiscovered = True Then

						For Y As Int32 = 0 To glMineralPropertyUB
							Dim lIdx As Int32 = goMinerals(X).GetPropertyIndex(glMineralPropertyIdx(Y))
							If lIdx <> -1 Then
								'Set the tooltip
								lblProps(Y).ToolTipText = goMinerals(X).SentenceDescription(lIdx)

								'Now, the actual value...
								lTemp = goMinerals(X).MinPropValueScore(lIdx) '\ 10
							Else
								lTemp = 0
								lblProps(Y).ToolTipText = ""
							End If


							If lTemp < 4 Then
								lblValue(Y).BackImageColor = Color.Red
							ElseIf lTemp < 7 Then
								lblValue(Y).BackImageColor = Color.Yellow
							Else : lblValue(Y).BackImageColor = Color.Green
							End If
							'lLeftOffset = ((lTemp \ 4) * 63) + 50
							'lTemp = lTemp Mod 4
							'lblValue(Y).ControlImageRect = Rectangle.FromLTRB(lLeftOffset, 215 + (lTemp * 9), lLeftOffset + 63, 215 + ((lTemp + 1) * 9))
                            lblValue(Y).ControlImageRect = grc_UI(elInterfaceRectangle.eMinBar_0 + lTemp)

                            ''Now, the Known Value...
                            'If lIdx <> -1 Then lTemp = goMinerals(X).MinPropKnownScore(lIdx) + 1 '\ 10
                            'If lTemp < 4 Then
                            '	lblKnown(Y).BackImageColor = Color.Red
                            'ElseIf lTemp < 7 Then
                            '	lblKnown(Y).BackImageColor = Color.Yellow
                            'Else : lblKnown(Y).BackImageColor = Color.Green
                            'End If
                            ''lLeftOffset = ((lTemp \ 4) * 63) + 50
                            ''lTemp = lTemp Mod 4
                            ''lblKnown(Y).ControlImageRect = Rectangle.FromLTRB(lLeftOffset, 215 + (lTemp * 9), lLeftOffset + 63, 215 + ((lTemp + 1) * 9))
                            'lblKnown(Y).ControlImageRect = Rectangle.FromLTRB(193, 157 + (lTemp * 9), 256, 157 + ((lTemp + 1) * 9))

                            lblProps(Y).Visible = True
                            lblValue(Y).Visible = True
                            'lblKnown(Y).Visible = goMineralProperty(Y).yKnowledgeLevel > eMinPropKnowledgeLevel.eExistence
                        Next Y
                    Else
                        btnResearch.Visible = True
                    End If



					Exit For
				End If
			Next X
		End If
	End Sub

	Private Sub btnHelp_Click(ByVal sName As String) Handles btnHelp.Click
		If goTutorial Is Nothing Then goTutorial = New TutorialManager()
		goTutorial.BeginNewStepType(TutorialManager.TutorialStepType.eMatRes)
	End Sub

	Private Sub frmMatRes_OnNewFrame() Handles Me.OnNewFrame

        If glCurrentCycle - mlLastListCheck > 30 Then ' AndAlso myClickedButton <> 0 Then
            mlLastListCheck = glCurrentCycle
            Dim bFilter As Boolean = chkFilterArchived.Value
            Dim lCnt As Int32 = 0
            For X As Int32 = 0 To glMineralUB
                If glMineralIdx(X) <> -1 Then
                    If goMinerals(X).bArchived = False OrElse bFilter = False Then
                        'If myClickedButton = 1 Then
                        '    If goMinerals(X).bDiscovered = False Then lCnt += 1
                        'ElseIf goMinerals(X).bDiscovered = True Then
                        '    lCnt += 1
                        'End If
                        lCnt += 1
                    End If
                End If
            Next X
            If lCnt <> lstMaterials.ListCount Then
                'If btnDiscover Is Nothing = False AndAlso btnStudy Is Nothing = False Then
                '    If myClickedButton = 1 Then
                '        btnDiscover_Click(btnDiscover.ControlName)
                '    Else : btnStudy_Click(btnStudy.ControlName)
                '    End If
                'End If
                FillMinList()
            End If
        End If

		If lstMaterials Is Nothing = False AndAlso lstMaterials.ListIndex > -1 Then
			For X As Int32 = 0 To glMineralUB
				If glMineralIdx(X) = mlCurrentMineralID Then
                    goMinerals(X).CheckRequestProps()
                    'If goMinerals(X).bRequestedProps = False Then
                    '	goMinerals(X).lLastMsgUpdate = 0
                    '	mlLastUpdateMsg = 0
                    '	goMinerals(X).bRequestedProps = True
                    '	Dim yMsg(5) As Byte
                    '	System.BitConverter.GetBytes(GlobalMessageCode.eRequestMineral).CopyTo(yMsg, 0)
                    '	System.BitConverter.GetBytes(mlCurrentMineralID).CopyTo(yMsg, 2)
                    '	MyBase.moUILib.SendMsgToPrimary(yMsg)
                    'End If
					If mlLastUpdateMsg <> goMinerals(X).lLastMsgUpdate Then
						lstMaterials_ItemClick(lstMaterials.ListIndex)
					End If
					Exit For
				End If
			Next X
		End If
	End Sub

	Private Sub chkFilterArchived_Click() Handles chkFilterArchived.Click
		MyBase.moUILib.GetMsgSys.LoadArchived()
        'If myClickedButton = 2 Then btnStudy_Click(btnStudy.ControlName) Else btnDiscover_Click(btnDiscover.ControlName)
        FillMinList()
	End Sub

	Private Sub chkSetFilter_Click() Handles chkSetFilter.Click
		If mbIgnoreSetFilter = True Then Return
        'If myClickedButton = 1 Then Return

		If lstMaterials Is Nothing = False AndAlso lstMaterials.ListIndex > -1 Then
			Dim lMineralID As Int32 = lstMaterials.ItemData(lstMaterials.ListIndex)

			For X As Int32 = 0 To glMineralUB
				If glMineralIdx(X) = lMineralID Then
                    goMinerals(X).bArchived = chkSetFilter.Value

                    'ok, archiving a mineral (or unarchiving), see if it has an alloy associated to it
                    For Y As Int32 = 0 To goCurrentPlayer.mlTechUB
                        Dim oTech As Base_Tech = goCurrentPlayer.moTechs(Y)
                        If oTech.ObjTypeID = ObjectType.eAlloyTech Then
                            With CType(oTech, AlloyTech)
                                If .AlloyResultID = lMineralID Then
                                    .bArchived = chkSetFilter.Value
                                    Exit For
                                End If
                            End With
                        End If
                    Next Y
                    Exit For
                End If
			Next X

			Dim yMsg(8) As Byte
			Dim lPos As Int32 = 0
			System.BitConverter.GetBytes(GlobalMessageCode.eSetArchiveState).CopyTo(yMsg, lPos) : lPos += 2
			System.BitConverter.GetBytes(lMineralID).CopyTo(yMsg, lPos) : lPos += 4
			System.BitConverter.GetBytes(ObjectType.eMineral).CopyTo(yMsg, lPos) : lPos += 2
			If chkSetFilter.Value = True Then yMsg(lPos) = 1 Else yMsg(lPos) = 0
			lPos += 1

			MyBase.moUILib.SendMsgToPrimary(yMsg)
            lstMaterials.ListIndex = -1
        End If
	End Sub

    Private Sub btnExport_Click(ByVal sName As String) Handles btnExport.Click
        Export_MineralInfo()
    End Sub
    Private Shared mbInExport As Boolean = False
    Private Sub Export_MineralInfo()
        If goCurrentPlayer Is Nothing Then Return
        If mbInExport = True Then Return
        mbInExport = True
        If goUILib Is Nothing = False Then
            goUILib.AddNotification("Downloading complete mineral property info.  This may take a few minutes.", Color.White, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
        End If

        Dim oThread As New Threading.Thread(AddressOf DoTheExport)
        oThread.Start()

    End Sub

    Private Sub DoTheExport()
        Try
            Dim bDone As Boolean = False
            Dim lToDo As Int32 = 0
            Dim lTotalToDo As Int32 = 0
            Dim oProgBar As frmProgressBar = New frmProgressBar(MyBase.moUILib)
            Dim lSorted() As Int32 = GetSortedMineralIdxArray(False, True)
            If lSorted Is Nothing = False Then
                While bDone = False
                    lToDo = 0
                    bDone = True
                    For X As Int32 = 0 To lSorted.GetUpperBound(0)
                        Dim oMineral As Mineral = goMinerals(lSorted(X))
                        Dim lIdx As Int32 = oMineral.GetPropertyIndex(1)
                        If oMineral.bRequestedPropsExport = False Then
                            oMineral.CheckRequestProps()
                            bDone = False
                            lToDo += 1
                        ElseIf lIdx = -1 Then
                            bDone = False
                            lToDo += 1
                        ElseIf oMineral.bReceivedProps = False Then
                            bDone = False
                            lToDo += 1
                        End If
                    Next X
                    If lTotalToDo = 0 AndAlso lToDo > 0 Then
                        lTotalToDo = lToDo
                    End If
                    oProgBar.Name = "Export Percent"
                    oProgBar.Percent = CSng((lTotalToDo - lToDo) / lTotalToDo)

                    If bDone = False Then
                        Threading.Thread.Sleep(10)
                    End If
                End While
                MyBase.moUILib.RemoveWindow("frmProgressBar")
                oProgBar = Nothing

                If muSettings.ExportedDataFormat = 1 Then
                    Export_MineralInfo_Csv()
                ElseIf muSettings.ExportedDataFormat = 2 Then
                    'Export_MineralInfo_Xml()
                End If
            End If
        Catch
            'Do nothing... we could alert the user that Export failed
        End Try

        'set our inExport flag before thread dies
        mbInExport = False
    End Sub

    Private Sub Export_MineralInfo_Csv()
        Dim sExportData As String = ""
        Dim lIdx As Int32 = -1
        Dim iProperty As Int32

        ' Header Row
        sExportData = "Mineral,IsArchived,IsMineral"
        For lIdx = 0 To glMineralPropertyUB
            sExportData &= ","
            sExportData &= goMineralProperty(lIdx).MineralPropertyName
        Next
        sExportData &= vbCrLf


        Dim lSorted() As Int32 = GetSortedMineralIdxArray(False, True)
        If lSorted Is Nothing = False Then
            ' Mineral Data
            For X As Int32 = 0 To lSorted.GetUpperBound(0)
                If goMinerals(lSorted(X)).ObjectID <= 157 OrElse goMinerals(lSorted(X)).ObjectID = 41991 Then

                    sExportData &= Replace(goMinerals(lSorted(X)).MineralName, ",", " ")
                    If goMinerals(lSorted(X)).bArchived Then
                        sExportData &= ",1"
                    Else
                        sExportData &= ",0"
                    End If

                    sExportData &= ",1"

                    For Y As Int32 = 0 To glMineralPropertyUB
                        lIdx = goMinerals(lSorted(X)).GetPropertyIndex(glMineralPropertyIdx(Y))
                        If lIdx <> -1 Then
                            iProperty = goMinerals(lSorted(X)).MinPropValueScore(lIdx)
                            sExportData &= ","
                            sExportData &= iProperty.ToString
                        End If
                    Next Y
                    sExportData &= vbCrLf
                End If
            Next X

            ' Alloy Data
            For X As Int32 = 0 To lSorted.GetUpperBound(0)
                If goMinerals(lSorted(X)).ObjectID > 157 AndAlso goMinerals(lSorted(X)).ObjectID <> 41991 Then
                    sExportData &= Replace(goMinerals(lSorted(X)).MineralName, ",", " ")
                    If goMinerals(lSorted(X)).bArchived Then
                        sExportData &= ",1"
                    Else
                        sExportData &= ",0"
                    End If
                    sExportData &= ",0"

                    For Y As Int32 = 0 To glMineralPropertyUB
                        lIdx = goMinerals(lSorted(X)).GetPropertyIndex(glMineralPropertyIdx(Y))
                        If lIdx <> -1 Then
                            iProperty = goMinerals(lSorted(X)).MinPropValueScore(lIdx)
                            sExportData &= ","
                            sExportData &= iProperty.ToString
                        End If
                    Next Y
                    sExportData &= vbCrLf
                End If
            Next X
        End If

        'Create the File
        If sExportData = "" Then Return
        Dim sFile As String = AppDomain.CurrentDomain.BaseDirectory
        If sFile.EndsWith("\") = False Then sFile = sFile & "\"
        sFile &= "ExportedData"
        If Exists(sFile) = False Then MkDir(sFile)
        If sFile.EndsWith("\") = False Then sFile &= "\"
        sFile &= "Minerals_" & goCurrentPlayer.PlayerName & "_" & Now.ToString("MM_dd_yyyy_HHmmss") & ".csv"

        Dim fsFile As IO.FileStream = New IO.FileStream(sFile, IO.FileMode.Create)

        Dim info As Byte() = New System.Text.UTF8Encoding(True).GetBytes(sExportData)
        fsFile.Write(info, 0, info.Length)
        fsFile.Close()
        fsFile.Dispose()
        If goUILib Is Nothing = False Then
            goUILib.AddNotification("Mineral Info Exported.", Color.White, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
        End If
    End Sub
End Class
