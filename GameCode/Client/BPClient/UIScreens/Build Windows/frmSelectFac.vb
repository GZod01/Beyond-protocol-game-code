Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Interface created from Interface Builder
Public Class frmSelectFac
    Inherits UIWindow

	Private lblTitle As UILabel
    Private WithEvents lstFacs As UIListBox
    Private lblModulesUsed As UILabel
	Private WithEvents btnClose As UIButton
    Private WithEvents btnDismantle As UIButton


    Private mlParentEntityIndex As Int32 = -1

    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)

        'frmSelectFac initial props
        With Me
            .lWindowMetricID = BPMetrics.eWindow.eSelectFac
            .ControlName = "frmSelectFac"
            Dim oTmpWin As UIWindow = Nothing
            oTmpWin = oUILib.GetWindow("frmBuildWindow")
            If oTmpWin Is Nothing Then oTmpWin = oUILib.GetWindow("frmResearchMain")

            If oTmpWin Is Nothing Then
                .Left = (oUILib.oDevice.PresentationParameters.BackBufferWidth \ 2) - 128
                .Top = (oUILib.oDevice.PresentationParameters.BackBufferHeight \ 2) - 96
            Else
                Dim lTemp As Int32
                lTemp = oTmpWin.Left - 256
                If lTemp < 0 Then lTemp = oTmpWin.Left + oTmpWin.Width
                .Left = lTemp
                .Top = oTmpWin.Top
                .Height = oTmpWin.Height
            End If
            .Width = 256
            '.Height = 192
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = False
            .mbAcceptReprocessEvents = True
        End With

        'lblTitle initial props
        lblTitle = New UILabel(oUILib)
        With lblTitle
            .ControlName = "lblTitle"
            .Left = 5
            .Top = 5
            .Width = 117
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Select a Facility:"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblTitle, UIControl))

        'lstFacs initial props
        lstFacs = New UIListBox(oUILib)
        With lstFacs
            .ControlName = "lstFacs"
            .Left = 5
            .Top = 30
            .Width = 245
            .Height = Me.Height - .Top - 6
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
            .mbAcceptReprocessEvents = True
        End With
        Me.AddChild(CType(lstFacs, UIControl))

        'TODO: Send MaxHP with info so we can show module count
        'lblModulesUsed initial props
        'lblModulesUsed = New UILabel(oUILib)
        'With lblModulesUsed
        '    .ControlName = "lblModulesUsed"
        '    .Left = 6
        '    .Top = Me.Height - 18 - 5
        '    .Width = 200
        '    .Height = 18
        '    .Enabled = True
        '    .Visible = True
        '    .Caption = "0/0 Modules Used."
        '    .ForeColor = muSettings.InterfaceBorderColor
        '    .FontFormat = CType(0, DrawTextFormat)
        '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
        '    .DrawBackImage = False
        'End With
        'Me.AddChild(CType(lblModulesUsed, UIControl))

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

		'btnDismantle initial props
		btnDismantle = New UIButton(oUILib)
		With btnDismantle
			.ControlName = "btnDismantle"
			.Left = btnClose.Left - 100
            .Top = Me.BorderLineWidth
			.Width = 100
			.Height = 24
			.Enabled = True
			.Visible = True
			.Caption = "Dismantle"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = True
			.FontFormat = CType(5, DrawTextFormat)
			.ControlImageRect = New Rectangle(0, 0, 120, 32)
		End With
		Me.AddChild(CType(btnDismantle, UIControl))

        MyBase.moUILib.RemoveWindow(Me.ControlName)
        MyBase.moUILib.AddWindow(Me)
    End Sub

    Public Sub SetFromEntityIndex(ByVal lIndex As Int32)
        If goCurrentEnvir Is Nothing Then Return
        If lIndex = -1 Then Return
        If goCurrentEnvir.lEntityIdx(lIndex) = -1 Then Return

        mlParentEntityIndex = lIndex

        lstFacs.Clear()
    End Sub

    Private Sub frmSelectFac_OnNewFrame() Handles Me.OnNewFrame
        If goCurrentEnvir Is Nothing Then Return
        If mlParentEntityIndex = -1 Then Return
        If goCurrentEnvir.lEntityIdx(mlParentEntityIndex) = -1 Then Return

        Dim lCnt As Int32 = 0
        For X As Int32 = 0 To goCurrentEnvir.oEntity(mlParentEntityIndex).lChildUB
            If goCurrentEnvir.oEntity(mlParentEntityIndex).lChildIdx(X) > -1 Then
                lCnt += 1
            End If
        Next X
        'TODO: Send MaxHP with info so we can show module count
        'If goCurrentEnvir.oEntity(mlParentEntityIndex).oUnitDef Is Nothing = False Then
        '    Dim sModCnt As String
        '    Dim x As EntityDef = goCurrentEnvir.oEntity(mlParentEntityIndex).oUnitDef

        '    Dim iMyHp As Int32 = goCurrentEnvir.oEntity(mlParentEntityIndex).oUnitDef.Structure_MaxHP
        '    Dim iModLimit As Int32 = iMyHp \ 50000
        '    sModCnt = lCnt.ToString & "/" & iModLimit.ToString & " Modules Used."
        '    If lblModulesUsed.Caption <> sModCnt Then lblModulesUsed.Caption = sModCnt
        'End If


        If lstFacs.ListCount <> lCnt + 1 OrElse goCurrentEnvir.oEntity(mlParentEntityIndex).bChildListUpdated = True Then
            lstFacs.Clear()

            'Add the parent facility
            lstFacs.AddItem(goCurrentEnvir.oEntity(mlParentEntityIndex).EntityName)
            lstFacs.ItemData(lstFacs.NewIndex) = -1
            lstFacs.ItemData2(lstFacs.NewIndex) = -1

            Dim lSortedIndex() As Int32
            ReDim lSortedIndex(goCurrentEnvir.oEntity(mlParentEntityIndex).lChildUB)

            For X As Int32 = 0 To lSortedIndex.GetUpperBound(0)
                lSortedIndex(X) = -1
            Next X

            With goCurrentEnvir.oEntity(mlParentEntityIndex)
                'First, sort our list...
                For X As Int32 = 0 To .lChildUB
                    If .lChildIdx(X) > 0 Then
                        'Find where this one goes...
                        For Y As Int32 = 0 To lSortedIndex.GetUpperBound(0)
                            If lSortedIndex(Y) = -1 Then
                                lSortedIndex(Y) = X
                                Exit For
                            ElseIf .oChild(lSortedIndex(Y)).oChildDef.DefName > .oChild(X).oChildDef.DefName Then
                                'Ok, found it... shift all items after this
                                For Z As Int32 = lSortedIndex.GetUpperBound(0) To Y + 1 Step -1
                                    lSortedIndex(Z) = lSortedIndex(Z - 1)
                                Next Z
                                lSortedIndex(Y) = X
                                Exit For
                            End If
                        Next Y
                    End If
                Next X

                'Now, fill our list...
                For X As Int32 = 0 To lSortedIndex.GetUpperBound(0)
                    If lSortedIndex(X) = -1 Then Continue For
                    If .oChild(lSortedIndex(X)).oChildDef Is Nothing = False Then
                        lstFacs.AddItem(.oChild(lSortedIndex(X)).oChildDef.DefName)
                        lstFacs.ItemData(lstFacs.NewIndex) = .oChild(lSortedIndex(X)).lChildID
                        lstFacs.ItemData2(lstFacs.NewIndex) = .oChild(lSortedIndex(X)).iChildTypeID
                    End If
                Next X
            End With

            goCurrentEnvir.oEntity(mlParentEntityIndex).bChildListUpdated = False

        End If

    End Sub

    Private Sub lstFacs_ItemClick(ByVal lIndex As Integer) Handles lstFacs.ItemClick
        If lIndex <> -1 Then
            'ok... 
            Dim lID As Int32 = lstFacs.ItemData(lIndex)
            Dim iTypeID As Int16 = CShort(lstFacs.ItemData2(lIndex))

			btnDismantle.Caption = "Dismantle"
            'now... get our windows...
            If lID = -1 OrElse iTypeID = -1 Then
                Dim frmBuild As frmBuildWindow = CType(MyBase.moUILib.GetWindow("frmBuildWindow"), frmBuildWindow)
                If frmBuild Is Nothing Then frmBuild = New frmBuildWindow(goUILib, False)
                If frmBuild Is Nothing = False Then
                    frmBuild.Visible = True
					frmBuild.UpdateFromEntity(mlParentEntityIndex, -1, -1)
				End If
				MyBase.moUILib.RemoveWindow("frmResearchMain")
				btnDismantle.Enabled = False
            Else
				Dim yProd As Byte
				btnDismantle.Enabled = True
                With goCurrentEnvir.oEntity(mlParentEntityIndex)
                    Dim oChild As StationChild = .GetChild(lID, iTypeID)
                    If oChild Is Nothing = False Then
                        yProd = oChild.oChildDef.ProductionTypeID

                        If yProd = ProductionType.eResearch Then
                            Dim frmResMain As frmResearchMain = CType(MyBase.moUILib.GetWindow("frmResearchMain"), frmResearchMain)
                            If frmResMain Is Nothing Then
                                frmResMain = New frmResearchMain(MyBase.moUILib)
                            End If
                            frmResMain.Visible = True

                            Me.Left = frmResMain.Left - Me.Width
                            If Me.Left < 0 Then
                                Me.Left = frmResMain.Left + frmResMain.Width
                            End If
                            Me.Top = frmResMain.Top

                            frmResMain = Nothing

                            MyBase.moUILib.RemoveWindow("frmBuildWindow")
                            lstFacs.HasFocus = True
                        Else
                            Dim frmBuild As frmBuildWindow = CType(MyBase.moUILib.GetWindow("frmBuildWindow"), frmBuildWindow)

                            If frmBuild Is Nothing Then
                                frmBuild = New frmBuildWindow(MyBase.moUILib, False)
							End If
                            frmBuild.Visible = True
                            frmBuild.UpdateFromEntity(mlParentEntityIndex, lID, iTypeID)
                            Me.Left = frmBuild.Left - Me.Width
                            If Me.Left < 0 Then
                                Me.Left = frmBuild.Left + frmBuild.Width
                            End If
                            Me.Top = frmBuild.Top
                            frmBuild = Nothing

                            MyBase.moUILib.RemoveWindow("frmResearchMain")
                        End If
                    End If
                End With
            End If

            'the rest are regardless...
            Dim ofrmProdStatus As frmProdStatus = CType(MyBase.moUILib.GetWindow("frmProdStatus"), frmProdStatus)
            If ofrmProdStatus Is Nothing = False Then ofrmProdStatus.SetFromEntity(mlParentEntityIndex, lID, iTypeID)
            ofrmProdStatus = Nothing

            Dim ofrmAdv As frmAdvanceDisplay = CType(MyBase.moUILib.GetWindow("frmAdvanceDisplay"), frmAdvanceDisplay)
            If ofrmAdv Is Nothing = False Then ofrmAdv.SetFromEntityWithChild(mlParentEntityIndex, lID, iTypeID)
            ofrmAdv = Nothing

        End If
    End Sub

    Public Function GetCurrentChild() As StationChild
        If lstFacs.ListIndex > -1 Then
            Dim lID As Int32 = lstFacs.ItemData(lstFacs.ListIndex)
            Dim iTypeID As Int16 = CShort(lstFacs.ItemData2(lstFacs.ListIndex))

            Return goCurrentEnvir.oEntity(mlParentEntityIndex).GetChild(lID, iTypeID)
        End If
        Return Nothing
    End Function

	Private Sub btnClose_Click(ByVal sName As String) Handles btnClose.Click
        MyBase.moUILib.RemoveWindow("frmBuildWindow")
        MyBase.moUILib.RemoveWindow("frmResearchMain")
        MyBase.moUILib.RemoveWindow(Me.ControlName)
	End Sub

	Private Sub btnDismantle_Click(ByVal sName As String) Handles btnDismantle.Click
		If btnDismantle.Caption.ToUpper = "CONFIRM" Then
			'dismantle
			Dim lID As Int32 = lstFacs.ItemData(lstFacs.ListIndex)
			Dim iTypeID As Int16 = CShort(lstFacs.ItemData2(lstFacs.ListIndex))
			If lID = -1 OrElse iTypeID = -1 Then Return

			Dim yMsg(13) As Byte
			Dim lPos As Int32 = 0
			System.BitConverter.GetBytes(GlobalMessageCode.eSetDismantleTarget).CopyTo(yMsg, lPos) : lPos += 2

			With goCurrentEnvir.oEntity(mlParentEntityIndex)
				.GetGUIDAsString.CopyTo(yMsg, lPos) : lPos += 6
				.RemoveChild(lID, iTypeID)
			End With
			System.BitConverter.GetBytes(lID).CopyTo(yMsg, lPos) : lPos += 4
			System.BitConverter.GetBytes(iTypeID).CopyTo(yMsg, lPos) : lPos += 2
			MyBase.moUILib.SendMsgToPrimary(yMsg)

			lstFacs.RemoveItem(lstFacs.ListIndex)
			btnDismantle.Caption = "Dismantle"
		Else
			btnDismantle.Caption = "Confirm"
		End If
		
    End Sub

    Private Sub frmSelectFac_WindowMoved() Handles Me.WindowMoved
        Dim ofrm As UIWindow = MyBase.moUILib.GetWindow("frmBuildWindow")
        If ofrm Is Nothing = False Then
            ofrm.Left = Me.Left + Me.Width
            ofrm.Top = Me.Top
        Else
            ofrm = MyBase.moUILib.GetWindow("frmResearchMain")
            If ofrm Is Nothing = False Then
                ofrm.Left = Me.Left + Me.Width
                ofrm.Top = Me.Top
            End If
        End If
    End Sub
End Class