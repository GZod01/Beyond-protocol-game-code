


Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

Partial Class frmDiplomacy

    'Interface created from Interface Builder
    Public Class fraPlayerRelList
        Inherits UIWindow

        Private WithEvents vscrScroll As UIScrollBar

        Private moItems() As ctlDiplomacy
        Private mlItemUB As Int32 = -1

        Private moList() As PlayerRel
        Private mlListUB As Int32 = -1
        Private mlSelectedItem As Int32 = -1

        Private mlDisplayItemCnt As Int32 = 0
        Private mbHasUnknowns As Boolean = False
        Private mlLastUpdate As Int32

        Private mbNeedNames As Boolean = True

        Public Event ListItemClicked(ByVal lPlayerID As Int32)
        Public Event ItemDataChanged(ByVal lPlayerID As Int32, ByVal lRelVal As Int32)

        Public Sub New(ByRef oUILib As UILib)
            MyBase.New(oUILib)

            'fraPlayerRelList initial props
            With Me
                .ControlName = "fraPlayerRelList"
                .Left = 82
                .Top = 66
                .Width = 735
                .Height = 280
                .Enabled = True
                .Visible = True
                .BorderColor = muSettings.InterfaceBorderColor
                .FillColor = muSettings.InterfaceFillColor
                .FullScreen = False
                .BorderLineWidth = 1
                .Moveable = False
                .mbAcceptReprocessEvents = True
            End With

            'vscrScroll initial props
            vscrScroll = New UIScrollBar(oUILib, True)
            With vscrScroll
                .ControlName = "vscrScroll"
                .Left = 711
                .Top = 4
                .Width = 24
                .Height = 272
                .Enabled = True
                .Visible = True
                .Value = 0
                .MaxValue = 0
                .MinValue = 0
                .SmallChange = 1
                .LargeChange = 4
                .ReverseDirection = True
                .mbAcceptReprocessEvents = True
            End With
            Me.AddChild(CType(vscrScroll, UIControl))

        End Sub

        Public Sub ClearAllItemsTempScore()
            For X As Int32 = 0 To mlItemUB
                If moItems(X) Is Nothing = False Then
                    moItems(X).ClearTempScore()
                End If
            Next X
        End Sub

        Protected Overrides Sub Finalize()
			If goResMgr Is Nothing = False Then
				goResMgr.DeleteTexture("Diplomacy.dds")
				goResMgr.DeleteTexture("DipPlayerBack.bmp")
				goResMgr.DeleteTexture("DipPlayerFore.dds")
			End If

            If ctlDiplomacy.moRelBarImg Is Nothing = False Then ctlDiplomacy.moRelBarImg.Dispose()
            ctlDiplomacy.moRelBarImg = Nothing

            If ctlDiplomacy.moIconBack Is Nothing = False Then ctlDiplomacy.moIconBack.Dispose()
            ctlDiplomacy.moIconBack = Nothing

            If ctlDiplomacy.moIconFore Is Nothing = False Then ctlDiplomacy.moIconFore.Dispose()
            ctlDiplomacy.moIconFore = Nothing

            If ctlDiplomacy.moSprite Is Nothing = False Then ctlDiplomacy.moSprite.Dispose()
			ctlDiplomacy.moSprite = Nothing
			MyBase.Finalize()
        End Sub

        Public Sub SetFromCurrentPlayer()
            With goCurrentPlayer
                If .yPlayerPhase = 1 Then
                    Dim oRel As PlayerRel = .GetPlayerRel(gl_HARDCODE_PIRATE_PLAYER_ID)
                    If oRel Is Nothing = True Then
                        .SetPlayerRel(gl_HARDCODE_PIRATE_PLAYER_ID, 60)
                        oRel = .GetPlayerRel(gl_HARDCODE_PIRATE_PLAYER_ID)
                        oRel.lPlayerRegards = glPlayerID
                        oRel.lThisPlayer = gl_HARDCODE_PIRATE_PLAYER_ID
                        oRel.WithThisScore = 60
                        Dim oPI As New PlayerIntel
                        oPI.EmpireName = "Aurelium"
                        oPI.bIsMale = True
                        oPI.lPlayerIcon = 721216
                        oPI.ObjectID = gl_HARDCODE_PIRATE_PLAYER_ID
                        oPI.ObjTypeID = ObjectType.ePlayerIntel
                        oPI.PlayerName = "Aurelium Pirates"
                        oPI.RaceName = "Pirate"
                        oPI.yCustomTitle = Player.eyCustomRank.King
                        oPI.yPlayerTitle = Player.eyCustomRank.King
                        oPI.yRegardsCurrentPlayer = 1
                        oPI.lTechnologyScore = Int32.MinValue
                        oPI.lDiplomacyScore = Int32.MinValue
                        oPI.lMilitaryScore = Int32.MinValue
                        oPI.lPopulationScore = Int32.MinValue
                        oPI.lProductionScore = Int32.MinValue
                        oPI.lTechnologyScore = Int32.MinValue
                        oPI.lWealthScore = Int32.MinValue
                        oPI.lTotalScore = 0
                        glPlayerIntelUB += 1
                        ReDim Preserve glPlayerIntelIdx(glPlayerIntelUB)
                        ReDim Preserve goPlayerIntel(glPlayerIntelUB)
                        glPlayerIntelIdx(glPlayerIntelUB) = gl_HARDCODE_PIRATE_PLAYER_ID
                        goPlayerIntel(glPlayerIntelUB) = oPI
                    End If
                End If

                mlListUB = .PlayerRelUB
                ReDim moList(mlListUB)

                For X As Int32 = 0 To mlListUB
                    moList(X) = .GetPlayerRelByIndex(X)
                Next X

                Dim lTmp As Int32 = mlListUB + 1
                lTmp -= mlDisplayItemCnt
                If lTmp < 0 Then
                    vscrScroll.Value = 0
                    vscrScroll.MaxValue = 0
                Else : vscrScroll.MaxValue = lTmp
                End If

                Me.IsDirty = True
            End With
        End Sub

        Private Sub ItemClicked(ByVal lPlayerID As Int32)
            mlSelectedItem = -1
            For X As Int32 = 0 To mlListUB
                If moList(X).oPlayerIntel Is Nothing = False AndAlso moList(X).oPlayerIntel.ObjectID = lPlayerID Then
                    mlSelectedItem = X
                    Exit For
                End If
            Next X
            RaiseEvent ListItemClicked(lPlayerID)
            Me.IsDirty = True
        End Sub
        Private Sub ItemRelChanged(ByVal lPlayerID As Int32, ByVal lRelVal As Int32)
            RaiseEvent ItemDataChanged(lPlayerID, lRelVal)
        End Sub

		Public Sub fraPlayerRelList_OnNewFrame() Handles Me.OnNewFrame
			If mbHasUnknowns = True Then
				mbHasUnknowns = False
				For X As Int32 = 0 To mlItemUB
					mbHasUnknowns = mbHasUnknowns OrElse moItems(X).HasUnknowns()
				Next X
            End If

            If mbNeedNames = True Then
                mbNeedNames = False
                Try
                    For X As Int32 = 0 To mlListUB
                        Dim sCache As String = GetCacheObjectValue(moList(X).lThisPlayer, ObjectType.ePlayer)
                        If sCache = "Unknown" Then
                            mbNeedNames = True
                        End If
                    Next X


                    If mbNeedNames = False Then
                        'Ok, let's sort our items...
                        Dim lIdx() As Int32 = GetSortedPlayerRelIdxArray()
                        Dim oTmp(mlListUB) As PlayerRel
                        For X As Int32 = 0 To mlListUB
                            oTmp(X) = goCurrentPlayer.GetPlayerRelByIndex(lIdx(X))
                        Next X
                        moList = oTmp
                    End If
                Catch
                    mbNeedNames = True
                End Try
            End If
		End Sub

		Private Sub fraPlayerRelList_OnRender() Handles Me.OnRender
			'First, check if we need to create the detail item classes...
			If mlItemUB <> mlDisplayItemCnt - 1 Then
				Dim bDone As Boolean = False
				While bDone = False
					bDone = True
					For X As Int32 = 0 To Me.ChildrenUB
						If Me.moChildren(X).ControlName.ToUpper = "CTLDIPLOMACY" Then
							bDone = False
							Dim oTmp As ctlDiplomacy = CType(moChildren(X), ctlDiplomacy)
                            RemoveHandler oTmp.ItemClicked, AddressOf ItemClicked
                            RemoveHandler oTmp.ItemRelChanged, AddressOf ItemRelChanged
							oTmp = Nothing
							Me.RemoveChild(X)
							Exit For
						End If
					Next X
				End While

				mlItemUB = mlDisplayItemCnt - 1
				ReDim moItems(mlItemUB)
				For X As Int32 = 0 To mlItemUB
					moItems(X) = New ctlDiplomacy(MyBase.moUILib)
					With moItems(X)
						.Left = 2
						.Top = 5 + (X * .Height)
						.Visible = False
					End With
					Me.AddChild(CType(moItems(X), UIControl))
                    AddHandler moItems(X).ItemClicked, AddressOf ItemClicked
                    AddHandler moItems(X).ItemRelChanged, AddressOf ItemRelChanged
				Next X
			End If

			'Now, refresh our display
			mbHasUnknowns = False
			For X As Int32 = 0 To mlDisplayItemCnt - 1
				Dim lIdx As Int32 = vscrScroll.Value + X
				If lIdx > mlListUB Then
					If moItems(X).Visible = True Then moItems(X).Visible = False
					moItems(X).Selected = False
				Else
					If moItems(X).Visible = False Then moItems(X).Visible = True
					moItems(X).Selected = (lIdx = mlSelectedItem)

					moItems(X).SetValues(moList(lIdx))
					mbHasUnknowns = moItems(X).HasUnknowns() OrElse mbHasUnknowns
				End If
            Next X
			mlLastUpdate = glCurrentCycle
		End Sub

		Private Sub fraPlayerRelList_OnRenderEnd() Handles Me.OnRenderEnd
            If ctlDiplomacy.moSprite Is Nothing OrElse ctlDiplomacy.moSprite.Disposed = True Then
                Device.IsUsingEventHandlers = False
                ctlDiplomacy.moSprite = New Sprite(MyBase.moUILib.oDevice)
                Device.IsUsingEventHandlers = True
            End If

			ctlDiplomacy.moSprite.Begin(SpriteFlags.AlphaBlend)
			Try
				Dim rcDest As Rectangle = New Rectangle(0, 0, 64, 64)
				For X As Int32 = 0 To mlItemUB
					If moItems(X).Visible = True Then
						Dim ptDest As Point
						ptDest.X = Me.Left + moItems(X).Left + 3
						ptDest.Y = Me.Top + moItems(X).Top + 3

						With moItems(X)
							ctlDiplomacy.moSprite.Draw2D(ctlDiplomacy.moIconBack, .rcBack, rcDest, ptDest, .clrBack)
							ctlDiplomacy.moSprite.Draw2D(ctlDiplomacy.moIconFore, .rcFore1, rcDest, ptDest, .clrFore1)
							ctlDiplomacy.moSprite.Draw2D(ctlDiplomacy.moIconFore, .rcFore2, rcDest, ptDest, .clrFore2)

							ptDest = ctlDiplomacy.ptRelBarLoc
							ptDest.Y += Me.Top + .Top

							ctlDiplomacy.moSprite.Draw2D(ctlDiplomacy.moRelBarImg, ctlDiplomacy.rcRelBarSrc, ctlDiplomacy.rcRelBarLoc, ptDest, Color.White)

							Dim lTempX As Int32 = ptDest.X
							ptDest.X += CInt(.yTheirRelScore)
							ptDest.Y -= 2
							ctlDiplomacy.moSprite.Draw2D(ctlDiplomacy.moRelBarImg, Rectangle.FromLTRB(11, 26, 20, 31), Rectangle.FromLTRB(0, 0, 9, 5), ptDest, Color.White)
							ptDest.X = lTempX + CInt(.yMyRelScore)
							ptDest.Y += ctlDiplomacy.rcRelBarLoc.Height
							ctlDiplomacy.moSprite.Draw2D(ctlDiplomacy.moRelBarImg, Rectangle.FromLTRB(0, 26, 9, 31), Rectangle.FromLTRB(0, 0, 9, 5), ptDest, Color.White)


						End With
					End If
				Next X
			Catch
				'do nothing?
			End Try
			ctlDiplomacy.moSprite.End()
		End Sub

        Private Sub fraPlayerRelList_OnResize() Handles Me.OnResize
            Dim lSize As Int32 = Me.Height - -10

            mlDisplayItemCnt = lSize \ ctlDiplomacy.ml_ITEM_HEIGHT
        End Sub

        Private Sub vscrScroll_ValueChange() Handles vscrScroll.ValueChange
            Me.IsDirty = True
            ClearAllItemsTempScore()
        End Sub

        Public Sub Export_DiplomacyInfo()
            If goCurrentPlayer Is Nothing Then Return
            If muSettings.ExportedDataFormat = 1 Then
                Export_DiplomacyInfo_Csv()
            ElseIf muSettings.ExportedDataFormat = 2 Then
                'Export_DiplomacyInfo_Xml()
            End If
        End Sub

        Private Sub Export_DiplomacyInfo_Csv()
            'TODO: Thread this to ensure cached Player names and Guild names

            Dim sExportData As String = ""
            Dim sName As String = ""
            Dim sTitle As String = ""
            Dim sGuildName As String = ""
            Dim oRel As PlayerRel
            Dim lTotalVoteStrength As Int32 = 0

            Dim lModMine As Int32 = goCurrentPlayer.lTotalScore
            Dim lIdx() As Int32 = GetSortedPlayerRelIdxArray()
            Dim oTmp(mlListUB) As PlayerRel
            For X As Int32 = 0 To mlListUB
                oTmp(X) = goCurrentPlayer.GetPlayerRelByIndex(lIdx(X))
            Next X
            sExportData &= "PlayerName,Rank,Votes,GuildName,TargetRelationship,TowardsThem,TowardsUs,ThreatLevel,TotalScore,Technology,Diplomacy,Military,Population,Production,Wealth" & vbCrLf

            sName = Replace(goCurrentPlayer.PlayerName, ",", " ")
            sTitle = Player.GetPlayerNameWithTitle(goCurrentPlayer.yPlayerTitle, "", goCurrentPlayer.bIsMale).Trim
            If Not goCurrentPlayer.oGuild Is Nothing Then
                sGuildName = goCurrentPlayer.oGuild.sName
            End If
            sExportData &= sName & "," & sTitle & "," & lTotalVoteStrength & "," & sGuildName & ",-1,-1,-1,-1,"
            sExportData &= goCurrentPlayer.lTotalScore.ToString
            sExportData &= ","
            sExportData &= goCurrentPlayer.lTechnologyScore.ToString
            sExportData &= ","
            sExportData &= goCurrentPlayer.lDiplomacyScore.ToString
            sExportData &= ","
            sExportData &= goCurrentPlayer.lMilitaryScore.ToString
            sExportData &= ","
            sExportData &= goCurrentPlayer.lPopulationScore.ToString
            sExportData &= ","
            sExportData &= goCurrentPlayer.lProductionScore.ToString
            sExportData &= ","
            sExportData &= goCurrentPlayer.lWealthScore.ToString
            sExportData &= vbCrLf


            For X As Int32 = 0 To mlListUB
                sName = Replace(GetCacheObjectValue(oTmp(X).lThisPlayer, ObjectType.ePlayer), ",", " ")
                oRel = goCurrentPlayer.GetPlayerRel(oTmp(X).lThisPlayer)
                sExportData &= sName
                sExportData &= ","

                sTitle = Player.GetPlayerNameWithTitle(oRel.oPlayerIntel.yPlayerTitle, "", oRel.oPlayerIntel.bIsMale).Trim
                sExportData &= sTitle
                sExportData &= ","
                sExportData &= oRel.oPlayerIntel.lTotalVoteStrength
                sExportData &= ","
                If oRel.oPlayerIntel.lGuildID <> -1 Then
                    sGuildName = GetCacheObjectValue(oRel.oPlayerIntel.lGuildID, ObjectType.eGuild)
                    If sGuildName <> "Unknown" Then
                        sExportData &= Replace(sGuildName, ",", " ")
                    End If
                End If
                sExportData &= ","
                sExportData &= oTmp(X).TargetScore.ToString
                sExportData &= ","
                sExportData &= oTmp(X).WithThisScore.ToString
                sExportData &= ","
                sExportData &= oRel.oPlayerIntel.yRegardsCurrentPlayer.ToString
                sExportData &= ","
                Dim lModTheirs As Int32 = oTmp(X).oPlayerIntel.lTotalScore


                If oTmp(X).oPlayerIntel.lTotalScore < 0 OrElse goCurrentPlayer.lTotalScore < 0 Then
                    Dim lMod As Int32 = Math.Abs(Math.Min(lModTheirs, lModMine))
                    lModTheirs += lMod
                    lModMine += lMod
                End If

                Dim fTemp As Single = 0.0F
                If lModTheirs + lModMine = 0 Then
                    fTemp = 0.5F
                Else : fTemp = CSng(lModTheirs / (lModTheirs + lModMine))
                End If

                Dim lValue As Int32 = CInt(fTemp * 10)
                If lValue < 0 Then lValue = 0
                If lValue > 9 Then lValue = 9
                lValue += 1
                sExportData &= lValue.ToString
                sExportData &= ","
                sExportData &= lModTheirs
                sExportData &= ","
                If oTmp(X).oPlayerIntel.lTechnologyScore >= 0 Then
                    sExportData &= oTmp(X).oPlayerIntel.lTechnologyScore.ToString
                End If
                sExportData &= ","
                If oTmp(X).oPlayerIntel.lDiplomacyScore >= 0 Then
                    sExportData &= oTmp(X).oPlayerIntel.lDiplomacyScore.ToString
                End If
                sExportData &= ","
                If oTmp(X).oPlayerIntel.lMilitaryScore >= 0 Then
                    sExportData &= oTmp(X).oPlayerIntel.lMilitaryScore.ToString
                End If
                sExportData &= ","
                If oTmp(X).oPlayerIntel.lPopulationScore >= 0 Then
                    sExportData &= oTmp(X).oPlayerIntel.lPopulationScore.ToString
                End If
                sExportData &= ","
                If oTmp(X).oPlayerIntel.lProductionScore >= 0 Then
                    sExportData &= oTmp(X).oPlayerIntel.lProductionScore.ToString
                End If
                sExportData &= ","
                If oTmp(X).oPlayerIntel.lWealthScore >= 0 Then
                    sExportData &= oTmp(X).oPlayerIntel.lWealthScore.ToString
                End If
                sExportData &= vbCrLf
            Next
            If sExportData = "" Then Return
            Dim sFile As String = AppDomain.CurrentDomain.BaseDirectory
            If sFile.EndsWith("\") = False Then sFile = sFile & "\"
            sFile &= "ExportedData"
            If Exists(sFile) = False Then MkDir(sFile)
            If sFile.EndsWith("\") = False Then sFile &= "\"
            sFile &= "Diplomacy_" & goCurrentPlayer.PlayerName & "_" & Now.ToString("MM_dd_yyyy_HHmmss") & ".csv"

            Dim fsFile As IO.FileStream = New IO.FileStream(sFile, IO.FileMode.Create)

            Dim info As Byte() = New System.Text.UTF8Encoding(True).GetBytes(sExportData)
            fsFile.Write(info, 0, info.Length)
            fsFile.Close()
            fsFile.Dispose()
            If goUILib Is Nothing = False Then
                goUILib.AddNotification("Diplomacy Info Exported.", Color.White, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            End If
        End Sub
    End Class
End Class