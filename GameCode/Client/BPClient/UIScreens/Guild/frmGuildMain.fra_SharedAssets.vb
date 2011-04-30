Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

Partial Class frmGuildMain
    Private Class fra_SharedAssets
        Inherits guildframe

        Private fraSharedAssetsList As UIWindow
        Private WithEvents btnRefresh As UIButton
        Private WithEvents tvwMain As UITreeView

        Private Shared moGuildShareAssets(-1) As GuildShareAsset
        Private mbLoading As Boolean = True
        Private mbHasUnknowns As Boolean = False
        Private mlLastUnknownCheck As Int32 = -1
        Private miLastRefresh As Int32 = -1

        Structure GuildShareAsset
            Dim iObjTypeID As Int16
            Dim lObjectID As Int32
            Dim iEnvirTypeID As Int16
            Dim lEnvirID As Int32
            Dim iLocX As Int32
            Dim iLocZ As Int32
            Dim sName As String
            'Dim lWarpointUpkeep As Int64
            Dim lIdx As Int32
            Dim oParent As UITreeView.UITreeViewItem
        End Structure

        Public Sub New(ByRef oUILib As UILib)
            MyBase.New(oUILib)

            With Me
                .ControlName = "fra_SharedAssets"
                .Left = 15
                .Top = 5
                .Width = 128
                .Height = 128
                .Enabled = True
                .Visible = True
                .BorderColor = muSettings.InterfaceBorderColor
                .FillColor = muSettings.InterfaceFillColor
                .FullScreen = False
                .BorderLineWidth = 2
                .Moveable = False
            End With

            'fraSharedAssetsList initial props
            fraSharedAssetsList = New UIWindow(MyBase.moUILib)
            With fraSharedAssetsList
                .ControlName = "fraSharedAssetsList"
                .Left = 15
                .Top = 5
                .Width = 266
                .Height = 515
                .Enabled = True
                .Visible = True
                .BorderColor = muSettings.InterfaceBorderColor
                .FillColor = muSettings.InterfaceFillColor
                .FullScreen = False
                .Moveable = False
                .BorderLineWidth = 2
                .Caption = "Your units shared with the Guild"
            End With
            Me.AddChild(CType(fraSharedAssetsList, UIControl))

            'tvwMain initial props
            tvwMain = New UITreeView(oUILib)
            With tvwMain
                .ControlName = "tvwMain"
                .Left = 10
                .Top = 10
                .Width = 246
                .Height = 472
                .Enabled = True
                .Visible = True
                .BorderColor = muSettings.InterfaceBorderColor
                .FillColor = muSettings.InterfaceFillColor
                .ForeColor = muSettings.InterfaceBorderColor
                .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
                .mbAcceptReprocessEvents = True
            End With
            fraSharedAssetsList.AddChild(CType(tvwMain, UIControl))

            'btnRefresh initial props
            btnRefresh = New UIButton(oUILib)
            With btnRefresh
                .ControlName = "btnRefresh"
                .Left = 10
                .Top = tvwMain.Top + tvwMain.Height + 5
                .Width = 100
                .Height = 24
                .Enabled = True
                .Visible = True
                .Caption = "Refresh"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = True
                .FontFormat = CType(5, DrawTextFormat)
                .ControlImageRect = New Rectangle(0, 0, 120, 32)
            End With
            fraSharedAssetsList.AddChild(CType(btnRefresh, UIControl))

            If goCurrentPlayer Is Nothing = True OrElse goCurrentPlayer.oGuild Is Nothing = True Then Exit Sub

            If moGuildShareAssets.GetUpperBound(0) = -1 Then
                RequestGuildSharedAssets()
            Else
                ReloadList()
            End If

            'Make temp fake list
            'Dim yData(10 * 52 + 4) As Byte
            'Dim lPos As Int32 = 0       'for msgcode
            'System.BitConverter.GetBytes(GlobalMessageCode.eAbandonColony).CopyTo(yData, lPos) : lPos += 2
            'System.BitConverter.GetBytes(10).CopyTo(yData, lPos) : lPos += 2
            'For x As Int32 = 0 To 9
            '    Dim iEnvirTypeID As Int16
            '    Dim lEnvirID As Int32
            '    Dim lOwnerID As Int32 = 3510
            '    Dim iObjTypeID As Int16 = ObjectType.eUnit
            '    Dim lObjectID As Int32 = -1
            '    Dim iLocX As Int32 = -1
            '    Dim iLocZ As Int32 = -1
            '    Dim iLocation As Int32 = CInt(3 * Rnd())
            '    Dim iThisLoc As Int32 = 0
            '    Dim bFound As Boolean = True
            '    For Y As Int32 = 0 To goGalaxy.mlSystemUB
            '        iThisLoc += 1
            '        If iThisLoc > iLocation Then
            '            iEnvirTypeID = ObjectType.eSolarSystem
            '            lEnvirID = goGalaxy.moSystems(Y).ObjectID
            '            Exit For
            '        End If
            '    Next
            '    Dim iWarpointUpkeep As Int64 = CInt(100000 * Rnd())
            '    Dim sName As String = "RandomUnit_" & x.ToString
            '    System.BitConverter.GetBytes(iEnvirTypeID).CopyTo(yData, lPos) : lPos += 2
            '    System.BitConverter.GetBytes(lEnvirID).CopyTo(yData, lPos) : lPos += 4
            '    System.BitConverter.GetBytes(lOwnerID).CopyTo(yData, lPos) : lPos += 4
            '    System.BitConverter.GetBytes(iObjTypeID).CopyTo(yData, lPos) : lPos += 2
            '    System.BitConverter.GetBytes(lObjectID).CopyTo(yData, lPos) : lPos += 4
            '    System.Text.ASCIIEncoding.ASCII.GetBytes(Mid$(sName, 1, 20)).CopyTo(yData, lPos) : lPos += 20
            '    System.BitConverter.GetBytes(iLocX).CopyTo(yData, lPos) : lPos += 4
            '    System.BitConverter.GetBytes(iLocZ).CopyTo(yData, lPos) : lPos += 4
            '    If x = 9 Then iWarpointUpkeep = Int64.MaxValue
            '    System.BitConverter.GetBytes(iWarpointUpkeep).CopyTo(yData, lPos) : lPos += 8
            'Next
            'HandleGetGuildShareAssets(yData)

            mbLoading = False

        End Sub

        Public Sub HandleGetGuildShareAssets(ByRef yData() As Byte)
            Dim lPos As Int32 = 2 'for msgcode
            Dim lUB As Int32 = System.BitConverter.ToInt16(yData, lPos) - 1 : lPos += 4
            Dim x As Int32 = 0
            ReDim moGuildShareAssets(lUB)
            Try

                For x = 0 To lUB
                    'If x = lUB Then Stop
                    With moGuildShareAssets(x)
                        '.lOwnerID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                        .iObjTypeID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
                        .lObjectID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                        Dim sName As String = GetCacheObjectValue(.lEnvirID, .iEnvirTypeID) 'Prefetch the cache
                        .iEnvirTypeID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
                        .lEnvirID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                        .iLocX = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                        .iLocZ = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                        .sName = GetStringFromBytes(yData, lPos, 20) : lPos += 20
                        '.lWarpointUpkeep = System.BitConverter.ToInt64(yData, lPos) : lPos += 8
                        .lIdx = -1
                    End With
                Next x
                ReloadList()
            Catch ex As Exception

            End Try
        End Sub

        Private Sub ReloadList()
            'Todo: Support Parent info for when it's a docked unit, or in the milky way.
            tvwMain.Clear()
            If moGuildShareAssets.GetUpperBound(0) = -1 Then Exit Sub
            For x As Int32 = 0 To moGuildShareAssets.GetUpperBound(0)
                Dim sName As String
                Select Case moGuildShareAssets(x).iEnvirTypeID
                    Case ObjectType.eSolarSystem
                        sName = GetCacheObjectValue(moGuildShareAssets(x).lEnvirID, ObjectType.eSolarSystem)
                    Case ObjectType.ePlanet
                        sName = GetCacheObjectValue(moGuildShareAssets(x).lEnvirID, ObjectType.ePlanet)
                    Case ObjectType.eGalaxy
                        sName = "Milky Way"
                        'Case ObjectType.eFacility, ObjectType.eUnit
                        '    sName = GetCacheObjectValue(moGuildShareAssets(x).iParentID, moGuildShareAssets(x).iParentTypeID)
                    Case Else
                        sName = "Unknown Location"
                End Select

                If sName.ToUpper = "UNKNOWN" Then mbHasUnknowns = True
                Dim oParentRoot As UITreeView.UITreeViewItem = Nothing
                oParentRoot = tvwMain.GetNodeByItemData3(moGuildShareAssets(x).lEnvirID, moGuildShareAssets(x).iEnvirTypeID, -1)
                If oParentRoot Is Nothing = True Then
                    Dim oCurr As UITreeView.UITreeViewItem = tvwMain.oRootNode
                    Dim oLast As UITreeView.UITreeViewItem = Nothing
                    Dim bFound As Boolean = False
                    While oCurr Is Nothing = False
                        If oCurr.lItemData3 = -1 Then
                            If sName.ToUpper > oCurr.sItem.Split(" "c)(0).ToUpper Then
                                oCurr = UITreeView.TraverseNextNode_NoExpand(oCurr)
                                Continue While
                            Else
                                oLast = oCurr
                                'If oLast Is Nothing = False Then oCurr = oLast
                                bFound = True
                                Exit While
                            End If
                        End If
                        oLast = oCurr
                        oCurr = UITreeView.TraverseNextNode_NoExpand(oCurr)
                    End While
                    If bFound = False OrElse oLast Is Nothing = True Then
                        oParentRoot = tvwMain.AddNode(sName, moGuildShareAssets(x).lEnvirID, moGuildShareAssets(x).iEnvirTypeID, -1, Nothing, Nothing)
                        oParentRoot.bItemBold = True
                    Else
                        oParentRoot = tvwMain.AddNode(sName, moGuildShareAssets(x).lEnvirID, moGuildShareAssets(x).iEnvirTypeID, -1, Nothing, oLast)
                        oParentRoot.bItemBold = True
                    End If
                    If sName <> "Unknown Location" Then
                        oParentRoot.sItem = oParentRoot.sItem & " (0)"
                    End If
                End If
                If moGuildShareAssets(x).iObjTypeID = -2 AndAlso moGuildShareAssets(x).lObjectID > 0 Then
                    moGuildShareAssets(x).sName = moGuildShareAssets(x).lObjectID.ToString & " additional units."
                Else
                    'If moGuildShareAssets(x).sName.Contains(" (") = False Then moGuildShareAssets(x).sName &= " (" & moGuildShareAssets(x).lWarpointUpkeep.ToString("#,##0") & ")"
                    'Dim sParent() As String
                    'Dim sParentWP As Int64 = 0
                    'oParentRoot.sItem = oParentRoot.sItem.Replace(")", "")
                    'sParent = oParentRoot.sItem.Split("("c)
                    'Try
                    '    sParentWP = CLng(sParent(1)) + moGuildShareAssets(x).lWarpointUpkeep
                    'Catch
                    'End Try
                    'oParentRoot.sItem = sName & " (" & sParentWP.ToString("#,##0") & ")"
                End If

                Dim bChildFound As Boolean = False
                Dim oChildCurr As UITreeView.UITreeViewItem = oParentRoot.oFirstChild
                Dim oChildLast As UITreeView.UITreeViewItem = Nothing
                While oChildCurr Is Nothing = False
                    If moGuildShareAssets(x).sName.ToUpper > oChildCurr.sItem.ToUpper Then
                        oChildCurr = oChildCurr.oNextSibling
                        Continue While
                    Else
                        oChildLast = oChildCurr
                        bChildFound = True
                        Exit While
                    End If
                    oChildLast = oChildCurr
                    oChildCurr = oChildCurr.oNextSibling
                End While
                If bChildFound = False OrElse oChildLast Is Nothing = True Then
                    tvwMain.AddNode(moGuildShareAssets(x).sName, moGuildShareAssets(x).lObjectID, moGuildShareAssets(x).iObjTypeID, ObjectType.eUnit, oParentRoot, Nothing)
                Else
                    tvwMain.AddNode(moGuildShareAssets(x).sName, moGuildShareAssets(x).lObjectID, moGuildShareAssets(x).iObjTypeID, ObjectType.eUnit, oParentRoot, oChildLast)
                End If
            Next x
        End Sub

        Private Sub tvwMain_NodeDoubleClicked() Handles tvwMain.NodeDoubleClicked
            If tvwMain.oSelectedNode Is Nothing = True OrElse tvwMain.oSelectedNode.lItemData3 = -1 OrElse moGuildShareAssets.GetUpperBound(0) = -1 Then Exit Sub

            Dim lID As Int32 = tvwMain.oSelectedNode.lItemData
            Dim iTypeID As Int16 = CShort(tvwMain.oSelectedNode.lItemData2)
            For x As Int32 = 0 To moGuildShareAssets.GetUpperBound(0)
                If moGuildShareAssets(x).iObjTypeID = iTypeID AndAlso moGuildShareAssets(x).lObjectID = lID Then
                    Dim utmp As PlayerComm.WPAttachment
                    With utmp
                        If moGuildShareAssets(x).iLocX = -1 OrElse moGuildShareAssets(x).iLocZ = -1 Then Exit Sub
                        .EnvirID = moGuildShareAssets(x).lEnvirID
                        .EnvirTypeID = moGuildShareAssets(x).iEnvirTypeID
                        .LocX = moGuildShareAssets(x).iLocX
                        .LocZ = moGuildShareAssets(x).iLocZ
                        .sWPName = ""
                        .AttachNumber = 0
                    End With
                    utmp.JumpToAttachment()
                End If
            Next x
        End Sub

        Public Overrides Sub NewFrame()
            'Lookup mbHasUnknown
            If mbHasUnknowns = True AndAlso glCurrentCycle - mlLastUnknownCheck > 300 Then
                mbHasUnknowns = False
                mlLastUnknownCheck = glCurrentCycle
                Dim oCurr As UITreeView.UITreeViewItem = tvwMain.oRootNode
                Dim bFound As Boolean = False
                While oCurr Is Nothing = False
                    If oCurr.lItemData3 = -1 Then
                        If oCurr.sItem.ToUpper.StartsWith("UNKNOWN") Then
                            Dim sName As String = GetCacheObjectValue(oCurr.lItemData, CShort(oCurr.lItemData2))
                            If oCurr.sItem.Split(" "c)(0) <> sName Then

                                Dim sParent() As String
                                Dim sParentWP As Int64 = 0
                                oCurr.sItem = oCurr.sItem.Replace(")", "")
                                sParent = oCurr.sItem.Split("("c)
                                Try
                                    sParentWP = CLng(sParent(1))
                                Catch
                                End Try

                                oCurr.sItem = sName & " (" & sParentWP.ToString("#,##0") & ")"
                                Me.IsDirty = True
                            End If
                            If sName.StartsWith("UNKNOWN") Then mbHasUnknowns = True
                        End If

                    End If
                    oCurr = UITreeView.TraverseNextNode_NoExpand(oCurr)
                End While
            End If

            If glCurrentCycle - miLastRefresh > 300 Then
                If btnRefresh.Enabled = False Then btnRefresh.Enabled = True
                If btnRefresh.Caption <> "Refresh" Then btnRefresh.Caption = "Refresh"
            End If

        End Sub

        Public Overrides Sub RenderEnd()
        End Sub

        Private Sub btnRefresh_Click(ByVal sName As String) Handles btnRefresh.Click
            RequestGuildSharedAssets()
        End Sub

        Private Sub RequestGuildSharedAssets()
            If glCurrentCycle - miLastRefresh > 300 Then
                If btnRefresh.Enabled = True Then btnRefresh.Enabled = False
                If btnRefresh.Caption <> "Refreshing..." Then btnRefresh.Caption = "Refreshing..."
                miLastRefresh = glCurrentCycle
                tvwMain.Clear()
                Dim yMsg(1) As Byte
                System.BitConverter.GetBytes(GlobalMessageCode.eGetGuildShareAssets).CopyTo(yMsg, 0)
                MyBase.moUILib.SendMsgToPrimary(yMsg)
            End If
        End Sub
    End Class
End Class
