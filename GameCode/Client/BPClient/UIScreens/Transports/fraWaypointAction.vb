Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Interface created from Interface Builder
Partial Class frmTransportOrders
    Private Class fraWaypointAction
        Inherits UIWindow

        Private lblQuantity As UILabel
        Private txtQuantity As UITextBox
        Private chkAsPercent As UICheckBox
        Private tvwActionItems As UITreeView
        Private WithEvents btnAddAction As UIButton
        Private WithEvents optLoad As UIOption
        Private WithEvents optUnload As UIOption

        Public Event AddAction(ByVal yActionTypeID As Byte, ByVal lExt1 As Int32, ByVal iExt2 As Int16, ByVal lExt3 As Int32)

        Public Sub New(ByRef oUILib As UILib)
            MyBase.New(oUILib)

            'fraWaypointAction initial props
            With Me
                .ControlName = "fraWaypointAction"
                .Left = 405
                .Top = 174
                .Width = 225
                .Height = 220
                .Enabled = True
                .Visible = True
                .BorderColor = muSettings.InterfaceBorderColor
                .FillColor = muSettings.InterfaceFillColor
                .FullScreen = False
            End With

            'optLoad initial props
            optLoad = New UIOption(oUILib)
            With optLoad
                .ControlName = "optLoad"
                .Left = 10
                .Top = 10
                .Width = 50
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Load"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(6, DrawTextFormat)
                .Value = True
            End With
            Me.AddChild(CType(optLoad, UIControl))

            'optUnload initial props
            optUnload = New UIOption(oUILib)
            With optUnload
                .ControlName = "optUnload"
                .Left = 125
                .Top = 10
                .Width = 60
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Unload"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(6, DrawTextFormat)
                .Value = False
            End With
            Me.AddChild(CType(optUnload, UIControl))

            'lblQuantity initial props
            lblQuantity = New UILabel(oUILib)
            With lblQuantity
                .ControlName = "lblQuantity"
                .Left = 10
                .Top = 35
                .Width = 50
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Quantity:"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblQuantity, UIControl))

            'txtQuantity initial props
            txtQuantity = New UITextBox(oUILib)
            With txtQuantity
                .ControlName = "txtQuantity"
                .Left = 70
                .Top = 35
                .Width = 80
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "0"
                .ForeColor = muSettings.InterfaceTextBoxForeColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
                .BackColorEnabled = muSettings.InterfaceTextBoxFillColor
                .BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
                .MaxLength = 7
                .BorderColor = muSettings.InterfaceBorderColor
                .bNumericOnly = True
            End With
            Me.AddChild(CType(txtQuantity, UIControl))

            'chkAsPercent initial props
            chkAsPercent = New UICheckBox(oUILib)
            With chkAsPercent
                .ControlName = "chkAsPercent"
                .Left = 165
                .Top = 35
                .Width = 50
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "As %"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(6, DrawTextFormat)
                .Value = False
            End With
            Me.AddChild(CType(chkAsPercent, UIControl))

            'tvwActionItems initial props
            tvwActionItems = New UITreeView(oUILib)
            With tvwActionItems
                .ControlName = "tvwActionItems"
                .Left = 10
                .Top = 60
                .Width = Me.Width - 20
                .Height = 125
                .Enabled = True
                .Visible = True
                .BorderColor = muSettings.InterfaceBorderColor
                .FillColor = muSettings.InterfaceFillColor
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
            End With
            Me.AddChild(CType(tvwActionItems, UIControl))

            'btnAddAction initial props
            btnAddAction = New UIButton(oUILib)
            With btnAddAction
                .ControlName = "btnAddAction"
                .Left = Me.Width - 110
                .Top = 190
                .Width = 100
                .Height = 24
                .Enabled = True
                .Visible = True
                .Caption = "Add Action"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = True
                .FontFormat = CType(5, DrawTextFormat)
                .ControlImageRect = New Rectangle(0, 0, 120, 32)
            End With
            Me.AddChild(CType(btnAddAction, UIControl))

            FillActionItems()
        End Sub

        Private Sub FillActionItems()
            tvwActionItems.Clear()

            'components and minerals

            tvwActionItems.AddNode("Any/All Cargo", -1, -1, -1, Nothing, Nothing)
            tvwActionItems.AddNode("Any/All Alloys and Minerals", -1, ObjectType.eMineral, -1, Nothing, Nothing)

            Dim oAlloys As UITreeView.UITreeViewItem = tvwActionItems.AddNode("Alloys", Int32.MinValue, Int32.MinValue, Int32.MinValue, Nothing, Nothing)
            Dim oComponents As UITreeView.UITreeViewItem = tvwActionItems.AddNode("Components", Int32.MinValue, Int32.MinValue, Int32.MinValue, Nothing, Nothing)
            Dim oMinerals As UITreeView.UITreeViewItem = tvwActionItems.AddNode("Minerals", Int32.MinValue, Int32.MinValue, Int32.MinValue, Nothing, Nothing)
            Dim oPersonnel As UITreeView.UITreeViewItem = tvwActionItems.AddNode("Personnel", Int32.MinValue, Int32.MaxValue, Int32.MinValue, Nothing, Nothing)

            oMinerals.bItemBold = True : oMinerals.bUseFillColor = True
            oAlloys.bItemBold = True : oAlloys.bUseFillColor = True
            oComponents.bItemBold = True : oComponents.bUseFillColor = True
            oPersonnel.bItemBold = True : oPersonnel.bUseFillColor = True

            tvwActionItems.AddNode("Colonists", 0, ObjectType.eColonists, -1, oPersonnel, Nothing)
            tvwActionItems.AddNode("Enlisted", 0, ObjectType.eEnlisted, -1, oPersonnel, Nothing)
            tvwActionItems.AddNode("Officers", 0, ObjectType.eOfficers, -1, oPersonnel, Nothing)

            '-1 needs to be all
            'tvwActionItems.AddNode("Any/All Minerals", -1, ObjectType.eMineral, -1, oMinerals, Nothing)
            tvwActionItems.AddNode("Any/All Components", -1, ObjectType.eComponentCache, -1, oComponents, Nothing)

            'All of a type
            Dim lMinIdx() As Int32 = GetSortedMineralIdxArray(True)
            If lMinIdx Is Nothing = False Then
                For X As Int32 = 0 To lMinIdx.GetUpperBound(0)
                    Dim lTempMin As Int32 = lMinIdx(X)
                    If lTempMin < 0 OrElse lTempMin > glMineralUB Then Continue For
                    Dim oMin As Mineral = goMinerals(lTempMin)
                    If oMin Is Nothing = False Then
                        If goMinerals(lMinIdx(X)).ObjectID <= 157 OrElse goMinerals(lMinIdx(X)).ObjectID = 41991 Then
                            tvwActionItems.AddNode(oMin.MineralName, oMin.ObjectID, ObjectType.eMineral, -1, oMinerals, Nothing)
                        Else
                            tvwActionItems.AddNode(oMin.MineralName, oMin.ObjectID, ObjectType.eMineral, -1, oAlloys, Nothing)
                        End If
                    End If
                Next X
            End If


            Dim oArmor As UITreeView.UITreeViewItem = tvwActionItems.AddNode("Armor Components", Int32.MinValue, Int32.MinValue, Int32.MinValue, oComponents, Nothing)
            Dim oEngine As UITreeView.UITreeViewItem = tvwActionItems.AddNode("Engine Components", Int32.MinValue, Int32.MinValue, Int32.MinValue, oComponents, Nothing)
            Dim oRadar As UITreeView.UITreeViewItem = tvwActionItems.AddNode("Radar Components", Int32.MinValue, Int32.MinValue, Int32.MinValue, oComponents, Nothing)
            Dim oShield As UITreeView.UITreeViewItem = tvwActionItems.AddNode("Shield Components", Int32.MinValue, Int32.MinValue, Int32.MinValue, oComponents, Nothing)
            Dim oWeapon As UITreeView.UITreeViewItem = tvwActionItems.AddNode("Weapon Components", Int32.MinValue, Int32.MinValue, Int32.MinValue, oComponents, Nothing)

            tvwActionItems.AddNode("Any/All Armor", -1, ObjectType.eArmorTech, -1, oArmor, Nothing)
            tvwActionItems.AddNode("Any/All Engines", -1, ObjectType.eEngineTech, -1, oEngine, Nothing)
            tvwActionItems.AddNode("Any/All Radars", -1, ObjectType.eRadarTech, -1, oRadar, Nothing)
            tvwActionItems.AddNode("Any/All Shields", -1, ObjectType.eShieldTech, -1, oShield, Nothing)
            tvwActionItems.AddNode("Any/All Weapons", -1, ObjectType.eWeaponTech, -1, oWeapon, Nothing)

            For X As Int32 = 0 To goCurrentPlayer.mlTechUB
                Dim oTech As Base_Tech = goCurrentPlayer.moTechs(X)
                If oTech Is Nothing = False Then
                    Select Case oTech.ObjTypeID
                        Case ObjectType.eArmorTech
                            tvwActionItems.AddNode(oTech.GetComponentName(), oTech.ObjectID, oTech.ObjTypeID, -1, oArmor, Nothing)
                        Case ObjectType.eEngineTech
                            tvwActionItems.AddNode(oTech.GetComponentName(), oTech.ObjectID, oTech.ObjTypeID, -1, oEngine, Nothing)
                        Case ObjectType.eRadarTech
                            tvwActionItems.AddNode(oTech.GetComponentName(), oTech.ObjectID, oTech.ObjTypeID, -1, oRadar, Nothing)
                        Case ObjectType.eShieldTech
                            tvwActionItems.AddNode(oTech.GetComponentName(), oTech.ObjectID, oTech.ObjTypeID, -1, oShield, Nothing)
                        Case ObjectType.eWeaponTech
                            tvwActionItems.AddNode(oTech.GetComponentName(), oTech.ObjectID, oTech.ObjTypeID, -1, oWeapon, Nothing)
                    End Select
                End If
            Next X

        End Sub

        Private Sub btnAddAction_Click(ByVal sName As String) Handles btnAddAction.Click
            Dim oNode As UITreeView.UITreeViewItem = tvwActionItems.oSelectedNode
            If oNode Is Nothing Then
                goUILib.AddNotification("Select an item to take action with.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return
            End If

            If oNode.lItemData = Int32.MinValue OrElse oNode.lItemData2 = Int32.MinValue Then
                goUILib.AddNotification("Select an item to take action with.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return
            End If

            If txtQuantity.Caption = "" Then
                goUILib.AddNotification("Enter a quantity.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return
            End If
            If oNode.lItemData2 = ObjectType.eColonists Then
                goUILib.AddNotification("Warning: Picking up colonists from a colony may result in destruction if all colonists are removed from the colony.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            End If

            'OK - create our action code
            Dim yAction As Byte = 0
            If optLoad.Value = True Then
                yAction = TransportRouteAction.TransportRouteActionType.eLoad
            Else
                yAction = TransportRouteAction.TransportRouteActionType.eUnload
            End If
            If chkAsPercent.Value = True Then
                yAction = yAction Or TransportRouteAction.TransportRouteActionType.ePercentage
            Else
                yAction = yAction Or TransportRouteAction.TransportRouteActionType.eQty
            End If

            'Ok, define our Extendeds
            Dim lExt1 As Int32 = oNode.lItemData
            Dim iExt2 As Int16 = CShort(oNode.lItemData2)
            Dim lExt3 As Int32 = CInt(Val(txtQuantity.Caption))

            RaiseEvent AddAction(yAction, lExt1, iExt2, lExt3)

        End Sub

        Private Sub optUnload_Click() Handles optUnload.Click
            optUnload.Value = True
            optLoad.Value = False
        End Sub

        Private Sub optLoad_Click() Handles optLoad.Click
            optUnload.Value = False
            optLoad.Value = True
        End Sub
    End Class
End Class
