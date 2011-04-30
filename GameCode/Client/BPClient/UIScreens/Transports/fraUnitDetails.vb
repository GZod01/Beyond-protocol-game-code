Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Interface created from Interface Builder
Partial Class frmTransportManagement
    Private Class fraUnitDetails
        Inherits UIWindow

        Private mlTransportID As Int32 = -1

        Private WithEvents txtName As UITextBox
        Private lblStatus As UILabel
        Private lblSpeed As UILabel
        Private lblManeuver As UILabel

        Private lblCargo As UILabel
        Private lstCargo As UIListBox
        Private WithEvents btnOrders As UIButton
        Private WithEvents btnDiscard As UIButton
        Private WithEvents btnBegin As UIButton
        Private WithEvents btnPause As UIButton
        Private WithEvents btnRecall As UIButton

        Private moModelViewTex As Texture
        Private mlModelViewTexID As Int32 = -1
        Private moShader As ModelShader = Nothing

        Public Sub New(ByRef oUILib As UILib)
            MyBase.New(oUILib)

            'fraUnitDetails initial props
            With Me
                .ControlName = "fraUnitDetails"
                .Left = 360
                .Top = 123
                .Width = 250
                .Height = 470
                .Enabled = True
                .Visible = True
                .BorderColor = muSettings.InterfaceBorderColor
                .FillColor = muSettings.InterfaceFillColor
                .FullScreen = False
                .Caption = "Unit Details"
            End With

            'lblName initial props
            txtName = New UITextBox(oUILib)
            With txtName
                .ControlName = "txtName"
                .Left = 7
                .Top = 10
                .Width = 240
                .Height = 14
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
            Me.AddChild(CType(txtName, UIControl))

            'lblStatus initial props
            lblStatus = New UILabel(oUILib)
            With lblStatus
                .ControlName = "lblStatus"
                .Left = 7
                .Top = 25
                .Width = 240
                .Height = 14
                .Enabled = True
                .Visible = True
                .Caption = "Status: Moving to Destination"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblStatus, UIControl))

            'lblSpeed initial props
            lblSpeed = New UILabel(oUILib)
            With lblSpeed
                .ControlName = "lblSpeed"
                .Left = 7
                .Top = 40
                .Width = 100
                .Height = 14
                .Enabled = True
                .Visible = True
                .Caption = "Speed: 255"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblSpeed, UIControl))

            'lblManeuver initial props
            lblManeuver = New UILabel(oUILib)
            With lblManeuver
                .ControlName = "lblManeuver"
                .Left = 150
                .Top = 40
                .Width = 100
                .Height = 14
                .Enabled = True
                .Visible = True
                .Caption = "Maneuver: 255"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblManeuver, UIControl))

            'lblCargo initial props
            lblCargo = New UILabel(oUILib)
            With lblCargo
                .ControlName = "lblCargo"
                .Left = 5
                .Top = 315
                .Width = 240
                .Height = 14
                .Enabled = True
                .Visible = True
                .Caption = "Cargo Capacity: 1,000,000 / 1,000,000"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblCargo, UIControl))


            'btnRecall initial props
            btnRecall = New UIButton(oUILib)
            With btnRecall
                .ControlName = "btnRecall"
                .Left = 5
                .Top = 290
                .Width = 100
                .Height = 24
                .Enabled = True
                .Visible = True
                .Caption = "Recall"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = True
                .FontFormat = CType(5, DrawTextFormat)
                .ControlImageRect = New Rectangle(0, 0, 120, 32)
            End With
            Me.AddChild(CType(btnRecall, UIControl))

            'btnPause initial props
            btnPause = New UIButton(oUILib)
            With btnPause
                .ControlName = "btnPause"
                .Left = 145
                .Top = 260
                .Width = 100
                .Height = 24
                .Enabled = True
                .Visible = True
                .Caption = "Pause"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = True
                .FontFormat = CType(5, DrawTextFormat)
                .ControlImageRect = New Rectangle(0, 0, 120, 32)
            End With
            Me.AddChild(CType(btnPause, UIControl))

            'btnBegin initial props
            btnBegin = New UIButton(oUILib)
            With btnBegin
                .ControlName = "btnBegin"
                .Left = 5
                .Top = 260
                .Width = 100
                .Height = 24
                .Enabled = True
                .Visible = True
                .Caption = "Begin Route"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = True
                .FontFormat = CType(5, DrawTextFormat)
                .ControlImageRect = New Rectangle(0, 0, 120, 32)
            End With
            Me.AddChild(CType(btnBegin, UIControl))

            'btnDiscard initial props
            btnDiscard = New UIButton(oUILib)
            With btnDiscard
                .ControlName = "btnDiscard"
                .Left = 80
                .Top = 438
                .Width = 100
                .Height = 24
                .Enabled = True
                .Visible = True
                .Caption = "Discard"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = True
                .FontFormat = CType(5, DrawTextFormat)
                .ControlImageRect = New Rectangle(0, 0, 120, 32)
            End With
            Me.AddChild(CType(btnDiscard, UIControl))

            'btnOrders initial props
            btnOrders = New UIButton(oUILib)
            With btnOrders
                .ControlName = "btnOrders"
                .Left = 145
                .Top = 290
                .Width = 100
                .Height = 24
                .Enabled = True
                .Visible = True
                .Caption = "Orders"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = True
                .FontFormat = CType(5, DrawTextFormat)
                .ControlImageRect = New Rectangle(0, 0, 120, 32)
            End With
            Me.AddChild(CType(btnOrders, UIControl))

            'lstCargo initial props
            lstCargo = New UIListBox(oUILib)
            With lstCargo
                .ControlName = "lstCargo"
                .Left = 5
                .Top = 330
                .Width = 240
                .Height = 100
                .Enabled = True
                .Visible = True
                .BorderColor = muSettings.InterfaceBorderColor
                .FillColor = muSettings.InterfaceFillColor
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Courier New", 9.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
            End With
            Me.AddChild(CType(lstCargo, UIControl))
        End Sub

        Public Sub SetCurrentTransport(ByVal lID As Int32)
            mlTransportID = lID
        End Sub

        Private Function HasAlterRights() As Boolean
            If gbAliased = False Then Return True
            If (glAliasRights And AliasingRights.eModifyBattleGroups) = 0 Then
                goUILib.AddNotification("You lack the rights to alter transports.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                goUILib.AddNotification("Request alias rights to Modify Battle Groups.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return False
            End If
            Return True
        End Function

        Private Sub btnBegin_Click(ByVal sName As String) Handles btnBegin.Click
            If HasAlterRights() = False Then Return
            Dim oTrans As Transport = frmTransportManagement.GetTransport(mlTransportID)
            If oTrans Is Nothing = False Then
                If (oTrans.TransFlags And Transport.elTransportFlags.eIdle) = 0 Then
                    goUILib.AddNotification("The transport must be Idle in order to Begin a route.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                    Return
                End If

                Dim yMsg(6) As Byte
                Dim lPos As Int32 = 0
                System.BitConverter.GetBytes(GlobalMessageCode.eSetTransportStatus).CopyTo(yMsg, lPos) : lPos += 2
                System.BitConverter.GetBytes(mlTransportID).CopyTo(yMsg, lPos) : lPos += 4
                yMsg(lPos) = frmTransportOrders.eyStatusCode.eBeginRoute : lPos += 1

                goUILib.SendMsgToPrimary(yMsg)
            End If
        End Sub

        Private Sub btnDiscard_Click(ByVal sName As String) Handles btnDiscard.Click
            If HasAlterRights() = False Then Return
            Dim oTrans As Transport = frmTransportManagement.GetTransport(mlTransportID)
            If oTrans Is Nothing = False Then
                If (oTrans.TransFlags And Transport.elTransportFlags.eIdle) = 0 Then
                    goUILib.AddNotification("The transport must be Idle in order to Discard cargo.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                    Return
                End If

                If lstCargo.ListIndex = -1 Then
                    goUILib.AddNotification("Select an item of cargo to discard.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                    Return
                End If

                Dim yMsg(16) As Byte
                Dim lPos As Int32 = 0
                System.BitConverter.GetBytes(GlobalMessageCode.eSetTransportStatus).CopyTo(yMsg, lPos) : lPos += 2
                System.BitConverter.GetBytes(mlTransportID).CopyTo(yMsg, lPos) : lPos += 4
                yMsg(lPos) = frmTransportOrders.eyStatusCode.eDiscardCargo : lPos += 1
                System.BitConverter.GetBytes(lstCargo.ItemData(lstCargo.ListIndex)).CopyTo(yMsg, lPos) : lPos += 4
                System.BitConverter.GetBytes(CShort(lstCargo.ItemData2(lstCargo.ListIndex))).CopyTo(yMsg, lPos) : lPos += 2
                System.BitConverter.GetBytes(lstCargo.ItemData3(lstCargo.ListIndex)).CopyTo(yMsg, lPos) : lPos += 4

                goUILib.SendMsgToPrimary(yMsg)
            End If
        End Sub

        Private Sub btnOrders_Click(ByVal sName As String) Handles btnOrders.Click
            If HasAlterRights() = False Then Return
            Dim oTrans As Transport = frmTransportManagement.GetTransport(mlTransportID)
            If oTrans Is Nothing Then Return
            Dim ofrm As New frmTransportOrders(goUILib)
            ofrm.Visible = True
            ofrm.SetTransport(mlTransportID)
        End Sub

        Private Sub btnPause_Click(ByVal sName As String) Handles btnPause.Click
            If HasAlterRights() = False Then Return
            Dim oTrans As Transport = frmTransportManagement.GetTransport(mlTransportID)
            If oTrans Is Nothing = False Then
                Dim yMsg(6) As Byte
                Dim lPos As Int32 = 0
                System.BitConverter.GetBytes(GlobalMessageCode.eSetTransportStatus).CopyTo(yMsg, lPos) : lPos += 2
                System.BitConverter.GetBytes(mlTransportID).CopyTo(yMsg, lPos) : lPos += 4
                yMsg(lPos) = frmTransportOrders.eyStatusCode.ePauseRoute : lPos += 1

                goUILib.SendMsgToPrimary(yMsg)
            End If
        End Sub

        Private Sub btnRecall_Click(ByVal sName As String) Handles btnRecall.Click
            If HasAlterRights() = False Then Return
            Dim oTrans As Transport = frmTransportManagement.GetTransport(mlTransportID)
            If oTrans Is Nothing = False Then
                Dim yMsg(6) As Byte
                Dim lPos As Int32 = 0
                System.BitConverter.GetBytes(GlobalMessageCode.eSetTransportStatus).CopyTo(yMsg, lPos) : lPos += 2
                System.BitConverter.GetBytes(mlTransportID).CopyTo(yMsg, lPos) : lPos += 4
                yMsg(lPos) = frmTransportOrders.eyStatusCode.eRecall : lPos += 1

                goUILib.SendMsgToPrimary(yMsg)
            End If
        End Sub
        Private mbNameChangeInProgress As Boolean = False
        Public Sub fraUnitDetails_OnNewFrame() Handles Me.OnNewFrame
            Dim oTransport As Transport = frmTransportManagement.GetTransport(mlTransportID)
            If oTransport Is Nothing = False Then
                With oTransport

                    If (.TransFlags And Transport.elTransportFlags.ePaused) <> 0 Then
                        If btnPause.Caption <> "Unpause" Then btnPause.Caption = "Unpause"
                    ElseIf btnPause.Caption <> "Pause" Then
                        btnPause.Caption = "Pause"
                    End If

                    .RequestDetails()
                    Dim sText As String = frmTransportManagement.GetTransportName(.TransportID)
                    If mbNameChangeInProgress = False AndAlso txtName.Caption <> sText Then txtName.Caption = sText

                    sText = .GetStatusText()
                    If lblStatus.Caption <> sText Then lblStatus.Caption = sText
                    sText = "Speed: " & .MaxSpeed.ToString()
                    If lblSpeed.Caption <> sText Then lblSpeed.Caption = sText
                    sText = "Maneuver: " & .Maneuver.ToString
                    If lblManeuver.Caption <> sText Then lblManeuver.Caption = sText
                    sText = "Cargo Capacity: " & .CargoCapAvail.ToString("#,##0") & " / " & .CargoCapTotal.ToString("#,##0")
                    If lblCargo.Caption <> sText Then lblCargo.Caption = sText

                    If .lCargoUB + 1 <> lstCargo.ListCount Then
                        lstCargo.Clear()
                        Try
                            For X As Int32 = 0 To .lCargoUB
                                lstCargo.AddItem(.oCargo(X).GetDisplay())
                                lstCargo.ItemData(lstCargo.NewIndex) = .oCargo(X).CargoID
                                lstCargo.ItemData2(lstCargo.NewIndex) = .oCargo(X).CargoTypeID
                                lstCargo.ItemData3(lstCargo.NewIndex) = .oCargo(X).OwnerID
                            Next X
                        Catch
                        End Try
                    Else
                        Try
                            For X As Int32 = 0 To .lCargoUB

                                If lstCargo.ItemData(X) <> .oCargo(X).CargoID OrElse lstCargo.ItemData2(X) <> .oCargo(X).CargoTypeID OrElse lstCargo.ItemData3(X) <> .oCargo(X).OwnerID Then
                                    lstCargo.Clear()
                                    Exit For
                                End If

                                sText = .oCargo(X).GetDisplay()
                                If lstCargo.List(X) <> sText Then
                                    lstCargo.List(X) = sText
                                    lstCargo.IsDirty = True
                                End If
                            Next X
                        Catch
                            lstCargo.Clear()
                        End Try
                    End If
                End With

                If oTransport.ModelID <> mlModelViewTexID Then
                    RenderModelView(oTransport.ModelID)
                    mlModelViewTexID = oTransport.ModelID
                    Me.IsDirty = True
                End If
            Else
                If lstCargo.ListCount > 0 Then lstCargo.Clear()
                If txtName.Caption <> "" Then txtName.Caption = ""
                If lblStatus.Caption <> "" Then lblStatus.Caption = ""
                If lblSpeed.Caption <> "" Then lblSpeed.Caption = ""
                If lblManeuver.Caption <> "" Then lblManeuver.Caption = ""
                If lblCargo.Caption <> "Cargo Capacity: 0/0" Then lblCargo.Caption = "Cargo Capacity: 0/0"
            End If

        End Sub

        Private Sub RenderModelView(ByVal iModelID As Int16)
            'Ok, gotta render the model... get our model id
            Dim oOriginal As Surface
            Dim oScene As Surface = Nothing
            Dim matView As Matrix
            Dim matProj As Matrix

            Dim lCX As Int32
            Dim lCY As Int32
            Dim lCZ As Int32
            Dim lCAtX As Int32
            Dim lCAtY As Int32
            Dim lCAtZ As Int32

            'Get our current camera location
            With goCamera
                lCX = .mlCameraX
                lCY = .mlCameraY
                lCZ = .mlCameraZ
                lCAtX = .mlCameraAtX
                lCAtY = .mlCameraAtY
                lCAtZ = .mlCameraAtZ
            End With

            'Set no events
            Device.IsUsingEventHandlers = False

            'Create our texture...
            If moModelViewTex Is Nothing Then moModelViewTex = New Texture(MyBase.moUILib.oDevice, 512, 512, 1, Usage.RenderTarget, Format.R5G6B5, Pool.Default)
            With MyBase.moUILib.oDevice

                'Store our matrices beforehand...
                matView = MyBase.moUILib.oDevice.Transform.View
                matProj = MyBase.moUILib.oDevice.Transform.Projection

                'Ok, store our original surface
                oOriginal = .GetRenderTarget(0)

                'Get our surface to render to
                oScene = moModelViewTex.GetSurfaceLevel(0)

                'Now, set our render target to the texture's surface
                .SetRenderTarget(0, oScene)

                '.RenderState.ZBufferEnable = False
                'Clear out our surface
                .Clear(ClearFlags.Target Or ClearFlags.ZBuffer, System.Drawing.Color.FromArgb(255, 0, 0, 0), 1.0F, 0)

                'Set up our new matrices
                goCamera.SetupMatrices(MyBase.moUILib.oDevice, glCurrentEnvirView)

                .RenderState.AlphaBlendEnable = False

                'render our model here... NOTE: No breaking out of the code here... the render targets need to be reset!
                If iModelID <> 0 Then
                    Dim oObj As BaseMesh = goResMgr.GetMesh(iModelID)

                    Dim fMaxSize As Single = Math.Max(Math.Max(Math.Max(oObj.vecDeathSeqSize.X, oObj.vecDeathSeqSize.Y), oObj.vecDeathSeqSize.Z), 100)
                    With goCamera
                        .mlCameraX = CInt(fMaxSize / 2)
                        .mlCameraY = .mlCameraX
                        .mlCameraZ = -CInt(fMaxSize)

                        .mlCameraAtX = 0
                        .mlCameraAtY = -.mlCameraX \ 4
                        .mlCameraAtZ = 0
                    End With
                    'Set up our new matrices
                    goCamera.SetupMatrices(MyBase.moUILib.oDevice, glCurrentEnvirView)

                    If oObj.oMesh Is Nothing = False Then
                        .Transform.World = Matrix.Identity
                        .RenderState.CullMode = Cull.None

                        If muSettings.LightQuality > EngineSettings.LightQualitySetting.VSPS1 Then
                            If moShader Is Nothing = True Then moShader = New ModelShader()
                            moShader.RenderHullBuilderMesh(oObj, .Transform.World, iModelID)
                        Else
                            For X As Int32 = 0 To oObj.NumOfMaterials - 1
                                .Material = oObj.Materials(X)
                                .SetTexture(0, oObj.Textures(0))
                                oObj.oMesh.DrawSubset(X)
                            Next X

                            'Now, check for a turret...
                            If oObj.bTurretMesh = True Then
                                .Material = oObj.Materials(0)
                                .SetTexture(0, oObj.Textures(0))
                                oObj.oTurretMesh.DrawSubset(0)
                            End If
                        End If

                    End If
                    'End of model rendering...
                    .RenderState.CullMode = Cull.CounterClockwise
                    oObj = Nothing      'do not dispose, simply release the pointer
                End If

                'Now, restore our original surface to the device
                .SetRenderTarget(0, oOriginal)

                'now, re-enable our Z Buffer
                '.RenderState.ZBufferWriteEnable = True
                .RenderState.AlphaBlendEnable = True

                'restore our matrices
                .Transform.View = matView
                .Transform.Projection = matProj
                .Transform.World = Matrix.Identity

                'Release all our objects
                If oScene Is Nothing = False Then oScene.Dispose()
                If oOriginal Is Nothing = False Then oOriginal.Dispose()
                oScene = Nothing
                oOriginal = Nothing

            End With

            'turn events back on
            Device.IsUsingEventHandlers = True

            'Reset our camera
            With goCamera
                .mlCameraAtX = lCAtX
                .mlCameraAtY = lCAtY
                .mlCameraAtZ = lCAtZ
                .mlCameraX = lCX
                .mlCameraY = lCY
                .mlCameraZ = lCZ
            End With
        End Sub

        Public Sub fraUnitDetails_OnRenderEnd() Handles Me.OnRenderEnd

            If moModelViewTex Is Nothing = False AndAlso mlModelViewTexID > -1 Then
                Dim oLoc As System.Drawing.Point = Me.GetAbsolutePosition()
                Dim lInternalY As Int32 = lblSpeed.Top + lblSpeed.Height + 5
                oLoc.Y += lInternalY
                oLoc.X += lblSpeed.Left

                Dim lIdealWidth As Int32 = Me.Width - (lblSpeed.Left * 2)
                Dim lIdealHeight As Int32 = btnBegin.Top - lInternalY
                Dim lWH As Int32 = Math.Min(lIdealWidth, lIdealHeight)
                If lWH <> lIdealWidth Then
                    oLoc.X += (lIdealWidth \ 2) - (lWH \ 2)
                End If
                If lWH <> lIdealHeight Then
                    oLoc.Y += (lIdealHeight - lWH) \ 2  '(lIdealHeight \ 2) - (lWH \ 2)
                End If

                Dim rcDest As Rectangle = New Rectangle(oLoc.X, oLoc.Y, lWH, lWH)

                BPSprite.Draw2DOnce(GFXEngine.moDevice, moModelViewTex, New Rectangle(0, 0, 512, 512), rcDest, Color.White, 512, 512)
            End If

        End Sub

        Private Sub fraUnitDetails_WindowClosed() Handles Me.WindowClosed
            On Error Resume Next
            If moModelViewTex Is Nothing = False Then moModelViewTex.Dispose()
            moModelViewTex = Nothing
            mlModelViewTexID = -1
            If moShader Is Nothing = False Then moShader.DisposeMe()
            moShader = Nothing
        End Sub

        Private Sub txtName_OnKeyDown(ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtName.OnKeyDown
            mbNameChangeInProgress = True
            If e.KeyCode = Keys.Enter AndAlso txtName.Locked = False Then
                Dim oTransport As Transport = frmTransportManagement.GetTransport(mlTransportID)
                If otransport Is Nothing = False Then
                    Try
                        'Submit a name change...
                        Dim yMsg(26) As Byte
                        Dim lPos As Int32 = 0
                        System.BitConverter.GetBytes(GlobalMessageCode.eSetTransportStatus).CopyTo(yMsg, lPos) : lPos += 2
                        System.BitConverter.GetBytes(mlTransportID).CopyTo(yMsg, lPos) : lPos += 4
                        yMsg(lPos) = frmTransportOrders.eyStatusCode.eRenameTransport : lPos += 1
                        System.Text.ASCIIEncoding.ASCII.GetBytes(Mid$(txtName.Caption, 1, 20)).CopyTo(yMsg, lPos) : lPos += 20
                        goUILib.SendMsgToPrimary(yMsg)

                        mbNameChangeInProgress = False
                    Catch
                    End Try
                End If
            End If
        End Sub
    End Class
End Class
