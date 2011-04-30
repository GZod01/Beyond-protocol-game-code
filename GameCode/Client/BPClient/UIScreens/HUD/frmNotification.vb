Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Notification Window
Public Class frmNotification
    Inherits UIWindow

    Private Declare Function timeGetTime Lib "winmm.dll" Alias "timeGetTime" () As Int32

    Private Structure ClickToItem
        Public lEnvirID As Int32
        Public iEnvirTypeID As Int16
        Public lEntityID As Int32
        Public iEntityTypeID As Int16

        'Not always used....
        Public lLocX As Int32
        Public lLocZ As Int32
    End Structure

    Private lblItem1 As UILabel
    Private lblItem2 As UILabel
    Private lblItem3 As UILabel
    Private lblItem4 As UILabel
    Private lblItem5 As UILabel

    Private mlElementTime() As Int32
    Private mlElementAlpha() As Int32
    Private msElements() As String
    Private mcElementColor() As System.Drawing.Color
    Private muElementClickTo() As ClickToItem

    Private mlCurrentClickToIndex As Int32 = -1

    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)

        'frmPlayerRels initial props
        With Me
            .lWindowMetricID = BPMetrics.eWindow.eNotification
            .ControlName = "frmNotification"
            .Left = CInt(oUILib.oDevice.PresentationParameters.BackBufferWidth / 2) - 230
            .Top = 0
            .Width = 460
            .Height = 78
            If .Left + .Width > oUILib.oDevice.PresentationParameters.BackBufferWidth Then .Left = oUILib.oDevice.PresentationParameters.BackBufferWidth - .Width
            If .Top + .Height > oUILib.oDevice.PresentationParameters.BackBufferHeight Then .Top = oUILib.oDevice.PresentationParameters.BackBufferHeight - .Height
            .Enabled = True
            .Visible = True
            .BorderColor = System.Drawing.Color.FromArgb(64, 255, 255, 255)
            .FillColor = System.Drawing.Color.FromArgb(64, 128, 128, 128)
            .FullScreen = True      'Always display the notification window
            .BorderLineWidth = 2
        End With
        Debug.Write(Me.ControlName & " Newed" & vbCrLf)

        'lblItem1 initial props
        lblItem1 = New UILabel(oUILib)
        With lblItem1
            .ControlName = "lblItem1"
            .Left = 5
            .Top = 5
            .Width = 450
            .Height = 12
            .Enabled = True
            .Visible = True
            .Caption = ""
            .ForeColor = System.Drawing.Color.FromArgb(255, 0, 255, 255)
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.VerticalCenter
            .bAcceptEvents = True
        End With
        Me.AddChild(CType(lblItem1, UIControl))

        'lblItem2 initial props
        lblItem2 = New UILabel(oUILib)
        With lblItem2
            .ControlName = "lblItem2"
            .Left = 5
            .Top = 19
            .Width = 450
            .Height = 12
            .Enabled = True
            .Visible = True
            .Caption = ""
            .ForeColor = System.Drawing.Color.FromArgb(255, 0, 255, 255)
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.VerticalCenter
            .bAcceptEvents = True
        End With
        Me.AddChild(CType(lblItem2, UIControl))

        'lblItem3 initial props
        lblItem3 = New UILabel(oUILib)
        With lblItem3
            .ControlName = "lblItem3"
            .Left = 5
            .Top = 33
            .Width = 450
            .Height = 12
            .Enabled = True
            .Visible = True
            .Caption = ""
            .ForeColor = System.Drawing.Color.FromArgb(255, 0, 255, 255)
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.VerticalCenter
            .bAcceptEvents = True
        End With
        Me.AddChild(CType(lblItem3, UIControl))

        'lblItem4 initial props
        lblItem4 = New UILabel(oUILib)
        With lblItem4
            .ControlName = "lblItem4"
            .Left = 5
            .Top = 47
            .Width = 450
            .Height = 12
            .Enabled = True
            .Visible = True
            .Caption = ""
            .ForeColor = System.Drawing.Color.FromArgb(255, 0, 255, 255)
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.VerticalCenter
            .bAcceptEvents = True
        End With
        Me.AddChild(CType(lblItem4, UIControl))

        'lblItem5 initial props
        lblItem5 = New UILabel(oUILib)
        With lblItem5
            .ControlName = "lblItem5"
            .Left = 5
            .Top = 61
            .Width = 450
            .Height = 12
            .Enabled = True
            .Visible = True
            .Caption = ""
            .ForeColor = System.Drawing.Color.FromArgb(255, 0, 255, 255)
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.VerticalCenter
            .bAcceptEvents = True
        End With
        Me.AddChild(CType(lblItem5, UIControl))

        'Ensure we are unique
        oUILib.RemoveWindow(Me.ControlName)

        Dim lIdx As Int32
        ReDim mlElementTime(4)
        ReDim mlElementAlpha(4)
        ReDim msElements(4)
        ReDim mcElementColor(4)
        ReDim muElementClickTo(4)
        For lIdx = 0 To mlElementTime.Length - 1
            mlElementTime(lIdx) = 0
            mlElementAlpha(lIdx) = 0
            msElements(lIdx) = ""
            muElementClickTo(lIdx).iEntityTypeID = -1
            muElementClickTo(lIdx).iEnvirTypeID = -1
            muElementClickTo(lIdx).lEntityID = -1
            muElementClickTo(lIdx).lEnvirID = -1
        Next lIdx

        AddHandler lblItem1.OnMouseDown, AddressOf Label1Clicked
        AddHandler lblItem2.OnMouseDown, AddressOf Label2Clicked
        AddHandler lblItem3.OnMouseDown, AddressOf Label3Clicked
        AddHandler lblItem4.OnMouseDown, AddressOf Label4Clicked
        AddHandler lblItem5.OnMouseDown, AddressOf Label5Clicked

        'Now, add me...
        oUILib.AddWindow(Me)
    End Sub

    Private Sub frmNotification_OnNewFrame() Handles MyBase.OnNewFrame
        'Ok, the window has already rendered...
        Dim X As Int32
        Dim Y As Int32
        Dim lCurrIdx As Int32
        Dim lCurrTime As Int32 = timeGetTime
        Dim yVis As Byte

        Dim oTmpWin As UIWindow = MyBase.moUILib.GetWindow("frmNoteHistory")
        If oTmpWin Is Nothing = False AndAlso oTmpWin.Visible = True Then
            Me.Visible = False
            Return
        End If

        If muSettings.NotificationDisplayTime <> -1 AndAlso ((goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.yPlayerPhase <> 0) OrElse glCurrentEnvirView = CurrentView.eArmorResearch OrElse glCurrentEnvirView = CurrentView.ePrototypeBuilder OrElse glCurrentEnvirView = CurrentView.eHullResearch) Then
            For X = 0 To mlElementTime.Length - 1
                If lCurrTime - mlElementTime(X) > muSettings.NotificationDisplayTime Then
                    mlElementAlpha(X) -= 10      '10?
                    If mlElementAlpha(X) < 0 Then
                        mlElementAlpha(X) = 0

                        'shift up...
                        lCurrIdx = X
                        For Y = X To mlElementAlpha.Length - 1
                            If mlElementAlpha(Y) <> 0 Then
                                mlElementAlpha(lCurrIdx) = mlElementAlpha(Y)
                                mlElementTime(lCurrIdx) = mlElementTime(Y)
                                msElements(lCurrIdx) = msElements(Y)
                                mcElementColor(lCurrIdx) = mcElementColor(Y)
                                muElementClickTo(lCurrIdx) = muElementClickTo(Y)

                                mlElementAlpha(Y) = 0

                                lCurrIdx += 1
                            End If
                        Next Y
                    End If
                End If
            Next X
        End If

        'Now, go through and do our stuff...
        yVis = 0
        For X = 0 To mlElementAlpha.Length - 1
            If mlElementAlpha(X) <> 0 Then yVis = 1 : Exit For
        Next X

        If yVis = 0 Then
            Me.Visible = False
        Else
            lCurrIdx = 0
            If lblItem1.Visible <> (mlElementAlpha(lCurrIdx) > 0) Then lblItem1.Visible = mlElementAlpha(lCurrIdx) > 0
            If lblItem1.Caption <> msElements(lCurrIdx) Then lblItem1.Caption = msElements(lCurrIdx)
            Dim clrVal As System.Drawing.Color = System.Drawing.Color.FromArgb(mlElementAlpha(lCurrIdx), mcElementColor(lCurrIdx).R, mcElementColor(lCurrIdx).G, mcElementColor(lCurrIdx).B)
            If lblItem1.ForeColor.ToArgb <> clrVal.ToArgb Then lblItem1.ForeColor = clrVal

            lCurrIdx = 1
            If lblItem2.Visible <> (mlElementAlpha(lCurrIdx) > 0) Then lblItem2.Visible = mlElementAlpha(lCurrIdx) > 0
            If lblItem2.Caption <> msElements(lCurrIdx) Then lblItem2.Caption = msElements(lCurrIdx)
            clrVal = System.Drawing.Color.FromArgb(mlElementAlpha(lCurrIdx), mcElementColor(lCurrIdx).R, mcElementColor(lCurrIdx).G, mcElementColor(lCurrIdx).B)
            If lblItem2.ForeColor.ToArgb <> clrVal.ToArgb Then lblItem2.ForeColor = clrVal

            lCurrIdx = 2
            If lblItem3.Visible <> (mlElementAlpha(lCurrIdx) > 0) Then lblItem3.Visible = mlElementAlpha(lCurrIdx) > 0
            If lblItem3.Caption <> msElements(lCurrIdx) Then lblItem3.Caption = msElements(lCurrIdx)
            clrVal = System.Drawing.Color.FromArgb(mlElementAlpha(lCurrIdx), mcElementColor(lCurrIdx).R, mcElementColor(lCurrIdx).G, mcElementColor(lCurrIdx).B)
            If lblItem3.ForeColor.ToArgb <> clrVal.ToArgb Then lblItem3.ForeColor = clrVal

            lCurrIdx = 3
            If lblItem4.Visible <> (mlElementAlpha(lCurrIdx) > 0) Then lblItem4.Visible = mlElementAlpha(lCurrIdx) > 0
            If lblItem4.Caption <> msElements(lCurrIdx) Then lblItem4.Caption = msElements(lCurrIdx)
            clrVal = System.Drawing.Color.FromArgb(mlElementAlpha(lCurrIdx), mcElementColor(lCurrIdx).R, mcElementColor(lCurrIdx).G, mcElementColor(lCurrIdx).B)
            If lblItem4.ForeColor.ToArgb <> clrVal.ToArgb Then lblItem4.ForeColor = clrVal

            lCurrIdx = 4
            If lblItem5.Visible <> (mlElementAlpha(lCurrIdx) > 0) Then lblItem5.Visible = mlElementAlpha(lCurrIdx) > 0
            If lblItem5.Caption <> msElements(lCurrIdx) Then lblItem5.Caption = msElements(lCurrIdx)
            clrVal = System.Drawing.Color.FromArgb(mlElementAlpha(lCurrIdx), mcElementColor(lCurrIdx).R, mcElementColor(lCurrIdx).G, mcElementColor(lCurrIdx).B)
            If lblItem5.ForeColor.ToArgb <> clrVal.ToArgb Then lblItem5.ForeColor = clrVal
        End If

    End Sub

    Private mbAddedToChat As Boolean = False
    Public Sub AddNotification(ByVal sText As String, ByVal cNoteColor As System.Drawing.Color, ByVal bForceDown As Boolean, ByVal lEnvirID As Int32, ByVal iEnvirTypeID As Int16, ByVal lEntityID As Int32, ByVal iEntityTypeID As Int16, ByVal lLocX As Int32, ByVal lLocZ As Int32)
        Dim X As Int32
        Dim Y As Int32
        Dim lCurrIdx As Int32
        Dim bFound As Boolean = False

        Try

            If mbAddedToChat = False Then
                Dim oChat As frmChat = CType(goUILib.GetWindow("frmChat"), frmChat)
                If oChat Is Nothing = False Then oChat.AddChatMessage(-1, sText, ChatMessageType.eNotificationMessage, Date.MinValue, True)
                oChat = Nothing
            End If
            mbAddedToChat = True

            If Me.Visible = False Then
                Dim ofrm As frmNoteHistory = CType(MyBase.moUILib.GetWindow("frmNoteHistory"), frmNoteHistory)
                If ofrm Is Nothing OrElse ofrm.Visible = False Then Me.Visible = True
            End If

            'Ok... split up our line if we need to...
            Dim lTempWidth As Int32 = lblItem1.GetTextDimensions(sText).Width
            If lTempWidth > lblItem1.Width AndAlso lTempWidth > 0 Then
                Dim fScaleFactor As Single = CSng(lblItem1.Width / lTempWidth)
                'Now... get our new len
                Dim lNewLen As Int32 = CInt(fScaleFactor * sText.Length)
                Dim lCharSpot As Int32
                If lNewLen <> 0 AndAlso lNewLen < sText.Length Then
                    lCharSpot = sText.LastIndexOf(" "c, lNewLen)
                    If lCharSpot = -1 Then lCharSpot = lNewLen

                    'Ok, now... we want to send the remainder to the new line...
                    AddNotification(sText.Substring(lCharSpot).Trim, cNoteColor, True, lEnvirID, iEnvirTypeID, lEntityID, iEntityTypeID, lLocX, lLocZ)
                    mbAddedToChat = True

                    'And we want to lop off that part of it
                    sText = sText.Substring(0, lCharSpot).Trim
                End If
            End If

            If bForceDown = False Then
                For X = 0 To mlElementAlpha.Length - 1
                    If mlElementAlpha(X) = 0 Then
                        mlElementAlpha(X) = 255
                        mlElementTime(X) = timeGetTime
                        msElements(X) = sText
                        mcElementColor(X) = cNoteColor
                        With muElementClickTo(X)
                            .lEnvirID = lEnvirID
                            .lEntityID = lEntityID
                            .iEnvirTypeID = iEnvirTypeID
                            .iEntityTypeID = iEntityTypeID
                            .lLocX = lLocX
                            .lLocZ = lLocZ
                        End With
                        bFound = True
                        Exit For
                    End If
                Next X
            End If

            If bFound = False Then
                'Ok, so, we need to shift all items up 1... element 0 is oldest
                mlElementAlpha(0) = 0
                lCurrIdx = 0

                'shift up...
                X = 0
                For Y = X To mlElementAlpha.Length - 1
                    If mlElementAlpha(Y) <> 0 Then
                        mlElementAlpha(lCurrIdx) = mlElementAlpha(Y)
                        mlElementTime(lCurrIdx) = mlElementTime(Y)
                        msElements(lCurrIdx) = msElements(Y)
                        mcElementColor(lCurrIdx) = mcElementColor(Y)
                        muElementClickTo(lCurrIdx) = muElementClickTo(Y)
                        lCurrIdx += 1
                    End If
                Next Y

                'Now set the last element
                X = mlElementAlpha.Length - 1
                mlElementAlpha(X) = 255
                mlElementTime(X) = timeGetTime
                msElements(X) = sText
                mcElementColor(X) = cNoteColor

                With muElementClickTo(X)
                    .lEnvirID = lEnvirID
                    .lEntityID = lEntityID
                    .iEnvirTypeID = iEnvirTypeID
                    .iEntityTypeID = iEntityTypeID
                    .lLocX = lLocX
                    .lLocZ = lLocZ
                End With
            End If

            frmNoteHistory.AddHistory(sText, cNoteColor, lEnvirID, iEnvirTypeID, lEntityID, iEntityTypeID, lLocX, lLocZ)

            If Me.Visible = True Then Me.IsDirty = True
            mbAddedToChat = False
        Catch
        End Try
    End Sub

    Protected Overrides Sub Finalize()
        Debug.Write(Me.ControlName & " Finalized" & vbCrLf)
        MyBase.Finalize()
    End Sub

    Private Sub Label1Clicked(ByVal lMouseX As Int32, ByVal lMouseY As Int32, ByVal lButton As MouseButtons)
        HandleClickToEvent(0)
    End Sub

    Private Sub Label2Clicked(ByVal lMouseX As Int32, ByVal lMouseY As Int32, ByVal lButton As MouseButtons)
        HandleClickToEvent(1)
    End Sub

    Private Sub Label3Clicked(ByVal lMouseX As Int32, ByVal lMouseY As Int32, ByVal lButton As MouseButtons)
        HandleClickToEvent(2)
    End Sub

    Private Sub Label4Clicked(ByVal lMouseX As Int32, ByVal lMouseY As Int32, ByVal lButton As MouseButtons)
        HandleClickToEvent(3)
    End Sub

    Private Sub Label5Clicked(ByVal lMouseX As Int32, ByVal lMouseY As Int32, ByVal lButton As MouseButtons)
        HandleClickToEvent(4)
    End Sub

    Private Sub HandleClickToEvent(ByVal lIdx As Int32)
        goCamera.TrackingIndex = -1
        mlCurrentClickToIndex = lIdx
        With muElementClickTo(lIdx)
            If .lEnvirID < 1 OrElse .iEnvirTypeID < 1 Then Return

            If .iEnvirTypeID = ObjectType.ePlanet OrElse .iEnvirTypeID = ObjectType.eSolarSystem Then
                '  are we in that environment? 
                If NewTutorialManager.TutorialOn = True Then
                    If MyBase.moUILib.CommandAllowed(True, Me.GetFullName) = False Then Return
                    NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eNotificationAlertClick, -1, -1, -1, "")
                End If

                If goCurrentEnvir Is Nothing = False Then
                    If goCurrentEnvir.ObjectID = .lEnvirID AndAlso goCurrentEnvir.ObjTypeID = .iEnvirTypeID Then
                        'Yes, we are... so no need to wait for switch
                        FinishClickToEvent()
                    Else
                        'no, change to that environment... but first set our select state
                        MyBase.moUILib.lUISelectState = UILib.eSelectState.eNotification_ClickTo
                        frmMain.ForceChangeEnvironment(.lEnvirID, .iEnvirTypeID)
                    End If
                End If
            End If
        End With
    End Sub

    Public Sub FinishClickToEvent()
        goCamera.TrackingIndex = -1
        If mlCurrentClickToIndex > -1 Then
            With muElementClickTo(mlCurrentClickToIndex)
                If goCurrentEnvir Is Nothing = False Then

                    If .iEnvirTypeID = ObjectType.ePlanet Then
                        glCurrentEnvirView = CurrentView.ePlanetMapView
                    Else : glCurrentEnvirView = CurrentView.eSystemMapView1
                    End If

                    With goCamera
                        .mlCameraAtX = 0 : .mlCameraAtY = 0 : .mlCameraAtZ = 0
                        .mlCameraX = 0 : .mlCameraY = 1000 : .mlCameraZ = -1000
                    End With

                    If .lLocX <> Int32.MinValue AndAlso .lLocZ <> Int32.MinValue Then
                        'Ok, use the locational coordinates
                        If .iEnvirTypeID = ObjectType.ePlanet Then
                            glCurrentEnvirView = CurrentView.ePlanetView
                            goCamera.mlCameraY = 1700
                        Else
                            glCurrentEnvirView = CurrentView.eSystemView
                            goCamera.mlCameraY = 3000
                        End If

                        goCamera.mlCameraAtX = .lLocX : goCamera.mlCameraAtY = 0 : goCamera.mlCameraAtZ = .lLocZ
                        goCamera.mlCameraX = goCamera.mlCameraAtX : goCamera.mlCameraZ = goCamera.mlCameraAtZ - 500
                        Try
                            If goCurrentEnvir.ObjTypeID = ObjectType.ePlanet Then
                                goCamera.mlCameraY = CInt(CType(goCurrentEnvir.oGeoObject, Planet).GetHeightAtPoint(goCamera.mlCameraAtX, goCamera.mlCameraAtZ, True)) + muSettings.lPlanetViewCameraY

                                goCamera.mlCameraX += muSettings.lPlanetViewCameraX
                                goCamera.mlCameraZ += muSettings.lPlanetViewCameraZ
                            End If
                        Catch
                        End Try
                    ElseIf .iEntityTypeID > -1 AndAlso .lEntityID > 0 Then
                        'Ok, now find that item...

                        For X As Int32 = 0 To goCurrentEnvir.lEntityUB
                            If goCurrentEnvir.lEntityIdx(X) = .lEntityID AndAlso goCurrentEnvir.oEntity(X).ObjTypeID = .iEntityTypeID Then

                                If .iEnvirTypeID = ObjectType.ePlanet Then
                                    glCurrentEnvirView = CurrentView.ePlanetView
                                    goCamera.mlCameraY = 1700
                                Else
                                    glCurrentEnvirView = CurrentView.eSystemView
                                    goCamera.mlCameraY = 3000
                                End If

                                With goCamera
                                    .mlCameraAtX = CInt(goCurrentEnvir.oEntity(X).LocX) : .mlCameraAtY = 0 : .mlCameraAtZ = CInt(goCurrentEnvir.oEntity(X).LocZ)
                                    .mlCameraX = .mlCameraAtX : .mlCameraZ = .mlCameraAtZ - 500

                                    Try
                                        If goCurrentEnvir.ObjTypeID = ObjectType.ePlanet Then
                                            goCamera.mlCameraY = CInt(CType(goCurrentEnvir.oGeoObject, Planet).GetHeightAtPoint(goCamera.mlCameraAtX, goCamera.mlCameraAtZ, True)) + muSettings.lPlanetViewCameraY

                                            goCamera.mlCameraX += muSettings.lPlanetViewCameraX
                                            goCamera.mlCameraZ += muSettings.lPlanetViewCameraZ
                                        End If
                                    Catch
                                    End Try
                                End With

                                'goCamera.TrackingIndex = X
                                Exit For
                            End If
                        Next X
                    End If
                End If

            End With
        End If
    End Sub

End Class

'Interface created from Interface Builder
Public Class frmNoteHistory
    Inherits UIWindow

    Private Structure ClickToItem
        Public lEnvirID As Int32
        Public iEnvirTypeID As Int16
        Public lEntityID As Int32
        Public iEntityTypeID As Int16

        'Not always used....
        Public lLocX As Int32
        Public lLocZ As Int32
    End Structure

    Private Shared msHistory() As String
    Private Shared mcHistoryColor() As System.Drawing.Color
    Private Shared muElementClickTo() As ClickToItem
    Private Shared mlHistoryUB As Int32 ' = -1
    Private Shared mbInSyncLock As Boolean = False

    Private moSysFont As System.Drawing.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25, FontStyle.Bold, GraphicsUnit.Point, 0)

    Private mlPreviousUB As Int32 = -1

    Private muCurrentClickItem As ClickToItem
    Private mbCurrentClickItemSet As Boolean = False

    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)

        'frmNoteHistory initial props
        With Me
            .lWindowMetricID = BPMetrics.eWindow.eNotificationHistory
            .ControlName = "frmNoteHistory"
            .Left = CInt(oUILib.oDevice.PresentationParameters.BackBufferWidth / 2) - 230
            .Top = 0
            .Width = 460
            .Height = 512
            .Enabled = True
            .Visible = True
            .BorderColor = System.Drawing.Color.FromArgb(64, 255, 255, 255)
            .FillColor = System.Drawing.Color.FromArgb(64, 128, 128, 128)
            .FullScreen = True
            .BorderLineWidth = 2
            .bRoundedBorder = True
        End With

        MyBase.moUILib.RemoveWindow(Me.ControlName)
        MyBase.moUILib.AddWindow(Me)
    End Sub

    Public Shared Sub AddHistory(ByVal sText As String, ByVal cColor As System.Drawing.Color, ByVal lEnvirID As Int32, ByVal iEnvirTypeID As Int16, ByVal lEntityID As Int32, ByVal iEntityTypeID As Int16, ByVal lLocX As Int32, ByVal lLocZ As Int32)
        WaitForSyncLock()
        mbInSyncLock = True
        Try
            If msHistory Is Nothing Then mlHistoryUB = -1
            mlHistoryUB += 1
            ReDim Preserve msHistory(mlHistoryUB)
            ReDim Preserve mcHistoryColor(mlHistoryUB)
            ReDim Preserve muElementClickTo(mlHistoryUB)
            For X As Int32 = mlHistoryUB To 1 Step -1
                msHistory(X) = msHistory(X - 1)
                mcHistoryColor(X) = mcHistoryColor(X - 1)
                muElementClickTo(X) = muElementClickTo(X - 1)
            Next X

            msHistory(0) = sText
            mcHistoryColor(0) = cColor
            With muElementClickTo(0)
                .lEnvirID = lEnvirID
                .iEnvirTypeID = iEnvirTypeID
                .lEntityID = lEntityID
                .iEntityTypeID = iEntityTypeID
                .lLocX = lLocX
                .lLocZ = lLocZ
            End With
        Catch ex As Exception
            'do nothing?
        End Try
        mbInSyncLock = False
    End Sub

    Private Shared Sub WaitForSyncLock()
        Dim lCnt As Int32 = 0
        While mbInSyncLock = True AndAlso lCnt < 100
            'Application.DoEvents()
            'Threading.Thread.Sleep(0)
            Threading.Thread.Sleep(1)
            lCnt += 1
        End While
    End Sub

    Private Sub frmNoteHistory_OnMouseDown(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles Me.OnMouseDown
        'figure out which item we hit...
        Dim lY As Int32 = lMouseY - Me.GetAbsolutePosition.Y

        'ok, now, we start at the bottom
        Dim lIdx As Int32 = 0
        lY = Me.Height - lY - 5
        'now, get the text height
        Dim lHt As Int32 = 16
        Try
            Using oFont As New Font(MyBase.moUILib.oDevice, moSysFont)
                Try
                    lHt = oFont.MeasureString(Nothing, "W", DrawTextFormat.Left Or DrawTextFormat.VerticalCenter, Color.White).Height
                Catch ex As Exception
                    'do nothing
                End Try
                oFont.Dispose()
            End Using
        Catch
            'do nothing
        End Try

        'now, what is the item we clicked?
        lIdx = lY \ lHt
        If lIdx <= mlHistoryUB Then
            'Ok, item clicked
            'Dim sText As String = msHistory(lIdx)
            'MyBase.moUILib.AddNotification("Item clicked: " & sText, mcHistoryColor(lIdx), -1, -1, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue) 

            HandleClickToEvent(lIdx)
        End If
    End Sub

    Private Sub frmNoteHistory_OnNewFrame() Handles Me.OnNewFrame
        If mlPreviousUB <> mlHistoryUB Then
            mlPreviousUB = mlHistoryUB
            Me.IsDirty = True
        End If
    End Sub

    Private Sub frmNoteHistory_OnRenderEnd() Handles Me.OnRenderEnd
        If GFXEngine.gbPaused = True OrElse GFXEngine.gbDeviceLost = True Then Return

        WaitForSyncLock()
        mbInSyncLock = True

        Try
            Using oFont As Font = New Font(MyBase.moUILib.oDevice, moSysFont)
                Using oTextSpr As Sprite = New Sprite(MyBase.moUILib.oDevice)
                    'Start our sprite
                    oTextSpr.Begin(SpriteFlags.AlphaBlend)

                    Try
                        Dim lHeight As Int32 = oFont.MeasureString(Nothing, "W", DrawTextFormat.Left Or DrawTextFormat.VerticalCenter, Color.White).Height
                        Dim lTop As Int32 = Me.Height - lHeight - 3
                        Dim rcDest As Rectangle = New Rectangle(5, lTop, Me.Width - 10, lHeight)

                        For X As Int32 = 0 To frmNoteHistory.mlHistoryUB
                            oFont.DrawText(oTextSpr, msHistory(X), rcDest, DrawTextFormat.Left Or DrawTextFormat.VerticalCenter, mcHistoryColor(X))
                            lTop -= lHeight
                            rcDest.Y = lTop
                            If lTop < 0 Then
                                mlHistoryUB = X
                                Exit For
                            End If
                        Next X
                    Catch
                        'do nothing
                    End Try

                    oTextSpr.End()
                    oTextSpr.Dispose()
                End Using
                oFont.Dispose()
            End Using

        Catch ex As Exception
            'do nothing?
        End Try

        mbInSyncLock = False
    End Sub

    Protected Overrides Sub Finalize()
        Dim oTmpWin As UIWindow = moUILib.GetWindow("frmNotification")
        If oTmpWin Is Nothing = False Then oTmpWin.Visible = True
        MyBase.Finalize()
    End Sub

    Private Sub HandleClickToEvent(ByVal lIdx As Int32)
        goCamera.TrackingIndex = -1
        muCurrentClickItem = muElementClickTo(lIdx)
        mbCurrentClickItemSet = True
        With muCurrentClickItem
            If .lEnvirID < 1 OrElse .iEnvirTypeID < 1 Then Return

            If .iEnvirTypeID = ObjectType.ePlanet OrElse .iEnvirTypeID = ObjectType.eSolarSystem Then
                '  are we in that environment? 

                If NewTutorialManager.TutorialOn = True Then
                    If MyBase.moUILib.CommandAllowed(True, Me.GetFullName) = False Then Return
                    NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eNotificationAlertClick, -1, -1, -1, "")
                End If

                If goCurrentEnvir Is Nothing = False Then
                    If goCurrentEnvir.ObjectID = .lEnvirID AndAlso goCurrentEnvir.ObjTypeID = .iEnvirTypeID Then
                        'Yes, we are... so no need to wait for switch
                        FinishClickToEvent()
                    Else
                        'no, change to that environment... but first set our select state
                        MyBase.moUILib.lUISelectState = UILib.eSelectState.eNotification_ClickTo
                        frmMain.ForceChangeEnvironment(.lEnvirID, .iEnvirTypeID)
                    End If
                End If
            End If
        End With
    End Sub

    Public Sub FinishClickToEvent()
        goCamera.TrackingIndex = -1
        If mbCurrentClickItemSet = True Then
            mbCurrentClickItemSet = False
            With muCurrentClickItem
                If goCurrentEnvir Is Nothing = False Then

                    If .iEnvirTypeID = ObjectType.ePlanet Then
                        glCurrentEnvirView = CurrentView.ePlanetMapView
                    Else : glCurrentEnvirView = CurrentView.eSystemMapView1
                    End If

                    With goCamera
                        .mlCameraAtX = 0 : .mlCameraAtY = 0 : .mlCameraAtZ = 0
                        .mlCameraX = 0 : .mlCameraY = 1000 : .mlCameraZ = -1000
                    End With
                    If .lLocX <> Int32.MinValue AndAlso .lLocZ <> Int32.MinValue Then
                        'Ok, use the locational coordinates
                        If .iEnvirTypeID = ObjectType.ePlanet Then
                            glCurrentEnvirView = CurrentView.ePlanetView
                            goCamera.mlCameraY = 1700
                        Else
                            glCurrentEnvirView = CurrentView.eSystemView
                            goCamera.mlCameraY = 3000
                        End If

                        goCamera.mlCameraAtX = .lLocX : goCamera.mlCameraAtY = 0 : goCamera.mlCameraAtZ = .lLocZ
                        goCamera.mlCameraX = goCamera.mlCameraAtX : goCamera.mlCameraZ = goCamera.mlCameraAtZ - 500
                        Try
                            If goCurrentEnvir.ObjTypeID = ObjectType.ePlanet Then
                                goCamera.mlCameraY = CInt(CType(goCurrentEnvir.oGeoObject, Planet).GetHeightAtPoint(goCamera.mlCameraAtX, goCamera.mlCameraAtZ, True)) + muSettings.lPlanetViewCameraY

                                goCamera.mlCameraX += muSettings.lPlanetViewCameraX
                                goCamera.mlCameraZ += muSettings.lPlanetViewCameraZ
                            End If
                        Catch
                        End Try
                    ElseIf .iEntityTypeID > -1 AndAlso .lEntityID > 0 Then
                        'Ok, now find that item...

                        For X As Int32 = 0 To goCurrentEnvir.lEntityUB
                            If goCurrentEnvir.lEntityIdx(X) = .lEntityID AndAlso goCurrentEnvir.oEntity(X).ObjTypeID = .iEntityTypeID Then

                                If .iEnvirTypeID = ObjectType.ePlanet Then
                                    glCurrentEnvirView = CurrentView.ePlanetView
                                    goCamera.mlCameraY = 1700
                                Else
                                    glCurrentEnvirView = CurrentView.eSystemView
                                    goCamera.mlCameraY = 3000
                                End If

                                With goCamera
                                    .mlCameraAtX = CInt(goCurrentEnvir.oEntity(X).LocX) : .mlCameraAtY = 0 : .mlCameraAtZ = CInt(goCurrentEnvir.oEntity(X).LocZ)
                                    .mlCameraX = .mlCameraAtX : .mlCameraZ = .mlCameraAtZ - 500

                                    Try
                                        If goCurrentEnvir.ObjTypeID = ObjectType.ePlanet Then
                                            goCamera.mlCameraY = CInt(CType(goCurrentEnvir.oGeoObject, Planet).GetHeightAtPoint(goCamera.mlCameraAtX, goCamera.mlCameraAtZ, True)) + muSettings.lPlanetViewCameraY

                                            goCamera.mlCameraX += muSettings.lPlanetViewCameraX
                                            goCamera.mlCameraZ += muSettings.lPlanetViewCameraZ
                                        End If
                                    Catch
                                    End Try
                                End With

                                'goCamera.TrackingIndex = X
                                Exit For
                            End If
                        Next X
                    End If
                End If
            End With
        End If
    End Sub

End Class