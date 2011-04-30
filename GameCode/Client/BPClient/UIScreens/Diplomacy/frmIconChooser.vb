Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Interface created from Interface Builder
Public Class frmIconChooser
    Inherits UIWindow

    Private fraResult As UIWindow
    Private WithEvents fraSpriteList As UIWindow
    Private WithEvents optBack As UIOption
    Private WithEvents optMiddle As UIOption
    Private WithEvents optFore2 As UIOption
    Private WithEvents btnAccept As UIButton
    Private WithEvents btnCancel As UIButton
    Private hscrSprites As UIScrollBar
    Private WithEvents fraColor As UIWindow
    Private hscrColor As UIScrollBar

    Private rcBack As Rectangle
    Private rcFore1 As Rectangle
    Private rcFore2 As Rectangle
    Private clrBack As System.Drawing.Color
    Private clrFore1 As System.Drawing.Color
    Private clrFore2 As System.Drawing.Color

    Private myCurrentBackImg As Byte
    Private myCurrentBackClr As Byte
    Private myCurrentFore1Img As Byte
    Private myCurrentFore1Clr As Byte
    Private myCurrentFore2Img As Byte
    Private myCurrentFore2Clr As Byte

    Private mbLoading As Boolean = True

    Private mbStandAlone As Boolean = False

    Private moIconFore As Texture
    Private moIconBack As Texture

    Public Sub New(ByRef oUILib As UILib, ByVal bStandalone As Boolean)
        MyBase.New(oUILib)

        'frmIconChooser initial props
        With Me
            .ControlName = "frmIconChooser"
            .Left = oUILib.oDevice.PresentationParameters.BackBufferWidth \ 2 - 125
            .Top = oUILib.oDevice.PresentationParameters.BackBufferHeight \ 2 - 125
            .Width = 250
            .Height = 250
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = True
            .Moveable = True
            .BorderLineWidth = 1
            If bStandalone = False Then
                .Caption = "Design Your Icon "
            End If
        End With

        'fraResult initial props
        fraResult = New UIWindow(oUILib)
        With fraResult
            .ControlName = "fraResult"
            .Left = 10
            .Top = 10
            .Width = 67
            .Height = 67
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = False
            .BorderLineWidth = 1
            .Moveable = False
        End With
        Me.AddChild(CType(fraResult, UIControl))

        'fraSpriteList initial props
        fraSpriteList = New UIWindow(oUILib)
        With fraSpriteList
            .ControlName = "fraSpriteList"
            .Left = 10
            .Top = 85
            .Width = 230
            .Height = 67
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = False
            .BorderLineWidth = 1
            .Moveable = False
        End With
        Me.AddChild(CType(fraSpriteList, UIControl))

        'optBack initial props
        optBack = New UIOption(oUILib)
        With optBack
            .ControlName = "optBack"
            .Left = 85
            .Top = 10
            .Width = 53
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Back"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = True
            .DisplayType = UIOption.eOptionTypes.eSmallOption
        End With
        Me.AddChild(CType(optBack, UIControl))

        'optMiddle initial props
        optMiddle = New UIOption(oUILib)
        With optMiddle
            .ControlName = "optMiddle"
            .Left = 85
            .Top = 30
            .Width = 63
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Middle"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = False
            .DisplayType = UIOption.eOptionTypes.eSmallOption
        End With
        Me.AddChild(CType(optMiddle, UIControl))

        'optFore2 initial props
        optFore2 = New UIOption(oUILib)
        With optFore2
            .ControlName = "optFore2"
            .Left = 85
            .Top = 50
            .Width = 52
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Front"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = False
            .DisplayType = UIOption.eOptionTypes.eSmallOption
        End With
        Me.AddChild(CType(optFore2, UIControl))

        'btnAccept initial props
        btnAccept = New UIButton(oUILib)
        With btnAccept
            .ControlName = "btnAccept"
            .Left = 160
            .Top = 10
            .Width = 80
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Accept"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnAccept, UIControl))

        'btnCancel initial props
        btnCancel = New UIButton(oUILib)
        With btnCancel
            .ControlName = "btnCancel"
            .Left = 160
            .Top = 50
            .Width = 80
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

        'hscrSprites initial props
        hscrSprites = New UIScrollBar(oUILib, False)
        With hscrSprites
            .ControlName = "hscrSprites"
            .Left = 10
            .Top = 154
            .Width = 230
            .Height = 24
            .Enabled = True
            .Visible = True
            .Value = 0
            .MaxValue = 100
            .MinValue = 0
            .SmallChange = 1
            .LargeChange = 4
            .ReverseDirection = False
        End With
        Me.AddChild(CType(hscrSprites, UIControl))

        'fraColor initial props
        fraColor = New UIWindow(oUILib)
        With fraColor
            .ControlName = "fraColor"
            .Left = 10
            .Top = 185
            .Width = 230
            .Height = 33
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = False
            .Moveable = False
            .BorderLineWidth = 1
        End With
        Me.AddChild(CType(fraColor, UIControl))

        'hscrColor initial props
        hscrColor = New UIScrollBar(oUILib, False)
        With hscrColor
            .ControlName = "hscrColor"
            .Left = 10
            .Top = 220
            .Width = 230
            .Height = 24
            .Enabled = True
            .Visible = True
            .Value = 0
            .MaxValue = 100
            .MinValue = 0
            .SmallChange = 1
            .LargeChange = 4
            .ReverseDirection = False
        End With
        Me.AddChild(CType(hscrColor, UIControl))

        If bStandalone = True Then
            MyBase.moUILib.RemoveWindow(Me.ControlName)
            MyBase.moUILib.AddWindow(Me)
        Else
            btnAccept.Visible = False
            btnCancel.Visible = False
        End If

        Dim lTemp As Int32 = CInt(Rnd() * (Int32.MaxValue - 1))
        SetStartingIcon(lTemp)
        SetScrollValues()

        mbStandAlone = bStandalone

        mbLoading = False
    End Sub

    Private Sub SetScrollValues()
        If optBack.Value = True Then
            hscrSprites.MaxValue = 16 - ((fraSpriteList.Width - 6) \ 64)
        Else : hscrSprites.MaxValue = 64 - ((fraSpriteList.Width - 6) \ 64)
        End If
        hscrColor.MaxValue = 16 - (fraColor.Width \ 32)
    End Sub

    Public Sub SetStartingIcon(ByVal lPlayerIcon As Int32)
        Dim yBackImg As Byte
        Dim yBackClr As Byte
        Dim yFore1Img As Byte
        Dim yFore1Clr As Byte
        Dim yFore2Img As Byte
        Dim yFore2Clr As Byte

        PlayerIconManager.FillIconValues(lPlayerIcon, yBackImg, yBackClr, yFore1Img, yFore1Clr, yFore2Img, yFore2Clr)

        rcBack = PlayerIconManager.ReturnImageRectangle(yBackImg, True)
        rcFore1 = PlayerIconManager.ReturnImageRectangle(yFore1Img, False)
        rcFore2 = PlayerIconManager.ReturnImageRectangle(yFore2Img, False)

        clrBack = PlayerIconManager.GetColorValue(yBackClr)
        clrFore1 = PlayerIconManager.GetColorValue(yFore1Clr)
        clrFore2 = PlayerIconManager.GetColorValue(yFore2Clr)

        myCurrentBackImg = yBackImg
        myCurrentBackClr = yBackClr
        myCurrentFore1Img = yFore1Img
        myCurrentFore1Clr = yFore1Clr
        myCurrentFore2Img = yFore2Img
        myCurrentFore2Clr = yFore2Clr
    End Sub

    Public Function GetStartingIcon() As Int32
        Return PlayerIconManager.CreatePlayerIconNumber(myCurrentBackImg, myCurrentBackClr, myCurrentFore1Img, myCurrentFore1Clr, myCurrentFore2Img, myCurrentFore2Clr)
    End Function

	Private Sub btnAccept_Click(ByVal sName As String) Handles btnAccept.Click
        'goUILib.AddNotification("That functionality is not implemented yet.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
        Dim lIcon As Int32 = GetStartingIcon()
        If lIcon = 0 Then
            MyBase.moUILib.AddNotification("Select a valid icon.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Return
        End If

        If goCurrentPlayer.blCredits > 1000000 Then
            goCurrentPlayer.lPlayerIcon = lIcon
            Dim yMsg(5) As Byte
            System.BitConverter.GetBytes(GlobalMessageCode.eUpdatePlayerDetails).CopyTo(yMsg, 0)
            System.BitConverter.GetBytes(lIcon).CopyTo(yMsg, 2)
            MyBase.moUILib.SendMsgToPrimary(yMsg)
            MyBase.moUILib.RemoveWindow(Me.ControlName)
        Else
            MyBase.moUILib.AddNotification("Insufficient credits. Recreating an icon requires a million credits.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
        End If

	End Sub

	Private Sub btnCancel_Click(ByVal sName As String) Handles btnCancel.Click
		MyBase.moUILib.RemoveWindow(Me.ControlName)
	End Sub

    Private Sub fraColor_OnMouseDown(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles fraColor.OnMouseDown
        lMouseX -= fraColor.GetAbsolutePosition.X
        Dim lIdx As Int32 = (lMouseX \ 32) + hscrColor.Value
        If optBack.Value = True Then
            clrBack = PlayerIconManager.GetColorValue(CByte(lIdx))
            myCurrentBackClr = CByte(lIdx)
        ElseIf optMiddle.Value = True Then
            clrFore1 = PlayerIconManager.GetColorValue(CByte(lIdx))
            myCurrentFore1Clr = CByte(lIdx)
        Else
            clrFore2 = PlayerIconManager.GetColorValue(CByte(lIdx))
            myCurrentFore2Clr = CByte(lIdx)
        End If
        Me.IsDirty = True
    End Sub

    Private Sub fraSpriteList_OnMouseDown(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles fraSpriteList.OnMouseDown
        lMouseX -= fraSpriteList.GetAbsolutePosition.X
        Dim lIdx As Int32 = (lMouseX \ 64) + hscrSprites.Value
        If optBack.Value = True Then
            rcBack = PlayerIconManager.ReturnImageRectangle(CByte(lIdx), True)
            myCurrentBackImg = CByte(lIdx)
        ElseIf optMiddle.Value = True Then
            rcFore1 = PlayerIconManager.ReturnImageRectangle(CByte(lIdx), False)
            myCurrentFore1Img = CByte(lIdx)
        Else
            rcFore2 = PlayerIconManager.ReturnImageRectangle(CByte(lIdx), False)
            myCurrentFore2Img = CByte(lIdx)
        End If
        Me.IsDirty = True
    End Sub

    Private Sub optBack_Click() Handles optBack.Click
        If mbLoading = True Then Return
        mbLoading = True
        optFore2.Value = False
        optMiddle.Value = False
        optBack.Value = True
        mbLoading = False
        SetScrollValues()
    End Sub

    Private Sub optFore2_Click() Handles optFore2.Click
        If mbLoading = True Then Return
        mbLoading = True
        optFore2.Value = True
        optMiddle.Value = False
        optBack.Value = False
        mbLoading = False
        SetScrollValues()
    End Sub

    Private Sub optMiddle_Click() Handles optMiddle.Click
        If mbLoading = True Then Return
        mbLoading = True
        optFore2.Value = False
        optMiddle.Value = True
        optBack.Value = False
        mbLoading = False
        SetScrollValues()
    End Sub

    Public Sub frmIconChooser_OnRenderEnd() Handles Me.OnRenderEnd
        If GFXEngine.gbPaused = True OrElse GFXEngine.gbDeviceLost = True Then Return
        If moIconFore Is Nothing Then moIconFore = goResMgr.GetTexture("DipPlayerFore.dds", GFXResourceManager.eGetTextureType.UserInterface, "Interface.pak")
        If moIconBack Is Nothing Then moIconBack = goResMgr.GetTexture("DipPlayerBack.bmp", GFXResourceManager.eGetTextureType.UserInterface, "Interface.pak")

        goUILib.oDevice.RenderState.ZBufferEnable = False

        Using oSprite As New Sprite(goUILib.oDevice)

            oSprite.Begin(SpriteFlags.AlphaBlend)
            Try
                'Ok, draw the Result Icon...
                Dim rcDest As Rectangle = New Rectangle(0, 0, 64, 64)
                Dim ptDest As Point

                If mbStandAlone = True Then
                    ptDest.X = fraResult.Left + 3
                    ptDest.Y = fraResult.Top + 3
                Else
                    ptDest = fraResult.GetAbsolutePosition()
                    ptDest.X += 2
                    ptDest.Y += 2
                End If
                oSprite.Draw2D(moIconBack, rcBack, rcDest, ptDest, clrBack)
                oSprite.Draw2D(moIconFore, rcFore1, rcDest, ptDest, clrFore1)
                oSprite.Draw2D(moIconFore, rcFore2, rcDest, ptDest, clrFore2)

                'Now, draw our lists... first... the sprite list
                Dim XMax As Int32 = (fraSpriteList.Width - 6) \ 64
                XMax += hscrSprites.Value - 1

                'Now, loop
                If mbStandAlone = True Then
                    ptDest.X = fraSpriteList.Left
                    ptDest.Y = fraSpriteList.Top
                Else
                    ptDest = fraSpriteList.GetAbsolutePosition
                End If

                Dim lExtentX As Int32 = ptDest.X + fraSpriteList.Width
                ptDest.X += 3 : ptDest.Y += 3
                rcDest = New Rectangle(0, 0, 64, 64)

                Dim clrVal As System.Drawing.Color = Color.White
                If optBack.Value = True Then
                    clrVal = clrBack
                ElseIf optMiddle.Value = True Then
                    clrVal = clrFore1
                Else : clrVal = clrFore2
                End If

                For lIdx As Int32 = hscrSprites.Value To XMax
                    Dim rcSrc As Rectangle = PlayerIconManager.ReturnImageRectangle(CByte(lIdx), optBack.Value)
                    'If ptDest.X + rcDest.Width > lExtentX Then
                    '    rcDest.Width -= ((ptDest.X + rcDest.Width) - lExtentX)
                    'End If

                    If optBack.Value = True Then
                        oSprite.Draw2D(moIconBack, rcSrc, rcDest, ptDest, clrVal)
                    Else
                        oSprite.Draw2D(moIconFore, rcSrc, rcDest, ptDest, clrVal)
                    End If
                    ptDest.X += 64
                Next lIdx

                'Now, the colors...
                If mbStandAlone = True Then
                    ptDest.X = fraColor.Left
                    ptDest.Y = fraColor.Top + 1
                Else
                    ptDest = fraColor.GetAbsolutePosition()
                    ptDest.X += 1
                    ptDest.Y += 1
                End If

                lExtentX = ptDest.X + fraColor.Width
                rcDest = New Rectangle(0, 0, 32, 32)
                XMax = (fraColor.Width \ 32) + hscrColor.Value - 1
                Dim rcClrSrc As Rectangle = New Rectangle(209, 15, 32, 32)
                For lIdx As Int32 = hscrColor.Value To XMax
                    oSprite.Draw2D(goUILib.oInterfaceTexture, rcClrSrc, rcDest, ptDest, PlayerIconManager.GetColorValue(CByte(lIdx)))
                    ptDest.X += 32
                Next lIdx
            Catch
                Me.IsDirty = True
            End Try

            oSprite.End()
        End Using
        goUILib.oDevice.RenderState.ZBufferEnable = True
    End Sub

    Protected Overrides Sub Finalize()
		If goResMgr Is Nothing = False Then
			goResMgr.DeleteTexture("DipPlayerFore.dds")
			goResMgr.DeleteTexture("DipPlayerBack.bmp")
		End If
		MyBase.Finalize()
    End Sub
End Class