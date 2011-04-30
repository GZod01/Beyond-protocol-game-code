Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

Public Class UIHullSlots
    Inherits UIControl

    Public Enum eyHullRequirements As Byte
        RequiresAnEngine = 1
        EngineIsToBeInRear = 2
        RequiresFiftyThousandStructHP = 4
        RequiresBayDoor = 8
    End Enum

    Private Const ml_GRID_SIZE_WH As Int32 = 30     'number of squares in each direction
    Private Const ml_SQUARE_SIZE As Int32 = 16      'Width and Height of each square
    Private Structure ValPair
        Public ValueConfig As SlotConfig
        Public ValueGroupNum As Int32
        Public Value As Int32
    End Structure

    Private moSlots(,) As HullSlot
    Public AutoRefresh As Boolean = True

    Private mcolVal() As System.Drawing.Color
    'Private mcolFiltered As System.Drawing.Color = System.Drawing.Color.FromArgb(255, 32, 32, 32)

    Private mlFrontSlots As Int32 = 0
    Private mlLeftSlots As Int32 = 0
    Private mlRearSlots As Int32 = 0
    Private mlRightSlots As Int32 = 0
    Private mlAllArcSlots As Int32 = 0

    Private mbRenderHullIcons As Boolean = True
    Public Shared moHullTex As Texture = Nothing

    Public LastHullSize As Int32 = 0

    Public Event HullSlotClick(ByVal lIndexX As Int32, ByVal lIndexY As Int32, ByVal lButton As System.Windows.Forms.MouseButtons)

    Public Property RenderHullIcons() As Boolean
        Get
            Return mbRenderHullIcons
        End Get
        Set(ByVal value As Boolean)
            If mbRenderHullIcons <> value Then
                mbRenderHullIcons = value
                Me.IsDirty = True

                If mbRenderHullIcons = True Then
                    ReDim mcolVal(5)
                    mcolVal(0) = System.Drawing.Color.FromArgb(255, 32, 32, 32)
                    mcolVal(1) = System.Drawing.Color.FromArgb(255, 0, 0, 255)
                    mcolVal(2) = System.Drawing.Color.FromArgb(255, 255, 0, 255)
                    mcolVal(3) = System.Drawing.Color.FromArgb(255, 0, 255, 0)
                    mcolVal(4) = System.Drawing.Color.FromArgb(255, 255, 0, 0)
                    mcolVal(5) = System.Drawing.Color.FromArgb(255, 255, 255, 0)
                Else
                    ReDim mcolVal(5)
                    mcolVal(0) = System.Drawing.Color.FromArgb(255, 32, 32, 32)
                    mcolVal(1) = System.Drawing.Color.FromArgb(255, 0, 0, 192)
                    mcolVal(2) = System.Drawing.Color.FromArgb(255, 192, 0, 192)
                    mcolVal(3) = System.Drawing.Color.FromArgb(255, 0, 192, 0)
                    mcolVal(4) = System.Drawing.Color.FromArgb(255, 192, 0, 0)
                    mcolVal(5) = System.Drawing.Color.FromArgb(255, 192, 192, 0)
                End If
            End If
        End Set
    End Property

    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)

        ReDim moSlots(ml_GRID_SIZE_WH - 1, ml_GRID_SIZE_WH - 1)

        For Y As Int32 = 0 To ml_GRID_SIZE_WH - 1
            For X As Int32 = 0 To ml_GRID_SIZE_WH - 1
                With moSlots(X, Y)
                    .lConfig = SlotConfig.eArmorConfig
                    .yType = SlotType.eUnused
                    .lGroupNum = 0
                End With
            Next X
        Next Y

        Dim oINI As InitFile = New InitFile
        If CInt(Val(oINI.GetString("INTERFACE", "HullSlots_RenderIcons", "1"))) = 0 Then
            Me.RenderHullIcons = False
        Else : Me.RenderHullIcons = True
        End If
        oINI = Nothing

        If mbRenderHullIcons = True Then
            ReDim mcolVal(5)
            mcolVal(0) = System.Drawing.Color.FromArgb(255, 32, 32, 32)
            mcolVal(1) = System.Drawing.Color.FromArgb(255, 0, 0, 255)
            mcolVal(2) = System.Drawing.Color.FromArgb(255, 255, 0, 255)
            mcolVal(3) = System.Drawing.Color.FromArgb(255, 0, 255, 0)
            mcolVal(4) = System.Drawing.Color.FromArgb(255, 255, 0, 0)
            mcolVal(5) = System.Drawing.Color.FromArgb(255, 255, 255, 0)
        Else
            ReDim mcolVal(5)
            mcolVal(0) = System.Drawing.Color.FromArgb(255, 32, 32, 32)
            mcolVal(1) = System.Drawing.Color.FromArgb(255, 0, 0, 192)
            mcolVal(2) = System.Drawing.Color.FromArgb(255, 192, 0, 192)
            mcolVal(3) = System.Drawing.Color.FromArgb(255, 0, 192, 0)
            mcolVal(4) = System.Drawing.Color.FromArgb(255, 192, 0, 0)
            mcolVal(5) = System.Drawing.Color.FromArgb(255, 192, 192, 0)
        End If

        Me.Width = (ml_SQUARE_SIZE * ml_GRID_SIZE_WH) + 2
        Me.Height = Me.Width
    End Sub

    Private Sub UIHullSlots_OnMouseMove(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles Me.OnMouseMove
        If lButton = MouseButtons.Left OrElse lButton = MouseButtons.Right Then
            Dim oLoc As System.Drawing.Point = Me.GetAbsolutePosition()
            Dim lTmpX As Int32 = lMouseX - oLoc.X - 1
            Dim lTmpY As Int32 = lMouseY - oLoc.Y - 1

            lTmpX \= ml_SQUARE_SIZE
            lTmpY \= ml_SQUARE_SIZE

            If lTmpX > -1 AndAlso lTmpY > -1 AndAlso lTmpX < ml_GRID_SIZE_WH AndAlso lTmpY < ml_GRID_SIZE_WH Then
                If moSlots(lTmpX, lTmpY).yType <> SlotType.eUnused AndAlso moSlots(lTmpX, lTmpY).bFiltered = False Then
                    RaiseEvent HullSlotClick(lTmpX, lTmpY, lButton)
                End If
            End If
        Else
            Dim oLoc As System.Drawing.Point = Me.GetAbsolutePosition()
            Dim lTmpX As Int32 = lMouseX - oLoc.X - 1
            Dim lTmpY As Int32 = lMouseY - oLoc.Y - 1

            lTmpX \= ml_SQUARE_SIZE
            lTmpY \= ml_SQUARE_SIZE

            If lTmpX > -1 AndAlso lTmpY > -1 AndAlso lTmpX < ml_GRID_SIZE_WH AndAlso lTmpY < ml_GRID_SIZE_WH Then
                If moSlots(lTmpX, lTmpY).yType <> SlotType.eUnused AndAlso moSlots(lTmpX, lTmpY).bFiltered = False Then

                    'If moSlots(lTmpX, lTmpY).lConfig = SlotConfig.eHangarDoor Then
                    '    Dim lGrpNum As Int32 = moSlots(lTmpX, lTmpY).lGroupNum

                    '    Dim lTotalCnt As Int32 = 0
                    '    For X As Int32 = 0 To ml_GRID_SIZE_WH - 1
                    '        For Y As Int32 = 0 To ml_GRID_SIZE_WH - 1
                    '            If moSlots(X, Y).yType <> SlotType.eUnused AndAlso moSlots(X, Y).lGroupNum = lGrpNum AndAlso moSlots(X, Y).lConfig = SlotConfig.eHangarDoor Then
                    '                lTotalCnt += 1
                    '            End If
                    '        Next Y
                    '    Next X

                    '    Dim fHullSize As Single = lTotalCnt * CSng(LastHullSize / TotalSlots())
                    '    MyBase.moUILib.SetToolTip("Bay Door " & lGrpNum & " size: " & fHullSize.ToString("#,##0.#"), lMouseX, lMouseY)

                    'End If

                    Dim lGrpNum As Int32 = moSlots(lTmpX, lTmpY).lGroupNum
                    Dim lConfig As Int32 = moSlots(lTmpX, lTmpY).lConfig
                    Dim yType As SlotType = moSlots(lTmpX, lTmpY).yType
                    Dim lTotalCnt As Int32 = 0
                    For X As Int32 = 0 To ml_GRID_SIZE_WH - 1
                        For Y As Int32 = 0 To ml_GRID_SIZE_WH - 1
                            If moSlots(X, Y).yType <> SlotType.eUnused AndAlso moSlots(X, Y).lGroupNum = lGrpNum AndAlso moSlots(X, Y).lConfig = lConfig Then
                                If lConfig = SlotConfig.eArmorConfig Then
                                    'only count armor slots if they are in the same Arc
                                    If moSlots(X, Y).yType = yType Then lTotalCnt += 1
                                Else : lTotalCnt += 1
                                End If
                            End If
                        Next Y
                    Next X
                    Dim fHullSize As Single = lTotalCnt * CSng(LastHullSize / TotalSlots())

                    Dim sText As String = ""
                    Select Case moSlots(lTmpX, lTmpY).lConfig
                        Case SlotConfig.eArmorConfig
                            sText = "Armor and Crew space: " & fHullSize.ToString("#,##0.#")
                            If yType = SlotType.eAllArc Then
                                sText &= vbCrLf & "AllArc can not be assigned as armor"
                            End If
                        Case SlotConfig.eCargoBay
                            sText = "Cargo Bay size: " & fHullSize.ToString("#,##0.#")
                        Case SlotConfig.eEngines
                            sText = "Engine size: " & fHullSize.ToString("#,##0.#")
                        Case SlotConfig.eHangar
                            sText = "Hanger size: " & fHullSize.ToString("#,##0.#")
                        Case SlotConfig.eHangarDoor
                            sText = "Bay Door " & lGrpNum & " size: " & fHullSize.ToString("#,##0.#")
                        Case SlotConfig.eRadar
                            sText = "Radar size: " & fHullSize.ToString("#,##0.#")
                        Case SlotConfig.eShields
                            sText = "Shield size: " & fHullSize.ToString("#,##0.#")
                        Case SlotConfig.eWeapons
                            sText = "Weapon group " & lGrpNum & " size: " & fHullSize.ToString("#,##0.#")
                    End Select
                    If sText <> "" Then MyBase.moUILib.SetToolTip(sText, lMouseX, lMouseY)

                End If
            End If
        End If
    End Sub

    Private Sub UIHullSlots_OnMouseUp(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles Me.OnMouseUp
        Dim oLoc As System.Drawing.Point = Me.GetAbsolutePosition()
        Dim lTmpX As Int32 = lMouseX - oLoc.X - 1
        Dim lTmpY As Int32 = lMouseY - oLoc.Y - 1

        lTmpX \= ml_SQUARE_SIZE
        lTmpY \= ml_SQUARE_SIZE

        If lTmpX > -1 AndAlso lTmpY > -1 AndAlso lTmpX < ml_GRID_SIZE_WH AndAlso lTmpY < ml_GRID_SIZE_WH Then
			If moSlots(lTmpX, lTmpY).yType <> SlotType.eUnused AndAlso moSlots(lTmpX, lTmpY).bFiltered = False Then
                RaiseEvent HullSlotClick(lTmpX, lTmpY, lButton)
			End If
        End If
    End Sub

    'Private Sub UIHullSlots_OnRender(ByRef oImgSprite As Sprite, ByRef oTextSprite As Sprite) Handles Me.OnRender
    Private Sub UIHullSlots_OnRender() Handles Me.OnRender
        Dim oLoc As System.Drawing.Point = GetAbsolutePosition()

        Dim oBorderLineVerts(4) As Vector2
        'Draw a box border around our window...
        With oBorderLineVerts(0)
            .X = oLoc.X
            .Y = oLoc.Y
        End With
        With oBorderLineVerts(1)
            .X = oLoc.X + Width
            .Y = oLoc.Y
        End With
        With oBorderLineVerts(2)
            .X = oLoc.X + Width
            .Y = oLoc.Y + Height
        End With
        With oBorderLineVerts(3)
            .X = oLoc.X
            .Y = oLoc.Y + Height
        End With
        With oBorderLineVerts(4)
            .X = oLoc.X
            .Y = oLoc.Y
        End With

        'Draw a box border around our window...
        Using moBorderLine As New Line(MyBase.moUILib.oDevice)
            moBorderLine.Antialias = True
            moBorderLine.Width = 2
            moBorderLine.Begin()
            moBorderLine.Draw(oBorderLineVerts, muSettings.InterfaceBorderColor)
            moBorderLine.End()
        End Using
        'End of the border drawing

        If mbRenderHullIcons = True AndAlso (moHullTex Is Nothing = True OrElse moHullTex.Disposed = True) Then
            moHullTex = goResMgr.GetTexture("HullBuilderIcons.dds", GFXResourceManager.eGetTextureType.UserInterface, "Interface.pak")
        End If

        'Now, render our stuff
        Dim lTmpX As Int32
        Dim lTmpY As Int32

        If GFXEngine.gbPaused = False AndAlso GFXEngine.gbDeviceLost = False Then

            If mbRenderHullIcons = True Then
                Using oSpr As Sprite = New Sprite(MyBase.moUILib.oDevice)

                    Dim rcSrc() As Rectangle
                    Dim rcDest As Rectangle
                    Dim fX As Single
                    Dim fY As Single

                    ReDim rcSrc(SlotConfig.eConfigCnt - 1)

                    rcSrc(SlotConfig.eArmorConfig) = New Rectangle(0, 0, 16, 16)
                    rcSrc(SlotConfig.eCargoBay) = New Rectangle(32, 0, 16, 16)
                    rcSrc(SlotConfig.eCrewQuarters) = New Rectangle(16, 0, 16, 16)
                    rcSrc(SlotConfig.eEngines) = New Rectangle(0, 16, 16, 16)
                    rcSrc(SlotConfig.eFuelBay) = New Rectangle(32, 16, 16, 16)
                    rcSrc(SlotConfig.eHangar) = New Rectangle(48, 16, 16, 16)
                    rcSrc(SlotConfig.eRadar) = New Rectangle(48, 0, 16, 16)
                    rcSrc(SlotConfig.eShields) = New Rectangle(16, 16, 16, 16)
                    rcSrc(SlotConfig.eWeapons) = New Rectangle(0, 32, 16, 16)
                    rcSrc(SlotConfig.eHangarDoor) = New Rectangle(16, 32, 16, 16)

                    oSpr.Begin(SpriteFlags.AlphaBlend)

                    For Y As Int32 = 0 To ml_GRID_SIZE_WH - 1
                        For X As Int32 = 0 To ml_GRID_SIZE_WH - 1
                            With moSlots(X, Y)
                                If .yType <> SlotType.eUnused Then
                                    lTmpX = oLoc.X + (X * ml_SQUARE_SIZE) + 1
                                    lTmpY = oLoc.Y + (Y * ml_SQUARE_SIZE) + 1

                                    fX = lTmpX * (rcSrc(.lConfig).Width / CSng(16))
                                    fY = lTmpY * (rcSrc(.lConfig).Height / CSng(16))

                                    rcDest = Rectangle.FromLTRB(oLoc.X, oLoc.Y, oLoc.X + 16, oLoc.Y + 16)

                                    If .bFiltered = True Then
                                        Dim clrFiltered As System.Drawing.Color = System.Drawing.Color.FromArgb(255, mcolVal(.yType).R \ 5, mcolVal(.yType).G \ 5, mcolVal(.yType).B \ 5)
                                        oSpr.Draw2D(moHullTex, rcSrc(.lConfig), rcDest, System.Drawing.Point.Empty, 0, New Point(CInt(fX), CInt(fY)), clrFiltered)
                                    Else : oSpr.Draw2D(moHullTex, rcSrc(.lConfig), rcDest, System.Drawing.Point.Empty, 0, New Point(CInt(fX), CInt(fY)), mcolVal(.yType))
                                    End If

                                End If
                            End With
                        Next X
                    Next Y
                    oSpr.End()
                    oSpr.Dispose()
                End Using

                'Now, go back through for our Weapons...
                Using oFont As Font = New Font(MyBase.moUILib.oDevice, New System.Drawing.Font("Arial", 10, FontStyle.Bold, GraphicsUnit.Point, 0))
                    Dim sCap As String
                    Dim rcDest As Rectangle

                    Using oTextSpr As Sprite = New Sprite(MyBase.moUILib.oDevice)
                        oTextSpr.Begin(SpriteFlags.AlphaBlend)
                        For Y As Int32 = 0 To ml_GRID_SIZE_WH - 1
                            For X As Int32 = 0 To ml_GRID_SIZE_WH - 1
                                With moSlots(X, Y)
                                    If .yType <> SlotType.eUnused AndAlso (.lConfig = SlotConfig.eWeapons OrElse .lConfig = SlotConfig.eHangarDoor) AndAlso .bFiltered = False Then
                                        lTmpX = oLoc.X + 1 + (X * ml_SQUARE_SIZE)
                                        lTmpY = oLoc.Y + 1 + (Y * ml_SQUARE_SIZE)

                                        rcDest = Rectangle.FromLTRB(lTmpX, lTmpY, lTmpX + 16, lTmpY + 16)
                                        sCap = .lGroupNum.ToString
                                        If sCap <> "" Then
                                            oFont.DrawText(oTextSpr, sCap, rcDest, DrawTextFormat.Center Or DrawTextFormat.VerticalCenter, Color.White)
                                        End If
                                    End If
                                End With
                            Next X
                        Next Y
                        oTextSpr.End()
                        oTextSpr.Dispose()
                    End Using

                    oFont.Dispose()
                End Using
            Else
                Using oSpr As Sprite = New Sprite(MyBase.moUILib.oDevice)

                    Dim rcSrc As Rectangle = grc_UI(elInterfaceRectangle.eCheck_Unchecked)
                    Dim rcDest As Rectangle
                    Dim fX As Single
                    Dim fY As Single

                    oSpr.Begin(SpriteFlags.AlphaBlend)

                    For Y As Int32 = 0 To ml_GRID_SIZE_WH - 1
                        For X As Int32 = 0 To ml_GRID_SIZE_WH - 1
                            With moSlots(X, Y)
                                If .yType <> SlotType.eUnused Then
                                    lTmpX = oLoc.X + (X * ml_SQUARE_SIZE) + 1
                                    lTmpY = oLoc.Y + (Y * ml_SQUARE_SIZE) + 1

                                    fX = lTmpX * (rcSrc.Width / CSng(16))
                                    fY = lTmpY * (rcSrc.Height / CSng(16))

                                    rcDest = Rectangle.FromLTRB(oLoc.X, oLoc.Y, oLoc.X + 16, oLoc.Y + 16)

                                    If .bFiltered = True Then
                                        Dim clrFiltered As System.Drawing.Color = System.Drawing.Color.FromArgb(255, mcolVal(.yType).R \ 5, mcolVal(.yType).G \ 5, mcolVal(.yType).B \ 5)
                                        oSpr.Draw2D(MyBase.moUILib.oInterfaceTexture, rcSrc, rcDest, System.Drawing.Point.Empty, 0, New Point(CInt(fX), CInt(fY)), clrFiltered)
                                    Else : oSpr.Draw2D(MyBase.moUILib.oInterfaceTexture, rcSrc, rcDest, System.Drawing.Point.Empty, 0, New Point(CInt(fX), CInt(fY)), mcolVal(.yType))
                                    End If

                                End If
                            End With
                        Next X
                    Next Y
                    oSpr.End()
                    oSpr.Dispose()
                End Using

                Using oFont As Font = New Font(MyBase.moUILib.oDevice, New System.Drawing.Font("Arial", 10, FontStyle.Bold, GraphicsUnit.Point, 0))
                    Dim sCap As String
                    Dim rcDest As Rectangle
                    Using oTextSpr As Sprite = New Sprite(MyBase.moUILib.oDevice)
                        oTextSpr.Begin(SpriteFlags.AlphaBlend)
                        For Y As Int32 = 0 To ml_GRID_SIZE_WH - 1
                            For X As Int32 = 0 To ml_GRID_SIZE_WH - 1
                                With moSlots(X, Y)
                                    lTmpX = oLoc.X + 1 + (X * ml_SQUARE_SIZE)
                                    lTmpY = oLoc.Y + 1 + (Y * ml_SQUARE_SIZE)

                                    rcDest = Rectangle.FromLTRB(lTmpX, lTmpY, lTmpX + 16, lTmpY + 16)
                                    If .yType <> SlotType.eUnused AndAlso .bFiltered = False Then
                                        Select Case .lConfig
                                            Case SlotConfig.eCargoBay
                                                sCap = "C"
                                            Case SlotConfig.eCrewQuarters
                                                sCap = "Q"
                                            Case SlotConfig.eEngines
                                                sCap = "E"
                                            Case SlotConfig.eFuelBay
                                                sCap = "F"
                                            Case SlotConfig.eHangar
                                                sCap = "H"
                                            Case SlotConfig.eRadar
                                                sCap = "R"
                                            Case SlotConfig.eShields
                                                sCap = "S"
                                            Case SlotConfig.eWeapons
                                                sCap = .lGroupNum.ToString
                                            Case SlotConfig.eHangarDoor
                                                sCap = "D"
                                            Case Else
                                                sCap = ""
                                        End Select
                                        If sCap <> "" Then
                                            oFont.DrawText(oTextSpr, sCap, rcDest, DrawTextFormat.Center Or DrawTextFormat.VerticalCenter, Color.White)
                                        End If
                                    End If
                                End With
                            Next X
                        Next Y
                        oTextSpr.End()
                        oTextSpr.Dispose()
                    End Using

                    oFont.Dispose()
                End Using
            End If
        End If
    End Sub

    Public Sub SetHullSlot(ByVal lX As Int32, ByVal lY As Int32, ByVal yType As SlotType, ByVal lConfig As SlotConfig, ByVal lGroupNum As Int32)
        If lX > -1 AndAlso lY > -1 AndAlso lX < ml_GRID_SIZE_WH AndAlso lY < ml_GRID_SIZE_WH Then
            With moSlots(lX, lY)
                .lConfig = lConfig
                .lGroupNum = lGroupNum
                If yType <> SlotType.eNoChange Then
                    Select Case .yType
                        Case SlotType.eFront
                            mlFrontSlots -= 1
                        Case SlotType.eLeft
                            mlLeftSlots -= 1
                        Case SlotType.eRear
                            mlRearSlots -= 1
                        Case SlotType.eRight
                            mlRightSlots -= 1
                        Case SlotType.eAllArc
                            mlAllArcSlots -= 1
                    End Select
                    .yType = yType
                    Select Case .yType
                        Case SlotType.eFront
                            mlFrontSlots += 1
                        Case SlotType.eLeft
                            mlLeftSlots += 1
                        Case SlotType.eRear
                            mlRearSlots += 1
                        Case SlotType.eRight
                            mlRightSlots += 1
                        Case SlotType.eAllArc
                            mlAllArcSlots += 1
                    End Select
                End If
            End With
            MyBase.IsDirty = True
        End If
    End Sub

    Public Sub GetHullSlotValues(ByVal lX As Int32, ByVal lY As Int32, ByRef yType As SlotType, ByRef lConfig As SlotConfig, ByRef lGroupNum As Int32)
        If lX > -1 AndAlso lY > -1 AndAlso lX < ml_GRID_SIZE_WH AndAlso lY < ml_GRID_SIZE_WH Then
            With moSlots(lX, lY)
                yType = .yType
                lConfig = .lConfig
                lGroupNum = .lGroupNum
            End With
        End If
    End Sub

    Public Sub SetTypeColor(ByVal lType As SlotType, ByVal colNewColor As System.Drawing.Color)
        mcolVal(lType) = colNewColor
    End Sub

    Public Sub SetByModelID(ByVal lModelID As Int32)
        Dim X As Int32
        Dim Y As Int32
        Dim lTemp As Int32

        For Y = 0 To ml_GRID_SIZE_WH - 1
            For X = 0 To ml_GRID_SIZE_WH - 1
                moSlots(X, Y).yType = SlotType.eUnused
                moSlots(X, Y).bFiltered = False
            Next X
        Next Y

        Dim oModelDef As ModelDef = goModelDefs.GetModelDef(lModelID)

        If oModelDef Is Nothing = False Then
            With oModelDef
                'Forward
                For lIdx As Int32 = 0 To .FrontLocs.GetUpperBound(0)
                    lTemp = .FrontLocs(lIdx)
                    Y = lTemp \ ml_GRID_SIZE_WH
                    X = lTemp - (Y * ml_GRID_SIZE_WH)
                    moSlots(X, Y).yType = SlotType.eFront
                Next lIdx

                'Left
                For lIdx As Int32 = 0 To .LeftLocs.GetUpperBound(0)
                    lTemp = .LeftLocs(lIdx)
                    Y = lTemp \ ml_GRID_SIZE_WH
                    X = lTemp - (Y * ml_GRID_SIZE_WH)
                    moSlots(X, Y).yType = SlotType.eLeft
                Next lIdx

                'Rear
                For lIdx As Int32 = 0 To .RearLocs.GetUpperBound(0)
                    lTemp = .RearLocs(lIdx)
                    Y = lTemp \ ml_GRID_SIZE_WH
                    X = lTemp - (Y * ml_GRID_SIZE_WH)
                    moSlots(X, Y).yType = SlotType.eRear
                Next lIdx

                'Right
                For lIdx As Int32 = 0 To .RightLocs.GetUpperBound(0)
                    lTemp = .RightLocs(lIdx)
                    Y = lTemp \ ml_GRID_SIZE_WH
                    X = lTemp - (Y * ml_GRID_SIZE_WH)
                    moSlots(X, Y).yType = SlotType.eRight
                Next lIdx

                'All Arc
                For lIdx As Int32 = 0 To .AllArcLocs.GetUpperBound(0)
                    lTemp = .AllArcLocs(lIdx)
                    Y = lTemp \ ml_GRID_SIZE_WH
                    X = lTemp - (Y * ml_GRID_SIZE_WH)
                    moSlots(X, Y).yType = SlotType.eAllArc
                Next lIdx
            End With

            'Now, go back and reset our counts...
            mlFrontSlots = 0 : mlLeftSlots = 0 : mlRearSlots = 0 : mlRightSlots = 0 : mlAllArcSlots = 0
            For Y = 0 To moSlots.GetUpperBound(1)
                For X = 0 To moSlots.GetUpperBound(0)
                    Select Case moSlots(X, Y).yType
                        Case SlotType.eAllArc
                            mlAllArcSlots += 1
                        Case SlotType.eFront
                            mlFrontSlots += 1
                        Case SlotType.eLeft
                            mlLeftSlots += 1
                        Case SlotType.eRear
                            mlRearSlots += 1
                        Case SlotType.eRight
                            mlRightSlots += 1
                    End Select
                Next X
            Next Y
        End If

    End Sub

    Public Function GetHullSummary(ByVal lHullSize As Int32) As String
        LastHullSize = lHullSize
        Dim X As Int32
        Dim Y As Int32
        Dim lIdx As Int32

        Dim vpTemp() As ValPair = Nothing
        Dim lUB As Int32
        Dim bFound As Boolean

        Dim lTmpGrp As Int32

        lUB = -1

        For X = 0 To ml_GRID_SIZE_WH - 1
            For Y = 0 To ml_GRID_SIZE_WH - 1
                If moSlots(X, Y).yType <> SlotType.eUnused Then

                    If moSlots(X, Y).lConfig = SlotConfig.eArmorConfig Then
                        lTmpGrp = moSlots(X, Y).yType
                    Else : lTmpGrp = moSlots(X, Y).lGroupNum
                    End If

                    bFound = False
                    For lIdx = 0 To lUB
                        If vpTemp(lIdx).ValueConfig = moSlots(X, Y).lConfig AndAlso vpTemp(lIdx).ValueGroupNum = lTmpGrp Then
                            vpTemp(lIdx).Value += 1
                            bFound = True
                            Exit For
                        End If
                    Next lIdx

                    If bFound = False Then
                        lUB = lUB + 1
                        ReDim Preserve vpTemp(0 To lUB)
                        vpTemp(lUB).ValueConfig = moSlots(X, Y).lConfig
                        vpTemp(lUB).ValueGroupNum = lTmpGrp
                        vpTemp(lUB).Value = 1
                    End If

                End If
            Next Y
        Next X

        'Now, clear our string
        Dim oSB As New System.Text.StringBuilder

        Dim sArmorFront As String = ""
        Dim sArmorLeft As String = ""
        Dim sArmorRight As String = ""
        Dim sArmorRear As String = ""
        Dim sArmorAllArc As String = ""

        Dim fHullPerSlot As Single
        If TotalSlots = 0 Then fHullPerSlot = lHullSize Else fHullPerSlot = CSng(lHullSize / TotalSlots)

        'And add our values... first, get the armor values...
        For X = 0 To lUB
            Select Case vpTemp(X).ValueConfig
                Case SlotConfig.eArmorConfig
                    Select Case vpTemp(X).ValueGroupNum
                        Case SlotType.eAllArc
                            sArmorAllArc = "Unallocated All-Arc: " & (vpTemp(X).Value * fHullPerSlot).ToString("#,###.0")
                        Case SlotType.eFront
                            sArmorFront = "Armor Front Side: " & (vpTemp(X).Value * fHullPerSlot).ToString("#,###.0")
                        Case SlotType.eLeft
                            sArmorLeft = "Armor Left Side: " & (vpTemp(X).Value * fHullPerSlot).ToString("#,###.0")
                        Case SlotType.eRight
                            sArmorRight = "Armor Right Side: " & (vpTemp(X).Value * fHullPerSlot).ToString("#,###.0")
                        Case SlotType.eRear
                            sArmorRear = "Armor Rear Side: " & (vpTemp(X).Value * fHullPerSlot).ToString("#,###.0")
                    End Select
            End Select
        Next X



        'Now, put the armor values in the list box
        oSB.AppendLine(sArmorFront)
        oSB.AppendLine(sArmorLeft)
        oSB.AppendLine(sArmorRight)
        oSB.AppendLine(sArmorRear)
        If sArmorAllArc <> "" Then oSB.AppendLine(sArmorAllArc)

        'Add all non-weapons..
        For X = 0 To lUB
            Select Case vpTemp(X).ValueConfig
                Case SlotConfig.eCargoBay
                    Dim fVal As Single = (vpTemp(X).Value * fHullPerSlot)
                    If fVal - CInt(fVal) < 0 Then fVal -= 0.1F
                    oSB.AppendLine("Cargo Bay: " & fVal.ToString("#,###.0"))
                Case SlotConfig.eCrewQuarters
                    Dim fVal As Single = (vpTemp(X).Value * fHullPerSlot)
                    If fVal - CInt(fVal) < 0 Then fVal -= 0.1F
                    oSB.AppendLine("Crew Quarters: " & fVal.ToString("#,###.0"))
                Case SlotConfig.eEngines
                    Dim fVal As Single = (vpTemp(X).Value * fHullPerSlot)
                    If fVal - CInt(fVal) < 0 Then fVal -= 0.1F
                    oSB.AppendLine("Engines: " & fVal.ToString("#,###.0"))
                Case SlotConfig.eFuelBay
                    Dim fVal As Single = (vpTemp(X).Value * fHullPerSlot)
                    If fVal - CInt(fVal) < 0 Then fVal -= 0.1F
                    oSB.AppendLine("Fuel Bay: " & fVal.ToString("#,###.0"))
                Case SlotConfig.eHangar
                    Dim fVal As Single = (vpTemp(X).Value * fHullPerSlot)
                    If fVal - CInt(fVal) < 0 Then fVal -= 0.1F
                    oSB.AppendLine("Hangar: " & fVal.ToString("#,###.0"))
                Case SlotConfig.eRadar
                    Dim fVal As Single = (vpTemp(X).Value * fHullPerSlot)
                    If fVal - CInt(fVal) < 0 Then fVal -= 0.1F
                    oSB.AppendLine("Radar/Comm: " & fVal.ToString("#,###.0"))
                Case SlotConfig.eShields
                    Dim fVal As Single = (vpTemp(X).Value * fHullPerSlot)
                    If fVal - CInt(fVal) < 0 Then fVal -= 0.1F
                    oSB.AppendLine("Shields: " & fVal.ToString("#,###.0"))
            End Select
        Next X

        'Do a quick sort on Hangar Bay Doors
        Dim lSortIdx(-1) As Int32
        For X = 0 To lUB
            If vpTemp(X).ValueConfig = SlotConfig.eHangarDoor Then
                lIdx = -1
                'Ok, ascending, find where this group num is less then lHDIdx
                For lTmpIdx As Int32 = 0 To lSortIdx.GetUpperBound(0)
                    If vpTemp(lSortIdx(lTmpIdx)).ValueGroupNum > vpTemp(X).ValueGroupNum Then
                        'Ok, we go here and shift all values up
                        lIdx = lTmpIdx
                        ReDim Preserve lSortIdx(lSortIdx.GetUpperBound(0) + 1)
                        For lShiftIdx As Int32 = lSortIdx.GetUpperBound(0) - 1 To lTmpIdx Step -1
                            lSortIdx(lShiftIdx + 1) = lSortIdx(lShiftIdx)
                        Next lShiftIdx
                        lSortIdx(lIdx) = X
                        Exit For
                    End If
                Next lTmpIdx

                If lIdx = -1 Then
                    ReDim Preserve lSortIdx(lSortIdx.GetUpperBound(0) + 1)
                    lSortIdx(lSortIdx.GetUpperBound(0)) = X
                End If
            End If
        Next X
        For X = 0 To lSortIdx.GetUpperBound(0)
            With vpTemp(lSortIdx(X))
                oSB.AppendLine("Hangar Bay Door " & .ValueGroupNum.ToString & ": " & (.Value * fHullPerSlot).ToString("#,###.0"))
            End With
        Next X

        'Do a quick sort on Weapons
        ReDim lSortIdx(-1)
        For X = 0 To lUB
            If vpTemp(X).ValueConfig = SlotConfig.eWeapons Then
                lIdx = -1
                'Ok, ascending, find where this group num is less then lHDIdx
                For lTmpIdx As Int32 = 0 To lSortIdx.GetUpperBound(0)
                    If vpTemp(lSortIdx(lTmpIdx)).ValueGroupNum > vpTemp(X).ValueGroupNum Then
                        'Ok, we go here and shift all values up
                        lIdx = lTmpIdx
                        ReDim Preserve lSortIdx(lSortIdx.GetUpperBound(0) + 1)
                        For lShiftIdx As Int32 = lSortIdx.GetUpperBound(0) - 1 To lTmpIdx Step -1
                            lSortIdx(lShiftIdx + 1) = lSortIdx(lShiftIdx)
                        Next lShiftIdx
                        lSortIdx(lIdx) = X
                        Exit For
                    End If
                Next lTmpIdx

                If lIdx = -1 Then
                    ReDim Preserve lSortIdx(lSortIdx.GetUpperBound(0) + 1)
                    lSortIdx(lSortIdx.GetUpperBound(0)) = X
                End If
            End If
        Next
        'Now, add weapons last
        For X = 0 To lSortIdx.GetUpperBound(0)
            With vpTemp(lSortIdx(X))
                Dim ySlotType As SlotType = GetWeaponsMainSlotType(.ValueGroupNum)
                Dim sSide As String = ""
                Select Case ySlotType
                    Case SlotType.eAllArc
                        sSide = "All-Arc"
                    Case SlotType.eFront
                        sSide = "Forward"
                    Case SlotType.eLeft
                        sSide = "Left"
                    Case SlotType.eRear
                        sSide = "Rear"
                    Case SlotType.eRight
                        sSide = "Right"
                    Case Else
                        sSide = "Unknown"
                End Select
                oSB.AppendLine("Weapon Group " & .ValueGroupNum.ToString & " (" & sSide & "): " & (.Value * fHullPerSlot).ToString("#,###.0#"))
            End With
        Next X

        Return oSB.ToString
    End Function


    Public Function GetSlotGrpNum(ByVal X As Int32, ByVal Y As Int32) As Int32
        Return moSlots(X, Y).lGroupNum
    End Function
    Public Function GetWeaponsMainSlotType(ByVal lWpnGrpNum As Int32) As SlotType

        Dim bOnEdge() As Boolean = {False, False, False, False}
        Dim lSideCnt() As Int32 = {0, 0, 0, 0}

        For Y As Int32 = 0 To ml_GRID_SIZE_WH - 1
            For X As Int32 = 0 To ml_GRID_SIZE_WH - 1
                With moSlots(X, Y)
                    If .yType <> SlotType.eNoChange AndAlso .yType <> SlotType.eUnused Then
                        If .lConfig = SlotConfig.eWeapons AndAlso .lGroupNum = lWpnGrpNum Then
                            'ok, determine what we are
                            If .yType = SlotType.eAllArc Then Return SlotType.eAllArc

                            Select Case .yType
                                Case SlotType.eLeft
                                    lSideCnt(UnitArcs.eLeftArc) += 1
                                    If bOnEdge(UnitArcs.eLeftArc) = False Then bOnEdge(UnitArcs.eLeftArc) = SlotIsOnEdge(X, Y)
                                Case SlotType.eRear
                                    lSideCnt(UnitArcs.eBackArc) += 1
                                    If bOnEdge(UnitArcs.eBackArc) = False Then bOnEdge(UnitArcs.eBackArc) = SlotIsOnEdge(X, Y)
                                Case SlotType.eRight
                                    lSideCnt(UnitArcs.eRightArc) += 1
                                    If bOnEdge(UnitArcs.eRightArc) = False Then bOnEdge(UnitArcs.eRightArc) = SlotIsOnEdge(X, Y)
                                Case SlotType.eFront
                                    lSideCnt(UnitArcs.eForwardArc) += 1
                                    If bOnEdge(UnitArcs.eForwardArc) = False Then bOnEdge(UnitArcs.eForwardArc) = SlotIsOnEdge(X, Y)
                            End Select
                        End If
                    End If
                End With
            Next X
        Next Y

        Dim lMax As Int32 = 0
        Dim lSideVal As Int32 = -1
        For X As Int32 = 0 To 3
            If bOnEdge(X) = True Then
                If lSideCnt(X) > lMax Then
                    lMax = lSideCnt(X)
                    lSideVal = X
                End If
            End If
        Next X

        'did we end up with a side?
        If lSideVal <> -1 Then
            Select Case lSideVal
                Case UnitArcs.eLeftArc
                    Return SlotType.eLeft
                Case UnitArcs.eForwardArc
                    Return SlotType.eFront
                Case UnitArcs.eRightArc
                    Return SlotType.eRight
                Case UnitArcs.eBackArc
                    Return SlotType.eRear
                Case Else
                    Return SlotType.eUnused
            End Select
        Else
            Return SlotType.eUnused
        End If

    End Function

    Public Function GetApproxPowerRequired(ByVal lFrameTypeID As Int32, ByVal lHullSize As Int32, ByVal lDensity As Int32) As Int32
        'Get power requirement for cargo cap
        Dim lCargoCap As Int32 = 0
        'Now, go through all the doors
        Dim lGroup() As Int32 = Nothing
        Dim lSize() As Int32 = Nothing
        Dim lGroupUB As Int32 = -1
        Dim lTotal As Int32

        'Load our door groups
        For Y As Int32 = 0 To ml_GRID_SIZE_WH - 1
            For X As Int32 = 0 To ml_GRID_SIZE_WH - 1
                If moSlots(X, Y).yType <> SlotType.eUnused Then
                    lTotal += 1

                    'Ok, now, is this a HangarDoor?
                    If moSlots(X, Y).lConfig = SlotConfig.eHangarDoor Then
                        'Ok, it is... find our group
                        Dim lIdx As Int32 = -1
                        For lTempIdx As Int32 = 0 To lGroupUB
                            If lGroup(lTempIdx) = moSlots(X, Y).lGroupNum Then
                                lIdx = lTempIdx
                                Exit For
                            End If
                        Next lTempIdx

                        If lIdx = -1 Then
                            lGroupUB += 1
                            ReDim Preserve lGroup(lGroupUB)
                            ReDim Preserve lSize(lGroupUB)
                            lIdx = lGroupUB
                        End If

                        lGroup(lIdx) = moSlots(X, Y).lGroupNum
                        lSize(lIdx) += 1
                    End If
                End If
            Next X
        Next Y

        Dim fHullPerSlot As Single = CSng(lHullSize / lTotal)

        If lFrameTypeID = 0 OrElse lFrameTypeID = 4 Then
            lCargoCap = CInt(GetSlotConfigCnt(SlotType.eNoChange, SlotConfig.eCargoBay) * fHullPerSlot)
            If lCargoCap <> 0 Then
                lCargoCap = CInt(Math.Ceiling(Math.Sqrt(lCargoCap) * 1.4F))
            End If
        ElseIf lFrameTypeID = 3 Then
            lCargoCap = CInt(GetSlotConfigCnt(SlotType.eNoChange, SlotConfig.eCargoBay) * fHullPerSlot)
            If lCargoCap <> 0 Then
                lCargoCap = CInt(Math.Ceiling(Math.Sqrt(lCargoCap) * 1.1F))
            End If
        End If

        Dim mlPowerRequired As Int32 = 0

        'Get power requirement for hangar
        Dim lHangarCap As Int32 = CInt(GetSlotConfigCnt(SlotType.eNoChange, SlotConfig.eHangar) * fHullPerSlot) 'GetHullAllocation(SlotConfig.eHangar)
        If lHangarCap = 0 Then
            mlPowerRequired = CInt(lCargoCap * 1.4F)
            Return mlPowerRequired
        Else : lHangarCap = CInt(Math.Ceiling(Math.Sqrt(lHangarCap) * 0.8F))
        End If

        'Get our HullPerSlot
        Dim lDoorPower As Int32 = 0
        Dim lStructDens As Int32 = lDensity * 25 'StructureMineral.GetPropertyValue(eMinPropID.Density)

        For lIdx As Int32 = 0 To lGroupUB
            If lSize(lIdx) <> 0 Then
                'Now, go through the doors and figure out their power
                Dim fScore As Single = ((lSize(lIdx) * (lStructDens / 256.0F)) * lSize(lIdx)) / 2.0F
                lDoorPower += CInt(Math.Ceiling(fScore / 100.0F))
            End If
        Next lIdx

        'mlPowerRequired = CInt((lCargoCap + lHangarCap + lDoorPower)) ' * (RandomSeed + 0.4F))
        Return CInt((lCargoCap + lHangarCap + lDoorPower) * 1.4F)
    End Function

    Public Sub SetFromHullTech(ByRef oHullTech As HullTech)
        ClearSelectingSlotConfig()
        If oHullTech Is Nothing = False Then
            For Y As Int32 = 0 To ml_GRID_SIZE_WH - 1
                For X As Int32 = 0 To ml_GRID_SIZE_WH - 1
                    moSlots(X, Y).yType = oHullTech.moSlots(X, Y).yType
                    moSlots(X, Y).lConfig = oHullTech.moSlots(X, Y).lConfig
                    moSlots(X, Y).lGroupNum = oHullTech.moSlots(X, Y).lGroupNum
                Next X
            Next Y
        Else
            For Y As Int32 = 0 To ml_GRID_SIZE_WH - 1
                For X As Int32 = 0 To ml_GRID_SIZE_WH - 1
                    moSlots(X, Y).yType = SlotType.eUnused
                    moSlots(X, Y).lConfig = SlotConfig.eArmorConfig
                Next X
            Next Y
        End If
        LastHullSize = oHullTech.HullSize
        Me.IsDirty = True
    End Sub

    ''' <summary>
    ''' Gets errors generated by the current hull design (used in frmHullBuilder)
    ''' Placed here for efficiency... returns whether the Error string, check the string for the word "ERROR"
    ''' </summary>
    ''' <param name="lTotalHull"> The Hull Size for the current object </param>
    ''' <param name="lCrewReqVal"> The Crew Requirement Percentage (as an Int32 = X * 100) for the object </param>
    ''' <param name="yHullReq"> Indicates the flags for the specific hull that need to be validated (enum of eyHullRequirements)</param>
    ''' <remarks></remarks>
    Public Function HasErrorList(ByVal lTotalHull As Int32, ByVal lCrewReqVal As Int32, ByVal yHullReq As eyHullRequirements) As String
        'RULES:
        'Should Contain an Engine (warning)
        'If Engine present, should contain a Fuel Bay (warning)
        'Armor should exist on all sides (warning)
        'Cross-Arc groups throws a (warning)
        'Armor cannot exist in an all-arc slot
        'Crew Quarters must equal or exceed the TotalHull * (lCrewReqVal / 100)
        'If bShip = True then Engines MUST have a least 1 EDGE REAR slot
        'Containers except for Crew Quarters must be touching the same container/group
        'Every weapon group must have at least 1 EDGE Slot... all EDGE slots must reside in the same arc

        Dim bHasEngine As Boolean = False
        'Dim bHasFuel As Boolean = False
        Dim bSideArmor(4) As Boolean
        Dim lCrewQuarters As Int32 = 0
        Dim bInsuffCrewQuarters As Boolean = False
        Dim fHullPerSlot As Single = CSng(lTotalHull / Me.TotalSlots)
        Dim bEngineOnRearEdge As Boolean = False
        Dim bBreakout As Boolean

        Dim bHasBayDoorInAllArc As Boolean = False
        Dim bHasBayDoors As Boolean = False
        Dim bHasHangar As Boolean = False

        Dim oSB As New System.Text.StringBuilder

        For Y As Int32 = 0 To ml_GRID_SIZE_WH - 1
            For X As Int32 = 0 To ml_GRID_SIZE_WH - 1
                With moSlots(X, Y)
                    If .yType <> SlotType.eUnused Then
                        Select Case .lConfig
                            Case SlotConfig.eArmorConfig
                                bSideArmor(.yType - 1) = True
                                lCrewQuarters += 1
                            Case SlotConfig.eEngines
                                bHasEngine = True

                                If Y < ml_GRID_SIZE_WH - 1 AndAlso .yType = SlotType.eRear Then
                                    If moSlots(X, Y + 1).yType = SlotType.eUnused Then
                                        bEngineOnRearEdge = True
                                    End If
                                ElseIf Y = ml_GRID_SIZE_WH - 1 AndAlso .yType = SlotType.eRear Then
                                    bEngineOnRearEdge = True
                                End If
                                'Case SlotConfig.eFuelBay
                                '    bHasFuel = True
                                'Case SlotConfig.eCrewQuarters
                                '    lCrewQuarters += 1
                            Case SlotConfig.eHangarDoor
                                bHasBayDoors = True
                                If .yType = SlotType.eAllArc Then bHasBayDoorInAllArc = True
                            Case SlotConfig.eHangar
                                bHasHangar = True
                        End Select
                    End If
                End With
            Next X
        Next Y

        'Set Insufficient Crew Quarters...
        bInsuffCrewQuarters = False '(lCrewQuarters * fHullPerSlot) < ((lCrewReqVal / 100.0F) * lTotalHull)

        'Check for Cross-Arc Groups...
        'TODO: if an item is in an "all-arc" slot it should not present a warning when the item is placed in other arcs...
        'if a weapon is in an "all-arc" slot, it is not permitted to be in any other slot
        Dim sCrossArcGroups() As String = Nothing
        Dim lCAUB As Int32 = -1
        Dim lWpnGrpAllArcCross(-1) As Int32
        For Y As Int32 = 0 To ml_GRID_SIZE_WH - 1
            For X As Int32 = 0 To ml_GRID_SIZE_WH - 1
                With moSlots(X, Y)
                    If .yType <> SlotType.eUnused AndAlso .lConfig <> SlotConfig.eArmorConfig AndAlso .lConfig <> SlotConfig.eCrewQuarters Then
                        'Now... go through the array again
                        bBreakout = False
                        For SubY As Int32 = 0 To ml_GRID_SIZE_WH - 1
                            For SubX As Int32 = 0 To ml_GRID_SIZE_WH - 1
                                If X <> SubX OrElse Y <> SubY Then
                                    If moSlots(SubX, SubY).lConfig = .lConfig AndAlso moSlots(SubX, SubY).lGroupNum = .lGroupNum Then
                                        If moSlots(SubX, SubY).yType <> .yType Then

                                            If .yType = SlotType.eAllArc OrElse moSlots(SubX, SubY).yType = SlotType.eAllArc Then
                                                If .lConfig = SlotConfig.eWeapons Then
                                                    Dim bFoundWpnGrpAllArcCross As Boolean = False
                                                    For f As Int32 = 0 To lWpnGrpAllArcCross.GetUpperBound(0)
                                                        If lWpnGrpAllArcCross(f) = .lGroupNum Then
                                                            bFoundWpnGrpAllArcCross = True
                                                            Exit For
                                                        End If
                                                    Next f
                                                    If bFoundWpnGrpAllArcCross = False Then
                                                        oSB.AppendLine("ERROR - Weapon Group " & .lGroupNum & " cannot be in an all-arc slot and cross arcs.")
                                                        ReDim Preserve lWpnGrpAllArcCross(lWpnGrpAllArcCross.GetUpperBound(0) + 1)
                                                        lWpnGrpAllArcCross(lWpnGrpAllArcCross.GetUpperBound(0)) = .lGroupNum
                                                    End If
                                                End If
                                            End If

                                            'Ok, cross-arc... see if we already know about it
                                            Dim bFound As Boolean = False
                                            Dim sTemp As String = CInt(.lConfig).ToString & "_" & CInt(.lGroupNum).ToString

                                            For lIdx As Int32 = 0 To lCAUB
                                                If sCrossArcGroups(lIdx) = sTemp Then
                                                    bFound = True
                                                    Exit For
                                                End If
                                            Next lIdx

                                            If bFound = False Then
                                                lCAUB += 1
                                                ReDim Preserve sCrossArcGroups(lCAUB)
                                                sCrossArcGroups(lCAUB) = sTemp
                                            End If

                                            bBreakout = True
                                            Exit For
                                        End If
                                    End If
                                End If
                            Next SubX
                            If bBreakout = True Then Exit For
                        Next SubY
                    End If
                End With
            Next X
        Next Y

        'Containers except for Crew Quarters, Armor and cargo bay must be touching the same container/group
        Dim vpVals() As ValPair = Nothing
        Dim lValUB As Int32 = -1
        For Y As Int32 = 0 To ml_GRID_SIZE_WH - 1
            For X As Int32 = 0 To ml_GRID_SIZE_WH - 1
                With moSlots(X, Y)
                    If .yType <> SlotType.eUnused AndAlso .lConfig <> SlotConfig.eArmorConfig AndAlso .lConfig <> SlotConfig.eCrewQuarters AndAlso .lConfig <> SlotConfig.eCargoBay Then
                        Dim bFound As Boolean = False
                        For lIdx As Int32 = 0 To lValUB
                            If vpVals(lIdx).ValueConfig = .lConfig AndAlso vpVals(lIdx).ValueGroupNum = .lGroupNum Then
                                bFound = True
                                Exit For
                            End If
                        Next lIdx
                        If bFound = False Then
                            lValUB += 1
                            ReDim Preserve vpVals(lValUB)
                            vpVals(lValUB).ValueConfig = .lConfig
                            vpVals(lValUB).ValueGroupNum = .lGroupNum
                            vpVals(lValUB).Value = 0
                        End If
                        .bVerified = False
                    End If
                End With
            Next X
        Next Y
        'Now, go back through our values
        For lIdx As Int32 = 0 To lValUB
            Dim bStarted As Boolean = False

            For Y As Int32 = 0 To ml_GRID_SIZE_WH - 1
                For X As Int32 = 0 To ml_GRID_SIZE_WH - 1
                    Dim bDecrementY As Boolean = False
                    Dim bDecrementX As Boolean = False

                    With moSlots(X, Y)
                        If .yType <> SlotType.eUnused AndAlso .lConfig = vpVals(lIdx).ValueConfig AndAlso .lGroupNum = vpVals(lIdx).ValueGroupNum Then
                            'Is this the first one?
                            If bStarted = False Then
                                bStarted = True
                                'Go ahead and verify this one...
                                .bVerified = True
                            End If

                            'Is this square verified?
                            If .bVerified = True Then
                                'yes, verify its neighbors
                                If Y <> 0 AndAlso moSlots(X, Y - 1).lConfig = .lConfig AndAlso moSlots(X, Y - 1).lGroupNum = .lGroupNum Then
                                    If moSlots(X, Y - 1).bVerified = False Then bDecrementY = True
                                    moSlots(X, Y - 1).bVerified = True
                                End If
                                If Y <> ml_GRID_SIZE_WH - 1 AndAlso moSlots(X, Y + 1).lConfig = .lConfig AndAlso moSlots(X, Y + 1).lGroupNum = .lGroupNum Then
                                    moSlots(X, Y + 1).bVerified = True
                                End If
                                If X <> 0 AndAlso moSlots(X - 1, Y).lConfig = .lConfig AndAlso moSlots(X - 1, Y).lGroupNum = .lGroupNum Then
                                    If moSlots(X - 1, Y).bVerified = False Then bDecrementX = True
                                    moSlots(X - 1, Y).bVerified = True
                                End If
                                If X <> ml_GRID_SIZE_WH - 1 AndAlso moSlots(X + 1, Y).lConfig = .lConfig AndAlso moSlots(X + 1, Y).lGroupNum = .lGroupNum Then
                                    moSlots(X + 1, Y).bVerified = True
                                End If
                            End If
                        End If
                    End With

                    If bDecrementY = True Then
                        Y = Y - 2
                        Exit For
                    ElseIf bDecrementX = True Then
                        X = X - 2
                    End If
                Next X
            Next Y
        Next lIdx

        'Now... go through and ensure all squares are valid
        For Y As Int32 = 0 To ml_GRID_SIZE_WH - 1
            For X As Int32 = 0 To ml_GRID_SIZE_WH - 1
                With moSlots(X, Y)
                    If .bVerified = False AndAlso .yType <> SlotType.eUnused AndAlso .lConfig <> SlotConfig.eCargoBay AndAlso .lConfig <> SlotConfig.eArmorConfig AndAlso .lConfig <> SlotConfig.eCrewQuarters Then
                        'Ok... got one... so find it in vpVals
                        For lIdx As Int32 = 0 To lValUB
                            If vpVals(lIdx).ValueConfig = .lConfig AndAlso vpVals(lIdx).ValueGroupNum = .lGroupNum Then
                                vpVals(lIdx).Value = 1
                                Exit For
                            End If
                        Next lIdx
                    End If
                End With
            Next X
        Next Y

        'Now, do one more loop through the vpVals array
        For lIdx As Int32 = 0 To lValUB
            If vpVals(lIdx).Value = 1 Then
                'Had an error in it...
                Select Case vpVals(lIdx).ValueConfig
                    Case SlotConfig.eCargoBay
                        'oSB.AppendLine("ERROR - Cargo Bay slots must be grouped together")
                    Case SlotConfig.eEngines
                        oSB.AppendLine("ERROR - Engine slots must be grouped together")
                        'Case SlotConfig.eFuelBay
                        '    oSB.AppendLine("ERROR - Fuel Bay slots must be grouped together")
                    Case SlotConfig.eHangar
                        oSB.AppendLine("ERROR - Hangar slots must be grouped together")
                    Case SlotConfig.eRadar
                        oSB.AppendLine("ERROR - Radar/Comm slots must be grouped together")
                    Case SlotConfig.eShields
                        oSB.AppendLine("ERROR - Shield slots must be grouped together")
                    Case SlotConfig.eWeapons
                        oSB.AppendLine("ERROR - Weapon " & vpVals(lIdx).ValueGroupNum & " slots must be grouped together")
                    Case SlotConfig.eHangarDoor
                        oSB.AppendLine("ERROR - Hangar Bay Door " & vpVals(lIdx).ValueGroupNum & " slots must be grouped together")
                End Select
            End If
        Next lIdx

        'Check Bay Door Edges
        Dim lBayGroups() As Int32 = Nothing
        Dim lBayGroupUB As Int32 = -1
        Dim bHasNoEdge() As Boolean = Nothing
        For Y As Int32 = 0 To ml_GRID_SIZE_WH - 1
            For X As Int32 = 0 To ml_GRID_SIZE_WH - 1
                With moSlots(X, Y)
                    If .yType <> SlotType.eUnused AndAlso .lConfig = SlotConfig.eHangarDoor Then
                        Dim lIdx As Int32 = -1
                        For lTempIdx As Int32 = 0 To lBayGroupUB
                            If lBayGroups(lTempIdx) = .lGroupNum Then
                                lIdx = lTempIdx
                                Exit For
                            End If
                        Next lTempIdx
                        If lIdx = -1 Then
                            lBayGroupUB += 1
                            ReDim Preserve lBayGroups(lBayGroupUB)
                            ReDim Preserve bHasNoEdge(lBayGroupUB)
                            lBayGroups(lBayGroupUB) = .lGroupNum
                            bHasNoEdge(lBayGroupUB) = False
                            lIdx = lBayGroupUB
                        End If

                        Dim bHasEdge As Boolean = False
                        If X > 0 AndAlso moSlots(X - 1, Y).yType = SlotType.eUnused Then
                            bHasEdge = True
                        ElseIf X < ml_GRID_SIZE_WH - 1 AndAlso moSlots(X + 1, Y).yType = SlotType.eUnused Then
                            bHasEdge = True
                        ElseIf Y > 0 AndAlso moSlots(X, Y - 1).yType = SlotType.eUnused Then
                            bHasEdge = True
                        ElseIf Y < ml_GRID_SIZE_WH - 1 AndAlso moSlots(X, Y + 1).yType = SlotType.eUnused Then
                            bHasEdge = True
                        End If

                        'now, if there is a slot that does not have an edge, we keep bHasNoEdge marked true
                        '  otherwise, if this slot does not have an edge, we mark bHasNoEdge as true
                        bHasNoEdge(lIdx) = bHasNoEdge(lIdx) OrElse (bHasEdge = False)

                    End If
                End With
            Next X
        Next Y

        If lBayGroupUB = -1 AndAlso (yHullReq And eyHullRequirements.RequiresBayDoor) <> 0 Then
            oSB.AppendLine("ERROR - Must contain at least one Hangar Bay Door!")
        End If
        If bHasEngine = False AndAlso (yHullReq And eyHullRequirements.RequiresAnEngine) <> 0 Then
            oSB.AppendLine("ERROR - Must contain room for an Engine (for power)!")
        End If

        'Now, check the edges
        For X As Int32 = 0 To lBayGroupUB
            If bHasNoEdge(X) = True Then
                oSB.AppendLine("ERROR - Hangar Bay Door " & lBayGroups(X) & " can only be placed on edge slots.")
            End If
        Next X
        'Now, check hangar door touches a hangar
        For lIdx As Int32 = 0 To lBayGroupUB
            Dim bFound As Boolean = False

            For Y As Int32 = 0 To ml_GRID_SIZE_WH - 1
                For X As Int32 = 0 To ml_GRID_SIZE_WH - 1
                    With moSlots(X, Y)
                        If .yType <> SlotType.eUnused AndAlso .lConfig = SlotConfig.eHangarDoor AndAlso .lGroupNum = lBayGroups(lIdx) Then
                            'check the neighbors
                            If X <> 0 AndAlso moSlots(X - 1, Y).lConfig = SlotConfig.eHangar Then
                                bFound = True
                            ElseIf X <> ml_GRID_SIZE_WH - 1 AndAlso moSlots(X + 1, Y).lConfig = SlotConfig.eHangar Then
                                bFound = True
                            ElseIf Y <> 0 AndAlso moSlots(X, Y - 1).lConfig = SlotConfig.eHangar Then
                                bFound = True
                            ElseIf Y <> ml_GRID_SIZE_WH - 1 AndAlso moSlots(X, Y + 1).lConfig = SlotConfig.eHangar Then
                                bFound = True
                            End If
                        End If
                    End With

                    If bFound = True Then Exit For
                Next X
                If bFound = True Then Exit For
            Next Y

            If bFound = False Then
                oSB.AppendLine("ERROR - Hangar Bay Door " & lBayGroups(lIdx) & " must touch the Hangar!")
            End If
        Next lIdx

        'Check Weapon Group Edges...
        Dim lWpnGroups() As Int32 = Nothing
        Dim lWpnGroupUB As Int32 = -1
        For Y As Int32 = 0 To ml_GRID_SIZE_WH - 1
            For X As Int32 = 0 To ml_GRID_SIZE_WH - 1
                With moSlots(X, Y)
                    If .yType <> SlotType.eUnused AndAlso .lConfig = SlotConfig.eWeapons Then
                        Dim bFound As Boolean = False
                        For lIdx As Int32 = 0 To lWpnGroupUB
                            If lWpnGroups(lIdx) = .lGroupNum Then
                                bFound = True
                                Exit For
                            End If
                        Next lIdx
                        If bFound = False Then
                            lWpnGroupUB += 1
                            ReDim Preserve lWpnGroups(lWpnGroupUB)
                            lWpnGroups(lWpnGroupUB) = .lGroupNum
                        End If
                    End If
                End With
            Next X
        Next Y
        Dim lWpnGroupEdges(lWpnGroupUB) As Int32
        For lIdx As Int32 = 0 To lWpnGroupUB
            'bBreakout = False
            Dim bLeftArcWpn As Boolean = False
            Dim bForeArcWpn As Boolean = False
            Dim bRightArcWpn As Boolean = False
            Dim bRearArcWpn As Boolean = False
            Dim bAllArcWpn As Boolean = False

            For Y As Int32 = 0 To ml_GRID_SIZE_WH - 1
                For X As Int32 = 0 To ml_GRID_SIZE_WH - 1
                    With moSlots(X, Y)
                        If .yType <> SlotType.eUnused AndAlso .lConfig = SlotConfig.eWeapons AndAlso lWpnGroups(lIdx) = .lGroupNum Then
                            'ok... now... is this an edge???
                            If (X > 0 AndAlso moSlots(X - 1, Y).yType = SlotType.eUnused) OrElse (X = 0) Then bLeftArcWpn = True
                            If (X < ml_GRID_SIZE_WH - 1 AndAlso moSlots(X + 1, Y).yType = SlotType.eUnused) OrElse (X = ml_GRID_SIZE_WH - 1) Then bRightArcWpn = True
                            If (Y > 0 AndAlso moSlots(X, Y - 1).yType = SlotType.eUnused) OrElse (Y = 0) Then bForeArcWpn = True
                            If (Y < ml_GRID_SIZE_WH - 1 AndAlso moSlots(X, Y + 1).yType = SlotType.eUnused) OrElse (Y = ml_GRID_SIZE_WH - 1) Then bRearArcWpn = True
                            If .yType = SlotType.eAllArc Then bAllArcWpn = True
                        End If
                    End With
                Next X
            Next Y

            If bLeftArcWpn = False AndAlso bForeArcWpn = False AndAlso bRightArcWpn = False AndAlso bRearArcWpn = False AndAlso bAllArcWpn = False Then
                'If (bLeftArcWpn Xor bForeArcWpn Xor bRightArcWpn Xor bRearArcWpn) = False AndAlso bAllArcWpn = False Then
                oSB.AppendLine("ERROR - Weapon " & lWpnGroups(lIdx) & " must touch one edge side")
            End If
        Next lIdx

        'Now, fill our lstVals listbox... warnings are always lower priority
        'If bSideArmor(4) = True Then oSB.AppendLine("ERROR - Armor and Crew cannot be allocated to All-Arc Slots")
        If bSideArmor(4) = True Then oSB.AppendLine("ERROR - All-Arc (Yellow) Slots cannot be left blank!")
        If bInsuffCrewQuarters = True Then
            oSB.AppendLine("ERROR - Insufficient allocation for Crew Quarters.")
            oSB.AppendLine("  Requires " & lCrewReqVal & "% of the hull.")
        End If
        If (yHullReq And eyHullRequirements.EngineIsToBeInRear) <> 0 AndAlso bEngineOnRearEdge = False Then
            oSB.AppendLine("ERROR - One Rear Edge Slot must be for Engines")
        End If
        If bHasBayDoorInAllArc = True Then
            oSB.AppendLine("ERROR - Hangar Bay Doors cannot be placed in All-Arc Slots!")
        End If

        'If bHasEngine = False AndAlso bLandBasedFac = False Then oSB.AppendLine("WARNING - Should contain slots for Engines")
        'If bLandBasedFac = False AndAlso bHasEngine = True AndAlso bHasFuel = False Then oSB.AppendLine("WARNING - Should contain slots for a Fuel Bay")
        If bSideArmor(0) = False Then oSB.AppendLine("WARNING - Front Arc has no room for Armor")
        If bSideArmor(1) = False Then oSB.AppendLine("WARNING - Left Arc has no room for Armor")
        If bSideArmor(2) = False Then oSB.AppendLine("WARNING - Rear Arc has no room for Armor")
        If bSideArmor(3) = False Then oSB.AppendLine("WARNING - Right Arc has no room for Armor")

        If bSideArmor(0) = False AndAlso bSideArmor(1) = False AndAlso bSideArmor(2) = False AndAlso bSideArmor(3) = False Then
            oSB.AppendLine("WARNING - Prototype may not have enough room for Crew Quarters")
        End If

        For lIdx As Int32 = 0 To lCAUB
            Dim sVals() As String = Split(sCrossArcGroups(lIdx), "_")
            Dim lConfig As Int32 = CInt(Val(sVals(0)))
            Dim lGroupNum As Int32 = CInt(Val(sVals(1)))

            Select Case CType(lConfig, SlotConfig)
                Case SlotConfig.eCargoBay
                    oSB.AppendLine("WARNING - Cargo Bay crosses arcs.")
                    oSB.AppendLine("  Increases the chance of components being damaged")
                Case SlotConfig.eEngines
                    oSB.AppendLine("WARNING - Engine crosses arcs.")
                    oSB.AppendLine("  Increases the chance of components being damaged")
                    'Case SlotConfig.eFuelBay
                    '    oSB.AppendLine("WARNING - Fuel Bay crosses arcs.")
                    '    oSB.AppendLine("  Increases the chance of components being damaged")
                Case SlotConfig.eHangar
                    oSB.AppendLine("WARNING - Hangar crosses arcs.")
                    oSB.AppendLine("  Increases the chance of components being damaged")
                Case SlotConfig.eRadar
                    oSB.AppendLine("WARNING - Radar/Comm crosses arcs.")
                    oSB.AppendLine("  Increases the chance of components being damaged")
                Case SlotConfig.eShields
                    oSB.AppendLine("WARNING - Shield crosses arcs.")
                    oSB.AppendLine("  Increases the chance of components being damaged")
                Case SlotConfig.eWeapons
                    oSB.AppendLine("WARNING - Weapon Group " & lGroupNum.ToString & " crosses arcs.")
                    oSB.AppendLine("  Increases the chance of components being damaged")
                Case SlotConfig.eHangarDoor
                    oSB.AppendLine("WARNING - Hangar Bay Door " & lGroupNum.ToString & " crosses arcs.")
                    oSB.AppendLine("  Increases the chance of components being damaged")
            End Select
        Next lIdx

        If bHasHangar = True AndAlso bHasBayDoors = False Then
            oSB.AppendLine("WARNING - Hangar does not have attached bay doors!")
        End If

        'ok, check if crewquarters is enough
        Dim lHullPerCrew As Int32 = 10
        If goCurrentPlayer Is Nothing = False Then lHullPerCrew = goCurrentPlayer.GetSpecialTechPropertyValue(PlayerSpecialAttributeID.eBetterHullToResidence)

        If (lCrewQuarters * fHullPerSlot) / lHullPerCrew < lTotalHull / 100 Then
            oSB.AppendLine("WARNING - Hull does not have sufficient room for crew." & vbCrLf & "  Components will need to compensate for the crew requirement.")
        End If

        Return oSB.ToString

    End Function

    Public Function GetMsgRdySlotList() As Byte()
        Dim yResp() As Byte
        Dim lPos As Int32 = 0
        Dim lCnt As Int32 = 0

        'Get the count of how many we have to send
        For Y As Int32 = 0 To ml_GRID_SIZE_WH - 1
            For X As Int32 = 0 To ml_GRID_SIZE_WH - 1
                If moSlots(X, Y).yType <> SlotType.eUnused AndAlso moSlots(X, Y).lConfig <> SlotConfig.eArmorConfig Then lCnt += 1
            Next X
        Next Y

        ReDim yResp((4 * lCnt) + 3)

        System.BitConverter.GetBytes(lCnt).CopyTo(yResp, 0) : lPos += 4

        For Y As Int32 = 0 To ml_GRID_SIZE_WH - 1
            For X As Int32 = 0 To ml_GRID_SIZE_WH - 1
                If moSlots(X, Y).yType <> SlotType.eUnused AndAlso moSlots(X, Y).lConfig <> SlotConfig.eArmorConfig Then
                    yResp(lPos) = CByte(X) : lPos += 1
                    yResp(lPos) = CByte(Y) : lPos += 1
                    yResp(lPos) = CByte(moSlots(X, Y).lConfig) : lPos += 1
                    yResp(lPos) = CByte(moSlots(X, Y).lGroupNum) : lPos += 1
                End If
            Next X
        Next Y

        Return yResp
    End Function

#Region "Slot counts"
    Public ReadOnly Property FrontSlots() As Int32
        Get
            Return mlFrontSlots
        End Get
    End Property

    Public ReadOnly Property LeftSlots() As Int32
        Get
            Return mlLeftSlots
        End Get
    End Property

    Public ReadOnly Property RearSlots() As Int32
        Get
            Return mlRearSlots
        End Get
    End Property

    Public ReadOnly Property RightSlots() As Int32
        Get
            Return mlRightSlots
        End Get
    End Property

    Public ReadOnly Property AllArcSlots() As Int32
        Get
            Return mlAllArcSlots
        End Get
    End Property

    Public ReadOnly Property TotalSlots() As Int32
        Get
            If mlFrontSlots = 0 OrElse mlLeftSlots = 0 OrElse mlRightSlots = 0 OrElse mlRearSlots = 0 Then
                mlFrontSlots = 0 : mlLeftSlots = 0 : mlRightSlots = 0 : mlAllArcSlots = 0 : mlRearSlots = 0
                For Y As Int32 = 0 To ml_GRID_SIZE_WH - 1
                    For X As Int32 = 0 To ml_GRID_SIZE_WH - 1
                        Select Case moSlots(X, Y).yType
                            Case SlotType.eAllArc
                                mlAllArcSlots += 1
                            Case SlotType.eFront
                                mlFrontSlots += 1
                            Case SlotType.eLeft
                                mlLeftSlots += 1
                            Case SlotType.eRear
                                mlRearSlots += 1
                            Case SlotType.eRight
                                mlRightSlots += 1
                        End Select
                    Next X
                Next Y
            End If
            Return mlFrontSlots + mlLeftSlots + mlRearSlots + mlRightSlots + mlAllArcSlots
        End Get
    End Property
#End Region

    Protected Overrides Sub Finalize()
        If goResMgr Is Nothing = False Then goResMgr.DeleteTexture("HullBuilderIcons.dds")
        moHullTex = Nothing
        MyBase.Finalize()
    End Sub

    Public Sub SetSelectingSlotConfig(ByVal lConfig As SlotConfig, ByVal lHullSizeMin As Int32, ByVal lHullSize As Int32)
        If lHullSize > 0 Then LastHullSize = lHullSize
        If lHullSizeMin = -1 Then
            For Y As Int32 = 0 To ml_GRID_SIZE_WH - 1
                For X As Int32 = 0 To ml_GRID_SIZE_WH - 1
                    If moSlots(X, Y).lConfig <> lConfig Then
                        moSlots(X, Y).bFiltered = True
                    Else : moSlots(X, Y).bFiltered = False
                    End If
                Next X
            Next Y
            Me.IsDirty = True
        Else
            'Go through and find what matches
            Dim vpVals() As ValPair = Nothing
            Dim lVPUB As Int32 = -1
            Dim bFound As Boolean = False

            For Y As Int32 = 0 To ml_GRID_SIZE_WH - 1
                For X As Int32 = 0 To ml_GRID_SIZE_WH - 1
                    If moSlots(X, Y).lConfig = lConfig Then
                        bFound = False
                        For lIdx As Int32 = 0 To lVPUB
                            If vpVals(lIdx).ValueConfig = lConfig AndAlso vpVals(lIdx).ValueGroupNum = moSlots(X, Y).lGroupNum Then
                                vpVals(lIdx).Value += 1
                                bFound = True
                                Exit For
                            End If
                        Next lIdx

                        If bFound = False Then
                            lVPUB += 1
                            ReDim Preserve vpVals(lVPUB)
                            With vpVals(lVPUB)
                                .Value = 1
                                .ValueConfig = lConfig
                                .ValueGroupNum = moSlots(X, Y).lGroupNum
                            End With
                        End If
                    End If
                Next X
            Next Y

            Dim fHullPerSlot As Single = CSng(lHullSize / Me.TotalSlots)
            For X As Int32 = 0 To lVPUB
                If vpVals(X).Value * fHullPerSlot >= lHullSizeMin Then
                    vpVals(X).Value = 0
                Else : vpVals(X).Value = 1
                End If
            Next X


            'Now, go back through...
            For Y As Int32 = 0 To ml_GRID_SIZE_WH - 1
                For X As Int32 = 0 To ml_GRID_SIZE_WH - 1
                    If moSlots(X, Y).lConfig <> lConfig Then
                        moSlots(X, Y).bFiltered = True
                    Else
                        'Now, find our vpVal
                        moSlots(X, Y).bFiltered = True
                        For lIdx As Int32 = 0 To lVPUB
                            If vpVals(lIdx).ValueConfig = lConfig AndAlso vpVals(lIdx).ValueGroupNum = moSlots(X, Y).lGroupNum AndAlso vpVals(lIdx).Value = 0 Then
                                moSlots(X, Y).bFiltered = False
                                Exit For
                            End If
                        Next lIdx
                    End If
                Next X
            Next Y

            Me.IsDirty = True
        End If
    End Sub

    Public Sub HighLightAcceptableSlots(ByVal lConfig As SlotConfig, ByVal lGrpNum As Int32)
        'Ok, a slot config has come in that we need to match
        'an acceptable slot is...
        ' if the config is armor/crew quarters or cargo, all work
        ' otherwise, if there is an existing slot configured with this config, then acceptable slots are those
        ' that are directly next to those slots
        If lConfig = SlotConfig.eArmorConfig OrElse lConfig = SlotConfig.eCrewQuarters OrElse lConfig = SlotConfig.eCargoBay Then
            ClearSelectingSlotConfig()
        Else

            'ok, find a slot
            Dim bFound As Boolean = False
            For Y As Int32 = 0 To ml_GRID_SIZE_WH - 1
                For X As Int32 = 0 To ml_GRID_SIZE_WH - 1
                    If (moSlots(X, Y).lGroupNum = lGrpNum OrElse lGrpNum < 1) AndAlso moSlots(X, Y).lConfig = lConfig Then
                        If moSlots(X, Y).yType <> SlotType.eUnused AndAlso moSlots(X, Y).yType <> SlotType.eNoChange Then
                            bFound = True
                            Exit For
                        End If
                    End If
                Next X
                If bFound = True Then Exit For
            Next Y

            If bFound = False Then
                If lConfig = SlotConfig.eHangarDoor Then
                    FilterEdgeSlots(True)
                ElseIf lConfig = SlotConfig.eWeapons Then
                    FilterEdgeSlots(True)
                    For Y As Int32 = 0 To ml_GRID_SIZE_WH - 1
                        For X As Int32 = 0 To ml_GRID_SIZE_WH - 1
                            If moSlots(X, Y).yType = SlotType.eAllArc Then moSlots(X, Y).bFiltered = False
                        Next X
                    Next Y
                Else : ClearSelectingSlotConfig()
                End If
            Else
                'Ok, go thru all slots
                For Y As Int32 = 0 To ml_GRID_SIZE_WH - 1
                    For X As Int32 = 0 To ml_GRID_SIZE_WH - 1
                        'Is the slot our config?
                        Dim bGood As Boolean = False
                        If (moSlots(X, Y).lGroupNum = lGrpNum OrElse lGrpNum < 1) AndAlso moSlots(X, Y).lConfig = lConfig Then
                            If moSlots(X, Y).yType <> SlotType.eUnused AndAlso moSlots(X, Y).yType <> SlotType.eNoChange Then
                                bGood = True

                                If lConfig = SlotConfig.eHangarDoor Then
                                    bGood = SlotIsOnEdge(X, Y)
                                End If
                            End If
                        End If

                        If bGood = True Then
                            moSlots(X, Y).bFiltered = False
                            If X <> ml_GRID_SIZE_WH - 1 Then
                                moSlots(X + 1, Y).bFiltered = Not (lConfig <> SlotConfig.eHangarDoor OrElse SlotIsOnEdge(X + 1, Y) = True) ' False
                            End If
                            If X <> 0 Then
                                moSlots(X - 1, Y).bFiltered = Not (lConfig <> SlotConfig.eHangarDoor OrElse SlotIsOnEdge(X - 1, Y) = True) 'False
                            End If
                            If Y <> ml_GRID_SIZE_WH - 1 Then
                                moSlots(X, Y + 1).bFiltered = Not (lConfig <> SlotConfig.eHangarDoor OrElse SlotIsOnEdge(X, Y + 1) = True) 'False
                            End If
                            If Y <> 0 Then
                                moSlots(X, Y - 1).bFiltered = Not (lConfig <> SlotConfig.eHangarDoor OrElse SlotIsOnEdge(X, Y - 1) = True) 'False
                            End If
                        Else
                            'check left and up
                            Dim bFilter As Boolean = True
                            If X <> 0 Then
                                If (moSlots(X - 1, Y).lGroupNum = lGrpNum OrElse lGrpNum < 1) AndAlso moSlots(X - 1, Y).lConfig = lConfig Then
                                    If moSlots(X - 1, Y).yType <> SlotType.eUnused AndAlso moSlots(X - 1, Y).yType <> SlotType.eNoChange Then
                                        bFilter = Not (lConfig <> SlotConfig.eHangarDoor OrElse SlotIsOnEdge(X, Y) = True)
                                    End If
                                End If
                            End If
                            If Y <> 0 AndAlso bFilter = True Then
                                If (moSlots(X, Y - 1).lGroupNum = lGrpNum OrElse lGrpNum < 1) AndAlso moSlots(X, Y - 1).lConfig = lConfig Then
                                    If moSlots(X, Y - 1).yType <> SlotType.eUnused AndAlso moSlots(X, Y - 1).yType <> SlotType.eNoChange Then
                                        bFilter = Not (lConfig <> SlotConfig.eHangarDoor OrElse SlotIsOnEdge(X, Y) = True)
                                    End If
                                End If
                            End If

                            moSlots(X, Y).bFiltered = bFilter
                        End If
                    Next X
                Next Y
            End If
        End If

        Me.IsDirty = True
    End Sub

    Public Sub ClearSelectingSlotConfig()
        For Y As Int32 = 0 To ml_GRID_SIZE_WH - 1
            For X As Int32 = 0 To ml_GRID_SIZE_WH - 1
                moSlots(X, Y).bFiltered = False
            Next X
        Next Y
        Me.IsDirty = True
    End Sub

    Public Sub ClearAllConfigs()
        For Y As Int32 = 0 To ml_GRID_SIZE_WH - 1
            For X As Int32 = 0 To ml_GRID_SIZE_WH - 1
                moSlots(X, Y).bFiltered = False
                moSlots(X, Y).lConfig = SlotConfig.eArmorConfig
                moSlots(X, Y).lGroupNum = 0
            Next X
        Next Y
    End Sub

    Public Sub ClearConfigGroup(ByVal lIndexX As Int32, ByVal lIndexY As Int32)
        Dim lSlotConfig As SlotConfig = moSlots(lIndexX, lIndexY).lConfig
        Dim lGroup As Int32 = moSlots(lIndexX, lIndexY).lGroupNum
        For Y As Int32 = 0 To ml_GRID_SIZE_WH - 1
            For X As Int32 = 0 To ml_GRID_SIZE_WH - 1
                If moSlots(X, Y).bFiltered = False AndAlso moSlots(X, Y).lConfig = lSlotConfig AndAlso moSlots(X, Y).lGroupNum = lGroup Then
                    moSlots(X, Y).lConfig = SlotConfig.eArmorConfig
                    moSlots(X, Y).lGroupNum = 0
                End If
            Next X
        Next Y
        Me.IsDirty = True
    End Sub

    Private Function SlotIsOnEdge(ByVal X As Int32, ByVal Y As Int32) As Boolean
        'Ok, check our edges
        If X <> 0 Then
            If moSlots(X - 1, Y).yType = SlotType.eUnused Then
                Return True
            End If
        End If
        If X <> ml_GRID_SIZE_WH - 1 Then
            If moSlots(X + 1, Y).yType = SlotType.eUnused Then
                Return True
            End If
        End If
        If Y <> 0 Then
            If moSlots(X, Y - 1).yType = SlotType.eUnused Then
                Return True
            End If
        End If
        If Y <> ml_GRID_SIZE_WH - 1 Then
            If moSlots(X, Y + 1).yType = SlotType.eUnused Then
                Return True
            End If
        End If
        Return False
    End Function

    Public Sub FilterEdgeSlots(ByVal bEdgeSlots As Boolean)
        ClearSelectingSlotConfig()
        For Y As Int32 = 0 To ml_GRID_SIZE_WH - 1
            For X As Int32 = 0 To ml_GRID_SIZE_WH - 1
                If moSlots(X, Y).yType <> SlotType.eUnused Then
                    moSlots(X, Y).bFiltered = bEdgeSlots

                    'Ok, check our edges
                    If X <> 0 Then
                        If moSlots(X - 1, Y).yType = SlotType.eUnused Then
                            moSlots(X, Y).bFiltered = Not bEdgeSlots
                            Continue For
                        End If
                    End If
                    If X <> ml_GRID_SIZE_WH - 1 Then
                        If moSlots(X + 1, Y).yType = SlotType.eUnused Then
                            moSlots(X, Y).bFiltered = Not bEdgeSlots
                            Continue For
                        End If
                    End If
                    If Y <> 0 Then
                        If moSlots(X, Y - 1).yType = SlotType.eUnused Then
                            moSlots(X, Y).bFiltered = Not bEdgeSlots
                            Continue For
                        End If
                    End If
                    If Y <> ml_GRID_SIZE_WH - 1 Then
                        If moSlots(X, Y + 1).yType = SlotType.eUnused Then
                            moSlots(X, Y).bFiltered = Not bEdgeSlots
                            Continue For
                        End If
                    End If
                End If
            Next X
        Next Y
    End Sub

    Public Function HasChanged(ByRef oTech As HullTech) As Boolean
        For Y As Int32 = 0 To ml_GRID_SIZE_WH - 1
            For X As Int32 = 0 To ml_GRID_SIZE_WH - 1
                If oTech.moSlots(X, Y).lConfig <> moSlots(X, Y).lConfig OrElse oTech.moSlots(X, Y).lGroupNum <> moSlots(X, Y).lGroupNum OrElse oTech.moSlots(X, Y).yType <> moSlots(X, Y).yType Then
                    Return True
                End If
            Next X
        Next Y
        Return False
    End Function

    Public Function GetSlotConfigCnt(ByVal lSlotType As SlotType, ByVal lSlotConfig As SlotConfig) As Int32
        Dim lCnt As Int32 = 0
        For Y As Int32 = 0 To ml_GRID_SIZE_WH - 1
            For X As Int32 = 0 To ml_GRID_SIZE_WH - 1
                If lSlotType = SlotType.eNoChange OrElse moSlots(X, Y).yType = lSlotType Then
                    If moSlots(X, Y).lConfig = lSlotConfig Then lCnt += 1
                End If
            Next X
        Next Y
        Return lCnt
    End Function

    Public Function GetSlotConfigCntWithGroup(ByVal lSlotType As SlotType, ByVal lSlotConfig As SlotConfig, ByVal lGroup As Int32) As Int32
        Dim lCnt As Int32 = 0
        For Y As Int32 = 0 To ml_GRID_SIZE_WH - 1
            For X As Int32 = 0 To ml_GRID_SIZE_WH - 1
                If lSlotType = SlotType.eNoChange OrElse moSlots(X, Y).yType = lSlotType Then
                    If moSlots(X, Y).lConfig = lSlotConfig AndAlso moSlots(X, Y).lGroupNum = lGroup Then lCnt += 1
                End If
            Next X
        Next Y
        Return lCnt
    End Function

    Public Function GetBiggestHangarDoorsize() As Int32
        Dim lCnts() As Int32 = Nothing
        Dim lGroupIDs() As Int32 = Nothing
        Dim lUB As Int32 = -1

        For Y As Int32 = 0 To ml_GRID_SIZE_WH - 1
            For X As Int32 = 0 To ml_GRID_SIZE_WH - 1
                If moSlots(X, Y).yType <> SlotType.eUnused Then
                    If moSlots(X, Y).lConfig = SlotConfig.eHangarDoor Then
                        Dim lIdx As Int32 = -1
                        For Z As Int32 = 0 To lUB
                            If lGroupIDs(Z) = moSlots(X, Y).lGroupNum Then
                                lIdx = Z
                                Exit For
                            End If
                        Next Z
                        If lIdx = -1 Then
                            lUB += 1
                            ReDim Preserve lCnts(lUB)
                            ReDim Preserve lGroupIDs(lUB)
                            lIdx = lUB
                            lCnts(lIdx) = 0
                            lGroupIDs(lIdx) = moSlots(X, Y).lGroupNum
                        End If
                        lCnts(lIdx) += 1
                    End If
                End If
            Next X
        Next Y

        Dim lCnt As Int32 = 0
        For X As Int32 = 0 To lUB
            If lCnts(X) > lCnt Then lCnt = lCnts(X)
        Next X

        Return lCnt
    End Function

    Public Function GetMaxWeaponGroupNum() As Int32
        Dim lMaxVal As Int32 = 0
        For Y As Int32 = 0 To ml_GRID_SIZE_WH - 1
            For X As Int32 = 0 To ml_GRID_SIZE_WH - 1
                If moSlots(X, Y).yType <> SlotType.eUnused Then
                    If moSlots(X, Y).lConfig = SlotConfig.eWeapons Then
                        If moSlots(X, Y).lGroupNum > lMaxVal Then lMaxVal = moSlots(X, Y).lGroupNum
                    End If
                End If
            Next X
        Next Y
        Return lMaxVal
    End Function

    Public Function WillCauseErrors(ByVal lX As Int32, ByVal lY As Int32, ByVal lConfig As SlotConfig, ByVal lGrpNum As Int32, ByVal lTotalHull As Int32, ByVal yHullReq As eyHullRequirements) As Boolean
        Dim yOrigType As SlotType = SlotType.eNoChange
        Dim lOrigConfig As SlotConfig = SlotConfig.eArmorConfig
        Dim lOrigGrpNum As Int32 = 0
        Me.GetHullSlotValues(lX, lY, yOrigType, lOrigConfig, lOrigGrpNum)

        Dim sErrors As String = Me.HasErrorList(lTotalHull, 0, yHullReq)

        Dim lErrorCntPre As Int32 = 0
        Dim sSplits() As String = Split(sErrors.ToUpper, "ERROR")
        If sSplits Is Nothing = False Then
            lErrorCntPre = sSplits.GetUpperBound(0)
        End If

        Me.SetHullSlot(lX, lY, SlotType.eNoChange, lConfig, lGrpNum)

        sErrors = Me.HasErrorList(lTotalHull, 0, yHullReq)
        sSplits = Split(sErrors.ToUpper, "ERROR")
        Dim lErrorCntPost As Int32 = 0
        If sSplits Is Nothing = False Then
            lErrorCntPost = sSplits.GetUpperBound(0)
        End If

        Me.SetHullSlot(lX, lY, SlotType.eNoChange, lOrigConfig, lOrigGrpNum)

        Return lErrorCntPre < lErrorCntPost
    End Function
End Class