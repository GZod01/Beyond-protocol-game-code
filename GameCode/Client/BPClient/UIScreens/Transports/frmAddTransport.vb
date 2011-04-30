Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Interface created from Interface Builder
Public Class frmAddTransport
    Inherits UIWindow

    Private lblTitle As UILabel
    Private lstUnits As UIListBox
    Private lblNote As UILabel
    Private lnDiv1 As UILine
    Private WithEvents btnAdd As UIButton
    Private WithEvents btnClose As UIButton

    Private mlLastRefresh As Int32 = -1

#Region "  Potential Transports Management  "
    Private Structure uPotential
        Public lID As Int32
        Public sDisplay As String
        Public bActive As Boolean
        Public sRealName As String
    End Structure
    Private Shared muPotentials() As uPotential
    Private Shared mlPotentialUB As Int32 = -1
    Private Shared mlLastPotentialRequest As Int32
    Private Shared mlLastPotentialEnvirID As Int32
    Private Shared miLastPotentialEnvirTypeID As Int16

    Public Shared Sub HandleRequestColonyTransportPotentials(ByVal yData() As Byte)
        Dim lPos As Int32 = 2
        Dim lCnt As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        Dim uTemp(lCnt - 1) As uPotential

        For X As Int32 = 0 To lCnt - 1
            With uTemp(X)
                .lID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                Dim sName As String = GetStringFromBytes(yData, lPos, 20) : lPos += 20
                Dim lCargo As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .sRealName = sName
                .sDisplay = sName.PadRight(20, " "c) & lCargo.ToString.PadLeft(7, " "c)
                .bActive = True
            End With
        Next X
        Try
            Dim lSorted() As Int32 = Nothing
            Dim lSortedUB As Int32 = -1
            For X As Int32 = 0 To lCnt - 1
                Dim lIdx As Int32 = -1

                Dim sTemp As String = uTemp(X).sDisplay.ToUpper
                For Y As Int32 = 0 To lSortedUB
                    If uTemp(lSorted(Y)).sDisplay.ToUpper > sTemp Then
                        lIdx = Y
                        Exit For
                    End If
                Next Y
                lSortedUB += 1
                ReDim Preserve lSorted(lSortedUB)
                If lIdx = -1 Then
                    lSorted(lSortedUB) = X
                Else
                    For Y As Int32 = lSortedUB To lIdx + 1 Step -1
                        lSorted(Y) = lSorted(Y - 1)
                    Next Y
                    lSorted(lIdx) = X
                End If
            Next X
            Dim uFinal(lCnt - 1) As uPotential
            For X As Int32 = 0 To lCnt - 1
                uFinal(X) = uTemp(lSorted(X))
            Next X
            uTemp = uFinal
        Catch
        End Try

        mlPotentialUB = -1
        muPotentials = uTemp
        mlPotentialUB = lCnt - 1
    End Sub

    Private Sub RequestColonyTransportPotentials()

        Try
            If goCurrentEnvir Is Nothing = False Then

                'Get the new envirID
                Dim lEnvID As Int32 = -1
                Dim iEnvirTypeID As Int16 = -1
                If goCurrentEnvir.ObjTypeID = ObjectType.eSolarSystem Then
                    For X As Int32 = 0 To goCurrentEnvir.lEntityUB
                        If goCurrentEnvir.lEntityIdx(X) > -1 Then
                            Dim oEnt As BaseEntity = goCurrentEnvir.oEntity(X)
                            If oEnt Is Nothing = False Then
                                If oEnt.bSelected = True AndAlso oEnt.yProductionType = ProductionType.eSpaceStationSpecial Then
                                    lEnvID = oEnt.ObjectID
                                    iEnvirTypeID = oEnt.ObjTypeID
                                    Exit For
                                End If
                            End If
                        End If
                    Next
                ElseIf goCurrentEnvir.ObjTypeID = ObjectType.ePlanet Then
                    lEnvID = goCurrentEnvir.ObjectID
                    iEnvirTypeID = goCurrentEnvir.ObjTypeID
                Else
                    'shouldn't be here so return
                    Return
                End If

                '5 seconds or we change environments/selections
                If glCurrentCycle - mlLastPotentialRequest > 150 OrElse mlLastPotentialEnvirID <> lEnvID OrElse miLastPotentialEnvirTypeID <> iEnvirTypeID Then
                    'if we are in space, request using the selceted space station
                    'if we are in planet, request using the planet
                    Dim yMsg(7) As Byte
                    System.BitConverter.GetBytes(GlobalMessageCode.eRequestColonyTransportPotentials).copyto(yMsg, 0)
                    System.BitConverter.GetBytes(lEnvID).CopyTo(yMsg, 2)
                    System.BitConverter.GetBytes(iEnvirTypeID).CopyTo(yMsg, 6)
                    goUILib.SendMsgToPrimary(yMsg)

                    mlLastPotentialRequest = glCurrentCycle
                    mlLastPotentialEnvirID = lEnvID
                    miLastPotentialEnvirTypeID = iEnvirTypeID
                End If
            End If
        Catch
        End Try

    End Sub

    Public Shared Sub HandleAddTransport(ByVal lUnitID As Int32)
        Dim lUB As Int32 = -1
        If muPotentials Is Nothing = False Then lUB = Math.Min(mlPotentialUB, muPotentials.GetUpperBound(0))

        For X As Int32 = 0 To lUB
            If muPotentials(X).lID = lUnitID Then
                muPotentials(X).bActive = False
                Exit For
            End If
        Next X
    End Sub
#End Region

    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)

        'frmAddTransport initial props
        With Me
            .ControlName = "frmAddTransport"
            .Left = 257
            .Top = 107
            .Width = 254
            .Height = 294
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = False
        End With

        'lblTitle initial props
        lblTitle = New UILabel(oUILib)
        With lblTitle
            .ControlName = "lblTitle"
            .Left = 5
            .Top = 5
            .Width = 157
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Add to Transport Fleet"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblTitle, UIControl))

        'lstUnits initial props
        lstUnits = New UIListBox(oUILib)
        With lstUnits
            .ControlName = "lstUnits"
            .Left = 5
            .Top = 105
            .Width = 245
            .Height = 150
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Courier New", 9.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
        End With
        Me.AddChild(CType(lstUnits, UIControl))

        'lblNote initial props
        lblNote = New UILabel(oUILib)
        With lblNote
            .ControlName = "lblNote"
            .Left = 5
            .Top = 30
            .Width = 245
            .Height = 70
            .Enabled = True
            .Visible = True
            .Caption = "Select a unit in the list to create as a transport. This process cannot be undone. The unit will be a transport and cannot become a normal unit again."
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.WordBreak Or DrawTextFormat.Center Or DrawTextFormat.VerticalCenter
        End With
        Me.AddChild(CType(lblNote, UIControl))

        'btnAdd initial props
        btnAdd = New UIButton(oUILib)
        With btnAdd
            .ControlName = "btnAdd"
            .Left = 70
            .Top = 263
            .Width = 120
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Create Transport"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnAdd, UIControl))

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

        'lnDiv1 initial props
        lnDiv1 = New UILine(oUILib)
        With lnDiv1
            .ControlName = "lnDiv1"
            .Left = Me.BorderLineWidth \ 2
            .Top = 27
            .Width = Me.Width - Me.BorderLineWidth
            .Height = 0
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(lnDiv1, UIControl))

        RequestColonyTransportPotentials()

        MyBase.moUILib.RemoveWindow(Me.ControlName)
        MyBase.moUILib.AddWindow(Me)

    End Sub

    Private Sub btnAdd_Click(ByVal sName As String) Handles btnAdd.Click
        If lstUnits.ListIndex > -1 Then
            Dim lID As Int32 = lstUnits.ItemData(lstUnits.ListIndex)
            Dim lUB As Int32 = -1
            If muPotentials Is Nothing = False Then lUB = Math.Min(muPotentials.GetUpperBound(0), mlPotentialUB)

            Dim lCnt As Int32 = 0
            For X As Int32 = 0 To lUB
                If muPotentials(X).bActive = True Then
                    If muPotentials(X).lID = lID Then

                        mlUnitID = lID
                        Dim oMsgBox As New frmMsgBox(goUILib, "Creating a transport from " & muPotentials(X).sRealName & " will result in the unit being lost and a new transport being created. This process cannot be undone. Do you wish to proceed?", MsgBoxStyle.YesNo, "Confirm Transport")
                        AddHandler oMsgBox.DialogClosed, AddressOf AddTransportResult

                        Exit For
                    End If
                End If
            Next X

        Else
            goUILib.AddNotification("Select a Unit to add as a transport.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Return
        End If
    End Sub

    Private Sub btnClose_Click(ByVal sName As String) Handles btnClose.Click
        MyBase.moUILib.RemoveWindow(Me.ControlName)
    End Sub

    Private Sub frmAddTransport_OnNewFrame() Handles Me.OnNewFrame
        If glCurrentCycle - mlLastRefresh > 30 Then     '1 second
            Try
                Dim lUB As Int32 = -1
                If muPotentials Is Nothing = False Then lUB = Math.Min(muPotentials.GetUpperBound(0), mlPotentialUB)

                Dim lCnt As Int32 = 0
                For X As Int32 = 0 To lUB
                    If muPotentials(X).bActive = True Then
                        lCnt += 1
                    End If
                Next X

                If lstUnits.ListCount <> lCnt Then
                    'just force refresh
                    lstUnits.Clear()
                    For X As Int32 = 0 To lUB
                        If muPotentials(X).bActive = True Then
                            lstUnits.AddItem(muPotentials(X).sDisplay, False)
                            lstUnits.ItemData(lstUnits.NewIndex) = muPotentials(X).lID
                        End If
                    Next X
                Else
                    'normal refresh
                    For X As Int32 = 0 To lUB
                        Dim bFound As Boolean = False
                        For Y As Int32 = 0 To lstUnits.ListCount - 1
                            If lstUnits.ItemData(Y) = muPotentials(X).lID Then
                                If muPotentials(X).bActive = False Then
                                    lstUnits.Clear()
                                    Return
                                End If
                                bFound = True
                                Exit For
                            End If
                        Next Y
                        If bFound = False AndAlso muPotentials(X).bActive = True Then
                            lstUnits.Clear()
                            Return
                        End If
                    Next X

                End If

                mlLastRefresh = glCurrentCycle
            Catch
                'just ignore
            End Try
        End If
    End Sub

    Private mlUnitID As Int32
    Private Sub AddTransportResult(ByVal lRes As MsgBoxResult)
        If lRes = MsgBoxResult.Yes Then
            Dim yMsg(5) As Byte
            System.BitConverter.GetBytes(GlobalMessageCode.eAddTransport).copyto(yMsg, 0)
            System.BitConverter.GetBytes(mlUnitID).CopyTo(yMsg, 2)
            goUILib.SendMsgToPrimary(yMsg)
        End If
    End Sub
End Class