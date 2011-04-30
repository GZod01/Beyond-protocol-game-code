Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

Public Class frmEnvirList
    Inherits UIWindow

    Private lblTitle As UILabel
    Private WithEvents btnClose As UIButton
    Private WithEvents lstEnvir As UIListBox
    Private lblStar As UILabel
    Private WithEvents txtStarName As UITextBox

    Private mbLoading As Boolean = True

    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)
        Dim ofrmED As frmEnvirDisplay = CType(goUILib.GetWindow("frmEnvirDisplay"), frmEnvirDisplay)

        'frmEnvirList initial props
        With Me
            .ControlName = "frmEnvirList"
            .Width = 211
            .Height = 275
            .Left = ofrmED.Left - .Width
            .Top = ofrmED.Top

            If .Left + .Width > oUILib.oDevice.PresentationParameters.BackBufferWidth Then .Left = oUILib.oDevice.PresentationParameters.BackBufferWidth - .Width
            If .Top + .Height > oUILib.oDevice.PresentationParameters.BackBufferHeight Then .Top = oUILib.oDevice.PresentationParameters.BackBufferHeight - .Height
            .Enabled = True
            .Visible = True
            .Moveable = False
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = False
            .BorderLineWidth = 2
        End With
        Debug.Write(Me.ControlName & " Newed" & vbCrLf)

        'lblTitle initial props
        lblTitle = New UILabel(oUILib)
        With lblTitle
            .ControlName = "lblTitle"
            .Left = 4
            .Top = 1
            .Width = 160
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Solar System List"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.0F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblTitle, UIControl))

        'lstEnvir initial props
        lstEnvir = New UIListBox(oUILib)
        With lstEnvir
            .ControlName = "lstEnvir"
            .Left = 2
            .Top = 27
            .Width = Me.Width - 4
            .Height = Me.Height - .Top - 19
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Courier New", 9.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
        End With
        Me.AddChild(CType(lstEnvir, UIControl))

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

        'lblStar initial props
        lblStar = New UILabel(oUILib)
        With lblStar
            .ControlName = "lblStar"
            .Width = 77
            .Height = 18
            .Left = 2
            .Top = Me.Height - .Height - 2
            .Enabled = True
            .Visible = True
            .Caption = "Quick Jump"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.0F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
            .ToolTipText = "Type the name of the star and press enter to quick jump"
        End With
        Me.AddChild(CType(lblStar, UIControl))

        'txtStarName initial props
        txtStarName = New UITextBox(oUILib)
        With txtStarName
            .ControlName = "txtStarName"
            .Left = lblStar.Left + lblStar.Width + 1
            .Width = Me.Width - .Left - 2
            .Height = 18
            .Top = Me.Height - .Height - 2
            .Enabled = True
            .Visible = True
            .Caption = ""
            .ForeColor = muSettings.InterfaceTextBoxForeColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
            .BackColorEnabled = muSettings.InterfaceTextBoxFillColor
            .BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
            .MaxLength = 25
            .BorderColor = muSettings.InterfaceBorderColor
            .ToolTipText = "Type the name of the star and press enter to quick jump"
        End With
        Me.AddChild(CType(txtStarName, UIControl))

        MyBase.moUILib.RemoveWindow(Me.ControlName)
        MyBase.moUILib.AddWindow(Me)
        mbLoading = False
        RefreshEnvironmentList()
        If MyBase.moUILib.FocusedControl Is Nothing = False Then MyBase.moUILib.FocusedControl.HasFocus = False
        MyBase.moUILib.FocusedControl = txtStarName
        txtStarName.HasFocus = True
    End Sub

    Private Sub RefreshEnvironmentList()
        lstEnvir.Clear()

        Dim lSorted() As Int32 = Nothing
        Dim lSortedUB As Int32 = -1
        For X As Int32 = 0 To goGalaxy.mlSystemUB
            Dim lIdx As Int32 = -1

            Dim sName As String = goGalaxy.moSystems(X).SystemName

            For Y As Int32 = 0 To lSortedUB
                Dim sOtherName As String = goGalaxy.moSystems(lSorted(Y)).SystemName
                If sOtherName > sName Then
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

        For X As Int32 = 0 To goGalaxy.mlSystemUB
            If goGalaxy.moSystems(lSorted(X)).ObjectID <> 88 AndAlso goGalaxy.moSystems(lSorted(X)).ObjectID <> 36 Then
                With goGalaxy.moSystems(lSorted(X))
                    Dim sText As String = .SystemName
                    lstEnvir.AddItem(sText, False)
                    lstEnvir.ItemData(lstEnvir.NewIndex) = .ObjectID
                End With
            End If
        Next X
    End Sub

    Private Sub btnClose_Click(ByVal sName As String) Handles btnClose.Click
        CloseForm()
    End Sub

    Private Sub lstEnvir_ItemDblClick(ByVal lIndex As Integer) Handles lstEnvir.ItemDblClick
        GotoEnvironmentWrapper(lstEnvir.ItemData(lIndex), ObjectType.eSolarSystem)
        CloseForm()
    End Sub

    Private Sub CloseForm()
        MyBase.moUILib.RemoveWindow(Me.ControlName)
        Me.Visible = False
    End Sub

    Private Sub txtStarName_OnKeyDown(ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtStarName.OnKeyDown
        If goGalaxy Is Nothing Then Return
        If e.KeyCode = Keys.Enter Then
            'Change environments to the star
            If goGalaxy.GotoStar(txtStarName.Caption) = True Then
                CloseForm()
            Else
                If txtStarName.Caption <> "" Then txtStarName.Caption = ""
            End If
        End If
    End Sub
End Class
