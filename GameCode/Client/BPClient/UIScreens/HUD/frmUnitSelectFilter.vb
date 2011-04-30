Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

Public Class frmUnitSelectFilter
    Inherits UIWindow

    Private lblTitle As UILabel
    Private WithEvents btnClose As UIButton
    Private lnDiv1 As UILine

    Private WithEvents optSelect As UIOption
    Private WithEvents optDeselect As UIOption

    Private chkUndamaged As UICheckBox
    Private chkDamaged As UICheckBox
    Private chkCritical As UICheckBox
    Private chkDisabled As UICheckBox

    Private chkUnits As UICheckBox
    Private chkFacilities As UICheckBox

    Private mbSelect As Boolean = True

    Private WithEvents btnFilter As UIButton

    Private mbLoading As Boolean = True

    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)

        'frmUnitSelectFilter initial props
        With Me
            .lWindowMetricID = BPMetrics.eWindow.eTransfer
            .ControlName = "frmUnitSelectFilter"
            .Width = 120
            .Height = 240
            .Top = CInt(oUILib.oDevice.PresentationParameters.BackBufferHeight \ 2) - (.Height \ 2)
            .Left = (oUILib.oDevice.PresentationParameters.BackBufferWidth \ 2) - (.Width \ 2)
            Dim oFrm As frmMultiDisplay = CType(MyBase.moUILib.GetWindow("frmMultiDisplay"), frmMultiDisplay)
            If oFrm Is Nothing = False Then
                .Top = oFrm.Top
                .Left = oFrm.Left + oFrm.Width
            End If
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = True
            .BorderLineWidth = 2
            .bRoundedBorder = True
        End With

        'lblTitle initial props
        lblTitle = New UILabel(oUILib)
        With lblTitle
            .ControlName = "lblTitle"
            .Left = 5
            .Top = 5
            .Width = 170
            .Height = 16
            .Enabled = True
            .Visible = True
            .Caption = "Filter"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblTitle, UIControl))

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
            .Left = Me.BorderLineWidth
            .Top = 25
            .Width = Me.Width - Me.BorderLineWidth
            .Height = 0
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(lnDiv1, UIControl))

        'optSelect initial props
        optSelect = New UIOption(oUILib)
        With optSelect
            .ControlName = "optSelect"
            .Left = 5
            .Top = 30
            .Width = 82
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Select Only"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .ToolTipText = "Select only chosen units or facilities"
            .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Right
            .Value = True
        End With
        Me.AddChild(CType(optSelect, UIControl))

        'optDeselect initial props
        optDeselect = New UIOption(oUILib)
        With optDeselect
            .ControlName = "optDeselect"
            .Left = 5
            .Top = optSelect.Top + 20
            .Width = 68
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Deselect"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .ToolTipText = "Deselect chosen units or facilities"
            .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Right
            .Value = False
        End With
        Me.AddChild(CType(optDeselect, UIControl))

        'chkUndamaged initial props
        chkUndamaged = New UICheckBox(oUILib)
        With chkUndamaged
            .ControlName = "chkUndamaged"
            .Left = 5
            .Top = optDeselect.Top + 25
            .Width = 90
            .Height = 16
            .Enabled = True
            .Visible = True
            .Caption = "Undamaged"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Right
            .Value = False
            .DisplayType = UICheckBox.eCheckTypes.eSmallX
        End With
        Me.AddChild(CType(chkUndamaged, UIControl))

        'chkDamaged initial props
        chkDamaged = New UICheckBox(oUILib)
        With chkDamaged
            .ControlName = "chkDamaged"
            .Left = 5
            .Top = chkUndamaged.Top + 20
            .Width = 75
            .Height = 16
            .Enabled = True
            .Visible = True
            .Caption = "Damaged"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Right
            .Value = False
            .DisplayType = UICheckBox.eCheckTypes.eSmallX
        End With
        Me.AddChild(CType(chkDamaged, UIControl))

        'chkCritical initial props
        chkCritical = New UICheckBox(oUILib)
        With chkCritical
            .ControlName = "chkCritical"
            .Left = 5
            .Top = chkDamaged.Top + 20
            .Width = 54
            .Height = 16
            .Enabled = True
            .Visible = True
            .Caption = "Critical"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Right
            .Value = False
            .DisplayType = UICheckBox.eCheckTypes.eSmallX
        End With
        Me.AddChild(CType(chkCritical, UIControl))

        'chkDisabled initial props
        chkDisabled = New UICheckBox(oUILib)
        With chkDisabled
            .ControlName = "chkDisabled"
            .Left = 5
            .Top = chkCritical.Top + 20
            .Width = 69
            .Height = 16
            .Enabled = True
            .Visible = True
            .Caption = "Disabled"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Right
            .Value = False
            .DisplayType = UICheckBox.eCheckTypes.eSmallX
        End With
        Me.AddChild(CType(chkDisabled, UIControl))

        'chkUnits initial props
        chkUnits = New UICheckBox(oUILib)
        With chkUnits
            .ControlName = "chkUnits"
            .Left = 5
            .Top = chkDisabled.Top + 25
            .Width = 44
            .Height = 16
            .Enabled = True
            .Visible = True
            .Caption = "Units"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Right
            .Value = False
            .DisplayType = UICheckBox.eCheckTypes.eSmallX
        End With
        Me.AddChild(CType(chkUnits, UIControl))

        'chkFacilities initial props
        chkFacilities = New UICheckBox(oUILib)
        With chkFacilities
            .ControlName = "chkFacilities"
            .Left = 5
            .Top = chkUnits.Top + 20
            .Width = 67
            .Height = 16
            .Enabled = True
            .Visible = True
            .Caption = "Facilities"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Right
            .Value = False
            .DisplayType = UICheckBox.eCheckTypes.eSmallX
        End With
        Me.AddChild(CType(chkFacilities, UIControl))

        'btnFilter initial props
        btnFilter = New UIButton(oUILib)
        With btnFilter
            .ControlName = "btnFilter"
            .Width = 100
            .Height = 24
            .Left = CInt((Me.Width - .Width) / 2)
            .Top = chkFacilities.Top + 25
            .Enabled = True
            .Visible = True
            .Caption = "Filter"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnFilter, UIControl))

        MyBase.moUILib.RemoveWindow(Me.ControlName)
        MyBase.moUILib.AddWindow(Me)
        mbLoading = False
    End Sub

    Private Sub btnClose_Click(ByVal sName As String) Handles btnClose.Click
        MyBase.moUILib.RemoveWindow(Me.ControlName)
    End Sub

    Private Sub btnFilter_Click(ByVal sName As String) Handles btnFilter.Click
        Dim oFrm As frmMultiDisplay = CType(MyBase.moUILib.GetWindow("frmMultiDisplay"), frmMultiDisplay)
        If oFrm Is Nothing = False Then
            oFrm.HandleSelectFilter(chkUndamaged.Value, chkDamaged.Value, chkCritical.Value, chkDisabled.Value, chkUnits.Value, chkFacilities.Value, mbSelect)
        End If
        MyBase.moUILib.RemoveWindow(Me.ControlName)
    End Sub

    Private Sub optDeselect_Click() Handles optDeselect.Click
        optDeselect.Value = True
        If mbSelect = False Then Return
        optSelect.Value = False
        mbSelect = False
    End Sub

    Private Sub optSelect_Click() Handles optSelect.Click
        optSelect.Value = True
        If mbSelect = True Then Return
        optDeselect.Value = False
        mbSelect = True
    End Sub

    Private Sub frmUnitSelectFilter_WindowMoved() Handles Me.WindowMoved
        Dim oWin As UIWindow = MyBase.moUILib.GetWindow("frmMultiDisplay")
        If oWin Is Nothing = False Then
            oWin.Top = Me.Top
            oWin.Left = Me.Left - oWin.Width
        End If
    End Sub
End Class
