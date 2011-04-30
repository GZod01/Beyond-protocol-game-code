Option Strict On
Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

Public Class frmBuildFilter
    Inherits UIWindow

    Private lblFilter As UILabel
    Private WithEvents cboBuildFilter As UIComboBox
    Private cboSetBuildFilter As UIComboBox
    Private WithEvents btnNew As UIButton
    Private WithEvents btnDelete As UIButton
    Private WithEvents btnRename As UIButton
    Private WithEvents btnAdd As UIButton

    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)

        With Me
            .lWindowMetricID = BPMetrics.eWindow.eBuildWindow
            .ControlName = "frmBuildFilter"
            .Left = 118
            .Top = 76
            .Width = 511
            .Height = 511
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = False
            .Moveable = True
            .mbAcceptReprocessEvents = True
            .bRoundedBorder = True
            .BorderLineWidth = 2
        End With



        'btnNew initial props
        btnNew = New UIButton(oUILib)
        With btnNew
            .ControlName = "btnNew"
            .Width = 44
            .Height = 18
            .Left = 45
            .Top = 25
            .Enabled = True
            .Visible = True
            .Caption = "New"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnNew, UIControl))

        'btnRename initial props
        btnRename = New UIButton(oUILib)
        With btnRename
            .ControlName = "btnRename"
            .Width = 65
            .Height = 18
            .Left = btnNew.Left + btnNew.Width + 1
            .Top = btnNew.Top
            .Enabled = True
            .Visible = True
            .Caption = "Rename"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnRename, UIControl))

        'btnDelete initial props
        btnDelete = New UIButton(oUILib)
        With btnDelete
            .ControlName = "btnDelete"
            .Width = 60
            .Height = 18
            .Left = btnRename.Left + btnRename.Width + 1
            .Top = btnNew.Top
            .Enabled = True
            .Visible = True
            .Caption = "Delete"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnDelete, UIControl))

        'btnAdd initial props
        btnAdd = New UIButton(oUILib)
        With btnAdd
            .ControlName = "btnAdd"
            .Width = 110
            .Height = 18
            .Left = 45
            .Top = 71
            .Enabled = True
            .Visible = True
            .Caption = "Add To Filter"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnAdd, UIControl))

        'cboSetBuildFilter initial props
        cboSetBuildFilter = New UIComboBox(oUILib)
        With cboSetBuildFilter
            .ControlName = "cboSetBuildFilter"
            .Top = btnDelete.Top + btnDelete.Height + 5
            .Left = 45
            .Width = 170
            .Height = 18
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceTextBoxFillColor
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
            .DropDownListBorderColor = muSettings.InterfaceBorderColor
            .mbAcceptReprocessEvents = True
        End With
        Me.AddChild(CType(cboSetBuildFilter, UIControl))




    End Sub

    Private Sub btnDelete_Click(ByVal sName As String) Handles btnDelete.Click
        'xxx
    End Sub

    Private Sub btnNew_Click(ByVal sName As String) Handles btnNew.Click
        'xxx
    End Sub

    Private Sub btnRename_Click(ByVal sName As String) Handles btnRename.Click
        'xxx
    End Sub

    Private Sub cboBuildFilter_ItemSelected(ByVal lItemIndex As Integer) Handles cboBuildFilter.ItemSelected
        'xxx
    End Sub
End Class

