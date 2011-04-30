Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Interface created from Interface Builder
Public Class ctlTransport
    Inherits UIWindow

    Private mlID As Int32 = -1
    Private mdtETA As DateTime = DateTime.MinValue

    Private lblName As UILabel
    Private lblStatus As UILabel
    Private lblTime As UILabel
    Private mbCurrSelectState As Boolean = False

    Public Event TransportSelected(ByVal lID As Int32)

    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)

        'ctlTransport initial props
        With Me
            .ControlName = "ctlTransport"
            .Left = 296
            .Top = 220
            .Width = 220
            .Height = 45
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = False
        End With

        'lblName initial props
        lblName = New UILabel(oUILib)
        With lblName
            .ControlName = "lblName"
            .Left = 5
            .Top = 5
            .Width = 165
            .Height = 14
            .Enabled = True
            .Visible = True
            .Caption = ""
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.0F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblName, UIControl))

        'lblStatus initial props
        lblStatus = New UILabel(oUILib)
        With lblStatus
            .ControlName = "lblStatus"
            .Left = 5
            .Top = 24
            .Width = 240
            .Height = 14
            .Enabled = True
            .Visible = True
            .Caption = ""
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblStatus, UIControl))

        'lblTime initial props
        lblTime = New UILabel(oUILib)
        With lblTime
            .ControlName = "lblTime"
            .Left = Me.Width - 70
            .Top = 5
            .Width = 65
            .Height = 14
            .Enabled = True
            .Visible = True
            .Caption = ""
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Right
        End With
        Me.AddChild(CType(lblTime, UIControl))

        AddHandler lblName.OnMouseDown, AddressOf ctlTransport_OnMouseDown
        AddHandler lblStatus.OnMouseDown, AddressOf ctlTransport_OnMouseDown
        AddHandler lblTime.OnMouseDown, AddressOf ctlTransport_OnMouseDown
    End Sub

    Private Sub ctlTransport_OnMouseDown(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles Me.OnMouseDown
        RaiseEvent TransportSelected(mlID)
    End Sub

    Public Sub SetData(ByVal sStatus As String, ByVal dtETA As DateTime, ByVal lID As Int32)
        mlID = lID
        mdtETA = dtETA

        lblName.Caption = frmTransportManagement.GetTransportName(lID)
        lblStatus.Caption = sStatus

        If dtETA <> DateTime.MinValue Then

            Dim lSeconds As Int32
            If dtETA.Subtract(Now).TotalSeconds > Int32.MaxValue Then lSeconds = Int32.MaxValue Else lSeconds = CInt(dtETA.Subtract(Now).TotalSeconds)
            If lSeconds < 0 Then
                lblTime.Caption = ""
                mdtETA = DateTime.MinValue
            Else
                lblTime.Caption = GetDurationFromSeconds(lSeconds, False)
            End If
        Else
            lblTime.Caption = ""
        End If
    End Sub

    Public Sub SetSelectedState(ByVal bValue As Boolean)
        If bValue = mbCurrSelectState Then Return
        mbCurrSelectState = bValue
        Dim clrText As System.Drawing.Color
        If bValue = True Then
            Me.FillColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
            clrText = System.Drawing.Color.FromArgb(255, 223, 191, 128)
        Else
            clrText = muSettings.InterfaceBorderColor
            Me.FillColor = muSettings.InterfaceFillColor
        End If
        If lblName Is Nothing = False Then lblName.ForeColor = clrText
        If lblStatus Is Nothing = False Then lblStatus.ForeColor = clrText
        If lblTime Is Nothing = False Then lblTime.ForeColor = clrText
    End Sub

    Public Sub ctlTransport_OnNewFrame() Handles Me.OnNewFrame
        If mlID > -1 Then
            Dim oTrans As Transport = frmTransportManagement.GetTransport(mlID)
            oTrans.RequestDetails()
            Dim sCap As String = ""
            If oTrans.ETA <> DateTime.MinValue Then
                Dim lSeconds As Int32
                If oTrans.ETA.Subtract(Now).TotalSeconds > Int32.MaxValue Then lSeconds = Int32.MaxValue Else lSeconds = CInt(oTrans.ETA.Subtract(Now).TotalSeconds)
                If lSeconds < 0 Then
                    'oTrans.ETA = DateTime.MinValue
                    sCap = "Imminent"
                Else
                    sCap = GetDurationFromSeconds(lSeconds, False)
                End If
            ElseIf (oTrans.TransFlags And Transport.elTransportFlags.eEnRoute) <> 0 Then
                sCap = "Imminent"
            End If
            If lblTime.Caption <> sCap Then lblTime.Caption = sCap

            If oTrans Is Nothing = False Then
                sCap = oTrans.GetStatusText()
                If lblStatus.Caption <> sCap Then lblStatus.Caption = sCap
                sCap = frmTransportManagement.GetTransportName(mlID)
                If lblName.Caption <> sCap Then lblName.Caption = sCap
            End If
        End If
    End Sub

    Public Function GetID() As Int32
        Return mlID
    End Function
End Class
