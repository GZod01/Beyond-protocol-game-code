Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Interface created from Interface Builder
Public Class frmCoeff
    Inherits UIWindow

    Private lblName() As UILabel
    Private txtValue() As UITextBox
    Private chkUseValues As UICheckBox
    Private WithEvents btnReload As UIButton
    Private WithEvents btnClose As UIButton

    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)

        'frmCoeff initial props
        With Me
            .ControlName = "frmCoeff"
            .Left = 266
            .Top = 163
            .Width = 127
            .Height = 511
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = True
            .Moveable = True
        End With

        ReDim lblName(14)
        ReDim txtValue(14)

        For X As Int32 = 0 To 14
            'lblName initial props
            lblName(X) = New UILabel(oUILib)
            With lblName(X)
                .ControlName = "lblName(" & X & ")"
                .Left = 5
                .Top = 5 + (X * 20)
                .Width = 60
                .Height = 14
                .Enabled = True
                .Visible = True
                .Caption = GetCoeffName(X)
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblName(X), UIControl))

            'txtHull initial props
            txtValue(X) = New UITextBox(oUILib)
            With txtValue(X)
                .ControlName = "txtValue(" & X & ")"
                .Left = 70
                .Top = lblName(X).Top
                .Width = 50
                .Height = 14
                .Enabled = True
                .Visible = True
                .Caption = "0"
                .ForeColor = muSettings.InterfaceTextBoxForeColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
                .BackColorEnabled = muSettings.InterfaceTextBoxFillColor
                .BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
                .MaxLength = 0
                .BorderColor = muSettings.InterfaceBorderColor
            End With
            Me.AddChild(CType(txtValue(X), UIControl))
        Next X
 
        chkUseValues = New UICheckBox(oUILib)
        With chkUseValues
            .ControlName = "chkUseValues"
            .Left = 5
            .Top = txtValue(14).Top + 20
            .Width = 100
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Use Values"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = False
            .DisplayType = UICheckBox.eCheckTypes.eSmallCheck
        End With
        Me.AddChild(CType(chkUseValues, UIControl))
 
        'btnReload initial props
        btnReload = New UIButton(oUILib)
        With btnReload
            .ControlName = "btnReload"
            .Left = 5
            .Top = chkUseValues.Top + chkUseValues.Height + 5
            .Width = 115
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Reload"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnReload, UIControl))

        'btnClose initial props
        btnClose = New UIButton(oUILib)
        With btnClose
            .ControlName = "btnClose"
            .Left = 5
            .Top = btnReload.Top + btnReload.Height + 5
            .Width = 115
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Close"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnClose, UIControl))

        MyBase.moUILib.RemoveWindow(Me.ControlName)
        MyBase.moUILib.AddWindow(Me)
    End Sub

    Private Function GetCoeffName(ByVal lVal As Int32) As String
        Select Case lVal
            Case TechBuilderComputer.elPropCoeffLookup.eColonist
                Return "Colonist"
            Case TechBuilderComputer.elPropCoeffLookup.eEnlisted
                Return "Enlisted"
            Case TechBuilderComputer.elPropCoeffLookup.eHull
                Return "Hull"
            Case TechBuilderComputer.elPropCoeffLookup.eMin1
                Return "Min 1"
            Case TechBuilderComputer.elPropCoeffLookup.eMin2
                Return "Min 2"
            Case TechBuilderComputer.elPropCoeffLookup.eMin3
                Return "Min 3"
            Case TechBuilderComputer.elPropCoeffLookup.eMin4
                Return "Min 4"
            Case TechBuilderComputer.elPropCoeffLookup.eMin5
                Return "Min 5"
            Case TechBuilderComputer.elPropCoeffLookup.eMin6
                Return "Min 6"
            Case TechBuilderComputer.elPropCoeffLookup.eOfficer
                Return "Officer"
            Case TechBuilderComputer.elPropCoeffLookup.ePower
                Return "Power"
            Case TechBuilderComputer.elPropCoeffLookup.eProdCost
                Return "Prod Cost"
            Case TechBuilderComputer.elPropCoeffLookup.eProdTime
                Return "Prod Time"
            Case TechBuilderComputer.elPropCoeffLookup.eResCost
                Return "Res Cost"
            Case TechBuilderComputer.elPropCoeffLookup.eResTime
                Return "Res Time"
        End Select
        Return ""
    End Function

    Public Sub SetCoeff(ByVal decVals() As Decimal)
        Try
            For X As Int32 = 0 To 14
                txtValue(X).Caption = decVals(X).ToString("###0.#####")
            Next X
        Catch
            MyBase.moUILib.AddNotification("Didn't run the thing before you tried the thing", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
        End Try
    End Sub

    Private mdecBase() As Decimal

    Public Function GetCoeff(ByVal decVals() As Decimal) As Decimal()
        mdecBase = decVals
        If chkUseValues.Value = True Then
            Dim decTemp(14) As Decimal
            For X As Int32 = 0 To 14
                decTemp(X) = CDec(txtValue(X).Caption)
            Next X
            Return decTemp
        Else
            Return decVals
        End If
    End Function

    Private Sub btnReload_Click(ByVal sName As String) Handles btnReload.Click
        SetCoeff(mdecBase)
    End Sub

    Private Sub btnClose_Click(ByVal sName As String) Handles btnClose.Click
        MyBase.moUILib.RemoveWindow(Me.ControlName)
    End Sub
End Class