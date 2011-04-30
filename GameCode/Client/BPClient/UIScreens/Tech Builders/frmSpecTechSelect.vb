Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Interface created from Interface Builder
Public Class frmSpecTechSelect
    Inherits UIWindow

    Private lblTitle As UILabel
    Private lnDiv1 As UILine
    Private lblDesc As UILabel
    Private WithEvents lstTechs As UIListBox
    Private WithEvents btnSubmit As UIButton
    Private WithEvents btnClose As UIButton
    Private lblSelected As UILabel

    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)

        'frmSpecTechSelect initial props
        With Me
            .lWindowMetricID = BPMetrics.eWindow.eSpecialTech
            .ControlName = "frmSpecTechSelect"
            .Left = (MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth \ 2) - 300
            .Top = (MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight \ 2) - 220
            .Width = 511
            .Height = 511
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = False

            .Moveable = True
            .BorderLineWidth = 1
            .mbAcceptReprocessEvents = True
        End With

        'lblTitle initial props
        lblTitle = New UILabel(oUILib)
        With lblTitle
            .ControlName = "lblTitle"
            .Left = 5
            .Top = 2
            .Width = 236
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Special Technology Selection"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 12.0F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblTitle, UIControl))

        'NewControl2 initial props
        lnDiv1 = New UILine(oUILib)
        With lnDiv1
            .ControlName = "lnDiv1"
            .Left = 1
            .Top = 25
            .Width = 511
            .Height = 0
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(lnDiv1, UIControl))

        'lblDesc initial props
        lblDesc = New UILabel(oUILib)
        With lblDesc
            .ControlName = "lblDesc"
            .Left = 1
            .Top = 30
            .Width = 511
            .Height = 116
            .Enabled = True
            .Visible = True
            .Caption = "Beyond Protocol does not utilize the traditional 'tech tree'. Instead, it" & vbCrLf & _
                       "uses a 'tech cloud'. Various links are made while researching items within" & vbCrLf & _
                       "the cloud while others are lost. When the end of the cloud has been reached," & vbCrLf & _
                       "broken links are repaired at a cost and then the cloud tries again." & vbCrLf & vbCrLf & _
                       "Below is a list of projects that will be available later in the game. Select one" & vbCrLf & _
                       "item in the list and that item is guaranteed to link before the end of the first cloud."
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(1, DrawTextFormat) Or DrawTextFormat.WordBreak
        End With
        Me.AddChild(CType(lblDesc, UIControl))

        'lstTechs initial props
        lstTechs = New UIListBox(oUILib)
        With lstTechs
            .ControlName = "lstTechs"
            .Left = 5
            .Top = 150 '100
            .Width = 502
            .Height = 310
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Courier New", 9.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
            .sHeaderRow = "Project Name".PadRight(32, " "c) & "Benefits"
        End With
        Me.AddChild(CType(lstTechs, UIControl))

        'btnSubmit initial props
        btnSubmit = New UIButton(oUILib)
        With btnSubmit
            .ControlName = "btnSubmit"
            .Left = 400
            .Top = 470
            .Width = 100
            .Height = 32
            .Enabled = True
            .Visible = True
            .Caption = "Submit"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnSubmit, UIControl))

        'btnClose initial props
        btnClose = New UIButton(oUILib)
        With btnClose
            .ControlName = "btnClose"
            .Left = Me.Width - 26
            .Top = 1
            .Width = 24
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "X"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnClose, UIControl))

        'lblSelected initial props
        lblSelected = New UILabel(oUILib)
        With lblSelected
            .ControlName = "lblSelected"
            .Left = 5
            .Top = 470
            .Width = 400
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Selected Guaranteed Project: None"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Left 'CType(1, DrawTextFormat) 'Or DrawTextFormat.WordBreak
        End With
        Me.AddChild(CType(lblSelected, UIControl))

        Dim yMsg(1) As Byte
        System.BitConverter.GetBytes(GlobalMessageCode.eGetSpecTechGuaranteeList).CopyTo(yMsg, 0)
        MyBase.moUILib.SendMsgToPrimary(yMsg)

        MyBase.moUILib.RemoveWindow(Me.ControlName)
        MyBase.moUILib.AddWindow(Me)
    End Sub

    Private Sub btnClose_Click(ByVal sName As String) Handles btnClose.Click
        MyBase.moUILib.RemoveWindow(Me.ControlName)
    End Sub

    Public Sub HandleSpecialTechListMsg(ByVal yData() As Byte)
        Dim lPos As Int32 = 2
        Dim lCnt As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        Dim lSelID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        lstTechs.Clear()

        For X As Int32 = 0 To lCnt - 1
            Dim lID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            Dim iLen As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
            Dim sName As String = GetStringFromBytes(yData, lPos, iLen) : lPos += CInt(iLen)
            iLen = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
            Dim sBenefit As String = GetStringFromBytes(yData, lPos, iLen) : lPos += CInt(iLen)

            If sName.Length > 31 Then
                sName = sName.Substring(0, 28) & "..."
            End If
            If lID = lSelID Then
                lblSelected.Caption = "Selected Guaranteed Project: " & sName
            End If
            lstTechs.AddItem(sName.PadRight(32, " "c) & sBenefit, False)
            lstTechs.ItemData(lstTechs.NewIndex) = lID
        Next X

    End Sub

    Private Sub btnSubmit_Click(ByVal sName As String) Handles btnSubmit.Click

        If lstTechs.ListIndex > -1 Then
            Dim sTech As String = lstTechs.List(lstTechs.ListIndex)
            sTech = sTech.Substring(0, 32).Trim

            Dim oMsgBox As New frmMsgBox(goUILib, "You are about to select " & sTech & " as your guaranteed link. Are you sure?", MsgBoxStyle.YesNo, "Confirmation")
            oMsgBox.Visible = True
            AddHandler oMsgBox.DialogClosed, AddressOf Submit_Result
        Else
            MyBase.moUILib.AddNotification("Select a tech in the list before clicking Submit.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
        End If
    End Sub

    Private Sub Submit_Result(ByVal lResult As MsgBoxResult)
        If lResult = MsgBoxResult.Yes Then
            'submit the tech listing...
            If lstTechs.ListIndex > -1 Then
                Dim lID As Int32 = lstTechs.ItemData(lstTechs.ListIndex)
                Dim yMsg(5) As Byte
                System.BitConverter.GetBytes(GlobalMessageCode.eSetSpecTechGuarantee).CopyTo(yMsg, 0)
                System.BitConverter.GetBytes(lID).CopyTo(yMsg, 2)
                MyBase.moUILib.SendMsgToPrimary(yMsg)
                MyBase.moUILib.RemoveWindow(Me.ControlName)
            End If
        End If
    End Sub


End Class