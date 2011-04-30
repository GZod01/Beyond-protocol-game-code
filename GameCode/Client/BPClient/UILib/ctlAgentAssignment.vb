Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Interface created from Interface Builder
Public Class ctlAgentAssignment
    Inherits UIWindow

	'Private lblAssignment As UILabel
	'Private lnDiv1 As UILine
    Private picAgentPic As UIWindow
    Private lblSkill1 As UILabel
    'Private lblOr1 As UILabel
    'Private lblSkill2 As UILabel
    'Private lblOr2 As UILabel
    'Private lblSkill3 As UILabel
    'Private lblOr3 As UILabel
    'Private lblSkill4 As UILabel 

    Private WithEvents btnAssign As UIButton

    Public Event AssignButtonClicked(ByRef oSkillSetSkill As SkillSet_Skill, ByRef oSender As ctlAgentAssignment)

    Private moAgent As Agent = Nothing
    Private moSkillSet_Skill As SkillSet_Skill

    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)

        'ctlAgentAssignment initial props
        With Me
            .ControlName = "ctlAgentAssignment"
            .Left = 345
            .Top = 312
            .Width = 150
            .Height = 185
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = False
            .BorderLineWidth = 2
			.bRoundedBorder = True
			.Moveable = False
        End With

        'btnAssign initial props
        btnAssign = New UIButton(oUILib)
        With btnAssign
            .ControlName = "btnAssign"
			.Left = 2
            .Top = 5
			.Width = Me.Width - 4
            .Height = 32
            .Enabled = True
            .Visible = True
            .Caption = "Assign"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnAssign, UIControl))

		''lblAssignment initial props
		'lblAssignment = New UILabel(oUILib)
		'With lblAssignment
		'    .ControlName = "lblAssignment"
		'    .Left = 0
		'    .Top = 2
		'    .Width = 150
		'    .Height = 35
		'    .Enabled = True
		'    .Visible = False
		'    .Caption = "Unassigned"
		'    .ForeColor = muSettings.InterfaceBorderColor
		'    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
		'    .DrawBackImage = False
		'    .FontFormat = DrawTextFormat.Center Or DrawTextFormat.VerticalCenter Or DrawTextFormat.WordBreak
		'End With
		'Me.AddChild(CType(lblAssignment, UIControl))

		''lnDiv1 initial props
		'lnDiv1 = New UILine(oUILib)
		'With lnDiv1
		'    .ControlName = "lnDiv1"
		'    .Left = 1
		'    .Top = 40
		'    .Width = Me.Width - 1
		'    .Height = 0
		'    .Enabled = True
		'    .Visible = True
		'    .BorderColor = muSettings.InterfaceBorderColor
		'End With
		'Me.AddChild(CType(lnDiv1, UIControl))

        'picAgentPic initial props
        picAgentPic = New UIWindow(oUILib)
        With picAgentPic
            .ControlName = "picAgentPic"
            .Left = 11
            .Top = 50
            .Width = 128
            .Height = 128
            .Enabled = True
            .Visible = False
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = False
            .Moveable = False
            .DrawBorder = False
        End With
        Me.AddChild(CType(picAgentPic, UIControl))

        'lblSkill1 initial props
        lblSkill1 = New UILabel(oUILib)
        With lblSkill1
            .ControlName = "lblSkill1"
            .Left = 11
            .Top = 50
            .Width = 128
            .Height = 128 '18
            .Enabled = True
            .Visible = False
            .Caption = ""
            .ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
			.FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Center Or DrawTextFormat.WordBreak
        End With
        Me.AddChild(CType(lblSkill1, UIControl))

        ''lblOr1 initial props
        'lblOr1 = New UILabel(oUILib)
        'With lblOr1
        '    .ControlName = "lblOr1"
        '    .Left = 11
        '    .Top = 65
        '    .Width = 128
        '    .Height = 18
        '    .Enabled = True
        '    .Visible = False
        '    .Caption = "OR"
        '    .ForeColor = muSettings.InterfaceBorderColor
        '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
        '    .DrawBackImage = False
        '    .FontFormat = CType(5, DrawTextFormat)
        'End With
        'Me.AddChild(CType(lblOr1, UIControl))

        ''lblSkill2 initial props
        'lblSkill2 = New UILabel(oUILib)
        'With lblSkill2
        '    .ControlName = "lblSkill2"
        '    .Left = 11
        '    .Top = 80
        '    .Width = 128
        '    .Height = 18
        '    .Enabled = True
        '    .Visible = False
        '    .Caption = "Engineering"
        '    .ForeColor = muSettings.InterfaceBorderColor
        '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
        '    .DrawBackImage = False
        '    .FontFormat = CType(5, DrawTextFormat)
        'End With
        'Me.AddChild(CType(lblSkill2, UIControl))

        ''lblOr2 initial props
        'lblOr2 = New UILabel(oUILib)
        'With lblOr2
        '    .ControlName = "lblOr2"
        '    .Left = 11
        '    .Top = 95
        '    .Width = 128
        '    .Height = 18
        '    .Enabled = True
        '    .Visible = False
        '    .Caption = "OR"
        '    .ForeColor = muSettings.InterfaceBorderColor
        '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
        '    .DrawBackImage = False
        '    .FontFormat = CType(5, DrawTextFormat)
        'End With
        'Me.AddChild(CType(lblOr2, UIControl))

        ''lblSkill3 initial props
        'lblSkill3 = New UILabel(oUILib)
        'With lblSkill3
        '    .ControlName = "lblSkill3"
        '    .Left = 11
        '    .Top = 110
        '    .Width = 128
        '    .Height = 18
        '    .Enabled = True
        '    .Visible = False
        '    .Caption = "Poisons"
        '    .ForeColor = muSettings.InterfaceBorderColor
        '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
        '    .DrawBackImage = False
        '    .FontFormat = CType(5, DrawTextFormat)
        'End With
        'Me.AddChild(CType(lblSkill3, UIControl))

        ''lblOr3 initial props
        'lblOr3 = New UILabel(oUILib)
        'With lblOr3
        '    .ControlName = "lblOr3"
        '    .Left = 11
        '    .Top = 125
        '    .Width = 128
        '    .Height = 18
        '    .Enabled = True
        '    .Visible = False
        '    .Caption = "OR"
        '    .ForeColor = muSettings.InterfaceBorderColor
        '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
        '    .DrawBackImage = False
        '    .FontFormat = CType(5, DrawTextFormat)
        'End With
        'Me.AddChild(CType(lblOr3, UIControl))

        ''lblSkill4 initial props
        'lblSkill4 = New UILabel(oUILib)
        'With lblSkill4
        '    .ControlName = "lblSkill4"
        '    .Left = 11
        '    .Top = 140
        '    .Width = 128
        '    .Height = 18
        '    .Enabled = True
        '    .Visible = False
        '    .Caption = "Athletics"
        '    .ForeColor = muSettings.InterfaceBorderColor
        '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
        '    .DrawBackImage = False
        '    .FontFormat = CType(5, DrawTextFormat)
        'End With
        'Me.AddChild(CType(lblSkill4, UIControl))

	End Sub
 
	Public Sub SetAsCoverAgent(ByRef oAgent As Agent)
		lblSkill1.Caption = "COVER AGENT"
		lblSkill1.Visible = True
		moAgent = oAgent
		Me.IsDirty = True
		Me.Visible = True
	End Sub

    Public Sub SetAgent(ByRef oAgent As Agent)
        moAgent = oAgent
        Me.IsDirty = True
        If moAgent Is Nothing = False Then
			'btnAssign.Visible = False
			'lblAssignment.Caption = moAgent.sAgentName
			'lblAssignment.Visible = True
            btnAssign.Caption = oAgent.sAgentName

            If moSkillSet_Skill Is Nothing = False AndAlso moSkillSet_Skill.oSkill Is Nothing = False Then
                Dim lID As Int32 = moSkillSet_Skill.oSkill.ObjectID
                Dim bHasNaturalTalents As Boolean = False
                Dim yRating As Byte = 0
                For X As Int32 = 0 To oAgent.SkillUB
                    If oAgent.Skills(X).ObjectID = lID Then
                        yRating = oAgent.SkillProf(X)
                        bHasNaturalTalents = False
                        Exit For
                    ElseIf oAgent.Skills(X).ObjectID = 36 Then      'naturally talented
                        bHasNaturalTalents = True
                        yRating = oAgent.SkillProf(X)
                    End If
                Next X

                Dim sFinal As String = moSkillSet_Skill.oSkill.SkillName
                If bHasNaturalTalents = True Then
                    sFinal &= vbCrLf & "Natural Talents"
                End If
                sFinal &= vbCrLf & "(" & yRating & ")"
                lblSkill1.Caption = sFinal
            End If

        Else
            btnAssign.Caption = "Assign"
            'btnAssign.Visible = True
            'lblAssignment.Visible = False
        End If
    End Sub

    Public Sub SetSkillSetSkill(ByRef oSkill As SkillSet_Skill)
        moSkillSet_Skill = oSkill
        If moSkillSet_Skill Is Nothing Then
            Me.Visible = False
        Else
            Me.Visible = True
            If moSkillSet_Skill.ToHitModifier < 0 Then
                lblSkill1.ForeColor = System.Drawing.Color.FromArgb(255, 255, 64, 64)
            ElseIf moSkillSet_Skill.ToHitModifier > 0 Then
                lblSkill1.ForeColor = System.Drawing.Color.FromArgb(255, 64, 255, 64)
            Else
                lblSkill1.ForeColor = muSettings.InterfaceBorderColor
            End If
			lblSkill1.Caption = moSkillSet_Skill.oSkill.SkillName '& " (" & moSkillSet_Skill.ToHitModifier.ToString & ")"
			lblSkill1.Visible = True
        End If
    End Sub

	Private Sub btnAssign_Click(ByVal sName As String) Handles btnAssign.Click
		RaiseEvent AssignButtonClicked(moSkillSet_Skill, Me)
    End Sub

    Private Sub ctlAgentAssignment_OnRender() Handles Me.OnRender
        If goCurrentPlayer Is Nothing = True Then Return
        If goCurrentPlayer.AgentUB = -1 Then Return

        If moAgent Is Nothing = False Then
            If AgentRenderer.goAgentRenderer Is Nothing Then AgentRenderer.goAgentRenderer = New AgentRenderer()

            Dim rcDest As Rectangle
            Dim oLoc As Point = GetAbsolutePosition()
            With rcDest
                .X = oLoc.X + picAgentPic.Left
                .Y = oLoc.Y + picAgentPic.Top
                .Width = picAgentPic.Width
                .Height = picAgentPic.Height
            End With
            AgentRenderer.goAgentRenderer.RenderAgent2(moAgent.ObjectID, rcDest, False, moAgent.bMale)
        End If

    End Sub

	Public ReadOnly Property SkillsetID() As Int32
		Get
			If moSkillSet_Skill Is Nothing OrElse moSkillSet_Skill.oSkillSet Is Nothing Then Return -1 Else Return moSkillSet_Skill.oSkillSet.SkillSetID
		End Get
	End Property

	Public ReadOnly Property SkillID() As Int32
		Get
			If moSkillSet_Skill Is Nothing OrElse moSkillSet_Skill.oSkill Is Nothing Then Return -1 Else Return moSkillSet_Skill.oSkill.ObjectID
		End Get
	End Property

End Class