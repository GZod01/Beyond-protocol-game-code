Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

Public Class WarpointRewards

    Private Structure Reward
        Public lLocX As Int32
        Public lLocZ As Int32
        Public fLocY As Single
        Private fAlpha As Single
        Public lStringRectID() As Int32
        Public lAlpha As Int32
        Public Active As Boolean

        Public lR As Int32
        Public lG As Int32
        Public lB As Int32

        Public Sub Update(ByVal lElapsed As Int32)
            fLocY += lElapsed * RewardLabelRiseRate
            fAlpha -= (lElapsed * RewardLabelFadeRate)

            lAlpha = CInt(fAlpha)
            If lAlpha < 0 Then lAlpha = 0
            If lAlpha > 255 Then lAlpha = 255
        End Sub
        Public Sub SetAlpha(ByVal lVal As Int32)
            lAlpha = lVal
            fAlpha = lAlpha
        End Sub

    End Structure

    Private muRewards(-1) As Reward
    Private mlLastUpdate As Int32 = Int32.MinValue

    Private Shared RewardLabelFadeRate As Single = 5.0F
    Private Shared RewardLabelRiseRate As Single = 15.0F

    Public Sub AddReward(ByVal lX As Int32, ByVal lY As Int32, ByVal lZ As Int32, ByVal lReward As Int32, ByVal bRewarded As Boolean)
        Dim lIdx As Int32 = -1
        For X As Int32 = 0 To muRewards.GetUpperBound(0)
            If muRewards(X).Active = False Then
                lIdx = X
                Exit For
            End If
        Next X
        If lIdx = -1 Then
            ReDim Preserve muRewards(muRewards.GetUpperBound(0) + 1)
            lIdx = muRewards.GetUpperBound(0)
        End If
        With muRewards(lIdx)
            .lLocX = lX
            .lLocZ = lZ
            .fLocY = lY
            .SetAlpha(255)

            If bRewarded = True Then
                'green
                .lR = 64
                .lG = 255
                .lB = 64
            Else
                'blue
                .lR = 64
                .lG = 128
                .lB = 255
            End If

            .lStringRectID = BPSprite.GetStringRectIDs("+" & lReward.ToString("#,##0"))
            .Active = True
        End With
    End Sub

    Public Sub RenderRewards(ByVal bUpdateNoRender As Boolean)
        If mlLastUpdate = Int32.MinValue Then
            mlLastUpdate = glCurrentCycle
            Return
        End If
        Dim lElapsed As Int32 = glCurrentCycle - mlLastUpdate
        mlLastUpdate = glCurrentCycle

        RewardLabelFadeRate = muSettings.RewardLabelFadeRate
        RewardLabelRiseRate = muSettings.RewardLabelRiseRate

        Dim oTexture As Texture = goResMgr.GetTexture("RewardText.dds", GFXResourceManager.eGetTextureType.UserInterface, "BPMisc.pak")

        Try
            Dim oSprite As BPSprite = Nothing
            Dim bBegun As Boolean = False

            For X As Int32 = 0 To muRewards.GetUpperBound(0)
                If muRewards(X).Active = True Then
                    If lElapsed > 0 Then
                        muRewards(X).Update(lElapsed)
                        If muRewards(X).lAlpha = 0 Then
                            muRewards(X).Active = False
                            Continue For
                        End If
                    End If

                    If bBegun = False Then
                        oSprite = New BPSprite()
                        oSprite.BeginRender(0, oTexture, GFXEngine.moDevice)
                        bBegun = True
                    End If

                    With muRewards(X)
                        oSprite.DrawStringIn3DSpace(.lStringRectID, .lLocX, .fLocY, .lLocZ, oTexture, System.Drawing.Color.FromArgb(.lAlpha, .lR, .lG, .lB), 40.0F)
                    End With
                End If
            Next X

            If bBegun = True Then
                With GFXEngine.moDevice
                    .SetTextureStageState(0, TextureStageStates.AlphaArgument1, TextureArgument.TextureColor)
                    .SetTextureStageState(0, TextureStageStates.AlphaArgument2, TextureArgument.Diffuse)
                    .SetTextureStageState(0, TextureStageStates.AlphaOperation, TextureOperation.Modulate)
                    .RenderState.SourceBlend = Blend.SourceAlpha
                    .RenderState.DestinationBlend = Blend.InvSourceAlpha
                    .RenderState.AlphaBlendEnable = True
                    .RenderState.CullMode = Cull.None
                    .RenderState.ZBufferEnable = False
                    .RenderState.Lighting = False
                End With
                oSprite.EndRenderNoStateChange()

                With GFXEngine.moDevice
                    .SetTextureStageState(0, TextureStageStates.AlphaArgument1, TextureArgument.TextureColor)
                    .SetTextureStageState(0, TextureStageStates.AlphaArgument2, TextureArgument.Current)
                    .SetTextureStageState(0, TextureStageStates.AlphaOperation, TextureOperation.SelectArg1)
                    .RenderState.SourceBlend = Blend.SourceAlpha 'Blend.One
                    .RenderState.DestinationBlend = Blend.InvSourceAlpha 'Blend.Zero
                    .RenderState.AlphaBlendEnable = True
                    .RenderState.ZBufferEnable = True
                    .RenderState.CullMode = Cull.CounterClockwise
                    .RenderState.Lighting = True
                End With
            End If
        Catch
        End Try

    End Sub

    Public Sub ClearAll()
        ReDim muRewards(-1)
    End Sub
End Class
