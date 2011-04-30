Option Strict On


Module WindowOps


    Public Sub OpenWindow_Guild()
        BPMetrics.MetricMgr.AddActivity(BPMetrics.eKeyPress.eF11Key, 0, 0)
        If goCurrentPlayer Is Nothing Then Return
        If NewTutorialManager.TutorialOn = True Then
            If goUILib Is Nothing = False Then
                If goUILib.CommandAllowed(False, "F11") = False Then Return
            End If
        End If

        If goCurrentPlayer Is Nothing = False Then
            If goCurrentPlayer.oGuild Is Nothing Then

                Dim oFrm As frmGuildSearch = CType(goUILib.GetWindow("frmGuildSearch"), frmGuildSearch)
                If ofrm Is Nothing = False Then
                    If ofrm.Visible = True Then
                        goUILib.RemoveWindow(ofrm.ControlName)
                    Else : ofrm.Visible = True
                    End If
                Else
                    ofrm = New frmGuildSearch(goUILib)
                    ofrm.Visible = True
                End If
                ofrm = Nothing
            Else
                Dim ofrm As frmGuildMain = CType(goUILib.GetWindow("frmGuildMain"), frmGuildMain)
                If ofrm Is Nothing = False Then
                    If ofrm.Visible = True Then
                        goUILib.RemoveWindow(ofrm.ControlName)
                    Else : ofrm.Visible = True
                    End If
                Else
                    ofrm = New frmGuildMain(goUILib)
                    ofrm.Visible = True
                End If
                ofrm = Nothing
            End If
        End If
    End Sub

    Public Sub OpenWindow_RouteTemplates()
        Dim oFrm As frmRouteTemplate = CType(goUILib.GetWindow("frmRouteTemplate"), frmRouteTemplate)
        If oFrm Is Nothing = False Then
            If oFrm.Visible = True Then
                goUILib.RemoveWindow(oFrm.ControlName)
            Else : oFrm.Visible = True
            End If
        Else
            oFrm = New frmRouteTemplate(goUILib)
            oFrm.Visible = True
        End If
    End Sub
End Module
