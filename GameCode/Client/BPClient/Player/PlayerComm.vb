Option Strict On

Public Class PlayerComm
    Inherits Base_GUID

    Public Structure WPAttachment
        Public AttachNumber As Byte
        Public EnvirID As Int32
        Public EnvirTypeID As Int16
        Public LocX As Int32
        Public LocZ As Int32
        Public sWPName As String

        Public Sub JumpToAttachment()

            If (EnvirID < 1 OrElse EnvirTypeID < 1) AndAlso goCurrentPlayer.yPlayerPhase = 255 Then Return

            If EnvirID > 500000000 AndAlso EnvirTypeID = ObjectType.ePlanet AndAlso goCurrentPlayer.yPlayerPhase <> 0 Then
                If goUILib Is Nothing = False Then goUILib.AddNotification("That location no longer exists.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return
            End If

            If EnvirTypeID = ObjectType.ePlanet OrElse EnvirTypeID = ObjectType.eSolarSystem Then
                If goCurrentEnvir Is Nothing = False Then

					If NewTutorialManager.TutorialOn = True Then
						NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eEmailWaypointJumpedTo, -1, -1, -1, "")
					End If

                    PlayerComm.l_JumpToX = LocX
					PlayerComm.l_JumpToZ = LocZ


                    If goCurrentPlayer.yPlayerPhase <> 255 AndAlso goCurrentPlayer.yPlayerPhase <> 1 Then
                        FinishJumpToEvent()
                        Return
                    End If


                    If goCurrentEnvir.ObjectID = EnvirID AndAlso goCurrentEnvir.ObjTypeID = EnvirTypeID Then
                        'we're already here... nothing to do...
                        FinishJumpToEvent()
                    Else
                        goUILib.lUISelectState = UILib.eSelectState.eEmailAttachmentJumpTo
                        frmMain.ForceChangeEnvironment(EnvirID, EnvirTypeID)
                    End If
                End If
            End If
        End Sub
    End Structure

    Public Shared l_JumpToX As Int32
    Public Shared l_JumpToZ As Int32
    Public Shared Sub FinishJumpToEvent()

        goCamera.TrackingIndex = -1

        If goCurrentEnvir Is Nothing = False Then
            If goCurrentEnvir.ObjTypeID = ObjectType.ePlanet Then
                glCurrentEnvirView = CurrentView.ePlanetMapView
            Else : glCurrentEnvirView = CurrentView.eSystemMapView1
            End If

            With goCamera
                .mlCameraAtX = 0 : .mlCameraAtY = 0 : .mlCameraAtZ = 0
                .mlCameraX = 0 : .mlCameraY = 1000 : .mlCameraZ = -1000
            End With
 
            'Ok, now find that item...
            If goCurrentEnvir.ObjTypeID = ObjectType.ePlanet Then
                glCurrentEnvirView = CurrentView.ePlanetView
                goCamera.mlCameraY = 1700
            Else
                glCurrentEnvirView = CurrentView.eSystemView
                'goCamera.mlCameraY = 3000
                goCamera.mlCameraY = 7000 'RTP: 3000 was just too zoomed in for waypoints.  90% of the time after going to a WP, first thing I do is zoom back some
            End If

            With goCamera
                .mlCameraAtX = CInt(l_JumpToX) : .mlCameraAtY = 0 : .mlCameraAtZ = CInt(l_JumpToZ)
                .mlCameraX = .mlCameraAtX : .mlCameraZ = .mlCameraAtZ - 500
            End With

            Try
                If goCurrentEnvir.ObjTypeID = ObjectType.ePlanet Then
					goCamera.mlCameraY = CInt(CType(goCurrentEnvir.oGeoObject, Planet).GetHeightAtPoint(goCamera.mlCameraAtX, goCamera.mlCameraAtZ, True)) + muSettings.lPlanetViewCameraY

                    goCamera.mlCameraX += muSettings.lPlanetViewCameraX
                    goCamera.mlCameraZ += muSettings.lPlanetViewCameraZ
                End If
            Catch
            End Try
        End If

    End Sub

    Public sSender As String
    Public lSendOn As Int32
    Public EncryptLevel As Short
    Public bMsgRead As Boolean = False

    Public SentToList As String
    Public BCCList As String
    Public MsgTitle As String
    Public MsgBody As String

	Public PCF_ID As Int32

	Public bRequestedDetails As Boolean = False
	Public lLastMsgUpdate As Int32 = 0

    Public uAttachments() As WPAttachment
    Public lAttachmentUB As Int32 = -1

    Public Sub FillFromMsg(ByRef yData() As Byte)
        Dim lPos As Int32 = 2       '2 byte msg code

		With Me
            .lLastMsgUpdate = glCurrentCycle
            .bRequestedDetails = True
			.ObjectID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			.ObjTypeID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

			.sSender = GetStringFromBytes(yData, lPos, 20) : lPos += 20
			.lSendOn = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			.EncryptLevel = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
			.bMsgRead = yData(lPos) <> 0 : lPos += 1
			.PCF_ID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

			Dim lTemp As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			If lTemp <> 0 Then .SentToList = GetStringFromBytes(yData, lPos, lTemp) : lPos += lTemp
			lTemp = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			If lTemp <> 0 Then .MsgTitle = GetStringFromBytes(yData, lPos, lTemp) : lPos += lTemp
			lTemp = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			If lTemp <> 0 Then .MsgBody = GetStringFromBytes(yData, lPos, lTemp) : lPos += lTemp
			lTemp = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			If lTemp <> 0 Then .BCCList = GetStringFromBytes(yData, lPos, lTemp) : lPos += lTemp

			lTemp = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			lAttachmentUB = lTemp - 1
			ReDim uAttachments(lAttachmentUB)
			For X As Int32 = 0 To lAttachmentUB
				With uAttachments(X)
					.AttachNumber = yData(lPos) : lPos += 1
					.EnvirID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
					.EnvirTypeID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
					.LocX = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
					.LocZ = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
					.sWPName = GetStringFromBytes(yData, lPos, 20) : lPos += 20
				End With
			Next X
		End With

    End Sub

End Class
