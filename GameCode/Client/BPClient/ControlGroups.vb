Public Class ControlGroups
    Private Class ControlGroup
        Public lCtrlGroupNum As Int32
        Public lEnvirID As Int32
        Public iEnvirTypeID As Int16

        Public lIDs() As Int32
        Public iTypeIDs() As Int16
        Public lEntityUB As Int32 = -1

        Public Sub Clear()
            lEntityUB = -1
            ReDim lIDs(-1)
            ReDim iTypeIDs(-1)
        End Sub

        Public Function GetSaveBytes() As Byte()
            'ok... it is saved as...
            '4, 4, 2, 4 (cnt), 6()...
            Dim yResp(13 + ((lEntityUB + 1) * 6)) As Byte
            Dim lPos As Int32 = 0

            System.BitConverter.GetBytes(lCtrlGroupNum).CopyTo(yResp, lPos) : lPos += 4
            System.BitConverter.GetBytes(lEnvirID).CopyTo(yResp, lPos) : lPos += 4
            System.BitConverter.GetBytes(iEnvirTypeID).CopyTo(yResp, lPos) : lPos += 2
            System.BitConverter.GetBytes(lEntityUB + 1).CopyTo(yResp, lPos) : lPos += 4

            For X As Int32 = 0 To lEntityUB
                System.BitConverter.GetBytes(lIDs(X)).CopyTo(yResp, lPos) : lPos += 4
                System.BitConverter.GetBytes(iTypeIDs(X)).CopyTo(yResp, lPos) : lPos += 2
            Next X

            Return yResp
        End Function

        Public Sub SelectObjects(ByVal bClearSelection As Boolean, ByVal bTrack As Boolean)
            If goCurrentEnvir Is Nothing Then Return 'should be checked before calling this...

			Dim lTrackIdx As Int32 = -1

            With goCurrentEnvir
                'First, clear all selections...
				If bClearSelection = True Then
					.DeselectAll()
				End If

                For X As Int32 = 0 To .lEntityUB
                    If .lEntityIdx(X) <> -1 Then
                        For Y As Int32 = 0 To lEntityUB
                            If lIDs(Y) = .lEntityIdx(X) AndAlso iTypeIDs(Y) = .oEntity(X).ObjTypeID Then
								.oEntity(X).bSelected = True
								If bTrack = True AndAlso goCamera Is Nothing = False Then
									lTrackIdx = X
									'goCamera.TrackingIndex = X
									bTrack = False
								End If
								If .oEntity(X).OwnerID <> glPlayerID Then lTrackIdx = -1
                                Exit For    'Exit the for y loop
                            End If
                        Next Y
                    End If
                Next X

                If lTrackIdx <> -1 AndAlso goCamera Is Nothing = False Then
                    goCamera.SimplyPlaceCamera(CInt(.oEntity(lTrackIdx).LocX), CInt(.oEntity(lTrackIdx).LocY), CInt(.oEntity(lTrackIdx).LocZ))
                End If
			End With

            
            'goCamera.TrackingIndex = lTrackIdx
        End Sub
    End Class

    Private mlControlGroupUB As Int32 = -1
    Private moControlGroups() As ControlGroup
    Private mlLastControlGroup As Int32 = -1
    Private mlLastControlGroupSelect As Int32 = -1

    Private Sub SetInternalControlGroup(ByVal lCtrlGroupNum As Int32, ByVal lEnvirID As Int32, ByVal iEnvirTypeID As Int16, ByVal lIDs() As Int32, ByVal iTypeIDs() As Int16)
        'find out if that control group already exists
        Dim lIdx As Int32 = -1
        Dim XMax As Int32 = Math.Max(lIDs.Length, iTypeIDs.Length) - 1

        For X As Int32 = 0 To mlControlGroupUB
            If moControlGroups(X).lCtrlGroupNum = lCtrlGroupNum AndAlso moControlGroups(X).lEnvirID = lEnvirID AndAlso moControlGroups(X).iEnvirTypeID = iEnvirTypeID Then
                'Ok found it
                lIdx = X
                Exit For
            End If
        Next X

        If lIdx = -1 Then
            'create a new control group
            mlControlGroupUB += 1
            ReDim Preserve moControlGroups(mlControlGroupUB)
            moControlGroups(mlControlGroupUB) = New ControlGroup()
            lIdx = mlControlGroupUB
        End If

        With moControlGroups(lIdx)
            .Clear()

            .lEnvirID = lEnvirID
            .iEnvirTypeID = iEnvirTypeID
            .lCtrlGroupNum = lCtrlGroupNum
            .lEntityUB = XMax
            ReDim .lIDs(.lEntityUB)
            ReDim .iTypeIDs(.lEntityUB)

            For X As Int32 = 0 To .lEntityUB
                .lIDs(X) = lIDs(X)
                .iTypeIDs(X) = iTypeIDs(X)
            Next X
        End With
    End Sub

    Public Sub SetControlGroup(ByVal lCtrlGroupNum As Int32)
		If goCurrentEnvir Is Nothing Then Return

		If NewTutorialManager.TutorialOn = True Then
			NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eControlGroupAssigned, lCtrlGroupNum, -1, -1, "")
		End If

        'find out if that control group already exists
        Dim lIdx As Int32 = -1
        Dim lEnvirID As Int32 = goCurrentEnvir.ObjectID
        Dim iEnvirTypeID As Int16 = goCurrentEnvir.ObjTypeID

        For X As Int32 = 0 To mlControlGroupUB
            If moControlGroups(X).lCtrlGroupNum = lCtrlGroupNum AndAlso moControlGroups(X).lEnvirID = lEnvirID AndAlso moControlGroups(X).iEnvirTypeID = iEnvirTypeID Then
                'Ok found it
                lIdx = X
                Exit For
            End If
        Next X

        If lIdx = -1 Then
            'create a new control group
            mlControlGroupUB += 1
            ReDim Preserve moControlGroups(mlControlGroupUB)
            moControlGroups(mlControlGroupUB) = New ControlGroup()
            lIdx = mlControlGroupUB
        End If

        With moControlGroups(lIdx)
            .Clear()

            .lEnvirID = lEnvirID
            .iEnvirTypeID = iEnvirTypeID
            .lCtrlGroupNum = lCtrlGroupNum

            For X As Int32 = 0 To goCurrentEnvir.lEntityUB
                If goCurrentEnvir.lEntityIdx(X) <> -1 AndAlso goCurrentEnvir.oEntity(X).bSelected = True Then
                    .lEntityUB += 1
                    ReDim Preserve .lIDs(.lEntityUB)
                    ReDim Preserve .iTypeIDs(.lEntityUB)

                    .lIDs(.lEntityUB) = goCurrentEnvir.oEntity(X).ObjectID
                    .iTypeIDs(.lEntityUB) = goCurrentEnvir.oEntity(X).ObjTypeID
                End If
            Next X

        End With
    End Sub

    Public Sub ClearControlGroup(ByVal lCtrlGroupNum As Int32)
        If goCurrentEnvir Is Nothing Then Return

        Dim lEnvirID As Int32 = goCurrentEnvir.ObjectID
        Dim iEnvirTypeID As Int16 = goCurrentEnvir.ObjTypeID

        For X As Int32 = 0 To mlControlGroupUB
            If moControlGroups(X).lCtrlGroupNum = lCtrlGroupNum AndAlso moControlGroups(X).lEnvirID = lEnvirID AndAlso moControlGroups(X).iEnvirTypeID = iEnvirTypeID Then
                moControlGroups(X).Clear()
                Exit For
            End If
        Next X
    End Sub

    Public Sub SelectControlGroup(ByVal lCtrlGroupNum As Int32, ByVal bClearSelection As Boolean)
		If goCurrentEnvir Is Nothing Then Return

		If NewTutorialManager.TutorialOn = True Then
			NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eControlGroupSelected, lCtrlGroupNum, -1, -1, "")
		End If

        Dim lEnvirID As Int32 = goCurrentEnvir.ObjectID
        Dim iEnvirTypeID As Int16 = goCurrentEnvir.ObjTypeID

        Dim bCenter As Boolean = False
        If lCtrlGroupNum = mlLastControlGroup AndAlso glCurrentCycle - mlLastControlGroupSelect < 15 Then
            bCenter = True

            If goCurrentEnvir.ObjTypeID = ObjectType.ePlanet Then
                If glCurrentEnvirView <> CurrentView.ePlanetView Then glCurrentEnvirView = CurrentView.ePlanetView
            ElseIf goCurrentEnvir.ObjTypeID = ObjectType.eSolarSystem Then
                If glCurrentEnvirView <> CurrentView.eSystemView Then
                    glCurrentEnvirView = CurrentView.eSystemView
                    goCamera.mlCameraY = 1000
                    goCamera.mlCameraAtY = 0
                End If
            End If
        End If
        mlLastControlGroup = lCtrlGroupNum
        mlLastControlGroupSelect = glCurrentCycle

        'Find our control group
        For X As Int32 = 0 To mlControlGroupUB
            If moControlGroups(X).lCtrlGroupNum = lCtrlGroupNum AndAlso moControlGroups(X).lEnvirID = lEnvirID AndAlso moControlGroups(X).iEnvirTypeID = iEnvirTypeID Then
                'Ok, found the group, call the control group's select routine
                moControlGroups(X).SelectObjects(bClearSelection, bCenter)

                'Now, call the SetControlGroup method
                Me.SetControlGroup(lCtrlGroupNum)
                Exit For
            End If
        Next X
    End Sub

    Public Sub New()
        'load our groups
        Dim sFile As String = AppDomain.CurrentDomain.BaseDirectory
        If sFile.EndsWith("\") = False Then sFile &= "\"
        If goCurrentPlayer Is Nothing Then 'Critical failure, load backup file.
            sFile &= "ctrlgrp.dat"
        Else
            sFile &= goCurrentPlayer.PlayerName & ".ctr"
            If My.Computer.FileSystem.FileExists(sFile) = False Then 'Check for backup file
                sFile = AppDomain.CurrentDomain.BaseDirectory
                If sFile.EndsWith("\") = False Then sFile &= "\"
                sFile &= "ctrlgrp.dat"
            End If
        End If
        If My.Computer.FileSystem.FileExists(sFile) = True Then
            Dim fsFile As IO.FileStream = New IO.FileStream(sFile, IO.FileMode.Open)
            Dim oReader As IO.BinaryReader = New IO.BinaryReader(fsFile)
            Dim lPos As Int32 = 0

            Dim lCtrlGroupNum As Int32
            Dim lEnvirID As Int32
            Dim iEnvirTypeID As Int16
            Dim lCnt As Int32
            Try
                While fsFile.Position < fsFile.Length
                    Dim yHdr() As Byte = oReader.ReadBytes(14)
                    If yHdr Is Nothing = False Then
                        lCtrlGroupNum = System.BitConverter.ToInt32(yHdr, 0)
                        lEnvirID = System.BitConverter.ToInt32(yHdr, 4)
                        iEnvirTypeID = System.BitConverter.ToInt16(yHdr, 8)
                        lCnt = System.BitConverter.ToInt32(yHdr, 10)

                        Dim yVals() As Byte = oReader.ReadBytes(6 * lCnt)
                        Dim lIDs(lCnt - 1) As Int32
                        Dim iTypeIDs(lCnt - 1) As Int16
                        lPos = 0
                        For X As Int32 = 0 To lCnt - 1
                            lIDs(X) = System.BitConverter.ToInt32(yVals, lPos) : lPos += 4
                            iTypeIDs(X) = System.BitConverter.ToInt16(yVals, lPos) : lPos += 2
                        Next X

                        SetInternalControlGroup(lCtrlGroupNum, lEnvirID, iEnvirTypeID, lIDs, iTypeIDs)
                    Else
                        Exit While
                    End If
                End While
            Catch
            End Try
            oReader.Close()
            oReader = Nothing
            fsFile.Close()
            fsFile.Dispose()
            fsFile = Nothing
        End If
    End Sub

    Protected Overrides Sub Finalize()
        SaveGroups()

        Erase moControlGroups
        mlControlGroupUB = -1

        MyBase.Finalize()
    End Sub

    Public Sub SaveGroups()
        'Now, save our control groups...
        If goControlGroups Is Nothing OrElse goControlGroups.mlControlGroupUB = -1 Then Return
        Dim yGroupSave() As Byte
        Dim sFile As String = AppDomain.CurrentDomain.BaseDirectory
        If sFile.EndsWith("\") = False Then sFile &= "\"
        If goCurrentPlayer Is Nothing Then
            sFile &= "ctrlgrp.dat"
        Else
            sFile &= goCurrentPlayer.PlayerName & ".ctr"
        End If

        Dim fsFile As IO.FileStream = New IO.FileStream(sFile, IO.FileMode.Create)

        For X As Int32 = 0 To mlControlGroupUB
            If moControlGroups(X).lEntityUB > -1 Then
                yGroupSave = moControlGroups(X).GetSaveBytes()
                'yGroupSave = EncBytes(yGroupSave)
                fsFile.Write(yGroupSave, 0, yGroupSave.Length)
            End If
        Next X
        fsFile.Close()
        fsFile.Dispose()
        'Delete old legacy file
        'sFile = AppDomain.CurrentDomain.BaseDirectory
        'If sFile.EndsWith("\") = False Then sFile &= "\"
        'sFile &= "ctrlgrp.dat"
        'If My.Computer.FileSystem.FileExists(sFile) = True Then
        '    My.Computer.FileSystem.DeleteFile(sFile)
        'End If

    End Sub

End Class