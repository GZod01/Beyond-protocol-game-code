Public Structure RouteItem
	Public lDestID As Int32
	Public iDestTypeID As Int16
	Public lLocationID As Int32
	Public iLocationTypeID As Int16
	Public lLoadItemID As Int32
	Public iLoadItemTypeID As Int16
	Public lLocX As Int32
	Public lLocZ As Int32
    Public lOrderNum As Int32
    Public yFlag As Byte

	Public Sub FillFromMsg(ByRef yData() As Byte, ByRef lPos As Int32)
		lDestID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		iDestTypeID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
		lLocationID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		iLocationTypeID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
		lLoadItemID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		iLoadItemTypeID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
		lLocX = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        lLocZ = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        yFlag = yData(lPos) : lPos += 1
	End Sub
End Structure

Public Class BaseEntity
    Inherits RenderObject

    Public CurrentStatus As Int32 ' elUnitStatus
    Public CritList As Int32 'elUnitStatus     'indicates with bits marked 1 what elements of the entity are destroyed

    Public OwnerID As Int32         'owning player id
    Public sOverrideOwnerName As String = ""
    Public oUnitDef As EntityDef

    'Destination data...
    Public DestX As Int32
    Public DestY As Int32
    Public DestZ As Int32
    Public DestAngle As Int16

    Public TrueDestAngle As Int16       'TODO: need to update this into the movement engine here locally

    Public LastUpdateCycle As Int32     'the last cycle this unit was moved

    'Velocity Data
    Public VelX As Single
    Public VelZ As Single

    'this units total velocity
    Public TotalVelocity As Single

    Public bIsTurning As Boolean

    Public yArmorHP(3) As Byte
    Public yShieldHP As Byte = 100
    Public yStructureHP As Byte = 100
    Public bHPChanged As Boolean = True
    Public lLastHPUpdate As Int32 = Int32.MinValue

    Public HPTexture As Microsoft.DirectX.Direct3D.Texture

    Public lEngineFXEmitterID() As Int32
    Private mbEngineFXReady As Boolean = False

    Public lUnitSoundIdx As Int32 = -1

    'The following is gathered from a detail message
    Public bRequestedDetails As Boolean = False
    Public EntityName As String
    'Public Fuel_Cap As Int16
    Public Ammo_Cap As Int16
    Public Exp_Level As Byte
    Public iTargetingTactics As Int16
    Public iCombatTactics As Int32 = 8
    Public lTetherPointX As Int32 = Int32.MinValue
    Public lTetherPointZ As Int32 = Int32.MinValue
    'END OF DETAIL MESSAGE ITEMS
    Public lMaxWpnRngValue As Int32 = -1

#Region "  Route  "
	Public uRoute() As RouteItem
	Public lRouteUB As Int32 = -1
	Public bRoutePaused As Boolean = False
#End Region

    'OPTIMIZATION PURPOSES
    Public lEnvirEntityIdx As Int32

    Public yChangeEnvironment As Byte = 0
	'Public fOriginalDist As Single

    Public yRelID As Byte = Byte.MaxValue       'Use this for Optimization Purposes

    Private mbSelected As Boolean = False

    Public yProductionType As Byte

    Public colHangar As Collection
    Public colCargo As Collection
    Public ContentsLastUpdate As Int32 = 0      'cycle number of last contents update

    'Specifically for Facilities... received from the GetContents call
    Public CurrentWorkers As Int32 = 0
    Public CurrentColonists As Int32 = 0
    Public MaxWorkers As Int32 = 0
    Public MaxColonists As Int32 = 0
    Public PowerConsumption As Int32 = 0
    Public AutoLaunch As Boolean

    Public ProductionQueue() As ProdQueueItem
    Public ProdQueueUB As Int32 = -1

    'For Production Status management...
	Public bProducing As Boolean = False
	Public iProducingModelID As Int16 = -1
	Public iProductionAngle As Int16 = 0
	Private mlProdStarted As Int32
	Private mlFinishCycle As Int32
	Private mfLastPerc As Single

	Public lProducingID As Int32 = -1
	Public iProducingTypeID As Int16

	Public RallyX As Int32 = Int32.MinValue
	Public RallyZ As Int32 = Int32.MinValue

	'For Old Data system...
	Public bObjectDestroyed As Boolean = False

	'For Space Station Members...
	Public lChildUB As Int32 = -1
	Public oChild() As StationChild
	Public lChildIdx() As Int32
    Public bChildListUpdated As Boolean = False

    Public bGuildMember As Boolean = False

	Public lLastWeaponUpdate As Int32 = 0

	Public lTargetIdx(3) As Int32		'indices within the array for the entity being targeted
	Public lTargetMsg As Int32 = 0

	Private moProducingMesh As BaseMesh = Nothing
	Private miProducingMeshID As Int16 = -1
	Public ReadOnly Property GetProducingModel() As BaseMesh
		Get
			If moProducingMesh Is Nothing OrElse iProducingModelID <> miProducingMeshID Then
				moProducingMesh = Nothing
				miProducingMeshID = iProducingModelID
				If miProducingMeshID <> -1 Then
					moProducingMesh = goResMgr.GetMesh(miProducingMeshID)
				End If
			End If
			Return moProducingMesh
		End Get
	End Property

	Public Function GetChild(ByVal lID As Int32, ByVal iTypeID As Int16) As StationChild
		If lID = -1 Then Return Nothing
		For X As Int32 = 0 To lChildUB
			If lChildIdx(X) = lID AndAlso oChild(X).iChildTypeID = iTypeID Then Return oChild(X)
		Next X
		Return Nothing
	End Function

	Public Sub RemoveChild(ByVal lID As Int32, ByVal iTypeID As Int16)
		For X As Int32 = 0 To lChildUB
			If lChildIdx(X) = lID AndAlso oChild(X).iChildTypeID = iTypeID Then
				lChildIdx(X) = -1
				oChild(X) = Nothing
			End If
		Next X
	End Sub

	Public Sub ReceiveProductionStatusMsg(ByVal lServerCurrent As Int32, ByVal lServerFinish As Int32, ByVal fPerc As Single, ByVal piModelID As Int16)
		iProducingModelID = piModelID

		If fPerc = 255.0F Then
			bProducing = False
			lProducingID = -1
		Else
			bProducing = True

			If fPerc <> mfLastPerc Then

				'Now... need to do some conversion...
				Dim lCurrentDiff As Int32 = glCurrentCycle - lServerCurrent
				Dim lServerCyclesRemaining As Int32 = lServerFinish - lServerCurrent
				Dim fDivisor As Single = (100.0F - CSng(fPerc)) / 100.0F

				'Ok, adjust the server's finish cycle to my cycle
				mlFinishCycle = lServerFinish + lCurrentDiff

				'Now, we can calculate our started...
				Try
					If fDivisor = 0 Then mlProdStarted = 0 Else mlProdStarted = CInt(mlFinishCycle - (lServerCyclesRemaining / fDivisor))
				Catch
					mlProdStarted = 0
				End Try

				mfLastPerc = fPerc
			End If
		End If
	End Sub

	Public Function GetProductionStatus() As Single
		'returns a number from 0 to 1 or -1
		Try
			If bProducing = False Then Return -1

			Dim blTotalTime As Int64 = CLng(mlFinishCycle) - CLng(mlProdStarted)

			If blTotalTime = 0 Then Return 0.0F
			Return CSng((glCurrentCycle - mlProdStarted) / blTotalTime)
		Catch
			Return 0.0F
		End Try
    End Function

    Public Function GetProductionRemaining() As String
        Try
            If bProducing = False Then Return ""
            Dim blTotalTime As Int64 = CLng(mlFinishCycle) - CLng(mlProdStarted)
            Dim lProdFactor As Int32 = 766
            Dim blPoints As Int64 = (blTotalTime - (glCurrentCycle - mlProdStarted))
            'Dim fMult As Single = (GetMaxFactionBonus() * GetMaxOtherFactionBonus())
            'blPoints = CLng(blPoints * fMult)
            'Dim blSecondsToFinish As Int64 = CLng(blPoints / (lProdFactor * 30))
            Dim blSecondsToFinish As Int64 = CLng(blPoints / (lProdFactor))
            'Ok, format seconds to finish to minutes, hours, whatever
            Dim blHours As Int64 = blSecondsToFinish \ 3600L
            blSecondsToFinish -= (blHours * 3600L)
            Dim blMinutes As Int64 = blSecondsToFinish \ 60L
            blSecondsToFinish -= (blMinutes * 60L)
            If blSecondsToFinish <> 0 Then blMinutes += 1

            If blMinutes = 60 Then
                blMinutes = 0
                blHours += 1
            End If

            Dim sTimeString As String = ""
            If blHours > 24 Then
                Dim blDays As Int64 = blHours \ 24L
                sTimeString = blDays.ToString & " days"
                blHours -= (blDays * 24L)
            End If
            If blHours > 0 Then
                If sTimeString.Length > 0 Then sTimeString &= ", "
                sTimeString &= blHours.ToString & " hours"
            End If
            If blMinutes > 0 Then
                If sTimeString.Length > 0 Then sTimeString &= ", "
                sTimeString &= blMinutes.ToString & " minutes"
            End If
            If sTimeString = "" Then sTimeString = "< 1 minute"
            Return sTimeString
        Catch
            Return ""
        End Try
    End Function

	'End Production Status Management

	Public Property bSelected() As Boolean
		Get
			Return mbSelected
		End Get
		Set(ByVal Value As Boolean)
			If Value = False Then
				If HPTexture Is Nothing = False Then HPTexture.Dispose()
				HPTexture = Nothing
				bHPChanged = True
			Else
				If goUILib Is Nothing = False Then goUILib.AddSelection(lEnvirEntityIdx)

				If goTutorial Is Nothing = False AndAlso goTutorial.TutorialOn = True Then
					If Me.yProductionType = ProductionType.eFacility Then
						goTutorial.HandleTutorialTrigger(TutorialManager.TutorialTriggerType.SelectedEngineer)
					ElseIf Me.yProductionType = ProductionType.eEnlisted Then
						goTutorial.HandleTutorialTrigger(TutorialManager.TutorialTriggerType.BarracksSelected)
					End If
				End If
			End If
			mbSelected = Value
		End Set
	End Property

	Public Sub RefreshHPTexture(ByRef oDevice As Microsoft.DirectX.Direct3D.Device)
		If bHPChanged = True Then
			If HPTexture Is Nothing = False Then
				HPTexture.Dispose()
				HPTexture = Nothing
			End If
			Microsoft.DirectX.Direct3D.Device.IsUsingEventHandlers = False
			HPTexture = New Microsoft.DirectX.Direct3D.Texture(oDevice, 32, 32, 1, Microsoft.DirectX.Direct3D.Usage.None, Microsoft.DirectX.Direct3D.Format.A8R8G8B8, Microsoft.DirectX.Direct3D.Pool.Managed)
			Microsoft.DirectX.Direct3D.Device.IsUsingEventHandlers = True

			'Remarked out this code and added code above to better handle these and avoid a crash bug...
			'  not sure why the crash bug was occurring, will research that (whether it is driver specific, etc...)
			'If HPTexture Is Nothing OrElse HPTexture.Disposed Then
			'    oDevice.IsUsingEventHandlers = False
			'    HPTexture = New Microsoft.DirectX.Direct3D.Texture(oDevice, 32, 32, 1, Microsoft.DirectX.Direct3D.Usage.None, Microsoft.DirectX.Direct3D.Format.A8R8G8B8, Microsoft.DirectX.Direct3D.Pool.Managed)
			'    oDevice.IsUsingEventHandlers = True
			'End If

			Dim lPitch As Int32
			Dim oStream As Microsoft.DirectX.GraphicsStream = HPTexture.LockRectangle(0, System.Drawing.Rectangle.Empty, Microsoft.DirectX.Direct3D.LockFlags.Discard, lPitch)

			Dim yData() As Byte

			Dim lIdx As Int32
			Dim lPos As Int32
			Dim lVal As Int32
			Dim lDataSize As Int32 = ((4 * 32) * 32) - 1

			Dim LocX As Int32
			Dim LocY As Int32

			Dim lD As Int32

			Dim rcShield As Rectangle = Rectangle.FromLTRB(4, 26, 27, 31)
			Dim rcTopArmor As Rectangle = Rectangle.FromLTRB(4, 21, 27, 26)
			Dim rcLeftArmor As Rectangle = Rectangle.FromLTRB(4, 5, 9, 21)
			Dim rcBottomArmor As Rectangle = Rectangle.FromLTRB(4, 0, 27, 5)
			Dim rcRightArmor As Rectangle = Rectangle.FromLTRB(22, 5, 27, 21)
			Dim rcStructure As Rectangle = Rectangle.FromLTRB(9, 5, 22, 21)

			Dim yForeArmorValue As Byte = yArmorHP(0)
			Dim yLeftArmorValue As Byte = yArmorHP(1)
			Dim yRearArmorValue As Byte = yArmorHP(2)
			Dim yRightArmorValue As Byte = yArmorHP(3)

			If yForeArmorValue = 255 Then yForeArmorValue = 0
			If yLeftArmorValue = 255 Then yLeftArmorValue = 0
			If yRearArmorValue = 255 Then yRearArmorValue = 0
			If yRightArmorValue = 255 Then yRightArmorValue = 0

			ReDim yData(lDataSize)

			lPos = 0
			oStream.Position = lPos
			lVal = oStream.Read(yData, 0, yData.Length)

			For LocX = 0 To 31
				For LocY = 0 To 31
					lIdx = (LocY * 128) + (LocX * 4)		 '128 = 32 * 4

					If LocX < 4 OrElse LocX > 27 Then
						yData(lIdx) = 0
						yData(lIdx + 1) = 0
						yData(lIdx + 2) = 0
						yData(lIdx + 3) = 0
						'make black borders
					ElseIf LocX = 4 OrElse LocX = 27 OrElse LocY = 0 OrElse LocY = 31 OrElse LocY = 26 OrElse LocY = 21 OrElse LocY = 5 OrElse (LocX = 9 AndAlso (LocY > 5 AndAlso LocY < 21)) OrElse (LocX = 22 AndAlso (LocY > 5 AndAlso LocY < 21)) Then
						yData(lIdx) = 0
						yData(lIdx + 1) = 0
						yData(lIdx + 2) = 0
						yData(lIdx + 3) = 255
					ElseIf rcShield.Contains(LocX, LocY) Then
						'I only care about X
						lD = LocX - rcShield.Left

						If lD = 0 AndAlso yShieldHP <> 0 Then
							yData(lIdx) = 255
							yData(lIdx + 1) = 128
							yData(lIdx + 2) = 0
							yData(lIdx + 3) = 255
						ElseIf (lD / rcShield.Width) * 100 > yShieldHP Then
							yData(lIdx) = 0
							yData(lIdx + 1) = 0
							yData(lIdx + 2) = 255
							yData(lIdx + 3) = 255
						Else
							yData(lIdx) = 255
							yData(lIdx + 1) = 128
							yData(lIdx + 2) = 0
							yData(lIdx + 3) = 255
						End If

					ElseIf rcTopArmor.Contains(LocX, LocY) Then
						'i only care about x
						lD = LocX - rcTopArmor.Left

						If lD = 0 AndAlso yForeArmorValue <> 0 Then
							yData(lIdx) = 0
							yData(lIdx + 1) = 255
							yData(lIdx + 2) = 0
							yData(lIdx + 3) = 255
						ElseIf (lD / rcTopArmor.Width) * 100 > yForeArmorValue Then
							yData(lIdx) = 0
							yData(lIdx + 1) = 0
							yData(lIdx + 2) = 255
							yData(lIdx + 3) = 255
						Else
							yData(lIdx) = 0
							yData(lIdx + 1) = 255
							yData(lIdx + 2) = 0
							yData(lIdx + 3) = 255
						End If
					ElseIf rcLeftArmor.Contains(LocX, LocY) Then
						'I only care about y
						lD = LocY - rcLeftArmor.Top

						If lD = 0 AndAlso yLeftArmorValue <> 0 Then
							yData(lIdx) = 0
							yData(lIdx + 1) = 255
							yData(lIdx + 2) = 0
							yData(lIdx + 3) = 255
						ElseIf (lD / rcLeftArmor.Height) * 100 > yLeftArmorValue Then
							yData(lIdx) = 0
							yData(lIdx + 1) = 0
							yData(lIdx + 2) = 255
							yData(lIdx + 3) = 255
						Else
							yData(lIdx) = 0
							yData(lIdx + 1) = 255
							yData(lIdx + 2) = 0
							yData(lIdx + 3) = 255
						End If
					ElseIf rcBottomArmor.Contains(LocX, LocY) Then
						'i only care about x
						lD = LocX - rcBottomArmor.Left

						If lD = 0 AndAlso yRearArmorValue <> 0 Then
							yData(lIdx) = 0
							yData(lIdx + 1) = 255
							yData(lIdx + 2) = 0
							yData(lIdx + 3) = 255
						ElseIf (lD / rcBottomArmor.Width) * 100 > yRearArmorValue Then
							yData(lIdx) = 0
							yData(lIdx + 1) = 0
							yData(lIdx + 2) = 255
							yData(lIdx + 3) = 255
						Else
							yData(lIdx) = 0
							yData(lIdx + 1) = 255
							yData(lIdx + 2) = 0
							yData(lIdx + 3) = 255
						End If
					ElseIf rcRightArmor.Contains(LocX, LocY) Then
						'i only care about y
						lD = LocY - rcRightArmor.Top

						If lD = 0 AndAlso yRightArmorValue <> 0 Then
							yData(lIdx) = 0
							yData(lIdx + 1) = 255
							yData(lIdx + 2) = 0
							yData(lIdx + 3) = 255
						ElseIf (lD / rcRightArmor.Height) * 100 > yRightArmorValue Then
							yData(lIdx) = 0
							yData(lIdx + 1) = 0
							yData(lIdx + 2) = 255
							yData(lIdx + 3) = 255
						Else
							yData(lIdx) = 0
							yData(lIdx + 1) = 255
							yData(lIdx + 2) = 0
							yData(lIdx + 3) = 255
						End If
					ElseIf rcStructure.Contains(LocX, LocY) Then
						'i only care about y
						lD = LocY - rcStructure.Top

						If lD = 0 AndAlso yStructureHP <> 0 Then
							yData(lIdx) = 0
							yData(lIdx + 1) = 255
							yData(lIdx + 2) = 0
							yData(lIdx + 3) = 255
						ElseIf (lD / rcStructure.Height) * 100 > yStructureHP Then
							yData(lIdx) = 0
							yData(lIdx + 1) = 0
							yData(lIdx + 2) = 255
							yData(lIdx + 3) = 255
						Else
							yData(lIdx) = 0
							yData(lIdx + 1) = 255
							yData(lIdx + 2) = 0
							yData(lIdx + 3) = 255
						End If
					Else
						yData(lIdx) = 0
						yData(lIdx + 1) = 0
						yData(lIdx + 2) = 0
						yData(lIdx + 3) = 0
					End If

				Next LocY
			Next LocX

			oStream.Position = lPos
			oStream.Write(yData, 0, yData.Length)
			HPTexture.UnlockRectangle(0)

			'now, check for whether the unit needs to burn
			TestForBurnFX()

			bHPChanged = False
		End If



	End Sub

	Public Sub TestForBurnFX()
		If goPFXEngine32 Is Nothing Then Return

		If muSettings.BurnFXParticles = 0 Then Return

		Dim lCnt As Int32
		Dim lIdx As Int32

		Dim lType As Int32
		Dim lPCntAdjust As Int32 = 0

		lType = BurnFX.ParticleEngine.EmitterType.eFireEmitter
		If goCurrentEnvir Is Nothing = False Then
			If goCurrentEnvir.ObjTypeID = ObjectType.ePlanet Then
				lType = BurnFX.ParticleEngine.EmitterType.eSmokeyFireEmitter
				lPCntAdjust = 10
			End If
		End If

		For lSide As Int32 = 0 To 3
			If yArmorHP(lSide) < 100 OrElse yStructureHP < 50 Then
				If yArmorHP(lSide) = 0 OrElse yStructureHP < 50 Then
					lIdx = Me.GetBurnFXLoc(lSide)
					While lIdx <> -1
						With Me.oMesh.muBurnLocs(lIdx)
							Me.mlBurnFXActive(lIdx) = goPFXEngine32.AddEmitter(CType(lType Or .ModValue, BurnFX.ParticleEngine.EmitterType), New Microsoft.DirectX.Vector3(.lOffsetX, .lOffsetY, .lOffsetZ), .lPCnt + lPCntAdjust, Me)
						End With
						lIdx = Me.GetBurnFXLoc(lSide)
					End While
				Else
					lCnt = 100 \ yArmorHP(lSide)
					If Me.mlBurnFXActive Is Nothing = False Then
						For X As Int32 = 0 To Me.mlBurnFXActive.GetUpperBound(0)
							If Me.mlBurnFXActive(X) <> -1 Then
								If Me.oMesh.muBurnLocs(X).lSide = lSide OrElse Me.oMesh.muBurnLocs(X).lSide = -1 Then
									lCnt -= 1
								End If
							End If
						Next X
					End If

					For X As Int32 = 0 To lCnt - 1
						lIdx = Me.GetBurnFXLoc(lSide)
						If lIdx = -1 Then Exit For
						With Me.oMesh.muBurnLocs(lIdx)
							Me.mlBurnFXActive(lIdx) = goPFXEngine32.AddEmitter(CType(lType Or .ModValue, BurnFX.ParticleEngine.EmitterType), New Microsoft.DirectX.Vector3(.lOffsetX, .lOffsetY, .lOffsetZ), .lPCnt + lPCntAdjust, Me)
						End With
					Next X
				End If
			Else
				'check for items on this side
				If Me.mlBurnFXActive Is Nothing = False Then
					For X As Int32 = 0 To Me.mlBurnFXActive.GetUpperBound(0)
						If Me.mlBurnFXActive(X) <> -1 Then
							If Me.oMesh.muBurnLocs(X).lSide = lSide OrElse Me.oMesh.muBurnLocs(X).lSide = -1 Then
								goPFXEngine32.StopEmitter(Me.mlBurnFXActive(X))
								Me.mlBurnFXActive(X) = -1
							End If
						End If
					Next X
				End If
			End If
		Next lSide

	End Sub

	Public Sub New()
		yArmorHP(0) = 100
		yArmorHP(1) = 100
		yArmorHP(2) = 100
		yArmorHP(3) = 100
		lTargetIdx(0) = -1 : lTargetIdx(1) = -1 : lTargetIdx(2) = -1 : lTargetIdx(3) = -1
	End Sub

	Public Sub SetDest(ByVal lX As Int32, ByVal lZ As Int32, ByVal iAngle As Int16)
        'then, we can set our dest up

        'goUILib.AddNotification("SetDest: " & Me.ObjectID & ", " & lX & ", " & lZ, Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
		DestX = lX
		DestZ = lZ
		DestAngle = iAngle

		yChangeEnvironment = 0

		If (CurrentStatus And elUnitStatus.eUnitMoving) = 0 Then
			LastUpdateCycle = glCurrentCycle
			'finally, set our movement status
			CurrentStatus = CurrentStatus Or elUnitStatus.eUnitMoving
		End If

		'By rule, a unit that is building something cannot be moving. If that was to change, we should change this
		Me.bProducing = False

		InitializeEngineFX()
	End Sub

    Public Sub InitializeEngineFX()
        Dim X As Int32
        If goPFXEngine32 Is Nothing Then Exit Sub

        If mbEngineFXReady = False Then
            Try
                If MyBase.oMesh Is Nothing = False Then
                    If MyBase.oMesh.EngineFXCnt > 0 Then
                        ReDim lEngineFXEmitterID(MyBase.oMesh.EngineFXCnt - 1)
                        For X = 0 To MyBase.oMesh.EngineFXCnt - 1
                            lEngineFXEmitterID(X) = goPFXEngine32.AddEmitter(BurnFX.ParticleEngine.EmitterType.eEngineEmitter, _
                              MyBase.oMesh.EngineFXOffset(X), MyBase.oMesh.EngineFXPCnt(X), Me)
                        Next X
                    End If
                End If
            Catch
            End Try
            mbEngineFXReady = True
        End If
    End Sub

    Public Sub ShutoffEngines()
        'tells us to stop the engine particle fx
        If goPFXEngine32 Is Nothing Then Exit Sub
        Dim X As Int32

        If mbEngineFXReady = True AndAlso lEngineFXEmitterID Is Nothing = False Then
            For X = 0 To lEngineFXEmitterID.GetUpperBound(0)
                Try
                    goPFXEngine32.StopEmitter(lEngineFXEmitterID(X))
                Catch
                    'do nothing
                End Try
            Next X
            ReDim lEngineFXEmitterID(-1)
            mbEngineFXReady = False
        End If
    End Sub

    Public Sub ClearParticleFX()
        Dim X As Int32

        If goPFXEngine32 Is Nothing = False Then
			If mbEngineFXReady = True Then
				If lEngineFXEmitterID Is Nothing = False Then
					For X = 0 To lEngineFXEmitterID.Length - 1
						goPFXEngine32.StopEmitter(lEngineFXEmitterID(X))
					Next X
				End If
			End If

			If Me.mlBurnFXActive Is Nothing = False Then
				For X = 0 To Me.mlBurnFXActive.GetUpperBound(0)
					goPFXEngine32.StopEmitter(Me.mlBurnFXActive(X))
				Next X
			End If
		End If

    End Sub

    Public Sub SetHPsFromBurstMsg(ByVal yValue As Byte)
        If yValue = 0 Then Return

        bHPChanged = True

        'Ok, here's how it works...
        'Bit 1, 2, 3, 4 indicate that the side is burning
        'Bit 5, 6, 7, 8 indicate that the side should be full burning

        'however, if yValue = 255, we can cut to the chase
        If yValue = 255 Then
            For X As Int32 = 0 To 3
                yArmorHP(X) = 0
            Next
            yStructureHP = 30
        Else
            'Do individual checks
            If (yValue And 16) <> 0 Then
                'ok, side 0 is full burning
                yArmorHP(0) = 0
            ElseIf (yValue And 1) <> 0 Then
                'ok, side 0 is just burning
                yArmorHP(0) = 75
            End If

            If (yValue And 32) <> 0 Then
                'side 1 is full burning
                yArmorHP(1) = 0
            ElseIf (yValue And 2) <> 0 Then
                'normal burn
                yArmorHP(1) = 75
            End If

            If (yValue And 64) <> 0 Then
                'side 2 is full burning
                yArmorHP(2) = 0
            ElseIf (yValue And 3) <> 0 Then
                yArmorHP(2) = 75
            End If

            If (yValue And 128) <> 0 Then
                yArmorHP(3) = 0
            ElseIf (yValue And 4) <> 0 Then
                yArmorHP(3) = 75
            End If
        End If

        TestForBurnFX()
	End Sub

    Public Shared Function GetExperienceLevelName(ByVal lExpLevel As Int32) As String
        Dim lXPRank As Int32 = Math.Abs((CInt(lExpLevel) - 1) \ 25)
        Dim sTemp As String
        Select Case lXPRank
            Case 0  'Green
                sTemp = "Green"
            Case 1  'Trained
                sTemp = "Trained"
            Case 2  'Experienced
                sTemp = "Experienced"
            Case 3  'Adept
                sTemp = "Adept"
            Case 4  'Veteran
                sTemp = "Veteran"
            Case 5  'Ace
                sTemp = "Ace"
            Case 6  'Top Ace
                sTemp = "Top Ace"
            Case 7  'Distinguished
                sTemp = "Distinguished"
            Case 8  'Revered
                sTemp = "Revered"
            Case Else   'Elite
                sTemp = "Elite"
        End Select
        Return sTemp
    End Function
	Public Shared Function GetExperienceLevelNameAndBenefits(ByVal lExpLevel As Int32) As String
		Dim lXPRank As Int32 = Math.Abs((CInt(lExpLevel) - 1) \ 25)	'  CInt(Math.Floor(Math.Abs(.Exp_Level - 1) / 25))
		Dim sTemp As String = "Unknown"
		Select Case lXPRank
			Case 0	'Green
				sTemp = "Experience: Green"
			Case 1	'Trained
				sTemp = "Experience: Trained" & vbCrLf & "+5 To-Hit Bonus"
			Case 2	'Experienced
				sTemp = "Experience: Experienced" & vbCrLf & "+8 To-Hit Bonus"
			Case 3	'Adept
				sTemp = "Experience: Adept" & vbCrLf & "+8 To-Hit Bonus" & vbCrLf & "+1 Speed"
			Case 4	'Veteran
				sTemp = "Experience: Veteran" & vbCrLf & "+8 To-Hit Bonus" & vbCrLf & "+1 Speed" & vbCrLf & "+1 Maneuver"
			Case 5	'Ace
				sTemp = "Experience: Ace" & vbCrLf & "+10 To-Hit Bonus" & vbCrLf & "+1 Speed" & vbCrLf & "+1 Maneuver"
			Case 6	'Top Ace
				sTemp = "Experience: Top Ace" & vbCrLf & "+10 To-Hit Bonus" & vbCrLf & "+1 Speed" & vbCrLf & "+1 Maneuver" & vbCrLf & "+5% Damage"
			Case 7	'Distinguished
				sTemp = "Experience: Distinguished" & vbCrLf & "+10 To-Hit Bonus" & vbCrLf & "+3 Speed" & vbCrLf & "+1 Maneuver" & vbCrLf & "+5% Damage"
			Case 8	'Revered
				sTemp = "Experience: Revered" & vbCrLf & "+10 To-Hit Bonus" & vbCrLf & "+4 Speed" & vbCrLf & "+2 Maneuver" & vbCrLf & "+5% Damage"
			Case Else	'Elite
				sTemp = "Experience: Elite" & vbCrLf & "+13 To-Hit Bonus" & vbCrLf & "+5 Speed" & vbCrLf & "+3 Maneuver" & vbCrLf & "+10% Damage"
        End Select

        'Now, an experience meter...
        If lXPRank < 9 Then
            Dim lXPReqForNextLevel As Int32 = lXPRank + 1
            lXPReqForNextLevel *= 25
            lXPReqForNextLevel += 1

            lXPReqForNextLevel -= lExpLevel
            sTemp &= vbCrLf & "Experience to Next Level: " & lXPReqForNextLevel.ToString
        End If

		Return sTemp
	End Function
End Class

'Like a watered down version of BaseEntity
Public Class StationChild
    Public lChildID As Int32 = -1
    Public iChildTypeID As Int16 = -1
    Public oChildDef As EntityDef = Nothing
    Public bProducing As Boolean = False
    Public lFinishCycle As Int32
    Public lProdStarted As Int32
    Public lChildStatus As Int32

    'What is currently being produced
    Public lProdID As Int32
    Public iProdTypeID As Int16

    Private mfLastPerc As Single

    Public Function GetProductionStatus() As Single
        'returns a number from 0 to 1 or -1
        If bProducing = False Then Return -1

		Dim fResult As Single = 0.0F
		Try
			Dim lTotalTime As Int32 = lFinishCycle - lProdStarted
			If lTotalTime = 0 Then Return 0.0F
			fResult = CSng((glCurrentCycle - lProdStarted) / lTotalTime)
		Catch
			fResult = 0.0F
		End Try
		Return fResult
    End Function

    Public Sub ReceiveProductionStatusMsg(ByVal lServerCurrent As Int32, ByVal lServerFinish As Int32, ByVal fPerc As Single)
        If fPerc = 255.0F Then
            bProducing = False
            lProdID = -1
        Else
            bProducing = True

            If fPerc <> mfLastPerc Then

                'Now... need to do some conversion...
                Dim lCurrentDiff As Int32 = glCurrentCycle - lServerCurrent
                Dim lServerCyclesRemaining As Int32 = lServerFinish - lServerCurrent
                Dim fDivisor As Single = (100.0F - CSng(fPerc)) / 100.0F

                'Ok, adjust the server's finish cycle to my cycle
                lFinishCycle = lServerFinish + lCurrentDiff

                'Now, we can calculate our started...
                If fDivisor < 0.0001 Then lProdStarted = 0 Else lProdStarted = CInt(lFinishCycle - (lServerCyclesRemaining / fDivisor))

                mfLastPerc = fPerc
            End If
        End If
    End Sub
End Class