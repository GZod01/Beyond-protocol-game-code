Option Strict On

Public Class PlayerItemIntel
    Public Enum PlayerItemIntelType As Byte
        eExistance = 0      'player knows of it
        eLocation = 1       'player knows where it is located
        eStatus = 2         'player knows the item's status
        eFullKnowledge = 4  'player knows all things
    End Enum
    Public lItemID As Int32                     'ID of the item
    Public iItemTypeID As Int16                 'TypeID of the item
    Public lItemIdx As Int32                    'index in the global array for this item (some items may not use this)
    Public yIntelType As PlayerItemIntelType    'Intel Type indicating knowledge known
    Public lOtherPlayerID As Int32              'other player ID who this intel is about
    Public lPlayerID As Int32                   'player who knows this intel
    Public dtTimeStamp As DateTime              'real life date/time when knowledge was acquired
	Public lValue As Int32						'arbitrary value at time of timestamp

	Public LocX As Int32 = Int32.MinValue		'Location of this item when the intel was discovered
	Public LocZ As Int32 = Int32.MinValue		'Location of this item when the intel was discovered
	Public EnvirID As Int32 = -1				'Environment where this item was located when this intel was discovered
	Public EnvirTypeID As Int16 = -1			'Environment type 
	Public StatusValue As Int32 = 0				'status value (if the intel type is status or better)
	Public FullKnowledge() As Byte = Nothing	'string text of the intel report (if the intel type is fullknowledge or better)

	Public bArchived As Boolean = False

	Private mbRequestedDetails As Boolean = False
	Public Sub RequestDetails()
		If mbRequestedDetails = False Then
			mbRequestedDetails = True

			If yIntelType > PlayerItemIntelType.eExistance Then
				Dim yMsg(11) As Byte
				Dim lPos As Int32 = 0
				System.BitConverter.GetBytes(GlobalMessageCode.eGetItemIntelDetail).CopyTo(yMsg, lPos) : lPos += 2
				System.BitConverter.GetBytes(lItemID).CopyTo(yMsg, lPos) : lPos += 4
				System.BitConverter.GetBytes(iItemTypeID).CopyTo(yMsg, lPos) : lPos += 2
				System.BitConverter.GetBytes(lOtherPlayerID).CopyTo(yMsg, lPos) : lPos += 4
				If goUILib Is Nothing = False Then goUILib.SendMsgToPrimary(yMsg)
			End If
		End If
	End Sub

	Public Sub FillDetails(ByRef yData() As Byte, ByVal lPos As Int32)
		If yIntelType >= PlayerItemIntelType.eLocation Then
			LocX = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			LocZ = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			EnvirID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            EnvirTypeID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

            'Update the Intel Form if it's open (so the goto button lights up)
            Dim ofrm As frmIntelReport = CType(goUILib.GetWindow("frmIntelReport"), frmIntelReport)
            If Not ofrm Is Nothing Then
                ofrm.SetFromItemIntel(Me)
            End If
            ofrm = Nothing
        End If
		If yIntelType >= PlayerItemIntelType.eStatus Then
			StatusValue = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		End If
		If yIntelType >= PlayerItemIntelType.eFullKnowledge Then
			Dim lSize As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			ReDim FullKnowledge(lSize - 1)
			Array.Copy(yData, lPos, FullKnowledge, 0, lSize)
			lPos += lSize
		End If
	End Sub

	Public Sub FillFromMsg(ByRef yData() As Byte)
		mbRequestedDetails = False
		Dim lPos As Int32 = 2	'for msgcode?
		lItemID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		lPos += 2	'for the object type of PlayerItemIntel
		iItemTypeID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
		lPlayerID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		lOtherPlayerID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		yIntelType = CType(yData(lPos), PlayerItemIntelType) : lPos += 1

		Dim lSeconds As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		dtTimeStamp = Now.Subtract(New TimeSpan(0, 0, -(Math.Abs(lSeconds))))

		lValue = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
	End Sub

	Public Shared Function GetGroupValue(ByVal iType As Int16) As Int32
		'  Agents
		'  Colony
		'  Facility
		'  Mineral
		'  PlayerRel
		'  Trade
		'  Unit
		Select Case iType
			Case ObjectType.eAgent
				Return 1
			Case ObjectType.eColony
				Return 2
			Case ObjectType.eFacility
				Return 3
			Case ObjectType.eMineral
				Return 4
			Case ObjectType.ePlayer	'playerrel
				Return 5
			Case ObjectType.eTrade
				Return 6
			Case ObjectType.eUnit
				Return 7
			Case Else
				Return 200
		End Select
	End Function
	Public Shared Function GetGroupValueName(ByVal iType As Int16) As String
		Select Case iType
			Case ObjectType.eAgent
				Return "AGENT"
			Case ObjectType.eColony
				Return "COLONY"
			Case ObjectType.eFacility
				Return "FACILITY"
			Case ObjectType.eMineral
				Return "MINERAL"
			Case ObjectType.ePlayer	'playerrel
				Return "PLAYER RELATIONS"
			Case ObjectType.eTrade
				Return "TRADES"
			Case ObjectType.eUnit
				Return "UNITS"
			Case Else
				Return "OTHER"
		End Select
	End Function

	Public Function GetIntelReport() As String
		Dim oSB As New System.Text.StringBuilder

		Dim sVal As String = GetGroupValueName(iItemTypeID).ToUpper
		sVal = sVal.Substring(0, 1).ToUpper & sVal.Substring(1).ToLower

		oSB.AppendLine(sVal & " report" & vbCrLf)
		oSB.AppendLine("Owner: " & GetCacheObjectValue(lOtherPlayerID, ObjectType.ePlayer))
        oSB.AppendLine("Age of Report: " & GetDurationFromSeconds(Math.Abs(CInt(Now.Subtract(dtTimeStamp).TotalSeconds)), True) & vbCrLf)

		If Me.yIntelType = PlayerItemIntelType.eExistance Then
			oSB.AppendLine("All we know is that it exists")
		End If
		If Me.yIntelType >= PlayerItemIntelType.eLocation Then
			If Me.iItemTypeID = ObjectType.eAgent Then
				oSB.AppendLine("We know that the agent has infiltrated " & GetCacheObjectValue(EnvirID, ObjectType.ePlayer) & ".")
			ElseIf Me.iItemTypeID = ObjectType.eColony OrElse Me.iItemTypeID = ObjectType.eFacility OrElse Me.iItemTypeID = ObjectType.eUnit Then
				oSB.AppendLine("We know that it is located in " & GetCacheObjectValue(EnvirID, EnvirTypeID) & ".")
			End If
		End If
		If Me.yIntelType >= PlayerItemIntelType.eStatus Then
			Select Case Me.iItemTypeID
				Case ObjectType.eAgent
                    oSB.AppendLine("Last Known Status: " & Agent.GetStatusText(StatusValue, EnvirID, EnvirTypeID, 255)) ' 255 means no infiltration known
				Case ObjectType.eColony
					oSB.AppendLine("Last Reported Population: " & StatusValue.ToString("#,##0"))
				Case ObjectType.eFacility
					If StatusValue = 0 Then oSB.AppendLine("The facility is unpowered") Else oSB.AppendLine("The facility is powered")
				Case ObjectType.ePlayer
					oSB.AppendLine("Last Known Relationship Standing: " & StatusValue.ToString("#,##0"))
				Case ObjectType.eUnit
					oSB.AppendLine("Experience Level: " & BaseEntity.GetExperienceLevelName(StatusValue))
			End Select
		End If
		If Me.yIntelType >= PlayerItemIntelType.eFullKnowledge Then
			'right now, there is only colony and trade, eventually, we could add agent, facility and unit
			Select Case Me.iItemTypeID
				Case ObjectType.eColony, ObjectType.eTrade
					oSB.AppendLine(System.Text.ASCIIEncoding.ASCII.GetString(FullKnowledge))
			End Select
		End If

		Return oSB.ToString
	End Function
End Class
