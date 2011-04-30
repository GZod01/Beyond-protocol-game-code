Public Class Mineral
    'Actual Minerals/alloys and actual values
    Inherits Epica_GUID
    Public ServerIndex As Int32 = -1    'index in the global array
    Public lAlloyTechID As Int32 = -1   'if -1, mineral is natural mineral, otherwise, it is alloy tech result
    Public lRarity As Int32

    Public MineralName(19) As Byte
    Private moMineralPropertyValues() As MineralPropertyValue
    Private mlMineralPropertyValueUB As Int32 = -1

    Public MineralValue As Int32 = 83

    Private moMinGeoRel() As MineralGeographyRel = Nothing
    Private mlMinGeoRelUB As Int32 = -1

	Public Function GetPropertyValue(ByVal lPropertyID As Int32) As Int32
		For X As Int32 = 0 To mlMineralPropertyValueUB
			If moMineralPropertyValues(X).lPropertyID = lPropertyID Then
				Return moMineralPropertyValues(X).lValue
			End If
		Next
		Return 0
	End Function

    Public Sub SetPropertyValue(ByVal lMinPropValID As Int32, ByVal lPropID As Int32, ByVal lValue As Int32)
        For X As Int32 = 0 To mlMineralPropertyValueUB
            If moMineralPropertyValues(X).lPropertyID = lPropID Then
                moMineralPropertyValues(X).lValue = lValue
                Return
            End If
        Next

        'Ok, didn't find it
        mlMineralPropertyValueUB += 1
        ReDim Preserve moMineralPropertyValues(mlMineralPropertyValueUB)
        moMineralPropertyValues(mlMineralPropertyValueUB) = New MineralPropertyValue
        With moMineralPropertyValues(mlMineralPropertyValueUB)
            .MineralPropertyValueID = lMinPropValID
            .lPropertyID = lPropID
            .lParentID = Me.ObjectID
            .lValue = lValue
            .oProperty = GetEpicaMineralProperty(lPropID)
        End With
    End Sub

    'NOTE: To send a mineral as an object, you must call the appropriate player's PlayerMineral.GetObjAsString

    Public Function SaveObject() As Boolean
        Dim bResult As Boolean = False
        Dim sSQL As String
        Dim oComm As OleDb.OleDbCommand

        'if the object's OBJECT ID is -1, then we know that this is the first time saving this object!

        Try
            If ObjectID = -1 Then
                'INSERT
                sSQL = "INSERT INTO tblMineral (MineralName, AlloyTechID, Rarity) VALUES ('" & _
                    MakeDBStr(BytesToString(MineralName)) & "', " & lAlloyTechID & ", " & lRarity & ")"
            Else
                'UPDATE
                sSQL = "UPDATE tblMineral SET MineralName ='" & MakeDBStr(BytesToString(MineralName)) & _
                    "', AlloyTechID = " & lAlloyTechID & ", Rarity = " & lRarity & " WHERE MineralID = " & ObjectID
            End If
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            If oComm.ExecuteNonQuery() = 0 Then
                Err.Raise(-1, "SaveObject", "No records affected when saving object!")
            End If
            If ObjectID = -1 Then
                Dim oData As OleDb.OleDbDataReader
                oComm = Nothing
                sSQL = "SELECT MAX(MineralID) FROM tblMineral WHERE MineralName = '" & MakeDBStr(BytesToString(MineralName)) & "'"
                oComm = New OleDb.OleDbCommand(sSQL, goCN)
                oData = oComm.ExecuteReader(CommandBehavior.Default)
                If oData.Read Then
                    ObjectID = CInt(oData(0))
                End If
                oData.Close()
                oData = Nothing
            End If

            'Now, save our values...
            For X As Int32 = 0 To mlMineralPropertyValueUB
                moMineralPropertyValues(X).lParentID = Me.ObjectID
                moMineralPropertyValues(X).SaveObject(0)
            Next X

            bResult = True
        Catch
            LogEvent(GlobalVars.LogEventType.CriticalError, "Unable to save object " & ObjectID & " of type " & ObjTypeID & ". Reason: " & Err.Description)
        Finally
            oComm = Nothing
        End Try
        Return bResult
    End Function

    Public Sub AddMinGeoRel(ByVal lGeoID As Int32, ByVal iGeoTypeID As Int16, ByVal lMaxQty As Int32, ByVal lMaxConc As Int32)
        For X As Int32 = 0 To mlMinGeoRelUB
            If moMinGeoRel(X).lGeoID = lGeoID AndAlso moMinGeoRel(X).iGeoTypeID = iGeoTypeID Then
                moMinGeoRel(X).lRedistMaxConc = lMaxConc
                moMinGeoRel(X).lRedistMaxQty = lMaxQty
                Return
            End If
        Next X

        mlMinGeoRelUB += 1
        ReDim Preserve moMinGeoRel(mlMinGeoRelUB)
        moMinGeoRel(mlMinGeoRelUB) = New MineralGeographyRel()
        With moMinGeoRel(mlMinGeoRelUB)
            .iGeoTypeID = iGeoTypeID
            .lGeoID = lGeoID
            .lMineralID = Me.ObjectID
            .lRedistMaxConc = lMaxConc
            .lRedistMaxQty = lMaxQty
        End With
    End Sub

    Private mlMineralScore As Int32 = Int32.MinValue
    Public ReadOnly Property GetMineralScore() As Int32
        Get
            If mlMineralScore = Int32.MinValue Then
                mlMineralScore = 0
                For X As Int32 = 0 To mlMineralPropertyValueUB
                    mlMineralScore += moMineralPropertyValues(X).lValue
                Next X
            End If
            Return mlMineralScore
        End Get
    End Property

    Private myNonOwnerMsg() As Byte = Nothing
    Public Function GetNonOwnerMsg() As Byte()
		If myNonOwnerMsg Is Nothing Then

			Dim lCnt As Int32 = 0
			For X As Int32 = 0 To mlMineralPropertyValueUB
				If moMineralPropertyValues(X).lPropertyID <> eMinPropID.MineralColor AndAlso moMineralPropertyValues(X).lPropertyID <> eMinPropID.Complexity Then
					lCnt += 1
				End If
			Next X

			ReDim myNonOwnerMsg(1 + (lCnt * 5))
			Dim lPos As Int32 = 0
			System.BitConverter.GetBytes(CShort(lCnt)).CopyTo(myNonOwnerMsg, lPos) : lPos += 2
			For X As Int32 = 0 To mlMineralPropertyValueUB
				If moMineralPropertyValues(X).lPropertyID <> eMinPropID.MineralColor AndAlso moMineralPropertyValues(X).lPropertyID <> eMinPropID.Complexity Then
					System.BitConverter.GetBytes(moMineralPropertyValues(X).lPropertyID).CopyTo(myNonOwnerMsg, lPos) : lPos += 4
					myNonOwnerMsg(lPos) = PlayerMineral.GetClientDisplayedPropertyValue(moMineralPropertyValues(X).lValue) : lPos += 1
					'Dim fTemp As Single = CSng(moMineralPropertyValues(X).lValue / moMineralPropertyValues(X).oProperty.MaximumValue)
					'myNonOwnerMsg(lPos) = CByte(Math.Floor(fTemp * 10)) : lPos += 1
				End If
			Next X
		End If
        Return myNonOwnerMsg
	End Function

	Public Sub RespawnMe(ByVal lOriginalEnvirID As Int32, ByVal iOriginalEnvirTypeID As Int16)
		Dim yMsg(13) As Byte
		Dim lPos As Int32 = 0
		System.BitConverter.GetBytes(GlobalMessageCode.eAddObjectCommand).CopyTo(yMsg, lPos) : lPos += 2
		Me.GetGUIDAsString.CopyTo(yMsg, lPos) : lPos += 6
		System.BitConverter.GetBytes(lOriginalEnvirID).CopyTo(yMsg, lPos) : lPos += 4
		System.BitConverter.GetBytes(iOriginalEnvirTypeID).CopyTo(yMsg, lPos) : lPos += 2
		goMsgSys.SendMsgToOperator(yMsg)
    End Sub

    Public Function GetForcePrimarySyncMsg() As Byte()
        Dim lCnt As Int32 = mlMineralPropertyValueUB + 1

        Dim yMsg(39 + (lCnt * 12)) As Byte
        Dim lPos As Int32 = 0
        System.BitConverter.GetBytes(GlobalMessageCode.eForcePrimarySync).CopyTo(yMsg, lPos) : lPos += 2
        Me.GetGUIDAsString.CopyTo(yMsg, lPos) : lPos += 6
        System.BitConverter.GetBytes(lAlloyTechID).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(lRarity).CopyTo(yMsg, lPos) : lPos += 4
        MineralName.CopyTo(yMsg, lPos) : lPos += 20
        System.BitConverter.GetBytes(lCnt).CopyTo(yMsg, lPos) : lPos += 4
        For X As Int32 = 0 To mlMineralPropertyValueUB
            System.BitConverter.GetBytes(moMineralPropertyValues(X).lPropertyID).CopyTo(yMsg, lPos) : lPos += 4
            System.BitConverter.GetBytes(moMineralPropertyValues(X).lValue).CopyTo(yMsg, lPos) : lPos += 4
            System.BitConverter.GetBytes(moMineralPropertyValues(X).MineralPropertyValueID).CopyTo(yMsg, lPos) : lPos += 4
        Next X
        Return yMsg
    End Function

    Public Sub FillFromPrimarySyncMsg(ByVal yData() As Byte)
        Dim lPos As Int32 = 2    'for msgcode
        With Me
            .ObjectID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .ObjTypeID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
            .lAlloyTechID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .lRarity = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            ReDim .MineralName(19)
            Array.Copy(yData, lPos, .MineralName, 0, 20) : lPos += 20
            mlMineralPropertyValueUB = System.BitConverter.ToInt32(yData, lPos) - 1 : lPos += 4

            ReDim moMineralPropertyValues(mlMineralPropertyValueUB)
            For X As Int32 = 0 To mlMineralPropertyValueUB

                Dim lPropID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                Dim lPropVal As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                Dim lMPVID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

                moMineralPropertyValues(X) = New MineralPropertyValue
                With moMineralPropertyValues(X)
                    .lParentID = Me.ObjectID
                    .lPropertyID = lPropID
                    .lValue = lPropVal
                    .MineralPropertyValueID = lMPVID
                    .oProperty = GetEpicaMineralProperty(lPropID)
                End With
            Next X

        End With
    End Sub
End Class
