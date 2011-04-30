Public Enum eyRouteLoadItemType As Byte
    eNoLoadAllItems = 0
    eLoadAnyAllItems = 1
    eLoadAllMinerals = 2
    eLoadAllComponents = 3
    eLoadAllArmor = 4
    eLoadAllEngines = 5
    eLoadAllRadar = 6
    eLoadAllShields = 7
    eLoadAllWeapons = 8
    eLoadAbsolutelyNothing = 9
    eLoadColonists = 10
    eLoadEnlisted = 11
    eLoadOfficers = 12
End Enum
'This structure is used for UNITS only... it indicates a point in a route of points...
Public Structure RouteItem
    'a route item is defined as:
    '  GUID - indicates the item to move to... can be an environment or another entity (for example, a facility/unit)
    Public oDest As Epica_GUID
    '  Loc - indicates the location within GUID to move to... if the GUID is another entity, then this value is unused
    Public lLocX As Int32
    Public lLocZ As Int32
    '  LoadItemGUID - indicates the item to load (a mineral/alloy) when the unit arrives at the destination
    Public oLoadItem As Epica_GUID

    Public yLoadAllItems As Byte
    Public yExtraFlags As Byte

    Public lOrderNum As Int32

    Public Sub SetLoadItem(ByVal lObjID As Int32, ByVal iObjTypeID As Int16, ByVal oPlayer As Player, ByVal yFlag As Byte)
        If lObjID = -1 Then
            Select Case iObjTypeID
                Case ObjectType.eMineral
                    yLoadAllItems = eyRouteLoadItemType.eLoadAllMinerals
                Case ObjectType.eComponentCache
                    yLoadAllItems = eyRouteLoadItemType.eLoadAllComponents
                Case ObjectType.eArmorTech
                    yLoadAllItems = eyRouteLoadItemType.eLoadAllArmor
                Case ObjectType.eEngineTech
                    yLoadAllItems = eyRouteLoadItemType.eLoadAllEngines
                Case ObjectType.eRadarTech
                    yLoadAllItems = eyRouteLoadItemType.eLoadAllRadar
                Case ObjectType.eShieldTech
                    yLoadAllItems = eyRouteLoadItemType.eLoadAllShields
                Case ObjectType.eWeaponTech
                    yLoadAllItems = eyRouteLoadItemType.eLoadAllWeapons
                Case -2
                    yLoadAllItems = eyRouteLoadItemType.eLoadAbsolutelyNothing
                Case ObjectType.eColonists
                    yLoadAllItems = eyRouteLoadItemType.eLoadColonists
                Case ObjectType.eEnlisted
                    yLoadAllItems = eyRouteLoadItemType.eLoadEnlisted
                Case ObjectType.eOfficers
                    yLoadAllItems = eyRouteLoadItemType.eLoadOfficers
                Case Else
                    yLoadAllItems = eyRouteLoadItemType.eLoadAnyAllItems
            End Select
            oLoadItem = New Epica_GUID()
            With oLoadItem
                .ObjectID = lObjID
                .ObjTypeID = iObjTypeID
            End With
        Else
            yLoadAllItems = eyRouteLoadItemType.eNoLoadAllItems
            If iObjTypeID = ObjectType.eMineral Then
                oLoadItem = GetEpicaMineral(lObjID)
            Else
                oLoadItem = oPlayer.GetTech(lObjID, iObjTypeID)
            End If
        End If

        yExtraFlags = yFlag

    End Sub
End Structure

'Public Class ObjectOrder
'    Inherits Epica_GUID

'    Public oParentObject As Object      'this is the object to whom this order is issued
'    Public OrderID As Short             'the order to execute (this is something we know of, for example, move would be 1
'    Public oTargetObject As Object      'the object to which this order is to be executed at, for example, attack order, this is who to attack
'    Public OrderLocX As Int32           'overloaded basically... for example, Move would not use Target Object
'    Public OrderLocY As Int32           'overloaded basically... for example, Move would not use Target Object
'    Public OrderNum As Int32            'the value representing this order's place in the order queue... TODO: This is not filled or used yet

'    Private mySendString() As Byte

'    Public Function GetObjAsString() As Byte()
'        'here we will return the entire object as a string
'        If mbStringReady = False Then
'            ReDim mySendString(31) '0 to 31 = 32 bytes

'            GetGUIDAsString.CopyTo(mySendString, 0)
'            CType(oParentObject, Epica_GUID).GetGUIDAsString.CopyTo(mySendString, 6)
'            System.BitConverter.GetBytes(OrderID).CopyTo(mySendString, 12)
'            If oTargetObject Is Nothing = False Then
'                CType(oTargetObject, Epica_GUID).GetGUIDAsString.CopyTo(mySendString, 14)
'            Else
'                System.BitConverter.GetBytes(-1).CopyTo(mySendString, 14)
'                System.BitConverter.GetBytes(-1).CopyTo(mySendString, 18)
'            End If
'            System.BitConverter.GetBytes(OrderLocX).CopyTo(mySendString, 20)
'            System.BitConverter.GetBytes(OrderLocY).CopyTo(mySendString, 24)
'            System.BitConverter.GetBytes(OrderNum).CopyTo(mySendString, 28)
'            mbStringReady = True
'        End If
'        Return mySendString
'    End Function

'    Public Function SaveObject() As Boolean
'        Dim bResult As Boolean = False
'        Dim sSQL As String
'        Dim oComm As OleDb.OleDbCommand

'        'if the object's OBJECT ID is -1, then we know that this is the first time saving this object!

'        Try
'            If ObjectID = -1 Then
'                'INSERT
'                sSQL = "INSERT INTO tblObjectOrder (ObjectID, ObjectTypeID, OrderID, OrderTargetID, " & _
'                  "OrderTargetTypeID, OrderLocX, OrderLocY, Order_OrderNum) VALUES (" & CType(oParentObject, Epica_GUID).ObjectID & _
'                  ", " & CType(oParentObject, Epica_GUID).ObjTypeID & ", " & OrderID & ", "
'                If oTargetObject Is Nothing = False Then
'                    sSQL = sSQL & CType(oTargetObject, Epica_GUID).ObjectID & ", " & CType(oTargetObject, Epica_GUID).ObjTypeID
'                Else
'                    sSQL = sSQL & "-1, -1"
'                End If
'                sSQL = sSQL & ", " & OrderLocX & ", " & OrderLocY & ", " & OrderNum & ")"
'            Else
'                'UPDATE
'                sSQL = "UPDATE tblObjectOrder SET ObjectID=" & CType(oParentObject, Epica_GUID).ObjectID & ", ObjectTypeID=" & _
'                  CType(oParentObject, Epica_GUID).ObjTypeID & ", OrderID=" & OrderID & ", OrderTargetID = "
'                If oTargetObject Is Nothing = False Then
'                    sSQL = sSQL & CType(oTargetObject, Epica_GUID).ObjectID & ", OrderTargetTypeID = " & CType(oTargetObject, Epica_GUID).ObjTypeID
'                Else
'                    sSQL = sSQL & "-1, -1"
'                End If
'                sSQL = sSQL & ", OrderLocX = " & OrderLocX & ", OrderLocY = " & OrderLocY & ", Order_OrderNum = " & _
'                  OrderNum & " WHERE ObjectOrderID = " & ObjectID
'            End If
'            oComm = New OleDb.OleDbCommand(sSQL, goCN)
'            If oComm.ExecuteNonQuery() = 0 Then
'                Err.Raise(-1, "SaveObject", "No records affected when saving object!")
'            End If
'            If ObjectID = -1 Then
'                Dim oData As OleDb.OleDbDataReader
'                oComm = Nothing
'                sSQL = "SELECT MAX(ObjectOrderID) FROM tblObjectOrder WHERE ObjectID = " & CType(oParentObject, Epica_GUID).ObjectID & _
'                  "AND ObjectTypeID = " & CType(oParentObject, Epica_GUID).ObjTypeID
'                oComm = New OleDb.OleDbCommand(sSQL, goCN)
'                oData = oComm.ExecuteReader(CommandBehavior.Default)
'                If oData.Read Then
'                    ObjectID = CInt(oData(0))
'                End If
'                oData.Close()
'                oData = Nothing
'            End If
'            bResult = True
'        Catch
'            LogEvent(GlobalVars.LogEventType.CriticalError, "Unable to save object " & ObjectID & " of type " & ObjTypeID & ". Reason: " & Err.Description)
'        Finally
'            oComm = Nothing
'        End Try
'        Return bResult
'    End Function
'End Class