Public Class Planet
    Inherits Epica_GUID

    Public Const PLANET_OBJ_MSG_LEN As Int32 = 71
    Public Const PLANET_ROTATE_ANGLE_MSG_LOC As Int32 = 69

    Public PlanetName(19) As Byte
    Public PlanetTypeID As Byte         'based on the PlanetType enum in EpicaShared
    Public PlanetSizeID As Byte         'used for determining Map Size: 0-Tiny, 1-Small, 2-Medium, 3-Large, 4-Huge
    Public PlanetRadius As Int16        'used for displaying the planet sphere
    Public ParentSystem As SolarSystem
    Public LocX As Int32
    Public LocY As Int32
    Public LocZ As Int32
    Public Vegetation As Byte
    Public Atmosphere As Byte
    Public Hydrosphere As Byte
    Public Gravity As Byte
    Public SurfaceTemperature As Byte
    Public RotationDelay As Int16       'cycles between incrementing the rotation angle (not used except on client)

    Public InnerRingRadius As Int32
    Public OuterRingRadius As Int32
    Public RingDiffuse As Int32
    Public RingMineralID As Int32 = 0
    Public RingConcentration As Int32 = 0

    Public PlayerSpawns As Int32 = 0    'number of player spawns used
    Public SpawnLocked As Boolean = False

    Public AxisAngle As Int32

    Public OwnerID As Int32 = -1

    Private Shared msw_Rotate As Stopwatch      'MSC 12/18/2006

	Private mySendString() As Byte

    Public Overloads Sub DataChanged()    'call this sub when changing properties of this object
        mbStringReady = False
        ParentSystem.mbDetailsReady = False
    End Sub

    Public Function GetObjAsString() As Byte()
        'here we will return the object populated as a string
        If mbStringReady = False Then
            'guid-6
            'name-20
            'typeid-1
            'sizeid-1
            'Radius-2
            'parentid-4
            'locx-4
            'locy-4
            'locz-4
            'veg-1
            'atmos-1
            'hyrdo-1
            'grav-1
            'surftemp-1
            'RotDelay-2
            'InnerRingRadius-4
            'OuterRingRadius-4
            'RingDiffuse-4
            'AxisAngle-4
            'CurrentAngle-2
            ReDim mySendString(PLANET_OBJ_MSG_LEN - 1)
            GetGUIDAsString.CopyTo(mySendString, 0)
            PlanetName.CopyTo(mySendString, 6)
            mySendString(26) = PlanetTypeID
            mySendString(27) = PlanetSizeID
            System.BitConverter.GetBytes(PlanetRadius).CopyTo(mySendString, 28)
            System.BitConverter.GetBytes(ParentSystem.ObjectID).CopyTo(mySendString, 30)
            System.BitConverter.GetBytes(LocX).CopyTo(mySendString, 34)
            System.BitConverter.GetBytes(LocY).CopyTo(mySendString, 38)
            System.BitConverter.GetBytes(LocZ).CopyTo(mySendString, 42)
            mySendString(46) = Vegetation
            mySendString(47) = Atmosphere
            mySendString(48) = Hydrosphere
            mySendString(49) = Gravity
            mySendString(50) = SurfaceTemperature
            System.BitConverter.GetBytes(RotationDelay).CopyTo(mySendString, 51)
            System.BitConverter.GetBytes(InnerRingRadius).CopyTo(mySendString, 53)
            System.BitConverter.GetBytes(OuterRingRadius).CopyTo(mySendString, 57)
            System.BitConverter.GetBytes(RingDiffuse).CopyTo(mySendString, 61)
            System.BitConverter.GetBytes(AxisAngle).CopyTo(mySendString, 65)
            System.BitConverter.GetBytes(GetRotateAngle).CopyTo(mySendString, 69)
            mbStringReady = True
        End If

        Return mySendString
    End Function

    Public Function GetRotateAngle() As Int16
        If msw_Rotate Is Nothing Then
            msw_Rotate = Stopwatch.StartNew
        End If
        'Lastly, the CurrentAngle changes constantly, send it back as of right now
        Return CShort(Math.Floor(msw_Rotate.ElapsedMilliseconds / RotationDelay) Mod 3600)
    End Function

    Public Function SaveObject() As Boolean
        Dim bResult As Boolean = False
        Dim sSQL As String
        Dim oComm As OleDb.OleDbCommand

        'if the object's OBJECT ID is -1, then we know that this is the first time saving this object!

        Try
            If ObjectID = -1 Then
                'INSERT
                sSQL = "INSERT INTO tblPlanet (PlanetTypeID, PlanetName, PlanetSizeID, ParentID, LocX, LocY, LocZ, " & _
                  "Vegetation, Atmosphere, Hydrosphere, Gravity, SurfaceTemperature, RotationDelay, PlanetRadius, PlayerSpawns, " & _
                  "SpawnLocked, RingInnerRadius, RingOuterRadius, RingDiffuse, AxisAngle, RingMineralID, RingMineralConcentration) VALUES (" & _
                  PlanetTypeID & ", '" & MakeDBStr(BytesToString(PlanetName)) & "', " & PlanetSizeID & ", " & ParentSystem.ObjectID & ", " & _
                  LocX & ", " & LocY & ", " & LocZ & ", " & Vegetation & ", " & Atmosphere & ", " & Hydrosphere & ", " & _
                  Gravity & ", " & SurfaceTemperature & ", " & RotationDelay & ", " & PlanetRadius & ", " & PlayerSpawns & ", "
                If SpawnLocked = True Then sSQL &= "1" Else sSQL &= "0"
                sSQL &= ", " & Me.InnerRingRadius & ", " & Me.OuterRingRadius & ", " & Me.RingDiffuse & ", " & AxisAngle & _
                ", " & RingMineralID & ", " & RingConcentration & ")"
            Else
                'UPDATE
                sSQL = "UPDATE tblPlanet SET PlanetTypeID = " & PlanetTypeID & ", PlanetName = '" & MakeDBStr(BytesToString(PlanetName)) & _
                  "', PlanetSizeID = " & PlanetSizeID & ", ParentID = " & ParentSystem.ObjectID & ", LocX = " & _
                  LocX & ", LocY = " & LocY & ", LocZ = " & LocZ & ", Vegetation = " & Vegetation & ", Atmosphere = " & _
                  Atmosphere & ", Hydrosphere = " & Hydrosphere & ", Gravity = " & Gravity & ", SurfaceTemperature = " & _
                  SurfaceTemperature & ", RotationDelay = " & RotationDelay & ", PlanetRadius = " & PlanetRadius & _
                  ", PlayerSpawns = " & PlayerSpawns & ", SpawnLocked = "
                If SpawnLocked = True Then sSQL &= "1" Else sSQL &= "0"
                sSQL &= ", RingInnerRadius = " & Me.InnerRingRadius & ", RingOuterRadius = " & Me.OuterRingRadius & ", RingDiffuse = " & Me.RingDiffuse
                sSQL &= ", AxisAngle = " & AxisAngle & ", RingMineralID = " & RingMineralID & ", RingMineralConcentration = " & RingConcentration & " WHERE PlanetID = " & ObjectID
            End If
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            If oComm.ExecuteNonQuery() = 0 Then
                Err.Raise(-1, "SaveObject", "No records affected when saving object!")
            End If
            If ObjectID = -1 Then
                Dim oData As OleDb.OleDbDataReader
                oComm = Nothing
                sSQL = "SELECT MAX(PlanetID) FROM tblPlanet WHERE PlanetName = '" & MakeDBStr(BytesToString(PlanetName)) & "'"
                oComm = New OleDb.OleDbCommand(sSQL, goCN)
                oData = oComm.ExecuteReader(CommandBehavior.Default)
                If oData.Read Then
                    ObjectID = CInt(oData(0))
                End If
                oData.Close()
                oData = Nothing
            End If
            bResult = True
        Catch
            LogEvent(GlobalVars.LogEventType.CriticalError, "Unable to save object " & ObjectID & " of type " & ObjTypeID & ". Reason: " & Err.Description)
        Finally
            oComm = Nothing
        End Try
        Return bResult
    End Function

    Private Shared mlWidthTotals() As Int32
    Public ReadOnly Property PlanetWidthTotal() As Int32
        Get
            If mlWidthTotals Is Nothing Then
                ReDim mlWidthTotals(4)
                mlWidthTotals(0) = gl_TINY_PLANET_CELL_SPACING * 240
                mlWidthTotals(1) = gl_SMALL_PLANET_CELL_SPACING * 240
                mlWidthTotals(2) = gl_MEDIUM_PLANET_CELL_SPACING * 240
                mlWidthTotals(3) = gl_LARGE_PLANET_CELL_SPACING * 240
                mlWidthTotals(4) = gl_HUGE_PLANET_CELL_SPACING * 240
            End If
            Return mlWidthTotals(PlanetSizeID)
        End Get
    End Property
 End Class
