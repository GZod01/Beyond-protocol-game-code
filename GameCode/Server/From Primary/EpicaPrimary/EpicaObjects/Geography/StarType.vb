Public Class StarType
    Public StarTypeID As Byte       'it is stored as a long
    Public StarTypeName(19) As Byte
    Public StarTypeAttrs As Int32   'bit-wise attributes
    Public StarTexture(19) As Byte
    Public StarRadius As Int32
    Public HeatIndex As Byte        '0 to 255, used for determining temperatures during Galaxy Generation, 100 is Sol
    Public MatDiffuse As Int32
    Public MatEmissive As Int32
    Public StarRarity As Byte       'how much of the galaxy is made up of this star, NOTE: the sum of these values should = 100
    Public LightRange As Int32
    Public LightDiffuse As Int32
    Public LightAmbient As Int32
    Public LightSpecular As Int32
    Public LightAtt1 As Single
    Public LightAtt2 As Single
    Public LightAtt3 As Single
    Public StarMapRectIdx As Byte

    Private mbStringReady As Boolean = False
    Private mySendString() As Byte

    'the data of a star type NEVER changes between reboots and there is no need to save it

    Public Function GetObjAsString() As Byte()
        If mbStringReady = False Then
            ReDim mySendString(87)      '0 to 87 = 88 bytes

            mySendString(0) = StarTypeID
            StarTypeName.CopyTo(mySendString, 1)
            System.BitConverter.GetBytes(StarTypeAttrs).CopyTo(mySendString, 21)
            StarTexture.CopyTo(mySendString, 25)
            'No one cares about min and Max radius
            mySendString(45) = HeatIndex
            System.BitConverter.GetBytes(MatDiffuse).CopyTo(mySendString, 46)
            System.BitConverter.GetBytes(MatEmissive).CopyTo(mySendString, 50)
            mySendString(54) = StarRarity       'more for show then anything
            System.BitConverter.GetBytes(LightRange).CopyTo(mySendString, 55)
            System.BitConverter.GetBytes(LightDiffuse).CopyTo(mySendString, 59)
            System.BitConverter.GetBytes(LightAmbient).CopyTo(mySendString, 63)
            System.BitConverter.GetBytes(LightSpecular).CopyTo(mySendString, 67)
            System.BitConverter.GetBytes(LightAtt1).CopyTo(mySendString, 71)
            System.BitConverter.GetBytes(LightAtt2).CopyTo(mySendString, 75)
            System.BitConverter.GetBytes(LightAtt3).CopyTo(mySendString, 79)
            mySendString(83) = StarMapRectIdx
            System.BitConverter.GetBytes(StarRadius).CopyTo(mySendString, 84)
        End If

        Return mySendString
    End Function
End Class
