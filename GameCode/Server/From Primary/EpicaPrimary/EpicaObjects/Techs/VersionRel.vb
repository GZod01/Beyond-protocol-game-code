Option Strict On

Public Class VersionRel
    Public lVersionNumber As Int32
    Public lOtherVersion As Int32
    Public lNoisePerc As Int32

    Public Sub New(ByVal lNum As Int32, ByVal lOther As Int32, ByVal lPerc As Int32)
        lVersionNumber = lNum
        lOtherVersion = lOther
        lNoisePerc = lPerc
    End Sub
End Class
