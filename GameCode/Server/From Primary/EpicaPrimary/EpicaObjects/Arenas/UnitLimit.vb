Option Strict On

Public Class UnitLimit
    Public yHullType As Byte        'indicates what hull to apply this limit to, 255 indicates not to refer to hulltype, 254 indicates All Unit limit, 253 indicates Ground unit, 252 indicates Flying Unit
    Public lHullSize As Int32       'maximum hullsize
    Public lMaxCnt As Int32         'maximum cnt

    'return 0 if this limit does not apply, 1 if it does apply, 2 if it applies but the hull size is exceeded
    Public Function DoesDefFitLimit(ByVal oDef As Epica_Entity_Def) As Byte
        If yHullType = 255 Then
            If lHullSize > 1 AndAlso oDef.HullSize > lHullSize Then Return 2 Else Return 1
        Else
            If oDef.oPrototype Is Nothing = False Then
                If oDef.oPrototype.oHullTech Is Nothing = False Then
                    If HullTech.GetHullTypeID(oDef.oPrototype.oHullTech.yTypeID, oDef.oPrototype.oHullTech.ySubTypeID) = yHullType Then
                        If lHullSize > 1 AndAlso oDef.HullSize > lHullSize Then Return 2 Else Return 1
                    End If
                End If
            End If
        End If

        Return 0
    End Function
End Class
