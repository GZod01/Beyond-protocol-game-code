Option Strict On

Public Class UnitLimit
    Public yHullType As Byte        'indicates what hull to apply this limit to, 255 indicates not to refer to hulltype, 254 indicates All Unit limit, 253 indicates Ground unit, 252 indicates Flying Unit
    Public lHullSize As Int32       'maximum hullsize
    Public lMaxCnt As Int32         'maximum cnt

    Public Function GetDisplayText() As String
        Dim sValue As String = ""
        Select Case yHullType
            Case 254
                sValue = "All Units: "
            Case 253
                sValue = "Ground Units: "
            Case 252
                sValue = "Flying Units: "
            Case Else
                sValue = Base_Tech.GetHullTypeName(yHullType) & ": "
        End Select

        If lMaxCnt > 0 Then
            sValue &= " Max " & lMaxCnt.ToString
        End If
        If lHullSize > 0 Then
            sValue &= " Max Hull " & lHullSize.ToString
        End If

        Return sValue
    End Function

    'return 0 if this limit does not apply, 1 if it does apply, 2 if it applies but the hull size is exceeded
    Public Function DoesDefFitLimit(ByVal lHull As Int32, ByVal yType As Byte) As Byte
        If yHullType = 255 Then
            If lHullSize > 1 AndAlso lHull > lHullSize Then Return 2 Else Return 1
        Else
            If yType = yHullType Then
                If lHullSize > 1 AndAlso lHull > lHullSize Then Return 2 Else Return 1
            End If
        End If

        Return 0
    End Function
End Class
