Option Strict On

Public Enum AgentEffectType As Byte
    eMorale = 0
    eProduction = 1
    eCargoBay = 2
    eHangarBay = 3
    eTradeIncome = 4
    'Per colony, per morale modifier type
    eColonyTaxMorale = 5
    eColonyHousingMorale = 6
    eColonyUnemploymentMorale = 7

    eSiphonTradepost = 8
    eSiphonMining = 9
    eSiphonCorporation = 10
    eGovernorMorale = 11
    eIncreasePowerNeed = 12
End Enum

Public Class AgentEffect
    Public lStartCycle As Int32     'when this effect began
    Public lDuration As Int32       'in cycles, how long this effect will last
    Public yType As AgentEffectType 'the type of effect this applies
    Public lAmount As Int32         'amount of effect applied
    Public bAmountAsPerc As Boolean 'if true, indicates that lAmount is percentage based (needs to be divided by 100)

    Public lLastVerification As Int32       'last cycle this agent effect was validated true

    Public lCausedByID As Int32     'playerid that caused the effect
End Class
