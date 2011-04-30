Module GlobalVars
    Public Enum elInterfaceRectangle As Int32
        eButton_Normal = 0
        eButton_Disabled = 1
        eButton_Down = 2
        eSmall_Button_Normal = 3
        eSmall_Button_Disabled = 4
        eSmall_Button_Down = 5
        eUpArrow_Button_Normal = 6
        eUpArrow_Button_Disabled = 7
        eUpArrow_Button_Down = 8
        eLeftArrow_Button_Normal = 9
        eLeftArrow_Button_Disabled = 10
        eLeftArrow_Button_Down = 11
        eRightArrow_Button_Normal = 12
        eRightArrow_Button_Disabled = 13
        eRightArrow_Button_Down = 14
        eDownArrow_Button_Normal = 15
        eDownArrow_Button_Disabled = 16
        eDownArrow_Button_Down = 17

        eDirectionalArrow = 18
        eLightning = 19
        eSingleDude = 20
        eHappySadFace = 21

        eSphere = 22
        ePlanetOrbit = 23

        eXPRank_EmptyDot = 24
        eXPRank_SolidDot = 25
        eXPRank_Arrow = 26
        eXPRank_Bar = 27

        eWrench = 28
        eDemolish = 29
        eAlarm = 30
        eKeyButton = 31
        eLock = 32

        eCheck_Unchecked = 33
        eCheck_Disabled = 34
        eCheck_Checked = 35
        eCheck_Xed = 36
        eCheck_Blocked = 37

        eOption_Normal = 38
        eOption_Disabled = 39
        eOption_Marked = 40

        eLeftExpander = 41
        eRightExpander = 42
        eUpExpander = 43
        eDownExpander = 44

        eColonyStatsMinimizer = 45

        eEmailIcon = 46
        ePlanetIcon = 47

        eMinBar_0 = 48
        eMinBar_1 = 49
        eMinBar_2 = 50
        eMinBar_3 = 51
        eMinBar_4 = 52
        eMinBar_5 = 53
        eMinBar_6 = 54
        eMinBar_7 = 55
        eMinBar_8 = 56
        eMinBar_9 = 57
        eMinBar_10 = 58

        eWhiteBox = 59

        eAgentScrollLeftButton = 60
        eAgentScrollRightButton = 61

        eQuickbar_Help = 62
        eQuickbar_Email = 63
        eQuickbar_Battlegroup = 64
        eQuickbar_Trade = 65
        eQuickbar_Diplomacy = 66
        eQuickbar_ColonyStats = 67
        eQuickbar_Budget = 68
        eQuickbar_Mining = 69
        eQuickbar_Agent = 70
        eQuickbar_Formations = 71
        eQuickbar_ColonyResearch = 72
        eQuickbar_AvailResources = 73
        eQuickbar_ChatConfig = 74
        eQuickbar_Command = 75
        eQuickBar_Senate = 76

        eLastElement
    End Enum
    Public grc_UI() As Rectangle
    Public Sub CreateGlobalRectangleList()

        ReDim grc_UI(elInterfaceRectangle.eLastElement - 1)

        grc_UI(elInterfaceRectangle.eAgentScrollLeftButton) = New Rectangle(214, 107, 21, 50)
        grc_UI(elInterfaceRectangle.eAgentScrollRightButton) = New Rectangle(235, 107, 21, 50)
        grc_UI(elInterfaceRectangle.eAlarm) = New Rectangle(192, 96, 16, 16)
        grc_UI(elInterfaceRectangle.eButton_Disabled) = New Rectangle(0, 33, 120, 32)
        grc_UI(elInterfaceRectangle.eButton_Down) = New Rectangle(1, 64, 118, 32)
        grc_UI(elInterfaceRectangle.eButton_Normal) = New Rectangle(1, 0, 118, 32)
        grc_UI(elInterfaceRectangle.eCheck_Blocked) = New Rectangle(40, 215, 10, 10)
        grc_UI(elInterfaceRectangle.eCheck_Checked) = New Rectangle(20, 215, 10, 10)
        grc_UI(elInterfaceRectangle.eCheck_Disabled) = New Rectangle(10, 215, 10, 10)
        grc_UI(elInterfaceRectangle.eCheck_Unchecked) = New Rectangle(0, 215, 10, 10)
        grc_UI(elInterfaceRectangle.eCheck_Xed) = New Rectangle(30, 215, 10, 10)
        grc_UI(elInterfaceRectangle.eColonyStatsMinimizer) = New Rectangle(0, 246, 159, 10)
        grc_UI(elInterfaceRectangle.eDemolish) = New Rectangle(192, 80, 16, 16)
        grc_UI(elInterfaceRectangle.eDirectionalArrow) = New Rectangle(100, 98, 16, 16)
        grc_UI(elInterfaceRectangle.eDownArrow_Button_Disabled) = New Rectangle(120, 96, 24, 24)
        grc_UI(elInterfaceRectangle.eDownArrow_Button_Down) = New Rectangle(168, 96, 24, 24)
        grc_UI(elInterfaceRectangle.eDownArrow_Button_Normal) = New Rectangle(144, 96, 24, 24)
        grc_UI(elInterfaceRectangle.eDownExpander) = New Rectangle(27, 236, 10, 10)
        grc_UI(elInterfaceRectangle.eEmailIcon) = New Rectangle(161, 239, 32, 17)
        grc_UI(elInterfaceRectangle.eHappySadFace) = New Rectangle(110, 173, 16, 16)
        grc_UI(elInterfaceRectangle.eKeyButton) = New Rectangle(208, 80, 16, 16)
        grc_UI(elInterfaceRectangle.eLeftArrow_Button_Disabled) = New Rectangle(120, 48, 24, 24)
        grc_UI(elInterfaceRectangle.eLeftArrow_Button_Down) = New Rectangle(168, 48, 24, 24)
        grc_UI(elInterfaceRectangle.eLeftArrow_Button_Normal) = New Rectangle(144, 48, 24, 24)

        grc_UI(elInterfaceRectangle.eLeftExpander) = New Rectangle(0, 235, 8, 10)

        grc_UI(elInterfaceRectangle.eLightning) = New Rectangle(111, 144, 16, 16)
        grc_UI(elInterfaceRectangle.eLock) = New Rectangle(208, 64, 16, 16)

        For X As Int32 = 0 To 10
            grc_UI(elInterfaceRectangle.eMinBar_0 + X) = New Rectangle(193, 157 + (X * 9), 64, 9)
        Next X

        grc_UI(elInterfaceRectangle.eOption_Disabled) = New Rectangle(10, 225, 10, 10)
        grc_UI(elInterfaceRectangle.eOption_Marked) = New Rectangle(20, 225, 10, 10)
        grc_UI(elInterfaceRectangle.eOption_Normal) = New Rectangle(0, 225, 10, 10)
        grc_UI(elInterfaceRectangle.ePlanetIcon) = New Rectangle(161, 208, 32, 32)
        grc_UI(elInterfaceRectangle.ePlanetOrbit) = New Rectangle(144, 144, 16, 16)
        grc_UI(elInterfaceRectangle.eQuickbar_Agent) = New Rectangle(64, 160, 32, 32)
        grc_UI(elInterfaceRectangle.eQuickbar_AvailResources) = New Rectangle(128, 192, 32, 32)
        grc_UI(elInterfaceRectangle.eQuickbar_Battlegroup) = New Rectangle(0, 160, 32, 32)
        grc_UI(elInterfaceRectangle.eQuickbar_Budget) = New Rectangle(64, 96, 32, 32)
        grc_UI(elInterfaceRectangle.eQuickbar_ChatConfig) = New Rectangle(129, 160, 32, 32)
        grc_UI(elInterfaceRectangle.eQuickbar_ColonyResearch) = New Rectangle(96, 192, 32, 32)
        grc_UI(elInterfaceRectangle.eQuickbar_ColonyStats) = New Rectangle(32, 160, 32, 32)
        grc_UI(elInterfaceRectangle.eQuickbar_Command) = New Rectangle(224, 64, 32, 32)
        grc_UI(elInterfaceRectangle.eQuickbar_Diplomacy) = New Rectangle(32, 128, 32, 32)
        grc_UI(elInterfaceRectangle.eQuickbar_Email) = New Rectangle(0, 128, 32, 32)
        grc_UI(elInterfaceRectangle.eQuickbar_Formations) = New Rectangle(64, 192, 32, 32)
        grc_UI(elInterfaceRectangle.eQuickbar_Help) = New Rectangle(0, 96, 32, 32)
        grc_UI(elInterfaceRectangle.eQuickbar_Mining) = New Rectangle(64, 128, 32, 32)
        grc_UI(elInterfaceRectangle.eQuickBar_Senate) = New Rectangle(182, 120, 32, 32)
        grc_UI(elInterfaceRectangle.eQuickbar_Trade) = New Rectangle(32, 96, 32, 32)
        grc_UI(elInterfaceRectangle.eRightArrow_Button_Disabled) = New Rectangle(120, 72, 24, 24)
        grc_UI(elInterfaceRectangle.eRightArrow_Button_Down) = New Rectangle(168, 72, 24, 24)
        grc_UI(elInterfaceRectangle.eRightArrow_Button_Normal) = New Rectangle(144, 72, 24, 24)

        grc_UI(elInterfaceRectangle.eRightExpander) = New Rectangle(7, 235, 8, 10)

        grc_UI(elInterfaceRectangle.eSingleDude) = New Rectangle(111, 158, 15, 15)
        grc_UI(elInterfaceRectangle.eSmall_Button_Disabled) = New Rectangle(120, 0, 24, 24)
        grc_UI(elInterfaceRectangle.eSmall_Button_Down) = New Rectangle(168, 0, 24, 24)
        grc_UI(elInterfaceRectangle.eSmall_Button_Normal) = New Rectangle(144, 0, 24, 24)
        grc_UI(elInterfaceRectangle.eSphere) = New Rectangle(128, 144, 16, 16)
        grc_UI(elInterfaceRectangle.eUpArrow_Button_Disabled) = New Rectangle(120, 24, 24, 24)
        grc_UI(elInterfaceRectangle.eUpArrow_Button_Down) = New Rectangle(168, 24, 24, 24)
        grc_UI(elInterfaceRectangle.eUpArrow_Button_Normal) = New Rectangle(144, 24, 24, 24)

        grc_UI(elInterfaceRectangle.eUpExpander) = New Rectangle(15, 236, 10, 10)

        grc_UI(elInterfaceRectangle.eWhiteBox) = New Rectangle(192, 0, 62, 64)
        grc_UI(elInterfaceRectangle.eWrench) = New Rectangle(192, 64, 16, 16)
        grc_UI(elInterfaceRectangle.eXPRank_Arrow) = New Rectangle(161, 192, 16, 16)
        grc_UI(elInterfaceRectangle.eXPRank_Bar) = New Rectangle(177, 191, 16, 16)
        grc_UI(elInterfaceRectangle.eXPRank_EmptyDot) = New Rectangle(161, 176, 16, 16)
        grc_UI(elInterfaceRectangle.eXPRank_SolidDot) = New Rectangle(177, 176, 16, 16)


    End Sub

End Module
