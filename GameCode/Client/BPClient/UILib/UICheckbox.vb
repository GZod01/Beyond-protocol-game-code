Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

Public Class UICheckBox
    Inherits UIOption

    Public Enum eCheckTypes As Integer
        eLargeCheckBox = 0          'this is now the same as SmallCheck!!!
        eSmallCheck
        eSmallX
        eSmallBlock
    End Enum
    Private mlType As eCheckTypes

    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)

        'load default images
        DisplayType = eCheckTypes.eLargeCheckBox
    End Sub

    Public Shadows Property DisplayType() As eCheckTypes
        Get
            Return mlType
        End Get
        Set(ByVal Value As eCheckTypes)
            mlType = Value

            'If mlType = eCheckTypes.eLargeCheckBox Then
            '    ControlImageRect_Disabled = System.Drawing.Rectangle.FromLTRB(120, 0, 143, 24)
            '    ControlImageRect_Normal = System.Drawing.Rectangle.FromLTRB(144, 0, 167, 24)
            '    ControlImageRect_Pressed = System.Drawing.Rectangle.FromLTRB(168, 0, 191, 24)
            'Else
            ControlImageRect_Disabled = grc_UI(elInterfaceRectangle.eCheck_Disabled)
            ControlImageRect_Normal = grc_UI(elInterfaceRectangle.eCheck_Unchecked)

            If mlType = eCheckTypes.eSmallBlock Then
                ControlImageRect_Pressed = grc_UI(elInterfaceRectangle.eCheck_Blocked)
            ElseIf mlType = eCheckTypes.eSmallCheck OrElse mlType = eCheckTypes.eLargeCheckBox Then
                ControlImageRect_Pressed = grc_UI(elInterfaceRectangle.eCheck_Checked)
            Else : ControlImageRect_Pressed = grc_UI(elInterfaceRectangle.eCheck_Xed)
            End If
            'End If
        End Set
    End Property

End Class
