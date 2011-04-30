'Option Strict On

'Imports Microsoft.DirectX
'Imports Microsoft.DirectX.Direct3D

''Interface created from Interface Builder
'Public Class frmTrade
'    Inherits UIWindow

'    Private lblTitle As UILabel
'    Private lnDiv1 As UILine
'    Private lblSource As UILabel
'    Private lblDest As UILabel
'    Private lblTradeables As UILabel
'    Private lblQuantity As UILabel
'    Private lnDiv2 As UILine
'    Private lblOffer As UILabel
'    Private lblTheyOffer As UILabel
'    Private lblNotes As UILabel
'    Private lnDiv3 As UILine
'    Private lnDiv4 As UILine

'    Private txtNotesToMe As UITextBox
'    Private txtQuantity As UITextBox
'    Private txtNotesToThem As UITextBox
'    Private cboDest As UIComboBox
'    Private chkAccept As UICheckBox

'    Private WithEvents chkTheirAccept As UICheckBox

'    Private WithEvents cboSource As UIComboBox

'    Private WithEvents lstTradeables As UIListBox
'    Private WithEvents lstOffer As UIListBox
'    Private WithEvents lstGetting As UIListBox

'    Private WithEvents btnRemove As UIButton
'    Private WithEvents btnAdd As UIButton
'    Private WithEvents btnSubmit As UIButton
'    Private WithEvents btnReject As UIButton
'    Private WithEvents btnClose As UIButton

'    Private mlLastUpdate As Int32

'    'Used for new trades
'    Private mlOtherPlayerID As Int32 = -1
'    'Used for previously created trades
'    Private moTrade As Trade = Nothing

'    Private mblQtys() As Int64

'    Public Sub New(ByRef oUILib As UILib)
'        MyBase.New(oUILib)

'        'frmTrade initial props
'        With Me
'            .ControlName = "frmTrade"
'            .Left = 213
'            .Top = 92
'            .Width = 560
'            .Height = 455
'            .Enabled = True
'            .Visible = True
'            .BorderColor = muSettings.InterfaceBorderColor
'            .FillColor = muSettings.InterfaceFillColor
'            .FullScreen = False
'            .BorderLineWidth = 1
'            .Moveable = True
'            .mbAcceptReprocessEvents = True
'        End With

'        'lblTitle initial props
'        lblTitle = New UILabel(oUILib)
'        With lblTitle
'            .ControlName = "lblTitle"
'            .Left = 5
'            .Top = 5
'            .Width = 650
'            .Height = 18
'            .Enabled = True
'            .Visible = True
'            .Caption = "Trade Agreement With GREMan"
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 11.25F, FontStyle.Bold, GraphicsUnit.Point, 0))
'            .DrawBackImage = False
'            .FontFormat = CType(4, DrawTextFormat)
'        End With
'        Me.AddChild(CType(lblTitle, UIControl))

'        'btnClose initial props
'        btnClose = New UIButton(oUILib)
'        With btnClose
'            .ControlName = "btnClose"
'            .Left = 536
'            .Top = 1
'            .Width = 24
'            .Height = 24
'            .Enabled = True
'            .Visible = True
'            .Caption = "X"
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
'            .DrawBackImage = True
'            .FontFormat = CType(5, DrawTextFormat)
'            .ControlImageRect = New Rectangle(0, 0, 120, 32)
'        End With
'        Me.AddChild(CType(btnClose, UIControl))

'        'lnDiv1 initial props
'        lnDiv1 = New UILine(oUILib)
'        With lnDiv1
'            .ControlName = "lnDiv1"
'            .Left = 0
'            .Top = 25
'            .Width = 560
'            .Height = 0
'            .Enabled = True
'            .Visible = True
'            .BorderColor = muSettings.InterfaceBorderColor
'        End With
'        Me.AddChild(CType(lnDiv1, UIControl))

'        'lblSource initial props
'        lblSource = New UILabel(oUILib)
'        With lblSource
'            .ControlName = "lblSource"
'            .Left = 5
'            .Top = 30
'            .Width = 70
'            .Height = 18
'            .Enabled = True
'            .Visible = True
'            .Caption = "Distributor:"
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
'            .DrawBackImage = False
'            .FontFormat = CType(4, DrawTextFormat)
'        End With
'        Me.AddChild(CType(lblSource, UIControl))

'        'lblDest initial props
'        lblDest = New UILabel(oUILib)
'        With lblDest
'            .ControlName = "lblDest"
'            .Left = 5
'            .Top = 55
'            .Width = 70
'            .Height = 18
'            .Enabled = True
'            .Visible = True
'            .Caption = "Receiver:"
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
'            .DrawBackImage = False
'            .FontFormat = CType(4, DrawTextFormat)
'        End With
'        Me.AddChild(CType(lblDest, UIControl))

'        'lblTradables initial props
'        lblTradeables = New UILabel(oUILib)
'        With lblTradeables
'            .ControlName = "lblTradeables"
'            .Left = 5
'            .Top = 80
'            .Width = 110
'            .Height = 18
'            .Enabled = True
'            .Visible = True
'            .Caption = "Tradeable Items:"
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
'            .DrawBackImage = False
'            .FontFormat = CType(4, DrawTextFormat)
'        End With
'        Me.AddChild(CType(lblTradeables, UIControl))

'        'lstTradables initial props
'        lstTradeables = New UIListBox(oUILib)
'        With lstTradeables
'            .ControlName = "lstTradeables"
'            .Left = 5
'            .Top = 100
'            .Width = 270
'            .Height = 100
'            .Enabled = True
'            .Visible = True
'            .BorderColor = muSettings.InterfaceBorderColor
'            .FillColor = muSettings.InterfaceFillColor
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
'            .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
'            .mbAcceptReprocessEvents = True
'        End With
'        Me.AddChild(CType(lstTradeables, UIControl))

'        'lblQuantity initial props
'        lblQuantity = New UILabel(oUILib)
'        With lblQuantity
'            .ControlName = "lblQuantity"
'            .Left = 5
'            .Top = 205
'            .Width = 50
'            .Height = 18
'            .Enabled = True
'            .Visible = True
'            .Caption = "Quantity:"
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
'            .DrawBackImage = False
'            .FontFormat = CType(4, DrawTextFormat)
'        End With
'        Me.AddChild(CType(lblQuantity, UIControl))

'        'txtQuantity initial props
'        txtQuantity = New UITextBox(oUILib)
'        With txtQuantity
'            .ControlName = "txtQuantity"
'            .Left = 65
'            .Top = 205
'            .Width = 72
'            .Height = 18
'            .Enabled = True
'            .Visible = True
'            .Caption = "1"
'            .ForeColor = muSettings.InterfaceTextBoxForeColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
'            .DrawBackImage = False
'            .FontFormat = CType(4, DrawTextFormat)
'            .BackColorEnabled = muSettings.InterfaceTextBoxFillColor
'            .BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
'            .MaxLength = 0
'            .BorderColor = muSettings.InterfaceBorderColor
'        End With
'        Me.AddChild(CType(txtQuantity, UIControl))

'        'btnAdd initial props
'        btnAdd = New UIButton(oUILib)
'        With btnAdd
'            .ControlName = "btnAdd"
'            .Left = 150
'            .Top = 205
'            .Width = 70
'            .Height = 21
'            .Enabled = True
'            .Visible = True
'            .Caption = "Add"
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
'            .DrawBackImage = True
'            .FontFormat = CType(5, DrawTextFormat)
'            .ControlImageRect = New Rectangle(0, 0, 120, 32)
'        End With
'        Me.AddChild(CType(btnAdd, UIControl))

'        'lnDiv2 initial props
'        lnDiv2 = New UILine(oUILib)
'        With lnDiv2
'            .ControlName = "lnDiv2"
'            .Left = 280
'            .Top = 25
'            .Width = 0
'            .Height = 385
'            .Enabled = True
'            .Visible = True
'            .BorderColor = muSettings.InterfaceBorderColor
'        End With
'        Me.AddChild(CType(lnDiv2, UIControl))

'        'lblOffer initial props
'        lblOffer = New UILabel(oUILib)
'        With lblOffer
'            .ControlName = "lblOffer"
'            .Left = 285
'            .Top = 30
'            .Width = 102
'            .Height = 18
'            .Enabled = True
'            .Visible = True
'            .Caption = "You are Offering:"
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
'            .DrawBackImage = False
'            .FontFormat = CType(4, DrawTextFormat)
'        End With
'        Me.AddChild(CType(lblOffer, UIControl))

'        'lstOffer initial props
'        lstOffer = New UIListBox(oUILib)
'        With lstOffer
'            .ControlName = "lstOffer"
'            .Left = 285
'            .Top = 50
'            .Width = 270
'            .Height = 150
'            .Enabled = True
'            .Visible = True
'            .BorderColor = muSettings.InterfaceBorderColor
'            .FillColor = muSettings.InterfaceFillColor
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
'            .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
'            .mbAcceptReprocessEvents = True
'        End With
'        Me.AddChild(CType(lstOffer, UIControl))

'        'lblTheyOffer initial props
'        lblTheyOffer = New UILabel(oUILib)
'        With lblTheyOffer
'            .ControlName = "lblTheyOffer"
'            .Left = 285
'            .Top = 235
'            .Width = 112
'            .Height = 18
'            .Enabled = True
'            .Visible = True
'            .Caption = "They are Offering:"
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
'            .DrawBackImage = False
'            .FontFormat = CType(4, DrawTextFormat)
'        End With
'        Me.AddChild(CType(lblTheyOffer, UIControl))

'        'lstGetting initial props
'        lstGetting = New UIListBox(oUILib)
'        With lstGetting
'            .ControlName = "lstGetting"
'            .Left = 285
'            .Top = 255
'            .Width = 270
'            .Height = 150
'            .Enabled = True
'            .Visible = True
'            .BorderColor = muSettings.InterfaceBorderColor
'            .FillColor = muSettings.InterfaceFillColor
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
'            .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
'            .mbAcceptReprocessEvents = True
'        End With
'        Me.AddChild(CType(lstGetting, UIControl))

'        'btnRemove initial props
'        btnRemove = New UIButton(oUILib)
'        With btnRemove
'            .ControlName = "btnRemove"
'            .Left = 455
'            .Top = 205
'            .Width = 100
'            .Height = 21
'            .Enabled = True
'            .Visible = True
'            .Caption = "Remove"
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
'            .DrawBackImage = True
'            .FontFormat = CType(5, DrawTextFormat)
'            .ControlImageRect = New Rectangle(0, 0, 120, 32)
'        End With
'        Me.AddChild(CType(btnRemove, UIControl))

'        'lblNotes initial props
'        lblNotes = New UILabel(oUILib)
'        With lblNotes
'            .ControlName = "lblNotes"
'            .Left = 5
'            .Top = 235
'            .Width = 111
'            .Height = 18
'            .Enabled = True
'            .Visible = True
'            .Caption = "Messages:"
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
'            .DrawBackImage = False
'            .FontFormat = CType(4, DrawTextFormat)
'        End With
'        Me.AddChild(CType(lblNotes, UIControl))

'        'lnDiv3 initial props
'        lnDiv3 = New UILine(oUILib)
'        With lnDiv3
'            .ControlName = "lnDiv3"
'            .Left = 0
'            .Top = 230
'            .Width = 280
'            .Height = 0
'            .Enabled = True
'            .Visible = True
'            .BorderColor = muSettings.InterfaceBorderColor
'        End With
'        Me.AddChild(CType(lnDiv3, UIControl))

'        'txtNotesToMe initial props
'        txtNotesToMe = New UITextBox(oUILib)
'        With txtNotesToMe
'            .ControlName = "txtNotesToMe"
'            .Left = 5
'            .Top = 255
'            .Width = 270
'            .Height = 70
'            .Enabled = True
'            .Visible = True
'            .Caption = ""
'            .ForeColor = muSettings.InterfaceTextBoxForeColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
'            .DrawBackImage = False
'            .FontFormat = CType(0, DrawTextFormat)
'            .BackColorEnabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
'            .BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
'            .MaxLength = 0
'            .BorderColor = muSettings.InterfaceBorderColor
'            .Locked = True
'            .MultiLine = True
'        End With
'        Me.AddChild(CType(txtNotesToMe, UIControl))

'        'txtNotesToThem initial props
'        txtNotesToThem = New UITextBox(oUILib)
'        With txtNotesToThem
'            .ControlName = "txtNotesToThem"
'            .Left = 5
'            .Top = 335
'            .Width = 270
'            .Height = 70
'            .Enabled = True
'            .Visible = True
'            .Caption = ""
'            .ForeColor = muSettings.InterfaceTextBoxForeColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
'            .DrawBackImage = False
'            .FontFormat = CType(0, DrawTextFormat)
'            .BackColorEnabled = muSettings.InterfaceTextBoxFillColor
'            .BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
'            .MaxLength = 0
'            .BorderColor = muSettings.InterfaceBorderColor
'            .Locked = False
'            .MultiLine = True
'        End With
'        Me.AddChild(CType(txtNotesToThem, UIControl))

'        'btnSubmit initial props
'        btnSubmit = New UIButton(oUILib)
'        With btnSubmit
'            .ControlName = "btnSubmit"
'            .Left = 325
'            .Top = 422
'            .Width = 100
'            .Height = 25
'            .Enabled = True
'            .Visible = True
'            .Caption = "Submit"
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
'            .DrawBackImage = True
'            .FontFormat = CType(5, DrawTextFormat)
'            .ControlImageRect = New Rectangle(0, 0, 120, 32)
'        End With
'        Me.AddChild(CType(btnSubmit, UIControl))

'        'btnReject initial props
'        btnReject = New UIButton(oUILib)
'        With btnReject
'            .ControlName = "btnReject"
'            .Left = 455
'            .Top = 422
'            .Width = 100
'            .Height = 25
'            .Enabled = True
'            .Visible = True
'            .Caption = "Reject"
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
'            .DrawBackImage = True
'            .FontFormat = CType(5, DrawTextFormat)
'            .ControlImageRect = New Rectangle(0, 0, 120, 32)
'        End With
'        Me.AddChild(CType(btnReject, UIControl))

'        'lnDiv4 initial props
'        lnDiv4 = New UILine(oUILib)
'        With lnDiv4
'            .ControlName = "lnDiv4"
'            .Left = 0
'            .Top = 410
'            .Width = 560
'            .Height = 0
'            .Enabled = True
'            .Visible = True
'            .BorderColor = muSettings.InterfaceBorderColor
'        End With
'        Me.AddChild(CType(lnDiv4, UIControl))

'        'chkAccept initial props
'        chkAccept = New UICheckBox(oUILib)
'        With chkAccept
'            .ControlName = "chkAccept"
'            .Left = 5
'            .Top = 415
'            .Width = 162
'            .Height = 18
'            .Enabled = True
'            .Visible = True
'            .Caption = "Accept This Agreement"
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
'            .DrawBackImage = False
'            .FontFormat = CType(6, DrawTextFormat)
'            .Value = False
'            .DisplayType = UICheckBox.eCheckTypes.eSmallCheck
'        End With
'        Me.AddChild(CType(chkAccept, UIControl))

'        'chkTheirAccept initial props
'        chkTheirAccept = New UICheckBox(oUILib)
'        With chkTheirAccept
'            .ControlName = "chkTheirAccept"
'            .Left = 5
'            .Top = 435
'            .Width = 161
'            .Height = 18
'            .Enabled = True
'            .Visible = True
'            .Caption = "Their Acceptance"
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
'            .DrawBackImage = False
'            .FontFormat = CType(6, DrawTextFormat)
'            .Value = False
'            .DisplayType = UICheckBox.eCheckTypes.eSmallCheck
'        End With
'        Me.AddChild(CType(chkTheirAccept, UIControl))

'        'cboDest initial props
'        cboDest = New UIComboBox(oUILib)
'        With cboDest
'            .ControlName = "cboDest"
'            .Left = 75
'            .Top = 55
'            .Width = 200
'            .Height = 18
'            .Enabled = True
'            .Visible = True
'            .BorderColor = muSettings.InterfaceBorderColor
'            .FillColor = System.Drawing.Color.FromArgb(255, 255, 255, 255)
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
'            .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
'            .DropDownListBorderColor = muSettings.InterfaceBorderColor
'            .mbAcceptReprocessEvents = True
'        End With
'        Me.AddChild(CType(cboDest, UIControl))

'        'cboSource initial props
'        cboSource = New UIComboBox(oUILib)
'        With cboSource
'            .ControlName = "cboSource"
'            .Left = 75
'            .Top = 30
'            .Width = 200
'            .Height = 18
'            .Enabled = True
'            .Visible = True
'            .BorderColor = muSettings.InterfaceBorderColor
'            .FillColor = System.Drawing.Color.FromArgb(255, 255, 255, 255)
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
'            .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
'            .DropDownListBorderColor = muSettings.InterfaceBorderColor
'            .mbAcceptReprocessEvents = True
'        End With
'        Me.AddChild(CType(cboSource, UIControl))

'        MyBase.moUILib.RemoveWindow(Me.ControlName)
'        MyBase.moUILib.AddWindow(Me)

'        Dim yMsg(5) As Byte
'        System.BitConverter.GetBytes(EpicaMessageCode.eGetColonyList).CopyTo(yMsg, 0)
'        System.BitConverter.GetBytes(glPlayerID).CopyTo(yMsg, 2)
'        oUILib.SendMsgToPrimary(yMsg)
'    End Sub

'    Public Sub SetFromTradeObject(ByRef oTrade As Trade)
'        moTrade = oTrade
'        chkAccept.Enabled = True

'        If moTrade.yTradeState = Trade.eTradeStateValues.Proposal Then
'            Dim sOther As String
'            If moTrade.lPlayer1ID = glPlayerID Then sOther = GetCacheObjectValue(moTrade.lPlayer2ID, ObjectType.ePlayer) Else sOther = GetCacheObjectValue(moTrade.lPlayer1ID, ObjectType.ePlayer)
'            lblTitle.Caption = "New Trade Proposal with " & sOther
'            chkTheirAccept.Caption = sOther & "'s Acceptance"
'            Dim rcTemp As Rectangle = chkTheirAccept.GetTextDimensions()
'            chkTheirAccept.Width = rcTemp.Width + 20
'        Else
'            Dim sOther As String
'            If moTrade.lPlayer1ID = glPlayerID Then sOther = GetCacheObjectValue(moTrade.lPlayer2ID, ObjectType.ePlayer) Else sOther = GetCacheObjectValue(moTrade.lPlayer1ID, ObjectType.ePlayer)
'            lblTitle.Caption = "Trade Agreement with " & sOther
'            chkTheirAccept.Caption = sOther & "'s Acceptance"
'            Dim rcTemp As Rectangle = chkTheirAccept.GetTextDimensions()
'            chkTheirAccept.Width = rcTemp.Width + 20
'        End If

'        lstOffer.Clear()
'        lstGetting.Clear()

'        If cboSource.ListCount = 0 OrElse cboDest.ListCount = 0 Then FillSourceAndDestLists()

'        With moTrade
'            If .lPlayer1ID = glPlayerID Then
'                'cboSource.FindComboItemData(.lP1SourceID)
'                'cboDest.FindComboItemData(.lP1DestID)

'                'Fill our offer list
'                ReDim mblQtys(.mlPlayer1ItemUB)
'                For X As Int32 = 0 To .mlPlayer1ItemUB
'                    Dim sValue As String = GetOfferDescription(.muPlayer1Items(X).ObjectID, .muPlayer1Items(X).ObjTypeID, .muPlayer1Items(X).Quantity) ' GetCacheObjectValue(.muPlayer1Items(X).ObjectID, .muPlayer1Items(X).ObjTypeID)

'                    lstOffer.AddItem(sValue)
'                    lstOffer.ItemData(lstOffer.NewIndex) = .muPlayer1Items(X).ObjectID
'                    lstOffer.ItemData2(lstOffer.NewIndex) = .muPlayer1Items(X).ObjTypeID
'                    mblQtys(lstOffer.NewIndex) = .muPlayer1Items(X).Quantity
'                Next X

'                'Fill our getting list
'                For X As Int32 = 0 To .mlPlayer2ItemUB
'                    Dim sValue As String = GetTheirOfferDescription(.muPlayer2Items(X).ObjectID, .muPlayer2Items(X).ObjTypeID, .muPlayer2Items(X).Quantity)
'                    lstGetting.AddItem(sValue)
'                    lstGetting.ItemData(lstGetting.NewIndex) = .muPlayer2Items(X).ObjectID
'                    lstGetting.ItemData2(lstGetting.NewIndex) = .muPlayer2Items(X).ObjTypeID
'                Next X

'                txtNotesToMe.Caption = .sP2Notes
'                txtNotesToThem.Caption = .sP1Notes

'                If (.yTradeState And Trade.eTradeStateValues.Player1Accepted) <> 0 Then
'                    chkAccept.Value = True
'                Else : chkAccept.Value = False
'                End If
'                If (.yTradeState And Trade.eTradeStateValues.Player2Accepted) <> 0 Then
'                    chkTheirAccept.Value = True
'                Else : chkTheirAccept.Value = False
'                End If
'            Else
'                'cboSource.FindComboItemData(.lP2SourceID)
'                'cboDest.FindComboItemData(.lP2DestID)

'                'Fill our offer list
'                ReDim mblQtys(.mlPlayer2ItemUB)
'                For X As Int32 = 0 To .mlPlayer2ItemUB
'                    Dim sValue As String = GetOfferDescription(.muPlayer2Items(X).ObjectID, .muPlayer2Items(X).ObjTypeID, .muPlayer2Items(X).Quantity) 'GetCacheObjectValue(.muPlayer2Items(X).ObjectID, .muPlayer2Items(X).ObjTypeID)

'                    lstOffer.AddItem(sValue)
'                    lstOffer.ItemData(lstOffer.NewIndex) = .muPlayer2Items(X).ObjectID
'                    lstOffer.ItemData2(lstOffer.NewIndex) = .muPlayer2Items(X).ObjTypeID
'                    mblQtys(lstOffer.NewIndex) = .muPlayer2Items(X).Quantity
'                Next X

'                'Fill our getting list
'                For X As Int32 = 0 To .mlPlayer1ItemUB
'                    Dim sValue As String = GetTheirOfferDescription(.muPlayer1Items(X).ObjectID, .muPlayer1Items(X).ObjTypeID, .muPlayer1Items(X).Quantity)
'                    lstGetting.AddItem(sValue)
'                    lstGetting.ItemData(lstGetting.NewIndex) = .muPlayer1Items(X).ObjectID
'                    lstGetting.ItemData2(lstGetting.NewIndex) = .muPlayer1Items(X).ObjTypeID
'                Next X

'                txtNotesToMe.Caption = .sP1Notes
'                txtNotesToThem.Caption = .sP2Notes

'                If (.yTradeState And Trade.eTradeStateValues.Player1Accepted) <> 0 Then
'                    chkTheirAccept.Value = True
'                Else : chkTheirAccept.Value = False
'                End If
'                If (.yTradeState And Trade.eTradeStateValues.Player2Accepted) <> 0 Then
'                    chkAccept.Value = True
'                Else : chkAccept.Value = False
'                End If
'            End If


'            If chkAccept.Value = True AndAlso chkTheirAccept.Value = True Then
'                btnReject.Enabled = False
'                btnSubmit.Enabled = False
'                btnAdd.Enabled = False
'                txtQuantity.Enabled = False
'                txtNotesToMe.Enabled = False
'                txtNotesToThem.Enabled = False
'                cboSource.Enabled = False
'                cboDest.Enabled = False
'                btnRemove.Enabled = False
'                chkAccept.Enabled = False
'                chkTheirAccept.Enabled = False
'            End If
'        End With


'    End Sub

'    Public Sub CreateNewTrade(ByVal lOtherPlayerID As Int32)
'        mlOtherPlayerID = lOtherPlayerID
'        moTrade = Nothing
'        chkAccept.Enabled = False
'        lstGetting.Clear()
'        lstOffer.Clear()

'        Dim sOther As String = GetCacheObjectValue(lOtherPlayerID, ObjectType.ePlayer)
'        chkTheirAccept.Caption = sOther & "'s Acceptance"
'        lblTitle.Caption = "New Trade Proposal for " & sOther
'        Dim rcTemp As Rectangle = chkTheirAccept.GetTextDimensions()
'        chkTheirAccept.Width = rcTemp.Width + 20
'    End Sub

'    Private Function ValidateData() As Boolean
'        If cboSource.ListIndex < 0 Then
'            goUILib.AddNotification("You must select a source colony.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
'            Return False
'        End If
'        If cboDest.ListIndex < 0 Then
'            goUILib.AddNotification("You must select a destination colony.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
'            Return False
'        End If
'        Return True
'    End Function

'    Private Sub CreateAndSendSubmitTradeMessage(ByVal bReject As Boolean)
'        If bReject = False AndAlso ValidateData() = False Then Return

'        'Ok, do all the work now
'        Dim lPos As Int32 = 0
'        Dim sNotes As String = txtNotesToThem.Caption.Trim
'        Dim yMsg(24 + sNotes.Length + (lstOffer.ListCount * 14)) As Byte

'        System.BitConverter.GetBytes(EpicaMessageCode.eSubmitTrade).CopyTo(yMsg, lPos) : lPos += 2
'        If moTrade Is Nothing = False Then
'            moTrade.GetGUIDAsString.CopyTo(yMsg, lPos) : lPos += 6

'            'Trade state can only be set if the trade is initiated... and then, the player can only set THEIR Accepted/Rejected
'            If bReject = True Then
'                yMsg(lPos) = 255
'            ElseIf chkAccept.Value = True Then
'                If moTrade.lPlayer1ID = glPlayerID Then
'                    yMsg(lPos) = Trade.eTradeStateValues.Player1Accepted
'                Else : yMsg(lPos) = Trade.eTradeStateValues.Player2Accepted
'                End If
'            End If
'            lPos += 1
'        Else
'            System.BitConverter.GetBytes(-Math.Abs(mlOtherPlayerID)).CopyTo(yMsg, lPos) : lPos += 4
'            System.BitConverter.GetBytes(ObjectType.eTrade).CopyTo(yMsg, lPos) : lPos += 2
'            yMsg(lPos) = 0
'            lPos += 1
'        End If

'        'Next is my dest ID, sourceid
'        If bReject = True Then
'            System.BitConverter.GetBytes(-1I).CopyTo(yMsg, lPos) : lPos += 4
'            System.BitConverter.GetBytes(-1I).CopyTo(yMsg, lPos) : lPos += 4
'        Else
'            System.BitConverter.GetBytes(cboDest.ItemData(cboDest.ListIndex)).CopyTo(yMsg, lPos) : lPos += 4
'            System.BitConverter.GetBytes(cboSource.ItemData(cboSource.ListIndex)).CopyTo(yMsg, lPos) : lPos += 4
'        End If

'        'Now, my notes
'        System.BitConverter.GetBytes(sNotes.Length).CopyTo(yMsg, lPos) : lPos += 4
'        If sNotes.Length > 0 Then
'            System.Text.ASCIIEncoding.ASCII.GetBytes(sNotes).CopyTo(yMsg, lPos)
'            lPos += sNotes.Length
'        End If

'        'Next is the count
'        If bReject = True Then
'            System.BitConverter.GetBytes(-1I).CopyTo(yMsg, lPos) : lPos += 4
'        Else
'            System.BitConverter.GetBytes(lstOffer.ListCount).CopyTo(yMsg, lPos) : lPos += 4
'            'And now the items...
'            For X As Int32 = 0 To lstOffer.ListCount - 1
'                System.BitConverter.GetBytes(lstOffer.ItemData(X)).CopyTo(yMsg, lPos) : lPos += 4
'                System.BitConverter.GetBytes(CShort(lstOffer.ItemData2(X))).CopyTo(yMsg, lPos) : lPos += 2
'                System.BitConverter.GetBytes(mblQtys(X)).CopyTo(yMsg, lPos) : lPos += 8
'            Next X
'        End If

'        MyBase.moUILib.SendMsgToPrimary(yMsg)

'        goUILib.AddNotification("New trade proposal details submitted.", Color.Green, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
'        btnClose_Click()
'    End Sub

'    Private Sub FillSourceAndDestLists()
'        cboSource.Clear()
'        cboDest.Clear()

'        For X As Int32 = 0 To goCurrentPlayer.mlColonyUB
'            If goCurrentPlayer.mlColonyIdx(X) <> -1 Then
'                With goCurrentPlayer.moColonies(X)
'                    Dim sValue As String
'                    sValue = .sName '& " located in " & .sParentName
'                    If .ParentEnvirTypeID = ObjectType.ePlanet Then
'                        sValue &= " (Planet)"
'                    Else : sValue &= " (System)"
'                    End If
'                    cboSource.AddItem(sValue)
'                    cboSource.ItemData(cboSource.NewIndex) = .ObjectID
'                    cboDest.AddItem(sValue)
'                    cboDest.ItemData(cboDest.NewIndex) = .ObjectID
'                End With
'            End If
'        Next X

'        If moTrade Is Nothing = False Then
'            If moTrade.lPlayer1ID = glPlayerID Then
'                cboSource.FindComboItemData(moTrade.lP1SourceID)
'                cboDest.FindComboItemData(moTrade.lP1DestID)
'            Else
'                cboSource.FindComboItemData(moTrade.lP2SourceID)
'                cboDest.FindComboItemData(moTrade.lP2DestID)
'            End If
'        End If

'    End Sub

'    Private Sub btnAdd_Click(ByVal sName As String) Handles btnAdd.Click
'        If lstTradeables.ListIndex > -1 Then
'            Dim lID As Int32 = lstTradeables.ItemData(lstTradeables.ListIndex)
'            Dim iTypeID As Int16 = CShort(lstTradeables.ItemData2(lstTradeables.ListIndex))

'            'Eventually, everything will be tradeable, but for now, I need to indicate what isn't implemented
'            Select Case iTypeID
'                Case ObjectType.eAgent
'                    goUILib.AddNotification("Trading agents is not implemented yet.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
'                    Return
'                Case ObjectType.eColony
'                    goUILib.AddNotification("Trading colonies is not implemented yet.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
'                    Return
'                Case ObjectType.eFood
'                    goUILib.AddNotification("Trading food is not implemented yet.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
'                    Return
'                Case ObjectType.eStock
'                    goUILib.AddNotification("Trading stock is not implemented yet.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
'                    Return
'                Case ObjectType.eAlloyTech, ObjectType.eArmorTech, ObjectType.eEngineTech, ObjectType.eHullTech, ObjectType.ePrototype, ObjectType.eRadarTech, ObjectType.eShieldTech, ObjectType.eSpecialTech, ObjectType.eWeaponTech
'                    goUILib.AddNotification("Trading component designs is not implemented yet.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
'                    Return
'                    'Case ObjectType.eFacility
'                    '    goUILib.AddNotification("Trading facilities is not implemented yet.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
'                    '    Return
'            End Select

'            If IsNumeric(txtQuantity.Caption) = True Then
'                Dim blQty As Int64 = CLng(Val(txtQuantity.Caption))
'                If blQty > 0 Then
'                    Dim lIdx As Int32 = -1

'                    For X As Int32 = 0 To lstOffer.ListCount - 1
'                        If lstOffer.ItemData(X) = lID AndAlso lstOffer.ItemData2(X) = iTypeID Then
'                            lIdx = X
'                            Exit For
'                        End If
'                    Next X

'                    If lIdx = -1 Then
'                        lstOffer.AddItem(GetOfferDescription(lID, iTypeID, blQty))
'                        lIdx = lstOffer.NewIndex
'                        lstOffer.ItemData(lIdx) = lID
'                        lstOffer.ItemData2(lIdx) = iTypeID
'                        ReDim Preserve mblQtys(lstOffer.ListCount - 1)
'                        mblQtys(lIdx) = 0
'                    Else
'                        'Ensure we have enough qty
'                        If ItemStackable(CShort(lstTradeables.ItemData2(lstTradeables.ListIndex))) = False Then
'                            goUILib.AddNotification("That item is already in the offer list.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
'                            Return
'                        End If
'                    End If

'                    mblQtys(lIdx) += blQty
'                    Dim sValue As String = GetOfferDescription(lID, iTypeID, mblQtys(lIdx))
'                    If lstOffer.List(lIdx) <> sValue Then lstOffer.List(lIdx) = sValue
'                Else : goUILib.AddNotification("Please enter a numeric value greater than 0 and try again.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
'                End If
'            Else : goUILib.AddNotification("Please enter a numeric value greater than 0 and try again.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
'            End If
'        Else : goUILib.AddNotification("Please select an item in the tradeables list and try again.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
'        End If
'    End Sub

'    Private Sub btnClose_Click(ByVal sName As String) Handles btnClose.Click
'        MyBase.moUILib.RemoveWindow(Me.ControlName)
'    End Sub

'    Private Sub btnReject_Click(ByVal sName As String) Handles btnReject.Click
'        If moTrade Is Nothing Then
'            MyBase.moUILib.RemoveWindow(Me.ControlName)
'            Return
'        End If

'        If btnReject.Caption.ToLower = "reject" Then
'            btnReject.Caption = "Confirm"
'            goUILib.AddNotification("Press Confirm to Reject the Trade. This action cannot be undone.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
'        Else
'            'reject the trade
'            CreateAndSendSubmitTradeMessage(True)
'        End If
'    End Sub

'    Private Sub btnRemove_Click(ByVal sName As String) Handles btnRemove.Click
'        If lstOffer.ListIndex > -1 Then
'            For X As Int32 = lstOffer.ListIndex To mblQtys.GetUpperBound(0) - 1
'                mblQtys(X) = mblQtys(X + 1)
'            Next X
'            ReDim Preserve mblQtys(mblQtys.GetUpperBound(0) - 1)
'            lstOffer.RemoveItem(lstOffer.ListIndex)
'        Else
'            goUILib.AddNotification("Please select an item in the list of items you are offering to remove and try again.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
'        End If
'    End Sub

'    Private Sub btnSubmit_Click(ByVal sName As String) Handles btnSubmit.Click
'        CreateAndSendSubmitTradeMessage(False)
'    End Sub

'    Private Sub cboSource_ItemSelected(ByVal lItemIndex As Integer) Handles cboSource.ItemSelected
'        'If cboSource.ListIndex > -1 Then
'        '    Dim yMsg(5) As Byte
'        '    System.BitConverter.GetBytes(EpicaMessageCode.eGetColonyTradeables).CopyTo(yMsg, 0)
'        '    System.BitConverter.GetBytes(cboSource.ItemData(cboSource.ListIndex)).CopyTo(yMsg, 2)
'        '    MyBase.moUILib.SendMsgToPrimary(yMsg)
'        'End If
'    End Sub

'    Private Sub lstGetting_ItemClick(ByVal lIndex As Integer) Handles lstGetting.ItemClick
'        With lstGetting
'            If .ListIndex > -1 Then
'                ShowItemDetails(.ItemData(.ListIndex), CShort(.ItemData2(.ListIndex)))
'            End If
'        End With
'    End Sub

'    Private Sub lstOffer_ItemClick(ByVal lIndex As Integer) Handles lstOffer.ItemClick
'        With lstOffer
'            If .ListIndex > -1 Then
'                ShowItemDetails(.ItemData(.ListIndex), CShort(.ItemData2(.ListIndex)))
'            End If
'        End With
'    End Sub

'    Private Sub lstTradables_ItemClick(ByVal lIndex As Integer) Handles lstTradeables.ItemClick
'        With lstTradeables
'            If .ListIndex > -1 Then
'                ShowItemDetails(.ItemData(.ListIndex), CShort(.ItemData2(.ListIndex)))
'                If .ItemData2(.ListIndex) = ObjectType.eUnit OrElse .ItemData2(.ListIndex) = ObjectType.eFacility Then
'                    txtQuantity.Enabled = False
'                    txtQuantity.Caption = "1"
'                Else
'                    txtQuantity.Enabled = True
'                End If
'            End If
'        End With
'    End Sub

'    Private Sub lstTradables_ItemDblClick(ByVal lIndex As Integer) Handles lstTradeables.ItemDblClick
'        btnAdd_Click()
'    End Sub

'    Private Function GetOfferDescription(ByVal lObjID As Int32, ByVal iObjTypeID As Int16, ByVal blQty As Int64) As String
'        Dim sValue As String = ""
'        Select Case Math.Abs(iObjTypeID)
'            Case ObjectType.eColony
'                sValue = "Entire Colony"
'                If cboSource.ListIndex > -1 Then
'                    Dim lID As Int32 = cboSource.ItemData(cboSource.ListIndex)
'                    For X As Int32 = 0 To goCurrentPlayer.mlColonyUB
'                        If goCurrentPlayer.mlColonyIdx(X) = lID Then
'                            sValue = goCurrentPlayer.moColonies(X).sName & " (Entire Colony)"
'                            Exit For
'                        End If
'                    Next X
'                End If
'            Case ObjectType.eColonists
'                sValue = "Colonists" ' (" & lQty & ")"
'            Case ObjectType.eEnlisted
'                sValue = "Enlisted" ' (" & lQty & ")"
'            Case ObjectType.eOfficers
'                sValue = "Officers" ' (" & lQty & ")"
'            Case ObjectType.eAgent
'                'TODO: Define this
'            Case ObjectType.eCredits
'                sValue = "Credits" ' (" & lQty.ToString("#,###") & ")"
'            Case ObjectType.eFood
'                sValue = "Food" ' (" & lQty & ")"
'            Case ObjectType.eAmmunition
'                Dim oTech As Epica_Tech = goCurrentPlayer.GetTech(lObjID, ObjectType.eWeaponTech)
'                Dim sTemp As String = ""
'                If oTech Is Nothing = False Then
'                    sTemp = oTech.GetComponentName
'                Else : sTemp = GetCacheObjectValue(lObjID, ObjectType.eWeaponTech)
'                End If
'                sValue = "Ammunition (" & sTemp & ")"
'            Case ObjectType.eStock
'                'TODO: Define this
'            Case ObjectType.eAlloyTech, ObjectType.eArmorTech, ObjectType.eEngineTech, ObjectType.eHullTech, ObjectType.ePrototype, ObjectType.eRadarTech, ObjectType.eShieldTech, ObjectType.eWeaponTech
'                Dim oTech As Epica_Tech = goCurrentPlayer.GetTech(lObjID, Math.Abs(iObjTypeID))
'                If oTech Is Nothing = False Then
'                    sValue = oTech.GetComponentName & " (" & Epica_Tech.GetComponentTypeName(Math.Abs(iObjTypeID)) & ")"
'                Else
'                    If iObjTypeID < 0 Then
'                        sValue = "Unknown Components"
'                    Else : sValue = "Unknown " & Epica_Tech.GetComponentTypeName(iObjTypeID) & " Design"
'                    End If
'                End If
'            Case ObjectType.eMineral, ObjectType.eMineralCache
'                sValue = "Unknown Minerals"
'                For X As Int32 = 0 To glMineralUB
'                    If glMineralIdx(X) = lObjID Then
'                        sValue = goMinerals(X).MineralName
'                        Exit For
'                    End If
'                Next X
'                'sValue &= " (" & lQty & ")"
'            Case Else
'                sValue = GetCacheObjectValue(lObjID, iObjTypeID)
'        End Select

'        If blQty > 1 Then sValue &= " (" & blQty.ToString("#,###") & ")"
'        Return sValue
'    End Function

'    Private Function GetTheirOfferDescription(ByVal lObjID As Int32, ByVal iObjTypeID As Int16, ByVal blQty As Int64) As String
'        Dim sValue As String = ""
'        Select Case Math.Abs(iObjTypeID)
'            Case ObjectType.eColony
'                sValue = "Entire Colony"
'                sValue = GetCacheObjectValue(lObjID, iObjTypeID)
'            Case ObjectType.eColonists
'                sValue = "Colonists" ' (" & lQty & ")"
'            Case ObjectType.eEnlisted
'                sValue = "Enlisted" ' (" & lQty & ")"
'            Case ObjectType.eOfficers
'                sValue = "Officers" ' (" & lQty & ")"
'            Case ObjectType.eAgent
'                'TODO: Define this
'            Case ObjectType.eCredits
'                sValue = "Credits" ' (" & lQty.ToString("#,###") & ")"
'            Case ObjectType.eFood
'                sValue = "Food" ' (" & lQty & ")"
'            Case ObjectType.eAmmunition
'                Dim oTech As Epica_Tech = goCurrentPlayer.GetTech(lObjID, ObjectType.eWeaponTech)
'                Dim sTemp As String = ""
'                If oTech Is Nothing = False Then
'                    sTemp = oTech.GetComponentName
'                Else : sTemp = GetCacheObjectValue(lObjID, ObjectType.eWeaponTech)
'                End If
'                sValue = "Ammunition (" & sTemp & ")"
'            Case ObjectType.eStock
'                'TODO: Define this
'            Case ObjectType.eAlloyTech, ObjectType.eArmorTech, ObjectType.eEngineTech, ObjectType.eHullTech, ObjectType.ePrototype, ObjectType.eRadarTech, ObjectType.eShieldTech, ObjectType.eWeaponTech
'                sValue = GetCacheObjectValue(lObjID, Math.Abs(iObjTypeID))
'            Case ObjectType.eMineral, ObjectType.eMineralCache
'                sValue = "Unknown Minerals"
'                For X As Int32 = 0 To glMineralUB
'                    If glMineralIdx(X) = lObjID Then
'                        sValue = goMinerals(X).MineralName
'                        Exit For
'                    End If
'                Next X
'                'sValue &= " (" & lQty & ")"
'            Case Else
'                sValue = GetCacheObjectValue(lObjID, iObjTypeID)
'        End Select

'        If blQty > 1 Then sValue &= " (" & blQty.ToString("#,###") & ")"
'        Return sValue
'    End Function

'    Private Sub EnsureTradeablesContains(ByVal lObjID As Int32, ByVal iObjTypeID As Int16, ByVal blQty As Int64)
'        Dim sValue As String

'        Dim blUsedQty As Int64 = 0

'        For X As Int32 = 0 To lstOffer.ListCount - 1
'            If lstOffer.ItemData(X) = lObjID AndAlso lstOffer.ItemData2(X) = iObjTypeID Then
'                blUsedQty = mblQtys(X)
'                Exit For
'            End If
'        Next X

'        Select Case Math.Abs(iObjTypeID)
'            Case ObjectType.eColony
'                sValue = "Entire Colony"
'                If cboSource.ListIndex > -1 Then
'                    Dim lID As Int32 = cboSource.ItemData(cboSource.ListIndex)
'                    For X As Int32 = 0 To goCurrentPlayer.mlColonyUB
'                        If goCurrentPlayer.mlColonyIdx(X) = lID Then
'                            sValue = goCurrentPlayer.moColonies(X).sName & " (Entire Colony)"
'                            Exit For
'                        End If
'                    Next X
'                End If
'            Case ObjectType.eColonists
'                sValue = "Colonists (" & (blQty - blUsedQty) & ")"
'            Case ObjectType.eEnlisted
'                sValue = "Enlisted (" & (blQty - blUsedQty) & ")"
'            Case ObjectType.eOfficers
'                sValue = "Officers (" & (blQty - blUsedQty) & ")"
'            Case ObjectType.eAgent
'                sValue = "Unknown Agent"
'                sValue = GetCacheObjectValue(lObjID, iObjTypeID)
'            Case ObjectType.eCredits
'                sValue = "Credits (" & (goCurrentPlayer.blCredits - blUsedQty).ToString("#,###") & ")"
'            Case ObjectType.eFood
'                sValue = "Food (" & (blQty - blUsedQty) & ")"
'            Case ObjectType.eAmmunition
'                sValue = "Ammunition (" & (blQty - blUsedQty) & ")"
'            Case ObjectType.eStock
'                sValue = "Unknown Stock"
'                sValue = GetCacheObjectValue(lObjID, iObjTypeID)
'            Case ObjectType.eAlloyTech, ObjectType.eArmorTech, ObjectType.eEngineTech, ObjectType.eHullTech, ObjectType.ePrototype, ObjectType.eRadarTech, ObjectType.eShieldTech, ObjectType.eWeaponTech
'                If iObjTypeID < 0 Then
'                    'represents a component cache
'                    sValue = "Unknown Components"
'                    sValue = GetCacheObjectValue(lObjID, Math.Abs(iObjTypeID)) & " (" & (blQty - blUsedQty) & ")"
'                Else
'                    'represents the tech design
'                    sValue = "Unknown " & Epica_Tech.GetComponentTypeName(iObjTypeID) & " Design"
'                    For X As Int32 = 0 To goCurrentPlayer.mlTechUB
'                        If goCurrentPlayer.moTechs(X).ObjectID = lObjID AndAlso goCurrentPlayer.moTechs(X).ObjTypeID = iObjTypeID Then
'                            sValue = goCurrentPlayer.moTechs(X).GetComponentName() & " (" & Epica_Tech.GetComponentTypeName(iObjTypeID) & " Design)"
'                            Exit For
'                        End If
'                    Next X
'                End If
'            Case ObjectType.eMineral, ObjectType.eMineralCache
'                'represents a mineral cache... ID should be mineral
'                sValue = "Unknown Minerals"
'                For X As Int32 = 0 To glMineralUB
'                    If glMineralIdx(X) = lObjID Then
'                        sValue = goMinerals(X).MineralName
'                        Exit For
'                    End If
'                Next X
'                sValue &= " (" & (blQty - blUsedQty) & ")"
'            Case Else
'                sValue = GetCacheObjectValue(lObjID, iObjTypeID)
'        End Select


'        For X As Int32 = 0 To lstTradeables.ListCount - 1
'            If lstTradeables.ItemData(X) = lObjID AndAlso lstTradeables.ItemData2(X) = iObjTypeID Then
'                If lstTradeables.List(X) <> sValue Then lstTradeables.List(X) = sValue
'                Return
'            End If
'        Next X

'        'If we are here, it does not contain it...
'        lstTradeables.AddItem(sValue)
'        lstTradeables.ItemData(lstTradeables.NewIndex) = lObjID
'        lstTradeables.ItemData2(lstTradeables.NewIndex) = iObjTypeID
'    End Sub

'    Private Sub frmTrade_OnNewFrame() Handles Me.OnNewFrame
'        If glCurrentCycle - mlLastUpdate > 60 Then      '2 seconds
'            mlLastUpdate = glCurrentCycle

'            If moTrade Is Nothing = False Then
'                If moTrade.lPlayer1ID = glPlayerID Then
'                    If txtNotesToMe.Caption <> moTrade.sP2Notes Then txtNotesToMe.Caption = moTrade.sP2Notes
'                Else
'                    If txtNotesToMe.Caption <> moTrade.sP1Notes Then txtNotesToMe.Caption = moTrade.sP1Notes
'                End If
'            End If

'            'Check Tradeables listbox (including qtys)
'            If cboSource.ListIndex > -1 Then
'                Dim oColony As Colony = Nothing
'                Dim lID As Int32 = cboSource.ItemData(cboSource.ListIndex)
'                For X As Int32 = 0 To goCurrentPlayer.mlColonyUB
'                    If goCurrentPlayer.mlColonyIdx(X) = lID Then
'                        oColony = goCurrentPlayer.moColonies(X)
'                        Exit For
'                    End If
'                Next X
'                If oColony Is Nothing = False Then
'                    EnsureTradeablesContains(oColony.ObjectID, oColony.ObjTypeID, 1)
'                    EnsureTradeablesContains(1, ObjectType.eCredits, 0)     'pass 0 to handle the int64 internally
'                    EnsureTradeablesContains(1, ObjectType.eColonists, oColony.Colonists)
'                    EnsureTradeablesContains(1, ObjectType.eEnlisted, oColony.Enlisted)
'                    EnsureTradeablesContains(1, ObjectType.eOfficers, oColony.Officers)
'                    'TODO: Agents are next
'                    EnsureTradeablesContains(1, ObjectType.eFood, oColony.Food)
'                    EnsureTradeablesContains(1, ObjectType.eAmmunition, oColony.Ammo)
'                    'TODO: Stocks go here

'                    'Tech Designs
'                    For X As Int32 = 0 To goCurrentPlayer.mlTechUB
'                        If goCurrentPlayer.moTechs(X).ObjTypeID <> ObjectType.eSpecialTech Then
'                            EnsureTradeablesContains(goCurrentPlayer.moTechs(X).ObjectID, goCurrentPlayer.moTechs(X).ObjTypeID, 1)
'                        End If
'                    Next X

'                    ''Component(Caches)
'                    'For X As Int32 = 0 To oColony.lCacheUB
'                    '    If oColony.lCacheIDs(X) <> -1 Then
'                    '        If oColony.iCacheTypeIDs(X) <> ObjectType.eMineralCache AndAlso oColony.iCacheTypeIDs(X) <> ObjectType.eMineral Then
'                    '            'Ok, must be component cache
'                    '            EnsureTradeablesContains(oColony.lCacheIDs(X), -Math.Abs(oColony.iCacheTypeIDs(X)), oColony.lCacheQtys(X))
'                    '        End If
'                    '    End If
'                    'Next X
'                    ''Mineral(Caches)
'                    'For X As Int32 = 0 To oColony.lCacheUB
'                    '    If oColony.lCacheIDs(X) <> -1 Then
'                    '        If oColony.iCacheTypeIDs(X) = ObjectType.eMineralCache OrElse oColony.iCacheTypeIDs(X) = ObjectType.eMineral Then
'                    '            'Ok, must be component cache
'                    '            EnsureTradeablesContains(oColony.lCacheIDs(X), oColony.iCacheTypeIDs(X), oColony.lCacheQtys(X))
'                    '        End If
'                    '    End If
'                    'Next X

'                    ''Facilities()
'                    'For X As Int32 = 0 To oColony.lFacUB
'                    '    If oColony.lFacIDs(X) <> -1 Then
'                    '        EnsureTradeablesContains(oColony.lFacIDs(X), ObjectType.eFacility, 1)
'                    '    End If
'                    'Next X
'                    ''Units()
'                    'For X As Int32 = 0 To oColony.lUnitUB
'                    '    If oColony.lUnitIDs(X) <> -1 Then
'                    '        EnsureTradeablesContains(oColony.lUnitIDs(X), ObjectType.eUnit, 1)
'                    '    End If
'                    'Next X
'                Else
'                    lstTradeables.Clear()
'                    lstOffer.Clear()
'                End If
'            ElseIf lstTradeables.ListCount > 0 Then
'                lstTradeables.Clear()
'                lstOffer.Clear()
'            End If

'            'Check Colonies cbos
'            Dim bFound(cboSource.ListCount - 1) As Boolean
'            For X As Int32 = 0 To goCurrentPlayer.mlColonyUB
'                If goCurrentPlayer.mlColonyIdx(X) <> -1 Then
'                    Dim lID As Int32 = goCurrentPlayer.moColonies(X).ObjectID
'                    Dim lIdx As Int32 = -1
'                    For Y As Int32 = 0 To cboSource.ListCount - 1
'                        If cboSource.ItemData(Y) = lID Then
'                            bFound(Y) = True
'                            lIdx = Y
'                            Exit For
'                        End If
'                    Next Y

'                    Dim sValue As String
'                    With goCurrentPlayer.moColonies(X)
'                        sValue = .sName '& " located in " & .sParentName
'                        If .ParentEnvirTypeID = ObjectType.ePlanet Then
'                            sValue &= " (Planet)"
'                        Else : sValue &= " (System)"
'                        End If
'                    End With

'                    If lIdx = -1 Then
'                        cboSource.AddItem(sValue)
'                        cboSource.ItemData(cboSource.NewIndex) = goCurrentPlayer.moColonies(X).ObjectID
'                        cboDest.AddItem(sValue)
'                        cboDest.ItemData(cboDest.NewIndex) = goCurrentPlayer.moColonies(X).ObjectID
'                    Else
'                        If cboSource.List(lIdx) <> sValue Then cboSource.List(lIdx) = sValue
'                        If cboDest.List(lIdx) <> sValue Then cboDest.List(lIdx) = sValue
'                    End If

'                End If
'            Next X
'            For X As Int32 = 0 To bFound.GetUpperBound(0)
'                If bFound(X) = False Then
'                    cboSource.RemoveItem(X)
'                    cboDest.RemoveItem(X)
'                End If
'            Next X

'            'Check lstGetting
'            If moTrade Is Nothing = False Then
'                With moTrade
'                    If .lPlayer1ID = glPlayerID Then
'                        For X As Int32 = 0 To .mlPlayer2ItemUB
'                            Dim lID As Int32 = .muPlayer2Items(X).ObjectID
'                            Dim iTypeID As Int16 = .muPlayer2Items(X).ObjTypeID
'                            Dim blQty As Int64 = .muPlayer2Items(X).Quantity

'                            'Dim sValue As String =  GetCacheObjectValue(lID, iTypeID)
'                            'If lQty > 1 Then sValue &= " (" & lQty & ")"
'                            Dim sValue As String = GetTheirOfferDescription(lID, iTypeID, blQty)

'                            Dim bGettingFound As Boolean = False
'                            For Y As Int32 = 0 To lstGetting.ListCount - 1
'                                If lstGetting.ItemData(Y) = lID AndAlso lstGetting.ItemData2(Y) = iTypeID Then
'                                    If lstGetting.List(Y) <> sValue Then lstGetting.List(Y) = sValue
'                                    bGettingFound = True
'                                    Exit For
'                                End If
'                            Next Y
'                            If bGettingFound = False Then
'                                lstGetting.AddItem(sValue)
'                                lstGetting.ItemData(lstGetting.NewIndex) = lID
'                                lstGetting.ItemData2(lstGetting.NewIndex) = iTypeID
'                            End If
'                        Next X
'                    Else
'                        For X As Int32 = 0 To .mlPlayer1ItemUB
'                            Dim lID As Int32 = .muPlayer1Items(X).ObjectID
'                            Dim iTypeID As Int16 = .muPlayer1Items(X).ObjTypeID
'                            Dim blQty As Int64 = .muPlayer1Items(X).Quantity

'                            'Dim sValue As String = GetCacheObjectValue(lID, iTypeID)
'                            'If lQty > 1 Then sValue &= " (" & lQty & ")"
'                            Dim sValue As String = GetTheirOfferDescription(lID, iTypeID, blQty)

'                            Dim bGettingFound As Boolean = False
'                            For Y As Int32 = 0 To lstGetting.ListCount - 1
'                                If lstGetting.ItemData(Y) = lID AndAlso lstGetting.ItemData2(Y) = iTypeID Then
'                                    If lstGetting.List(Y) <> sValue Then lstGetting.List(Y) = sValue
'                                    bGettingFound = True
'                                    Exit For
'                                End If
'                            Next Y
'                            If bGettingFound = False Then
'                                lstGetting.AddItem(sValue)
'                                lstGetting.ItemData(lstGetting.NewIndex) = lID
'                                lstGetting.ItemData2(lstGetting.NewIndex) = iTypeID
'                            End If
'                        Next X
'                    End If
'                End With
'            End If
'        End If
'    End Sub

'    Private Sub ShowItemDetails(ByVal lObjectID As Int32, ByVal iObjTypeID As Int16)
'        Select Case iObjTypeID
'            Case ObjectType.eColonists, ObjectType.eEnlisted, ObjectType.eOfficers, ObjectType.eCredits, ObjectType.eFood, ObjectType.eAmmunition, ObjectType.eColony
'                'No details to display, so remove the window
'                MyBase.moUILib.RemoveWindow("frmNonOwnerItemDetails")
'                MyBase.moUILib.RemoveWindow("frmMinDetail")
'            Case ObjectType.eMineral, ObjectType.eMineralCache
'                'If the player knows of the mineral, display the mineral details
'                For X As Int32 = 0 To glMineralUB
'                    If glMineralIdx(X) = lObjectID Then
'                        If goMinerals(X).bDiscovered = True Then
'                            Dim ofrmMinDet As New frmMinDetail(goUILib)
'                            ofrmMinDet.ShowMineralDetail(MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth \ 2 - 200, Me.Top, 500, lObjectID)
'                            ofrmMinDet = Nothing
'                        End If
'                        Exit For
'                    End If
'                Next X
'            Case Else
'                'Everything else calls the window
'                Dim ofrmDetails As New frmNonOwnerItemDetails(goUILib)
'                ofrmDetails.SetFromItem(lObjectID, Math.Abs(iObjTypeID))
'                ofrmDetails = Nothing
'        End Select

'    End Sub

'    Private Sub chkTheirAccept_Click() Handles chkTheirAccept.Click
'        If moTrade Is Nothing = False Then
'            If moTrade.lPlayer1ID = glPlayerID Then
'                chkTheirAccept.Value = (moTrade.yTradeState And Trade.eTradeStateValues.Player2Accepted) <> 0
'            Else : chkTheirAccept.Value = (moTrade.yTradeState And Trade.eTradeStateValues.Player1Accepted) <> 0
'            End If
'        Else : chkTheirAccept.Value = False
'        End If
'    End Sub

'    Private Function ItemStackable(ByVal iTypeID As Int16) As Boolean
'        Select Case iTypeID
'            Case ObjectType.eAgent, ObjectType.eColony, ObjectType.eAlloyTech, ObjectType.eArmorTech, ObjectType.eEngineTech, _
'              ObjectType.eHullTech, ObjectType.ePrototype, ObjectType.eRadarTech, ObjectType.eShieldTech, ObjectType.eWeaponTech, _
'              ObjectType.eUnit, ObjectType.eFacility
'                Return False
'            Case Else
'                Return True
'        End Select
'    End Function
'End Class