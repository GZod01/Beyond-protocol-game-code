Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D
Imports System.Xml

'Interface created from Interface Builder
Public Class frmTutorialTOC
    Inherits UIWindow

    Private lblTitle As UILabel
    Private WithEvents tvwTOC As UITreeView
    Private WithEvents btnDisplay As UIButton
    Private WithEvents btnClose As UIButton
    Private WithEvents btnSearch As UIButton
    Private WithEvents txtSearch As UITextBox
    Private WithEvents optSubject As UIOption
    Private WithEvents optBody As UIOption


    Private mlFoundCategory As Int32 = -1
    Private mlFoundIdx As Int32 = -1
    Private mlFoundPage As Int32 = -1

    Private mbLoading As Boolean = True
    Private moTOCList As TOCList = Nothing

    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)

        'frmTutorialTOC initial props
        With Me
            .lWindowMetricID = BPMetrics.eWindow.eTutorialTOC
            .ControlName = "frmTutorialTOC"

            Dim lLeft As Int32 = -1 'muSettings.TutorialTOCX
            Dim lTop As Int32 = -1 'muSettings.TutorialTOCY

            If NewTutorialManager.TutorialOn = False Then
                lLeft = muSettings.TutorialTOCX
                lTop = muSettings.TutorialTOCY
            End If
            If lLeft < 0 Then lLeft = MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth \ 2 - 256
            If lTop < 0 Then lTop = MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight \ 2 - 256
            If lLeft + 511 > MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth Then lLeft = MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth - 512
            If lTop + 511 > MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight Then lTop = MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight - 512

            .Left = lLeft '275
            .Top = lTop '188
            .Width = 511
            .Height = 511
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = True
            .BorderLineWidth = 1
            .Moveable = True
        End With

        'lblTitle initial props
        lblTitle = New UILabel(oUILib)
        With lblTitle
            .ControlName = "lblTitle"
            .Left = 5
            .Top = 5
            .Width = 218
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Quick Help Table of Contents"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblTitle, UIControl))

        'btnClose initial props
        btnClose = New UIButton(oUILib)
        With btnClose
            .ControlName = "btnClose"
            .Width = 100
            .Height = 25
            .Left = Me.Width - .Width
            .Top = Me.Height - .Height - 1
            .Enabled = True
            .Visible = True
            .Caption = "Close"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnClose, UIControl))

        'btnDisplay initial props
        btnDisplay = New UIButton(oUILib)
        With btnDisplay
            .ControlName = "btnDisplay"
            .Width = 100
            .Height = 25
            .Left = btnClose.Left - .Width - 1 '125 '5
            .Top = Me.Height - .Height - 1 '480 '230
            .Enabled = False
            .Visible = True
            .Caption = "Display"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnDisplay, UIControl))

        'btnSearch initial props
        btnSearch = New UIButton(oUILib)
        With btnSearch
            .ControlName = "btnSearch"
            .Width = 100
            .Height = 25
            .Left = btnDisplay.Left - .Width - 1
            .Top = Me.Height - .Height - 1
            .Enabled = True
            .Visible = True
            .Caption = "Search"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnSearch, UIControl))

        'optHeader initial props
        txtSearch = New UITextBox(oUILib)
        With txtSearch
            .ControlName = "txtSearch"
            .Left = 5
            .Top = btnSearch.Top + 1
            .Width = btnSearch.Left - 10
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = ""
            .ForeColor = muSettings.InterfaceTextBoxForeColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
            .BackColorEnabled = muSettings.InterfaceTextBoxFillColor
            .BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
            .MaxLength = 50
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(txtSearch, UIControl))

        'optSubject initial props
        optSubject = New UIOption(oUILib)
        With optSubject
            .ControlName = "optSubject"
            .Width = 70
            .Height = 18
            .Left = txtSearch.Left
            .Top = txtSearch.Top - .Height - 1
            .Enabled = True
            .Visible = True
            .Caption = "Subject"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = True
        End With
        Me.AddChild(CType(optSubject, UIControl))

        'optBody initial props
        optBody = New UIOption(oUILib)
        With optBody
            .ControlName = "optBody"
            .Width = 60
            .Height = 18
            .Left = optSubject.Left + optSubject.Width + 25
            .Top = optSubject.Top
            .Enabled = True
            .Visible = True
            .Caption = "Body"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = False
        End With
        Me.AddChild(CType(optBody, UIControl))

        'tvwTOC initial props
        tvwTOC = New UITreeView(oUILib)
        With tvwTOC
            .ControlName = "tvwTOC"
            .Left = 5
            .Top = 25
            .Width = 500
            .Height = optSubject.Top - .Top - 5
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .ForeColor = muSettings.InterfaceBorderColor
            .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
        End With
        Me.AddChild(CType(tvwTOC, UIControl))

        FillTOCList()

        mbLoading = False

        MyBase.moUILib.RemoveWindow(Me.ControlName)
        MyBase.moUILib.AddWindow(Me)
    End Sub

    Private Sub FillTOCList()
        If goTutorial Is Nothing Then goTutorial = New TutorialManager()

        'lstTOC.Clear()
        If moTOCList Is Nothing Then moTOCList = New TOCList()
        moTOCList.FillTreeView(tvwTOC)

        'Dim sItems() As String = Nothing
        'Dim lItemID() As Int32 = Nothing
        'Dim lItemUB As Int32 = -1

        'For X As Int32 = 0 To goTutorial.mlStepUB
        '    With goTutorial
        '        If .muSteps(X).sStepTitle <> "" Then

        '            lItemUB += 1
        '            ReDim Preserve sItems(lItemUB)
        '            ReDim Preserve lItemID(lItemUB)

        '            Dim bFound As Boolean = False
        '            For Y As Int32 = 0 To lItemUB
        '                If sItems(Y) > .muSteps(X).sStepTitle Then
        '                    For Z As Int32 = lItemUB To Y + 1 Step -1
        '                        sItems(Z) = sItems(Z - 1)
        '                        lItemID(Z) = lItemID(Z - 1)
        '                    Next Z

        '                    sItems(Y) = .muSteps(X).sStepTitle
        '                    lItemID(Y) = .muSteps(X).lStepID
        '                    bFound = True
        '                    Exit For
        '                End If
        '            Next Y

        '            If bFound = False Then
        '                sItems(lItemUB) = .muSteps(X).sStepTitle
        '                lItemID(lItemUB) = .muSteps(X).lStepID
        '            End If
        '        End If
        '    End With
        'Next X

        'For X As Int32 = 0 To lItemUB
        '    lstTOC.AddItem(sItems(X))
        '    lstTOC.ItemData(lstTOC.NewIndex) = lItemID(X)
        'Next X

        'If goTutorial Is Nothing = False Then chkTutorial.Value = goTutorial.TutorialOn Else chkTutorial.Value = False
    End Sub

    Private Sub btnDisplay_Click(ByVal sName As String) Handles btnDisplay.Click
        Dim oNode As UITreeView.UITreeViewItem = tvwTOC.oSelectedNode
        If oNode Is Nothing = False Then
            BPMetrics.MetricMgr.AddActivity(BPMetrics.eInterface.eHelpWindowItemClick, oNode.lItemData, oNode.lItemData2)

            'Now, do it
            If oNode.lItemData2 <> TOCList.lItemType.eTopic Then Return
            If oNode.lItemData2 = TOCList.lItemType.eTopic AndAlso oNode.oRelatedObject Is Nothing = False Then
                Dim oFrm As frmHelpItem = CType(MyBase.moUILib.GetWindow("frmHelpItem"), frmHelpItem)
                If oFrm Is Nothing Then oFrm = New frmHelpItem(goUILib)
                oFrm.Top = Me.Top
                oFrm.Left = Me.Left + Me.Width
                oFrm.Visible = True
                oFrm.SetFromTopicNode(CType(oNode.oRelatedObject, TOCList.TOCItem))
            Else
                MyBase.moUILib.AddNotification("Select a topic to display first.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            End If
        End If
        'If lstTOC.ListIndex > -1 Then
        '	Dim lID As Int32 = lstTOC.ItemData(lstTOC.ListIndex)
        '	If goTutorial.TutorialOn = False Then
        '		For X As Int32 = 0 To goTutorial.mlStepUB
        '			If goTutorial.muSteps(X).lStepID = lID Then
        '				goTutorial.lTemporaryOnGroupID = goTutorial.muSteps(X).GroupID
        '				Exit For
        '			End If
        '		Next X
        '		goTutorial.TutorialOn = True
        '	End If
        '	goTutorial.ResetGroupOfStep(lID)
        '	goTutorial.ExecuteTutorialStep(lid)
        'End If
    End Sub

    Private Sub btnClose_Click(ByVal sName As String) Handles btnClose.Click
        goUILib.RemoveWindow("frmHelpItem")
        MyBase.moUILib.RemoveWindow(Me.ControlName)
    End Sub

    'Private Sub chkTutorial_Click() Handles chkTutorial.Click
    '    If mbLoading = True Then Return
    '    If goTutorial Is Nothing Then
    '        If chkTutorial.Value = True Then
    '            goTutorial = New TutorialManager()
    '            goTutorial.TutorialOn = True
    '        End If
    '    Else : goTutorial.TutorialOn = chkTutorial.Value
    '    End If

    '    If goTutorial.TutorialOn = True Then goTutorial.ContinueTutorial()
    'End Sub

    Private Sub tvwTOC_NodeDoubleClicked() Handles tvwTOC.NodeDoubleClicked
        btnDisplay_Click("btnDisplay")
    End Sub

    Private Sub tvwTOC_NodeExpanded(ByVal oNode As UITreeView.UITreeViewItem) Handles tvwTOC.NodeExpanded
        BPMetrics.MetricMgr.AddActivity(BPMetrics.eInterface.eHelpWindowTopicExpand, oNode.lItemData, oNode.lItemData2)
    End Sub

    Private Sub frmTutorialTOC_WindowMoved() Handles Me.WindowMoved
        If mbLoading = False Then
            muSettings.TutorialTOCX = Me.Left
            muSettings.TutorialTOCY = Me.Top
        End If
        Dim ofrm As UIWindow = MyBase.moUILib.GetWindow("frmHelpItem")

        If ofrm Is Nothing = False Then
            ofrm.Top = Me.Top
            ofrm.Left = Me.Left + Me.Width
        End If
    End Sub

    Private Sub tvwTOC_NodeSelected(ByRef oNode As UITreeView.UITreeViewItem) Handles tvwTOC.NodeSelected
        btnDisplay.Enabled = oNode.lItemData2 = TOCList.lItemType.eTopic
    End Sub

    Private Sub txtSearch_OnKeyPress(ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtSearch.OnKeyPress
        If mlFoundIdx > 0 Then
            mlFoundCategory = -1
            mlFoundIdx = 0
            mlFoundPage = 0
        End If
    End Sub

    Private Sub btnSearch_Click(ByVal sName As String) Handles btnSearch.Click
        Dim sFind As String = txtSearch.Caption.Trim
        If sFind.Length = 0 Then Return
        If moTOCList Is Nothing = False Then
            'Collapse all
            For X As Int32 = 0 To moTOCList.oDocs(1).oItems.GetUpperBound(0)
                If moTOCList.oDocs(1).oItems(X).oTVWItem.bExpanded = True Then
                    moTOCList.oDocs(1).oItems(X).oTVWItem.bExpanded = False
                End If
            Next
            'Find it
            If optSubject.Value = True Then
                FindDoc(txtSearch.Caption.ToUpper, mlFoundIdx, False)
            Else
                FindDoc(txtSearch.Caption.ToUpper, mlFoundIdx, True)
            End If
        End If
    End Sub

    Private Sub FindDoc(ByVal sSearchFor As String, ByVal lId As Int32, ByVal bSearchBody As Boolean)
        If moTOCList Is Nothing = False AndAlso moTOCList.oDocs Is Nothing = False Then
            For X As Int32 = 0 To moTOCList.oDocs(1).oItems.GetUpperBound(0)
                If X > mlFoundCategory Then
                    If bSearchBody = False Then
                        mlFoundPage = -1
                        For Y As Int32 = 0 To moTOCList.oDocs(1).oItems(X).oItems.GetUpperBound(0)
                            If Y > mlFoundIdx AndAlso moTOCList.oDocs(1).oItems(X).oItems(Y).sText.ToUpper.Contains(sSearchFor) = True Then
                                Debug.Print("ID0=" & moTOCList.oDocs(1).oItems(X).oItems(Y).oTVWItem.lItemData.ToString)
                                Debug.Print("ID1=" & moTOCList.oDocs(1).oItems(X).oItems(Y).oTVWItem.lItemData2.ToString)
                                Debug.Print("ID2=" & moTOCList.oDocs(1).oItems(X).oItems(Y).oTVWItem.lItemData3.ToString)
                                moTOCList.oDocs(1).oItems(X).oItems(Y).oTVWItem.ExpandToRoot()
                                tvwTOC.SelectNodeByObject(CType(moTOCList.oDocs(1).oItems(X).oItems(Y), Object))
                                mlFoundCategory = X
                                mlFoundIdx = Y
                                mlFoundPage = -1
                                btnDisplay_Click("btnDisplay")
                                Return
                            End If
                        Next Y
                    Else
                        For Y As Int32 = 0 To moTOCList.oDocs(1).oItems(X).oItems.GetUpperBound(0)
                            If Y > mlFoundIdx Then mlFoundPage = -1
                            For Z As Int32 = 0 To moTOCList.oDocs(1).oItems(X).oItems(Y).oItems.GetUpperBound(0)
                                If Z > mlFoundPage AndAlso moTOCList.oDocs(1).oItems(X).oItems(Y).oItems(Z).sText.ToUpper.Contains(sSearchFor) = True Then
                                    moTOCList.oDocs(1).oItems(X).oItems(Y).oTVWItem.ExpandToRoot()
                                    tvwTOC.SelectNodeByObject(CType(moTOCList.oDocs(1).oItems(X).oItems(Y), Object))
                                    mlFoundCategory = X
                                    mlFoundIdx = Y
                                    mlFoundPage = Z
                                    btnDisplay_Click("btnDisplay")
                                    If Z > 0 Then
                                        Dim oFrm As frmHelpItem = CType(MyBase.moUILib.GetWindow("frmHelpItem"), frmHelpItem)
                                        If Not oFrm Is Nothing Then
                                            For P As Int32 = 1 To Z
                                                oFrm.ChangePage(1)
                                            Next P
                                        End If
                                    End If
                                    Return
                                End If
                            Next Z
                        Next Y
                    End If
                End If
            Next X
        End If
        If mlFoundIdx > 0 Then
            goUILib.AddNotification("No additional Documents found.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
        Else
            goUILib.AddNotification("Help Document not found", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
        End If
        mlFoundCategory = -1
        mlFoundIdx = -1
        mlFoundPage = -1
    End Sub

    'Public Sub FindDocInNode(ByRef tvw As UITreeView, ByRef oNode As UITreeView.UITreeViewItem, ByVal lParentDocID As Int32)
    '    If oItems Is Nothing = False Then
    '        For X As Int32 = 0 To oItems.GetUpperBound(0)
    '            If oItems(X) Is Nothing = False Then
    '                'ok, we add items if the type is not page
    '                If oItems(X).lType <> lItemType.ePage Then
    '                    Dim oNewNode As UITreeView.UITreeViewItem = tvw.AddNode(oItems(X).sText, X, oItems(X).lType, lParentDocID, oNode, Nothing)

    '                    If oItems(X).lType = lItemType.eTopic Then
    '                        oNewNode.oRelatedObject = oItems(X)
    '                    End If

    '                    oItems(X).FindDocInNode(tvw, oNewNode, lParentDocID)
    '                End If
    '            End If
    '        Next X
    '    End If
    'End Sub

    Public Sub ShowForTrigger(ByVal lTrigger As Int32)
        'Ok, go thru our tvw for that trigger
        If moTOCList Is Nothing = False AndAlso moTOCList.oDocs Is Nothing = False Then
            For X As Int32 = 0 To moTOCList.oDocs.GetUpperBound(0)
                Dim oTmp As TOCList.TOCItem = moTOCList.oDocs(X).FindTrigger(lTrigger)
                If oTmp Is Nothing = False Then
                    tvwTOC.SelectNodeByObject(CType(oTmp, Object))
                    Exit For
                End If
            Next X

            btnDisplay_Click("btnDisplay")
        End If
    End Sub

    Private Sub optSubject_Click() Handles optSubject.Click
        optSubject.Value = True
        optBody.Value = False
    End Sub

    Private Sub optBody_Click() Handles optBody.Click
        optSubject.Value = False
        optBody.Value = True
    End Sub
End Class


Public Class TOCList

    Public Enum lItemType As Int32
        eDoc = 0
        eCategory = 1
        eTopic = 2
        ePage = 3
    End Enum

    Public Class TOCItem
        Public sText As String
        Public lType As lItemType
        Public lTrigger As Int32 = 0
        Public oItems() As TOCItem = Nothing
        Public oTVWItem As UITreeView.UITreeViewItem = Nothing

        Public Function AddItem(ByVal sVal As String) As TOCItem
            If oItems Is Nothing = True Then ReDim oItems(-1)

            Dim lUB As Int32 = oItems.GetUpperBound(0) + 1

            ReDim Preserve oItems(lUB)
            oItems(lUB) = New TOCItem
            With oItems(lUB)
                .sText = sVal
            End With
            Return oItems(lUB)
        End Function

        Public Sub FillTreeViewNode(ByRef tvw As UITreeView, ByRef oNode As UITreeView.UITreeViewItem, ByVal lParentDocID As Int32)
            If oItems Is Nothing = False Then
                For X As Int32 = 0 To oItems.GetUpperBound(0)
                    If oItems(X) Is Nothing = False Then
                        'ok, we add items if the type is not page
                        If oItems(X).lType <> lItemType.ePage Then
                            Dim oNewNode As UITreeView.UITreeViewItem = tvw.AddNode(oItems(X).sText, X, oItems(X).lType, lParentDocID, oNode, Nothing)
                            oItems(X).oTVWItem = oNewNode

                            If oItems(X).lType = lItemType.eTopic Then
                                oNewNode.oRelatedObject = oItems(X)
                            End If

                            oItems(X).FillTreeViewNode(tvw, oNewNode, lParentDocID)
                        End If
                    End If
                Next X
            End If
        End Sub

        Public Function FindTrigger(ByVal plTrigger As Int32) As TOCItem
            If plTrigger = lTrigger Then Return Me
            If oItems Is Nothing = False Then
                For X As Int32 = 0 To oItems.GetUpperBound(0)
                    If oItems(X) Is Nothing = False Then
                        Dim oTmp As TOCItem = oItems(X).FindTrigger(plTrigger)
                        If oTmp Is Nothing = False Then Return oTmp
                    End If
                Next X
            End If
            Return Nothing
        End Function
    End Class
    Public oDocs() As TOCItem = Nothing

    Private Sub LoadAllTOC()
        Dim sPath As String = AppDomain.CurrentDomain.BaseDirectory
        If sPath.EndsWith("\") = False Then sPath &= "\"
        sPath &= "Help"

        If Exists(sPath) = False Then Return

        Erase oDocs
        oDocs = Nothing

        Dim colResults As System.Collections.ObjectModel.ReadOnlyCollection(Of String) = My.Computer.FileSystem.GetFiles(sPath, FileIO.SearchOption.SearchAllSubDirectories, "*.xml")
        For Each sValue As String In colResults
            LoadTOCXML(sValue)
        Next

        'Now, sort our Docs and the Doc's Children
        If oDocs Is Nothing = False Then
            For X As Int32 = 0 To oDocs.GetUpperBound(0)
                SortItem(oDocs(X))
            Next X

            'Now, let's sort the docs
            Dim lSorted() As Int32 = Nothing
            Dim lSortedUB As Int32 = -1
            For X As Int32 = 0 To oDocs.GetUpperBound(0)
                Dim lIdx As Int32 = -1

                For Y As Int32 = 0 To lSortedUB
                    If oDocs(lSorted(Y)).sText.ToUpper > oDocs(X).sText.ToUpper Then 'goMinerals(lSorted(Y)).MineralName.ToUpper > goMinerals(X).MineralName.ToUpper Then
                        lIdx = Y
                        Exit For
                    End If
                Next Y
                lSortedUB += 1
                ReDim Preserve lSorted(lSortedUB)
                If lIdx = -1 Then
                    lSorted(lSortedUB) = X
                Else
                    For Y As Int32 = lSortedUB To lIdx + 1 Step -1
                        lSorted(Y) = lSorted(Y - 1)
                    Next Y
                    lSorted(lIdx) = X
                End If
            Next X

            Dim oTmp(lSortedUB) As TOCItem
            For X As Int32 = 0 To lSortedUB
                oTmp(X) = oDocs(lSorted(X))
            Next X
            oDocs = oTmp
        End If

    End Sub

    Private Sub SortItem(ByRef oItem As TOCItem)
        'Now, let's sort the docs
        If oItem Is Nothing OrElse oItem.oItems Is Nothing Then Return

        Dim lSorted() As Int32 = Nothing
        Dim lSortedUB As Int32 = -1
        For X As Int32 = 0 To oItem.oItems.GetUpperBound(0)
            Dim lIdx As Int32 = -1

            For Y As Int32 = 0 To lSortedUB
                If oItem.oItems(lSorted(Y)).sText.ToUpper > oItem.oItems(X).sText.ToUpper Then 'goMinerals(lSorted(Y)).MineralName.ToUpper > goMinerals(X).MineralName.ToUpper Then
                    lIdx = Y
                    Exit For
                End If
            Next Y
            lSortedUB += 1
            ReDim Preserve lSorted(lSortedUB)
            If lIdx = -1 Then
                lSorted(lSortedUB) = X
            Else
                For Y As Int32 = lSortedUB To lIdx + 1 Step -1
                    lSorted(Y) = lSorted(Y - 1)
                Next Y
                lSorted(lIdx) = X
            End If
        Next X

        Dim oTmp(lSortedUB) As TOCItem
        For X As Int32 = 0 To lSortedUB
            oTmp(X) = oItem.oItems(lSorted(X))
        Next X
        oItem.oItems = oTmp
    End Sub

    Private Sub LoadTOCXML(ByVal sFile As String)

        Dim oDoc As XmlDocument
        oDoc = New XmlDocument()
        oDoc.Load(sFile)

        Dim oNode As XmlNode = oDoc.FirstChild()
        While oNode Is Nothing = False
            If oNode.Name.ToUpper = "DOC" Then
                Exit While
            Else : oNode = oNode.NextSibling
            End If
        End While
        'If oNode Is Nothing = False Then
        '    If oNode.FirstChild Is Nothing = False Then
        '        oNode = oNode.FirstChild
        '        While oNode Is Nothing = False
        '            If oNode.Name.ToUpper <> "CATEGORIES" Then
        '                oNode = oNode.NextSibling
        '            Else
        '                Exit While
        '            End If
        '        End While
        '    End If
        'End If

        If oNode Is Nothing = False AndAlso oNode.Name.ToUpper = "DOC" Then
            'oNode = oNode.FirstChild

            If oDocs Is Nothing Then ReDim oDocs(-1)
            Dim lUB As Int32 = oDocs.GetUpperBound(0)

            While oNode Is Nothing = False
                lUB += 1
                ReDim Preserve oDocs(lUB)

                oDocs(lUB) = New TOCItem
                'Dim oRootTVW As TreeNode = tvwMain.Nodes.Add("NewCat")
                oDocs(lUB).lType = lItemType.eDoc
                ParseNode(oNode, oDocs(lUB))
                oNode = oNode.NextSibling
            End While
        End If
    End Sub
    Private Sub ParseNode(ByVal oRootNode As XmlNode, ByRef oRootItem As TOCItem)
        Dim oNode As XmlNode = oRootNode.FirstChild

        While oNode Is Nothing = False
            Select Case oNode.Name.ToUpper
                Case "TITLE"
                    oRootItem.sText = oNode.InnerText
                Case "PAGES"
                    ParseNode(oNode, oRootItem)
                Case "PAGE"
                    Dim oNewNode As TOCItem = Nothing 'TreeNode = Nothing
                    'If tvwRootNode Is Nothing = False Then oNewNode = tvwRootNode.Nodes.Add("NewPage") Else oNewNode = tvwMain.Nodes.Add("NewPage")
                    oNewNode = oRootItem.AddItem("Empty Page") 'Else oNewNode = tvwRootNode.AddItem("NewPage") 'tvwMain.Nodes.Add("NewPage")
                    oNewNode.lType = lItemType.ePage
                    oNewNode.sText = oNode.InnerText.Replace("[VBCRLF]", vbCrLf)
                Case "TOPICS"
                    ParseNode(oNode, oRootItem)
                Case "TOPIC"
                    Dim oNewNode As TOCItem = Nothing
                    'If tvwRootNode Is Nothing = False Then oNewNode = tvwRootNode.Nodes.Add("NewTopic") Else oNewNode = tvwMain.Nodes.Add("NewTopic")
                    oNewNode = oRootItem.AddItem("Untitled Topic") 'Else oNewNode = tvwMain.Nodes.Add("NewTopic")
                    oNewNode.lType = lItemType.eTopic
                    ParseNode(oNode, oNewNode)
                Case "CATEGORY"
                    Dim oNewNode As TOCItem = Nothing
                    'If tvwRootNode Is Nothing = False Then oNewNode = tvwRootNode.Nodes.Add("NewCat") Else oNewNode = tvwMain.Nodes.Add("NewCat")
                    oNewNode = oRootItem.AddItem("Untitled Category") 'Else oNewNode = tvwMain.Nodes.Add("NewCat")
                    oNewNode.lType = lItemType.eCategory
                    ParseNode(oNode, oNewNode)
                Case "CATEGORIES"
                    ParseNode(oNode, oRootItem)
                Case "TRIGGER"
                    oRootItem.lTrigger = CInt(Val(oNode.InnerText))
            End Select
            oNode = oNode.NextSibling
        End While
    End Sub

    Public Sub FillTreeView(ByRef tvw As UITreeView)
        tvw.Clear()

        If oDocs Is Nothing = False Then
            For X As Int32 = 0 To oDocs.GetUpperBound(0)
                Dim oNewNode As UITreeView.UITreeViewItem = tvw.AddNode(oDocs(X).sText, X, oDocs(X).lType, -1, Nothing, Nothing)
                oDocs(X).oTVWItem = oNewNode
                oDocs(X).FillTreeViewNode(tvw, oNewNode, X)
            Next X
        End If
    End Sub

    Public Sub New()
        LoadAllTOC()
    End Sub


End Class