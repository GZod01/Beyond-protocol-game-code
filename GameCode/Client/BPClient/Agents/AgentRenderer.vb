Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

Public Class AgentRenderer
    Private Shared moSprite As Sprite
    Private Shared moSrcTexture As Texture
    Private mbSpriteBegun As Boolean = False

    Private moGenerator As facemaker

    Private moAgentTex() As Texture
    Private mlAgentIdx() As Int32
    Private mlAgentUB As Int32 = -1
    Private Shared msDestPath As String

    Public bHasUnknowns As Boolean = False

    Private Sub FaceGenerated(ByVal lID As Int32, ByVal oBMP As Bitmap)

        Try
            If msDestPath Is Nothing Then
                msDestPath = AppDomain.CurrentDomain.BaseDirectory
                If msDestPath.EndsWith("\") = False Then msDestPath &= "\"
                msDestPath &= "Data\Agents\"
                If Exists(msDestPath) = False Then
                    MkDir(msDestPath)
                End If
            End If
            Dim sFile As String = msDestPath & "A" & lID.ToString & ".bmp"
            oBMP.Save(sFile, Imaging.ImageFormat.Bmp)

            Dim oTex As Texture = TextureLoader.FromFile(moSprite.Device, sFile)
            For X As Int32 = 0 To mlAgentUB
                If mlAgentIdx(X) = lID Then
                    moAgentTex(X) = oTex
                End If
            Next X

            Kill(sFile)
        Catch
        End Try
    End Sub

    Protected Overrides Sub Finalize()
		If moSprite Is Nothing = False Then moSprite.Dispose()
        moSprite = Nothing
        If moSrcTexture Is Nothing = False Then moSrcTexture.Dispose()
		moSrcTexture = Nothing
		MyBase.Finalize()
    End Sub

    Public Sub PrepareForBatch()
        ValidateResources()
        If mbSpriteBegun = False Then moSprite.Begin(SpriteFlags.AlphaBlend)
        mbSpriteBegun = True
    End Sub

    Public Sub FinishBatch()
        If mbSpriteBegun = True Then moSprite.End()
        mbSpriteBegun = False
    End Sub

    Public Sub RenderAgent2(ByVal lAgentID As Int32, ByVal rcDestRect As Rectangle, ByVal bFaded As Boolean, ByVal bIsMale As Boolean)
        ValidateResources()

        bHasUnknowns = False

        If moGenerator Is Nothing Then
            moGenerator = New facemaker("")
            AddHandler moGenerator.createface_completed, AddressOf FaceGenerated
            AddHandler moGenerator.generateFace_completed, AddressOf FaceGenerated
        End If

        Dim lIdx As Int32 = -1
        Dim lFirstIdx As Int32 = -1
        For X As Int32 = 0 To mlAgentUB
            If mlAgentIdx(X) = lAgentID Then
                lIdx = X
                Exit For
            ElseIf lFirstIdx = -1 AndAlso mlAgentIdx(X) = -1 Then
                lFirstIdx = X
            End If
        Next X
        If lIdx = -1 Then
            If lFirstIdx = -1 Then
                mlAgentUB += 1
                ReDim Preserve mlAgentIdx(mlAgentUB)
                ReDim Preserve moAgentTex(mlAgentUB)
                lFirstIdx = mlAgentUB
            End If
            mlAgentIdx(lFirstIdx) = lAgentID
            moAgentTex(lFirstIdx) = Nothing 'moSrcTexture
            lIdx = lFirstIdx
        End If

        If moAgentTex(lIdx) Is Nothing Then
            moAgentTex(lIdx) = moSrcTexture
            bHasUnknowns = True
            moGenerator.generateFace(lAgentID, bIsMale)
        End If

        If moAgentTex(lIdx) Is Nothing Then Return
        Dim bEndSprite As Boolean

        Try
            Dim oDesc As SurfaceDescription = moAgentTex(lIdx).GetSurfaceLevel(0).Description
            Dim lImgW As Int32 = moAgentTex(lIdx).GetSurfaceLevel(0).Description.Width
            Dim lImgH As Int32 = moAgentTex(lIdx).GetSurfaceLevel(0).Description.Height

            bEndSprite = Not mbSpriteBegun
            If mbSpriteBegun = False Then
                moSprite.Begin(SpriteFlags.AlphaBlend)
                mbSpriteBegun = True
            End If
            Dim rcSrc As Rectangle = New Rectangle(0, 0, lImgW, lImgH)

            Dim fMultX As Single = CSng(lImgW / rcDestRect.Width)
            Dim fMultY As Single = CSng(lImgH / rcDestRect.Height)

            Dim oLoc As Point = rcDestRect.Location
            Dim pt As Point = New Point(CInt(oLoc.X * fMultX), CInt(oLoc.Y * fMultY))

            If bFaded = True Then
                moSprite.Draw2D(moAgentTex(lIdx), rcSrc, rcDestRect, Point.Empty, 0, pt, System.Drawing.Color.FromArgb(192, 255, 255, 255))
            Else
                moSprite.Draw2D(moAgentTex(lIdx), rcSrc, rcDestRect, Point.Empty, 0, pt, Color.White)
            End If
        Catch
        End Try

        If bEndSprite = True Then
            Try
                moSprite.End()
            Catch
            End Try
            mbSpriteBegun = False
        End If

    End Sub

    Private Sub RenderAgent(ByVal lAgentID As Int32, ByVal rcDestRect As Rectangle, ByVal bFaded As Boolean)
        'Here, we will do the actual work of rendering an agent...
        ValidateResources()

        Dim bEndSprite As Boolean = Not mbSpriteBegun
        If mbSpriteBegun = False Then
            moSprite.Begin(SpriteFlags.AlphaBlend)
            mbSpriteBegun = True
        End If

        If bFaded = True Then
            moSprite.Draw2D(moSrcTexture, Rectangle.Empty, rcDestRect, Point.Empty, 0, rcDestRect.Location, System.Drawing.Color.FromArgb(192, 255, 255, 255))
        Else
            moSprite.Draw2D(moSrcTexture, Rectangle.Empty, rcDestRect, Point.Empty, 0, rcDestRect.Location, Color.White)
        End If

        If bEndSprite = True Then
            moSprite.End()
            mbSpriteBegun = False
        End If
    End Sub

    Public Sub ClearTextures()
        Try
            For X As Int32 = 0 To mlAgentUB
                If moAgentTex(X) Is Nothing = False Then moAgentTex(X).Dispose()
                moAgentTex(X) = Nothing
            Next X
        Catch
        End Try
        moAgentTex = Nothing
        mlAgentIdx = Nothing
        mlAgentUB = -1
    End Sub

    Private Shared Sub ValidateResources()
        Device.IsUsingEventHandlers = False
        If moSprite Is Nothing OrElse moSprite.Disposed = True Then moSprite = New Sprite(goUILib.oDevice)
		If moSrcTexture Is Nothing OrElse moSrcTexture.Disposed = True Then moSrcTexture = goResMgr.LoadScratchTexture("AgentIcon.dds", "apics.pak")
        Device.IsUsingEventHandlers = True
    End Sub

    Public Shared Sub ReleaseSprite()
        Try
            If moSprite Is Nothing = False Then moSprite.Dispose()
            moSprite = Nothing
        Catch
        End Try
    End Sub

    Public Shared goAgentRenderer As AgentRenderer

End Class



Public Class facemaker
    Public Structure FACE_ATTRIBUTES
        Public fileID As Byte
        Public R, G, B As Byte
    End Structure

    Public Event createface_completed(ByVal id As Integer, ByVal bmp As Bitmap)
    Public Event generateFace_completed(ByVal id As Integer, ByVal bmp As Bitmap)
    Public conf As New CONFIG

    Private Function random_number(ByVal k As Integer, ByVal max As Integer) As Integer ' random number is created 
        Dim finalnumber As Int32 = 0
        If max = 1 Then
            Return 1 ' its pointless to go further
        End If
        max = max + 1 ' so that the max number is also included 
        Do
            finalnumber = CInt(Int((max * Rnd())))

            If k > 10000000 Then
                k -= 8000000
            ElseIf k > 1000000 Then
                k -= 800000
            ElseIf k > 100000 Then
                k -= 80000
            ElseIf k > 10000 Then
                k -= 8000
            ElseIf k > 1000 Then
                k -= 800
            Else
                k -= 1
            End If
        Loop While k > 0

        Return finalnumber
    End Function
    Private Function getrandomcolor(ByVal k As Integer, ByVal min As Integer, ByVal max As Integer) As Integer
        Dim finalnumber As Int32 = 0
        If max = 1 Then
            Return 1 ' its pointless to go further
        End If
        max = max + 1
        Do
            finalnumber = CInt(Int((-max * Rnd()) - 5))

            If k > 10000000 Then
                k -= 8000000
            ElseIf k > 1000000 Then
                k -= 800000
            ElseIf k > 100000 Then
                k -= 80000
            ElseIf k > 10000 Then
                k -= 8000
            ElseIf k > 1000 Then
                k -= 800
            Else
                k -= 1
            End If
        Loop While k > 0
        Do While finalnumber < min
            finalnumber += CInt(Int((100 * Rnd())))
        Loop
        Do While finalnumber > (max - 1)
            finalnumber -= CInt(Int((10 * Rnd())))
        Loop
        Return finalnumber
    End Function

    Private Sub do_depend_list(ByVal k As Integer, ByVal list() As String, ByRef img() As Bitmap, ByRef donelist As String, ByVal male As Boolean) ' the dependant list is done
        Dim gender As String = "male"
        Dim dependlist As Array = conf.set_get_data.male_dependent_Layer.Split("/"c)
        If Not male Then
            gender = "female"
            dependlist = conf.set_get_data.female_dependent_Layer.Split("/"c)
        End If
        ' the Array that fits the gender is chosen 
        ' the rest of this code is just prasing the text set_get_data.female_dependent_Layer
        For Each dependitems As String In dependlist
            Dim itm_list() As String = dependitems.Split("|"c)
            Dim Values(itm_list.Length - 1) As System.Collections.ObjectModel.ReadOnlyCollection(Of String)
            Dim fnames(itm_list.Length - 1) As String
            Dim same_filename() As String = Nothing
            Dim ext_for As Boolean = False
            Dim _i_ As Integer = 0
            Dim _j_ As Integer = 0
            For _i As Integer = 0 To itm_list.Length - 1
                fnames(_i) = conf.set_get_data.load_dir & "\" & gender & "\" & itm_list(_i)
                Values(_i) = FileIO.FileSystem.GetFiles(fnames(_i))
                If Array.IndexOf(list, itm_list(_i)) < 0 Then
                    ext_for = True
                    Exit For
                End If
            Next
            If Not ext_for Then
                ReDim same_filename(Values(0).Count - 1)
                For Each fname As String In Values(0)
                    Dim __i As Integer = 1
                    For __i = 1 To itm_list.Length - 1
                        If FileIO.FileSystem.FileExists(fname.Replace("\" & itm_list(0) & "\", "\" & itm_list(__i) & "\")) Then
                            same_filename(_j_) = _i_.ToString
                            _j_ += 1
                        End If
                    Next
                    _i_ += 1
                Next
                Dim rn As Integer = random_number(k, Values(0).Count - 1)
                ' random file is selected 
                'do_random_selection:

                Dim lCnt As Int32 = 0
                Dim bDone As Boolean = False

                While bDone = False AndAlso lCnt < 100
                    lCnt += 1

                    Dim basic_file_name As String = Values(0).Item(rn)
                    Dim filenames(itm_list.Length - 1) As String
                    For __ii As Integer = 0 To itm_list.Length - 1
                        ' the corresponding file is loaded from all other dependent folders 
                        filenames(__ii) = basic_file_name.Replace("\" & itm_list(__ii) & "\", "\" & itm_list(__ii) & "\")

                        Try
                            Dim iindex As Integer = Array.IndexOf(list, itm_list(__ii))
                            If filenames(__ii).IndexOf("Thumbs.db") >= 0 Then
                                rn -= 1
                                Continue While
                            End If
                            img(iindex) = CType(Bitmap.FromFile(filenames(__ii)), Bitmap)
                        Catch ex As Exception
                        End Try

                        donelist &= " " & itm_list(__ii) & " "
                        ' the donelist is updated so that we don't overright this image with a random image 
                    Next
                    bDone = True
                End While

            End If

        Next
    End Sub
    Private ReadOnly Property _getRandomimageList(ByVal k As Integer, ByVal male As Boolean) As Bitmap()
        Get
            Dim img() As Bitmap = Nothing
            Dim i As Integer = 0
            Dim donelist As String = ""
            Dim gender As String = "male"
            Dim ary As Array = conf.set_get_data.male_layernames
            If Not male Then
                gender = "female"
                ary = conf.set_get_data.female_layernames
            End If
            ' the Array that fits the gender is chosen 
            ReDim img(ary.Length - 1)
            do_depend_list(k, conf.set_get_data.male_layernames, img, donelist, male)
            'all the dependant object's images are loaded, the dependant object might include ->'eye|eye_shadow/hair|hair_shade|hair2/'
            For Each s As String In ary
                If Not (donelist.IndexOf(" " & s & " ") >= 0) Then
                    ' the donelist is ingnoreded cause the values are already loaded by do_depend_list
                    Dim Values As System.Collections.ObjectModel.ReadOnlyCollection(Of String)
                    Values = FileIO.FileSystem.GetFiles(conf.set_get_data.load_dir & "\" & gender & "\" & s)
                    Dim rn As Integer = random_number(k, Values.Count - 1)
                    ' a random file is selected from the list 
                    'do_random_selection:
                    Try
                        Dim lCnt As Int32 = 0
                        Dim bDone As Boolean = False
                        While bDone = False AndAlso lCnt < 100
                            lCnt += 1
                            Dim ma As String = Values.Item(rn)
                            If Values.Item(rn).IndexOf("Thumbs.db") >= 0 Then
                                rn -= 1
                                bDone = False
                                Continue While
                            End If
                            bDone = True
                            img(i) = CType(Bitmap.FromFile(Values.Item(rn)), Bitmap) ' the random file is loaded to a image 
                        End While

                        'img(i).MakeTransparent(Color.White)
                    Catch ex As Exception

                    End Try
                End If
                i += 1
            Next

            Return img
        End Get
    End Property


    Private ReadOnly Property _getDefaultvalues() As CONFIG.config 'loads all the default values
        Get
            Dim confi As New CONFIG.config
            confi.load_dir = AppDomain.CurrentDomain.BaseDirectory
            If confi.load_dir.EndsWith("\") = False Then confi.load_dir &= "\"
            confi.load_dir &= "Agents"
            Dim i As Integer = 0
            Dim imglist(16) As String
            imglist(0) = "bg"
            imglist(1) = "neck"
            imglist(2) = "cloth"
            imglist(3) = "ear"
            imglist(4) = "face"
            imglist(5) = "iris" '"eye"
            imglist(6) = "eye_shade"
            imglist(7) = "eye" '"iris"
            imglist(8) = "brow"
            imglist(9) = "mouth_shade"
            imglist(10) = "mouth"
            imglist(11) = "nose"
            imglist(12) = "beard"
            imglist(13) = "hair_shadow"
            imglist(14) = "hair2"
            imglist(15) = "hair"
            imglist(16) = "acc"

            confi.male_layernames = imglist
            confi.male_dependent_Layer = "eye|eye_shade/mouth|mouth_shade/hair|hair_shadow/"
            ReDim imglist(13)
            imglist(0) = "bg"
            imglist(1) = "neck"
            imglist(2) = "cloth"
            imglist(3) = "ear"
            imglist(4) = "face"
            imglist(5) = "eye_shadow"
            imglist(6) = "iris"
            imglist(7) = "eye"
            imglist(8) = "brow"
            imglist(9) = "mouth_shade"
            imglist(10) = "mouth"
            imglist(11) = "nose"
            imglist(12) = "hair"
            imglist(13) = "acc"
            confi.female_layernames = imglist
            confi.female_dependent_Layer = "eye|eye_shadow/mouth|mouth_shade"

            ReDim imglist(8)
            imglist(0) = "neck"
            imglist(1) = "mouth"
            imglist(2) = "mouth_shade"
            imglist(3) = "nose"
            imglist(4) = "ear"
            imglist(5) = "eye"
            imglist(6) = "face"
            imglist(7) = "hair_shadow"
            imglist(8) = "eye_shade"

            confi.face_part = imglist
            confi.new_random_color_part = ("brow  ").Split(" "c)

            ReDim imglist(1)
            imglist(0) = "hair2"
            imglist(1) = "hair"
            confi.hair = imglist

            ReDim imglist(7)
            imglist(0) = "neck"
            imglist(1) = "mouth"
            imglist(2) = "eye_shade"
            imglist(3) = "nose"
            imglist(4) = "ear"
            imglist(5) = "eye"
            imglist(6) = "face"
            imglist(7) = "eye_shadow"
            confi.female_face_part = imglist
            confi.female_new_random_color_part = ("brow  ").Split(" "c)
            ReDim imglist(0)
            imglist(0) = "hair"
            confi.female_hair = imglist
            Dim cl(2) As Color
            cl(0) = Color.FromArgb(255, 248, 248, 248)
            cl(1) = Color.FromArgb(255, 238, 205, 209)
            cl(2) = Color.FromArgb(255, 200, 180, 180)
            confi.skinTones = cl
            Return confi
        End Get
    End Property
    Public Sub New(ByVal configFilename As String)
        'If configFilename = "" Then
        conf.set_get_data = _getDefaultvalues
        '    'load default values 
        'Else
        '    Dim xs As New System.Runtime.Serialization.Formatters.Soap.SoapFormatter
        '    Dim strm As New System.IO.FileStream(configFilename, IO.FileMode.OpenOrCreate)
        '    conf = xs.Deserialize(strm)
        '    strm.Close()
        '    ' load the values from the file
        'End If
        'save_mod_config(conf, "Config")
        ' load the config for future use 
    End Sub
    'Sub save_mod_config(ByVal _config As CONFIG, ByVal filename As String)
    '    Dim xs As New System.Runtime.Serialization.Formatters.Soap.SoapFormatter
    '    Dim strm As New System.IO.FileStream(filename, IO.FileMode.Create)
    '    xs.Serialize(strm, _config)
    '    strm.Close()
    'End Sub

    Public Sub generateFace(ByVal ID As Integer, ByVal bIsMale As Boolean)
        Dim th As New Threading.Thread(AddressOf getRandomimage)
        'Dim obj() As Object = {ID, bIsMale}
        Dim oParms As FuncCallParms = New FuncCallParms(ID, bIsMale)
        th.Start(oParms)
        'start a new thread for face genration
    End Sub
    Public Sub CreateFace(ByVal ID As Integer, ByVal bIsMale As Boolean, ByVal c() As FACE_ATTRIBUTES)
        Dim th As New Threading.Thread(AddressOf getimage)
        'Dim obj() As Object = {ID, bIsMale, c}
        Dim oParms As FuncCallParms = New FuncCallParms(ID, bIsMale, c)
        th.Start(oParms)
        'start a new thread for face creation
    End Sub


    Private Sub getRandomimage(ByVal oArgs As Object)
        Dim oParms As FuncCallParms = CType(oArgs, FuncCallParms)
        Dim id As Integer = oParms.lID
        Dim ismale As Boolean = oParms.bIsMale
        Dim imglist() As Bitmap = _getRandomimageList(id, ismale)
        ' random image list is loaded 
        Dim i As Integer = 0

        Dim returnimg As New Bitmap(imglist(0).Width, imglist(0).Height) ' the final image 
        Dim gr As Graphics = Graphics.FromImage(returnimg) ' this enables us to draw diffrenet layers on the iamge 


        Dim tone As Color = conf.set_get_data.skinTones(random_number(id, conf.set_get_data.skinTones.Length - 1))
        Dim __r, __b, __g, __a As Byte
        __a = tone.A
        __r = CByte(tone.R + getrandomcolor(id, -5, 5))
        __b = CByte(tone.B + getrandomcolor(id, -5, 5))
        __g = CByte(tone.G + getrandomcolor(id, -5, 5))
        ' random color is selected from the list  for skin color

        Dim h__r, h__b, h__g As Byte
        h__r = CByte(random_number(id, 255))
        h__b = CByte(random_number(id, 255))
        h__g = CByte(random_number(id, 255))
        ' random color is selected for hair 

        For i = 0 To imglist.Length - 1 ' each image is drawn on the return image

            Dim faceary As Array = conf.set_get_data.face_part
            Dim newary As Array = conf.set_get_data.new_random_color_part
            Dim hair As Array = conf.set_get_data.hair
            Dim ary() As String = conf.set_get_data.male_layernames

            If Not ismale Then
                faceary = conf.set_get_data.female_face_part
                newary = conf.set_get_data.female_new_random_color_part
                hair = conf.set_get_data.female_hair
                ary = conf.set_get_data.female_layernames
            End If
            ' the Array that correspondes to the gender is loaded 

            Dim im As New Imaging.ImageAttributes ' has the .... image attributes 

            If Array.IndexOf(faceary, ary(i)) >= 0 Then
                ' all the face parts from the array faceary are colored 
                Dim ptsArray As Single()() = { _
                          New Single() {__r / 255.0F, 0, 0, 0, 0}, _
                          New Single() {0, __b / 255.0F, 0, 0, 0}, _
                          New Single() {0, 0, __g / 255.0F, 0, 0}, _
                          New Single() {0, 0, 0, __a / 255.0F, 0}, _
                          New Single() {0, 0, 0, 0, 1}}
                Dim cMatrix As New Imaging.ColorMatrix(ptsArray)
                im.SetColorMatrix(cMatrix)
            End If

            If Array.IndexOf(hair, ary(i)) >= 0 Then
                ' all the hair parts from the array hair are colored
                Dim ptsArray As Single()() = { _
                          New Single() {h__r / 255.0F, 0, 0, 0, 0}, _
                          New Single() {0, h__b / 255.0F, 0, 0, 0}, _
                          New Single() {0, 0, h__g / 255.0F, 0, 0}, _
                          New Single() {0, 0, 0, 1, 0}, _
                          New Single() {0, 0, 0, 0, 1}}
                Dim cMatrix As New Imaging.ColorMatrix(ptsArray)
                im.SetColorMatrix(cMatrix)
            End If


            If Array.IndexOf(newary, ary(i)) >= 0 Then
                ' all the other parts from the array newarry are colored
                Dim _r, _b, _g As Byte
                _r = CByte(random_number(id, 255))
                _b = CByte(random_number(id, 255))
                _g = CByte(random_number(id, 255))
                Dim ptsArray As Single()() = { _
                          New Single() {_r / 255.0F, 0, 0, 0, 0}, _
                          New Single() {0, _b / 255.0F, 0, 0, 0}, _
                          New Single() {0, 0, _g / 255.0F, 0, 0}, _
                          New Single() {0, 0, 0, 1, 0}, _
                          New Single() {0.2, 0.2, 0.2, 0, 1}}
                Dim cMatrix As New Imaging.ColorMatrix(ptsArray)
                im.SetColorMatrix(cMatrix)
            End If


            Try
                gr.DrawImage(imglist(i), New Rectangle(0, 0, imglist(i).Width, imglist(i).Height), _
             0, 0, imglist(i).Width, imglist(i).Height, GraphicsUnit.Pixel, im)
            Catch ex As Exception
                ' ignored if an error occured 
                ' its a good idea cause sometime some of the objects in array are not removed but rather changed to "nothing"
            End Try
        Next

        RaiseEvent generateFace_completed(id, returnimg)
    End Sub
    Private Sub getimage(ByVal oArgs As Object)
        Dim oParms As FuncCallParms = CType(oArgs, FuncCallParms)
        Dim id As Integer = oParms.lID
        Dim ismale As Boolean = oParms.bIsMale
        Dim c() As FACE_ATTRIBUTES = oParms.c
        Dim imglist(c.Length - 1) As Bitmap
        Dim i As Integer = 0
        Dim gender As String = "male"
        Dim ary() As String = conf.set_get_data.male_layernames
        If Not ismale Then
            gender = "female"
            ary = conf.set_get_data.female_layernames
        End If

        For Each attributes As FACE_ATTRIBUTES In c 'all the images are loaded as per the user's input 
            Dim fname As String = conf.set_get_data.load_dir & "\" & gender & "\" & ary(i) & "\"
            Dim Values As System.Collections.ObjectModel.ReadOnlyCollection(Of String)
            Values = FileIO.FileSystem.GetFiles(fname)
            imglist(i) = CType(Bitmap.FromFile(Values.Item(id)), Bitmap)
            i += 1
        Next

        Dim returnimg As New Bitmap(imglist(0).Width, imglist(0).Height) ' the final image 
        Dim gr As Graphics = Graphics.FromImage(returnimg)
        i = 0

        For i = 0 To imglist.Length - 1 ' the images are drawn on the returnimage with the color
            Dim im As New Imaging.ImageAttributes
            Dim ptsArray As Single()() = { _
            New Single() {c(i).R / 255.0F, 0, 0, 0, 0}, _
            New Single() {0, c(i).G / 255.0F, 0, 0, 0}, _
            New Single() {0, 0, c(i).B / 255.0F, 0, 0}, _
            New Single() {0, 0, 0, 1, 0}, _
            New Single() {0, 0, 0, 0, 1}}
            Dim cMatrix As New Imaging.ColorMatrix(ptsArray)
            im.SetColorMatrix(cMatrix)
            gr.DrawImage(imglist(i), New Rectangle(0, 0, imglist(i).Width, imglist(i).Height), _
         0, 0, imglist(i).Width, imglist(i).Height, GraphicsUnit.Pixel, im)
        Next
        RaiseEvent createface_completed(id, returnimg)
    End Sub

    Private Class FuncCallParms
        Public lID As Int32
        Public bIsMale As Boolean
        Public c() As FACE_ATTRIBUTES

        Public Sub New(ByVal plID As Int32, ByVal pbIsMale As Boolean)
            lID = plID
            bIsMale = pbIsMale
        End Sub
        Public Sub New(ByVal plID As Int32, ByVal pbIsMale As Boolean, ByVal pc() As FACE_ATTRIBUTES)
            lID = plID
            bIsMale = pbIsMale
            c = pc
        End Sub
    End Class

    <Serializable()> Public Class CONFIG
        <Serializable()> Structure config
            Public load_dir As String ' the dir to start from
            Public male_layernames() As String ' Array ' all the layer name of the male character in the right order 
            Public female_layernames() As String ' As Array 'same list but for female character 
            Public female_dependent_Layer As String
            'Example of a dependent layer would be eye and eye_shade so there fore it would be represented like ->
            'eye|eye_shade/ or eye_shade|eye/. another example would be hair|hair_shadow|hair2/
            Public male_dependent_Layer As String 'same for male 

            Public hair() As String ' Array
            Public face_part() As String ' Array
            Public new_random_color_part() As String ' Array

            Public female_hair() As String ' Array
            Public female_face_part() As String ' Array
            Public female_new_random_color_part() As String ' Array

            Public skinTones() As System.Drawing.Color ' Array
        End Structure
        Dim conf As config
        Property set_get_data() As config
            Get
                Return conf
            End Get
            Set(ByVal value As config)
                conf = value
            End Set
        End Property
    End Class

    ' if you need to change/add/remove any of the folders/colors then simpley load the 
    ' "facemaker.conf.set_get_data" to a "facemaker.CONFIG.config" variable and edit any information you need 
    ' if you want to save those values then call facemaker.save_mod_config to save the current "config" the class has 
    ' If anyone has any question about this code please feel free to contact me. 
    ' Regards and Goodluck 
    ' by :
    ' Ajay
End Class