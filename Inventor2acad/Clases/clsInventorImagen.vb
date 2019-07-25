Partial Public Class Inventor2acad
    Public Function ToIPictureDisp(ByVal ico As System.Drawing.Icon) As stdole.IPictureDisp
        Return modPictureDispConverter.ToIPictureDisp(ico)
    End Function

    'Converts an image into a IPictureDisp
    Public Function ToIPictureDisp(ByVal picture As System.Drawing.Image) As stdole.IPictureDisp
        Return modPictureDispConverter.ToIPictureDisp(picture)
    End Function

    Public Function ToImage(ByVal objImage As stdole.IPictureDisp) As System.Drawing.Image
        Return modPictureDispConverter.ToImage(objImage)
    End Function

    Public Shared Function Image2Bytes(ByVal img As System.Drawing.Image) As Byte()
        Dim sTemp As String = System.IO.Path.GetTempFileName()
        Dim fs As New System.IO.FileStream(sTemp, System.IO.FileMode.OpenOrCreate, System.IO.FileAccess.ReadWrite)
        img.Save(fs, System.Drawing.Imaging.ImageFormat.Png)
        fs.Position = 0
        '
        Dim imgLength As Integer = CInt(fs.Length)
        Dim bytes(0 To imgLength - 1) As Byte
        fs.Read(bytes, 0, imgLength)
        fs.Close()
        Return bytes
    End Function

    Public Shared Function Bytes2Image(ByVal bytes() As Byte) As System.Drawing.Image
        If bytes Is Nothing Then Return Nothing
        '
        Dim ms As New System.IO.MemoryStream(bytes)
        Dim bm As System.Drawing.Bitmap = Nothing
        Try
            bm = New System.Drawing.Bitmap(ms)
        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine(ex.Message)
        End Try
        Return bm
    End Function

    Sub ChangeThumbnail(ByVal oDoc As Inventor.Document, ByVal queImagen As String)

        ' Set a reference to the active document
        'Dim oDoc As Document
        'oDoc = oAp.ActiveDocument

        ' Get the "Summary Information" property set
        Dim oPropSet As Inventor.PropertySet
        oPropSet = oDoc.PropertySets("{F29F85E0-4FF9-1068-AB91-08002B27B3D9}")

        ' Get the "Thumbnail" property    
        Dim oProp As Inventor.Property = Nothing
        oProp = oPropSet.ItemByPropId(17)

        Dim oDisp As stdole.IPictureDisp
        oDisp = clsIp.GetIPictureDispFromImage(System.Drawing.Image.FromFile(queImagen))
        'oDisp = LoadPicture("C:\temp\thumbnail.bmp")

        ' Set the value of the thumbnail file property
        oProp.Value = oDisp

        ' Save the document
        oDoc.Save()
    End Sub

    '' ***** USAR EN DLL, NO DA ERROR *****
    ' Display name: Información de resumen de Inventor
    ' Internal name: {F29F85E0-4FF9-1068-AB91-08002B27B3D9}
    ' Nombre: Thumbnail (Miniatura) / Valor:  / Id: 17
    ' Dim prop As Inventor.Property
    ' prop = oD.PropertySets.Item("{F29F85E0-4FF9-1068-AB91-08002B27B3D9}").ItemByPropId(17)
    Public Function DameThumbnailInventorDoc(Optional ByVal oDoc As Inventor.Document = Nothing, Optional ByVal queCamino As String = "") As System.Drawing.Image
        Dim resultado As System.Drawing.Image = Nothing
        Dim estabaabierto As Boolean = False

        If oDoc Is Nothing And queCamino = "" Then
            Return resultado
            Exit Function
        ElseIf oDoc Is Nothing And IO.File.Exists(queCamino) Then
            If FicheroAbierto(queCamino) Then
                oDoc = oAppI.Documents.ItemByName(queCamino)
                estabaabierto = True
            Else
                oAppI.SilentOperation = True
                oDoc = oAppI.Documents.Open(queCamino, False)
                estabaabierto = False
                oAppI.SilentOperation = False
            End If
        End If

        Try
            'If oDoc IsNot Nothing Then
            'resultado = Nothing ' My.Resources.SinImagen.GetThumbnailImage(tamaño, tamaño, Nothing, System.IntPtr.Zero)
            If oDoc IsNot Nothing Then resultado = clsIp.GetImageFromIPictureDisp(oDoc.Thumbnail)
            '' Tiempo de espera 40 segundos máximo.
            'Dim tiempo As Date = Date.Now
            'While oDoc.Thumbnail Is Nothing
            '    If oDoc.Thumbnail IsNot Nothing Then Exit While
            '    If Date.Now > tiempo.AddSeconds(40) Then Exit While
            'End While
            'resultado = Microsoft.VisualBasic.Compatibility.VB6.IPictureToImage(oDoc.Thumbnail)
            'End If
        Catch ex As Exception
            'resultado = Nothing ' My.Resources.SinImagen.GetThumbnailImage(tamaño, tamaño, Nothing, System.IntPtr.Zero)
            'If oDoc IsNot Nothing Then resultado = PictureToImage(oDoc.Thumbnail)
        End Try
        'If resultado Is Nothing Then resultado = My.Resources.SinImagen.GetThumbnailImage(tamaño, tamaño, Nothing, System.IntPtr.Zero)
        If estabaabierto = False Then
            oDoc.Close(True)
            oDoc = Nothing
        End If
        'If resultado IsNot Nothing Then resultado = resultado.GetThumbnailImage(tamaño, tamaño, Nothing, System.IntPtr.Zero)
        Return resultado
    End Function

    Public Function DameThumbnailInventorPro(Optional ByVal oDoc As Inventor.Document = Nothing, Optional ByVal queCamino As String = "") As System.Drawing.Image
        Dim resultado As System.Drawing.Image = Nothing
        Dim estabaabierto As Boolean = False

        If oDoc Is Nothing And queCamino = "" Then
            Return resultado
            Exit Function
        ElseIf oDoc Is Nothing And IO.File.Exists(queCamino) Then
            If FicheroAbierto(queCamino) Then
                oDoc = oAppI.Documents.ItemByName(queCamino)
                estabaabierto = True
            Else
                oAppI.SilentOperation = True
                oDoc = oAppI.Documents.Open(queCamino, False)
                estabaabierto = False
                oAppI.SilentOperation = False
            End If
        End If

        Dim oProp As Inventor.Property = Nothing
        Try
            ' Set a reference to the active document
            'Dim oDoc As Inventor.Document
            'oDoc = oAp.ActiveDocument

            ' Get the "Summary Information" property set
            Dim oPropSet As Inventor.PropertySet
            oPropSet = oDoc.PropertySets("{F29F85E0-4FF9-1068-AB91-08002B27B3D9}")

            ' Get the "Thumbnail" property    
            'Dim oProp As Inventor.Property = Nothing
            oProp = oPropSet.ItemByPropId(17)
            'Debug.Print(oProp.Type.ToString & " / " & oProp.GetType.ToString)
            If oProp IsNot Nothing Then resultado = clsIp.GetImageFromIPictureDisp(oProp.Value)
            'If oProp IsNot Nothing Then resultado = Microsoft.VisualBasic.Compatibility.VB6.IPictureToImage(oProp.Value)
        Catch ex As Exception
            'resultado = Nothing ' My.Resources.SinImagen.GetThumbnailImage(tamaño, tamaño, Nothing, System.IntPtr.Zero)
            'If oProp IsNot Nothing Then resultado = PictureToImage(oProp.Value)
        End Try


        If estabaabierto = False And oDoc IsNot Nothing Then
            oDoc.Close(True)
            oDoc = Nothing
        End If
        Return resultado
    End Function

    ''' <summary>
    ''' Devuelve la imagen previa de un fichero (tamaño small, medium, large o extralarge)
    ''' </summary>
    ''' <param name="camino">Camino completo del fichero</param>
    ''' <param name="tamaño">small, medium, large o extralarge</param>
    ''' <returns>Devuelve System.Drawing.Image</returns>
    ''' <remarks>Tamaño small, medium, large o extralarge</remarks>
    Public Function DameImagenWinShell(camino As String, Optional tamaño As IEnum.TipoVista = IEnum.TipoVista.ExtraLarge) As System.Drawing.Image
        Dim resultado As System.Drawing.Image = Nothing
        'tamaño 0 (pequeño), 1 (Normal), 2(media), 3(Larga), 4(extra-larga)
        Dim st As Microsoft.WindowsAPICodePack.Shell.ShellFile = Nothing
        st.Thumbnail.AllowBiggerSize = True
        st = Microsoft.WindowsAPICodePack.Shell.ShellFile.FromFilePath(camino)
        If tamaño = IEnum.TipoVista.Small Then
            resultado = st.Thumbnail.SmallBitmap
        ElseIf tamaño = IEnum.TipoVista.Medium Then
            resultado = st.Thumbnail.MediumBitmap
        ElseIf tamaño = IEnum.TipoVista.Large Then
            resultado = st.Thumbnail.LargeBitmap
        ElseIf tamaño = IEnum.TipoVista.ExtraLarge Then
            resultado = st.Thumbnail.ExtraLargeBitmap
        Else
            resultado = st.Thumbnail.Bitmap
        End If

        Return resultado
    End Function

    ''' <summary>
    ''' Devuelve el icono de un fichero (tamaño = 0 a 4)
    ''' </summary>
    ''' <param name="camino">Camino completo del fichero</param>
    ''' <param name="tamaño">small, medium, large o extralarge</param>
    ''' <returns>Devuelve System.Drawing.Icono</returns>
    ''' <remarks>Tamaño small, medium, large o extralarge</remarks>
    Public Function DameIconoWinShell(camino As String, Optional tamaño As IEnum.TipoVista = IEnum.TipoVista.ExtraLarge) As System.Drawing.Icon
        Dim resultado As System.Drawing.Icon = Nothing

        'tamaño 0 (pequeño), 1 (Normal), 2(mediano), 3 (Largo), 4(extralargo)
        Dim st As Microsoft.WindowsAPICodePack.Shell.ShellFile = Nothing
        st.Thumbnail.AllowBiggerSize = True
        st = Microsoft.WindowsAPICodePack.Shell.ShellFile.FromFilePath(camino)
        If tamaño = IEnum.TipoVista.Small Then
            resultado = st.Thumbnail.SmallIcon
        ElseIf tamaño = IEnum.TipoVista.Medium Then
            resultado = st.Thumbnail.MediumIcon
        ElseIf tamaño = IEnum.TipoVista.Large Then
            resultado = st.Thumbnail.LargeIcon
        ElseIf tamaño = IEnum.TipoVista.ExtraLarge Then
            resultado = st.Thumbnail.ExtraLargeIcon
        Else
            resultado = st.Thumbnail.Icon
        End If

        Return resultado
    End Function

    Public Function DameThumbnailAprenticeDoc(ByVal queDoc As String) As System.Drawing.Image
        ' Declare the Apprentice object
        Dim oApprentice As New Inventor.ApprenticeServerComponent

        ' Open a document using Apprentice
        Dim oApprenticeDoc As Inventor.ApprenticeServerDocument
        oApprenticeDoc = oApprentice.Open(queDoc)  '  "C:\Test\part.ipt")

        '' Tiempo de espera 40 segundos máximo.
        Dim tiempo As Date = Date.Now
        While oApprenticeDoc.Thumbnail Is Nothing
            If oApprenticeDoc.Thumbnail IsNot Nothing Then Exit While
            If Date.Now > tiempo.AddSeconds(40) Then Exit While
        End While
        Dim oImagen As System.Drawing.Image = Nothing
        oImagen = clsIp.GetImageFromIPictureDisp(oApprenticeDoc.Thumbnail)
        oApprentice.Close()

        DameThumbnailAprenticeDoc = oImagen
        Exit Function
    End Function

    Public Function DameThumbnailAprenticeProp(ByVal queDoc As String) As System.Drawing.Image
        ' Declare the Apprentice object
        Dim oApprentice As New Inventor.ApprenticeServerComponent

        ' Open a document using Apprentice
        Dim oApprenticeDoc As Inventor.ApprenticeServerDocument
        oApprenticeDoc = oApprentice.Open(queDoc)  '  "C:\Test\part.ipt")

        '' TAMBIÉN VALDRÍA
        ' oApprenticeDoc.Thumbnail


        ' Obtain the PropertySets collection
        Dim oPropsets As Inventor.PropertySets
        oPropsets = oApprenticeDoc.PropertySets

        ' Get the "Summary Information" property set
        Dim oPropSet As Inventor.PropertySet
        oPropSet = oPropsets("{F29F85E0-4FF9-1068-AB91-08002B27B3D9}")

        ' Get the "Thumbnail" property    
        Dim oProp As Inventor.Property = Nothing
        oProp = oPropSet.ItemByPropId(17)
        'Debug.Print(oProp.Type.ToString & " / " & oProp.GetType.ToString)

        Dim oImagen As System.Drawing.Image
        oImagen = clsIp.GetImageFromIPictureDisp(oProp.Value)
        oApprentice.Close()

        DameThumbnailAprenticeProp = oImagen
        Exit Function
    End Function
End Class
