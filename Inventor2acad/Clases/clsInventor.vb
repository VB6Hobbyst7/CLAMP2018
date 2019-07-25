Option Compare Text

Imports Inventor
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports Microsoft.Win32
Imports System.Linq
Imports System.IO
Imports Microsoft.VisualBasic
Imports Microsoft.WindowsAPICodePack.Shell
'
Partial Public Class Inventor2acad
    Public Sub New(ByVal queApp As Inventor.Application)
        LlenaObjetosPrincipalesClase(queApp)
        Try
            If Me.Log Then PonLog("Nueva instancia creada de " & app_nameandversion, True)
        Catch ex As Exception
            MsgBox("No se puede crear " & app_log & vbCrLf & vbCrLf & ex.ToString)
        End Try
    End Sub
    Protected Overrides Sub Finalize()
        VaciaTodo()
        MyBase.Finalize()
    End Sub
    Public Sub LlenaObjetosPrincipalesClase(ByVal queApp As Inventor.Application)
        If queApp Is Nothing Then
            Try
                oAppI = GetObject(, "Inventor.Application")
            Catch ex As Exception
                MsgBox("Inventor no está abierto... Cerrando aplicación")
                Exit Sub
            End Try
        Else
            oAppI = queApp
        End If
        'If oAppCls Is Nothing Then oAppCls = queApp 'oAppCls = GetObject(, "Inventor.Application")
        If (oAppI IsNot Nothing) Then
            oAppIEv = oAppI.ApplicationEvents
            oTg = oAppI.TransientGeometry
            oTo = oAppI.TransientObjects
            oTBr = oAppI.TransientBRep
            oCm = oAppI.CommandManager
            If oAppI.Documents.Count > 0 Then
                oSelSet = oAppI.ActiveDocument.SelectSet
            Else
                oSelSet = Nothing
            End If
        End If
        dirProyectoInv = oAppI.FileLocations.Workspace
        If dirProyectoInv.EndsWith("\") = False Then dirProyectoInv &= "\"
        ' Inicializar clases
        clsIp = New clsiPictureToImage()
    End Sub

    Public Function UnidadesEs(queValor As String) As String
        Dim oUOM As Inventor.UnitsOfMeasure = oAppI.UnitsOfMeasure

        Dim queUni As UnitsTypeEnum = oAppI.ActiveDocument.UnitsOfMeasure.LengthUnits
        If IsNumeric(queValor) Then
            Return queValor.Replace(".", ",")
        Else
            Dim valores() As String
            valores = Split(queValor)
            Return oUOM.GetLocaleCorrectedExpression(queValor, valores(1))
        End If
    End Function

    Public Sub MensajeInventor(ByVal quetexto As String)
        Call oCm.PromptMessage(quetexto, vbOK, "Avisos")
    End Sub

    Public Function FicheroAbierto(ByVal queFichero As String) As Boolean
        Dim resultado As Boolean = False

        For Each oD As Inventor.Document In Me.oAppI.Documents
            If oD.FullFileName = queFichero Then
                resultado = True
                Exit For
            End If
        Next
        Return resultado
    End Function

    Public Function FicheroVisible(ByVal queFichero As String) As Boolean
        Dim resultado As Boolean = False

        For Each oD As Inventor.Document In Me.oAppI.Documents.VisibleDocuments
            If oD.FullFileName = queFichero Then
                resultado = True
                Exit For
            End If
        Next
        Return resultado
    End Function

    Public Function FicheroCierra(ByVal queFichero As String) As Boolean
        Dim resultado As Boolean = False

        For Each oD As Inventor.Document In Me.oAppI.Documents
            If oD.FullFileName = queFichero Then
                Try
                    oD.Close(True)
                    resultado = True
                Catch ex As Exception
                    resultado = False
                End Try
                Exit For
            End If
        Next
        Return resultado
    End Function

    Public Sub DoEventsInventor(Optional ByVal tambienInventor As Boolean = True)
        System.Windows.Forms.Application.DoEvents()
        If tambienInventor Then oAppI.UserInterfaceManager.DoEvents()
    End Sub

    Private Function iMateDamePorNombre(oCo As ComponentOccurrence, Name As String) As iMateDefinition
        Dim iMateDef As iMateDefinition = Nothing
        ''
        Dim oDef As Object = Nothing
        If oCo.DefinitionDocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
            oDef = CType(oCo.Definition, AssemblyComponentDefinition)
        ElseIf oCo.DefinitionDocumentType = DocumentTypeEnum.kPartDocumentObject Then
            oDef = CType(oCo.Definition, PartComponentDefinition)
        Else
            Return iMateDef
            Exit Function
        End If
        '' Es un ensamblaje o un pieza. Iteramos con los iMateDefinitions que tengan.
        For Each imate As iMateDefinition In oDef.iMateDefinitions
            If imate.Name = Name Then
                iMateDef = imate
                Exit For
            End If
        Next
        ' Devuelve el iMateDefinition o Nothing si no se ha encontrado.
        Return iMateDef
    End Function

    Public Function InsertarComponenteGetPoint(ByVal queFichero As String, ensPadre As AssemblyDocument) As ComponentOccurrence
        ''
        Dim resultado As ComponentOccurrence = Nothing
        ''
        If oAppI.ActiveDocumentType <> Inventor.DocumentTypeEnum.kAssemblyDocumentObject Then
            MsgBox("Macro only for Assemblies...")
            Return resultado
            Exit Function
        End If
        ''Utilidades.GuardaMensaje("***** INSERTAR PIEZA NUEVA EN ENSAMBLAJE (INSERTAREF) *****")
        Dim tg As Inventor.TransientGeometry
        tg = oAppI.TransientGeometry
        ''
        Dim compD As AssemblyComponentDefinition
        compD = ensPadre.ComponentDefinition
        ''
        Dim omatrix As Inventor.Matrix
        omatrix = tg.CreateMatrix
        ''
        '' Get inertion point
        Dim oPt As Inventor.Point = Get3dPointInFace(False)
        If oPt IsNot Nothing Then
            Call omatrix.SetTranslation(tg.CreateVector(oPt.X, oPt.Y, oPt.Z))
            ''
            Try
                oAppI.SilentOperation = True
                resultado = ensPadre.ComponentDefinition.Occurrences.Add(queFichero, omatrix)
                ensPadre.Rebuild2()
                ensPadre.Update2()
            Catch ex As Exception
                ''Utilidades.GuardaMensaje("***** ERROR INSERTANDO PIEZA NUEVA *****")
            Finally
                ''Utilidades.GuardaMensaje("***** PIEZA NUEVA INSERTADA CORRECTAMENTE (INSERTAREF) *****")
                oAppI.SilentOperation = False
            End Try
        End If
        ''
        Return resultado
    End Function
    '
    Public Function InsertCreateOccProjectionSeleccionaFace(m_inApp As Inventor.Application,
                                              queOut As String,
                                         Optional crearExtrusion As Boolean = False,
                                         Optional oOcc As ComponentOccurrence = Nothing,
                                         Optional queBoceto As String = "",
                                         Optional queFiInsertCreate As String = "",
                                         Optional quePlantilla As String = "") As PlanarSketchProxy
        ''
        Dim resultado As PlanarSketchProxy = Nothing
        ''
        If m_inApp.ActiveEditDocument.DocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then
            MsgBox("Only for ActiveEditDocument = Async...", MsgBoxStyle.Critical, "ERROR")
            Return resultado
            Exit Function
        End If
        '' Comprobar valores opciones y rellenarlos si estan vacios "" o Nothing
        If quePlantilla = "" OrElse IO.File.Exists(quePlantilla) = False Then
            quePlantilla = m_inApp.GetTemplateFile(DocumentTypeEnum.kPartDocumentObject)
        End If
        If queFiInsertCreate = "" OrElse IO.File.Exists(queFiInsertCreate) = False Then
            'Dim nameFi As String = IO.Path.GetRandomFileName
            Dim nameFi As String = "2aCADTemp.ipt"
            queFiInsertCreate = IO.Path.Combine(IO.Path.GetTempPath, nameFi)
        End If
        ' queOut = "Ent" (Entidadas), "Rec" (Rectangulo 2D) o "Loo" (EdgeLoop Exterior)
        Dim colqueOut As String() = New String() {"Ent", "Rec", "Loo"}
        If queOut = "" Or colqueOut.Contains(queOut) = False Then
            queOut = "Ent"
        End If
        ''
        Try
            ' Set a reference to the assembly component definintion.
            ' This assumes an assembly document is open.
            Dim oAsmCompDef As AssemblyComponentDefinition
            oAsmCompDef = m_inApp.ActiveEditDocument.ComponentDefinition
            ' Ask the user to select a face
            ' Call the Pick method of the clsSelect object and set
            ' the filter to pick any face.
            Dim oFace As FaceProxy
            oFace = SelectionDame(SelectionFilterEnum.kPartFaceFilter, "Select Face...")
            'oSelect = Nothing
            ' Check to make sure a face was selected.
            If oFace Is Nothing Then
                Return resultado
                Exit Function
            End If
            ''
            If oOcc Is Nothing Then
                ' Create a matrix. A new matrix is initialized with an identity matrix.
                Dim oMatrix As Matrix = Me.oTg.CreateMatrix
                '
                If queFiInsertCreate = "" OrElse IO.File.Exists(queFiInsertCreate) = False Then
                    Dim oPartDoc As PartDocument = m_inApp.Documents.Add(DocumentTypeEnum.kPartDocumentObject, quePlantilla, False)
                    Call oPartDoc.SaveAs(queFiInsertCreate, False) : oPartDoc.Close()
                End If
                ' Add the occurrence.
                oOcc = oAsmCompDef.Occurrences.Add(queFiInsertCreate, oMatrix)
            End If
            '
            Dim oOccDef As PartComponentDefinition = oOcc.Definition
            ' Comprobar si creanos nuevo boceto o cogemos, si existe, el nombre de queBoceto
            Dim oSketch As PlanarSketch = Nothing
            If queBoceto = "" Then
                oSketch = oOccDef.Sketches.Add(oOccDef.WorkPlanes(3), False) '' Crear PlanarSketch en plano 3 (XY)
            Else
                For Each oSketch In oOccDef.Sketches
                    If oSketch.Name = queBoceto Then
                        Exit For
                    End If
                Next
                ''
                If oSketch Is Nothing Then
                    oSketch = oOccDef.Sketches.Add(oOccDef.WorkPlanes(3), False)
                End If
            End If
            '
            ' Create a proxy for the sketch in the newly created part.
            Dim oSketchProxyTemp As Object = Nothing    ' PlanarSketchProxy = Nothing
            Dim oSketchProxy As PlanarSketchProxy = Nothing
            Call oOcc.CreateGeometryProxy(oSketch, oSketchProxyTemp)
            oSketchProxy = CType(oSketchProxyTemp, PlanarSketchProxy)
            ''
            '' Crear el objeto NonParametricBaseFeature con el FaceProxy
            Dim oParFea As NonParametricBaseFeature = BaseFeatureSurfaceBodyDame_InFace(oFace.NativeObject, oFace.ContainingOccurrence, oOcc)
            If oParFea Is Nothing OrElse TypeOf oParFea.Faces(1).Geometry Is Cone Then
                If oParFea IsNot Nothing Then oParFea.Delete()
                oSketch.Delete()
                Return Nothing
                Exit Function
            End If
            '' Si oParFea.Faces(1).Geometry = Cylinder o Cono. Crear eje entre los 2 circulos
            Dim oFaNon As Face = oParFea.Faces(1)
            If TypeOf oFaNon.Geometry Is Cylinder Then
                Dim oSk3D As Sketch3D = FaceCylinderConeCreaWorkPointEje_DameSketch3D(oParFea.Faces(1), oOcc.Definition)
                If oSk3D Is Nothing Then
                    oSketch.Delete()
                    Return Nothing
                    Exit Function
                End If
                Dim escorrecto As Boolean = SurfaceCilindricalCreaContorno2DSketchLine3D(oSk3D, oOcc.Definition, oSketchProxy.NativeObject.Name)
                If escorrecto = False Then
                    oSketch.Delete()
                    Return Nothing
                    Exit Function
                End If
            ElseIf TypeOf oFaNon.Geometry Is Plane Then
                If FacePlaneDireccionHaciaAbajo(oFace, oAsmCompDef) Then
                    MsgBox("Hacia arriba" & vbCrLf & "Angulo = " & FacePlaneDameAnguloSobreXY(oFace, oAsmCompDef))
                Else
                    MsgBox("Hacia abajo o perpendicular" & vbCrLf & "Angulo = " & FacePlaneDameAnguloSobreXY(oFace, oAsmCompDef))
                    Return Nothing
                    Exit Function
                End If
                ''
                Dim oSketchEnt As SketchEntity = Nothing
                Select Case queOut
                    Case "Ent"          '' Proyectas todas las entidas
                        For Each oEdge As EdgeProxy In oFace.Edges
                            oSketchEnt = oSketchProxy.AddByProjectingEntity(oEdge)
                        Next
                    Case "Rec"          '' Crear Rectangulo 2D con el Evaluator--RangeBox (Coger solo coordenadas 2D)
                        Dim oBox As Box = oFace.Evaluator.RangeBox
                        Dim ptMin2D As Point2d = oTg.CreatePoint2d(oBox.MinPoint.X, oBox.MinPoint.Y)
                        Dim ptMax2D As Point2d = oTg.CreatePoint2d(oBox.MaxPoint.X, oBox.MaxPoint.Y)

                        Dim oSketchEntEnum As SketchEntitiesEnumerator = oSketchProxy.SketchLines.AddAsTwoPointRectangle(ptMin2D, ptMax2D)
                    Case "Loo"          '' Proyectar sólo las entidades externas
                        For Each oLoop As EdgeLoopProxy In oFace.EdgeLoops
                            If oLoop.IsOuterEdgeLoop = False Then Continue For
                            ''
                            For Each oEdge As EdgeProxy In oFace.Edges
                                oSketchEnt = oSketchProxy.AddByProjectingEntity(oEdge)
                            Next
                            Exit For
                        Next
                End Select
            End If
            ''
            If crearExtrusion = True Then
                '' Crear la extrusión contra la cara de oParFea
                Dim oExf As ExtrudeFeature = ExtrudeFaceFeature(oParFea, oSketchProxy.NativeObject)
            End If
            oAppI.ActiveView.Update()
            ''
            resultado = oSketchProxy
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
            resultado = Nothing
        End Try
        '
        Return resultado
    End Function
    '' Busca un fichero de plano (DWG o IDW) que se llame igual que solo el nombre o que empiece por el nombre del fichero
    '' GetExtension(C:\pp.txt) devuelve ".txt"
    ''' <summary>
    ''' Busca en el directorio del fichero "queFichero" planos IDW o DWG
    ''' También en los subdirectorios que cuelguen de este directorio
    ''' No tiene en cuenta los que se encuentran en "OldVersions" no los DWG de AutoCAD
    ''' </summary>
    ''' <param name="queFichero">Camino completo del fichero: IAM, IPT o IPN</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ExisteFicheroPlanoEnDirBasico(ByVal queFichero As String) As ArrayList
        Dim resultado As New ArrayList

        Dim directorio As String = IO.Path.GetDirectoryName(queFichero)
        Dim soloNombre As String = IO.Path.GetFileNameWithoutExtension(queFichero)
        '' Llenamos un arraylist con todos los ficheros IDW y DWG del directori y subdirectorios indicados.
        Dim ficheros As New ArrayList
        ficheros.AddRange(IO.Directory.GetFiles(directorio, "*.idw", IO.SearchOption.AllDirectories))
        ficheros.AddRange(IO.Directory.GetFiles(directorio, "*.dwg", IO.SearchOption.AllDirectories))

        For Each f As String In ficheros
            Dim Nombre As String = IO.Path.GetFileName(f)
            If Nombre.StartsWith(soloNombre) = True AndAlso f.ToLower.Contains("oldversions") = False Then
                '' Si es DWG, pero de AutoCAD, pasamos al siguiente.
                If IO.Path.GetExtension(f).EndsWith("dwg") AndAlso oAppI.FileManager.IsInventorDWG(f) = False Then Continue For
                If resultado.Contains(f) = False Then resultado.Add(f)
            End If
        Next
        ficheros = Nothing

        Return resultado
    End Function

    '' Busca un fichero de plano (DWG o IDW) que se llame igual que solo el nombre o que empiece por el nombre del fichero
    '' GetExtension(C:\pp.txt) devuelve ".txt"
    ''' <summary>
    ''' Busca en el directorio del fichero "queFichero" planos IDW o DWG
    ''' También en los subdirectorios que cuelguen de este directorio
    ''' No tiene en cuenta los que se encuentran en "OldVersions" no los DWG de AutoCAD
    ''' *** Lo abre y comprueba si la vista base (DrawingView.ParentView=Nothin) tiene este documento.
    ''' </summary>
    ''' <param name="queFichero">Camino completo del fichero: IAM, IPT o IPN</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ExisteFicheroPlanoEnDirProfundoInv(ByVal queFichero As String) As ArrayList
        Dim resultado As New ArrayList
        Dim arrPlanos As New ArrayList

        Dim directorio As String = IO.Path.GetDirectoryName(queFichero)
        Dim soloNombre As String = IO.Path.GetFileNameWithoutExtension(queFichero)
        '' Llenamos un arraylist con todos los ficheros IDW y DWG del directori y subdirectorios indicados.
        Dim ficheros As New ArrayList
        ficheros.AddRange(IO.Directory.GetFiles(directorio, "*.idw", IO.SearchOption.AllDirectories))
        ficheros.AddRange(IO.Directory.GetFiles(directorio, "*.dwg", IO.SearchOption.AllDirectories))
        '' Buscamos en "ficheros" los planos que se llamen igual o empiecen por "soloNombre"
        '' No añadimos los que estén en "OldVersions" no los que sean DWG de AutoCAD
        For Each f As String In ficheros
            'Dim Nombre As String = IO.Path.GetFileName(f)
            'If Nombre.StartsWith(soloNombre) = True AndAlso f.Contains("OldVersions") = False Then
            '' Si es DWG, pero es de AutoCAD, continuamos al siguiente.
            If f.ToLower.Contains("oldversions") = True Then Continue For
            If IO.Path.GetExtension(f).EndsWith("dwg") AndAlso oAppI.FileManager.IsInventorDWG(f) = False Then Continue For
            If arrPlanos.Contains(f) = False Then arrPlanos.Add(f)
            'End If
        Next
        ficheros = Nothing

        '' Ahora abrimos cada plano de "arrPlanos" y buscaremos en su vista principal (ParentView)
        '' Si el documento que refleja "queFichero" es el FullFilenamo del objeto Documento que contiene.
        Dim oDib As DrawingDocument = Nothing

        oAppI.SilentOperation = True
        For Each queF As String In arrPlanos
            ' Crear un nuevo NameValueMap object
            Dim oDocOpenOptions As NameValueMap
            oDocOpenOptions = oAppI.TransientObjects.CreateNameValueMap
            'PrivateRepresentationFileName  (Type=String / Valid Documents=Assembly)   
            'DesignViewRepresentation  (Type=String / Valid Documents=Assembly)   
            'PositionalRepresentation  (Type=String / Valid Documents=Assembly)    
            'LevelOfDetailRepresentation    (Type=String / Valid Documents=Assembly)  NOTA: Typically, the LevelOfDetailRepresentation to use should be provided in the form of a FulDocumentName (first argument). But if this is provided separately, you should make sure that it does not conflict with the FullDocumentName argument by providing FullFileName as the first argument rather than a FullDocumentName.  
            'DeferUpdates    (Type=Boolean / Valid Documents=Drawing)      
            'FileVersionOption  (Type=FileVersionEnum / Valid Documents=All)    NOTA: Valid values for FileVersionEnum are kOpenOldVersion, kOpenCurrentVersion and kRestoreOldVersionToCurrent. If set to kOpenOldVersion, save will not be allowed on the opened document. kRestoreOldVersionToCurrent is valid only if no other versions are open and the current version is not checked out.  
            'ImportNonInventorDWG  (Type=Boolean / Valid Documents=Imports the DWG file to an IDW if True, Opens it into Inventor DWG if False)  NOTA: When opening non-Inventor DWG files, this method honors the application option to decide between open and import, unless an override is specified in the Options argument.  
            'Password  (Type=String / Valid Documents=All)    
            'SkipAllUnresolvedFiles  (Type=Boolean / Valid Documents=All)    

            ' Set the representations to use when opening the document.
            'Call oDocOpenOptions.Add("LevelOfDetailRepresentation", "MyLODRep")
            'Call oDocOpenOptions.Add("PositionalRepresentation", "MyPositionalRep")
            'Call oDocOpenOptions.Add("DesignViewRepresentation", "MyDesignViewRep")
            Call oDocOpenOptions.Add("DeferUpdates", True)
            'oDib = oAppCls.Documents.Open(queF, False)
            Try
                oDib = oAppI.Documents.OpenWithOptions(queF, oDocOpenOptions, False)
                For Each oSh As Sheet In oDib.Sheets
                    For Each oV As DrawingView In oSh.DrawingViews
                        '' oV.ParentView Is Nothing para la Vista Base Principal.
                        If oV.ParentView Is Nothing AndAlso oV.ReferencedDocumentDescriptor.ReferencedFileDescriptor.FullFileName = queFichero Then
                            If resultado.Contains(queF) = Nothing Then resultado.Add(queF)
                            Exit For
                        End If
                    Next
                Next
                oDib.DrawingSettings.DeferUpdates = False
            Catch ex As Exception
                'MsgBox("ExisteFicheroPlanoEnDirProfundoInv :  " & vbCrLf & vbCrLf & ex.Message)
                Continue For
            End Try
        Next
        oAppI.SilentOperation = True

        Return resultado
    End Function

    Public Function ExisteFicheroPlanoEnArray(ByVal queArr As ArrayList, ByVal fullFichero As String) As String
        Dim resultado As String = ""
        Dim SoloNombre As String = IO.Path.GetFileNameWithoutExtension(fullFichero)
        Dim SoloExtension As String = IO.Path.GetExtension(fullFichero)

        For Each fichero As String In queArr
            Dim ficheroNom As String = IO.Path.GetFileNameWithoutExtension(fichero)
            Dim ficheroExt As String = IO.Path.GetExtension(fichero)
            If ficheroNom = SoloNombre And ficheroExt = SoloExtension Then
                resultado = fichero
                Exit For
            End If
            If ficheroNom.StartsWith(SoloNombre) And fichero.EndsWith(SoloExtension) Then
                resultado = fichero
                Exit For
            End If
        Next

        Return resultado
    End Function
    Public Sub RibbonTabActiva(ByVal queNombreRibbon As String)
        Try
            If Me.oAppI.UserInterfaceManager.ActiveEnvironment.Ribbon.RibbonTabs.Item(queNombreRibbon).Active = False Then _
        Me.oAppI.UserInterfaceManager.ActiveEnvironment.Ribbon.RibbonTabs.Item(queNombreRibbon).Active = True
        Catch ex As Exception
            Debug.Print("El Ribbon " & queNombreRibbon & " no existe en entorno " & Me.oAppI.UserInterfaceManager.ActiveEnvironment.DisplayName)
        End Try
    End Sub

    Public Function DameSurfBody(ByRef oC As ComponentOccurrence) As SurfaceBodyProxy
        Dim gn As SurfaceBody = oC.Definition.SurfaceBodies(1)
        Dim gn1 As SurfaceBodyProxy = Nothing

        oC.CreateGeometryProxy(gn, CType(gn1, SurfaceBodyProxy))

        Return gn1
    End Function

    Public Function BASEDame(ByVal oCEns As ComponentOccurrence, Optional ByVal queProp As String = "_BASE", Optional ByVal Proxy As Boolean = True) As Object
        Dim oW As Object
        Dim valor As String
        Dim oD As Inventor.AssemblyDocument
        Dim oP As Inventor.PartDocument


        valor = "Centro"
        oD = oCEns.ReferencedDocumentDescriptor.ReferencedDocument

        'On Error Resume Next
        Try
            valor = oD.PropertySets.Item("User Defined Properties").Item(queProp).Value
        Catch
            ' Si da un error, es que no existe la propiedad en el ensamblaje padre. valor = "Centro"
        End Try
        oP = oD.ComponentDefinition.Occurrences.ItemByName(oCEns.Name).ReferencedDocumentDescriptor.ReferencedDocument
        oW = oP.ComponentDefinition.WorkPoints.Item(1)
        If oW.Name = "Center Point" Then oW.Name = "Centro"
        Try
            oW = oP.ComponentDefinition.WorkPoints.Item(valor)
        Catch ex As Exception
            oW = oP.ComponentDefinition.WorkPoints.Item(1)
        End Try
        If Proxy = True Then oCEns.SubOccurrences.Item(1).CreateGeometryProxy(oW, oW)
        BASEDame = oW
        Exit Function
    End Function

    Public Function DamePunto(ByVal oC As ComponentOccurrence, ByVal nombre As String, Optional ByVal Proxy As Boolean = True) As Object
        Dim oP As PartDocument
        Dim oA As AssemblyDocument
        Dim oProx As Object = Nothing   'WorkPointProxy = Nothing ' Object = Nothing
        Dim resultado As WorkPoint = Nothing
        Try
            If oC.DefinitionDocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                oA = oC.ReferencedDocumentDescriptor.ReferencedDocument
                'On Error Resume Next
                If Proxy = True Then
                    oC.CreateGeometryProxy(oA.ComponentDefinition.WorkPoints.Item(nombre), oProx)
                    resultado = CType(oProx, WorkPointProxy).Point
                Else
                    resultado = oA.ComponentDefinition.WorkPoints.Item(nombre).Point
                End If
            ElseIf oC.DefinitionDocumentType = DocumentTypeEnum.kPartDocumentObject Then
                oP = oC.ReferencedDocumentDescriptor.ReferencedDocument
                'On Error Resume Next
                If Proxy = True Then
                    oC.CreateGeometryProxy(oP.ComponentDefinition.WorkPoints.Item(nombre), oProx)
                    resultado = CType(oProx, WorkPointProxy)
                Else
                    resultado = oP.ComponentDefinition.WorkPoints.Item(nombre)
                End If
            End If
        Catch ex As Exception
            'MsgBox("No existe el Punto de Trabajo (" & nombre & ") Creelo y vuelva a intentarlo")
            'resultado = oAp.TransientGeometry.CreatePoint(0, 0, 0)
            resultado = Nothing
        End Try

        DamePunto = resultado
        Exit Function
    End Function


    Public Function DamePlano(ByVal oC As ComponentOccurrence, ByVal quePlano As Integer, Optional ByVal Proxy As Boolean = True) As Object
        Dim oP As PartDocument
        Dim oA As AssemblyDocument
        Dim oProx As Object = Nothing   'WorkPlaneProxy = Nothing ' Object = Nothing
        Dim resultado As WorkPlane = Nothing
        ' quePlano = 1 (YZ Plane)
        ' quePlano = 2 (XZ Plane)
        ' quePlano = 3 (XY Plane)
        Try
            If oC.DefinitionDocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                oA = oC.ReferencedDocumentDescriptor.ReferencedDocument
                'On Error Resume Next
                If Proxy = True Then
                    oC.CreateGeometryProxy(oA.ComponentDefinition.WorkPlanes.Item(quePlano), oProx)
                    resultado = CType(oProx, WorkPlaneProxy).Plane
                Else
                    resultado = oA.ComponentDefinition.WorkPlanes.Item(quePlano).Plane
                End If
            ElseIf oC.DefinitionDocumentType = DocumentTypeEnum.kPartDocumentObject Then
                oP = oC.ReferencedDocumentDescriptor.ReferencedDocument
                'On Error Resume Next
                If Proxy = True Then
                    oC.CreateGeometryProxy(oP.ComponentDefinition.WorkPlanes.Item(quePlano), oProx)
                    resultado = CType(oProx, WorkPlaneProxy)
                Else
                    resultado = oP.ComponentDefinition.WorkPlanes.Item(quePlano)
                End If
            End If
        Catch ex As Exception
            'MsgBox("No existe el Punto de Trabajo (" & nombre & ") Creelo y vuelva a intentarlo")
            'resultado = oAp.TransientGeometry.CreatePoint(0, 0, 0)
            resultado = Nothing
        End Try

        Return resultado
    End Function

    Public Function DameRad(ByVal grados As Object) As Double
        Dim resultado As Double
        If IsNumeric(grados) Then
            resultado = CDbl(FormatNumber(grados, 2, Microsoft.VisualBasic.TriState.False, , Microsoft.VisualBasic.TriState.False))
            'resultado = CDbl(Format(grados, "###0.00"))
        Else
            MsgBox("El valor enviado a DAMERAD no es un número que pueda convertirse en grados..")
            resultado = 0.0#
        End If
        DameRad = resultado * (Math.PI / 180)
        Exit Function
    End Function
    ''
    Public Sub MueveComponenteRelativo(ByVal queEns As AssemblyDocument, ByVal queOc As String, ByVal distancia As Double, ByVal queEje As String)
        Dim oMatrixTemp, oMatrix As Matrix
        Dim oC As ComponentOccurrence = queEns.ComponentDefinition.Occurrences.ItemByName(queOc)
        If Me.oTg Is Nothing Then Me.oTg = CType(queEns.ComponentDefinition.Application, Inventor.Application).TransientGeometry
        'Dim oTg As TransientGeometry = CType(queEns.ComponentDefinition.Application, Application).TransientGeometry
        ' queEje podrá ser X ,Y o Z
        '' Creaos Matrix temporal para reflejar el movimiento en X,Y ó Z
        oMatrixTemp = Me.oTg.CreateMatrix()
        '' Guardamos en oMatrix la actual Matrix del componente.
        oMatrix = oC.Transformation
        If queEje.ToLower = "x" Then
            oMatrixTemp.SetTranslation(oTg.CreateVector(distancia, 0, 0))
        ElseIf queEje.ToLower = "y" Then
            oMatrixTemp.SetTranslation(oTg.CreateVector(0, distancia, 0))
        ElseIf queEje.ToLower = "z" Then
            oMatrixTemp.SetTranslation(oTg.CreateVector(0, 0, distancia))
        End If
        '' Le sumamos a oMatriz + oMatrixTemp (esto sumará los valores)
        oMatrix.TransformBy(oMatrixTemp)
        '' Le aplicamos al componente oMatrix con los valores sumados (se producirá el movimiento)
        oC.Transformation = oMatrix
        oMatrixTemp = Nothing
        oMatrix = Nothing
    End Sub

    Public Sub MueveComponenteRelativo(ByVal oC As ComponentOccurrence, ByVal distancia As Double, ByVal queEje As String)
        Dim oMatrixTemp, oMatrix As Matrix
        'Dim oC As ComponentOccurrence = queEns.ComponentDefinition.Occurrences.ItemByName(queOc)
        If Me.oTg Is Nothing Then Me.oTg = CType(oC.Definition.Application, Inventor.Application).TransientGeometry
        'Dim oTg As TransientGeometry = CType(queEns.ComponentDefinition.Application, Application).TransientGeometry
        ' queEje podrá ser X ,Y o Z
        '' Creaos Matrix temporal para reflejar el movimiento en X,Y ó Z
        oMatrixTemp = Me.oTg.CreateMatrix()
        '' Guardamos en oMatrix la actual Matrix del componente.
        oMatrix = oC.Transformation
        If queEje.ToLower = "x" Then
            oMatrixTemp.SetTranslation(oTg.CreateVector(distancia, 0, 0))
        ElseIf queEje.ToLower = "y" Then
            oMatrixTemp.SetTranslation(oTg.CreateVector(0, distancia, 0))
        ElseIf queEje.ToLower = "z" Then
            oMatrixTemp.SetTranslation(oTg.CreateVector(0, 0, distancia))
        End If
        '' Le sumamos a oMatriz + oMatrixTemp (esto sumará los valores)
        oMatrix.TransformBy(oMatrixTemp)
        '' Le aplicamos al componente oMatrix con los valores sumados (se producirá el movimiento)
        oC.Transformation = oMatrix
        oMatrixTemp = Nothing
        oMatrix = Nothing
    End Sub


    Public Sub MueveComponenteAbsoluto(ByVal oC As ComponentOccurrence,
                                   Optional ByVal queX As Double = 0,
                                   Optional ByVal queY As Double = 0,
                                   Optional ByVal queZ As Double = 0)
        '' Si no hemos puesto valores. Salimos
        If queX + queY + queZ = 0 Then Exit Sub
        '' Si ya está en la nueva posición. Salimos
        If oC.Transformation.Translation.X = queX And
        oC.Transformation.Translation.Y = queY And
        oC.Transformation.Translation.Z = queZ Then Exit Sub

        Dim oTg As TransientGeometry = CType(oC.Definition.Application, Inventor.Application).TransientGeometry
        Dim oMatrix As Matrix = oTg.CreateMatrix
        '' Guardamos el punto origen y destino del movimiento

        oMatrix.SetTranslation(oTg.CreateVector(queX, queY, queZ))
        oC.Transformation = oMatrix
        oMatrix = Nothing
    End Sub

    Public Sub MueveComponenteAbsolutoSobreEje(ByVal oC As ComponentOccurrence, ByVal distancia As Double, ByVal queEje As String)
        Dim oMatrixTemp, oMatrix As Matrix
        'Dim oC As ComponentOccurrence = queEns.ComponentDefinition.Occurrences.ItemByName(queOc)
        If Me.oTg Is Nothing Then Me.oTg = CType(oC.Definition.Application, Inventor.Application).TransientGeometry
        'Dim oTg As TransientGeometry = CType(queEns.ComponentDefinition.Application, Application).TransientGeometry
        ' queEje podrá ser X ,Y o Z
        '' Creaos Matrix temporal para reflejar el movimiento en X,Y ó Z
        oMatrixTemp = Me.oTg.CreateMatrix()
        '' Guardamos en oMatrix la actual Matrix del componente.
        oMatrix = oC.Transformation
        If queEje.ToLower = "x" Then
            oMatrixTemp.SetTranslation(oTg.CreateVector(distancia, oMatrix.Translation.Y, oMatrix.Translation.Z))
        ElseIf queEje.ToLower = "y" Then
            oMatrixTemp.SetTranslation(oTg.CreateVector(oMatrix.Translation.X, distancia, oMatrix.Translation.Z))
        ElseIf queEje.ToLower = "z" Then
            oMatrixTemp.SetTranslation(oTg.CreateVector(oMatrix.Translation.X, oMatrix.Translation.Y, distancia))
        End If
        '' Le sumamos a oMatriz + oMatrixTemp (esto sumará los valores)
        'oMatrix.TransformBy(oMatrixTemp)
        '' Le aplicamos al componente oMatrix con los valores sumados (se producirá el movimiento)
        oC.Transformation = oMatrixTemp
        oMatrixTemp = Nothing
        oMatrix = Nothing
    End Sub

    Public Sub BrowseNode(ByVal oComp As ComponentOccurrence, ByVal expandir As Boolean, Optional ByVal padre As Boolean = False)
        Dim oNodePadre As BrowserNode = Nothing
        Dim oNode As BrowserNode = Nothing
        Dim oNNativo As NativeBrowserNodeDefinition = oAppI.ActiveDocument.BrowserPanes.GetNativeBrowserNodeDefinition(oComp)
        Dim oNodeCarpeta As BrowserFolder = Nothing
        oNodePadre = oAppI.ActiveDocument.BrowserPanes.ActivePane.TopNode
        Try
            If padre = True Then
                oNode = oNodePadre
            Else
                oNode = oNodePadre.AllReferencedNodes(oNNativo).Item(1)
            End If

            oNode.Expanded = expandir
        Catch ex As Exception
            '' No hacemos nada y continuamos
            Debug.Print(ex.Message)
        End Try
    End Sub

    Public Function OperacionesPiezaDesactivadasBorra(ByVal queApp As Inventor.Application, ByVal queFile As String) As Integer
        Dim contador As Integer = 0
        'Dim ContadorPasadas As Integer = 0
        Dim quePieza As PartDocument = Nothing
        Dim cerrar As Boolean = False
        If queApp.ActiveDocumentType = DocumentTypeEnum.kPartDocumentObject AndAlso queApp.ActiveDocument.FullFileName = queFile Then
            quePieza = queApp.ActiveDocument
        Else
            quePieza = queApp.Documents.Open(queFile)
            cerrar = True
        End If

        'OTRAVEZ:
        'If ContadorPasadas <= 1 Then
        For Each queOpe As PartFeature In quePieza.ComponentDefinition.Features
            Try
                Dim feaPadre As PartFeature = queOpe.OwnedBy
                Dim retener As Boolean = False
                '' Si hay una feature padre y esta activa. No la borraremos. retener=True
                If feaPadre IsNot Nothing AndAlso feaPadre.Suppressed = False Then retener = True
                If queOpe.Suppressed = True Then
                    queOpe.Delete(False, retener, False)   ' Retener sólo features.
                    contador += 1
                End If
            Catch ex As Exception
                Continue For
            End Try
        Next
        For Each queBoc As PlanarSketch In quePieza.ComponentDefinition.Sketches
            Try
                If queBoc.Consumed = False Then
                    queBoc.Delete()
                    contador += 1
                Else
                    queBoc.Visible = False
                    queBoc.Shared = False
                End If
            Catch ex As Exception
                Continue For
            End Try
        Next
        For Each queBoc As Sketch3D In quePieza.ComponentDefinition.Sketches3D
            Try
                If queBoc.Consumed = False Then
                    queBoc.Delete()
                    contador += 1
                Else
                    queBoc.Visible = False
                    queBoc.Shared = False
                End If
            Catch ex As Exception
                Continue For
            End Try
        Next
        '' Este apartado si queremos borrar también los puntos de alineación. Quitar comentarios.
        'For Each quePunto As WorkPoint In quePieza.ComponentDefinition.WorkPoints
        'Try
        'If quePunto.Consumed = False Then
        'quePunto.Delete()
        'Else
        'quePunto.Visible = False
        'quePunto.Shared = False
        'End If
        'Catch ex As Exception
        'Continue For
        'End Try
        'Next
        '**************************************
        For Each quePlano As WorkPlane In quePieza.ComponentDefinition.WorkPlanes
            Try
                If quePlano.Consumed = False Then
                    quePlano.
                quePlano.Delete()
                    contador += 1
                Else
                    quePlano.Visible = False
                    quePlano.Shared = False
                End If
            Catch ex As Exception
                Continue For
            End Try
        Next
        'ContadorPasadas += 1
        'GoTo OTRAVEZ            ' Haremos otra pasada, para ver si hay Features sin consumir.
        'End If

        Call quePieza.Update2()
        quePieza.Save2()
        If cerrar = True Then quePieza.Close(True)
        OperacionesPiezaDesactivadasBorra = contador
        Exit Function
    End Function

    Public Function OperacionesPiezaActivadasLee(ByVal queApp As Inventor.Application, ByVal queFile As String) As Hashtable
        Dim queOpeAct As New Hashtable
        Dim quePieza As PartDocument = Nothing
        Dim cerrar As Boolean = False
        If queApp.ActiveDocumentType = DocumentTypeEnum.kPartDocumentObject Then
            quePieza = queApp.ActiveDocument
        Else
            quePieza = queApp.Documents.Open(queFile)
            cerrar = True
        End If

        For Each queOpe As PartFeature In quePieza.ComponentDefinition.Features
            Try
                If queOpe.Suppressed = False Then
                    queOpeAct.Add(queOpe.Name, queOpe)
                End If
            Catch ex As Exception
                Continue For
            End Try
        Next
        If cerrar = True Then quePieza.Close(True)
        OperacionesPiezaActivadasLee = queOpeAct
        Exit Function
    End Function

    Public Sub GiraCompRelativoProyecto(ByVal oC As ComponentOccurrence, ByVal grados As Double, ByVal tipo As String)
        Dim oCHijo As ComponentOccurrence = Nothing
        Dim ptProx As Object = Nothing
        Dim oAsm As Inventor.AssemblyDocument = Nothing
        ' Quitamos FIJO del ensamblaje Padre (el que tiene todos los componentes)
        oC.Grounded = False
        ' Objeto AssemblyDocument con el Padre.
        oAsm = oC.ReferencedDocumentDescriptor.ReferencedDocument
        ' Objeto ComponentOccurrence del primer hijo (que se llama igual que el padre)
        oCHijo = oAsm.ComponentDefinition.Occurrences.Item(1)

        ptProx = BASEDame(oC, , True)

        Dim oMatrixTemp As Matrix
        Dim oMatrix As Matrix

        If oTg Is Nothing Then oTg = oAppI.TransientGeometry
        '' Matriz temporal para almacenar el giro de X grados (en radianes) a derecha, sobre el punto de alineación.
        oMatrixTemp = oTg.CreateMatrix
        'Call oMatrixTemp.SetToRotation(-DameRad(45), oTg.CreateVector(0, 0, 1), oWp.Point)
        If tipo = "de" Then
            Call oMatrixTemp.SetToRotation(-DameRad(grados), oTg.CreateVector(0, 0, 1), CType(ptProx, WorkPointProxy).Point)
        ElseIf tipo = "iz" Then
            Call oMatrixTemp.SetToRotation(DameRad(grados), oTg.CreateVector(0, 0, 1), CType(ptProx, WorkPointProxy).Point)
        ElseIf tipo = "" Then
            oC.Transformation = colMatrix(oC.Name)  ' oMatrixInicial
            'oCHijo.Transformation.SetToRotation(0, oTg.CreateVector(0, 0, 1), oWp.Point)
        End If

        oAppI.UserInterfaceManager.DoEvents()
        If tipo <> "" Then
            '' Matriz actual con los datos del Componente.
            oMatrix = oC.Transformation ' oC.Transformation
            '' Sobre la matriz del componente aplicamos la temporal (pre-multiplica sus valores)
            '' giro de oMatrixTemp + posición de oMatrix.
            Call oMatrix.TransformBy(oMatrixTemp)
            '' Aplicamos oMatrix al componente.
            oC.Transformation = oMatrix
        End If
        oC.Grounded = True
        oMatrixTemp = Nothing
        oMatrix = Nothing

        'oAppCls.ActiveView.Update()
    End Sub
    ' quePadre será el camino completo
    ' queHijo será el nombre del componente "componente:1")
    Public Sub GiraCompAbsolutoPadre(ByVal quePadre As String, ByVal queHijo As String, ByVal grados As Double)
        If grados = 0 Then Exit Sub
        If Dir(quePadre) = "" Then
            MsgBox("GiraComponenteAbsolutoPadre --> No existe el fichero PADRE " & quePadre)
            Exit Sub
        End If
        oAppI.SilentOperation = True
        Dim ensamPadre As AssemblyDocument = oAppI.Documents.Open(quePadre, True)
        Dim oCHijo As ComponentOccurrence = ensamPadre.ComponentDefinition.Occurrences.ItemByName(queHijo)

        Dim ptProx As Object = Nothing
        ' Quitamos FIJO del ensamblaje Padre (el que tiene todos los componentes)
        oCHijo.Grounded = False
        ' Objeto AssemblyDocument con el Padre.
        'oAsm = oC.ReferencedDocumentDescriptor.ReferencedDocument
        ' Objeto ComponentOccurrence del primer hijo (que se llama igual que el padre)
        'oCHijo = oAsm.ComponentDefinition.Occurrences.Item(1)

        ptProx = BASEDame(oCHijo, , True)

        Dim oMatrixTemp As Matrix
        Dim oMatrix As Matrix

        If oTg Is Nothing Then oTg = oAppI.TransientGeometry
        '' Matriz temporal para almacenar el giro de X grados (en radianes) a derecha, sobre el punto de alineación.
        oMatrixTemp = oTg.CreateMatrix
        'Call oMatrixTemp.SetToRotation(-DameRad(45), oTg.CreateVector(0, 0, 1), oWp.Point)
        Call oMatrixTemp.SetToRotation(DameRad(grados), oTg.CreateVector(0, 0, 1), CType(ptProx, WorkPointProxy).Point)



        '' Matriz actual con los datos del Componente.
        oMatrix = oCHijo.Transformation ' oC.Transformation
        '' Sobre la matriz del componente aplicamos la temporal (pre-multiplica sus valores)
        '' giro de oMatrixTemp + posición de oMatrix.
        Call oMatrix.TransformBy(oMatrixTemp)
        '' Aplicamos oMatrix al componente.
        oCHijo.Transformation = oMatrix

        oCHijo.Grounded = True
        oMatrixTemp = Nothing
        oMatrix = Nothing
        ensamPadre.Update2()
        ensamPadre.Save2(False)
        ensamPadre.Close()
        oAppI.SilentOperation = False
        'oAppCls.ActiveView.Update()
    End Sub

    Public Sub VerticalRestringePuntos(ByVal iamProyecto As AssemblyDocument, ByVal oC As ComponentOccurrence, ByVal pt1Proyecto As Point)
        Dim oCHijo As ComponentOccurrence = Nothing
        Dim iamPadre As AssemblyDocument = Nothing
        Dim ptProx1 As Object = Nothing
        Dim pt1 As WorkPoint = Nothing

        ' Quitamos FIJO del ensamblaje Padre (el que tiene todos los componentes)
        If oC.Grounded = True Then oC.Grounded = False
        ' Objeto AssemblyDocument del oC que recibimos.
        iamPadre = oC.ReferencedDocumentDescriptor.ReferencedDocument
        ' Objeto ComponentOccurrence del primer hijo (que se llama igual que el padre)
        oCHijo = iamPadre.ComponentDefinition.Occurrences.Item(1)

        Dim proxTemp As WorkPointProxy = BASEDame(oC, , )

        If proxTemp.Name = ptIAM Then
            ptProx1 = Me.DamePunto(oCHijo, ptIAM, True)
        Else
            ptProx1 = Me.DamePunto(oCHijo, ptIAM1, True)
        End If

        iamProyecto.ComponentDefinition.Occurrences.ItemByName(oC.Name).CreateGeometryProxy(ptProx1, ptProx1)


        pt1 = iamProyecto.ComponentDefinition.WorkPoints.AddFixed(pt1Proyecto)  'CType(ptProx1, WorkPointProxy).Point)
        pt1.Grounded = True

        Dim flusCons1 As MateConstraint = Nothing
        flusCons1 = iamProyecto.ComponentDefinition.Constraints.AddMateConstraint(ptProx1, pt1, 0)
        If iamProyecto.RequiresUpdate Then iamProyecto.Update2()
        oC.Grounded = True
        flusCons1.Delete()
        pt1.Delete(True)
        oAppI.ActiveView.Update()
        'MsgBox("Ver todo")
    End Sub

    Public Sub HorizontalRestringePuntos(ByVal iamProyecto As AssemblyDocument, ByVal oC As ComponentOccurrence, ByVal pt1Proyecto As Point, ByVal pt2Proyecto As Point)
        oAppI.ScreenUpdating = False
        Dim oCHijo As ComponentOccurrence = Nothing
        Dim iamPadre As AssemblyDocument = Nothing
        Dim ptProx1 As Object = Nothing
        Dim ptProx2 As Object = Nothing
        Dim pt1 As WorkPoint = Nothing
        Dim pt2 As WorkPoint = Nothing
        '' XY Plane del ensamblaje del proyecto
        Dim XYwp1 As WorkPlane = iamProyecto.ComponentDefinition.WorkPlanes(3)
        '' YZ Plane Proxy del componente Hijo
        Dim XZprox2 As Object = Nothing
        '' XY Plane Proxy del componente Hijo
        Dim XYprox3 As Object = Nothing
        Dim altura As Double
        '' AddMateConstraint (cara opuesta a cara)
        '' AddFlushConstraint (cara alineada a cara)
        ' Quitamos FIJO del ensamblaje Padre (el que tiene todos los componentes)
        oC.Grounded = False
        ' Objeto AssemblyDocument del oC que recibimos.
        iamPadre = oC.ReferencedDocumentDescriptor.ReferencedDocument
        ' Objeto ComponentOccurrence del primer hijo (que se llama igual que el padre)
        oCHijo = iamPadre.ComponentDefinition.Occurrences.ItemByName(oC.Name)
        'XYprox1 = oCHijo.ContextDefinition

        ''***** LOG PARA CONTROL DE ERRORES *****
        If Log Then PonLog(vbCrLf & "Vamos a alinear " & oC.Name & vbCrLf)
        ''*****************************************
        Dim proxTemp As WorkPointProxy = BASEDame(oC, , )

        If proxTemp.Name = ptIAM Then
            ptProx1 = Me.DamePunto(oCHijo, ptIAM, True)
            ptProx2 = Me.DamePunto(oCHijo, ptIAM1, True)
        Else
            ptProx1 = Me.DamePunto(oCHijo, ptIAM1, True)
            ptProx2 = Me.DamePunto(oCHijo, ptIAM, True)
        End If


        iamProyecto.ComponentDefinition.Occurrences.ItemByName(oC.Name).CreateGeometryProxy(ptProx1, ptProx1)
        iamProyecto.ComponentDefinition.Occurrences.ItemByName(oC.Name).CreateGeometryProxy(ptProx2, ptProx2)


        pt1 = iamProyecto.ComponentDefinition.WorkPoints.AddFixed(pt1Proyecto)  'CType(ptProx1, WorkPointProxy).Point)
        pt1.Grounded = True
        pt2 = iamProyecto.ComponentDefinition.WorkPoints.AddFixed(pt2Proyecto)              'CType(ptProx2, WorkPointProxy).Point)
        pt2.Grounded = True
        Dim distancia As Double = Me.oAppI.MeasureTools.GetMinimumDistance(pt1, pt2)
        'Dim angulo As Double = Me.oAppCls.MeasureTools.GetAngle(pt1, pt2)
        ''***** LOG PARA CONTROL DE ERRORES *****
        If Log Then PonLog("Punto1: " & pt1.Point.X & ", " & pt1.Point.Y & ", " & pt1.Point.Z & vbCrLf)
        If Log Then PonLog("Punto2: " & pt2.Point.X & ", " & pt2.Point.Y & ", " & pt2.Point.Z & vbCrLf)
        If Log Then PonLog("Distancia: " & distancia & vbCrLf)
        If Log Then PonLog("Ahora pasamos la distancia a 'ds_lar'" & vbCrLf)
        ''*****************************************
        Try
            Dim queDist As Double = Me.ParametroLeeDouble(oCHijo.ReferencedDocumentDescriptor.ReferencedDocument, "ds_lar")
            If FormatNumber(distancia, 8) <> FormatNumber(queDist, 8) Then
                Me.ParametroEscribeDouble(oCHijo.ReferencedDocumentDescriptor.ReferencedDocument, "", "ds_lar", distancia)
                CType(oCHijo.ReferencedDocumentDescriptor.ReferencedDocument, Inventor.Document).Save2()
                If Log Then PonLog("**  ds_lar (Antes = " & FormatNumber(queDist, 8) & " ) (Ahora = " & FormatNumber(distancia, 8) & " )" & vbCrLf)
            End If
            queDist = Me.ParametroLeeDouble(oCHijo.ReferencedDocumentDescriptor.ReferencedDocument, "ds_lar")
            If FormatNumber(distancia, 8) <> FormatNumber(queDist, 8) Then
                If Log Then PonLog("** ERROR. La distancia real no se ha guardado en 'ds_lar'" & vbCrLf)
            End If

        Catch ex As Exception
            If Log Then PonLog("Error pasando la distancia a 'ds_lar'. Continua la aplicación" & vbCrLf)
        End Try

        '' Solo aplicaremos las restricciones si los puntos son diferentes. Así ganamos tiempo.
        If pt1.Point.X <> CType(ptProx1, Inventor.WorkPointProxy).Point.X Or
    pt1.Point.Y <> CType(ptProx1, Inventor.WorkPointProxy).Point.Y Or
    pt1.Point.Z <> CType(ptProx1, Inventor.WorkPointProxy).Point.Z Or
    pt2.Point.X <> CType(ptProx2, Inventor.WorkPointProxy).Point.X Or
    pt2.Point.Y <> CType(ptProx2, Inventor.WorkPointProxy).Point.Y Or
    pt2.Point.Z <> CType(ptProx2, Inventor.WorkPointProxy).Point.Z Then

            Dim flusCons1 As MateConstraint = Nothing
            Dim flusCons1a As FlushConstraint = Nothing
            Dim flusCons2 As MateConstraint = Nothing
            Dim angCons1 As AngleConstraint = Nothing

            '' Ponemos la restriccion de Centro 1 y 2. Los puntos se han calculado bien antes.
            Try
                If Log Then PonLog("Restringimos punto 1" & vbCrLf)
                flusCons1 = iamProyecto.ComponentDefinition.Constraints.AddMateConstraint(ptProx1, pt1, 0)
            Catch ex As Exception
                If Log Then PonLog("ERROR. No se ha podido restringir Punto 1" & vbCrLf)
            End Try

            Try
                If Log Then PonLog("Restringimos punto 2" & vbCrLf)
                flusCons2 = iamProyecto.ComponentDefinition.Constraints.AddMateConstraint(ptProx2, pt2, 0)
            Catch ex As Exception
                If Log Then PonLog("ERROR. No se ha podido restringir Punto 2" & vbCrLf)
            End Try

            '' Solo restringimos los planos si la Z es igual en los 2 puntos (si no daría error)
            '' Esto es para que no se gire sobre su eje Z la pieza horizontal.
            '' En horizontales inclinadas ponemos restriccion para que no gire.
            If pt1.Point.Z = pt2.Point.Z Then
                If Log Then PonLog("No es inclinada. Ponemos restriccion FlusConstraint en plano XY para que no gire" & vbCrLf)
                altura = pt1.Point.Z
                XYprox3 = Me.DamePlano(oCHijo, 3, True)   ' 3 = XY Plane
                iamProyecto.ComponentDefinition.Occurrences.ItemByName(oC.Name).CreateGeometryProxy(XYprox3, XYprox3)
                Try
                    flusCons1a = iamProyecto.ComponentDefinition.Constraints.AddFlushConstraint(XYwp1, XYprox3, altura)
                Catch ex As Exception
                    If Log Then PonLog("ERROR. Con restriccion FlusConstraint en plano XZ" & vbCrLf)
                End Try
            Else
                If Log Then PonLog("Es inclinada. Ponemos restriccion AngleConstraint en plano XZ para que no gire" & vbCrLf)
                XZprox2 = Me.DamePlano(oCHijo, 2, True)   ' 2 = XZ Plane
                iamProyecto.ComponentDefinition.Occurrences.ItemByName(oC.Name).CreateGeometryProxy(XZprox2, XZprox2)
                Try
                    angCons1 = iamProyecto.ComponentDefinition.Constraints.AddAngleConstraint(XZprox2, XYwp1, DameRad(90))   ', AngleConstraintSolutionTypeEnum.kDirectedSolution)
                Catch ex As Exception
                    Try
                        angCons1 = iamProyecto.ComponentDefinition.Constraints.AddAngleConstraint(XZprox2, XYwp1, -DameRad(90))  ', AngleConstraintSolutionTypeEnum.kDirectedSolution)
                    Catch ex1 As Exception
                        If Log Then PonLog("ERROR. Con restriccion AngleConstraint en plano XZ" & vbCrLf)
                        'MsgBox("Error con restriccion de angulo...")
                    End Try
                End Try

            End If
            If iamProyecto.RequiresUpdate Then iamProyecto.Update2()
            Try
                oC.Grounded = True
                'If log Then
                'PonLog("No Borramos las Restricciones en modo LOG y fijamos el componente" & vbCrLf)
                'Else
                flusCons1.Delete() : flusCons2.Delete()
                If flusCons1a IsNot Nothing Then flusCons1a.Delete()
                If angCons1 IsNot Nothing Then angCons1.Delete()
                'End If
            Catch ex As Exception
                If Log Then PonLog("ERROR. Fijando el componente o borranlo las restricciones" & vbCrLf)
            End Try
        Else
            If Log Then PonLog("No hace falta alinear. Ya está posicionado en su sitio" & vbCrLf)
        End If
        If oC.Grounded = False Then oC.Grounded = True
        Try
            'If log Then
            'og("No Borramos los puntos de trabajo usados para las restricciones en modo LOG" & vbCrLf)
            'Else
            pt1.Delete(True) : pt2.Delete(True)
            'End If
        Catch ex As Exception
            If Log Then PonLog("ERROR. Borrando los puntos de trabajo" & vbCrLf)
        End Try
        oAppI.ActiveView.Update()
        oAppI.ScreenUpdating = True
        If Log Then PonLog("Finalizado HoriontalRestringePuntos " & oC.Name & vbCrLf & StrDup(50, "*") & vbCrLf)
        'MsgBox("Ver todo")
    End Sub

    Public Sub HorizontalHazSimetriaGiro(ByVal oC1 As ComponentOccurrence, ByVal nBase As String, ByVal ptGiro As Point)
        oC1.Grounded = False
        Dim dPi As Double
        Dim ptOrigen As WorkPoint = Nothing
        Dim ptOrigenProx As Object = Nothing      ' Punto Origen (Proxy) (HORIZONTALES)
        Dim ptOtro As WorkPoint = Nothing        ' Punto contrario al origen (HORIZONTALES)
        Dim ptOtroProx As Object = Nothing        ' Punto contrario (Proxy) al origen (HORIZONTALES)
        Dim oA As AssemblyDocument = Nothing
        Dim oP As PartDocument = Nothing
        oP = oC1.SubOccurrences(1).ReferencedDocumentDescriptor.ReferencedDocument
        oA = oC1.ReferencedDocumentDescriptor.ReferencedDocument
        'ptBase.Parent.Document()

        'If nBase = "Centro" Then
        Try
            ptOrigen = oP.ComponentDefinition.WorkPoints("Centro")
            ptOtro = oP.ComponentDefinition.WorkPoints("Centro1")
        Catch ex As Exception
            oP = Nothing
            oA = Nothing
            Exit Sub
        End Try
        If Me.PropiedadLeeUsuario(CType(oA, Inventor.Document), "_BASE") = "Centro" Then
            Me.PropiedadEscribeUsuario(CType(oA, Inventor.Document), "_BASE", "Centro1", , False)
        Else
            Me.PropiedadEscribeUsuario(CType(oA, Inventor.Document), "_BASE", "Centro", , False)
        End If
        'Else
        'ptOrigen = oP.ComponentDefinition.WorkPoints("Centro1")
        'ptOtro = oP.ComponentDefinition.WorkPoints("Centro")
        'Me.PropiedadEscribeUsuario(oA, "_BASE", "Centro")
        'End If
        'Set oP = oA.ComponentDefinition.Occurrences.ItemByName(oC1.Name)
        'Call oC.CreateGeometryProxy(oP.ComponentDefinition.WorkPoints(1), ptOrigen)
        dPi = Math.Atan(1) * 4

        oTg = oAppI.TransientGeometry

        Dim oM As Matrix
        'Set oM = oTG.CreateMatrix
        oM = oC1.Transformation
        'Cos(dPi / 4) = 45 grados   Cos(dPi / 2) = 90 grados
        '' Cogemos el punto origen.
        Call oAppI.ActiveDocument.ComponentDefinition.Occurrences.ItemByName(oC1.Name).SubOccurrences(1).CreateGeometryProxy(ptOrigen, ptOrigenProx)
        Call oAppI.ActiveDocument.ComponentDefinition.Occurrences.ItemByName(oC1.Name).SubOccurrences(1).CreateGeometryProxy(ptOtro, ptOtroProx)

        oM.SetToRotation(((180 * dPi) / 180), oTg.CreateVector(0, 0, 1), ptOrigenProx.Point) ' ptOrigenProx.Point)    'ptGiro)    ' 

        'oC1.Transformation = oM  '.TransformBy oM  '.SetToRotation 0.5, oTG.CreateVector(0, 0, 1), ptBase.Point
        oM.PostMultiplyBy(oC1.Transformation)
        '' Hacermos el giro y actualizamos ensamblaje Proyecto
        oM.SetTranslation(oTg.CreateVector(ptOtroProx.Point.X, ptOtroProx.Point.Y, ptOtroProx.Point.Z))
        oC1.Transformation = oM
        oAppI.ActiveDocument.Update2()
        oC1.Grounded = True
    End Sub

    Public Sub HorizontalHazSimetriaRestriccion(ByVal oC1 As ComponentOccurrence, ByVal nBase As String)
        oC1.Grounded = False
        Dim dPi As Double
        Dim ptOrigen As WorkPoint = Nothing
        Dim ptOrigenProx As Object = Nothing      ' Punto Origen (Proxy) (HORIZONTALES)
        Dim ptOtro As WorkPoint = Nothing        ' Punto contrario al origen (HORIZONTALES)
        Dim ptOtroProx As Object = Nothing        ' Punto contrario (Proxy) al origen (HORIZONTALES)
        Dim oA As AssemblyDocument = Nothing
        Dim oP As PartDocument = Nothing
        oP = oC1.SubOccurrences(1).ReferencedDocumentDescriptor.ReferencedDocument
        oA = oC1.ReferencedDocumentDescriptor.ReferencedDocument
        'ptBase.Parent.Document()

        If nBase = "Centro" Then
            ptOrigen = oP.ComponentDefinition.WorkPoints("Centro")
            ptOtro = oP.ComponentDefinition.WorkPoints("Centro1")
            Me.PropiedadEscribeUsuario(CType(oA, Inventor.Document), "_BASE", "Centro1", , False)
        Else
            ptOrigen = oP.ComponentDefinition.WorkPoints("Centro1")
            ptOtro = oP.ComponentDefinition.WorkPoints("Centro")
            Me.PropiedadEscribeUsuario(CType(oA, Inventor.Document), "_BASE", "Centro", , False)
        End If
        'Set oP = oA.ComponentDefinition.Occurrences.ItemByName(oC1.Name)
        'Call oC.CreateGeometryProxy(oP.ComponentDefinition.WorkPoints(1), ptOrigen)
        Call oAppI.ActiveDocument.ComponentDefinition.Occurrences.ItemByName(oC1.Name).SubOccurrences(1).CreateGeometryProxy(ptOrigen, ptOrigenProx)
        Call oAppI.ActiveDocument.ComponentDefinition.Occurrences.ItemByName(oC1.Name).SubOccurrences(1).CreateGeometryProxy(ptOtro, ptOtroProx)
        dPi = Math.Atan(1) * 4

        oTg = oAppI.TransientGeometry

        Dim oM As Matrix
        'Set oM = oTG.CreateMatrix
        oM = oC1.Transformation
        'Cos(dPi / 4) = 45 grados   Cos(dPi / 2) = 90 grados
        oM.SetToRotation(((180 * dPi) / 180), oTg.CreateVector(0, 0, 1), ptOrigenProx.Point)

        'oC1.Transformation = oM  '.TransformBy oM  '.SetToRotation 0.5, oTG.CreateVector(0, 0, 1), ptBase.Point
        oM.PostMultiplyBy(oC1.Transformation)
        oM.SetTranslation(oTg.CreateVector(ptOtroProx.Point.X, ptOtroProx.Point.Y, ptOtroProx.Point.Z))
        oC1.Transformation = oM
        oAppI.ActiveDocument.Update2()
        oC1.Grounded = True
    End Sub


    Public Function CambiaElementoEnsamblaje(ByVal oE As AssemblyDocument, ByVal nOcu1 As String, ByVal fullOcu1 As String, ByVal bolTodas As Boolean, Optional borrarAntes As Boolean = False) As PartDocument
        Dim resultado As PartDocument = Nothing
        Dim oC As ComponentOccurrence = oE.ComponentDefinition.Occurrences.ItemByName(nOcu1)
        Dim dirEns As String = DameParteCamino(oE.FullFileName, IEnum.ParteCamino.CaminoSinFicheroBarra)
        '' dirPilares  ' Directorio donde están las plantillas de pilares.
        Dim caminoAntes As String = oC.ReferencedDocumentDescriptor.ReferencedFileDescriptor.FullFileName
        '' Si es el mismo no hacemos nada y salimos.
        If caminoAntes = fullOcu1 Then
            resultado = oC.Definition.Document
            CambiaElementoEnsamblaje = resultado
            Exit Function
        End If
        Dim dirPilAhora As String = DameParteCamino(oC.ReferencedDocumentDescriptor.ReferencedFileDescriptor.FullFileName, IEnum.ParteCamino.CaminoSinFicheroBarra)
        Dim nomPilAhora As String = DameParteCamino(oC.ReferencedDocumentDescriptor.ReferencedFileDescriptor.FullFileName, IEnum.ParteCamino.SoloFicheroConExtension)
        Dim dirPilBibli As String = DameParteCamino(fullOcu1, IEnum.ParteCamino.CaminoSinFicheroBarra)
        Dim nomPilBibli As String = DameParteCamino(fullOcu1, IEnum.ParteCamino.SoloFicheroConExtension)


        Try
            If caminoAntes <> (dirEns & nomPilBibli) Then '' Si no existe el pilar, lo copiaremos desde las plantillas (si existe)
                If Dir(dirEns & nomPilBibli) = "" And Dir(fullOcu1) <> "" Then
                    IO.File.Copy(fullOcu1, (dirEns & nomPilBibli))
                End If
                Call oC.Replace((dirEns & nomPilBibli), bolTodas)
                '' Borramos el anterior si borrarAntes = true
                If borrarAntes = True And Dir(caminoAntes) <> "" Then IO.File.Delete(caminoAntes)
                'If caminoAntes.StartsWith(dirEns) = True And Dir(caminoAntes) <> "" Then IO.File.Delete(caminoAntes)
            End If
            resultado = oC.ReferencedDocumentDescriptor.ReferencedDocument
        Catch ex As Exception
            '' No hacemos nada porque ha dado un error.
        End Try
        CambiaElementoEnsamblaje = resultado
    End Function

    Public Function MatrizEnsamblajeCambiaElemento(ByVal oE As AssemblyDocument, ByVal nMatriz As String, ByVal fullOcu1 As String) As Inventor.Document
        'oens.ComponentDefinition.Features.RectangularPatternFeatures.Item("PatrónEstribosTramo1:1")
        Dim resultado As Inventor.Document = Nothing
        Dim oRp As OccurrencePattern = Nothing
        Dim oC As ComponentOccurrence = Nothing
        Try
            oRp = oE.ComponentDefinition.OccurrencePatterns.Item(nMatriz)
            'oC = oE.ComponentDefinition.Occurrences.ItemByName(oRp.OccurrencePatternElements(1).Occurrences(1).Name)
            oC = MatrizEnsamblajeDameOcurrencia1(oE, nMatriz)
            If oC.ReferencedDocumentDescriptor.ReferencedFileDescriptor.FullFileName <> fullOcu1 Then
                oC.Replace(fullOcu1, True) 'oC.ReferencedDocumentDescriptor.ReferencedFileDescriptor.ReplaceReference(fullOcu1)   ', False)
                'oEns.Rebuild2()
                If oE.RequiresUpdate Then
                    oE.Update2()
                    oE.Save2()
                End If
            End If
            resultado = oC.ContextDefinition.Document 'oC.Definition.Document  ' .ReferencedDocumentDescriptor.ReferencedDocument
        Catch ex As Exception
            '' No hacemos nada porque ha dado un error.
            MsgBox("Error al cambiar el Documento de Matriz --> " & nMatriz)
        End Try
        MatrizEnsamblajeCambiaElemento = resultado
    End Function


    Public Function MatrizEnsamblajeDameOcurrencia1(ByVal oE As AssemblyDocument, ByVal nMatriz As String) As ComponentOccurrence
        'oens.ComponentDefinition.Features.RectangularPatternFeatures.Item("PatrónEstribosTramo1:1")
        Dim resultado As ComponentOccurrence = Nothing
        Dim oRp As OccurrencePattern = Nothing
        Try
            oRp = oE.ComponentDefinition.OccurrencePatterns.Item(nMatriz)
            resultado = oE.ComponentDefinition.Occurrences.ItemByName(oRp.OccurrencePatternElements(1).Occurrences(1).Name)
        Catch ex As Exception
            '' No hacemos nada porque ha dado un error.
        End Try
        MatrizEnsamblajeDameOcurrencia1 = resultado
    End Function

    Public Function ComponentOccurrenceDame(ByVal queEns As AssemblyDocument, ByVal queCaminoCompleto As String) As ComponentOccurrence
        Dim oCo As ComponentOccurrence = Nothing
        Dim soloNombre As String = IO.Path.GetFileNameWithoutExtension(queCaminoCompleto)
        For Each oCo In queEns.ComponentDefinition.Occurrences
            If oCo.ReferencedDocumentDescriptor.ReferencedFileDescriptor.FullFileName.ToLower = queCaminoCompleto.ToLower Or
            oCo.Name.ToLower.StartsWith(soloNombre.ToLower & ":") Then
                Exit For
            End If
        Next
        ComponentOccurrenceDame = oCo
        Exit Function
    End Function
    '
    Public Function ComponentOccurrenceDamePrimero(ByVal queEns As AssemblyDocument, ByVal enumType As Inventor.DocumentTypeEnum, txtType As String) As ComponentOccurrence
        Dim resultado As ComponentOccurrence = Nothing
        For Each oCo As ComponentOccurrence In queEns.ComponentDefinition.Occurrences
            Dim queType As String = PropiedadLeeUsuario(CType(oCo.ReferencedDocumentDescriptor.ReferencedDocument, Inventor.Document), "TYPE")
            If oCo.ReferencedDocumentDescriptor.ReferencedDocumentType = enumType And queType.ToUpper.Equals(txtType.ToUpper) Then
                resultado = queEns.ComponentDefinition.Occurrences.ItemByName(oCo.Name)
                Exit For
            Else
                If oCo.SubOccurrences IsNot Nothing AndAlso oCo.SubOccurrences.Count > 0 Then
                    For Each oCo1 As ComponentOccurrence In oCo.SubOccurrences
                        If oCo1.ReferencedDocumentDescriptor.ReferencedDocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                            resultado = ComponentOccurrenceDamePrimeroRecursivo(queEns, enumType, txtType)
                            If resultado IsNot Nothing Then Exit For
                        End If
                    Next
                    If resultado IsNot Nothing Then Exit For
                Else
                    resultado = Nothing
                End If
            End If
        Next
        ''
        Return resultado
        Exit Function
    End Function
    ''

    Public Function ComponentOccurrenceDamePrimeroRecursivo(ByVal queEns As AssemblyDocument, ByVal enumType As Inventor.DocumentTypeEnum, txtType As String) As ComponentOccurrence
        Dim resultado As ComponentOccurrence = Nothing
        For Each oCo As ComponentOccurrence In queEns.ComponentDefinition.Occurrences
            Dim queType As String = PropiedadLeeUsuario(CType(oCo.ReferencedDocumentDescriptor.ReferencedDocument, Inventor.Document), "TYPE")
            If oCo.ReferencedDocumentDescriptor.ReferencedDocumentType = enumType And queType.ToUpper.Equals(txtType.ToUpper) Then
                resultado = queEns.ComponentDefinition.Occurrences.ItemByName(oCo.Name)
                Exit For
            Else
                If oCo.SubOccurrences IsNot Nothing AndAlso oCo.SubOccurrences.Count > 0 Then
                    For Each oCo1 As ComponentOccurrence In oCo.SubOccurrences
                        If oCo1.ReferencedDocumentDescriptor.ReferencedDocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                            resultado = ComponentOccurrenceDamePrimeroRecursivo(queEns, enumType, txtType)
                            If resultado IsNot Nothing Then Exit For
                        End If
                    Next
                    If resultado IsNot Nothing Then Exit For
                Else
                    resultado = Nothing
                End If
            End If
        Next
        ''
        Return resultado
        Exit Function
    End Function

    Public Sub ListaMaterialesRenumera(ByVal queEns As AssemblyDocument)
        '' Renumeramos los elementos de la lista de piezas. Vista "Estructurado"
        Dim oB As BOM = queEns.ComponentDefinition.BOM
        Dim oBv As BOMView = Nothing
        oB.StructuredViewEnabled = True
        oB.PartsOnlyViewEnabled = True

        oBv = oB.BOMViews.Item(2)   '(2)("Estructurado")("Structured") '(3)("Parts Only")("Solo piezas")  * (Sólo piezas) en 2011 y (Solo piezas) en 2012
        For x As Integer = 1 To oBv.BOMRows.Count
            Dim oBr As BOMRow = oBv.BOMRows.Item(x)
            oBr.ItemNumberLocked = False
            oBr.ItemNumber = x.ToString
            oBr.ItemNumberLocked = True
        Next
        oBv = Nothing
        '' Renumeramos los elementos de la lista de piezas. Vista "Solo piezas"
        oBv = oB.BOMViews.Item(3)   '("Solo piezas")   '("Structured") ("Estructurado") '("Parts Only") ("Sólo piezas")  
        For x As Integer = 1 To oBv.BOMRows.Count
            Dim oBr As BOMRow = oBv.BOMRows.Item(x)
            oBr.ItemNumberLocked = False
            oBr.ItemNumber = x.ToString
            oBr.ItemNumberLocked = True
        Next
    End Sub
#Region "FileReferences"
#Region "FileReferencesDameTodasInventor"


    ''' <summary>
    ''' TODO, recursivamente. Ficheros de Inventor y no de Inventor vinculados o insertados, objetos OLE, etc.
    ''' también los IPN y los planos IDW y DWG (Busquedabasica)
    ''' </summary>
    ''' <param name="queFichero">FullFilename del fichero (si oD=Nothing)</param>
    ''' <param name="bolVisible">True/False para que este Visible o No</param>
    ''' <param name="tambienIDW">True : También buscaremos planos IDW y DWG</param>
    ''' <param name="Busquedabasica">True : Busqueda sólo nombre o Abrir y ver la Vista Base</param>
    ''' <returns>Devuelve un ArrayList con los FullFilename de los ficheros a copiar</returns>
    ''' <remarks>Mejor indicar el IDW o DWG padre de todos los demás. Así también localizará los IPN</remarks>
    Public Function FileReferencesDameTodasInventor(ByVal queFichero As String,
                                                ByVal bolVisible As Boolean,
                                                ByVal tambienIDW As Boolean,
                                                ByVal Busquedabasica As Boolean,
                                                Optional tambienPadre As Boolean = True) As ArrayList
        Dim resultado As New ArrayList
        Dim estabaabierto As Boolean '= True
        Dim oDoc As Inventor.Document = Nothing '= oD
        If queFichero = "" Then
            MsgBox("Error : Debe indicar el FullFilename de un fichero")
            Return resultado
            'Exit Function
        End If

        If queFichero <> "" And IO.File.Exists(queFichero) Then
            oAppI.SilentOperation = True
            estabaabierto = Me.FicheroAbierto(queFichero)
            If estabaabierto = False Then
                oDoc = oAppI.Documents.Open(queFichero, bolVisible)
            Else
                oDoc = oAppI.Documents.ItemByName(queFichero)
            End If
            Try
                oDoc.Update2(True)
            Catch ex As Exception
                '' No se ha podido actualizar
            End Try
            If oDoc.Dirty = True Then oDoc.Save2(True)
            '' Si es un ensamblaje, activaremos la representacion "Principal"
            If oDoc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                RepresentacionActivaCrea(CType(oDoc, AssemblyDocument), False, "")  '' Poner activa "Principal" (Master)
            End If
            oAppI.SilentOperation = False
        End If

        If oDoc Is Nothing Then
            resultado = Nothing
            Return resultado
            Exit Function
        End If

        Dim caminoFull As String = oDoc.FullFileName
        If tambienPadre = True AndAlso resultado.Contains(caminoFull) = False Then resultado.Add(caminoFull)

        If caminoFull.ToLower.EndsWith(".iam") Or caminoFull.ToLower.EndsWith(".ipt") Or caminoFull.ToLower.EndsWith(".ipn") _
    And tambienIDW = True Then
            '' **** Buscamos también si tiene plano IDW o DWG para añadirlo a la colección.
            Dim planos As ArrayList = Nothing
            If Busquedabasica = True Then
                planos = ExisteFicheroPlanoEnDirBasico(caminoFull)
            Else
                planos = ExisteFicheroPlanoEnDirProfundoInv(caminoFull)
            End If

            If planos IsNot Nothing AndAlso planos.Count > 0 Then
                For Each queF As String In planos
                    If resultado.Contains(queF) = False Then resultado.Add(queF)
                Next
            End If
            planos = Nothing
            '' ***************************************************************************
        Else
            '' No hacemos nada, ya que es un plano DWG o IDW
        End If

        Call FileReferencesDameTodasInventorRecursivo(oDoc.File, resultado, tambienIDW, Busquedabasica)

        If estabaabierto = False Then oDoc.Close(True)
        oDoc = Nothing

        GC.WaitForPendingFinalizers()
        GC.Collect()
        '' Finalmente buscamos dentro de las referencias para ver si
        '' Las piezas, ensamblajes o presentaciones tienen un plano (IDW o DWG no de Inventor) que se llame igual
        '' En sus respectivos directorios. Lo añadiremos, si no existe
        If resultado IsNot Nothing And tambienIDW = True Then
            For Each queF As String In resultado
                If queF.ToLower.EndsWith(".iam") Or queF.ToLower.EndsWith(".ipt") Or
             queF.ToLower.EndsWith(".ipn") Then
                    Dim planoIDW As String = DameParteCamino(queF, IEnum.ParteCamino.SoloCambiaExtension, ".idw")
                    Dim planoDWG As String = DameParteCamino(queF, IEnum.ParteCamino.SoloCambiaExtension, ".dwg")
                    If IO.File.Exists(planoIDW) AndAlso resultado.Contains(planoIDW) = False Then _
                resultado.Add(planoIDW)
                    If IO.File.Exists(planoDWG) AndAlso oAppI.FileManager.IsInventorDWG(planoDWG) = True AndAlso resultado.Contains(planoDWG) = False Then _
                resultado.Add(planoDWG)
                End If
            Next
        End If

        GC.WaitForPendingFinalizers()
        GC.Collect()
        Return resultado
    End Function

    Private Sub FileReferencesDameTodasInventorRecursivo(ByVal oFile As Inventor.File,
                                                     ByRef resultado As ArrayList,
                                                     ByVal tambienIDW As Boolean,
                                                     ByVal Busquedabasica As Boolean)
        Dim oFileDescriptor As FileDescriptor
        For Each oFileDescriptor In oFile.ReferencedFileDescriptors
            'If oFileDescriptor.FullFileName.Contains("OldVersions") = True Then Continue For
            If oFileDescriptor.LocationType = LocationTypeEnum.kLibraryLocation Then Continue For
            '' Si no está desaparecido el componente (REferenceMising)
            If Not oFileDescriptor.ReferenceMissing Then
                Dim caminoFull As String = oFileDescriptor.FullFileName
                '' Si no existe el camino completo, lo añadimos. Añadimos cualquier fichero vinculado (menos librerías)
                If Not resultado.Contains(caminoFull) Then resultado.Add(caminoFull)

                '' Incluiremos también los planos DWG e IDW, sólo si es un ensamblaje o una pieza.
                If caminoFull.ToLower.EndsWith(".iam") Or caminoFull.ToLower.EndsWith(".ipt") _
            Or caminoFull.ToLower.EndsWith(".ipn") And tambienIDW = True Then

                    '' **** Buscamos también si tiene plano IDW o DWG para añadirlo a la colección.
                    '' planos ya tiene quitados (los de AutoCAD, los de OldVersions, etc.)
                    Dim planos As ArrayList = Nothing
                    If Busquedabasica = True Then
                        planos = ExisteFicheroPlanoEnDirBasico(caminoFull)
                    Else
                        planos = ExisteFicheroPlanoEnDirProfundoInv(caminoFull)
                    End If

                    If planos IsNot Nothing AndAlso planos.Count > 0 Then
                        For Each queF As String In planos
                            If resultado.Contains(queF) = False Then resultado.Add(queF)
                        Next
                    End If
                    planos = Nothing
                    '' ***************************************************************************
                End If

                ' Si no es un fichero No de Inventor (xls, jpg, etc.) kForeignFileType, procesa recursivamente.
                If Not oFileDescriptor.ReferencedFileType = FileTypeEnum.kForeignFileType Then
                    Call FileReferencesDameTodasInventorRecursivo(oFileDescriptor.ReferencedFile, resultado, tambienIDW, Busquedabasica)
                End If
            End If
        Next
    End Sub
#End Region
    ''' <summary>
    ''' Cambia el fichero referenciado (camino completo) por otro con Inventor
    ''' </summary>
    ''' <param name="listaNombres">Hashtable con Key=NombreViejo / Value=NombreNuevo</param>
    ''' <param name="queFichero">Camino completo del fichero a Cambiar</param>
    ''' <returns>Retorna una cadena con los ficheros cambiados viejo/nuevo</returns>
    ''' <remarks>Solo procesamos el fichero indicado en "queFichero"</remarks>
    Public Function FileReferenciaCambiaUnoInventor(ByVal listaNombres As Hashtable, ByVal queFichero As String) As String
        Dim resultado As String = ""
        Dim cabecera As String = ""
        Dim cambiados As String = ""
        Dim oDoc As Inventor.Document = Nothing
        Dim oEns As AssemblyDocument = Nothing
        Dim oEnsHijo As AssemblyDocument = Nothing
        Dim oDocDes As DocumentDescriptorsEnumerator = Nothing

        If Dir(queFichero) <> "" Then
            oAppI.SilentOperation = True
            ' Abrir el Documento en Inventor.
            oDoc = Me.oAppI.Documents.Open(queFichero, False) ' Visible, quitarlo después de pruebas.
            oDocDes = oDoc.ReferencedDocumentDescriptors
            If oDoc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                oEns = oDoc ' oAppCls.Documents.ItemByName(oDoc.FullDocumentName)
                Me.RepresentacionActivaCrea(oEns, False)    ' Activa Representación "Principal"
            End If

            cabecera &= StrDup(80, "*") & vbCrLf
            cabecera &= "Nombre   : " & oDoc.DisplayName & vbCrLf
            cabecera &= "Completo : " & oDoc.FullFileName & vbCrLf
            cabecera &= StrDup(80, "-") & vbCrLf
            For Each fileRef As Inventor.DocumentDescriptor In oDoc.ReferencedDocumentDescriptors
                '' Si el nombre viejo completo está en la colección "listaNombres"
                '' Creamos el mensaje del cambio viejo/nuevo y lo cambiamos.
                Dim fD As FileDescriptor = fileRef.ReferencedFileDescriptor

                If listaNombres.ContainsKey(fD.FullFileName) AndAlso IO.File.Exists(listaNombres(fD.FullFileName)) Then
                    ''resultado &= StrDup(80, "*") & vbCrLf
                    cambiados &= "Viejo : " & fD.FullFileName & vbCrLf
                    fD.ReplaceReference(listaNombres(fD.FullFileName))
                    ''oApprentice.FileManager.RefreshAllDocuments()
                    cambiados &= "Nuevo : " & fD.FullFileName & vbCrLf  ' listaNombres(strRef.FullFileName) & vbCrLf
                    If fD.ReferencedFileType = FileTypeEnum.kAssemblyFileType Then
                        oEnsHijo = oAppI.Documents.ItemByName(fileRef.FullDocumentName)
                        Me.RepresentacionActivaCrea(oEnsHijo, True, nivelDetalleDefecto)    ' Activa Representación "Desactivados"
                    End If
                End If
                oEnsHijo = Nothing
                Me.PropiedadEscribe(fileRef.ReferencedDocument, "Nº de pieza", DameParteCamino(fileRef.ReferencedFileDescriptor.FullFileName, IEnum.ParteCamino.SoloFicheroSinExtension).ToUpper)
                MyClass.oAppI.UserInterfaceManager.DoEvents()

            Next
            Me.PropiedadEscribe(oDoc, "Nº de pieza", DameParteCamino(oDoc.FullFileName, IEnum.ParteCamino.SoloFicheroSinExtension).ToUpper)

            Try
                If oDoc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                    oEns = oDoc ' oAppCls.Documents.ItemByName(oDoc.FullDocumentName)
                    Me.RepresentacionActivaCrea(oEns, True, nivelDetalleDefecto)    ' Activa Representación "Desactivados"
                End If
                oEns = Nothing

                If oDoc IsNot Nothing Then
                    If (oDoc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Or
                    oDoc.DocumentType = DocumentTypeEnum.kPartDocumentObject) AndAlso
                    oDoc.RequiresUpdate Then oDoc.Update2()
                    oDoc.Save2()
                    oDoc.Close()
                End If
            Catch ex As Exception
                '' No hacemos nada. El fichero ya estaba cerrado.
            End Try
            oDoc = Nothing
        End If

        oAppI.SilentOperation = False
        ''oApprentice.Close()
        'If oDoc.Dirty = False Then oDoc.
        'oDoc.Close()
        If cambiados <> "" Then
            resultado = cabecera & cambiados
        Else
            resultado = "NO EXISTE--> " & queFichero
        End If

        Return resultado
    End Function

    Public Function FileReferenciaDameApprentice(ByVal queDoc As String) As String
        Dim contador As Integer = 1
        Dim mensaje As String = ""

        If Dir(queDoc) = "" Then GoTo Fin
        ' Create a new instance of Apprentice.
        Dim oApprentice As New ApprenticeServerComponent
        ' Open a document.
        Dim oDoc As ApprenticeServerDocument
        oDoc = oApprentice.Open(queDoc)
        mensaje &= StrDup(80, "*") & vbCrLf
        mensaje &= "Nombre   : " & oDoc.DisplayName & vbCrLf
        mensaje &= "Completo : " & oDoc.FullFileName & vbCrLf
        mensaje &= StrDup(80, "-") & vbCrLf
        If oDoc.File.AllReferencedFiles.Count > 0 Then
            For Each strRef As Inventor.File In oDoc.File.AllReferencedFiles
                mensaje &= contador & ".- " & strRef.FullFileName & vbCrLf
                contador += 1
            Next
            'Debug.Print(mensaje & vbCrLf)
        Else
            mensaje &= "NO TIENE REFERENCIAS" & vbCrLf
        End If
        oDoc.Close()
        mensaje &= StrDup(80, "*") & vbCrLf

Fin:
        FileReferenciaDameApprentice = mensaje
        Exit Function
    End Function


    '' Solo ficheros Inventor. AllReferencedFiles
    Public Function FileReferenciasInventorTodasApp(ByVal queF As String) As ArrayList
        If IO.File.Exists(queF) = False Then Return Nothing

        Dim resultado As New ArrayList
        resultado.Add(queF)
        ' Create a new instance of Apprentice.
        Dim oApprentice As New ApprenticeServerComponent

        ' Open a document.
        Dim oDoc As ApprenticeServerDocument

        oDoc = oApprentice.Open(queF)

        ' Check to make sure the document is an assembly.
        For Each oFi As Inventor.File In oDoc.File.AllReferencedFiles
            If Not resultado.Contains(oFi.FullFileName) Then resultado.Add(oFi.FullFileName)
        Next
        oApprentice.Close()
        Return resultado
    End Function

    '' Cambiaremos todas las referencias antigüas, por las nuevas. Abriendo el Padre principal con Inventor.
    '' Indicamos Fullname del fichero (queFichero) y Hastable con Fullnames Key=Viejo,Value=Nuevo (queVieNue)
    Public Sub FileReferenciaCambiaTodoInventorPartNumberNoEnNumericos(ByVal queFichero As String,
                                                ByVal queVieNue As Hashtable,
                                                Optional ByVal recursivo As Boolean = True)
        Dim oDoc1 As Inventor.Document = Nothing
        Dim oEns1 As Inventor.AssemblyDocument = Nothing
        Dim oPie1 As Inventor.PartDocument = Nothing
        Dim oIdw1 As Inventor.DrawingDocument = Nothing
        Dim estabaabierto As Boolean = False

        oAppI.SilentOperation = True
        '' Llenamos el objeto Documento (oDoc) con el documento abierto o lo abrimos, si estaba cerrado.
        If Me.FicheroAbierto(queFichero) = True Then
            oDoc1 = oAppI.Documents.ItemByName(queFichero)
            'oDoc1.Activate()
            estabaabierto = True
        Else
            oDoc1 = Me.oAppI.Documents.Open(queFichero, False) ' Visible, quitarlo después de pruebas.
            estabaabierto = False
        End If
        '
        ' Cambiar Part Number (Solo si no es numérico [Centro de Contenido])
        Dim soloNombre As String = IO.Path.GetFileNameWithoutExtension(queFichero)
        If IsNumeric(soloNombre) = True Then
            Exit Sub
        End If
        '
        Me.PropiedadEscribeDesignTracking(Nothing, "Part Number", soloNombre, queFichero)
        '
        If oDoc1.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
            oEns1 = oAppI.Documents.ItemByName(queFichero) 'oDoc ' oAppCls.Documents.ItemByName(oDoc.FullDocumentName)
            Me.RepresentacionActivaCrea(oEns1, False)    ' Activa Representación "Principal"
            'oEns.Update2() : oEns.Save2()
            Call FileReferenciaCambiaTodoInventorRecursivo(oEns1.File, queVieNue, recursivo)
            'oEns1.Rebuild2(True)
            If oEns1.RequiresUpdate Then oEns1.Update2(True)
            oEns1.Save2(True)
            'Me.RepresentacionActivaCrea(oEns1, True, nivelDetalleDefecto)    ' Activa Representación "Desactivados"
            'oEns1.Save2(True)
        ElseIf oDoc1.DocumentType = DocumentTypeEnum.kDrawingDocumentObject Then
            oIdw1 = oAppI.Documents.ItemByName(queFichero) 'oDoc
            If oIdw1.FullFileName.EndsWith(".dwg") AndAlso oIdw1.IsInventorDWG = False Then Exit Sub
            Call FileReferenciaCambiaTodoInventorRecursivo(oIdw1.File, queVieNue, recursivo)
            oIdw1.Update2(True) : oIdw1.Save2(True)
        ElseIf oDoc1.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
            oPie1 = oAppI.Documents.ItemByName(queFichero) 'oDoc
            Call FileReferenciaCambiaTodoInventorRecursivo(oPie1.File, queVieNue, recursivo)
            'oPie1.Rebuild2(True)
            If oPie1.RequiresUpdate Then oPie1.Update2(True)
            oPie1.Save2(True)
        End If

        '' Por si da error RequiresUpdate o Update2
        Try
            If oEns1 IsNot Nothing And estabaabierto = False Then
                oEns1.Close(True)
            ElseIf oIdw1 IsNot Nothing And estabaabierto = False Then
                oIdw1.Close(True)
            ElseIf oPie1 IsNot Nothing And estabaabierto = False Then
                oPie1.Close(True)
            End If
        Catch ex As Exception
            'oDoc.Save2(False)
            Debug.Print("Error al actualizar/guardar")
        End Try

        oAppI.SilentOperation = False
        oEns1 = Nothing
        oIdw1 = Nothing
        oPie1 = Nothing
        oDoc1 = Nothing
    End Sub

    '' Cambiaremos todas las referencias antigüas, por las nuevas. Abriendo el Padre principal con Inventor.
    '' Indicamos Fullname del fichero (queFichero) y Hastable con Fullnames Key=Viejo,Value=Nuevo (queVieNue)
    Public Sub FileReferenciaCambiaTodoInventor(ByVal queFichero As String,
                                                ByVal queVieNue As Hashtable,
                                                Optional ByVal recursivo As Boolean = True)
        Dim oDoc1 As Inventor.Document = Nothing
        Dim oEns1 As Inventor.AssemblyDocument = Nothing
        Dim oPie1 As Inventor.PartDocument = Nothing
        Dim oIdw1 As Inventor.DrawingDocument = Nothing
        Dim estabaabierto As Boolean = False

        oAppI.SilentOperation = True
        '' Llenamos el objeto Documento (oDoc) con el documento abierto o lo abrimos, si estaba cerrado.
        If Me.FicheroAbierto(queFichero) = True Then
            oDoc1 = oAppI.Documents.ItemByName(queFichero)
            'oDoc1.Activate()
            estabaabierto = True
        Else
            oDoc1 = Me.oAppI.Documents.Open(queFichero, False) ' Visible, quitarlo después de pruebas.
            estabaabierto = False
        End If
        '
        ' Cambiar Part Number (Solo si no es numérico [Centro de Contenido])
        Dim soloNombre As String = IO.Path.GetFileNameWithoutExtension(queFichero)
        Me.PropiedadEscribeDesignTracking(Nothing, "Part Number", soloNombre, queFichero)
        '
        If oDoc1.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
            oEns1 = oAppI.Documents.ItemByName(queFichero) 'oDoc ' oAppCls.Documents.ItemByName(oDoc.FullDocumentName)
            Me.RepresentacionActivaCrea(oEns1, False)    ' Activa Representación "Principal"
            'oEns.Update2() : oEns.Save2()
            Call FileReferenciaCambiaTodoInventorRecursivo(oEns1.File, queVieNue, recursivo)
            'oEns1.Rebuild2(True)
            If oEns1.RequiresUpdate Then oEns1.Update2(True)
            oEns1.Save2(True)
            'Me.RepresentacionActivaCrea(oEns1, True, nivelDetalleDefecto)    ' Activa Representación "Desactivados"
            'oEns1.Save2(True)
        ElseIf oDoc1.DocumentType = DocumentTypeEnum.kDrawingDocumentObject Then
            oIdw1 = oAppI.Documents.ItemByName(queFichero) 'oDoc
            If oIdw1.FullFileName.EndsWith(".dwg") AndAlso oIdw1.IsInventorDWG = False Then Exit Sub
            Call FileReferenciaCambiaTodoInventorRecursivo(oIdw1.File, queVieNue, recursivo)
            oIdw1.Update2(True) : oIdw1.Save2(True)
        ElseIf oDoc1.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
            oPie1 = oAppI.Documents.ItemByName(queFichero) 'oDoc
            Call FileReferenciaCambiaTodoInventorRecursivo(oPie1.File, queVieNue, recursivo)
            'oPie1.Rebuild2(True)
            If oPie1.RequiresUpdate Then oPie1.Update2(True)
            oPie1.Save2(True)
        End If

        '' Por si da error RequiresUpdate o Update2
        Try
            If oEns1 IsNot Nothing And estabaabierto = False Then
                oEns1.Close(True)
            ElseIf oIdw1 IsNot Nothing And estabaabierto = False Then
                oIdw1.Close(True)
            ElseIf oPie1 IsNot Nothing And estabaabierto = False Then
                oPie1.Close(True)
            End If
        Catch ex As Exception
            'oDoc.Save2(False)
            Debug.Print("Error al actualizar/guardar")
        End Try

        oAppI.SilentOperation = False
        oEns1 = Nothing
        oIdw1 = Nothing
        oPie1 = Nothing
        oDoc1 = Nothing
    End Sub

    Private Sub FileReferenciaCambiaTodoInventorRecursivo(ByVal oFile As Inventor.File, ByVal queVieNue As Hashtable, Optional ByVal recursivo As Boolean = True)
        ' Cambiar Part Number
        Me.PropiedadEscribeDesignTracking(Nothing, "Part Number", IO.Path.GetFileNameWithoutExtension(oFile.FullFileName), oFile.FullFileName)
        '
        For Each oFD As FileDescriptor In oFile.ReferencedFileDescriptors
            '' Continuaremos si: 
            '' - Componentes fijos. No hay que cambiarlos, son de biblioteca.
            '' - Ya está cambiado
            '' - No está en la queVieNue
            '' Componente fijo.
            If oFD.FullFileName.Contains("\ARMADOS_FIJOS\") Then Continue For
            '' Si no estan renombrados pasamos al siguiente.
            If oFD.FullFileName.Contains("newVer.") Then Continue For
            '' Si no esta en la colección, mostramos error y pasamos a la siguiente
            If queVieNue.ContainsKey(oFD.FullFileName) = False Then
                If Log Then PonLog("NO EN COLECCIÓN Key=" & oFD.FullFileName)
                Continue For
            Else
                If oFD.FullFileName = queVieNue(oFD.FullFileName) Then
                    '' Si ya está cambiado, continuamos con el siguiente.
                    If Log Then PonLog("YA CAMBIADO " & queVieNue(oFD.FullFileName))
                    Continue For
                End If
            End If
            '' Si no existe destino, mostramos error y pasamos a la siguiente
            If IO.File.Exists(queVieNue(oFD.FullFileName)) = False Then
                If Log Then PonLog("NO EXISTE " & queVieNue(oFD.FullFileName))
                Continue For
            End If
            ''
            '' Si no es un referencia desaparecida...
            ''
            If Not oFD.ReferenceMissing Then
                If oFD.LocationType = LocationTypeEnum.kLibraryLocation Then Continue For
                '' Cambiamos la referencia si la plantilla inicial esta y si existe el fichero final.
                If queVieNue.ContainsKey(oFD.FullFileName) AndAlso IO.File.Exists(queVieNue(oFD.FullFileName)) Then
                    oFD.ReplaceReference(queVieNue(oFD.FullFileName))
                End If

                Me.DoEventsInventor(True)
                ' Recursivamente salvo que sea un kForeignFileType (Otros documentos NO de Inventor)
                If Not oFD.ReferencedFileType = FileTypeEnum.kForeignFileType AndAlso recursivo = True Then
                    Call FileReferenciaCambiaTodoInventorRecursivo(oFD.ReferencedFile, queVieNue)
                End If
            End If
        Next
    End Sub

    Public Function FileReferencesTODOaFichero(Optional ByVal oDoc As Inventor.Document = Nothing, Optional ByVal queFichero As String = "") As String
        Dim estabaabierto As Boolean = False
        If oDoc Is Nothing And queFichero = "" Then
            MsgBox("Debe indicar on Objeto Documento abierto o un camino completo a un fichero de Inventor...")
            Return ""
        End If

        If oDoc Is Nothing And queFichero <> "" Then
            If IO.File.Exists(queFichero) = False Then
                MsgBox("No existe " & queFichero)
                Return ""
            Else
                estabaabierto = FicheroAbierto(queFichero)
                If estabaabierto = True Then
                    oDoc = Me.oAppI.Documents.ItemByName(queFichero)
                Else
                    Me.oAppI.SilentOperation = True
                    oDoc = Me.oAppI.Documents.Open(queFichero, True)
                End If
            End If
        ElseIf oDoc IsNot Nothing Then
            estabaabierto = True
        End If


        Dim queF As String = My.Application.Info.DirectoryPath & "\Ficheros.txt"
        If IO.File.Exists(queF) Then IO.File.Delete(queF)

        FileOpen(1, queF, OpenMode.Append, OpenAccess.Write)
        WriteLine(1, "Referencias de : " & oDoc.FullFileName & vbNewLine & vbNewLine)

        Dim arrF As New ArrayList
        Dim oFi As Inventor.FileDescriptor
        For Each oFi In oDoc.File.ReferencedFileDescriptors
            If arrF.Contains(oFi.FullFileName) = False Then
                arrF.Add(oFi.FullFileName)
                'Debug.Print(oFi.FullFileName)
                WriteLine(1, oFi.FullFileName)
                Try
                    Dim oFds As FileDescriptorsEnumerator = oFi.ReferencedFile.ReferencedFileDescriptors
                    For Each oFi1 As FileDescriptor In oFds
                        If arrF.Contains(oFi1.FullFileName) = False Then
                            arrF.Add(oFi1.FullFileName)
                            'Debug.Print(oFi.FullFileName)
                            WriteLine(1, oFi1.FullFileName)
                        End If
                    Next
                Catch ex As Exception
                    Continue For
                End Try
            End If
        Next
        FileClose(1)

        If estabaabierto = False Then oDoc.Close(True)
        Me.oAppI.SilentOperation = False

        Return queF
    End Function

#End Region

    Public Sub AbreActualizaGuarda(ByVal queFichero As String, Optional ByVal elIDW As Boolean = False, Optional ByVal elDWG As Boolean = False)
        'Dim estaVisible As Boolean = False
        If queFichero.Contains("OldVersions") Then Exit Sub
        Dim procesado As Integer = 0
TambienDibujo:
        Dim oDoc As Inventor.Document = Nothing

        Try
            oAppI.SilentOperation = True
            For Each oDoch In oAppI.Documents
                If oDoch.FullFileName = queFichero Then
                    oDoc = oDoch
                    Exit For
                End If
            Next

            oAppI.SilentOperation = True
            If Dir(queFichero) = "" Then Exit Sub
            If (procesado = 2) AndAlso oAppI.FileManager.IsInventorDWG(queFichero) = False Then Exit Sub
            If oDoc Is Nothing Then
                oDoc = oAppI.Documents.Open(queFichero, False)
            End If
            '' Si es un dibujo DWG y NO es de Inventor, salimos.
            If oDoc.DocumentType = DocumentTypeEnum.kDrawingDocumentObject AndAlso CType(oDoc, DrawingDocument).IsInventorDWG = False Then Exit Sub

            If procesado = 0 Then   ' Si es el ensamblaje. Ponemos Representación nivelDetalleDefecto como activa.
                'Me.RepresentacionActivaCrea(CType(oDoc, AssemblyDocument), True, nivelDetalleDefecto)
                Me.RepresentacionActivaCrea(CType(oDoc, AssemblyDocument), False)
                oDoc.Update2(True)
            End If
            oDoc.Save2(True)
            Try
                'oDoc.ReleaseReference()
                oDoc.Close(True)
            Catch ex As Exception
                ' No hacemos nada
                Debug.Print("Error en AbreActualizaGuarda..." & vbCrLf & ex.Message)
            End Try
            'If oDoc.RequiresUpdate Then
            'Else
            '' No actualizamos ni guardamos. Cerraremos el fichero sin guardar.
            'End If
            oDoc = Nothing

            If elIDW = True And procesado = 0 Then
                '' ***** Para sacar el plano IDW que tenga
                Dim planos As ArrayList
                If Busquedabasica = True Then
                    planos = ExisteFicheroPlanoEnDirBasico(queFichero)
                Else
                    planos = ExisteFicheroPlanoEnDirProfundoInv(queFichero)
                End If
                Dim planoIDW As String = ""
                If planos IsNot Nothing AndAlso planos.Count > 0 Then
                    planoIDW = ExisteFicheroPlanoEnArray(planos, IO.Path.ChangeExtension(queFichero, ".idw"))
                End If
                '' *************************************************
                procesado = 1
                GoTo TambienDibujo
                '' *****************************************************************************
            ElseIf elDWG = True And procesado = 1 Then
                '' ***** Para sacar el plano DWG que tenga
                Dim planos As ArrayList
                If Busquedabasica = True Then
                    planos = ExisteFicheroPlanoEnDirBasico(queFichero)
                Else
                    planos = ExisteFicheroPlanoEnDirProfundoInv(queFichero)
                End If
                Dim planoDWG As String = ""
                If planos IsNot Nothing AndAlso planos.Count > 0 Then
                    planoDWG = ExisteFicheroPlanoEnArray(planos, IO.Path.ChangeExtension(queFichero, ".dwg"))
                End If
                '' *************************************************
            End If

            Me.oAppI.SilentOperation = False
        Catch ex As Exception
            Debug.Print("ALBERTO-->Error en clsInventor-->AbreActualizaGuarda" & vbCrLf & ex.Message)
            oAppI.SilentOperation = False
        End Try
    End Sub


    Public Sub ActualizaGuardaDibujo(ByRef oDoc As Inventor.DrawingDocument)
        'Dim estaVisible As Boolean = False
        Try
            Dim oS As Sheet
            Dim oV As DrawingView
            Dim procesados As New ArrayList
            For Each oS In oDoc.Sheets
                For Each oV In oS.DrawingViews
                    oAppI.SilentOperation = True
                    Try
                        Dim oD As Inventor.Document = oV.ReferencedDocumentDescriptor.ReferencedDocument
                        If oD IsNot Nothing AndAlso procesados.Contains(oD.FullFileName) = False Then
                            oD.Update2() : oD.Save2()
                            procesados.Add(oD.FullFileName)
                        End If
                    Catch ex As Exception
                        Continue For
                    End Try
                    Me.oAppI.SilentOperation = False
                Next oV
            Next oS
            oDoc.Save2()
        Catch ex As Exception
            'MsgBox("ALBERTO-->Error en clsInventor-->ActualizaGuardaDibujo" & vbCrLf & ex.Message)
            Debug.Print("ALBERTO-->Error en clsInventor-->AbreActualizaGuarda" & vbCrLf & ex.Message)
        End Try
    End Sub

    Public Function DameComponentesTreeNode(ByVal oD As Inventor.AssemblyDocument,
                                ByVal esH As Boolean, Optional ByVal solopadres As Boolean = True) As System.Windows.Forms.TreeNode()
        Dim tvn As New ArrayList
        Dim arrNodos As TreeNode() = Nothing
        Dim avisoError As String = ""
        '' Copiar el arraylist en en array
        'arrNodos = tvn.ToArray()
        Dim oE As AssemblyDocument = Nothing

        For Each oCo As Inventor.ComponentOccurrence In oD.ComponentDefinition.Occurrences
            ''Dim oDocu As Inventor.Document = oC.ReferencedDocumentDescriptor.ReferencedDocument
            '' Si es un ensamblaje.
            If oCo.ReferencedDocumentDescriptor.ReferencedDocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                oE = oCo.ReferencedDocumentDescriptor.ReferencedDocument
            Else
                Continue For
            End If
            '' Si la propiedad _TIPO no es "PADRE" este no es un ensamblaje padre y
            '' pasamos al siguiente componente. Si solopadres=true
            If solopadres = True AndAlso PropiedadLeeUsuario(oE, "_TIPO", , True, "").ToUpper <> "PADRE" Then _
        Continue For

            '' Si esH = True. Buscaremos sólo los elementos Horizontales
            '' Si esH = False. Buscaremos sólo los elementos Verticales (PI en Category)
            '' Category = PI o Category = FAMILIA·Tipo (FUTURA·fu90)
            Dim categoria As String = PropiedadLeeCategoria(oE)
            'Dim categoria As String = PropiedadLeeCategoriaApprentice(oE.FullFileName)

            '' Comprobaremos si la propiedad category tiene los valores correctos
            If categoria.ToLower = "pi" Or categoria.Contains("·") = True Then
                '' Es un elementos Vertical u Horizontal
                avisoError = ""
            Else
                '' Es vertical u horizontal, pero la propiedad category está MAL.
                avisoError = " Categoria ERROR"
            End If

            '' Ahora crearemos el Treenode y lo añadiremos a tvn
            '' Pasaremos tvn como array finalmente a arrNodos
            Dim tN As TreeNode = New TreeNode
            tN.Name = oE.DisplayName
            tN.Tag = oE.FullFileName
            If avisoError = "" Then
                tN.Text = oE.DisplayName
                tN.Checked = True
                tN.ForeColor = System.Drawing.Color.Black
            Else
                tN.Text = oE.DisplayName & avisoError
                tN.Checked = False
                tN.ForeColor = System.Drawing.Color.Red
            End If
            tvn.Add(tN)
            tN = Nothing
        Next

        If tvn.Count > 0 Then arrNodos = tvn.ToArray()

        DameComponentesTreeNode = arrNodos
        Exit Function
    End Function


    Public Function DameComponentesTreeNodeGeneral(ByVal oD As Inventor.AssemblyDocument) _
                                As System.Windows.Forms.TreeNode()
        Dim tvn As New ArrayList
        Dim arrNodos As TreeNode() = Nothing
        Dim avisoError As String = ""
        '' Copiar el arraylist en en array
        'arrNodos = tvn.ToArray()
        Dim oD1 As Inventor.Document = Nothing

        For Each oCo As Inventor.ComponentOccurrence In oD.ComponentDefinition.Occurrences

            oD1 = oCo.ReferencedDocumentDescriptor.ReferencedDocument

            '' Ahora crearemos el Treenode y lo añadiremos a tvn
            '' Pasaremos tvn como array finalmente a arrNodos
            Dim tN As TreeNode = New TreeNode
            tN.Name = oD.DisplayName
            tN.Text = oD.DisplayName
            tN.Tag = oD.FullFileName
            tN.Checked = True
            tN.ForeColor = System.Drawing.Color.Black
            tvn.Add(tN)
            tN = Nothing
        Next

        If tvn.Count > 0 Then arrNodos = tvn.ToArray()

        Return arrNodos
    End Function

    Public Function DameCadenaTextoViejo(ByVal queTexto As String) As String
        Dim resultado As String = ""
        Dim contador As Integer = 1
        Dim letra As String = Mid(queTexto, contador, 1)

        If IsNumeric(letra) = True Then
            ' Si la primera letra es un número saldremos con resultado "" de la función.
        Else
            ' Si la primera letra es un texto continuaremos hasta el final de la palabra.
            resultado = letra
            contador += 1
            While contador <= queTexto.Length
                letra = Mid(queTexto, contador, 1)
                If IsNumeric(letra) = False Then
                    resultado &= letra
                    contador += 1
                Else
                    Exit While
                End If
            End While
        End If

        DameCadenaTextoViejo = resultado.Replace("·", "")
        Exit Function
    End Function

    Public Sub BocetosVisibles(ByVal oP As Inventor.PartDocument, ByVal queBocetos As ArrayList, ByVal visibles As Boolean)
        Dim contador As Integer = 0
        For Each queB As Sketch In oP.ComponentDefinition.Sketches
            If queBocetos.Contains(queB.Name) Then
                queB.Visible = visibles
                'oP.Update2()
                oP.Parent.ActiveView.Update()
                contador += 1
                If contador >= queBocetos.Count Then Exit For
            End If
        Next
    End Sub

    Public Sub Boceto3DBorra(ByRef oCd As PartComponentDefinition, nombreboceto As String)
        For Each queB As Sketch3D In oCd.Sketches3D
            If queB.Name.ToUpper = nombreboceto.ToUpper Then
                queB.Delete()
                Exit For
            End If
        Next
    End Sub
    ''
    Public Sub ComponentOccurrencePiezaBorraCosas(ByRef oCd As PartComponentDefinition, nombres() As String)
        Dim objBorrar As ObjectCollection = oTo.CreateObjectCollection
        Dim procesados As New ArrayList
        ''
        For Each nombre As String In nombres
            If procesados.Contains(nombre.ToUpper) Then Continue For
            '' PlanarSketch
            For Each oSk2D As PlanarSketch In oCd.Sketches
                If oSk2D.Name.ToUpper = nombre.ToUpper Then
                    If oSk2D.Consumed = False Then
                        objBorrar.Add(oSk2D)
                    Else
                        For Each depend As Object In oSk2D.Dependents
                            objBorrar.Add(depend)
                        Next
                    End If
                    procesados.Add(nombre.ToUpper)
                End If
            Next
            ''
            If procesados.Contains(nombre.ToUpper) Then Continue For
            '' Sketch3D
            For Each oSk3D As Sketch3D In oCd.Sketches3D
                If oSk3D.Name.ToUpper = nombre.ToUpper Then
                    If oSk3D.Consumed = False Then
                        objBorrar.Add(oSk3D)
                    Else
                        For Each depend As Object In oSk3D.Dependents
                            objBorrar.Add(depend)
                        Next
                    End If
                    procesados.Add(nombre.ToUpper)
                End If
            Next
        Next

        oCd.DeleteObjects(objBorrar, False, False, False)
        ''
        Dim oPart As PartDocument = oCd.Document    ' CType(oCd.Document, PartDocument)
        If oPart.RequiresUpdate Then
            oPart.Update2()
        End If
        If oPart.Dirty Then oPart.Save2()
    End Sub

    Public Sub Boceto2DBorra(ByRef oCd As ComponentDefinition, nombreboceto As String)
        If TypeOf oCd Is AssemblyComponentDefinition Then
            oCd = DirectCast(DirectCast(oCd.Document, AssemblyDocument).ComponentDefinition, AssemblyComponentDefinition)
        ElseIf TypeOf oCd Is PartComponentDefinition Then
            oCd = DirectCast(DirectCast(oCd.Document, PartDocument).ComponentDefinition, PartComponentDefinition)
        End If
        ''
        For Each queB As PlanarSketch In oCd.Sketches
            If queB.Name.ToUpper = nombreboceto.ToUpper Then
                queB.Delete()
                Exit For
            End If
        Next
    End Sub
    Public Sub Boceto2DBorra(ByRef oCd As AssemblyComponentDefinition, nombreboceto As String)
        For Each queB As PlanarSketch In oCd.Sketches
            If queB.Name.ToUpper = nombreboceto.ToUpper Then
                queB.Delete()
                Exit For
            End If
        Next
    End Sub
    Public Sub Boceto2DBorra(ByRef oCd As PartComponentDefinition, nombreboceto As String)
        For Each queB As PlanarSketch In oCd.Sketches
            If queB.Name.ToUpper = nombreboceto.ToUpper Then
                queB.Delete()
                Exit For
            End If
        Next
    End Sub
    Public Function ProxyExtrusion(ByVal asmPadre As AssemblyDocument, ByVal queOC As ComponentOccurrence, ByVal queEX As String) As ExtrudeFeatureProxy
        Dim resultado As ExtrudeFeatureProxy = Nothing

        Dim oPie As PartDocument = queOC.SubOccurrences.Item(1).ReferencedDocumentDescriptor.ReferencedDocument
        Dim OC1 As ComponentOccurrence = queOC.SubOccurrences.Item(1)
        Dim oPartCompDef As PartComponentDefinition
        oPartCompDef = oPie.ComponentDefinition


        Dim oExtrudeFeature As ExtrudeFeature
        Dim oExtrudeFeatureProxy As Object = Nothing    ' ExtrudeFeatureProxy = Nothing
        Try
            oExtrudeFeature = oPartCompDef.Features.ExtrudeFeatures.Item(queEX)  '.Item(1)
            Call OC1.CreateGeometryProxy(oExtrudeFeature, oExtrudeFeatureProxy)
        Catch ex As Exception
            oExtrudeFeatureProxy = Nothing
        End Try

        If Not (oExtrudeFeatureProxy Is Nothing) Then resultado = oExtrudeFeatureProxy

        ProxyExtrusion = resultado
    End Function

    Public Sub FotoPantalla()
        Dim tieneTriada As Boolean = oAppI.GeneralOptions.Show3DIndicator
        oAppI.GeneralOptions.Show3DIndicator = False
        Dim fd As New System.Windows.Forms.SaveFileDialog
        fd.AddExtension = True
        fd.DefaultExt = "png"
        fd.Filter = "Imagen PNG|*.png|Imagen BMP|*.bmp|Imagen JPG|*.jpg"
        fd.FilterIndex = 1
        fd.InitialDirectory = My.Application.Info.DirectoryPath
        fd.FileName = "PantallaInventor"
        If fd.ShowDialog = System.Windows.Forms.DialogResult.OK Then
            Me.oAppI.ActiveView.SaveAsBitmap(fd.FileName, 1024, 768)
            Threading.Thread.Sleep(2000)
            Call Process.Start(fd.FileName)
        End If
        oAppI.GeneralOptions.Show3DIndicator = tieneTriada
    End Sub


    Public Sub CopiaConApprentice(ByVal ficheroOrigen As String, ByVal ficheroDestino As String,
                              Optional ByVal sobreescribir As Boolean = False, Optional ByVal cerrar As Boolean = True)
        If Dir(ficheroOrigen) <> "" Then
            ' Create a new instance of Apprentice.

            Dim oApprentice As New ApprenticeServerComponent
            ' Open a document.
            Dim oDoc As ApprenticeServerDocument

            oDoc = oApprentice.Open(ficheroOrigen)
            Try
                If My.Computer.FileSystem.FileExists(ficheroDestino) = True And
            sobreescribir = True Then
                    My.Computer.FileSystem.DeleteFile(ficheroDestino)
                End If
            Catch ex As Exception

            End Try
            Try
                If oDoc.NeedsMigrating = False Then
                    oApprentice.FileSaveAs.AddFileToSave(oDoc, ficheroDestino)
                    oApprentice.FileSaveAs.ExecuteSaveCopyAs()
                    If cerrar = True Then oDoc.Close()
                Else
                    MsgBox("Hay que abrir antes el fichero y guardarlo en versión actual. Está en versión antigüa")
                End If
            Catch ex As Exception
                MsgBox("Error CopiaConApprentice")
            End Try
        End If
    End Sub


    Public Function ComponenteCuantosOccAsm(ByVal queOC As ComponentOccurrence, ByVal queAs As AssemblyDocument, ByVal quePrefijoComponente As String) As Integer
        Dim resultado As Integer = 0
        If Not (queOC Is Nothing) AndAlso queOC.SubOccurrences.Count > 0 Then
            Try
                For Each oC As ComponentOccurrence In queOC.SubOccurrences
                    If oC.Name.StartsWith(quePrefijoComponente) Then resultado += 1
                Next
            Catch ex As Exception
                MsgBox("Error en ComponenteCuantos --> " & vbCrLf & ex.Message & vbCrLf & queOC.Name)
            Finally
            End Try
        End If
        If Not (queAs Is Nothing) AndAlso queAs.ComponentDefinition.Occurrences.Count > 0 Then
            Try
                For Each oC As ComponentOccurrence In queAs.ComponentDefinition.Occurrences
                    If oC.Name.StartsWith(quePrefijoComponente) Then resultado += 1
                Next
            Catch ex As Exception
                MsgBox("Error en ComponenteCuantos --> " & vbCrLf & ex.Message & vbCrLf & queOC.Name)
            Finally
            End Try
        End If
        ComponenteCuantosOccAsm = resultado
        Exit Function
    End Function

    Public Function DameComponentesArrTreeNodes(ByVal oD As Inventor.AssemblyDocument,
                                ByVal esH As Boolean, Optional ByVal solopadres As Boolean = True) As System.Windows.Forms.TreeNode()
        Dim tvn As New ArrayList
        Dim arrNodos(-1) As TreeNode
        ''Dim arrNodos As System.Array = Nothing
        Dim avisoError As String = ""
        '' Copiar el arraylist en en array
        'arrNodos = tvn.ToArray()
        Dim oDocCom As Inventor.Document = Nothing

        For Each oCo As Inventor.ComponentOccurrence In oD.ComponentDefinition.Occurrences
            ''Dim oDocu As Inventor.Document = oC.ReferencedDocumentDescriptor.ReferencedDocument
            '' Si es un ensamblaje.
            Dim CaminoTodo As String = oCo.ReferencedDocumentDescriptor.ReferencedFileDescriptor.FullFileName
            Dim CaminoDir As String = DameParteCamino(CaminoTodo, IEnum.ParteCamino.CaminoConFicheroSinExtensionBarra)
            Dim NombreSolo As String = DameParteCamino(CaminoTodo, IEnum.ParteCamino.SoloFicheroSinExtension)
            If oCo.ReferencedDocumentDescriptor.ReferencedDocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                oDocCom = oCo.ReferencedDocumentDescriptor.ReferencedDocument
            Else
                Continue For
            End If
            '' Si la propiedad _TIPO no es "PADRE" este no es un ensamblaje padre y
            '' pasamos al siguiente componente. Si solopadres=true
            If solopadres = True AndAlso PropiedadLeeUsuario(oDocCom, "_TIPO", , True, "").ToUpper <> "PADRE" Then _
        Continue For

            '' Si esH = True. Buscaremos sólo los elementos Horizontales
            '' Si esH = False. Buscaremos sólo los elementos Verticales (PI en Category)
            '' Category = PI o Category = (FU90)    Con FU leemos ADAPRETERSA.ini [CAMINOS]-->FU=FUTURA 'FAMILIA·Tipo (FUTURA·fu90)
            Dim categoria As String = PropiedadLeeCategoria(oDocCom)
            'Dim categoria As String = PropiedadLeeCategoriaApprentice(oDocCom.FullFileName)

            '' Comprobaremos si la propiedad category tiene los valores correctos
            'If categoria.ToLower = "pi" Or categoria.Contains("·") = True Then
            If (categoria <> "") = True Then
                '' Es un elementos Vertical u Horizontal
                avisoError = ""
            Else
                '' Es vertical u horizontal, pero la propiedad category está MAL.
                avisoError = " Categoria ERROR"
            End If
            '' Elementos HORIZONTALES sólo (menos las Delta [DE])
            '' tn.Name = COM·H·001·1
            '' tn.Text = COM·H·001·1:1
            '' tn.Tag = C:\[directorios]\COM·H·001·1.iam
            '' Sin Delta
            'If (esH = True And categoria.ToLower <> "pi" And _
            'categoria <> "" And _
            'categoria.ToLower.StartsWith("de·") = False) Then
            '' Con Delta.
            If (esH = True And categoria.ToLower <> "pi" And categoria <> "") Then
                '' Ahora crearemos el Treenode y lo añadiremos a tvn
                '' Pasaremos tvn como array finalmente a arrNodos
                '' Si es Horizontal y también la Delta.
                Dim tN As TreeNode = New TreeNode
                tN.Name = NombreSolo    ' oCo.Name
                'tN.Name = oDocCom.DisplayName
                tN.Tag = CaminoTodo ' oDocCom.FullFileName
                'If avisoError = "" Then
                tN.Text = oCo.Name
                '' Si existe el directorio y ensamblaje armado y plano armado y pieza de parámetros. El check estará a false
                If Dir(CaminoDir) <> "" And
                Dir(CaminoDir & NombreSolo & "·armado.iam") <> "" And
                Dir(CaminoDir & NombreSolo & "·armado.idw") <> "" And
                Dir(CaminoDir & NombreSolo & "·parametros.ipt") <> "" Then
                    tN.Checked = False
                Else
                    tN.Checked = True
                End If
                'If Dir(CaminoDir & NombreSolo & "·armado.ini") = "" Then tN.Checked = True
                tN.ForeColor = System.Drawing.Color.Black
                If tvn.Contains(tN) = False Then tvn.Add(tN)
                tN = Nothing
                '' Elementos VERTICALES sólo
            ElseIf (esH = False And categoria = "pi") Then  ' Es un pilar Vertical                '' Ahora crearemos el Treenode y lo añadiremos a tvn
                '' Pasaremos tvn como array finalmente a arrNodos
                Dim tN As TreeNode = New TreeNode
                tN.Name = NombreSolo    ' oCo.Name
                'tN.Name = oDocCom.DisplayName
                tN.Tag = CaminoTodo ' oDocCom.FullFileName
                tN.Text = oCo.Name
                tN.Checked = False
                If Dir(CaminoDir) = "" Then tN.Checked = True
                'If Dir(CaminoDir & NombreSolo & "·armado.ini") = "" Then tN.Checked = True
                If Dir(CaminoDir & NombreSolo & "·armado.iam") = "" Then tN.Checked = True
                If Dir(CaminoDir & NombreSolo & "·armado.idw") = "" Then tN.Checked = True
                tN.ForeColor = System.Drawing.Color.Black
                If tvn.Contains(tN) = False Then tvn.Add(tN)
                tN = Nothing
            End If
        Next

        Try
            If tvn.Count > 0 Then
                ReDim arrNodos(tvn.Count - 1)
                tvn.CopyTo(arrNodos)
            End If
        Catch ex As Exception
            MsgBox("Error al copiar en arrNodos" & vbCrLf & vbCrLf & ex.Message)
            DameComponentesArrTreeNodes = Nothing
        End Try
        Return arrNodos
    End Function

    Public Sub ZoomAllFit3DBest(Optional ByVal ponerIso As Boolean = True)
        If ponerIso Then
            oAppI.ActiveView.Camera.ViewOrientationType = ViewOrientationTypeEnum.kIsoTopRightViewOrientation
            oAppI.ActiveView.Camera.ApplyWithoutTransition()
        End If
        ' Retornar a la vista Inicio (Home)
        Dim oCd As Inventor.ControlDefinition = oAppI.CommandManager.ControlDefinitions.Item("AppViewCubeHomeCmd")
        Call oCd.Execute2(True)
        'Zoom Todo
        'oAppCls.ActiveView.Fit()
        oCd = Nothing
    End Sub

    Public Sub ZoomAllFit3D(Optional ByVal ponerIso As Boolean = True)
        Dim oVie As Inventor.View
        Dim oCam As Inventor.Camera
        oVie = Me.oAppI.ActiveView  'oPie.Views.Item(1)   ' oApp.ActiveView
        oCam = oVie.Camera
        If ponerIso Then oCam.ViewOrientationType = ViewOrientationTypeEnum.kIsoTopRightViewOrientation
        oCam.ApplyWithoutTransition()
        oCam.Fit()
        oCam.ApplyWithoutTransition()
        ' Fit the view to see the result.
        ' Dim cam As Camera
        ' Set cam = ThisApplication.ActiveView.Camera
        ' cam.Fit()
        ' cam.Apply()

    End Sub

    Public Sub ZoomAllFit2D(ByVal oSh As Sheet)
        Dim oVie As Inventor.View
        Dim oCam As Inventor.Camera
        oVie = Me.oAppI.ActiveView  'oPie.Views.Item(1)   ' oApp.ActiveView
        oCam = oVie.Camera
        oCam.SetExtents(oSh.Width + 1, oSh.Height + 1)
        'oCam.Fit()
        oCam.ApplyWithoutTransition()
    End Sub

    Public Sub ZoomOccurrence(ByVal oC As ComponentOccurrence, ByVal pt3D As Point, ByVal pt2D As Point2d)
        Dim ancho, alto, ancho2d, alto2d As Double
        Dim pt1, pt2, ptCentro3D As Inventor.Point
        Dim pt1_2d, pt2_2d, ptC2d_Destino As Point2d
        Dim oVie As Inventor.View = oAppI.ActiveView
        Dim oCam As Camera = oVie.Camera
        'Dim oCProx As ComponentOccurrenceProxy = Nothing
        'oC.CreateGeometryProxy(oC.ContextDefinition.Occurrences(1).SubOccurrences(1), oCProx)
        pt1 = Nothing : pt2 = Nothing : pt1_2d = Nothing : pt2_2d = Nothing
        'queEns.SelectSet.Select(oC)
        'Dim oSel As SelectSet = queEns.SelectSet
        'oSel.Item(1).Rangebox()
        'MsgBox(oSel.Item(1).GetType.ToString)
        'Dim oH As HighlightSet = oAppCls.ActiveDocument.CreateHighlightSet
        'oH.AddItem(oC)
        'Dim oSb As SurfaceBody = oC.SubOccurrences(1).SurfaceBodies.Item(1)
        'Dim oRb As Box = oSb.RangeBox
        'Dim oSbP As SurfaceBodyProxy = Nothing
        'oC.SubOccurrences(1).CreateGeometryProxy(oSb, oSbP)
        'Dim oRbP As Box = oSbP.RangeBox

        ' oAppCls.ActiveDocument.CreateHighlightSet.AddItem(oC)
        'oAppCls.ActiveDocument.

        oCam.GetExtents(ancho, alto)
        Dim iB As Inventor.Box = oC.Definition.RangeBox
        pt1 = iB.MinPoint
        pt2 = iB.MaxPoint
        pt1_2d = oCam.ModelToViewSpace(pt1)
        pt2_2d = oCam.ModelToViewSpace(pt2)
        ancho2d = Math.Abs(pt2_2d.X - pt1_2d.X)
        alto2d = Math.Abs(pt2_2d.Y - pt1_2d.Y)
        'ptC2d_Destino = oAppCls.TransientGeometry.CreatePoint2d(pt1_2d.X + (ancho2d / 2), pt1_2d.X + (alto2d / 2))
        ptC2d_Destino = oCam.ModelToViewSpace(pt1)
        ptCentro3D = oCam.ViewToModelSpace(ptC2d_Destino)
        'MsgBox("Vista Inventor : " & ancho & " / " & alto & vbCrLf & _
        '"En Componente  : " & pt2.X - pt1.X & " / " & pt2.Y - pt1.Y & vbCrLf & _
        '"En Componente1  : " & oRb.MaxPoint.X - oRb.MinPoint.X & " / " & oRb.MaxPoint.Y - oRb.MinPoint.Y & vbCrLf & _
        '"En Componente Proxy  : " & oRbP.MaxPoint.X - oRbP.MinPoint.X & " / " & oRbP.MaxPoint.Y - oRbP.MinPoint.Y & vbCrLf & _
        '"En Vista 2D    : " & pt2_2d.X - pt1_2d.X & " / " & Math.Abs(pt2_2d.Y - pt1_2d.Y) & vbCrLf & _
        '"Centro Vista   : " & -ptC2d_Destino.X & " / " & -ptC2d_Destino.Y)
        'oCam.ComputeWithMouseInput(oCam.ModelToViewSpace(oCam.Eye), ptC2d_Destino, 0, ViewOperationTypeEnum.kPanViewOperation)
        oCam.ComputeWithMouseInput(pt2D, oCam.ModelToViewSpace(oCam.Eye), 0, ViewOperationTypeEnum.kPanViewOperation)
        'oCam.ComputeWithMouseInput(pt1_2d, pt2_2d, 0, ViewOperationTypeEnum.kZoomViewOperation)
        oCam.Apply()
        'oCam.Eye = ptCentro3D
        'oCam.Target = ptCentro3D
        oCam.SetExtents(ancho2d * 8, alto2d * 8)
        oCam.Apply()
    End Sub
    '

    'Public Sub ZoomOccurrence(ByVal oC As ComponentOccurrence,
    '                          Optional pt3D As Point = Nothing)
    '    Dim ancho, alto, ancho2d, alto2d As Double
    '    Dim pt1, pt2, target As Inventor.Point
    '    Dim pt1_2d, pt2_2d As Point2d
    '    Dim oVie As Inventor.View = oAppCls.ActiveView
    '    Dim oCam As Camera = oVie.Camera
    '    '
    '    pt1 = Nothing : pt2 = Nothing : pt1_2d = Nothing : pt2_2d = Nothing
    '    '
    '    oCam.GetExtents(ancho, alto)
    '    Dim iB As Inventor.Box = oC.Definition.RangeBox
    '    pt1 = iB.MinPoint
    '    pt2 = iB.MaxPoint
    '    pt1_2d = oCam.ModelToViewSpace(pt1)
    '    pt2_2d = oCam.ModelToViewSpace(pt2)
    '    ancho2d = Math.Abs(pt2_2d.X - pt1_2d.X) / 75
    '    alto2d = Math.Abs(pt2_2d.Y - pt1_2d.Y) / 75
    '    'ancho2d = 15    'Math.Abs(pt2_2d.X - pt1_2d.X)
    '    'alto2d = 11 ' Math.Abs(pt2_2d.Y - pt1_2d.Y)
    '    '
    '    If pt3D Is Nothing Then
    '        target = oC.Definition.SurfaceBodies(1).Faces(1).PointOnFace
    '    Else
    '        target = pt3D
    '    End If
    '    oCam.Target = target
    '    '
    '    Dim ancho2dahora, alto2dahora As Double
    '    oCam.GetExtents(ancho2dahora, alto2dahora)
    '    '
    '    If ancho2dahora > ancho2d Then
    '        oCam.SetExtents(ancho2d, alto2d)
    '    End If
    '    oCam.Apply()
    '    'oCam.SetExtents(ancho2d * 8, alto2d * 8)
    '    'oCam.Apply()
    'End Sub
    'Public Sub ZoomOccurrenceFace(ByVal oFaceProxy As FaceProxy)
    '    Dim ancho, alto, ancho2d, alto2d As Double
    '    Dim pt1, pt2, target As Inventor.Point
    '    Dim pt1_2d, pt2_2d As Point2d
    '    Dim oVie As Inventor.View = oAppCls.ActiveView
    '    Dim oCam As Camera = oVie.Camera
    '    '
    '    pt1 = Nothing : pt2 = Nothing : pt1_2d = Nothing : pt2_2d = Nothing
    '    oCam.GetExtents(ancho, alto)
    '    '
    '    Dim iB As Inventor.Box = oFaceProxy.Evaluator.RangeBox
    '    pt1 = iB.MinPoint
    '    pt2 = iB.MaxPoint
    '    pt1_2d = oCam.ModelToViewSpace(pt1)
    '    pt2_2d = oCam.ModelToViewSpace(pt2)
    '    ancho2d = 13    ' Math.Abs(pt2_2d.X - pt1_2d.X)
    '    alto2d = 10 ' Math.Abs(pt2_2d.Y - pt1_2d.Y)
    '    'ptC2d_Destino = oAppCls.TransientGeometry.CreatePoint2d(pt1_2d.X + (ancho2d / 2), pt1_2d.X + (alto2d / 2))
    '    'ptC2d_Destino = oCam.ModelToViewSpace(pt1)
    '    'ptCentro3D = oCam.ViewToModelSpace(ptC2d_Destino)
    '    target = oFaceProxy.PointOnFace
    '    oCam.Target = target
    '    'oCam.ComputeWithMouseInput(pt1_2d, pt2_2d, 0, ViewOperationTypeEnum.kPanViewOperation)
    '    'oCam.Apply()
    '    Dim ancho2dahora, alto2dahora As Double
    '    oCam.GetExtents(ancho2dahora, alto2dahora)
    '    'If ancho2dahora <> ancho2d * 0.1 Then
    '    '    oCam.SetExtents(ancho2d * 0.1, alto2d * 0.1)
    '    'End If
    '    If ancho2dahora <> ancho2d Then
    '        oCam.SetExtents(ancho2d, alto2d)
    '    End If
    '    oCam.Apply()
    '    'oCam.SetExtents(ancho2d * 8, alto2d * 8)
    '    'oCam.Apply()
    'End Sub
    Public Sub ZoomSelection(oObj As Object)
        'oAppCls.ActiveDocument.SelectSet.Clear()
        'oAppCls.ActiveDocument.SelectSet.Select(oObj)
        oAppI.CommandManager.ControlDefinitions.Item("AppZoomSelectCmd").Execute()
        'oAppCls.ActiveDocument.SelectSet.Clear()
        'Dim ZoomCommand As String : ZoomCommand = "AppZoomSelectCmd"
        'Call ThisApplication.CommandManager.ControlDefinitions(ZoomCommand).Execute
    End Sub
    '
    Public Function SelectionDame(filtro As Inventor.SelectionFilterEnum, mensaje As String) As Object
        ' Get a feature selection from the user
        Dim oObject As Object = Nothing
        oObject = oAppI.CommandManager.Pick(filtro, mensaje)
        Return oObject
    End Function
    '
    Public Sub DesignViewRepresentationsActivaCrea(ByVal Ensamblaje As AssemblyDocument, ByVal bolActivar As Boolean, Optional ByVal queNombre As String = "Materiales")
        If bolActivar = False Then
            ' Principal
            Ensamblaje.ComponentDefinition.RepresentationsManager.DesignViewRepresentations(1).Activate()
            Exit Sub
        End If

        If Ensamblaje.ComponentDefinition.RepresentationsManager.ActiveDesignViewRepresentation.Name = queNombre Then
            Exit Sub
        End If

        '' Ponemos la representación "queNombre" como activa. Si no existe, la creamos antes.
        Dim viewMateriales As DesignViewRepresentation = Nothing
        Try
            viewMateriales = Ensamblaje.ComponentDefinition.RepresentationsManager.DesignViewRepresentations.Item(queNombre)
        Catch ex As Exception
            viewMateriales = Ensamblaje.ComponentDefinition.RepresentationsManager.DesignViewRepresentations.Add(queNombre)
        End Try
        viewMateriales.Activate()
        '' ****************************************************
        '' La 1 es la principal.
        'Dim x As Integer
        'For x = 1 To oAsmC.RepresentationsManager.LevelOfDetailRepresentations.Count
        'Debug.Print(x & ".- " & oAsmC.RepresentationsManager.LevelOfDetailRepresentations.Item(x).Name)
        'Next
        '' Item 1.- Principal
        '' Item 2.- Todos los componentes desactivados
        '' Item 3.- Todas las piezas desactivadas
        '' Item 4.- Todo el Centro de contenido desactivado
        '' Item 5.- Desactivados (en el desarrollo de Pretersa)
    End Sub
    Public Sub RepresentacionDesactivaArmadoCrea(ByVal queEns As AssemblyDocument, ByVal queRepre As String, Optional ByVal queDesactivamos As String = "·armado")
        '"armado.iam"   "armado_cabeza.iam"
        Dim repArmado As LevelOfDetailRepresentation = Nothing
        Try
            repArmado = queEns.ComponentDefinition.RepresentationsManager.LevelOfDetailRepresentations.Item(queRepre)    '("Desactivados") ("Armado_sin") ("Principal")
        Catch ex As Exception
            '' No existe la representación.
            '' ** Primer ponemos la principal activa. Para, a partir de ella, crear la nueva
            Call queEns.ComponentDefinition.RepresentationsManager.LevelOfDetailRepresentations.Item(1).Activate()
            repArmado = queEns.ComponentDefinition.RepresentationsManager.LevelOfDetailRepresentations.Add(queRepre)
        End Try
        repArmado.Activate()
        'oCEnsam.SubOccurrences(2).Suppress()
        For Each oC As ComponentOccurrence In queEns.ComponentDefinition.Occurrences
            'If oC.Name.Contains(queDesactivamos) = True Then oC.Suppress(True)
            If oC.SubOccurrences.Count = 0 Then Continue For
            For Each oc1 As ComponentOccurrence In oC.SubOccurrences
                If oc1.Name.Contains(queDesactivamos) = True Then oc1.Suppress(True)
            Next
        Next
        If queEns.RequiresUpdate = True Then
            queEns.Update2()
            queEns.Save2()
        End If
    End Sub

    Public Sub Representaciones_Vista_Activa(ByVal Ensamblaje As AssemblyDocument, ByVal bolActivar As Boolean, Optional ByVal queNombre As String = "Default")
        If bolActivar = False Then
            ' Default
            Ensamblaje.ComponentDefinition.RepresentationsManager.DesignViewRepresentations.Item(2).Activate()
            Exit Sub
        End If

        If Ensamblaje.ComponentDefinition.RepresentationsManager.ActiveDesignViewRepresentation.Name = queNombre Then
            Exit Sub
        End If
        '' Ponemos la representación "queNombre" como activa. Si no existe, la creamos antes.
        Dim queVista As DesignViewRepresentation = Nothing
        Try
            queVista = Ensamblaje.ComponentDefinition.RepresentationsManager.LevelOfDetailRepresentations.Item(queNombre)
            queVista.Activate()
        Catch ex As Exception
            ''
        End Try
        '' ****************************************************
        '' La 1 es la principal.
        'Dim x As Integer
        'For x = 1 To oAsmC.RepresentationsManager.LevelOfDetailRepresentations.Count
        'Debug.Print(x & ".- " & oAsmC.RepresentationsManager.LevelOfDetailRepresentations.Item(x).Name)
        'Next
        '' Item 1.- Principal
        '' Item 2.- Default
        '' Otras)
    End Sub

    Public Sub RepresentacionActivaCrea(ByVal Ensamblaje As AssemblyDocument, ByVal bolActivar As Boolean, Optional ByVal queNombre As String = "Desactivados")
        If bolActivar = False Then
            ' Principal
            Ensamblaje.ComponentDefinition.RepresentationsManager.LevelOfDetailRepresentations.Item(1).Activate()
            Exit Sub
        End If

        If Ensamblaje.ComponentDefinition.RepresentationsManager.ActiveLevelOfDetailRepresentation.Name = queNombre Then
            Exit Sub
        End If

        'If queNombre = "" Then
        'MsgBox("Hay que especificar un nombre de Representación para crear y activar")
        'Exit Sub
        'End If

        '' Ponemos la representación "queNombre" como activa. Si no existe, la creamos antes.
        Dim repArmadoSin As LevelOfDetailRepresentation = Nothing
        Try
            repArmadoSin = Ensamblaje.ComponentDefinition.RepresentationsManager.LevelOfDetailRepresentations.Item(queNombre)
        Catch ex As Exception
            repArmadoSin = Ensamblaje.ComponentDefinition.RepresentationsManager.LevelOfDetailRepresentations.Add(queNombre)
        End Try
        repArmadoSin.Activate(True)
        '' ****************************************************
        '' La 1 es la principal.
        'Dim x As Integer
        'For x = 1 To oAsmC.RepresentationsManager.LevelOfDetailRepresentations.Count
        'Debug.Print(x & ".- " & oAsmC.RepresentationsManager.LevelOfDetailRepresentations.Item(x).Name)
        'Next
        '' Item 1.- Principal
        '' Item 2.- Todos los componentes desactivados
        '' Item 3.- Todas las piezas desactivadas
        '' Item 4.- Todo el Centro de contenido desactivado
        '' Item 5.- Desactivados (en el desarrollo de Pretersa)
    End Sub

    ''' <summary>
    ''' 0    Principal
    ''' 1    Todos los componentes desactivados
    ''' 2    Todas las piezas desactivadas
    ''' 3    Todo el Centro de contenido desactivado
    ''' 4    Desactivados (en el desarrollo de Pretersa)
    ''' </summary>
    ''' <param name="queFichero">Nombre Full del Fichero</param>
    ''' <param name="queRep">Numero (0 al 3) para las estandar (4 a ...) para las personalizadas</param>
    ''' <returns>NombreFull.iam[menorque]Representacion[mayorque]</returns>
    ''' <remarks></remarks>
    Public Function RepresentacionDameFullDoc(ByVal queFichero As String, ByVal queRep As Integer) As String
        Dim resultado As String = ""

        If queFichero.ToLower.EndsWith(".iam") = False Then
            Return resultado
            Exit Function
        End If
        Dim RepPrincipal As String = oAppI.FileManager.GetLevelOfDetailRepresentations(queFichero)(queRep)
        'RepPrincipal.Name
        Return oAppI.FileManager.GetFullDocumentName(queFichero, RepPrincipal)
        '' ****************************************************
    End Function

    '' Me da el nombre (camino+nombre) de la representación ligera de un Ensamblaje,
    '' para leer cosas, no guardar o crear.
    Public Function DameCaminoSinComponentesInventor(ByVal queFichero As String) As String
        Dim resultado As String = ""    '"Todos los componentes desactivados"

        'Dim strFileName As String
        'strFileName = "C:\Program Files\Autodesk\Inventor 2008\Tutorial Files\engine_assy.iam"
        ' Respectively for Inventor 11:   strFileName = "C:\Program Files\Autodesk\Inventor 11\Tutorial Files\engine_assy.iam"


        Dim straLOD() As String
        straLOD = oAppI.FileManager.GetLevelOfDetailRepresentations(queFichero)

        ' if it's a pre Inventor 11 file, then it only has a "Master" representation [UBound() = 0]
        ' if it's an Inventor 11 file or later, then it has an "All Components Suppressed" representation as well [UBound() > 0]
        '' Returned in the following order: 
        '' 1.   Master
        '' 2.   All Parts Suppressed 
        '' 3.   All Components Suppressed 
        '' 4.   All Content Suppressed 
        '' 5.   Other

        If UBound(straLOD) > 0 Then
            For Each nombre As String In straLOD
                Debug.Print(nombre)
            Next
            resultado = oAppI.FileManager.GetFullDocumentName(queFichero, straLOD(2))  'LevelOfDetailEnum.kAllComponentsSuppressedLevelOfDetail.ToString)    ' straLOD(2))
        End If


        System.GC.WaitForPendingFinalizers()
        System.GC.Collect()
        Return resultado
    End Function

    Public Sub TextoCrearEnHoja(ByVal oDoc As DrawingDocument, ByVal queTexto As String, Optional ByVal queSk As String = "2acad", Optional ByVal borra As Boolean = False)
        If oDoc Is Nothing Then Exit Sub
        If Me.oAppI Is Nothing Then Exit Sub
        AppActivate(oAppI.Caption)

        Dim oSh As Sheet = oDoc.ActiveSheet
        Dim oSk As DrawingSketch = Nothing

        ' Open the sketched symbol definition's sketch for edit. This is done by calling the Edit
        ' method of the SketchedSymbolDefinition to obtain a DrawingSketch. This actually creates
        ' a copy of the sketched symbol definition's and opens it for edit.
        Try
            oSk = oDoc.ActiveSheet.Sketches.Item(queSk)
            oSk.Delete()
        Catch ex As Exception
            '' Si existe el Boceto, lo borramos.
        End Try
        If borra = True Then Exit Sub
        oSk = oDoc.ActiveSheet.Sketches.Add
        oSk.Name = queSk
        oSk.Edit()

        Dim oTG As TransientGeometry
        oTG = oAppI.TransientGeometry

        Dim sText As String
        Dim fontSize As Integer = CInt(oSh.Width / 20) ' 2
        sText = "<StyleOverride Font='Arial' Bold='True' FontSize='" & fontSize & "'>" & queTexto & "</StyleOverride>"
        Dim oTextBox As Inventor.TextBox
        Dim pt1, pt2 As Inventor.Point2d
        pt1 = oTG.CreatePoint2d(0, oSh.Height)
        pt2 = oTG.CreatePoint2d(oSh.Width, 0)

        'oTextBox = oSk.TextBoxes.AddFitted(oTG.CreatePoint2d(oSh.Width / 2, oSh.Height / 2), sText)
        oTextBox = oSk.TextBoxes.AddByRectangle(pt1, pt2, "")
        'oTextBox.VerticalJustification = VerticalTextAlignmentEnum.kAlignTextMiddle
        'oTextBox.HorizontalJustification = HorizontalTextAlignmentEnum.kAlignTextCenter
        oTextBox.VerticalJustification = VerticalTextAlignmentEnum.kAlignTextLower
        oTextBox.HorizontalJustification = HorizontalTextAlignmentEnum.kAlignTextLeft
        oTextBox.FormattedText = sText

        ' Add a prompted text field at the center of the sketch circle.
        'Call oSketchedSymbolDef.ExitEdit(True)
        oSk.ExitEdit()
    End Sub

    Public Function BuscaPlanosEnDirRefInventor(ByVal arrFicheros As ArrayList, Optional ByVal dirBuscar As String = "", Optional ByVal queTb As System.Windows.Forms.TextBox = Nothing) As ArrayList
        Dim colPlanos As New ArrayList
        Dim colFinal As New ArrayList

        If dirBuscar = "" Then dirBuscar = dirProyectoInv

        If queTb IsNot Nothing Then queTb.Text = ""
        'clbFicheros.Items.AddRange(IO.Directory.GetFileSystemEntries(dirProyectoInv, "*.idw"))
        '' ***** Todos los DIRECTORIOS ( , "SIN MASCARA")
        'clbFicheros.Items.AddRange(IO.Directory.GetFileSystemEntries(dirProyectoInv))
        '' ***** Todos los FICHEROS IDW ( , "*.idw")

        colPlanos.AddRange(IO.Directory.GetFiles(dirBuscar, "*.idw", IO.SearchOption.AllDirectories))
        colPlanos.AddRange(IO.Directory.GetFiles(dirBuscar, "*.dwg", IO.SearchOption.AllDirectories))

        'Dim oAppr As New Inventor.ApprenticeServerComponent

        oAppI.SilentOperation = True
        For Each queF As String In colPlanos
            '' Si el fichero está en un directorio OldVersions
            If queF.ToLower.Contains("oldversions") Then Continue For
            DoEventsInventor()
            If queF.ToLower.EndsWith(".dwg") AndAlso oAppI.FileManager.IsInventorDWG(queF) = False Then
                Continue For
            End If

            If queTb IsNot Nothing Then queTb.Text &= queF & vbCrLf
            Dim oDocRef As Inventor.Document = Nothing
            oDocRef = oAppI.Documents.Open(queF, False)

            For Each oDocRef In oDocRef.AllReferencedDocuments
                DoEventsInventor()
                If queTb IsNot Nothing Then queTb.Text &= "     " & oDocRef.FullFileName & vbCrLf
                If queTb IsNot Nothing Then queTb.Select(queTb.TextLength - 1, 1)
                If queTb IsNot Nothing Then queTb.ScrollToCaret()
                If arrFicheros.Contains(oDocRef.FullFileName) Then
                    If colFinal.Contains(queF) = False Then colFinal.Add(queF)
                End If
            Next
            If oDocRef.NeedsMigrating = True Then oDocRef.Save2(True)
            oDocRef.ReleaseReference()
            oDocRef.Close(True)
            oDocRef = Nothing
        Next
        oAppI.SilentOperation = False

        Return colFinal
    End Function

    Public Sub CreaDibujo(ByVal oDoc As Inventor.Document, Optional ByVal quePlantilla As String = "", Optional ByVal creacotasBase As Boolean = True)
        'AppActivate(Me.titulo)
        Dim queCamino As String = oDoc.FullFileName
        'Dim oDoc As Inventor.Document = oAppCls.Documents.ItemByName(queCamino)
        '' Si no indicamos plantilla. Cogeremos el nombre plantilla por defecto para planos.
        If quePlantilla = "" Then quePlantilla = oAppI.FileManager.GetTemplateFile(DocumentTypeEnum.kDrawingDocumentObject)
        '' Añadimos el nuevo dibujo IDW
        Dim oIdw As DrawingDocument = oAppI.Documents.Add(DocumentTypeEnum.kDrawingDocumentObject, quePlantilla)
        '' Creamos el nombre del plano desde el nombre de la pieza o ensamblaje que indicamos "queCamino"
        Dim extension As String = DameParteCamino(quePlantilla, IEnum.ParteCamino.SoloExtension)
        Dim NombrePlano As String = DameParteCamino(queCamino, IEnum.ParteCamino.SoloCambiaExtension, extension)
        '' Guardamos el plano en el mismo directorio que la pieza o ensamblaje indicado.
        'oIdw.SaveAs(NombrePlano, False)
        'oIdw.FullFileName = NombrePlano
        'oIdw.File.FullFileName = NombrePlano

        '' El objeto hoja activo (Sheet)
        Dim oSheet As Sheet = oIdw.ActiveSheet

        '' Cogemos los puntos medios de ancho y alto
        Dim oPoint1 As Point2d
        Dim px, py As Double
        px = (oSheet.Width / 2)
        py = (oSheet.Height / 2)
        oPoint1 = oTg.CreatePoint2d(CInt(px), CInt(py))


        '' ***** Crear Vista BASE
        Dim oView1 As Inventor.DrawingView
        oView1 = oSheet.DrawingViews.AddBaseView(oDoc,
     oPoint1, 1.0#, ViewOrientationTypeEnum.kFrontViewOrientation, DrawingViewStyleEnum.kHiddenLineDrawingViewStyle)
        oView1.Name = "VistaBase"

        ' Escalar y Mover la vista Base para posicionarla en cuadrante sup.-izda.
        '' Cogemos los valores que necesitamos (Height y Width de la DrawingView)
        Dim escalaFinal As String = ""
        Dim oView1H As Double = oView1.Height
        Dim oView1W As Double = oView1.Width
        Dim cuartoH As Double = ((oSheet.Height / 2) - 4)
        Dim cuartoW As Double = ((oSheet.Width / 2) - 3)
        '' Si la vista es más grande (alto y ancho) que el cuarto de la hoja
        If oView1H > cuartoH Or oView1W > cuartoW Then
            Dim denominador As Integer = 2
            Do While oView1H > cuartoH Or oView1W > cuartoW
                escalaFinal = "1:" & denominador
                oView1.ScaleString = escalaFinal
                'oView1H = (oView1.Height) * CInt(numerador)
                'oView1W = (oView1.Width) * CInt(numerador)
                oView1H = (oView1.Height)
                oView1W = (oView1.Width)
                denominador += 1
                'oView1.Scale -= 0.1
                DoEventsInventor()
            Loop
        ElseIf oView1H < cuartoH Or oView1W < cuartoW Then    '' Si es más pequeña.
            Dim numerador As Integer = 2
            Do While oView1H < cuartoH Or oView1W < cuartoW
                escalaFinal = numerador & ":1"
                oView1.ScaleString = escalaFinal
                'oView1H = (oView1.Height) * CInt(numerador)
                'oView1W = (oView1.Width) * CInt(numerador)
                oView1H = (oView1.Height)
                oView1W = (oView1.Width)
                numerador += 1
                'oView1.Scale -= 0.1
                DoEventsInventor()
            Loop
        End If

        '' Aplicamos la escala calculada a la DrawingView y la situaos en el centro
        'oView1.ScaleString = escalaFinal
        'oView1.Position = oPoint1
        oPoint1 = Nothing

        Do While oView1.Left > 3
            Dim pt2 As Point2d = Nothing
            pt2 = oTg.CreatePoint2d(oView1.Position.X - 1, oView1.Position.Y)
            oView1.Position = pt2
            pt2 = Nothing
            DoEventsInventor()
        Loop

        Do While oView1.Top < oSheet.Height - 3
            Dim pt2 As Point2d = Nothing
            pt2 = oTg.CreatePoint2d(oView1.Position.X, oView1.Position.Y + 1)
            oView1.Position = pt2
            pt2 = Nothing
            DoEventsInventor()
        Loop
        '' Sacamos todas las cotas de los bocetos que haya en la pieza/ensamblaje 3D
        If creacotasBase = True Then DrawingViewAcotaTodo(oView1, False)
        '********************************************************************************

        '' Los movimientos que tenemos que hacer para las vistas proyectas (la mitad de el ancho y el alto)
        Dim movX As Double = oSheet.Width / 2
        Dim movY As Double = oSheet.Height / 2.5

        '***** Crear vista derecha
        px = (oView1.Position.X + movX)
        py = oView1.Position.Y
        oPoint1 = oTg.CreatePoint2d(px, py)

        Dim oView2 As Inventor.DrawingView
        oView2 = oSheet.DrawingViews.AddProjectedView(oView1,
     oPoint1, DrawingViewStyleEnum.kFromBaseDrawingViewStyle)
        oView2.Name = "VistaDerecha"
        '' Sacamos todas las cotas de los bocetos que haya en la pieza/ensamblaje 3D
        If creacotasBase = True Then DrawingViewAcotaTodo(oView2, False)
        '********************************************************************************

        '***** Crear vista inferior
        px = (oView1.Position.X)
        py = oView1.Position.Y - movY
        oPoint1 = oTg.CreatePoint2d(px, py)

        Dim oView3 As Inventor.DrawingView
        oView3 = oSheet.DrawingViews.AddProjectedView(oView1,
     oPoint1, DrawingViewStyleEnum.kFromBaseDrawingViewStyle)
        oView3.Name = "VistaInferior"
        '' Sacamos todas las cotas de los bocetos que haya en la pieza/ensamblaje 3D
        If creacotasBase = True Then DrawingViewAcotaTodo(oView3, False)
        '********************************************************************************

        '***** Crear vista iso
        px = oView1.Position.X + movX
        py = oView1.Position.Y - movY
        oPoint1 = oTg.CreatePoint2d(px, py)

        Dim oView4 As Inventor.DrawingView
        oView4 = oSheet.DrawingViews.AddProjectedView(oView1,
     oPoint1, DrawingViewStyleEnum.kShadedDrawingViewStyle, oView1.Scale * 0.7)
        oView4.Name = "VistaIso"
        '' Sacamos todas las cotas de los bocetos que haya en la pieza/ensamblaje 3D
        If creacotasBase = True Then DrawingViewAcotaTodo(oView4, False)
        '********************************************************************************
    End Sub

    '' Devuelve un ArrayList (coleccion de Inventor.DrawingCurve)
    Public Function DrawingViewDameEntidades(ByVal vista As Inventor.DrawingView) As ArrayList
        Dim resultado As New ArrayList
        Dim mensaje As String = ""
        Dim dtc As Inventor.DrawingCurve
        For Each dtc In vista.DrawingCurves
            mensaje &= dtc.CurveType.ToString & vbCrLf
            resultado.Add(dtc)
        Next

        Return resultado
    End Function


    Public Sub DrawingViewAcotaTodo(ByVal vista As Inventor.DrawingView, Optional ByVal creaTodo As Boolean = False)
        Dim resultado As New ArrayList
        Dim mensaje As String = ""
        Dim dtc As Inventor.DrawingCurve
        For Each dtc In vista.DrawingCurves
            mensaje &= dtc.CurveType.ToString & vbCrLf
            resultado.Add(dtc)
        Next

        ' Recuperar todas las cotas de los bocetos (se puede especificar cuales recuperar)
        Call vista.Parent.DrawingDimensions.GeneralDimensions.Retrieve(vista)

        If creaTodo = False Then Exit Sub
        ' Acotar todas las lineas y arcos del dibujo
        Dim dd As Inventor.GeneralDimension
        For Each dc As DrawingCurve In vista.DrawingCurves
            'Dim gi As Inventor.GeometryIntent
            'gi.Geometry
            Dim pt1, pt2 As Double
            If dc.CenterPoint Is Nothing Then
                pt1 = dc.EndPoint.X + 2
                pt2 = dc.EndPoint.Y + 2
                dd = vista.Parent.DrawingDimensions.GeneralDimensions.AddLinear(
           oTg.CreatePoint2d(pt1, pt2), vista.Parent.CreateGeometryIntent(dc))
            Else
                pt1 = dc.CenterPoint.X + 2
                pt2 = dc.CenterPoint.Y + 2
                dd = vista.Parent.DrawingDimensions.GeneralDimensions.AddRadius(
              oTg.CreatePoint2d(pt1, pt2), vista.Parent.CreateGeometryIntent(dc))
            End If
        Next
    End Sub

    Public Sub TextoPonEnPantallaPon(ByVal queTexto As String, ByVal queTiempo As Integer)
        If oAppI.Documents.Count = 0 Then Exit Sub
        ' Set a reference to the document.
        Dim oDoc As Document
        oDoc = oAppI.ActiveDocument

        ' Set a reference to the component definition.
        ' This assumes that the active document is a part or an assembly.
        Dim oCompDef As ComponentDefinition
        oCompDef = oDoc.ComponentDefinition

        ' Attempt to get the existing client graphics object. If it exists
        ' delete it so the rest of the code can continue as if it never existed.
        Dim oClientGraphics As ClientGraphics

        Try
            oClientGraphics = oCompDef.ClientGraphicsCollection.Item("queTexto")
            oClientGraphics.Delete()
            oAppI.ActiveView.Update()
        Catch ex As Exception
            ' Create a new ClientGraphics object.
            oClientGraphics = oCompDef.ClientGraphicsCollection.Add("queTexto")
        End Try

        Try
            ' Create a graphics node.
            Dim oNode As GraphicsNode
            oNode = oClientGraphics.AddNode(1)

            ' Create text graphics.
            Dim oTextGraphics As TextGraphics
            oTextGraphics = oNode.AddTextGraphics

            ' Set the properties of the text.
            oTextGraphics.Text = queTexto
            oTextGraphics.Bold = True
            oTextGraphics.FontSize = 30
            Call oTextGraphics.PutTextColor(0, 255, 0)

            Dim oAnchorPoint As Point
            oAnchorPoint = oAppI.TransientGeometry.CreatePoint(1, 1, 1)

            ' Set the text's anchor in model space.
            oTextGraphics.Anchor = oAnchorPoint

            ' Anchor the text graphics in the view.
            Call oTextGraphics.SetViewSpaceAnchor(
            oAnchorPoint, oAppI.TransientGeometry.CreatePoint2d(30, 30), ViewLayoutEnum.kTopLeftViewCorner)

            ' Update the view to see the text.
            oAppI.ActiveView.Update()
            If queTiempo > 0 Then
                accion = "TextoPonEnPantallaBorra"
                Timer1 = New Timer
                Timer1.Interval = queTiempo
                Timer1.Start()
            End If
        Catch ex As Exception
            '' No hacemos nada.
        End Try
    End Sub
    '' Cambiar el primer texto de un PlanarSketch. O buscar un texto que contenga 'queBusco' y cambiarlo.
    Public Sub TextoCambiaPlanarSketch(ByRef quePsk As PlanarSketch,
                                   nuevoTexto As String,
                                   Optional enPrimero As Boolean = True,
                                   Optional queBusco As String = "·")
        '
        For Each queT As Inventor.TextBox In quePsk.TextBoxes
            If enPrimero Then
                queT.Text = nuevoTexto
                Exit For
            Else
                If queT.Text.Contains(queBusco) Then
                    queT.Text = nuevoTexto
                    Exit For
                End If
            End If
        Next
        quePsk.UpdateProfiles()

        'CType(oCo.ReferencedDocumentDescriptor.ReferencedDocument, Inventor.Document).Update2()
        'CType(oCo.ReferencedDocumentDescriptor.ReferencedDocument, Inventor.Document).Save2()
    End Sub
    Public Sub TextoPonEnPantallaBorra()
        If oAppI.Documents.Count = 0 Then Exit Sub
        Dim contador As Integer = 0
        Try
            ' Set a reference to the document.
            Dim oDoc As Document
            oDoc = oAppI.ActiveDocument

            ' Set a reference to the component definition.
            ' This assumes that the active document is a part or an assembly.
            Dim oCompDef As ComponentDefinition
            oCompDef = oDoc.ComponentDefinition

            ' Attempt to get the existing client graphics object. If it exists
            ' delete it so the rest of the code can continue as if it never existed.
            Dim oClientGraphics As ClientGraphics
            For Each oClientGraphics In oCompDef.ClientGraphicsCollection
                Try
                    oClientGraphics.Delete()
                    contador += 1
                Catch ex As Exception

                End Try
            Next
            If contador > 0 Then oAppI.ActiveView.Update()
        Catch ex As Exception
            '' No hacemos nada
        End Try
    End Sub



    Public Function EstilosActualiza_EnsPiePre(ByVal queDoc As Inventor.Document,
                                       ByVal queFichero As String) As String
        Dim resultado As String = ""
        Dim contadorL As Integer = 0
        Dim contadorM As Integer = 0
        Dim contadorR As Integer = 0
        Dim estababierto As Boolean = True
        oAppI.SilentOperation = True

        If queDoc Is Nothing And queFichero = "" Then
            MsgBox("Se debe especificar un Objeto Document o el Fullfilame de un fichero Pieza/Ensamblaje que exista")
            Return resultado
            Exit Function
        ElseIf queDoc Is Nothing AndAlso queFichero <> "" AndAlso IO.File.Exists(queFichero) = False Then
            MsgBox("Se debe especificar un Objeto Document o el Fullfilame de un fichero Pieza/Ensamblaje que exista")
            Return resultado
            Exit Function
        ElseIf queDoc Is Nothing AndAlso queFichero <> "" AndAlso IO.File.Exists(queFichero) = True Then
            If queFichero.ToLower.Contains("oldversions") Then
                Return resultado
                Exit Function
            End If
            estababierto = FicheroAbierto(queFichero)
            If estababierto = True Then
                queDoc = oAppI.Documents.ItemByName(queFichero)
            Else
                queDoc = oAppI.Documents.Open(queFichero, False)
            End If
        End If


        Dim EstilosL As Inventor.LightingStyles = Nothing
        Dim Materiales As Inventor.Materials = Nothing
        Dim Renders As Inventor.RenderStyles = Nothing

        If queDoc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
            Dim queEns As AssemblyDocument
            queEns = queDoc
            EstilosL = queEns.LightingStyles
            Materiales = queEns.Materials
            Renders = queEns.RenderStyles
        ElseIf queDoc.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
            Dim quePie As PartDocument
            quePie = queDoc
            EstilosL = quePie.LightingStyles
            Materiales = quePie.Materials
            Renders = quePie.RenderStyles
        ElseIf queDoc.DocumentType = DocumentTypeEnum.kPresentationDocumentObject Then
            Dim quePre As PresentationDocument
            quePre = queDoc
            EstilosL = quePre.LightingStyles
            'Materiales = quePre.Materials
            Renders = quePre.RenderStyles
        End If

        '' ***** LightinStyle (Iluminacion)
        If EstilosL IsNot Nothing Then
            Dim estilo As LightingStyle
            For Each estilo In EstilosL
                If estilo.StyleLocation = StyleLocationEnum.kLocalStyleLocation Then
                    'Debug.Print("LOCAL : " & estilo.Name & " / " & estilo.InUse)
                ElseIf estilo.StyleLocation = StyleLocationEnum.kLibraryStyleLocation Then
                    'Debug.Print("LIBRERIA : " & estilo.Name & " / " & estilo.InUse)
                ElseIf estilo.StyleLocation = StyleLocationEnum.kBothStyleLocation Then
                    'Debug.Print("AMBOS : " & estilo.Name & " / " & estilo.InUse)
                    If estilo.UpToDate = False Then
                        estilo.UpdateFromGlobal()
                        contadorL += 1
                    End If
                End If
            Next
            resultado = IIf(contadorL = 0, "", "(" & contadorL & ") Luces  /")
        End If
        '' ***** Material
        If Materiales IsNot Nothing Then
            Dim material As Inventor.Material
            For Each material In Materiales
                If material.StyleLocation = StyleLocationEnum.kLocalStyleLocation Then
                    'Debug.Print("LOCAL : " & material.Name & " / " & material.InUse)
                ElseIf material.StyleLocation = StyleLocationEnum.kLibraryStyleLocation Then
                    'Debug.Print("LIBRERIA : " & material.Name & " / " & material.InUse)
                ElseIf material.StyleLocation = StyleLocationEnum.kBothStyleLocation Then
                    'Debug.Print("AMBOS : " & material.Name & " / " & material.InUse)
                    If material.UpToDate = False Then
                        material.UpdateFromGlobal()
                        contadorM += 1
                    End If
                End If
            Next
            resultado &= IIf(contadorM = 0, "", "  (" & contadorM & ") Materiales  /")
        End If
        '' ***** RenderStyle (Color)
        If Renders IsNot Nothing Then
            Dim render As Inventor.RenderStyle
            For Each render In Renders
                If render.StyleLocation = StyleLocationEnum.kLocalStyleLocation Then
                    'Debug.Print("LOCAL : " & render.Name & " / " & render.InUse)
                ElseIf render.StyleLocation = StyleLocationEnum.kLibraryStyleLocation Then
                    'Debug.Print("LIBRERIA : " & render.Name & " / " & render.InUse)
                ElseIf render.StyleLocation = StyleLocationEnum.kBothStyleLocation Then
                    'Debug.Print("AMBOS : " & render.Name & " / " & render.InUse)
                    If render.UpToDate = False Then
                        render.UpdateFromGlobal()
                        contadorR += 1
                    End If
                End If
            Next
            resultado &= IIf(contadorR = 0, "", "  (" & contadorR & ") Colores")
        End If

        If resultado <> "" Then
            Try
                queDoc.Rebuild2()
            Catch ex As Exception
                Debug.Print("Da error Rebuild2()")
                Try
                    queDoc.Update2()
                Catch ex1 As Exception
                    Debug.Print("Da error Update2()")
                End Try
            End Try
            queDoc.Save2(False)
        End If
        If estababierto = False Then queDoc.Close(True)
        oAppI.SilentOperation = False

        Return resultado
    End Function

    Public Function EstilosActualiza_Dib(ByRef queDib As DrawingDocument,
                                       ByVal queFichero As String) As String
        Dim resultado As String = ""
        Dim contador As Integer = 0
        Dim estababierto As Boolean = True

        oAppI.SilentOperation = True
        If queDib Is Nothing And queFichero = "" Then
            MsgBox("Se debe especificar un Objeto DrawingDocument o el Fullfilame del Dibujo que tiene que existir")
            Return resultado
            Exit Function
        ElseIf queDib Is Nothing AndAlso queFichero <> "" AndAlso IO.File.Exists(queFichero) = False Then
            MsgBox("Se debe especificar un Objeto DrawingDocument o el Fullfilame del Dibujo que tiene que existir")
            Return resultado
            Exit Function
        ElseIf queDib Is Nothing AndAlso queFichero <> "" AndAlso IO.File.Exists(queFichero) = True Then
            If queFichero.ToLower.Contains("oldversions") Then
                Return resultado
                Exit Function
            End If
            estababierto = FicheroAbierto(queFichero)

            If estababierto = True Then
                queDib = oAppI.Documents.ItemByName(queFichero)
            Else
                queDib = oAppI.Documents.Open(queFichero, False)
            End If
        End If

        Dim Estilos As Inventor.Styles
        Estilos = queDib.StylesManager.Styles

        If Estilos IsNot Nothing Then
            Dim estilo As Inventor.Style
            For Each estilo In Estilos
                If estilo.StyleLocation = StyleLocationEnum.kLocalStyleLocation Then
                    'Debug.Print("LOCAL : " & estilo.Name & " / " & estilo.InUse)
                ElseIf estilo.StyleLocation = StyleLocationEnum.kLibraryStyleLocation Then
                    'Debug.Print("LIBRERIA : " & estilo.Name & " / " & estilo.InUse)
                ElseIf estilo.StyleLocation = StyleLocationEnum.kBothStyleLocation Then
                    'Debug.Print("AMBOS : " & estilo.Name & " / " & estilo.InUse)
                    If estilo.UpToDate = False Then
                        estilo.UpdateFromGlobal()
                        contador += 1
                    End If
                End If
                'estilo.UpdateFromGlobal
                'Debug.Print estilo.Name & " / " & estilo.InternalName & " / " & estilo.InUse
            Next
            resultado = IIf(contador = 0, "", "(" & contador & ") Estilos")
        End If

        If resultado <> "" Then
            Try
                queDib.Update2()
            Catch ex As Exception
                Debug.Print("Da error Update2()")
            End Try
            'ThisApplication.ActiveView.Update
            queDib.Save2(False)
        End If
        If estababierto = False Then queDib.Close(True)
        oAppI.SilentOperation = False

        Return resultado
        Exit Function
    End Function

    '' Rellena un Arraylist con los directorios del proyecto (Workspace, bibliotecas y grupos de trabajo)
    Public Sub DirectoriosProyecto(ByVal oFl As FileLocations, ByVal queArray As ArrayList)
        'Dim resultado As New ArrayList

        ' Set a reference to the FileLocations object.
        'Dim oFileLocations As FileLocations
        'oFileLocations = ThisApplication.FileLocations

        ' Display the workspace.
        'Debug.Print("Workspace: " & oFl.Workspace)

        '' **** Sólo para versiones 2010 o inferiores.
        'If oApp.SoftwareVersion.DisplayVersion <= "2010" Then
        'End If
        If oAppI.SoftwareVersion.Major <= 14 Then
            Dim asNames() As String = Nothing
            Dim asPaths() As String = Nothing

            Try
                ' Get the list of workgroup paths.
                Dim iNumWorkgroups As Long
                Call oFl.Workgroups(iNumWorkgroups, asNames, asPaths)
                If iNumWorkgroups > 0 Then
                    'Debug.Print("Workgroup Paths")
                    ' Iterate through the list of workgroups.  The array is filled
                    ' zero based, so the iteration begins a zero.
                    For i = 0 To iNumWorkgroups - 1
                        'Debug.Print("   " & asNames(i) & " = " & asPaths(i))
                        If queArray.Contains(asPaths(i)) = False Then queArray.Add(asPaths(i))
                    Next
                End If
            Catch ex As Exception
                Debug.Print(ex.Message)
            End Try

            Try
                ' Get the list of library paths.
                Dim iNumLibraries As Long
                Call oFl.Libraries(iNumLibraries, asNames, asPaths)
                If iNumLibraries > 0 Then
                    'Debug.Print("Library Paths")
                    ' Iterate through the list of libraries.  The array is filled
                    ' zero based, so the iteration begins a zero.
                    For i = 0 To iNumLibraries - 1
                        'Debug.Print("   " & asNames(i) & " = " & asPaths(i))
                        If queArray.Contains(asPaths(i)) = False Then queArray.Add(asPaths(i))
                    Next
                End If
            Catch ex As Exception
                Debug.Print(ex.Message)
            End Try
        End If

        Try
            'Debug.Print("Project File: " & oFl.FileLocationsFile)
            'Debug.Print("Directory for project file shortcuts: " & oFl.FileLocationsFilesDir)
            If queArray.Contains(oFl.FileLocationsFilesDir) = False Then queArray.Add(oFl.Workspace)
        Catch ex As Exception
            Debug.Print(ex.Message)
        End Try

        'Return resultado
    End Sub

    '' Para saber si un fichero o carpeta están dentro del proyecto actual (todos los caminos del proyecto)
    Public Function FicheroEstaEnProyecto(ByVal queF As String, ByVal arrDirProy As ArrayList) As Boolean
        Dim resultado As Boolean = False

        For Each nDir As String In arrDirProy
            If queF.StartsWith(nDir) Then
                resultado = True
                Exit For
            End If
        Next
        Return resultado
    End Function

    '***** StyleLocationEnum    (para Dibujo)
    'kBothStyleLocation  style is in both local and global locations
    'kLocalStyleLocation  style is only local
    'kLibraryStyleLocation  style is only in the library

    '***** StyleTypeEnum.
    '  kStandardStyleType = 71425
    '  kBalloonStyleType = 71426
    '  kCentermarkStyleType = 71427
    '  kDatumTargetStyleType = 71428
    '  kDimensionStyleType = 71429
    '  kFeatureControlFrameStyleType = 71430
    '  kHatchStyleType = 71431
    '  kHoleTableStyleType = 71432
    '  kIDStyleType = 71433
    '  kLayerStyleType = 71434
    '  kLeaderStyleType = 71435
    '  kObjectDefaultsStyleType = 71436
    '  kPartsListStyleType = 71437
    '  kRevisionTableStyleType = 71438
    '  kSurfaceTextureStyleType = 71439
    '  kTableStyleType = 71440
    '  kTextStyleType = 71441
    '  kViewAnnotationStyleType = 71444
    '  kWeldSymbolStyleType = 71442
    '  kWeldBeadStyleType = 71443
    '  kSheetMetalStyleType = 71445
    '  kUnfoldMethodType = 71446

    'Dim oApp AS OBJECT = Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication;
    '// Call the Quit() method of the AcadApplication object:
    'oApp.GetType().InvokeMember("QUIT", BindingFlags.InvokeMethod, null, oApp, null);
    Private Sub Timer1_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Try
            If accion = "TextoPonEnPantallaBorra" Then
                Me.TextoPonEnPantallaBorra()
            Else

            End If
        Catch ex As Exception
            '' No hacemos nada.
        Finally
            Timer1.Stop()
        End Try
    End Sub

    Public Sub BOMActiva(ByRef oEns As AssemblyDocument, Optional queVista As IEnum.EstructuraBOM = IEnum.EstructuraBOM.Piezas)
        Dim oBom As BOM = oEns.ComponentDefinition.BOM
        '' Activamos la vista elegida, si estaba desactiva.
        If queVista = IEnum.EstructuraBOM.Estructurado Then
            If oBom.StructuredViewEnabled = False Then oBom.StructuredViewEnabled = True
        ElseIf queVista = IEnum.EstructuraBOM.Piezas Then
            If oBom.PartsOnlyViewEnabled = False Then oBom.PartsOnlyViewEnabled = True
        End If
    End Sub

    ''' <summary>
    ''' Le damos objeto Document y le indicamos que propiedad queremos: "Peso", "Volumen" o "Area"
    ''' </summary>
    ''' <param name="oDoc">Objecto Inventor.Document</param>
    ''' <param name="queDoy">string "Peso", "Volumen" o "Area"</param>
    ''' <returns>una cadena con el valor de la propiedad solicitada</returns>
    ''' <remarks></remarks>
    Public Function PesoDame(oDoc As Inventor.Document, Optional queDoy As queDoy = queDoy.Peso, Optional guardarvalor As Boolean = True) As String
        Dim oControlDef As ControlDefinition = Me.oAppI.CommandManager.ControlDefinitions.Item("AppUpdateMassPropertiesCmd")
        oControlDef.Execute2(True)
        '
        Dim resultado As String = ""
        Dim oMp As MassProperties = Nothing
        If oDoc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
            oMp = CType(oDoc, AssemblyDocument).ComponentDefinition.MassProperties
        ElseIf oDoc.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
            oMp = CType(oDoc, PartDocument).ComponentDefinition.MassProperties
        End If
        '' **** PARA RECALCULAR EL PESO ****
        If oMp.AvailableAccuracy = MassPropertiesAccuracyEnum.k_VeryHigh Then
            oMp.Accuracy = MassPropertiesAccuracyEnum.k_VeryHigh
        ElseIf oMp.AvailableAccuracy = MassPropertiesAccuracyEnum.k_High Then
            oMp.Accuracy = MassPropertiesAccuracyEnum.k_High
        ElseIf oMp.AvailableAccuracy = MassPropertiesAccuracyEnum.k_Medium Then
            oMp.Accuracy = MassPropertiesAccuracyEnum.k_Medium
        End If
        'Set CacheResultsOnCompute property to False
        'Para que no se guarde Mass con el documento
        'y el documento is not 'dirtied'. (No hace falta guardarlo)
        ' True = Si guarda el resultado
        oMp.CacheResultsOnCompute = guardarvalor
        '' *********************************
        Select Case queDoy
            Case queDoy.Peso
                resultado = oMp.Mass    ' PropiedadLeeDesignTracking(oDoc, "Mass")  '(oCoGeo.ReferencedDocumentDescriptor.ReferencedDocument, "Mass")
            Case queDoy.Volumen ' "Volumen"
                resultado = oMp.Volume
            Case queDoy.Area    ' "Area"
                resultado = oMp.Area
        End Select

        Return resultado
    End Function

    Public Enum queDoy
        Peso
        Volumen
        Area
    End Enum

    ''' <summary>
    ''' Le damos objeto MassProperties del objeto y le indicamos que propiedad queremos: "Peso", "Volumen" o "Area"
    ''' </summary>
    ''' <param name="oMp">objetoc MassProperties</param>
    ''' <param name="queDoy">"Peso", "Volumen" o "Area"</param>
    ''' <returns>una cadena con el valor de la propiedad solicitada</returns>
    ''' <remarks></remarks>
    Public Function PesoDameCom(oMp As MassProperties, Optional queDoy As String = "Peso") As String
        Dim oControlDef As ControlDefinition = Me.oAppI.CommandManager.ControlDefinitions.Item("AppUpdateMassPropertiesCmd")
        oControlDef.Execute2(True)
        '
        Dim resultado As String = ""
        '' **** PARA RECALCULAR EL PESO ****
        If oMp.AvailableAccuracy = MassPropertiesAccuracyEnum.k_VeryHigh Then
            oMp.Accuracy = MassPropertiesAccuracyEnum.k_VeryHigh
        ElseIf oMp.AvailableAccuracy = MassPropertiesAccuracyEnum.k_High Then
            oMp.Accuracy = MassPropertiesAccuracyEnum.k_High
        ElseIf oMp.AvailableAccuracy = MassPropertiesAccuracyEnum.k_Medium Then
            oMp.Accuracy = MassPropertiesAccuracyEnum.k_Medium
        End If
        'Set CacheResultsOnCompute property to False
        'Para que no se guarde Mass con el documento
        'y el documento is not 'dirtied'. (No hace falta guardarlo)
        oMp.CacheResultsOnCompute = False
        '' *********************************
        Select Case queDoy
            Case "Peso"
                resultado = oMp.Mass    ' PropiedadLeeDesignTracking(oDoc, "Mass")  '(oCoGeo.ReferencedDocumentDescriptor.ReferencedDocument, "Mass")
            Case "Volumen"
                resultado = oMp.Volume
            Case "Area"
                resultado = oMp.Area
        End Select

        Return resultado
    End Function

    Public Function ProgressBarInventor(EnBarraTareas As Boolean, TotalPasos As Long, Titulo As String) As Inventor.ProgressBar

        ' Create a new ProgressBar object.
        Dim oProgressBar As Inventor.ProgressBar
        oProgressBar = oAppI.CreateProgressBar(EnBarraTareas, TotalPasos, Titulo)

        ' Set the message for the progress bar
        oProgressBar.Message = "Procesando elementos... "
        Return oProgressBar
    End Function

    Public Sub ProgressBarTestDialog(queApp As Inventor.Application)

        Dim iStepCount As Long
        iStepCount = 50

        ' Create a new ProgressBar object.
        Dim oProgressBar As Inventor.ProgressBar
        oProgressBar = queApp.CreateProgressBar(False, iStepCount, "Test Progress")

        ' Set the message for the progress bar
        oProgressBar.Message = "Executing some process"

        Dim i As Long
        For i = 1 To iStepCount
            ' Sleep 0.2 sec to simulate some process
            Retardo(1)
            oProgressBar.Message = "Executing some process - " & i
            oProgressBar.UpdateProgress()
        Next

        ' Terminate the progress bar.
        oProgressBar.Close()

    End Sub

    Public Sub ProgressBarTestStatusBar(queApp As Inventor.Application)
        Dim iStepCount As Long
        iStepCount = 50

        ' Create a new ProgressBar object.
        Dim oProgressBar As Inventor.ProgressBar
        oProgressBar = queApp.CreateProgressBar(True, iStepCount, "Test Progress")

        ' Set the message for the progress bar
        oProgressBar.Message = "Executing some process"

        Dim i As Long
        For i = 1 To iStepCount
            ' Sleep 0.2 sec to simulate some process
            Retardo(1)
            oProgressBar.Message = "Executing some process - " & i
            oProgressBar.UpdateProgress()
        Next

        ' Terminate the progress bar.
        oProgressBar.Close()
    End Sub


    '
    Public Sub ComponentoccurrenceSurfaceBodyLlenaRecursivo(ByRef oOcu As ComponentOccurrence, dicProxys As System.Collections.Generic.Dictionary(Of String, Inventor.SurfaceBody))
        '' Es una pieza o ensamblaje vacio
        For Each oSb As SurfaceBody In oOcu.SurfaceBodies
            dicProxys.Add(oOcu._DisplayName & "·" & oSb.Name, oSb)
        Next
        ''
        If oOcu.SubOccurrences IsNot Nothing AndAlso oOcu.SubOccurrences.Count > 0 Then
            '' Es un ensamblaje
            For Each oOcuHijo As ComponentOccurrence In oOcu.SubOccurrences
                ComponentoccurrenceSurfaceBodyLlenaRecursivo(oOcuHijo, dicProxys)
            Next
        End If
    End Sub
    Public Function ComponentOccurrenceOccurrenceDame(txtContiene As String) As ComponentOccurrence
        Dim resultado As ComponentOccurrence = Nothing
        If Me.oAppI.ActiveDocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
            Dim oCd As AssemblyComponentDefinition = CType(Me.oAppI.ActiveEditDocument, AssemblyDocument).ComponentDefinition
            For Each oOcu As ComponentOccurrence In oCd.Occurrences
                If oOcu.Name.ToUpper.Contains(txtContiene.ToUpper) Then
                    resultado = oOcu
                    Exit For
                End If
            Next
            'If resultado Is Nothing Then resultado = oCd.Occurrences.Item(1)
        Else
            resultado = Nothing
        End If
        ''
        Return resultado
    End Function
    ''
    Public Sub GiraComponentOccurrenceRelativo(ByRef queOccu As ComponentOccurrence,
                                   radianes As Double,
                                   vector As Inventor.Vector,
                                   base As Inventor.Point)
        ''
        queOccu.Grounded = False
        '' Si no indicamos giro, salir sin girar.
        If radianes = 0 Then
            queOccu.Grounded = True
            Exit Sub
        End If
        Dim oTg As TransientGeometry = Me.oAppI.TransientGeometry
        ''
        Dim oMat As Inventor.Matrix = queOccu.Transformation
        Dim oMatTemp As Inventor.Matrix = oTg.CreateMatrix()
        oMatTemp.SetToRotation(radianes, vector, base)
        ''
        oMat.PreMultiplyBy(oMatTemp)
        queOccu.Transformation = oMat
        oTg = Nothing
        ''
        queOccu.Grounded = True
        CType(queOccu.Parent.Document, AssemblyDocument).Rebuild2()
        'oAppCls.ActiveView.Update()
    End Sub
    ''
    Public Sub GiraComponentOccurrenceAbsoluto(ByRef queOccu As ComponentOccurrence,
                                   radianes As Double,
                                   vector As Inventor.Vector,
                                   base As Inventor.Point)
        ''
        Dim fijo As Boolean = queOccu.Grounded
        If queOccu.Grounded = True Then queOccu.Grounded = False
        Dim oTg As TransientGeometry = Me.oAppI.TransientGeometry
        Dim oMatTemp As Inventor.Matrix = oTg.CreateMatrix()
        oMatTemp.SetToRotation(radianes, vector, base)
        ''
        queOccu.Transformation = oMatTemp
        oTg = Nothing
        ''
        If fijo = True Then queOccu.Grounded = True
        CType(queOccu.Parent.Document, AssemblyDocument).Rebuild2()
        'oAppCls.ActiveView.Update()
    End Sub

    ''
    Public Sub MueveComponentOccurrenceRelativo(ByRef queOccu As ComponentOccurrence,
                                   vector As Inventor.Vector,
                                   Optional resetrotation As Boolean = False)
        ''
        Dim fijo As Boolean = queOccu.Grounded
        If queOccu.Grounded = True Then queOccu.Grounded = False
        ''
        Dim oMat As Inventor.Matrix = queOccu.Transformation
        'Dim VectorSuma As Vector = oMat.Translation
        'VectorSuma.AddVector(vector)
        'oMat.SetTranslation(VectorSuma)
        ''
        oMat.Cell(1, 4) = oMat.Cell(1, 4) + vector.X
        oMat.Cell(2, 4) = oMat.Cell(2, 4) + vector.Y
        oMat.Cell(3, 4) = oMat.Cell(3, 4) + vector.Z
        ''
        queOccu.Transformation = oMat
        ''
        If fijo = True Then queOccu.Grounded = True
        CType(queOccu.Parent.Document, AssemblyDocument).Rebuild2()
        '' Get a reference to an existing occurrence.
        'Dim oAsmDoc As AssemblyDocument
        'Set oAsmDoc = ThisApplication.ActiveDocument
        'Dim oOcc As ComponentOccurrence
        'Set oOcc = oAsmDoc.ComponentDefinition.Occurrences.ItemByName("Arrow:1")
        'Dim oTG As TransientGeometry
        'Set oTG = ThisApplication.TransientGeometry
        ''
        ''
        '' Move the occurrence to (3, 2, 1).
        'Dim oMatrix As Matrix
        'Set oMatrix = oOcc.Transformation
        'Call oMatrix.SetTranslation(oTG.CreateVector(3, 2, 1))
        'oOcc.Transformation = oMatrix
        ''
        ''
        '' Move the occurrence 5 cm in the X direction by changing the matrix directly.
        'Set oMatrix = oOcc.Transformation
        'oMatrix.Cell(1, 4) = oMatrix.Cell(1, 4) + 5
        'oOcc.Transformation = oMatrix
        '
        ' En Y = oMatrix.Cell(2, 4) = oMatrix.Cell(2, 4) + 5
        ' En Z = oMatrix.Cell(3, 4) = oMatrix.Cell(3, 4) + 5
        'oAppCls.ActiveView.Update()
    End Sub
    ''
    Public Sub MueveComponentOccurrenceAbsoluto(ByRef queOccu As ComponentOccurrence,
                                   vector As Inventor.Vector,
                                   Optional resetrotation As Boolean = False)
        ''
        Dim fijo As Boolean = queOccu.Grounded
        If queOccu.Grounded = True Then queOccu.Grounded = False
        ''
        Dim oMat As Inventor.Matrix = queOccu.Transformation
        Call oMat.SetTranslation(vector, resetrotation)
        ''
        oMat.Cell(1, 4) = vector.X
        oMat.Cell(2, 4) = vector.Y
        oMat.Cell(3, 4) = vector.Z
        ''
        queOccu.Transformation = oMat
        ''
        If fijo = True Then queOccu.Grounded = True
        CType(queOccu.Parent.Document, AssemblyDocument).Rebuild2()
        'oAppCls.ActiveView.Update()
    End Sub

#Region "UTILES"
    Public Function BotonDameControlDefinition(intNombre As String) As ButtonDefinition
        Dim resultado As ButtonDefinition = Nothing
        ''
        Dim controlDefs As ControlDefinitions = oAppI.CommandManager.ControlDefinitions
        ''
        For Each controlDef As ControlDefinition In controlDefs
            If (controlDef.InternalName = intNombre) Then
                resultado = CType(controlDef, ButtonDefinition)
                Exit For
            End If
        Next
        ''
        Return resultado
    End Function
    ''
    ' This finds iMates the entity with an specified name.  This
    ' allows to be used as a generic naming mechansim.
    Private Function GetNamedEntity(PartOccurrence As ComponentOccurrence, Name As String) As Object
        Dim resultEntity As Object = Nothing
        ' Look for the iMate that has the specified name in the referenced file.
        Dim iMate As iMateDefinition
        Dim partDef As PartComponentDefinition = PartOccurrence.Definition
        For Each iMate In partDef.iMateDefinitions
            ' Check to see if this iMate has the correct name.
            If UCase(iMate.Name) = UCase(Name) Then
                ' Get the geometry assocated with the iMate.
                Dim entity As Object = iMate
                ' Create a proxy.
                Call PartOccurrence.CreateGeometryProxy(entity, resultEntity)
                Exit For
            End If
        Next

        ' Return the found entity, or Nothing if a match wasn't found.
        Return resultEntity
    End Function
    ''
    Public Function PiezaDameCreaEnEnsamblaje(queAsm As AssemblyDocument,
                                          piezaBorrar As String,
                                          Optional borrar As Boolean = False,
                                          Optional queFullPlantilla As String = "") As ComponentOccurrence
        Dim resultado As ComponentOccurrence = Nothing
        '' Buscar y devolver el ComponentOccurrence [Nombre ASM]_BASE (Si existe)
        For Each oCo As ComponentOccurrence In queAsm.ComponentDefinition.Occurrences
            If oCo.DefinitionDocumentType = DocumentTypeEnum.kPartDocumentObject AndAlso
            IO.Path.GetFileName(oCo.ReferencedDocumentDescriptor.FullDocumentName).ToUpper = piezaBorrar.ToUpper Then
                If borrar Then
                    oCo.Delete()
                    queAsm.Save2()
                    IO.File.Delete(piezaBorrar)
                Else
                    Return oCo
                    Exit Function
                End If
            End If
        Next
        ''
        '' No existe o se ha borrado, habrá que crearlo e insertarlo.
        'queAsm.Update2()
        'queAsm.Rebuild2()
        'If queAsm.Dirty = True Then queAsm.Save2()
        '' Crear e insertar la pieza.
        Dim asmDir As String = IO.Path.GetDirectoryName(queAsm.FullFileName)
        Dim pieNameExt As String = IO.Path.GetFileName(piezaBorrar)
        Dim fullPieza As String = IO.Path.Combine(asmDir, pieNameExt)
        ''
        If queFullPlantilla = "" OrElse IO.File.Exists(queFullPlantilla) = False Then
            queFullPlantilla = oAppI.FileManager.GetTemplateFile(DocumentTypeEnum.kPartDocumentObject)
        End If

        '' Nuevo documento
        Dim oPie As PartDocument = Nothing
        oPie = oAppI.Documents.Add(DocumentTypeEnum.kPartDocumentObject, queFullPlantilla, False)
        oPie.FullFileName = fullPieza
        oPie.Save2()
        ''
        Dim options As NameValueMap = oTo.CreateNameValueMap
        Dim Position As Matrix = oTg.CreateMatrix()
        resultado = queAsm.ComponentDefinition.Occurrences.AddWithOptions(fullPieza, Position, options)
        ''
        Return resultado
    End Function
    '
    Public Function PiezaCreaEnEnsamblajeNumerada(queAsm As AssemblyDocument,
                                          quePrefijo As String,
                                          queFullPlantilla As String,
                                          Optional Position As Matrix = Nothing,
                                          Optional Bloqueada As Boolean = True) As ComponentOccurrence
        Dim resultado As ComponentOccurrence = Nothing
        '
        Dim asmDir As String = IO.Path.GetDirectoryName(queAsm.FullFileName)
        ''
        If queFullPlantilla = "" OrElse IO.File.Exists(queFullPlantilla) = False Then
            queFullPlantilla = oAppI.FileManager.GetTemplateFile(DocumentTypeEnum.kPartDocumentObject)
        End If
        '
        ' Buscar en el mismo directorio del ensamblaje piezas numeradas con el mismo prefijo "quePrefijo"
        Dim nInicio As Integer = 1
        Dim fullPieza As String = ""      ' Nuevo nombre numerado.
        For x As Integer = nInicio To 999
            Dim txtNumero As String = x.ToString.PadLeft(3, "0")
            Dim nombreext As String = quePrefijo & txtNumero & ".ipt"
            fullPieza = IO.Path.Combine(asmDir, nombreext)
            If IO.File.Exists(fullPieza) = False Then
                Exit For
            End If
        Next
        '
        ' Nuevo documento
        Dim oPie As PartDocument = Nothing
        oPie = oAppI.Documents.Add(DocumentTypeEnum.kPartDocumentObject, queFullPlantilla, False)
        oPie.FullFileName = fullPieza
        oPie.Save2()
        'oPie.Close(True)
        ''
        Dim options As NameValueMap = oTo.CreateNameValueMap
        If Position Is Nothing Then
            ' Si no hemos especificado posición, en el 0,0,0
            Position = oTg.CreateMatrix()
        End If
        resultado = queAsm.ComponentDefinition.Occurrences.Add(fullPieza, Position)
        'resultado = queAsm.ComponentDefinition.Occurrences.AddWithOptions(fullPieza, Position, options)
        If Bloqueada = True Then
            resultado.Grounded = Bloqueada
        End If
        ''
        Return resultado
    End Function
    Public Sub PiezaCreaBocetoPuntosCirculos(queOcu As ComponentOccurrence,
                                        quePie As PartDocument,
                                         dicSb As System.Collections.Generic.Dictionary(Of String, SurfaceBody))
        '' Borrar el PlanarSketch, si existía (BasePuntos)
        Dim nombresBorrar As String() = New String() {"BasePuntos", "BasePuntos2D", "BasePuntos3D"}
        'Boceto2DBorra(quePie.ComponentDefinition, "BasePuntos")
        'Boceto2DBorra(quePie.ComponentDefinition, nombreboceto2D)
        'Boceto3DBorra(quePie.ComponentDefinition, nombreboceto3D)
        ComponentOccurrencePiezaBorraCosas(quePie.ComponentDefinition, nombresBorrar)
        '' Crear el PlanarSketch "BasePuntos"
        Dim oPSketch As Inventor.PlanarSketch = Nothing
        oPSketch = quePie.ComponentDefinition.Sketches.Add(WorkPlaneBaseDame(quePie, IEnum.Datos.PlanoXY))
        If oPSketch Is Nothing Then Exit Sub
        oPSketch.Name = nombresBorrar(1)
        ''
        Dim oPSketch3D As Inventor.Sketch3D = Nothing
        oPSketch3D = quePie.ComponentDefinition.Sketches3D.Add()
        If oPSketch3D Is Nothing Then Exit Sub
        oPSketch3D.Name = nombresBorrar(2)
        ''
        '' Colecciones de objectos para puntos 3D (Z=0), puntos 3D y Lineas 3D
        Dim ocol2d As ObjectCollection = oTo.CreateObjectCollection
        Dim ocol3d As ObjectCollection = oTo.CreateObjectCollection
        Dim ocolLines As ObjectCollection = oTo.CreateObjectCollection
        ''
        Dim arrPuntos2D As New ArrayList
        '' Sacar de dicSb los SurfaceBodyProxy.
        For Each oSb As SurfaceBody In dicSb.Values
            For Each oFs As FaceShell In oSb.FaceShells
                '' No procesar los FaceShell abiertos ni interiores.
                If oFs.IsClosed = False Then Continue For
                If oFs.IsVoid Then Continue For
                ''
                For Each oFace As Face In oFs.Faces
                    '' Si tiene más de 1 EdgeLoop, continuar
                    'If oFace.EdgeLoops.Count > 1 Then Continue For
                    '' Si la cara no es cilindrica, continuar
                    If oFace.SurfaceType <> SurfaceTypeEnum.kCylinderSurface And oFace.SurfaceType <> SurfaceTypeEnum.kConeSurface Then Continue For
                    '' Si la cara no es exterior, continuar
                    'If CylindricalFaceEsExterior(oFace) = False Then Continue For
                    If FaceEsExteriorCylinderCone(oFace) = False Then Continue For
                    ''
                    For Each oEl As EdgeLoop In oFace.EdgeLoops
                        '' Si es interno, continuar.
                        If oEl.IsOuterEdgeLoop = False Then Continue For
                        For Each oEd As Edge In oEl.Edges
                            ''
                            If oEd.CurveType = CurveTypeEnum.kCircleCurve Then
                                '' Si no es un circulo cerrado, continuar
                                If oEd.StartVertex.Equals(oEd.StopVertex) = False Then Continue For
                                ''
                                Dim oEdProx As EdgeProxy = Nothing
                                queOcu.CreateGeometryProxy(oEd, CType(oEdProx, Inventor.EdgeProxy))
                                Dim oCir As Inventor.Circle = oEdProx.Curve(CurveTypeEnum.kCircleCurve)
                                Dim radio As Double = oCir.Radius
                                ''
                                Dim oPt As Point = oCir.Center
                                'Dim oPt As Point = CType(oEdProx.Curve(CurveTypeEnum.kCircleCurve), Inventor.Circle).Center
                                Dim oTg As TransientGeometry = oAppI.TransientGeometry
                                Dim oPt2D As Point2d = oTg.CreatePoint2d(oPt.X, oPt.Y)
                                ''
                                Dim clave As String = oPt2D.X.ToString & oPt2D.Y.ToString
                                '' Solo añadimos el punto si no existía
                                If arrPuntos2D.Contains(clave) = False Then
                                    arrPuntos2D.Add(clave)
                                    'Dim oSpt As SketchPoint = oPSketch.SketchPoints.Add(oTg.CreatePoint2d(oPt.X, oPt.Y))
                                    Dim oSpt3D0 As SketchPoint3D = oPSketch3D.SketchPoints3D.Add(oTg.CreatePoint(oPt.X, oPt.Y, 0))
                                    Dim oSpt3D1 As SketchPoint3D = oPSketch3D.SketchPoints3D.Add(oPt)
                                    Dim oSl3D As SketchLine3D = oPSketch3D.SketchLines3D.AddByTwoPoints(oSpt3D1, oSpt3D0)
                                    ocol3d.Add(oSpt3D1)
                                    ocol2d.Add(oSpt3D0)
                                    ocolLines.Add(oSl3D)
                                End If
                            End If
                        Next
                    Next
                Next
            Next
        Next
        '' Crear la superficie reglada.
        ' Dim pDefinition As SweepDefinition = quePie.ComponentDefinition.Features.SweepFeatures.CreateSweepDefinition()
        'quePie.ComponentDefinition.Features.SweepFeatures.Add(pDefinition)
        '' Crear SketchSplines3D
        'ObjectCollectionSketchPoint3DOrdena(ocol2d, queOcu)
        'ObjectCollectionSketchPoint3DOrdena(ocol3d, queOcu)
        Dim sp3D0 As SketchSpline3D = Nothing
        Dim sp3D1 As SketchSpline3D = Nothing
        Try
            sp3D0 = oPSketch3D.SketchSplines3D.Add(ocol2d, SplineFitMethodEnum.kACADSplineFit)
            sp3D1 = oPSketch3D.SketchSplines3D.Add(ocol3d, SplineFitMethodEnum.kACADSplineFit)
        Catch ex As Exception

        End Try
        If quePie.RequiresUpdate Then quePie.Update2()
        If quePie.Dirty Then quePie.Save2()
        oPSketch = Nothing
        oPSketch3D = Nothing
        ocol2d = Nothing
        ocol3d = Nothing
        ocolLines = Nothing
        arrPuntos2D = Nothing
        sp3D0 = Nothing
        sp3D1 = Nothing
    End Sub
    'Public Function CylindricalFaceEsExterior(oFace As Face) As Boolean
    '    Dim resultado As Boolean = True
    '    '' Evaluaremos si Face es CylinderSurface o ConeSurface
    '    If oFace.SurfaceType <> SurfaceTypeEnum.kCylinderSurface And oFace.SurfaceType <> SurfaceTypeEnum.kConeSurface Then
    '        Return False
    '        Exit Function
    '    End If
    '    ''
    '    Dim oCylinder As Object = Nothing
    '    oCylinder = oFace.Geometry
    '    '
    '    Dim params(1) As Double
    '    params(0) = 0.5
    '    params(1) = 0.5

    '    ' Get point on surface at param .5,.5
    '    Dim points(2) As Double
    '    Call oFace.Evaluator.GetPointAtParam(params, points)

    '    ' Create point object
    '    Dim oPoint As Point
    '    oPoint = oTg.CreatePoint(points(0), points(1), points(2))

    '    ' Get normal at this point
    '    Dim normals(2) As Double
    '    Call oFace.Evaluator.GetNormal(params, normals)

    '    ' Create normal vector object
    '    Dim oNormal As Vector
    '    oNormal = oTg.CreateVector(normals(0), normals(1), normals(2))

    '    ' Scale vector by radius of the cylinder
    '    'oNormal.ScaleBy(IIf(oCylinder IsNot Nothing, oCylinder.Radius, oCone.Radius))
    '    oNormal.ScaleBy(oCylinder.Radius)

    '    ' Find the sampler point on the normal by adding the
    '    ' scaled normal vector to the point at .5,.5 param.
    '    Dim oSamplePoint As Point
    '    oSamplePoint = oPoint

    '    oSamplePoint.TranslateBy(oNormal)

    '    ' Check if the sample point lies on the cylinder axis.
    '    ' If it does, we have a hollow face.

    '    ' Create a line describing the cylinder axis
    '    Dim oAxisLine As Line
    '    'oAxisLine = oTg.CreateLine(
    '    '    IIf(oCylinder IsNot Nothing, oCylinder.BasePoint, oCone.BasePoint),
    '    '    IIf(oCylinder IsNot Nothing, oCylinder.AxisVector.AsVector, oCone.AxisVector.AsVector))
    '    oAxisLine = oTg.CreateLine(oCylinder.BasePoint, oCylinder.AxisVector.AsVector)

    '    'Create a line parallel to the axis passing thru the sample point.
    '    Dim oSampleLine As Line
    '    'oSampleLine = oTg.CreateLine(oSamplePoint,
    '    '                             IIf(oCylinder IsNot Nothing, oCylinder.AxisVector.AsVector, oCone.AxisVector.AsVector))
    '    oSampleLine = oTg.CreateLine(oSamplePoint, oCylinder.AxisVector.AsVector)

    '    If oSampleLine.IsColinearTo(oAxisLine) Then
    '        resultado = False
    '    Else
    '        resultado = True
    '    End If
    '    ''
    '    Return resultado
    'End Function

    Public Function VectorBaseDame(queAsm As Inventor.AssemblyDocument, dato As IEnum.Datos) As Inventor.Vector
        Dim resultado As Inventor.Vector = Nothing
        Select Case dato
            Case IEnum.Datos.EjeX
                'resultado = oTg.CreateVector(1, 0, 0)
                resultado = WorkAxisBaseDame(queAsm, IEnum.Datos.EjeX).Line.Direction.AsVector
            Case IEnum.Datos.EjeY
                'resultado = oTg.CreateVector(0, 1, 0)
                resultado = WorkAxisBaseDame(queAsm, IEnum.Datos.EjeY).Line.Direction.AsVector
            Case IEnum.Datos.EjeZ
                'resultado = oTg.CreateVector(0, 0, 1)
                resultado = WorkAxisBaseDame(queAsm, IEnum.Datos.EjeZ).Line.Direction.AsVector
        End Select
        ''
        Return resultado
    End Function

    Public Function VectorBaseDame(quePart As Inventor.PartDocument, dato As IEnum.Datos) As Inventor.Vector
        Dim resultado As Inventor.Vector = Nothing
        Select Case dato
            Case IEnum.Datos.EjeX
                'resultado = oTg.CreateVector(1, 0, 0)
                resultado = WorkAxisBaseDame(quePart, IEnum.Datos.EjeX).Line.Direction.AsVector
            Case IEnum.Datos.EjeY
                'resultado = oTg.CreateVector(0, 1, 0)
                resultado = WorkAxisBaseDame(quePart, IEnum.Datos.EjeY).Line.Direction.AsVector
            Case IEnum.Datos.EjeZ
                'resultado = oTg.CreateVector(0, 0, 1)
                resultado = WorkAxisBaseDame(quePart, IEnum.Datos.EjeZ).Line.Direction.AsVector
        End Select
        ''
        Return resultado
    End Function
    Public Function VectorBaseDame(dato As IEnum.Datos) As Inventor.Vector
        Dim resultado As Inventor.Vector = Nothing
        Dim oTg As TransientGeometry = Me.oAppI.TransientGeometry
        Select Case dato
            Case IEnum.Datos.EjeX
                resultado = oTg.CreateVector(1, 0, 0)
            Case IEnum.Datos.EjeY
                resultado = oTg.CreateVector(0, 1, 0)
            Case IEnum.Datos.EjeZ
                resultado = oTg.CreateVector(0, 0, 1)
        End Select
        oTg = Nothing
        ''
        Return resultado
    End Function
    Public Function VectorDame(x As Double, y As Double, z As Double) As Inventor.Vector
        Dim resultado As Inventor.Vector = Nothing
        Dim oTg As TransientGeometry = Me.oAppI.TransientGeometry
        resultado = oTg.CreateVector(x, y, z)
        oTg = Nothing
        ''
        Return resultado
    End Function
    Public Function PointOrigenDame(queAsm As Inventor.AssemblyDocument) As Inventor.Point
        Return WorkPointBaseDame(queAsm).Point
    End Function
    Public Function PointOrigenDame(quePie As Inventor.PartDocument) As Inventor.Point
        Return WorkPointBaseDame(quePie).Point
    End Function
    Public Function PointOrigenDame() As Inventor.Point
        Dim resultado As Inventor.Point = Nothing
        Dim oTg As TransientGeometry = Me.oAppI.TransientGeometry
        resultado = oTg.CreatePoint(0, 0, 0)
        oTg = Nothing
        ''
        Return resultado
    End Function
    Public Function PointDame(x As Double, y As Double, z As Double) As Inventor.Point
        Dim resultado As Inventor.Point = Nothing
        Dim oTg As TransientGeometry = Me.oAppI.TransientGeometry
        resultado = oTg.CreatePoint(x, y, z)
        oTg = Nothing
        ''
        Return resultado
    End Function
    Public Function WorkPlaneBaseDame(ByRef oDoc As Inventor.Document, dato As IEnum.Datos) As Inventor.WorkPlane
        Dim resultado As Inventor.WorkPlane = Nothing
        Dim oDef As Object = Nothing
        If oDoc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
            oDef = CType(oDoc, AssemblyDocument).ComponentDefinition
        ElseIf oDoc.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
            oDef = CType(oDoc, PartDocument).ComponentDefinition
        Else
            Return Nothing
            Exit Function
        End If
        Select Case dato
            Case IEnum.Datos.PlanoXY
                resultado = oDef.WorkPlanes.Item(3)
            Case IEnum.Datos.PlanoXZ
                resultado = oDef.WorkPlanes.Item(2)
            Case IEnum.Datos.PlanoYZ
                resultado = oDef.WorkPlanes.Item(1)
        End Select
        ''
        Return resultado
    End Function
    Public Function WorkPlaneBaseDame(ByRef oAsm As AssemblyDocument, dato As IEnum.Datos) As Inventor.WorkPlane
        Dim resultado As Inventor.WorkPlane = Nothing
        Dim oAsmDef As AssemblyComponentDefinition = oAsm.ComponentDefinition
        Select Case dato
            Case IEnum.Datos.PlanoXY
                resultado = oAsmDef.WorkPlanes.Item(3)
            Case IEnum.Datos.PlanoXZ
                resultado = oAsmDef.WorkPlanes.Item(2)
            Case IEnum.Datos.PlanoYZ
                resultado = oAsmDef.WorkPlanes.Item(1)
        End Select
        ''
        Return resultado
    End Function
    Public Function WorkPlaneBaseDame(ByRef oPart As PartDocument, dato As IEnum.Datos) As Inventor.WorkPlane
        Dim resultado As Inventor.WorkPlane = Nothing
        Dim oPartDef As PartComponentDefinition = oPart.ComponentDefinition
        Select Case dato
            Case IEnum.Datos.PlanoXY
                resultado = oPartDef.WorkPlanes.Item(3)
            Case IEnum.Datos.PlanoXZ
                resultado = oPartDef.WorkPlanes.Item(2)
            Case IEnum.Datos.PlanoYZ
                resultado = oPartDef.WorkPlanes.Item(1)
        End Select
        ''
        Return resultado
    End Function
    Public Function WorkPointBaseDame(ByRef oAsm As AssemblyDocument) As Inventor.WorkPoint
        Dim resultado As Inventor.WorkPoint = Nothing
        Dim oAsmDef As AssemblyComponentDefinition = oAsm.ComponentDefinition
        resultado = oAsmDef.WorkPoints.Item(1)
        ''
        Return resultado
    End Function
    Public Function WorkPointBaseDame(ByRef oPart As PartDocument) As Inventor.WorkPoint
        Dim resultado As Inventor.WorkPoint = Nothing
        Dim oPartDef As PartComponentDefinition = oPart.ComponentDefinition
        resultado = oPartDef.WorkPoints.Item(1)
        ''
        Return resultado
    End Function
    Public Function WorkAxisBaseDame(ByRef oAsm As AssemblyDocument, dato As IEnum.Datos) As Inventor.WorkAxis
        Dim resultado As Inventor.WorkAxis = Nothing
        Dim oAsmDef As AssemblyComponentDefinition = oAsm.ComponentDefinition
        Select Case dato
            Case IEnum.Datos.EjeX
                resultado = oAsmDef.WorkAxes.Item(1)
            Case IEnum.Datos.EjeY
                resultado = oAsmDef.WorkAxes.Item(2)
            Case IEnum.Datos.EjeZ
                resultado = oAsmDef.WorkAxes.Item(3)
        End Select
        ''
        Return resultado
    End Function
    Public Function WorkAxisBaseDame(ByRef oPart As PartDocument, dato As IEnum.Datos) As Inventor.WorkAxis
        Dim resultado As Inventor.WorkAxis = Nothing
        Dim oPartDef As PartComponentDefinition = oPart.ComponentDefinition
        Select Case dato
            Case IEnum.Datos.EjeX
                resultado = oPartDef.WorkAxes.Item(1)
            Case IEnum.Datos.EjeY
                resultado = oPartDef.WorkAxes.Item(2)
            Case IEnum.Datos.EjeZ
                resultado = oPartDef.WorkAxes.Item(3)
        End Select
        ''
        Return resultado
    End Function

    'Public Function FaceDameUnitVector(oFace As Face, queDato As FaceData) As UnitVector
    '    Dim resultado As UnitVector = Nothing
    '    '
    '    Dim facePoint As Point = oFace.PointOnFace
    '    Dim surfEval As SurfaceEvaluator = oFace.Evaluator
    '    Dim points(2) As Double
    '    points(0) = facePoint.X
    '    points(1) = facePoint.Y
    '    points(2) = facePoint.Z
    '    Dim guessparams As Double() = New Double(1) {0, 0}
    '    Dim maxDeviations As Double() = New Double() {}
    '    Dim params As Double() = New Double() {}
    '    Dim solutionNatures As SolutionNatureEnum() = New SolutionNatureEnum() {}
    '    '
    '    Call surfEval.GetParamAtPoint(points, guessparams, maxDeviations, params, solutionNatures)
    '    '
    '    ' Calcular la Normal
    '    Dim normal(2) As Double
    '    Call surfEval.GetNormal(params, normal)
    '    Dim oNormal As UnitVector = oTg.CreateUnitVector(normal(0), normal(1), normal(2))
    '    '
    '    ' Calcular xDir
    '    Dim uTangents As Double() = New Double() {}
    '    Dim vTangents As Double() = New Double() {}
    '    Call surfEval.GetTangents(params, uTangents, vTangents)
    '    Dim xDir As UnitVector = oTg.CreateUnitVector(uTangents(0), uTangents(1), uTangents(2))
    '    '
    '    ' Calcular yDir
    '    Dim yDir As UnitVector = oNormal.CrossProduct(xDir)
    '    '
    '    ' Create a transform to position the text on the mid face.
    '    Dim transform As Matrix = oTg.CreateMatrix
    '    Call transform.SetCoordinateSystem(facePoint, xDir.AsVector, yDir.AsVector, oNormal.AsVector)
    '    Dim ori As Point = Nothing
    '    Dim xVec As Vector = Nothing
    '    Dim yVec As Vector = Nothing
    '    Dim zVec As Vector = Nothing
    '    transform.GetCoordinateSystem(ori, xVec, yVec, zVec)

    '    Select Case queDato
    '        Case FaceData.Normal, FaceData.DireccionZ
    '            resultado = oNormal
    '        Case FaceData.DireccionX
    '            resultado = xDir
    '        Case FaceData.DireccionY
    '            resultado = yDir
    '        Case FaceData.DireccionZMedio
    '            resultado = zVec
    '    End Select
    '    ''
    '    Return resultado
    'End Function
    ''
    'Public Function FaceDamePuntoMedio(oFace As Face) As Point
    '    Dim resultado As Point = Nothing
    '    '
    '    Dim facePoint As Point = oFace.PointOnFace
    '    Dim surfEval As SurfaceEvaluator = oFace.Evaluator
    '    Dim points(2) As Double
    '    points(0) = facePoint.X
    '    points(1) = facePoint.Y
    '    points(2) = facePoint.Z
    '    Dim guessparams As Double() = New Double(1) {0, 0}
    '    Dim maxDeviations As Double() = New Double() {}
    '    Dim params As Double() = New Double() {}
    '    Dim solutionNatures As SolutionNatureEnum() = New SolutionNatureEnum() {}
    '    '
    '    Call surfEval.GetParamAtPoint(points, guessparams, maxDeviations, params, solutionNatures)
    '    '
    '    ' Calcular la Normal
    '    Dim normal(2) As Double
    '    Call surfEval.GetNormal(params, normal)
    '    Dim oNormal As UnitVector = oTg.CreateUnitVector(normal(0), normal(1), normal(2))
    '    '
    '    ' Calcular xDir
    '    Dim uTangents As Double() = New Double() {}
    '    Dim vTangents As Double() = New Double() {}
    '    Call surfEval.GetTangents(params, uTangents, vTangents)
    '    Dim xDir As UnitVector = oTg.CreateUnitVector(uTangents(0), uTangents(1), uTangents(2))
    '    '
    '    ' Calcular yDir
    '    Dim yDir As UnitVector = oNormal.CrossProduct(xDir)
    '    '
    '    ' Create a transform to position the text on the mid face.
    '    Dim transform As Matrix = oTg.CreateMatrix
    '    Call transform.SetCoordinateSystem(facePoint, xDir.AsVector, yDir.AsVector, oNormal.AsVector)
    '    Dim ori As Point = Nothing
    '    Dim xVec As Vector = Nothing
    '    Dim yVec As Vector = Nothing
    '    Dim zVec As Vector = Nothing
    '    transform.GetCoordinateSystem(ori, xVec, yVec, zVec)
    '    '
    '    Return ori
    'End Function
    ''
    Public Sub HighlightSetObjetos(queCol As ICollection, nColor As IEnum.nombreColor, Optional queOpacity As Double = 1)
        ' Create a new highlight set for the start face(s).
        Dim oStartHLSet As HighlightSet = oAppI.ActiveEditDocument.CreateHighlightSet

        ' Change the highlight color for the set to red.
        Dim nC As Color = Nothing
        Select Case nColor
            Case IEnum.nombreColor.Rojo
                nC = oTo.CreateColor(255, 0, 0)
            Case IEnum.nombreColor.Verde
                nC = oTo.CreateColor(0, 255, 0)
            Case IEnum.nombreColor.Azul
                nC = oTo.CreateColor(0, 0, 255)
            Case IEnum.nombreColor.Blanco
                nC = oTo.CreateColor(255, 255, 255)
            Case IEnum.nombreColor.Negro
                nC = oTo.CreateColor(0, 0, 0)
        End Select

        ' Set the opacity
        nC.Opacity = queOpacity
        oStartHLSet.Color = nC
        ''
        For Each queEnt As Object In queCol
            oStartHLSet.AddItem(queEnt)
        Next
    End Sub
    ''
    Public Function FeatureExtrudeMarkCreaEnPlano(ByRef oCo As ComponentOccurrence,
                                          ancho As Double,
                                             profundidad As Double,
                                             tolerancia As Double,
                                             optMed As Point,
                                             oLineSegment As LineSegment) As ExtrudeFeature
        Dim resultado As ExtrudeFeature = Nothing
        '
        If oCo Is Nothing OrElse oCo.DefinitionDocumentType <> DocumentTypeEnum.kPartDocumentObject Then
            Return resultado
            Exit Function
        End If
        Dim oCd As PartComponentDefinition = oCo.Definition
        '
        'If ancho <= 0 Then
        '    '' Ancho de la torre (Width) o 6 si no existe Width.
        '    ancho = ParametroLee(oCd.Document, "Width")
        '    If ancho = 0 Then ancho = 10
        'End If
        ''
        Dim oWpProxy As WorkPlaneProxy = CreateGeometryProxy(oCo, oCd.WorkPlanes.Item("WP_Top"))
        ''
        If oWpProxy IsNot Nothing Then
            Dim oPSk As PlanarSketch = oCd.Sketches.Add(oWpProxy)       '' Nuevo PlanarSketch
            ''
            Dim oPt2DMed As Point2d = oPSk.ModelToSketchSpace(optMed)      '' Punto2D proyectado desde punto medio
            Dim oPt2DTop As Point2d = oPSk.ModelToSketchSpace(oLineSegment.EndPoint) '' Punto2D poryectado desde fin linea
            '
            ' Puntos base y bloquearlos
            Dim pt2DMed As SketchPoint = oPSk.SketchPoints.Add(oPt2DMed)      '' SketchPoint para insertar el bloque
            Dim pt2DTop As SketchPoint = oPSk.SketchPoints.Add(oPt2DTop)       '' SketchPoint para orientar bloque
            Call oPSk.GeometricConstraints.AddGround(pt2DMed)
            Call oPSk.GeometricConstraints.AddGround(pt2DTop)
            '
            ' Linea 1 y bloquearla
            Dim oSl1 As SketchLine = oPSk.SketchLines.AddByTwoPoints(pt2DMed, pt2DTop)  '' Para sacar la longitud y sumarle tolerancia
            oSl1.Construction = True
            Call oPSk.GeometricConstraints.AddGround(oSl1)
            CType(oCd.Document, Inventor.Document).Update2()
            '
            ' Linea 2. Poner cota y restricción
            Dim oSl2 As SketchLine = oPSk.SketchLines.AddByTwoPoints(pt2DMed, oTg.CreatePoint2d(pt2DMed.Geometry.X + 5, pt2DMed.Geometry.Y))
            oSl2.Construction = True
            Call oPSk.GeometricConstraints.AddPerpendicular(oSl1, oSl2)
            Dim cotaL1 As TwoPointDistanceDimConstraint
            cotaL1 = oPSk.DimensionConstraints.AddTwoPointDistance(oSl2.StartSketchPoint, oSl2.EndSketchPoint, DimensionOrientationEnum.kAlignedDim, oPt2DTop)
            CType(oCd.Document, Inventor.Document).Update2()
            '
            ' Linea 3. Y poner cota y restricción
            Dim oSl3 = oPSk.SketchLines.AddByTwoPoints(oSl2.EndSketchPoint, oTg.CreatePoint2d(oSl2.EndSketchPoint.Geometry.X, oSl2.EndSketchPoint.Geometry.Y + 5))
            oSl3.Construction = True
            Call oPSk.GeometricConstraints.AddPerpendicular(oSl2, oSl3)
            Dim cotaL2 As TwoPointDistanceDimConstraint
            cotaL2 = oPSk.DimensionConstraints.AddTwoPointDistance(oSl3.StartSketchPoint, oSl3.EndSketchPoint, DimensionOrientationEnum.kAlignedDim, oPt2DTop)
            CType(oCd.Document, Inventor.Document).Update2()
            '
            ' Cotas de longitud correctas para oSl2 y oSl3 (Añadimos tolerancia a ambas)
            ' MUY IMPORTANTE: Actualizar documento para que se recalculen la cotas
            cotaL1.Parameter.Value = ancho + tolerancia
            cotaL2.Parameter.Value = oSl1.Length + tolerancia
            CType(oCd.Document, Inventor.Document).Update2()
            '
            ' Crear el rectángulo y bloquear las lineas
            Dim oRec As Inventor.SketchEntitiesEnumerator
            oRec = oPSk.SketchLines.AddAsThreePointCenteredRectangle(oSl2.StartSketchPoint, oSl2.EndSketchPoint, oSl3.EndSketchPoint.Geometry)
            For Each oSl As SketchLine In oRec
                Call oPSk.GeometricConstraints.AddGround(oSl)
            Next
            '
            ' Crear la extrusión
            Dim oProf As Profile = oPSk.Profiles.AddForSolid
            Dim oExtDef As ExtrudeDefinition = oCd.Features.ExtrudeFeatures.CreateExtrudeDefinition(oProf, PartFeatureOperationEnum.kCutOperation)
            oExtDef.SetDistanceExtent(profundidad, PartFeatureExtentDirectionEnum.kNegativeExtentDirection)
            resultado = oCd.Features.ExtrudeFeatures.Add(oExtDef)
            resultado.Name = ExtrudeFeatureDameNombreNumerado(oCd, "ExtrudeMark")
            '
            ' Cambiar el color. Appearance. Lo editamos para que ActiveEditDocument lo lea.
            oCo.Edit()
            Try
                Dim oAsset As Asset = AssetApparenceCreaDame("Rojo", System.Drawing.Color.Red)
                resultado.Appearance = CType(oCd.Document, Inventor.PartDocument).AppearanceAssets.Item("Rojo")
            Catch ex As Exception
                Debug.Print(ex.ToString)
            End Try
            oCo.ExitEdit(ExitTypeEnum.kExitToTop)
            '
            ' Actualizar y guardar documento de pieza
            With CType(oCd.Document, Inventor.Document)
                .Rebuild()
                .Update2()
                .Save2()
            End With
            oAppI.ActiveEditDocument.Update2()
            oAppI.ActiveView.Update()
        End If
        '
        Return resultado
    End Function
    '
    Public Function TowerSupportCreaExtrusion(oFaceProxy As FaceProxy, oCo As ComponentOccurrence, tolerancia As Double) As ExtrudeFeature
        Dim resultado As ExtrudeFeature = Nothing
        ''
        If oAppI.ActiveEditDocument.DocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then
            MsgBox("Only for ActiveEditDocument = Async...", MsgBoxStyle.Critical, "ERROR")
            Return resultado
            Exit Function
        End If
        '
        If oFaceProxy Is Nothing Then
            Return resultado
            Exit Function
        End If
        '
        Dim oCd As PartComponentDefinition = oCo.Definition
        '
        Try
            Dim oSketch As PlanarSketch = oCd.Sketches.Add(oCd.WorkPlanes(3), False)
            '
            ' Create a proxy for the sketch in the newly created part.
            'Dim oSketchProxy As PlanarSketchProxy = CreateGeometryProxy(oCo, oSketch)
            ''
            '' Crear el objeto NonParametricBaseFeature con el FaceProxy
            Dim oParFea As NonParametricBaseFeature = BaseFeatureSurfaceBodyDame_InFace(oFaceProxy.NativeObject, oFaceProxy.ContainingOccurrence, oCo)
            If oParFea Is Nothing OrElse TypeOf oParFea.Faces(1).Geometry IsNot Plane Then
                If oParFea IsNot Nothing Then oParFea.Delete()
                oSketch.Delete()
                Return Nothing
                Exit Function
            End If
            '' Si oParFea.Faces(1).Geometry = Cylinder o Cono. Crear eje entre los 2 circulos
            Dim oFaNon As Face = oParFea.Faces(1)
            Dim oSketchEnt As SketchEntity = Nothing
            Dim oSketchEntCol As ObjectCollection = oTo.CreateObjectCollection
            Dim oSketchEntEnum As SketchEntitiesEnumerator = Nothing
            Dim oSketchEntEnum1 As SketchEntitiesEnumerator = Nothing
            If TypeOf oFaNon.Geometry Is Plane Then
                '
                '' 1.- Crear Rectangulo 2D con el Evaluator--RangeBox (Coger solo coordenadas 2D)
                Dim oBox As Box = oFaNon.Evaluator.RangeBox
                Dim ptMin2D As Point2d = oTg.CreatePoint2d(oBox.MinPoint.X, oBox.MinPoint.Y)
                Dim ptMax2D As Point2d = oTg.CreatePoint2d(oBox.MaxPoint.X, oBox.MaxPoint.Y)
                oSketchEntEnum = oSketch.SketchLines.AddAsTwoPointRectangle(ptMin2D, ptMax2D)
                ''
                '' 2.- Proyectar TODAS las entidades.
                'oSketchEntEnum = oSketchProxy.SketchLines.AddAsTwoPointRectangle(ptMin2D, ptMax2D)
                'For Each oEdge As EdgeProxy In oFaNon.Edges
                '    oSketchEnt = oSketchProxy.AddByProjectingEntity(oEdge)
                'Next
                '
                ' 3.- Proyectar solo las entidades exteriores y lineas
                'For Each oLoop As EdgeLoop In oFaNon.EdgeLoops
                '    Dim encontrado As Boolean = False
                '    If oLoop.IsOuterEdgeLoop = True Then
                '        For Each oEdge As Edge In oFaNon.Edges
                '            If oEdge.GeometryType = CurveTypeEnum.kLineSegmentCurve Then
                '                oSketchEnt = oSketch.AddByProjectingEntity(oEdge)
                '                oSketchEntCol.Add(oSketchEnt)
                '            End If
                '        Next
                '        encontrado = True
                '        Exit For
                '    End If
                '    If encontrado Then Exit For
                'Next
            End If
            ''
            '' Crear la extrusión contra la cara de oParFea
            If tolerancia > 0 Then
                For Each oSkl As SketchLine In oSketchEntEnum ' oSketchEntCol    ' oSketchEntEnum (Este si proyectamos Rectangulo 2D)
                    oSketch.GeometricConstraints.AddGround(oSkl)
                    If oSkl.Construction = False And oSkl.Centerline = False Then
                        oSketchEntCol.Add(oSkl)
                        oSkl.Construction = True
                    End If
                Next
                oSketchEntEnum1 = oSketch.OffsetSketchEntitiesUsingDistance(oSketchEntCol, tolerancia, True)
                resultado = ExtrudeFaceFeature(oParFea, oSketch) ', oSketchProxy.NativeObject)
            Else
                resultado = ExtrudeFaceFeature(oParFea, oSketch)   ', oSketchProxy.NativeObject)
            End If
            '
            oAppI.ActiveView.Update()
        Catch ex As Exception
            'MessageBox.Show(ex.ToString)
            resultado = Nothing
        End Try
        '
        Return resultado
    End Function
    '
    Public Function TowerSupportCambiaExtrusion(oFaceProxy As FaceProxy,
                                            oCo As ComponentOccurrence,
                                            nameExtrudeSupportTower As String,
                                            nameSketchSupportTower As String,
                                            tolerancia As Double) As ExtrudeFeature
        Dim oExtrudeFeature As ExtrudeFeature = Nothing
        ''
        oCo.Edit()
        '
        If oAppI.ActiveEditDocument.DocumentType <> DocumentTypeEnum.kPartDocumentObject Then
            MsgBox("Only for ActiveEditDocument = Parts...", MsgBoxStyle.Critical, "ERROR")
            Return oExtrudeFeature
            Exit Function
        End If
        '
        If oFaceProxy Is Nothing Then
            Return oExtrudeFeature
            Exit Function
        End If
        '
        Dim oCd As PartComponentDefinition = oCo.Definition
        ' Poner el final de pieza al principio.
        oCd.SetEndOfPartToTopOrBottom(True)
        ' Crear NonParametricBaseFeature con el FaceProxy recibido
        Dim oParFea As NonParametricBaseFeature = Nothing
        Try
            oParFea = BaseFeatureSurfaceBodyDame_InFace(oFaceProxy.NativeObject, oFaceProxy.ContainingOccurrence, oCo)
        Catch ex As Exception
            Return oParFea
            Exit Function
        End Try
        ' Cogemos el boceto 'SketchSupportTower"
        Dim oSketch As PlanarSketch = Nothing
        Try
            oSketch = oCd.Sketches.Item(nameSketchSupportTower)
            ' Borrar todas las entidades que tuviera.
            For Each oSkEnt As SketchEntity In oSketch.SketchEntities
                Try
                    oSkEnt.Delete()
                Catch ex As Exception
                    Continue For
                End Try
            Next
        Catch ex As Exception
            Return oSketch
            Exit Function
        End Try
        ' Cogemos la extrusion 'ExtrudeSupportTower'
        Try
            oExtrudeFeature = oCd.Features.ExtrudeFeatures.Item(nameExtrudeSupportTower)
        Catch ex As Exception
            Return oExtrudeFeature
            Exit Function
        End Try
        '
        Dim oFaNon As Face = oParFea.Faces(1)
        Try
            Dim oSketchEnt As SketchEntity = Nothing
            Dim oSketchEntCol As ObjectCollection = oTo.CreateObjectCollection
            Dim oSketchEntEnum As SketchEntitiesEnumerator = Nothing
            Dim oSketchEntEnum1 As SketchEntitiesEnumerator = Nothing
            If TypeOf oFaNon.Geometry Is Plane Then
                '
                '' 1.- Crear Rectangulo 2D con el Evaluator--RangeBox (Coger solo coordenadas 2D)
                Dim oBox As Box = oFaNon.Evaluator.RangeBox
                Dim ptMin2D As Point2d = oTg.CreatePoint2d(oBox.MinPoint.X, oBox.MinPoint.Y)
                Dim ptMax2D As Point2d = oTg.CreatePoint2d(oBox.MaxPoint.X, oBox.MaxPoint.Y)
                oSketchEntEnum = oSketch.SketchLines.AddAsTwoPointRectangle(ptMin2D, ptMax2D)
                ''
                '' 2.- Proyectar TODAS las entidades.
                'oSketchEntEnum = oSketchProxy.SketchLines.AddAsTwoPointRectangle(ptMin2D, ptMax2D)
                'For Each oEdge As EdgeProxy In oFaNon.Edges
                '    oSketchEnt = oSketchProxy.AddByProjectingEntity(oEdge)
                'Next
                '
                ' 3.- Proyectar solo las entidades exteriores y lineas
                'For Each oLoop As EdgeLoop In oFaNon.EdgeLoops
                '    Dim encontrado As Boolean = False
                '    If oLoop.IsOuterEdgeLoop = True Then
                '        For Each oEdge As Edge In oFaNon.Edges
                '            If oEdge.GeometryType = CurveTypeEnum.kLineSegmentCurve Then
                '                oSketchEnt = oSketch.AddByProjectingEntity(oEdge)
                '                oSketchEntCol.Add(oSketchEnt)
                '            End If
                '        Next
                '        encontrado = True
                '        Exit For
                '    End If
                '    If encontrado Then Exit For
                'Next
            End If
            ''
            '' Crear la extrusión contra la cara de oParFea
            If tolerancia > 0 Then
                For Each oSkl As SketchLine In oSketchEntEnum ' oSketchEntCol    ' oSketchEntEnum (Este si proyectamos Rectangulo 2D)
                    oSketch.GeometricConstraints.AddGround(oSkl)
                    oSketchEntCol.Add(oSkl)
                    oSkl.Construction = True
                Next
                oSketchEntEnum1 = oSketch.OffsetSketchEntitiesUsingDistance(oSketchEntCol, tolerancia, True)
                oSketch.UpdateProfiles()
                CType(oCd.Document, PartDocument).Rebuild2()
                CType(oCd.Document, PartDocument).Update2()
                oAppI.ActiveView.Update()
            End If
            ' Poner el final de pieza antes de la extrusión
            oExtrudeFeature.SetEndOfPart(True)
            ' Cambiar la extrusión (Definition) para el boceto y hasta cara.
            Dim oProfile As Profile = oSketch.Profiles.AddForSolid  '(True, oSketchEntCol)
            oExtrudeFeature.Definition.Profile = oProfile
            oExtrudeFeature.Definition.SetToExtent(oFaNon, True)
            'CType(oCd.Document, PartDocument).Rebuild2()
            'CType(oCd.Document, PartDocument).Update2()
            oExtrudeFeature.SetEndOfPart(False)
            CType(oCd.Document, PartDocument).Rebuild2()
            CType(oCd.Document, PartDocument).Update2()
            oAppI.ActiveView.Update()
            '
            Dim oSkCode As PlanarSketch = oCd.Sketches.Item("Sketch_Code")
            oSkCode.PlanarEntity = oExtrudeFeature.SideFaces(1)
            CType(oCd.Document, PartDocument).Rebuild2()
            CType(oCd.Document, PartDocument).Update2()
            oSkCode.SetEndOfPart(False)
            oAppI.ActiveView.Update()
        Catch ex As Exception
            'MessageBox.Show(ex.ToString)
            oExtrudeFeature = Nothing
        End Try
        ' Poner el final de pieza al final.
        oCd.SetEndOfPartToTopOrBottom(False)
        oAppI.ActiveView.Update()
        oCo.ExitEdit(ExitTypeEnum.kExitToTop)
        '
        Return oExtrudeFeature
    End Function
    Public Function AssetApparenceCreaDame(queNombre As String, queColor As System.Drawing.Color) As Asset
        Dim resultado As Asset = Nothing
        Dim doc As Document = oAppI.ActiveEditDocument
        ' Only document appearances can be edited, so that's what's created.
        ' This assumes a part or assembly document is active.
        Dim docAssets As Assets = doc.Assets

        ' Coger o Crear un nuevo appearance asset.
        Try
            resultado = docAssets.Item(queNombre)
        Catch ex As Exception
            resultado = docAssets.Add(AssetTypeEnum.kAssetTypeAppearance, "Generic", queNombre, queNombre)
            Dim color As ColorAssetValue = resultado.Item("generic_diffuse")
            color.Value = oTo.CreateColor(queColor.R, queColor.G, queColor.B)
            Dim floatValue As FloatAssetValue = resultado.Item("generic_reflectivity_at_0deg")
            floatValue.Value = 0.5
            floatValue = resultado.Item("generic_reflectivity_at_90deg")
            floatValue.Value = 0.5
            doc.Save2()
        End Try
        ''
        Return resultado
    End Function

    Public Function AssetApparenceDame(oDoc As Inventor.Document, Optional nombreAsset As String = "Bamboo") As Asset
        ' El Asset tiene que existir en el documento. Si no existe, hay que cargarlo
        ' previamente en el documento para poder utilizarlo.
        Dim localAsset As Asset = Nothing
        ''
        On Error Resume Next
        If oDoc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
            localAsset = CType(oDoc, AssemblyDocument).Assets.Item(nombreAsset)
        ElseIf oDoc.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
            localAsset = CType(oDoc, PartDocument).Assets.Item(nombreAsset)
        Else
            Return Nothing
            Exit Function
        End If
        If Err.Number <> 0 Then
            On Error GoTo 0

            ' Falla al cargar el Asset del documento. Hay que importarlo

            ' Cargar asset library por nombre.  El nombre depende del idioma,
            ' por lo que podemos cargarla tambien por el ID.
            Dim assetLib As AssetLibrary = oAppI.AssetLibraries.Item("Autodesk Appearance Library")
            'Dim assetLib = oAppCls.AssetLibraries.Item("314DE259-5443-4621-BFBD-1730C6CC9AE9")

            ' Cargar un Asset de la librería. Puede ser por nombre interno o display name
            Dim libAsset As Asset = assetLib.AppearanceAssets.Item(nombreAsset)
            'Dim libAsset = assetLib.AppearanceAssets.Item("ACADGen-082")
            '
            ' Copiar el asset localmente al documento
            localAsset = libAsset.CopyTo(oDoc)
        End If
        On Error GoTo 0

        ' Seleccionar una Occurrence (Solo en Ensamblajes y ponerle el Asset)
        'If oDoc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
        'Dim occ As ComponentOccurrence = oAppCls.CommandManager.Pick(SelectionFilterEnum.kAssemblyOccurrenceFilter, "Seleccionar un componente.")
        ' Assignarle el Asset al componente.
        'occ.Appearance = localAsset
        'End If
        '
        Return localAsset
    End Function
    Public Function PlanarSketchDameNombreNumerado(oCd As PartComponentDefinition, Optional preNombre As String = "SketchMark") As String
        Dim nuevonombre As String = ""
        For x As Integer = 1 To 100
            Dim nom As String = preNombre & x.ToString
            Dim existe As Boolean = False
            For Each oPsk As PlanarSketch In oCd.Sketches
                If oPsk.Name = nom Then
                    existe = True
                    Exit For
                End If
            Next
            '
            If existe = False Then
                nuevonombre = nom
                Exit For
            End If
        Next
        Return nuevonombre
    End Function
    Public Function ExtrudeFeatureDameNombreNumerado(oCd As PartComponentDefinition, Optional preNombre As String = "ExtrudeMark") As String
        Dim nuevonombre As String = ""
        For x As Integer = 1 To 100
            Dim nom As String = preNombre & x.ToString
            Dim existe As Boolean = False
            For Each oEf As ExtrudeFeature In oCd.Features.ExtrudeFeatures
                If oEf.Name = nom Then
                    existe = True
                    Exit For
                End If
            Next
            '
            If existe = False Then
                nuevonombre = nom
                Exit For
            End If
        Next
        Return nuevonombre
    End Function
    Public Function CreateGeometryProxy(oCo As ComponentOccurrence, queEntity As Object) As Object
        Dim queEntityTemp As Object = Nothing
        oCo.CreateGeometryProxy(queEntity, queEntityTemp)
        If queEntityTemp Is Nothing Then
            Return Nothing
        Else
            Return queEntityTemp
        End If
    End Function
#End Region
#Region "Calcular Angulos Occurrence"
    '' Sacar los angulos de componentes del ensamblajes
    'Sub GetAngles()
    '    Dim oDoc As AssemblyDocument
    'Set oDoc = ThisApplication.ActiveDocument
    'Dim oOcc As ComponentOccurrence
    'Set oOcc = oDoc.ComponentDefinition.Occurrences(1)
    'Dim oMat As Matrix
    'Set oMat = oOcc.Transformation
    'Dim aRotAngles(2) As Double
    '    Call CalculateRotationAngles(oMat, aRotAngles)
    '    ''
    '    ' Print results
    '    Dim i As Integer
    '    For i = 0 To 2
    '        Debug.Print FormatNumber(aRotAngles(i), 3)
    'Next i
    '    Beep()
    'End Sub

    Sub CalculateRotationAngles(ByVal oMatrix As Inventor.Matrix, ByRef aRotAngles() As Double)
        Const PI = 3.14159265358979
        Const TODEGREES As Double = 180 / PI
        Dim dB As Double
        Dim dC As Double
        Dim dNumer As Double
        Dim dDenom As Double
        Dim dAcosValue As Double
        Dim oRotate As Inventor.Matrix
        Dim oAxis As Inventor.Vector
        Dim oCenter As Inventor.Point
        oRotate = oAppI.TransientGeometry.CreateMatrix
        oAxis = oAppI.TransientGeometry.CreateVector
        oCenter = oAppI.TransientGeometry.CreatePoint
        oCenter.X = 0
        oCenter.Y = 0
        oCenter.Z = 0
        ' Choose aRotAngles[0] about x which transforms axes[2] onto the x-z plane
        '
        dB = oMatrix.Cell(2, 3)
        dC = oMatrix.Cell(3, 3)
        dNumer = dC
        dDenom = Math.Sqrt(dB * dB + dC * dC)
        ' Make sure we can do the division.  If not, then axes[2] is already in the x-z plane
        If (Math.Abs(dDenom) <= 0.000001) Then
            aRotAngles(0) = 0#
        Else
            If (dNumer / dDenom >= 1.0#) Then
                dAcosValue = 0#
            Else
                If (dNumer / dDenom <= -1.0#) Then
                    dAcosValue = PI
                Else
                    dAcosValue = Acos(dNumer / dDenom)
                End If
            End If
            aRotAngles(0) = Math.Sign(dB) * dAcosValue
            oAxis.X = 1
            oAxis.Y = 0
            oAxis.Z = 0
            Call oRotate.SetToRotation(aRotAngles(0), oAxis, oCenter)
            Call oMatrix.PreMultiplyBy(oRotate)
        End If
        '
        ' Choose aRotAngles[1] about y which transforms axes[3] onto the z axis
        '
        If (oMatrix.Cell(3, 3) >= 1.0#) Then
            dAcosValue = 0#
        Else
            If (oMatrix.Cell(3, 3) <= -1.0#) Then
                dAcosValue = PI
            Else
                dAcosValue = Acos(oMatrix.Cell(3, 3))
            End If
        End If
        ''
        aRotAngles(1) = Math.Sign(-oMatrix.Cell(1, 3)) * dAcosValue
        oAxis.X = 0
        oAxis.Y = 1
        oAxis.Z = 0
        Call oRotate.SetToRotation(aRotAngles(1), oAxis, oCenter)
        Call oMatrix.PreMultiplyBy(oRotate)
        '
        ' Choose aRotAngles[2] about z which transforms axes[0] onto the x axis
        '
        If (oMatrix.Cell(1, 1) >= 1.0#) Then
            dAcosValue = 0#
        Else
            If (oMatrix.Cell(1, 1) <= -1.0#) Then
                dAcosValue = PI
            Else
                dAcosValue = Acos(oMatrix.Cell(1, 1))
            End If
        End If
        ''
        aRotAngles(2) = Math.Sign(-oMatrix.Cell(2, 1)) * dAcosValue

        'if you want to get the result in degrees
        aRotAngles(0) = aRotAngles(0) * TODEGREES
        aRotAngles(1) = aRotAngles(1) * TODEGREES
        aRotAngles(2) = aRotAngles(2) * TODEGREES
    End Sub
    ''
    Public Function Acos(value As Double) As Double
        Acos = Math.Atan(-value / Math.Sqrt(-value * value + 1)) + 2 * Math.Atan(1)
    End Function
#End Region

#Region "Image Converters"



#End Region

End Class

#Region "iComparer"
Public Class clsiComparer
        Implements IComparer

        ' Calls CaseInsensitiveComparer.Compare with the parameters reversed.
        Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer _
         Implements IComparer.Compare
            Return New CaseInsensitiveComparer().Compare(y, x)
        End Function 'IComparer.Compare

    End Class 'clsiComparer
    ''
    '*** In case if arraylist of coordinates
    Public Class clsiComparerYCoordinates
        Implements IComparer
        Function Compare(x As Object, y As Object) As Integer Implements IComparer.Compare
            Dim a() As Double = DirectCast(x, Double())
            Dim b() As Double = DirectCast(y, Double())
            Dim result As Integer = a(1).CompareTo(b(1))
            Return result
        End Function
    End Class

    Public Class clsiComparerXYCoordinates
        Implements IComparer
        Function Compare(x As Object, y As Object) As Integer Implements IComparer.Compare
            Dim a() As Double = DirectCast(x, Double())
            Dim b() As Double = DirectCast(y, Double())
            Dim result As Integer = a(0).CompareTo(b(0)) Or a(1).CompareTo(b(1))
            Return result
        End Function
    End Class
    ''Usage: (sort arraylist twice)
    'first by Y
    'myarray.Sort(New PointYComparer())     'XComparer o ZComparer
    'then by X and Y
    ''
    ''
    '*** In case if arraylist of SketchPoint3D
    Public Class clsiComparerXSketchPoint3D
        Implements IComparer
        Function Compare(x As Object, y As Object) As Integer Implements IComparer.Compare
            Dim a As SketchPoint3D = DirectCast(x, SketchPoint3D)
            Dim b As SketchPoint3D = DirectCast(y, SketchPoint3D)
            Dim result As Integer = a.Geometry.X.CompareTo(b.Geometry.X)
            Return result
        End Function
    End Class
    Public Class clsiComparerYSketchPoint3D
        Implements IComparer
        Function Compare(x As Object, y As Object) As Integer Implements IComparer.Compare
            Dim a As SketchPoint3D = DirectCast(x, SketchPoint3D)
            Dim b As SketchPoint3D = DirectCast(y, SketchPoint3D)
            Dim result As Integer = a.Geometry.Y.CompareTo(b.Geometry.Y)
            Return result
        End Function
    End Class
    Public Class clsiComparerZSketchPoint3D
        Implements IComparer
        Function Compare(x As Object, y As Object) As Integer Implements IComparer.Compare
            Dim a As SketchPoint3D = DirectCast(x, SketchPoint3D)
            Dim b As SketchPoint3D = DirectCast(y, SketchPoint3D)
            Dim result As Integer = a.Geometry.Z.CompareTo(b.Geometry.Z)
            Return result
        End Function
    End Class

    Public Class clsiComparerXYSketchPoint3D
        Implements IComparer
        Function Compare(x As Object, y As Object) As Integer Implements IComparer.Compare
            Dim a As SketchPoint3D = DirectCast(x, SketchPoint3D)
            Dim b As SketchPoint3D = DirectCast(y, SketchPoint3D)
            Dim result As Integer = a.Geometry.X.CompareTo(b.Geometry.X) And a.Geometry.Y.CompareTo(b.Geometry.Y)
            Return result
        End Function
    End Class
    ''Usage: (sort twice)
    'myarray.Sort(New PointYComparer())         XComparer o ZComparer
    'myarray.Sort(New PointXYComparer())
#End Region
    Public Class PropEsEn
        Public nEs As String = ""
        Public nEn As String = ""
        Public Valor As String = ""

        Sub New(ByVal nEsp As String, ByVal nEng As String, ByVal queValor As String)
            nEs = nEsp
            nEn = nEng
            Valor = queValor
        End Sub
    End Class

#Region "hWnd Wrapper Class"
    ' This class is used to wrap a Win32 hWnd as a .Net IWind32Window =class.
    ' This is primarily used for parenting a dialog to the Inventor =window.
    '
    ' For example:
    ' myForm.Show(New =WindowWrapper(m_inventorApplication.MainFrameHWND))
    '
    ' Private Sub m_featureCountButtonDef_OnExecute( =... )
    '' Display the dialog.
    'Dim myForm As New InsertBoltForm
    'myForm.Show(New =WindowWrapper(m_inventorApplication.MainFrameHWND))
    'Sub

    Public Class WindowWrapper
        Implements System.Windows.Forms.IWin32Window

        Private _hwnd As IntPtr

        Public Sub New(ByVal handle As IntPtr)
            _hwnd = handle
        End Sub

        Public ReadOnly Property Handle() As IntPtr _
      Implements System.Windows.Forms.IWin32Window.Handle
            Get
                Return _hwnd
            End Get
        End Property

    End Class

#End Region

#Region "RIBBONS"
    'Private Function AddRibbonTab(ByVal objEnvironment As Environment, ByVal strDisplayName As String, ByVal strInternalName As String, ByVal strCLSID As String) As RibbonTab
    '    Dim objRibbonTab As RibbonTab = Nothing
    '    Try
    '        Try
    '            objRibbonTab = objEnvironment.Ribbon.RibbonTabs.Add(strDisplayName, strInternalName, strCLSID)
    '        Catch ex As Exception
    '            objRibbonTab = objEnvironment.Ribbon.RibbonTabs.Item(strInternalName)
    '        End Try
    '    Catch ex As Exception
    '        HandleException(ex)
    '    End Try
    '    Return objRibbonTab
    'End Function
    'Private Function AddRibbonPanel(ByVal objRibbonTab As RibbonTab, ByVal strDisplayName As String, ByVal strInternalName As String, ByVal strCLSID As String) As RibbonPanel
    '    Dim objRibbonPanel As RibbonPanel = Nothing
    '    Try
    '        Try
    '            objRibbonPanel = objRibbonTab.RibbonPanels.Add(strDisplayName, strInternalName, strCLSID)
    '        Catch ex As Exception
    '            objRibbonPanel = objRibbonTab.RibbonPanels.Item(strInternalName)
    '        End Try
    '    Catch ex As Exception
    '        HandleException(ex)
    '    End Try
    '    Return objRibbonPanel
    'End Function

    ' ***** RIBBONSTAB DE INVENTOR y LOS RIBBONSPANELS que tienen.
    'ZeroDoc / 8(ACTIVO)
    '    Para empezar / id_GetStarted
    '    Herramientas / id_TabTools
    '    Vault / id_TabVault
    '    Vault / id_TabVault_Upgrade
    '    Complementos / id_AddInsTab
    '    2aCAD / id_2aCAD(ACTIVO)
    '    3DA / id_Tab3DA_ZeroDoc
    '    Pretersa / id_ADAPretersa
    'Part / 36
    '    Chapa / id_TabSheetMetal
    '    Desarrollo / id_TabFlatPattern
    '    Modelo / id_TabModel
    '    Inspeccionar / id_TabInspect
    '    Herramientas / id_TabTools
    '    Administrar / id_TabManage
    '    Vista / id_TabView
    '    Entornos / id_TabEnvironments(ACTIVO)
    '    Vault / id_TabVault
    '    Vault / id_TabVault_Upgrade
    '    Para empezar / id_GetStarted
    '    Complementos / id_AddInsTab
    '    Boceto / id_TabSketch
    '    Salir de boceto 2D / id_TabSketch_Exit
    '    Boceto 3D / id_Tab3DSketch
    '    Sale de boceto 3D. / id_Tab3DSketch_Exit
    '    Construcción / id_TabConstruction
    '    Salir de construcción / id_TabConstruction_Exit
    '    Editar sólido base / id_TabEditBaseSolid
    '    Salir de sólido base / id_TabEditBaseSolid_Exit
    '    Enrutamiento / id_TabRoute
    '    Salir de enrutamiento / id_TabRoute_Exit
    '    Renderizar / id_TabRender
    '    Salir de Studio / id_TabRender_Exit
    '    Salir de análisis de tensión / id_TabStressAnalysis_Exit
    '    Análisis de tensión / id_TabAFEA
    '    Salir de análisis de tensión / id_TabAFEA_Exit
    '    Intercambio AEC / id_TabAECExchange
    '    Salir de intercambio AEC / id_TabAECExchange_Exit
    '    Volver / id_TabReturn
    '    Administrador de cambios de Fusion / FTC.Tab
    '    2aCAD / id_2aCAD
    '    Alias / id_TabInvAlias
    '    Feature Recognition / {F77AF03E-B9CD-4D3E-AA6C-F69B3BD52802}
    '    3DA / id_Tab3DA_Part
    '    Pretersa / id_ADAPretersa
    'Assembly / 40
    '    Presentación de molde / MoldTabLayout
    '    Núcleo/Cavidad / MoldTabCoreCavity
    '    Ensamblaje de molde / MoldTabMoldBase
    '    Ensamblar / id_TabAssemble
    '    Diseño / id_TabDesign
    '    Modelo / id_TabModel
    '    Soldadura / id_TabWeld
    '    Vuelta de soldadura a padre / id_TabWeld_ReturnParent
    '    Inspeccionar / id_TabInspect
    '    Herramientas / id_TabTools
    '    Administrar / id_TabManage
    '    Vista / id_TabView
    '    Entornos / id_TabEnvironments
    '    Vault / id_TabVault
    '    Vault / id_TabVault_Upgrade
    '    Para empezar / id_GetStarted(ACTIVO)
    '    Complementos / id_AddInsTab
    '    Boceto / id_TabSketch
    '    Salir de boceto 2D / id_TabSketch_Exit
    '    Tubos y tuberías / id_TabTube_Pipe
    '    Salir de tubos y tuberías / id_TabTube_Pipe_Exit
    '    Conducto de tubería / id_TabTube_Pipe_Run
    '    Salir de conducto de tubos y tuberías / id_TabTube_Pipe_Run_Exit
    '    Cable y arnés / id_TabCable_Harness
    '    Salir de cable y arnés / id_TabCable_Harness_Exit
    '    Renderizar / id_TabRender
    '    Salir de Studio / id_TabRender_Exit
    '    Simulación dinámica / id_TabSimulation
    '    Salir de simulación dinámica / id_TabSimulation_Exit
    '    Intercambio AEC / id_TabAECExchange
    '    Salir de intercambio AEC / id_TabAECExchange_Exit
    '    Volver / id_TabReturn
    '    An/ 22
    '    Insertar vistas / id_TabPlaceViews
    '    Anotar / id_TabAnnotate
    '    Anotar (ESKD) / id_TabAnnotateESKD
    '    Herramientas / id_TabTools
    '    Administrar / id_TabManage
    '    Vista / id_TabView
    '    Entornos / id_TabEnvironments
    '    Tabla de clavos / id_TabNailboard
    '    Vault / id_TabVault
    '    Vault / id_TabVault_Upgrade
    '    Para empezar / id_GetStarted
    '    Complementos / id_AddInsTab
    '    Boceto / id_TabSketch
    '    Salir de boceto 2D / id_TabSketch_Exit
    '    Salir de boceto de tabla de clavos / id_TabNailboard_Exit
    '    Revisar / id_TabReview
    '    Salir de revisión / id_TabNailboard_Review
    '    Volver / id_TabReturn
    '    2aCAD / id_2aCAD(ACTIVO)
    '    Feature Recognition / {F77AF03E-B9CD-4D3E-AA6C-F69B3BD52802}
    '    3DA / id_Tab3DA_Drawing
    '    Pretersa / id_ADAPretersa
    'Presentation / 11
    '    Presentación / id_TabManage(ACTIVO)
    '    Herramientas / id_TabTools
    '    Vista / id_TabView
    '    Entornos / id_TabEnvironments
    '    Vault / id_TabVault
    '    Vault / id_TabVault_Upgrade
    '    Para empezar / id_GetStarted
    '    Complementos / id_AddInsTab
    '    Volver / id_TabReturn
    '    Feature Recognition / {F77AF03E-B9CD-4D3E-AA6C-F69B3BD52802}
    '    3DA / id_Tab3DA_Presentation
    'iFeatures / 9
    '    iFeature / id_TabiFeature(ACTIVO)
    '    Herramientas / id_TabTools
    '    Vista / id_TabView
    '    Entornos / id_TabEnvironments
    '    Vault / id_TabVault
    '    Vault / id_TabVault_Upgrade
    '    Para empezar / id_GetStarted
    '    Complementos / id_AddInsTab
    '    Feature Recognition / {F77AF03E-B9CD-4D3E-AA6C-F69B3BD52802}
    'UnknownDocument / 10
    '    Vista personalizada / id_TabCustomView(ACTIVO)
    '    Cuaderno del ingeniero / id_TabEngineersNotebook
    '    Herramientas / id_TabEngineersNotebookTools
    '    Vista / id_TabEngineersNotebookView
    '    Vault / id_TabVault
    '    Vault / id_TabVault_Upgrade
    '    Para empezar / id_GetStarted
    '    Feature Recognition / {F77AF03E-B9CD-4D3E-AA6C-F69B3BD52802}
    '    Complementos / id_AddInsTab
    '     / id_TabView
#End Region