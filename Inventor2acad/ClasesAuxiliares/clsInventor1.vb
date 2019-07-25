Option Compare Text

Imports Inventor
'Imports Autodesk.iLogic
Imports System.Runtime.InteropServices
Imports System.Windows.forms
Imports Microsoft.Win32
Imports System.Data
''Imports System.IO
Imports Microsoft.VisualBasic
'Imports AdnInvUtilityLib
'Imports Autodesk.ADN.InvUtility

Public Class clsInventor
    Public oAppCls As Inventor.Application
    Public WithEvents oAppClsEv As Inventor.ApplicationEvents
    Public oGN As Inventor.GraphicsNode
    Public oRS As Inventor.RenderStyle = Nothing
    'Public oAsm As Inventor.AssemblyDocument
    Public oTg As Inventor.TransientGeometry
    Public cadenaMensajes As String = ""
    Public oCm As CommandManager
    Public dirProyectoInv As String = ""    '' Proyecto de Inventor que activaremos.
    Public Const nivelDetalleDefecto As String = "Desactivados"
    Public Const nivelDetalleDefectoCompleto As String = "<Desactivados>"
    Public WithEvents Timer1 As System.Windows.Forms.Timer = Nothing
    Public accion As String = ""
    Public colMatrix As ArrayList   ' Inventor.Matrix
    Public ptIAM As String = "Centro"   ' Nombre del origen del ensamblaje (punto 0,0,0)
    Public ptIAM1 As String = "Centro1"   ' Nombre de otro punto del ensamblaje (punto opuesto a ptIAM)
    Public Busquedabasica As Boolean = True

    Public Sub New(ByVal queApp As Inventor.Application)
        'Me.oAp = queapp.Application
        'Me.oTb = oAp.TransientBRep
        'Me.oSel = New clsSelect
        LlenaObjetosPrincipalesClase(queApp)
    End Sub

    Public Sub VaciaTodo()
        ' Liberar Objetos usados por esta clase, antes de cerrarla.
        If Not (oAppCls Is Nothing) Then Marshal.ReleaseComObject(oAppCls)
        oAppCls = Nothing

        If Not (oAppClsEv Is Nothing) Then Marshal.ReleaseComObject(oAppClsEv)
        oAppClsEv = Nothing

        If Not (oGN Is Nothing) Then Marshal.ReleaseComObject(oGN)
        oGN = Nothing

        If Not (oRS Is Nothing) Then Marshal.ReleaseComObject(oRS)
        oRS = Nothing

        If Not (oTg Is Nothing) Then Marshal.ReleaseComObject(oTg)
        oTg = Nothing

        System.GC.WaitForPendingFinalizers()
        System.GC.Collect()
    End Sub

    Public Sub LlenaObjetosPrincipalesClase(ByVal queApp As Inventor.Application)
        If clsIpic Is Nothing Then clsIpic = New clsiPictureToImage
        If queApp Is Nothing Then
            Try
                MyClass.oAppCls = GetObject(, "Inventor.Application")
            Catch ex As Exception
                MsgBox("Inventor no está abierto... Cerrando aplicación")
                Exit Sub
            End Try
        Else
            oAppCls = queApp
        End If
        'If oAppCls Is Nothing Then oAppCls = queApp 'oAppCls = GetObject(, "Inventor.Application")
        If (oAppCls IsNot Nothing) Then
            oAppClsEv = oAppCls.ApplicationEvents
            oTg = oAppCls.TransientGeometry
            oCm = oAppCls.CommandManager
        End If
        dirProyectoInv = oAppCls.FileLocations.Workspace
        If dirProyectoInv.EndsWith("\") = False Then dirProyectoInv &= "\"
    End Sub
    '
    Public Function UnidadesEs(queValor As String) As String
        Dim oUOM As Inventor.UnitsOfMeasure = oAppCls.UnitsOfMeasure

        Dim queUni As UnitsTypeEnum = oAppCls.ActiveEditDocument.UnitsOfMeasure.LengthUnits
        If IsNumeric(queValor) Then
            Return queValor.Replace(".", ",")
        Else
            Dim valores() As String
            valores = Split(queValor)
            Return oUOM.GetLocaleCorrectedExpression(queValor, valores(1))
        End If
    End Function

    Public Sub MensajeInventor(ByVal quetexto As String)
        Call oCm.PromptMessage(quetexto, vbOK, "Avisos PRETERSACAD")
    End Sub

    Public Function FicheroAbierto(ByVal queFichero As String) As Boolean
        Dim resultado As Boolean = False

        For Each oD As Inventor.Document In Me.oAppCls.Documents
            If oD.FullFileName = queFichero Then
                resultado = True
                Exit For
            End If
        Next
        Return resultado
    End Function

    Public Function FicheroVisible(ByVal queFichero As String) As Boolean
        Dim resultado As Boolean = False

        For Each oD As Inventor.Document In Me.oAppCls.Documents.VisibleDocuments
            If oD.FullFileName = queFichero Then
                resultado = True
                Exit For
            End If
        Next
        Return resultado
    End Function
    ''
    Public Sub CierraDocumentos(Optional queTipo As String = "*", Optional guardar As Boolean = True)
        For Each oD As Inventor.Document In Me.oAppCls.Documents
            If queTipo = "*" Then
                oD.Close(guardar)
                Continue For
            Else
                If oD.FullFileName.ToLower.EndsWith(queTipo) Then
                    oD.Close(guardar)
                Else
                    Continue For
                End If
            End If
        Next
    End Sub
    ''
    Public Sub DoEventsInventor(Optional ByVal tambienInventor As Boolean = True)
        System.Windows.Forms.Application.DoEvents()
        If tambienInventor Then oAppCls.UserInterfaceManager.DoEvents()
    End Sub

    Public Sub PropiedadesCopiadasModelo_Actualiza()
        '' UpdateCopiedModeliPropertiesCmd / Act&ualizar iProperties de modelo copiadas / Vuelve a copiar las iProperties elegidas del modelo de origen en el dibujo
        If oAppCls.ActiveDocumentType = DocumentTypeEnum.kDrawingDocumentObject Then
            Try
                oAppCls.CommandManager.ControlDefinitions.Item("UpdateCopiedModeliPropertiesCmd").Execute2(True)
            Catch ex As Exception
                '' Si da error sería que no tiene propiedades copiadas.
                '' Si da error sería que no tiene propiedades copiadas.
            End Try
        End If
    End Sub

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
                If IO.Path.GetExtension(f).EndsWith("dwg") AndAlso oAppCls.FileManager.IsInventorDWG(f) = False Then Continue For
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
            If IO.Path.GetExtension(f).EndsWith("dwg") AndAlso oAppCls.FileManager.IsInventorDWG(f) = False Then Continue For
            If arrPlanos.Contains(f) = False Then arrPlanos.Add(f)
            'End If
        Next
        ficheros = Nothing

        '' Ahora abrimos cada plano de "arrPlanos" y buscaremos en su vista principal (ParentView)
        '' Si el documento que refleja "queFichero" es el FullFilenamo del objeto Documento que contiene.
        Dim oDib As DrawingDocument = Nothing

        oAppCls.SilentOperation = True
        For Each queF As String In arrPlanos
            ' Crear un nuevo NameValueMap object
            Dim oDocOpenOptions As NameValueMap
            oDocOpenOptions = oAppCls.TransientObjects.CreateNameValueMap
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
                oDib = oAppCls.Documents.OpenWithOptions(queF, oDocOpenOptions, False)
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
                'MsgBox("ExisteFicheroPlanoEnDirProfundoInv : " & vbCrLf & vbCrLf & ex.Message)
                Continue For
            End Try
        Next
        oAppCls.SilentOperation = True

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
    ''
    Public Sub RibbonTabActiva(ByVal queNombreRibbon As String)
        Try
            If Me.oAppCls.UserInterfaceManager.ActiveEnvironment.Ribbon.RibbonTabs.Item(queNombreRibbon).Active = False Then _
            Me.oAppCls.UserInterfaceManager.ActiveEnvironment.Ribbon.RibbonTabs.Item(queNombreRibbon).Active = True
        Catch ex As Exception
            Debug.Print("El Ribbon " & queNombreRibbon & " no existe en entorno " & Me.oAppCls.UserInterfaceManager.ActiveEnvironment.DisplayName)
        End Try
    End Sub
    ''
    Public Function DameSurfBody(ByRef oC As ComponentOccurrence) As SurfaceBodyProxy
        Dim gn As SurfaceBody = oC.Definition.SurfaceBodies(1)
        Dim gn1 As SurfaceBodyProxy = Nothing

        oC.CreateGeometryProxy(gn, CType(gn1, Object))

        Return gn1
    End Function
    ''
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


    Public Sub MueveComponenteAbsoluto(ByVal oC As ComponentOccurrence, _
                                       Optional ByVal queX As Double = 0, _
                                       Optional ByVal queY As Double = 0, _
                                       Optional ByVal queZ As Double = 0)
        '' Si no hemos puesto valores. Salimos
        If queX + queY + queZ = 0 Then Exit Sub
        '' Si ya está en la nueva posición. Salimos
        If oC.Transformation.Translation.X = queX And _
            oC.Transformation.Translation.Y = queY And _
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

    Public Function ParametroLeeString(ByVal queCom As Inventor.Document, ByVal quePar As String, Optional queF As String = "") As String
        Dim resultado As String = Nothing
        Dim oPar As Parameter = Nothing
        '' Por si le damos fullFilename (queF) en vez de Document
        Dim estabaabierto As Boolean = True
        If queCom Is Nothing AndAlso queF <> "" AndAlso IO.File.Exists(queF) = True Then
            estabaabierto = FicheroAbierto(queF)
            oAppCls.SilentOperation = True
            If estabaabierto = True Then
                queCom = oAppCls.Documents.ItemByName(queF)
            Else
                queCom = oAppCls.Documents.Open(queF, False)
            End If
            oAppCls.SilentOperation = False
        End If
        '' ***********************************************
        Try
            If queCom.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                oPar = CType(queCom, AssemblyDocument).ComponentDefinition.Parameters.Item(quePar)
            ElseIf queCom.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
                oPar = CType(queCom, PartDocument).ComponentDefinition.Parameters.Item(quePar)
            End If
            resultado = oPar.Value
        Catch ex As Exception
            'MsgBox("Error ParametroASMLee. El parametro (" & quePar & ") no existe.")
            resultado = ""
        End Try
        If estabaabierto = False Then queCom.Close(True)
        ParametroLeeString = resultado
        Exit Function
    End Function
    '
    Public Function ParametroLeeDouble(ByVal queCom As Inventor.Document, ByVal quePar As String, Optional queF As String = "") As Double
        Dim resultado As Double = Nothing
        Dim oPar As Parameter = Nothing
        '' Por si le damos fullFilename (queF) en vez de Document
        Dim estabaabierto As Boolean = True
        If queCom Is Nothing AndAlso queF <> "" AndAlso IO.File.Exists(queF) = True Then
            estabaabierto = FicheroAbierto(queF)
            oAppCls.SilentOperation = True
            If estabaabierto = True Then
                queCom = oAppCls.Documents.ItemByName(queF)
            Else
                queCom = oAppCls.Documents.Open(queF, False)
            End If
            oAppCls.SilentOperation = False
        End If
        '' ***********************************************
        Try
            If queCom.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                oPar = CType(queCom, AssemblyDocument).ComponentDefinition.Parameters.Item(quePar)
            ElseIf queCom.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
                oPar = CType(queCom, PartDocument).ComponentDefinition.Parameters.Item(quePar)
            End If
            resultado = oPar.Value
        Catch ex As Exception
            'MsgBox("Error ParametroASMLee. El parametro (" & quePar & ") no existe.")
            resultado = 0
        End Try
        If estabaabierto = False Then queCom.Close(True)
        ParametroLeeDouble = resultado
        Exit Function
    End Function

    Public Function ParametroLeeBoolean(ByVal queCom As Inventor.Document, ByVal quePar As String, Optional queF As String = "") As Boolean
        Dim resultado As Boolean = False
        Dim oPar As Parameter = Nothing
        '' Por si le damos fullFilename (queF) en vez de Document
        Dim estabaabierto As Boolean = True
        If queCom Is Nothing AndAlso queF <> "" AndAlso IO.File.Exists(queF) = True Then
            estabaabierto = FicheroAbierto(queF)
            oAppCls.SilentOperation = True
            If estabaabierto = True Then
                queCom = oAppCls.Documents.ItemByName(queF)
            Else
                queCom = oAppCls.Documents.Open(queF, False)
            End If
            oAppCls.SilentOperation = False
        End If
        '' ***********************************************
        Try
            If queCom.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                oPar = CType(queCom, AssemblyDocument).ComponentDefinition.Parameters.Item(quePar)
            ElseIf queCom.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
                oPar = CType(queCom, PartDocument).ComponentDefinition.Parameters.Item(quePar)
            End If
            resultado = oPar.Value
        Catch ex As Exception
            'MsgBox("Error ParametroASMLee. El parametro (" & quePar & ") no existe.")
            resultado = False
        End Try
        If estabaabierto = False Then queCom.Close(True)
        ParametroLeeBoolean = resultado
        Exit Function
    End Function
    Public Sub ParametroEscribe(ByVal queDoc As Inventor.Document, ByVal queFi As String, ByVal quePar As String, ByVal queVal As Object, Optional ByVal queOperacion As OperacionValor = OperacionValor.cambiar, Optional ByVal cerrar As Boolean = False)
        If queFi <> "" AndAlso IO.File.Exists(queFi) Then
            oAppCls.SilentOperation = True
            queDoc = oAppCls.Documents.Open(queFi, False)
            oAppCls.SilentOperation = False
        End If
        oAppCls.ScreenUpdating = False
        ' queVal vendrá siempre en cm. Ya cambiamos a mm si procede.
        Dim oPar As Parameter = Nothing
        Try
            If queDoc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                oPar = CType(queDoc, AssemblyDocument).ComponentDefinition.Parameters.Item(quePar)
            ElseIf queDoc.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
                oPar = CType(queDoc, PartDocument).ComponentDefinition.Parameters.Item(quePar)
            End If
        Catch ex As Exception
            'MsgBox("Error ParametroASMEscribe. El parametro (" & quePar & ") no existe. O valor (" & queVal.ToString & ") incorrecto.")
            Debug.Print("Error ParametroASMEscribe. El parametro (" & quePar & ") no existe. O valor (" & queVal.ToString & ") incorrecto.")
            Exit Sub
        End Try
        '' Solo lo actualizaremos si es un parámetro modificable.
        If oPar.ParameterType <> ParameterTypeEnum.kDerivedParameter And
            oPar.ParameterType <> ParameterTypeEnum.kReferenceParameter And
            oPar.ParameterType <> ParameterTypeEnum.kTableParameter Then 'And _
            'IsNumeric(Left(oPar.Expression, 1)) = True Then
            ' Dim valor As Object = queVal
            Select Case queOperacion
                Case OperacionValor.cambiar
                    If oPar.Value.ToString <> queVal.ToString Then
                        If IsNumeric(queVal) And oPar.ModelValueType <> ModelValueTypeEnum.kNominalValue Then
                            If oPar.Value <> CDbl(queVal) Then oPar.Value = CDbl(queVal)
                        ElseIf oPar.Value <> queVal And oPar.ModelValueType = ModelValueTypeEnum.kNominalValue Then
                            oPar.Value = queVal
                        Else
                            Try
                                If oPar.Expression <> queVal.ToString Then oPar.Expression = queVal.ToString
                            Catch ex As Exception
                                If log Then PonLog("Error en ParametroEscribe con parametro " & oPar.Name)
                            End Try
                        End If
                    End If
                Case OperacionValor.sumar
                    If IsNumeric(queVal) Then
                        If oPar.Value <> oPar.Value + CDbl(queVal) Then _
                            oPar.Value = oPar.Value + CDbl(queVal)
                    End If
                Case OperacionValor.restar
                    If IsNumeric(queVal) Then
                        If oPar.Value <> oPar.Value - CDbl(queVal) Then _
                            oPar.Value = oPar.Value - CDbl(queVal)
                    End If
            End Select
        End If

        oPar = Nothing
        If queDoc.RequiresUpdate Then queDoc.Update2()
        If cerrar = True Then
            queDoc.Save2()
            queDoc.Close(True)
        End If

        oAppCls.ScreenUpdating = True
    End Sub

    Public Sub ParametrosEscriteTODOS(ByVal quePieza As PartDocument, ByVal queFila As DataRow)

        oAppCls.ScreenUpdating = False
        '' Hemos configurado las propiedades de algunas operaciones para que se activen
        '' si el parámetro que las controla es superior a 0,1 (los valores 0 los convertimos a 0,1)
        Dim arrExternas As New ArrayList
        '' Con este bucle cambiamos TODOS los parámetros que se llamen igual que los campos BD.
        'For Each oP As Inventor.Parameter In quePieza.ComponentDefinition.Parameters
        'Debug.Print(oP.Name & " / " & oP.Expression)
        'Next
        '' RECORREMOS TODOS LOS PARÁMETROS
        For Each queCol As DataColumn In queFila.Table.Columns
            Dim quePar As UserParameter
            'Dim encontrado As Boolean = True
            Dim valor As Object = queFila.Item(queCol.ColumnName)
            'Debug.Print(queFila.Table.Columns.Item(queCol.ColumnName).DataType.ToString)

            If IsDBNull(valor) Then valor = 0
            Try
                '' Parametro de la BD a parametro Inventor.
                Dim nombreReal As String = ""
                Select Case queCol.ColumnName
                    Case "COM_DIMX"
                        nombreReal = "ds_lar"
                    Case "COM_DIMY"
                        nombreReal = "ds_anc"
                    Case "COM_DIMZ"
                        nombreReal = "ds_alt"
                        '' La altura será "COM_DIMZ" + "lo_emp"
                        valor = CDbl(queFila.Item(queCol.ColumnName)) + CDbl(queFila.Item("lo_emp"))
                    Case Else
                        nombreReal = queCol.ColumnName.Trim
                End Select

                quePar = quePieza.ComponentDefinition.Parameters.UserParameters.Item(nombreReal)
                '' Si el parámetro se ha externalizado lo guardamos en la colección para
                '' no sobrescribir su valor cuando escribamos las iProperties (bucle siguiente)
                If quePar.ExposedAsProperty = True AndAlso arrExternas.Contains(quePar.Name) = False Then arrExternas.Add(quePar.Name)


                '' Si los valores son iguales. No hacemos nada con el valor.
                If valor = quePar.Value Then Continue For
                If quePar.Units = "gr" Then Continue For
                'If IsNumeric(quePar.Expression) = False Then Continue For

                If valor = 0 And quePar.Units = "su" Then
                    If quePar.Expression <> "0 " & quePar.Units Then _
                        quePar.Expression = "0 " & quePar.Units
                    'quePar.Value = 1
                ElseIf valor = 0 Then
                    If quePar.Expression <> "0,1 " & " " & quePar.Units Then _
                        quePar.Expression = "0,1 " & " " & quePar.Units
                    'quePar.Value = 0.1
                Else
                    If quePar.Expression <> FormatNumber(valor, 2, , , Microsoft.VisualBasic.TriState.False) & " " & quePar.Units Then _
                        quePar.Expression = FormatNumber(valor, 2, , , Microsoft.VisualBasic.TriState.False) & " " & quePar.Units
                    'quePar.Expression = Format(valor, "f") & " " & quePar.Units
                    'quePar.Value = FormatNumber(valor, 2)
                End If
                'oAp.UserInterfaceManager.DoEvents()
            Catch ex As Exception
                '' Si no existe un parámetro=nombre del campo pasamos al siguiente campo.
                Continue For
            End Try
        Next
        'oAp.UserInterfaceManager.DoEvents()
        '' RECORREMOS TODAS LAS PROPIEDADES
        For Each queCol As DataColumn In queFila.Table.Columns
            Dim quePro As Inventor.Property
            Dim valor As Object = queFila.Item(queCol.ColumnName)
            If IsDBNull(valor) Then valor = ""
            Try
                '' Propiedad de la BD a parametro Inventor.
                If arrExternas.Contains(queCol.ColumnName) = False Then
                    quePro = quePieza.PropertySets.Item("{D5CDD505-2E9C-101B-9397-08002B2CF9AE}").Item(queCol.ColumnName)
                    If Trim(valor) <> quePro.Value Then
                        'Debug.Print("Propiedad --> " & queCol.ColumnName & " = " & Trim(queFila.Item(queCol.ColumnName)))
                        quePro.Value = Trim(valor)
                    End If
                Else
                    Continue For
                End If
                'oAp.UserInterfaceManager.DoEvents()
            Catch ex As Exception
                '' Si no existe un parámetro=nombre del campo pasamos al siguiente campo.
                Continue For
            End Try
            '' Si el campo BD era un Parametro o una Propiedad, pas
        Next
        'oAp.UserInterfaceManager.DoEvents()
        If Trim(queFila.Item("COM_PLANTI")).ToLower = "pi" Then
            '' Para activar o desactivar la operación de pilar redondo o de esquina.
            Dim resultado As Object = queFila.Item("sec_pie")
            If IsDBNull(resultado) = False AndAlso Trim(queFila.Item("sec_pie")) = "R" Then
                Try
                    Dim pfLat As PartFeature = quePieza.ComponentDefinition.Features.Item("PilarLat Redondo")
                    Dim pfEsq1 As PartFeature = quePieza.ComponentDefinition.Features.Item("PilarEsq Redondo1")
                    If IsDBNull(queFila.Item("pos_nav")) = False AndAlso Trim(queFila.Item("pos_nav")) = "L" Then
                        pfLat.Suppressed = False
                        pfEsq1.Suppressed = True
                    ElseIf IsDBNull(queFila.Item("pos_nav")) = False AndAlso Trim(queFila.Item("pos_nav")) = "E" Then
                        pfLat.Suppressed = True
                        pfEsq1.Suppressed = False
                    Else
                        pfLat.Suppressed = True
                        pfEsq1.Suppressed = True
                    End If
                Catch ex As Exception
                    '' No existe esta Feature. No hacemos nada
                    Debug.Print("No existen las operaciones...")
                End Try
            End If
        End If
        If quePieza.RequiresUpdate = True Then quePieza.Update2()
        oAppCls.ScreenUpdating = True
    End Sub

    '' Le indicamos nombre completo del fichero Inventor y colDatos de la clase clsDatosFila (arrFilas(COM_CLAVE))
    Public Sub ParametrosEscribeTODOSCaminoHash(ByVal queFichero As String, ByVal colDatosCls As Hashtable)
        '' Hemos configurado las propiedades de algunas operaciones para que se activen
        '' si el parámetro que las controla es superior a 0,1 (los valores 0 los convertimos a 0,1)
        '' En este arraylist tendremos los nombres de los Parametros que se externalizan como Propiedades
        '' para no volver a escribirlos cuando pongamos las propiedades.
        Dim oDoc As Inventor.Document = Nothing

        If Dir(queFichero) = "" Then
            MsgBox("El fichero " & queFichero & vbCrLf & "NO EXISTE...")
            Exit Sub
        End If

        Dim EstabaAbierto As Boolean = Me.FicheroAbierto(queFichero)

        Me.oAppCls.SilentOperation = True

        If EstabaAbierto = True Then
            oDoc = Me.oAppCls.Documents.ItemByName(queFichero)
        Else
            oDoc = Me.oAppCls.Documents.Open(queFichero, False)
        End If
        ParametrosEscribeTODOSCaminoHashDoc(oDoc, colDatosCls)

        If oDoc IsNot Nothing AndAlso EstabaAbierto = False Then oDoc.Close(True)

        If Not (oDoc Is Nothing) Then Marshal.ReleaseComObject(oDoc)
        oDoc = Nothing

        System.GC.Collect()
        System.GC.WaitForPendingFinalizers()
        System.GC.Collect()
    End Sub


    '' Le indicamos nombre completo del fichero Inventor y colDatos de la clase clsDatosFila (arrFilas(COM_CLAVE))
    Public Sub ParametrosEscribeTODOSCaminoHashDoc(ByRef queDoc As Inventor.Document, ByVal colDatosCls As Hashtable)
        '' Hemos configurado las propiedades de algunas operaciones para que se activen
        '' si el parámetro que las controla es superior a 0,1 (los valores 0 los convertimos a 0,1)
        '' En este arraylist tendremos los nombres de los Parametros que se externalizan como Propiedades
        '' para no volver a escribirlos cuando pongamos las propiedades.
        Dim arrExternas As New ArrayList
        Dim oUps As Inventor.UserParameters = Nothing
        Dim nP As String = ""


        Me.oAppCls.SilentOperation = True
        oAppCls.ScreenUpdating = False

        '' ***** RECORREMOS TODOS LOS PARÁMETROS

        If queDoc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
            'oDoc = CType(oDoc, Inventor.AssemblyDocument)
            oUps = CType(queDoc, Inventor.AssemblyDocument).ComponentDefinition.Parameters.UserParameters
        ElseIf queDoc.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
            'oDoc = CType(oDoc, Inventor.PartDocument)
            oUps = CType(queDoc, Inventor.PartDocument).ComponentDefinition.Parameters.UserParameters
        End If

        For Each quePar As UserParameter In oUps
            nP = quePar.Name
            '' ***** Si no existe el nombre en colDatosCls, pasamos al siguiente parámetro.
            If Not colDatosCls.ContainsKey(quePar.Name) Then
                'If quePar.Name <> "COM_DIMX" And quePar.Name <> "COM_DIMY" And quePar.Name <> "COM_DIMZ" Then
                Continue For
                'End If
            End If
            Dim valor As Object
            Try
                valor = colDatosCls(quePar.Name)
            Catch ex As Exception
                Continue For
            End Try

            Dim parEx As Object = quePar.Expression
            Dim parExSin As String = Trim(quePar.Expression.Replace(quePar.Units, ""))

            '' ***** TODO ESTO PARA SABER SI LO RELLENAMOS O PASAMOS AL SIGUIENTE *****

            '' Si tiene una expresion, pasamos al siguiente.
            'If parEx IsNot Nothing Then Continue For
            Try
                If IsNumeric(parExSin) = False Then Continue For
                If quePar.Expression.Contains("_") Then Continue For
                If IsNumeric(quePar.Expression.Substring(0, 1)) = False Then Continue For

                If IsDBNull(valor) AndAlso quePar.Value.GetTypeCode = TypeCode.String Then valor = ""
                If IsDBNull(valor) AndAlso _
                (quePar.Value.GetTypeCode = TypeCode.Decimal Or quePar.Value.GetTypeCode = TypeCode.Double) _
                Then valor = 0

                If valor.ToString = "" Then Continue For
                If valor = quePar.Value Then Continue For
                '' No cambiamos los diámetros si vienen con valor 0 o ""
                If nP.StartsWith("di_") And IsNumeric(valor) AndAlso valor = 0 Then Continue For
                If nP.StartsWith("di_") And IsNumeric(valor) = False AndAlso valor = "" Then Continue For
                '' ***************************************************************************************
                '' Si el parámetro se ha externalizado lo guardamos en la colección para
                '' no sobrescribir su valor cuando escribamos las iProperties (bucle siguiente)
                If quePar.ExposedAsProperty = False Then quePar.ExposedAsProperty = True
                If arrExternas.Contains(quePar.Name) = False Then arrExternas.Add(quePar.Name)

                If quePar.Units <> "su" Then
                    If quePar.Value.GetTypeCode = TypeCode.String AndAlso quePar.Value = "0" Then _
                        If quePar.Value <> "0,01" Then _
                            quePar.Value = "0,01"
                    If quePar.Value.GetTypeCode = TypeCode.Decimal AndAlso quePar.Value = 0 Then _
                        If quePar.Value <> 0.01 Then _
                            quePar.Value = 0.01
                    If quePar.Value.GetTypeCode = TypeCode.Double AndAlso quePar.Value = 0 Then _
                        If quePar.Value <> 0.01 Then _
                            quePar.Value = 0.01
                    '' Valor que viene de colDatosCls, para evaluarlo
                End If
            Catch ex As Exception
                ''***** LOG PARA CONTROL DE ERRORES *****
                If log Then PonLog(vbCrLf & "Error en ParametrosEscribeTODOSCaminoHashDoc con " & nP & " y valor " & valor.ToString & vbCrLf)
                ''*****************************************
                MsgBox("Error en ParametrosEscribeTODOSCaminoHashDoc. Al evaluar valores")
                Continue For
            End Try

            Try
                '' Parametro de la BD a parametro Inventor.
                'Dim nombreReal As String = quePar.Name
                'If quePar.Name = "ds_alt" Then
                'If quePar.Value <> CDbl(colDatosCls("ds_alt")) Then quePar.Value = CDbl(colDatosCls("AltTot")) ' AltTol ya tiene el valor calculado previamente. CDbl(valor) + CDbl(colDatosCls("lo_emp"))
                'Continue For
                'End If

                '' Si los valores son iguales. No hacemos nada con el valor.
                If valor = quePar.Value Then Continue For
                If valor.ToString = quePar.Expression Then Continue For
                If quePar.Units = "gr" Then Continue For

                If valor = 0 And quePar.Units = "su" Then
                    If quePar.Expression <> ("0 " & quePar.Units) Then _
                        quePar.Expression = ("0 " & quePar.Units)
                    'quePar.Value = 1
                ElseIf valor = 0 And quePar.Units = "cm" Then
                    'If quePar.Expression <> ("0,01 " & quePar.Units) Then quePar.Expression = ("0,01 " & quePar.Units)
                    If quePar.Value <> 0.01 Then _
                        quePar.Value = 0.01
                ElseIf valor = 0 And quePar.Units = "mm" Then
                    'If quePar.Expression <> ("0,01 " & quePar.Units) Then quePar.Expression = ("0,01 " & quePar.Units)
                    If quePar.Value <> 0.1 Then _
                        quePar.Value = 0.1
                    'ElseIf quePar.Name = "ds_lar" Then
                    'If quePar.Value <> CDbl(colDatosCls("ds_lar")) Then quePar.Value = CDbl(colDatosCls("ds_lar"))
                Else
                    'If quePar.Expression <> (FormatNumber(valor, 2, , , Microsoft.VisualBasic.TriState.False) & " " & quePar.Units) Then _
                    'quePar.Expression = FormatNumber(valor, 2, , , Microsoft.VisualBasic.TriState.False) & " " & quePar.Units
                    If quePar.Expression <> valor & " " & quePar.Units Then _
                        quePar.Expression = valor & " " & quePar.Units
                    'quePar.Expression = Format(valor, "f") & " " & quePar.Units
                    'quePar.Value = FormatNumber(valor, 2)
                End If
                'oAp.UserInterfaceManager.DoEvents()
            Catch ex As Exception
                ''***** LOG PARA CONTROL DE ERRORES *****
                If log Then PonLog(vbCrLf & "Error en ParametrosEscribeTODOSCaminoHashDoc con parametro (" & nP & ") y valor " & valor.ToString & vbCrLf)
                ''*****************************************
                'MsgBox("Error en ParametrosEscribeTODOSCaminoHash. Con parametro (" & nP & ") con el valor --> " & valor.ToString)
                Continue For
            End Try
            Me.DoEventsInventor(True)
            'oAppCls.UserInterfaceManager.DoEvents()
        Next
        'oAp.UserInterfaceManager.DoEvents()
        '' RECORREMOS TODAS LAS PROPIEDADES de Usuario
        Dim oPS As PropertySet = queDoc.PropertySets.Item("{D5CDD505-2E9C-101B-9397-08002B2CF9AE}")
        For Each oPro As Inventor.Property In oPS
            nP = oPro.Name
            'Dim nP As String = oPro.Name
            Dim valor As Object = Nothing
            Try
                'If oPro.Expression IsNot Nothing Then Continue For
                If arrExternas.Contains(oPro.Name) = True Then Continue For
                If colDatosCls.ContainsKey(oPro.Name) = False Then Continue For
                If oPro.Expression <> oPro.Value Then Continue For

                valor = colDatosCls(oPro.Name)
                If IsDBNull(valor) Then valor = ""

                If valor.ToString = "" Then Continue For
                If oPro.Value.ToString = valor.ToString Then Continue For

                oPro.Value = Trim(valor.ToString)
            Catch ex As Exception
                ''***** LOG PARA CONTROL DE ERRORES *****
                If log Then PonLog(vbCrLf & "Error en ParametrosEscribeTODOSCaminoHashDoc con propiedad (" & nP & ") y valor " & valor.ToString & vbCrLf)
                ''*****************************************
                'MsgBox("Error en ParametrosEscribeTODOSCaminoHashDoc. Con propiedad (" & nP & ") con el valor --> " & valor.ToString)
                Continue For
            End Try
            Me.DoEventsInventor(True)
        Next
        queDoc.Update2()
        queDoc.Save2(False)

        If Not (oPS Is Nothing) Then Marshal.ReleaseComObject(oPS)
        oPS = Nothing
        If Not (oUps Is Nothing) Then Marshal.ReleaseComObject(oUps)
        oUps = Nothing

        System.GC.Collect()
        System.GC.WaitForPendingFinalizers()
        System.GC.Collect()

        oAppCls.ScreenUpdating = True
        Me.oAppCls.SilentOperation = False
    End Sub


    Public Sub ParametrosEscribeHijos(ByVal docPadreCamino As String, ByVal docHijo As PartDocument)
        '' "docPadre" es el ensamblaje PADRE origen (el que tiene en item(1) la pieza con todos los parámetros)
        '' Buscaremos en él todos los documentos referenciados (no incluir item(1))
        '' "docHijo" contiene todos los Userparameters a crear/cambiar en el resto.
        '' "pOrigen" el la colección de parametros originales a crear/cambiar referencias de "docPadre"
        Dim pOrigen As UserParameters = docHijo.ComponentDefinition.Parameters.UserParameters
        Dim dirModificar As String = modUtilidades.DameParteCamino(docPadreCamino, ParteCamino.CaminoConFicheroSinExtensionBarra) ' docPadre.FullFileName.Replace(".iam", "\")
        If IO.Directory.Exists(dirModificar) = False Then Exit Sub

        For Each fichero As String In IO.Directory.GetFiles(dirModificar, "*.i*", IO.SearchOption.TopDirectoryOnly)
            If fichero.ToLower.EndsWith(".iam") Or fichero.ToLower.EndsWith(".ipt") Then
                Me.ParametrosEscribeHijosUserParameters(fichero, pOrigen)
            End If
            oAppCls.UserInterfaceManager.DoEvents()
        Next
    End Sub

    Public Sub ParametrosEscribeHijosUserParameters(ByVal queF As String, ByVal pOrigen As UserParameters)

        oAppCls.ScreenUpdating = False
        Dim oD As Inventor.Document = Nothing
        Try
            oD = oAppCls.Documents.Open(queF, False)
        Catch ex As Exception
            '' Si da error al abrir salimos fuera, porque no será de inventor. O está bloqueado.
            Exit Sub
        End Try

        Me.PropiedadEscribe(oD, "Nº de pieza", modUtilidades.DameParteCamino(oD.FullFileName, ParteCamino.SoloFicheroSinExtension))
        Dim oPs As UserParameters = Nothing
        If oD.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
            oPs = CType(oD, AssemblyDocument).ComponentDefinition.Parameters.UserParameters
        ElseIf oD.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
            oPs = CType(oD, PartDocument).ComponentDefinition.Parameters.UserParameters
        End If

        For Each oP As UserParameter In oPs
            'oAp.UserInterfaceManager.DoEvents()
            Try
                If oP.Units <> pOrigen.Item(oP.Name).Units Then _
                    oP.Units = pOrigen.Item(oP.Name).Units
                If oP.Expression <> pOrigen.Item(oP.Name).Expression Then _
                    oP.Expression = pOrigen.Item(oP.Name).Expression
            Catch ex As Exception
                Continue For
                '' No hacemos nada. El parametro no existe
            End Try
        Next
        If oD.RequiresUpdate Then oD.Update2()

        oAppCls.ScreenUpdating = True
        'oD.Close()  ' Cerramos guardando todo
    End Sub
    '' Para poner la expresion de un parámetro dentro de una fórmula.
    Public Sub ParametroPonFormula(ByRef quePie As Inventor.PartDocument, ByVal nombreP As String, ByVal formula As String)
        If quePie Is Nothing Then Exit Sub
        If quePie.ComponentDefinition.Parameters.UserParameters Is Nothing Then Exit Sub
        If quePie.ComponentDefinition.Parameters.UserParameters.Count = 0 Then Exit Sub
        oAppCls.ScreenUpdating = False
        Try
            Dim queP As Inventor.UserParameter = quePie.ComponentDefinition.Parameters.UserParameters.Item(nombreP)
            ''***** Comprobamos si existe el parámetro nombreP Ejemplo: "ds_lar_tot"
            '' y si tiene la fórmula correcta formula(valores) Ejemplo: "floor(valores sumanos)"
            If Not queP.Expression.Contains(formula) Then
                'isolate(floor(largo + alto + ancho);su;mm)
                Dim expresion As String = "isolate(" & formula & "(" & queP.Expression & ");su;" & queP.Units & ")"
                queP.Expression = expresion
            End If
            quePie.Save2(False)
            ''********************************************************
        Catch ex As Exception
            '' No existe el parámetro en UserParameters.
            Exit Sub
        Finally
            oAppCls.ScreenUpdating = True
        End Try
    End Sub
    ''Summary Information, {F29F85E0-4FF9-1068-AB91-08002B27B3D9}
    ''Document Summary Information, {D5CDD502-2E9C-101B-9397-08002B2CF9AE}
    ''Design Tracking Properties, {32853F0F-3444-11D1-9E93-0060B03C1CA6}
    ''User Defined Properties, {D5CDD505-2E9C-101B-9397-08002B2CF9AE}
    Public Function PropiedadLeeTodasInventorArray(ByVal queDoc As String) As Object()
        Dim resultado(2) As Object
        Dim colEn As New Hashtable(StringComparer.OrdinalIgnoreCase)
        Dim colEs As New Hashtable(StringComparer.OrdinalIgnoreCase)
        Dim imagen As System.Drawing.Image = Nothing
        Dim estababierto As Boolean = True

        oAppCls.SilentOperation = True

        estababierto = FicheroAbierto(queDoc)

        Dim queDocSin As String = ""
        If queDoc.ToUpper.EndsWith(".iam") = True Then
            queDocSin = Me.DameCaminoSinComponentesInventor(queDoc)
        Else    'If queDoc.EndsWith(".ipt") Then
            queDocSin = queDoc
        End If

        ' Abrir un documento.
        Dim oDoc As Inventor.Document = Nothing
        Dim oProSs As PropertySets = Nothing

        Try
            If estababierto = True Then
                oDoc = oAppCls.Documents.ItemByName(queDoc)
            Else
                oDoc = oAppCls.Documents.Open(queDocSin, False)
            End If
            oDoc = oAppCls.Documents.Open(queDocSin, False)
            oProSs = oDoc.PropertySets
        Catch ex As Exception
            MsgBox("Error : " & vbCrLf & vbCrLf & ex.Message)
        End Try

        Try
            For Each oProS As PropertySet In oProSs
                For Each oPro As Inventor.Property In oProS
                    If oPro.Name = "Thumbnail" Or oPro.DisplayName = "Miniatura" Then
                        imagen = DameImagen(oPro.Value)
                    Else
                        If colEs.ContainsKey(oPro.DisplayName) Then _
                            colEs.Add(oPro.DisplayName, oPro.Value.ToString)
                        If colEn.ContainsKey(oPro.Name) Then _
                            colEn.Add(oPro.Name, oPro.Value.ToString)
                    End If
                Next
            Next
        Catch ex As Exception
            MsgBox("Error PropiedadLeeTodasApprenticeArray...")
        End Try

        If estababierto = False Then oDoc.Close(True)

        ''Liberar Objetos.
        If Not (oDoc Is Nothing) Then Marshal.ReleaseComObject(oDoc)
        oDoc = Nothing
        If Not (oProSs Is Nothing) Then Marshal.ReleaseComObject(oProSs)
        oProSs = Nothing

        System.GC.WaitForPendingFinalizers()
        System.GC.Collect()
        oAppCls.SilentOperation = False
        '' Guardamos en el Array los 3 valores Es, En, imagen
        resultado(0) = colEs
        resultado(1) = colEn
        resultado(2) = imagen

        Return resultado
        Exit Function
    End Function

    Public Function PropiedadLeeCategoria(ByVal queDoc As Inventor.Document, Optional queF As String = "", Optional ByRef quePss As PropertySets = Nothing) As String
        '' Por si le damos fullFilename (queF) en vez de Document
        Dim estabaabierto As Boolean = True
        If queDoc Is Nothing AndAlso queF <> "" AndAlso IO.File.Exists(queF) = True Then
            estabaabierto = FicheroAbierto(queF)
            oAppCls.SilentOperation = True
            If estabaabierto = True Then
                queDoc = oAppCls.Documents.ItemByName(queF)
            Else
                queDoc = oAppCls.Documents.Open(queF, False)
            End If
            oAppCls.SilentOperation = False
        End If
        '' ***********************************************
        '' Lee un valor de texto en una iProperty de usuario. Si no existe la crea con valor "".
        Dim resultado As String = ""
        '' Información resumen documento Inventor / Información resumen documento Inventor
        '' Internal name: {D5CDD502-2E9C-101B-9397-08002B2CF9AE}
        '' Nombre: Category (Categoría) / Valor:  / Id: 2

        Dim oProS As PropertySet = queDoc.PropertySets.Item("{D5CDD502-2E9C-101B-9397-08002B2CF9AE}")
        Dim oPro As Inventor.Property = oProS.ItemByPropId(2)
        resultado = oPro.Value.ToString
        If estabaabierto = False Then queDoc.Close(True)
        PropiedadLeeCategoria = resultado
        Exit Function
    End Function

    Public Sub PropiedadEscribe(ByRef queCom As Inventor.Document, ByVal quePro As String, ByVal queVal As Object, Optional queF As String = "")
        '' Por si le damos fullFilename (queF) en vez de Document
        Dim estabaabierto As Boolean = True
        If queCom Is Nothing AndAlso queF <> "" AndAlso IO.File.Exists(queF) = True Then
            estabaabierto = FicheroAbierto(queF)
            oAppCls.SilentOperation = True
            If estabaabierto = True Then
                queCom = oAppCls.Documents.ItemByName(queF)
            Else
                queCom = oAppCls.Documents.Open(queF, False)
            End If
            oAppCls.SilentOperation = False
        End If
        '' ***********************************************
        Dim oProSs As PropertySets = queCom.PropertySets

        Try
            For Each oProS As PropertySet In oProSs
                For Each oPro As Inventor.Property In oProS
                    If oPro.Name = quePro Or oPro.DisplayName = quePro Then
                        '' Usaremos esto como expresión. Por si es Nothing oPro.Expression.
                        Dim oProExp As String = IIf(oPro.Expression Is Nothing, "", oPro.Expression)

                        '' Solo escribiremos el valor si: la propiedad no es una expresion o le mandamos una nueva expresion y si
                        '' no es una expresión pero es diferente.
                        If queVal = oPro.Value Or queVal = oProExp Then
                            GoTo FINAL
                        ElseIf queVal.ToString.StartsWith("=") = True And queVal <> oProExp Then
                            oPro.Expression = queVal.ToString
                            GoTo FINAL
                        ElseIf queVal.ToString.StartsWith("=") = True And queVal = oProExp Then
                            GoTo FINAL
                        ElseIf oProExp.StartsWith("=") Then     '' Si hay una expresión ya. No hacemos nada.
                            GoTo FINAL
                        ElseIf queVal <> oPro.Value Then
                            oPro.Value = queVal
                            GoTo FINAL
                        End If
                    End If
                Next
            Next
        Catch ex As Exception
            ''***** LOG PARA CONTROL DE ERRORES *****
            If log Then PonLog(vbCrLf & "Error PropiedadEscribe. El parametro (" & quePro & ") no existe. O valor (" & queVal.ToString & ") incorrecto." & vbCrLf)
            ''*****************************************
            'MsgBox("Error PropiedadEscribe. El parametro (" & quePro & ") no existe. O valor (" & queVal.ToString & ") incorrecto.")
        End Try
FINAL:
        Try
            If queCom.Dirty = True Then queCom.Save2(False)
            If estabaabierto = False Then queCom.Close(True)
        Catch ex As Exception
            ' No hacemos nada.
            Debug.Print(ex.Message)
        End Try
    End Sub

    Public Sub PropiedadEscribeUsuario(ByRef queDoc As Inventor.Document,
                                       ByVal quePro As String,
                                       ByVal queVal As Object,
                                       Optional ByVal queFi As String = "",
                                       Optional ByVal cerrar As Boolean = False,
                                       Optional CREAR As Boolean = True,
                                       Optional sobrescribir As Boolean = False)
        If (queVal = "" Or queVal Is Nothing) And CREAR = False Then Exit Sub

        Dim estabaabierto As Boolean = False
        If queDoc Is Nothing Then
            If FicheroAbierto(queFi) = False Then
                queDoc = oAppCls.Documents.Open(queFi, False)
                estabaabierto = False
            Else
                queDoc = oAppCls.Documents.ItemByName(queFi)
                estabaabierto = True
            End If
        End If
        '' Escribe un valor de texto en una iProperty. Si no existe la crea con valor "".
        Dim oProS As PropertySet = queDoc.PropertySets.Item("{D5CDD505-2E9C-101B-9397-08002B2CF9AE}")
        Dim oPro As Inventor.Property = Nothing

        Try
            oPro = oProS.Item(quePro)
        Catch ex As Exception
            If CREAR = True Then
                oPro = oProS.Add(queVal.ToString, quePro)
                GoTo FINAL
            Else
                GoTo FINAL
            End If
        End Try
        '' Usaremos esto como expresión. Por si es Nothing oPro.Expression.
        Dim oProExp As String = IIf(oPro.Expression Is Nothing, "", oPro.Expression)
        '' Si sobrescribir=true. Siempre sobrescribimos el valor
        If sobrescribir Then
            Try
                oPro.Expression = queVal.ToString
                GoTo FINAL
            Catch ex As Exception
                If log Then PonLog("Error en PropiedadEscribeUsuario con oPro.Expression")
                GoTo FINAL
            End Try
        End If
        '' Solo escribiremos el valor si: la propiedad no es una expresion o le mandamos una nueva expresion y si
        '' no es una expresión pero es diferente.
        If queVal.ToString.StartsWith("=") = True AndAlso queVal <> oProExp Then
            Try
                oPro.Expression = queVal.ToString
                GoTo FINAL
            Catch ex As Exception
                If log Then PonLog("Error en PropiedadEscribeUsuario con oPro.Expression")
                GoTo FINAL
            End Try
        ElseIf queVal.ToString.StartsWith("=") = True AndAlso queVal = oProExp Then
            GoTo FINAL
        ElseIf oProExp.StartsWith("=") Then     '' Si hay una expresión ya. No hacemos nada.
            GoTo FINAL
        ElseIf queVal.Equals(oPro.Value) = False Then
            oPro.Value = queVal
            GoTo FINAL
        ElseIf queVal = oPro.Value Then
            GoTo FINAL
        End If

FINAL:
        oAppCls.SilentOperation = True
        If estabaabierto = False And cerrar = True Then
            '' Si queremos que se actualice y se guarde. Lo quitamos para ganar tiempo. Lo haremos al final.
            Try
                If queDoc.RequiresUpdate Then queDoc.Update2()
            Catch ex As Exception
                '' Continuamos.
            End Try
            Try
                If queDoc.Dirty = True Then queDoc.Save2(False)
            Catch ex As Exception
                '' No lo guardamos si da error.
            End Try
            Try
                queDoc.Close(True)
            Catch ex As Exception
                '' No hacemos nada. Lo dejamos abierto.
            End Try
        End If
        oPro = Nothing : oProS = Nothing
        oAppCls.SilentOperation = False
    End Sub

    Public Sub PropiedadBorraUsuario(ByRef queDoc As Inventor.Document,
                                       ByVal quePro As String,
                                       Optional ByVal queFi As String = "",
                                       Optional ByVal cerrar As Boolean = False)
        Dim estabaabierto As Boolean = False
        If queDoc Is Nothing Then
            If FicheroAbierto(queFi) = False Then
                queDoc = oAppCls.Documents.Open(queFi, False)
                estabaabierto = False
            Else
                queDoc = oAppCls.Documents.ItemByName(queFi)
                estabaabierto = True
            End If
        End If
        '' Escribe un valor de texto en una iProperty. Si no existe la crea con valor "".
        Dim oProS As PropertySet = queDoc.PropertySets.Item("{D5CDD505-2E9C-101B-9397-08002B2CF9AE}")
        Dim oPro As Inventor.Property = Nothing
        Try
            oPro = oProS.Item(quePro)
            oPro.Delete()
        Catch ex As Exception
            GoTo FINAL
        End Try
FINAL:

        If estabaabierto = False And cerrar = True Then
            '' Si queremos que se actualice y se guarde. Lo quitamos para ganar tiempo. Lo haremos al final.
            Try
                If queDoc.RequiresUpdate Then queDoc.Update2()
            Catch ex As Exception
                '' Continuamos.
            End Try
            Try

                If queDoc.Dirty = True Then queDoc.Save2(False)
            Catch ex As Exception
                '' No lo guardamos si da error.
            End Try
            Try
                queDoc.Close(True)
            Catch ex As Exception
                '' No hacemos nada. Lo dejamos abierto.
            End Try
        End If
        oPro = Nothing : oProS = Nothing
    End Sub

    '''' Summary Information / Información de resumen de Inventor
    '' Internal name: {F29F85E0-4FF9-1068-AB91-08002B27B3D9}
    Public Sub PropiedadEscribeInformacionResumenInventor(ByRef queDoc As Inventor.Document, ByVal quePro As String, ByVal queVal As Object, Optional ByVal queFi As String = "", Optional ByVal cerrar As Boolean = False)
        If queVal = "" Or queVal Is Nothing Then Exit Sub

        Dim estabaabierto As Boolean = False
        If queDoc Is Nothing Then
            If FicheroAbierto(queFi) = False Then
                queDoc = oAppCls.Documents.Open(queFi, False)
                estabaabierto = False
            Else
                queDoc = oAppCls.Documents.ItemByName(queFi)
                estabaabierto = True
            End If
        End If
        '' Escribe un valor de texto en una iProperty. Si no existe la crea con valor "".
        Dim oProS As PropertySet = queDoc.PropertySets.Item("{F29F85E0-4FF9-1068-AB91-08002B27B3D9}")
        Dim oPro As Inventor.Property = Nothing

        Try
            oPro = oProS.Item(quePro)
        Catch ex As Exception
            '' No hacemos nada. Aquí no se pueden crear propiedades nuevas.
        End Try
        '' Usaremos esto como expresión. Por si es Nothing oPro.Expression.
        Dim oProExp As String = IIf(oPro.Expression Is Nothing, "", oPro.Expression)

        '' Solo escribiremos el valor si: la propiedad no es una expresion o le mandamos una nueva expresion y si
        '' no es una expresión pero es diferente.
        If queVal.ToString.StartsWith("=") = True And queVal <> oProExp Then
            'oPro.Expression = queVal.ToString
            oPro.Expression = queVal.ToString
            GoTo FINAL
        ElseIf queVal.ToString.StartsWith("=") = True And queVal = oProExp Then
            GoTo FINAL
        ElseIf oProExp.StartsWith("=") Then     '' Si hay una expresión ya. No hacemos nada.
            GoTo FINAL
        ElseIf queVal <> oPro.Value Then
            oPro.Value = queVal
            GoTo FINAL
        ElseIf queVal = oPro.Value Then
            GoTo FINAL
        End If

FINAL:
        If estabaabierto = False And cerrar = True Then
            '' Si queremos que se actualice y se guarde. Lo quitamos para ganar tiempo. Lo haremos al final.
            Try
                If queDoc.RequiresUpdate Then queDoc.Update2()
            Catch ex As Exception
                '' Continuamos.
                Debug.Print("error")
            End Try
            Try

                If queDoc.Dirty = True Then queDoc.Save2(False)
            Catch ex As Exception
                '' No lo guardamos si da error.
                Debug.Print("error")
            End Try
            Try
                queDoc.Close(True)
            Catch ex As Exception
                '' No hacemos nada. Lo dejamos abierto.
                Debug.Print("error")
            End Try
        End If
        oPro = Nothing : oProS = Nothing
    End Sub

    '' Document Summary Information / Información resumen documento Inventor
    '' Internal name: {D5CDD502-2E9C-101B-9397-08002B2CF9AE}
    Public Sub PropiedadEscribeInformacionResumenDocumentoInventor(ByRef queDoc As Inventor.Document, ByVal quePro As String, ByVal queVal As Object, Optional ByVal queFi As String = "", Optional ByVal cerrar As Boolean = False)
        If queVal = "" Or queVal Is Nothing Then Exit Sub

        Dim estabaabierto As Boolean = False
        If queDoc Is Nothing Then
            If FicheroAbierto(queFi) = False Then
                queDoc = oAppCls.Documents.Open(queFi, False)
                estabaabierto = False
            Else
                queDoc = oAppCls.Documents.ItemByName(queFi)
                estabaabierto = True
            End If
        End If
        '' Escribe un valor de texto en una iProperty. Si no existe la crea con valor "".
        Dim oProS As PropertySet = queDoc.PropertySets.Item("{D5CDD502-2E9C-101B-9397-08002B2CF9AE}")
        Dim oPro As Inventor.Property = Nothing

        Try
            oPro = oProS.Item(quePro)
        Catch ex As Exception
            '' No hacemos nada. Aquí no se pueden crear propiedades nuevas.
        End Try
        '' Usaremos esto como expresión. Por si es Nothing oPro.Expression.
        Dim oProExp As String = IIf(oPro.Expression Is Nothing, "", oPro.Expression)

        '' Solo escribiremos el valor si: la propiedad no es una expresion o le mandamos una nueva expresion y si
        '' no es una expresión pero es diferente.
        If queVal.ToString.StartsWith("=") = True And queVal <> oProExp Then
            'oPro.Expression = queVal.ToString
            oPro.Expression = queVal.ToString
            GoTo FINAL
        ElseIf queVal.ToString.StartsWith("=") = True And queVal = oProExp Then
            GoTo FINAL
        ElseIf oProExp.StartsWith("=") Then     '' Si hay una expresión ya. No hacemos nada.
            GoTo FINAL
        ElseIf queVal <> oPro.Value Then
            oPro.Value = queVal
            GoTo FINAL
        ElseIf queVal = oPro.Value Then
            GoTo FINAL
        End If

FINAL:
        If estabaabierto = False And cerrar = True Then
            '' Si queremos que se actualice y se guarde. Lo quitamos para ganar tiempo. Lo haremos al final.
            Try
                If queDoc.RequiresUpdate Then queDoc.Update2()
            Catch ex As Exception
                '' Continuamos.
                Debug.Print("error")
            End Try
            Try

                If queDoc.Dirty = True Then queDoc.Save2(False)
            Catch ex As Exception
                '' No lo guardamos si da error.
                Debug.Print("error")
            End Try
            Try
                queDoc.Close(True)
            Catch ex As Exception
                '' No hacemos nada. Lo dejamos abierto.
                Debug.Print("error")
            End Try
        End If
        oPro = Nothing : oProS = Nothing
    End Sub
    ''
    ''Propiedades de Design Tracking / Propiedades de Design Tracking
    'Internal name: {32853F0F-3444-11D1-9E93-0060B03C1CA6}
    Public Sub PropiedadEscribeDesignTracking(ByRef queDoc As Inventor.Document, ByVal quePro As String, ByVal queVal As Object, Optional ByVal queFi As String = "", Optional ByVal cerrar As Boolean = False)
        If queVal = "" Or queVal Is Nothing Then Exit Sub

        Dim estabaabierto As Boolean = False
        If queDoc Is Nothing Then
            If FicheroAbierto(queFi) = False Then
                queDoc = oAppCls.Documents.Open(queFi, False)
                estabaabierto = False
            Else
                queDoc = oAppCls.Documents.ItemByName(queFi)
                estabaabierto = True
            End If
        End If
        '' Escribe un valor de texto en una iProperty. Si no existe la crea con valor "".
        Dim oProS As PropertySet = queDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}")
        Dim oPro As Inventor.Property = Nothing

        Try
            oPro = oProS.Item(quePro)
        Catch ex As Exception
            '' No hacemos nada. Aquí no se pueden crear propiedades nuevas.
        End Try
        '' Usaremos esto como expresión. Por si es Nothing oPro.Expression.
        Dim oProExp As String = IIf(oPro.Expression Is Nothing, "", oPro.Expression)

        '' Solo escribiremos el valor si: la propiedad no es una expresion o le mandamos una nueva expresion y si
        '' no es una expresión pero es diferente.
        If queVal.ToString.StartsWith("=") = True And queVal <> oProExp Then
            'oPro.Expression = queVal.ToString
            oPro.Expression = queVal.ToString
            GoTo FINAL
        ElseIf queVal.ToString.StartsWith("=") = True And queVal = oProExp Then
            GoTo FINAL
        ElseIf oProExp.StartsWith("=") Then     '' Si hay una expresión ya. No hacemos nada.
            GoTo FINAL
        ElseIf queVal <> oPro.Value Then
            oPro.Value = queVal
            GoTo FINAL
        ElseIf queVal = oPro.Value Then
            GoTo FINAL
        End If

FINAL:
        If estabaabierto = False And cerrar = True Then
            '' Si queremos que se actualice y se guarde. Lo quitamos para ganar tiempo. Lo haremos al final.
            Try
                If queDoc.RequiresUpdate Then queDoc.Update2()
            Catch ex As Exception
                '' Continuamos.
                Debug.Print("error")
            End Try
            Try

                If queDoc.Dirty = True Then queDoc.Save2(False)
            Catch ex As Exception
                '' No lo guardamos si da error.
                Debug.Print("error")
            End Try
            Try
                queDoc.Close(True)
            Catch ex As Exception
                '' No hacemos nada. Lo dejamos abierto.
                Debug.Print("error")
            End Try
        End If
        oPro = Nothing : oProS = Nothing
    End Sub

    Public Function PropiedadLeeUsuario(ByVal queDoc As Inventor.Document, ByVal quePro As String,
                                        Optional queF As String = "",
                                        Optional crear As Boolean = False,
                                        Optional valor As String = "",
                                        Optional ByRef quePss As PropertySets = Nothing) As String
        '' Por si le damos fullFilename (queF) en vez de Document
        Dim estabaabierto As Boolean = True
        If queDoc Is Nothing AndAlso queF <> "" AndAlso IO.File.Exists(queF) = True Then
            estabaabierto = FicheroAbierto(queF)
            oAppCls.SilentOperation = True
            If estabaabierto = True Then
                queDoc = oAppCls.Documents.ItemByName(queF)
            Else
                queDoc = oAppCls.Documents.Open(queF, False)
            End If
            oAppCls.SilentOperation = False
        End If
        '' ***********************************************
        '' Lee un valor de texto en una iProperty de usuario. Si no existe la crea con valor "".
        Dim resultado As String = ""
        If queDoc Is Nothing Then
            Return resultado
            Exit Function
        End If
        Dim oProS As PropertySet = Nothing
        If quePss IsNot Nothing Then
            oProS = quePss.Item("{D5CDD505-2E9C-101B-9397-08002B2CF9AE}")
        Else
            oProS = queDoc.PropertySets.Item("{D5CDD505-2E9C-101B-9397-08002B2CF9AE}")
        End If
        Dim oPro As Inventor.Property = Nothing
        Try
            oPro = oProS.Item(quePro)
            resultado = oPro.Value.ToString
        Catch ex As Exception
            If crear = True Then
                'Me.PropiedadEscribeUsuario(queDoc, quePro, "Centro", , False)
                Me.PropiedadEscribeUsuario(queDoc, quePro, valor, , False, True)
                oPro = oProS.Item(quePro)
                resultado = oPro.Value.ToString
            Else
                resultado = ""
            End If
        End Try
        'resultado = oPro.Value.ToString
        oProS = Nothing
        oPro = Nothing
        If estabaabierto = False Then queDoc.Close(True)
        Return resultado
    End Function

    Public Function PropiedadLeeResumenInventor(ByRef queDoc As Inventor.Document, ByVal quePro As String, Optional ByRef quePss As PropertySets = Nothing) As String
        'Información de resumen de Inventor / Información de resumen de Inventor
        'Internal name: {F29F85E0-4FF9-1068-AB91-08002B27B3D9}
        '
        ' Nombre: Title (Título) / Valor:  / Id: 2
        ' Nombre: Subject (Asunto) / Valor:  / Id: 3
        ' Nombre: Author (Autor) / Valor: Raul / Id: 4
        ' Nombre: Keywords (Palabras clave) / Valor:  / Id: 5
        ' Nombre: Comments (Comentarios) / Valor:  / Id: 6
        ' Nombre: Last Saved By (Guardado por última vez por) / Valor:  / Id: 8
        ' Nombre: Revision Number (Nº de revisión) / Valor:  / Id: 9
        ' Nombre: Thumbnail (Miniatura) / Valor:  / Id: 17

        '' Lee un valor de texto en una iProperty de Resumen Inventor. Si no existe devuelve "".
        Dim resultado As String = ""
        Dim oProS As PropertySet = Nothing
        Dim oPro As Inventor.Property = Nothing
        Try
            If quePss IsNot Nothing Then
                oProS = quePss.Item("{D5CDD505-2E9C-101B-9397-08002B2CF9AE}")
            Else
                oProS = queDoc.PropertySets.Item("{F29F85E0-4FF9-1068-AB91-08002B27B3D9}")
            End If
            oPro = oProS.Item(quePro)
            resultado = oPro.Value.ToString
        Catch ex As Exception
            ' No existe la Propiedad indicada en quePro
        End Try

        oProS = Nothing
        oPro = Nothing
        Return resultado
    End Function

    Public Function PropiedadLeeResumenDocumento(ByRef queDoc As Inventor.Document, ByVal quePro As String, Optional ByRef quePss As PropertySets = Nothing) As String
        'Información resumen documento Inventor / Información resumen documento Inventor
        'Internal name: {D5CDD502-2E9C-101B-9397-08002B2CF9AE}

        'Nombre: Category (Categoría) / Valor:  / Id: 2
        'Nombre: Manager (Responsable) / Valor:  / Id: 14
        'Nombre: Company (Empresa) / Valor:  / Id: 15
        '' Lee un valor de texto en una iProperty de Resumen Documento. Si no existe devuelve "".
        Dim resultado As String = ""
        Dim oProS As PropertySet = Nothing
        Dim oPro As Inventor.Property = Nothing
        Try
            If quePss IsNot Nothing Then
                oProS = quePss.Item("{D5CDD505-2E9C-101B-9397-08002B2CF9AE}")
            Else
                oProS = queDoc.PropertySets.Item("{D5CDD502-2E9C-101B-9397-08002B2CF9AE}")
            End If
            oPro = oProS.Item(quePro)
            resultado = oPro.Value.ToString
        Catch ex As Exception
            ' No existe la Propiedad indicada en quePro
        End Try

        oProS = Nothing
        oPro = Nothing
        Return resultado
    End Function


    Public Function PropiedadLeeDesignTracking(queDoc As Inventor.Document, ByVal quePro As String, Optional ByRef quePss As PropertySets = Nothing) As String
        '' Lee un valor de texto en una iProperty de DesignTracking. Si no existe devuelve "".
        Dim resultado As String = ""
        Dim oProS As PropertySet = Nothing
        Dim oPro As Inventor.Property = Nothing
        Try
            If quePss IsNot Nothing Then
                oProS = quePss.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}")
            Else
                oProS = queDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}")
            End If
            oPro = oProS.Item(quePro)
            resultado = oPro.Value.ToString
        Catch ex As Exception
            ' No existe la Propiedad indicada en quePro
        End Try

        oProS = Nothing
        oPro = Nothing
        Return resultado
    End Function


    Public Function PropiedadLeeUsuarioDoc(ByRef queDoc As Inventor.Document, ByVal quePro As String, Optional ByRef quePss As PropertySets = Nothing) As String
        '' Propiedades de Inventor definidas por el usuario / Propiedades de Inventor definidas por el usuario
        '' Internal name: {D5CDD505-2E9C-101B-9397-08002B2CF9AE}
        '' Lee un valor de texto en una iProperty de DesignTracking. Si no existe devuelve "".
        Dim resultado As String = ""
        Dim oProS As PropertySet = Nothing
        Dim oPro As Inventor.Property = Nothing
        Try
            If quePss IsNot Nothing Then
                oProS = quePss.Item("{D5CDD505-2E9C-101B-9397-08002B2CF9AE}")
            Else
                oProS = queDoc.PropertySets.Item("{D5CDD505-2E9C-101B-9397-08002B2CF9AE}")
            End If
            oPro = oProS.Item(quePro)
            resultado = oPro.Value.ToString
        Catch ex As Exception
            ' No existe la Propiedad indicada en quePro
        End Try

        oProS = Nothing
        oPro = Nothing
        Return resultado
    End Function

    Public Function PropiedadLeeUsuarioHashtable(ByVal queDoc As Inventor.Document, ByVal queFichero As String, Optional ByRef quePss As PropertySets = Nothing) As Hashtable
        '' Lee un valor de texto en una iProperty de usuario. Si no existe la crea con valor "".
        Dim resultado As Hashtable = Nothing
        Dim oProS As PropertySet
        Dim estaabierto As Boolean = False


        If queDoc IsNot Nothing Then
            estaabierto = True
        ElseIf (queDoc Is Nothing) AndAlso queFichero <> "" AndAlso IO.File.Exists(queFichero) Then
            estaabierto = FicheroAbierto(queFichero)
            If estaabierto = True Then
                queDoc = oAppCls.Documents.ItemByName(queFichero)
            Else
                '' El FullDocumentName sin componentes. Más rápido para abrir.
                If queFichero.ToLower.EndsWith(".iam") = True Then
                    queFichero = Me.DameCaminoSinComponentesInventor(queFichero)
                End If

                oAppCls.SilentOperation = True
                queDoc = oAppCls.Documents.Open(queFichero, False)
                oAppCls.SilentOperation = False
            End If
        End If

        oProS = queDoc.PropertySets.Item("{D5CDD505-2E9C-101B-9397-08002B2CF9AE}")

        Try
            For Each oPro As Inventor.Property In oProS
                If resultado.ContainsKey(oPro.Name) = False Then resultado.Add(oPro.Name, oPro.Value)
            Next
        Catch ex As Exception
            '' Error leyendo propiedades
        End Try
        '' Si no estaba abierto antes. Lo cerramos.
        If estaabierto = False Then queDoc.Close(True)

        oProS = Nothing
        Return resultado
    End Function

    Public Sub BrowseNode(ByVal oComp As ComponentOccurrence, ByVal expandir As Boolean, Optional ByVal padre As Boolean = False)
        Dim oNodePadre As BrowserNode = Nothing
        Dim oNode As BrowserNode = Nothing
        Dim oNNativo As NativeBrowserNodeDefinition = oAppCls.ActiveEditDocument.BrowserPanes.GetNativeBrowserNodeDefinition(oComp)
        Dim oNodeCarpeta As BrowserFolder = Nothing
        oNodePadre = oAppCls.ActiveEditDocument.BrowserPanes.ActivePane.TopNode
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
        If queApp.ActiveDocumentType = DocumentTypeEnum.kPartDocumentObject AndAlso queApp.ActiveEditDocument.FullFileName = queFile Then
            quePieza = queApp.ActiveEditDocument
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
        Dim queOpeAct As New Hashtable(StringComparer.OrdinalIgnoreCase)
        Dim quePieza As PartDocument = Nothing
        Dim cerrar As Boolean = False
        If queApp.ActiveDocumentType = DocumentTypeEnum.kPartDocumentObject Then
            quePieza = queApp.ActiveEditDocument
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
    ''
    Public Sub InsertarComponente(ByVal queFichero As String, ensPadre As AssemblyDocument,
                                  Optional x As Double = 0,
                                  Optional y As Double = 0)
        If Me.oAppCls.ActiveDocumentType <> Inventor.DocumentTypeEnum.kAssemblyDocumentObject Then
            MsgBox("Macro sólo para Ensamblajes...")
            Exit Sub
        End If
        ''Utilidades.GuardaMensaje("***** INSERTAR PIEZA NUEVA EN ENSAMBLAJE (INSERTAREF) *****")
        Dim tg As Inventor.TransientGeometry
        tg = Me.oAppCls.TransientGeometry

        Dim compD As AssemblyComponentDefinition
        compD = oAsm.ComponentDefinition

        Dim omatrix As Inventor.Matrix
        omatrix = tg.CreateMatrix

        Call omatrix.SetTranslation(tg.CreateVector(x, y, dblZ))
        dblZ += 10
        Try
            Me.oAppCls.SilentOperation = True
            modVariables.oOcuNew = oAsm.ComponentDefinition.Occurrences.Add(queFichero, omatrix)
            ensPadre.Rebuild2()
            ensPadre.Update2()
            'oAsm.Save()
            'ensPadre.ReleaseReference()
            'Dim od As Inventor.Document
            'od = oOcu.ReferencedDocumentDescriptor.ReferencedDocument
            'od.ReleaseReference()
            oOcuNew = Nothing
        Catch ex As Exception
            ''Utilidades.GuardaMensaje("***** ERROR INSERTANDO PIEZA NUEVA *****")
        Finally
            ''Utilidades.GuardaMensaje("***** PIEZA NUEVA INSERTADA CORRECTAMENTE (INSERTAREF) *****")
            Me.oAppCls.SilentOperation = False
        End Try
    End Sub
    ''
    Public Function InsertarComponenteAndGet(ByVal queFichero As String, ensPadre As AssemblyDocument,
                                  Optional x As Double = 0,
                                  Optional y As Double = 0) As ComponentOccurrence
        ''
        Dim resultado As ComponentOccurrence = Nothing
        ''
        If Me.oAppCls.ActiveDocumentType <> Inventor.DocumentTypeEnum.kAssemblyDocumentObject Then
            MsgBox("Macro only for Assemblies...")
            Return resultado
            Exit Function
        End If
        ''Utilidades.GuardaMensaje("***** INSERTAR PIEZA NUEVA EN ENSAMBLAJE (INSERTAREF) *****")
        Dim tg As Inventor.TransientGeometry
        tg = Me.oAppCls.TransientGeometry

        Dim compD As AssemblyComponentDefinition
        compD = oAsm.ComponentDefinition

        Dim omatrix As Inventor.Matrix
        omatrix = tg.CreateMatrix

        Call omatrix.SetTranslation(tg.CreateVector(x, y, dblZ))
        dblZ += 10

        Try
            Me.oAppCls.SilentOperation = True
            resultado = oAsm.ComponentDefinition.Occurrences.Add(queFichero, omatrix)
            ensPadre.Rebuild2()
            ensPadre.Update2()
        Catch ex As Exception
            ''Utilidades.GuardaMensaje("***** ERROR INSERTANDO PIEZA NUEVA *****")
        Finally
            ''Utilidades.GuardaMensaje("***** PIEZA NUEVA INSERTADA CORRECTAMENTE (INSERTAREF) *****")
            Me.oAppCls.SilentOperation = False
        End Try
        ''
        Return resultado
    End Function
    ''
    ''
    Public Sub InsertarComponentePonPropiedad(ByVal queFichero As String, ensPadre As AssemblyDocument, quePro As String, queVal As String)
        If Me.oAppCls.ActiveDocumentType <> Inventor.DocumentTypeEnum.kAssemblyDocumentObject Then
            MsgBox("Macro sólo para Ensamblajes...")
            Exit Sub
        End If
        ''Utilidades.GuardaMensaje("***** INSERTAR PIEZA NUEVA EN ENSAMBLAJE (INSERTAREF) *****")
        Dim tg As Inventor.TransientGeometry
        tg = Me.oAppCls.TransientGeometry

        Dim compD As AssemblyComponentDefinition
        compD = oAsm.ComponentDefinition

        Dim omatrix As Inventor.Matrix
        omatrix = tg.CreateMatrix

        Call omatrix.SetTranslation(tg.CreateVector(0, 0, dblZ))
        dblZ += 30
        Try
            Me.oAppCls.SilentOperation = True
            Dim oOcu As ComponentOccurrence = oAsm.ComponentDefinition.Occurrences.Add(queFichero, omatrix)
            '' Poner propiedad
            Me.PropiedadEscribeUsuario(oOcu.Definition.Document, quePro, queVal, , , True)
            ''
            ensPadre.Rebuild2()
            ensPadre.Update2()
            'oAsm.Save()
            'ensPadre.ReleaseReference()
            'Dim od As Inventor.Document
            'od = oOcu.ReferencedDocumentDescriptor.ReferencedDocument
            'od.ReleaseReference()
            oOcu = Nothing
        Catch ex As Exception
            ''Utilidades.GuardaMensaje("***** ERROR INSERTANDO PIEZA NUEVA *****")
        Finally
            ''Utilidades.GuardaMensaje("***** PIEZA NUEVA INSERTADA CORRECTAMENTE (INSERTAREF) *****")
            Me.oAppCls.SilentOperation = False
        End Try
    End Sub
    ''
    Public Function CambiaElementoEnsamblaje(ByVal oE As AssemblyDocument, ByVal nOcu1 As String, ByVal fullOcu1 As String, ByVal bolTodas As Boolean, Optional borrarAntes As Boolean = False) As PartDocument
        Dim resultado As PartDocument = Nothing
        Dim oC As ComponentOccurrence = oE.ComponentDefinition.Occurrences.ItemByName(nOcu1)
        Dim dirEns As String = modUtilidades.DameParteCamino(oE.FullFileName, ParteCamino.CaminoSinFicheroBarra)
        '' dirPilares  ' Directorio donde están las plantillas de pilares.
        Dim caminoAntes As String = oC.ReferencedDocumentDescriptor.ReferencedFileDescriptor.FullFileName
        '' Si es el mismo no hacemos nada y salimos.
        If caminoAntes = fullOcu1 Then
            resultado = oC.Definition.Document
            CambiaElementoEnsamblaje = resultado
            Exit Function
        End If
        Dim dirPilAhora As String = modUtilidades.DameParteCamino(oC.ReferencedDocumentDescriptor.ReferencedFileDescriptor.FullFileName, ParteCamino.CaminoSinFicheroBarra)
        Dim nomPilAhora As String = modUtilidades.DameParteCamino(oC.ReferencedDocumentDescriptor.ReferencedFileDescriptor.FullFileName, ParteCamino.SoloFicheroConExtension)
        Dim dirPilBibli As String = modUtilidades.DameParteCamino(fullOcu1, ParteCamino.CaminoSinFicheroBarra)
        Dim nomPilBibli As String = modUtilidades.DameParteCamino(fullOcu1, ParteCamino.SoloFicheroConExtension)


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

    Public Function ComponenteDame(ByVal queEns As AssemblyDocument, ByVal queCaminoCompleto As String) As ComponentOccurrence
        Dim resultado As ComponentOccurrence = Nothing
        For Each resultado In queEns.ComponentDefinition.Occurrences
            If resultado.ReferencedDocumentDescriptor.ReferencedFileDescriptor.FullFileName = queCaminoCompleto Then Exit For
        Next
        ComponenteDame = resultado
        Exit Function
    End Function
    ''
    '' Devolver caminos únicos para arrRefOrdenMod, arrRefOrdenEst y arrRefOrdenPie
    '' SortedList de indice en base 0
    Public Function ArrayList_arrRefOrdenX_DameUnicos(queArr As ArrayList) As SortedList
        Dim resultado As New SortedList
        Dim contador As Integer = 0
        For x As Integer = 0 To queArr.Count - 1
            Dim partes() As String = queArr(x).ToString.Split("·"c)
            Dim codigo As String = partes(0)
            Dim camino As String = partes(1)
            If resultado.ContainsValue(camino) = False Then
                resultado.Add(contador, camino)
                contador += 1
            End If
        Next
        Return resultado
    End Function
    ''
#Region "FileReferences"
#Region "FileReferencesDameTodasInventor"
    ''
    Private Sub FileReferencesDameTodasBOMSinNombreRecursivo(queRow As BOMRow, ByRef arrCam As ArrayList)
        ''
        For Each oRow As BOMRow In queRow.ChildRows
            'Set a reference to the primary ComponentDefinition of the row
            Dim oCompDef As Inventor.ComponentDefinition
            oCompDef = oRow.ComponentDefinitions.Item(1)
            '' Si es componente virtual, continuamos.
            If TypeOf oCompDef Is VirtualComponentDefinition Then Continue For
            ''
            Dim oD As Inventor.Document = oCompDef.Document
            Dim caminoFull As String = oD.FullFileName
            If arrCam.Contains(caminoFull) = False Then arrCam.Add(caminoFull)

            If caminoFull.ToLower.EndsWith(".iam") Or caminoFull.ToLower.EndsWith(".ipt") Then
                '' **** Buscamos también si tiene plano IDW o DWG para añadirlo a la colección.
                Dim planos As ArrayList = Nothing
                planos = clsI.ExisteFicheroPlanoEnDirBasico(caminoFull)
                ''
                If planos IsNot Nothing AndAlso planos.Count > 0 Then
                    For Each queF As String In planos
                        If arrCam.Contains(queF) = False Then arrCam.Add(queF)
                    Next
                End If
                planos = Nothing
            End If
            ''
            '' Recursivamente, los hijos que tenga.
            FileReferencesDameTodasBOMSinNombreRecursivo(oRow, arrCam)
        Next
    End Sub

    ''' <summary>
    ''' All Childs files in parentFile. Of Inventor or not.
    ''' </summary>
    ''' <param name="parentFile">FullFilename partnFile</param>
    ''' <param name="bolVisible">True/False for visible</param>
    ''' <param name="addDrawing">True : Add IDW and DWG</param>
    ''' <param name="BasicSearch">True : Only same name</param>
    ''' <returns>Return ArrayList with FullFilenames Child files</returns>
    ''' <remarks>Optimize if parentFile is IDW or DWG of Inventor</remarks>
    Public Function FileReferencesReadAllInventor(ByVal parentFile As String,
                                                    ByVal bolVisible As Boolean,
                                                    ByVal addDrawing As Boolean,
                                                    ByVal BasicSearch As Boolean,
                                                    Optional addParent As Boolean = True,
                                                    Optional addCC As Boolean = False) As ArrayList
        On Error Resume Next
        Dim result As New ArrayList
        Dim isOpen As Boolean '= True
        Dim oDoc As Inventor.Document = Nothing '= oD
        If parentFile = "" Then
            MsgBox("Error : parentFile is nothing")
            Return result
            Exit Function
        End If

        If parentFile <> "" And IO.File.Exists(parentFile) Then
            oAppCls.SilentOperation = True
            isOpen = Me.FicheroAbierto(parentFile)
            If isOpen = False Then
                oDoc = oAppCls.Documents.Open(parentFile, bolVisible)
            Else
                oDoc = oAppCls.Documents.ItemByName(parentFile)
            End If
            'Try
            oDoc.Update2(True)
            If oDoc.Dirty = True Then oDoc.Save2(True)
            '' Activate representation "Principal"
            If oDoc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                '' All active and visible
                RepresentacionActivaCrea(CType(oDoc, AssemblyDocument), False, "")
            End If
            oAppCls.SilentOperation = False
        End If

        If oDoc Is Nothing Then
            result = Nothing
            Return result
            Exit Function
        End If

        Dim pathFull As String = oDoc.FullFileName
        If addParent = True AndAlso result.Contains(pathFull) = False Then result.Add(pathFull)

        If pathFull.ToLower.EndsWith(".iam") Or pathFull.ToLower.EndsWith(".ipt") Or pathFull.ToLower.EndsWith(".ipn") _
        And addDrawing = True Then
            '' Add Drawing (IDW or DWG Inventor)
            Dim planos As ArrayList = Nothing
            If BasicSearch = True Then
                planos = clsI.ExisteFicheroPlanoEnDirBasico(pathFull)
            Else
                planos = clsI.ExisteFicheroPlanoEnDirProfundoInv(pathFull)
            End If

            If planos IsNot Nothing AndAlso planos.Count > 0 Then
                For Each queF As String In planos
                    If result.Contains(queF) = False Then result.Add(queF)
                Next
            End If
            planos = Nothing
            '' ***************************************************************************
        Else
            '' Nothing
        End If

        Call FileReferencesReadAllInventorRecursive(oDoc.File, result, addDrawing, BasicSearch, addCC)

        If isOpen = False Then oDoc.Close(True)
        oDoc = Nothing

        GC.WaitForPendingFinalizers()
        GC.Collect()
        '' addDrawing in BasicSearch
        If result IsNot Nothing And addDrawing = True Then
            For Each queF As String In result
                On Error Resume Next
                If queF.ToLower.EndsWith(".iam") Or queF.ToLower.EndsWith(".ipt") Or
                     queF.ToLower.EndsWith(".ipn") Then
                    Dim planoIDW As String = DameParteCamino(queF, ParteCamino.SoloCambiaExtension, ".idw")
                    Dim planoDWG As String = DameParteCamino(queF, ParteCamino.SoloCambiaExtension, ".dwg")
                    If result.Contains(planoIDW) Or result.Contains(planoDWG) Then Continue For
                    If IO.File.Exists(planoIDW) Then result.Add(planoIDW)
                    If IO.File.Exists(planoDWG) AndAlso oAppCls.FileManager.IsInventorDWG(planoDWG) = True Then _
                    result.Add(planoDWG)
                End If
            Next
        End If

        GC.WaitForPendingFinalizers()
        GC.Collect()
        Return result
    End Function

    Private Sub FileReferencesReadAllInventorRecursive(ByVal oFile As Inventor.File,
                                                         ByRef result As ArrayList,
                                                         ByVal addDrawing As Boolean,
                                                         ByVal BasicSearch As Boolean,
                                                         Optional addCC As Boolean = False)
        Dim oFileDescriptor As FileDescriptor
        For Each oFileDescriptor In oFile.ReferencedFileDescriptors
            If addCC = False And oFileDescriptor.LocationType = LocationTypeEnum.kLibraryLocation Then
                Continue For
            End If
            '' if not ReferenceMising
            If Not oFileDescriptor.ReferenceMissing Then
                Dim FullPath As String = oFileDescriptor.FullFileName
                '' Add to ArrayList
                If Not result.Contains(FullPath) Then result.Add(FullPath)

                '' Add Drawing. Only if iam or ipt.
                If FullPath.ToLower.EndsWith(".iam") Or FullPath.ToLower.EndsWith(".ipt") _
                Or FullPath.ToLower.EndsWith(".ipn") And addDrawing = True Then
                    Dim planos As ArrayList = Nothing
                    If BasicSearch = True Then
                        planos = clsI.ExisteFicheroPlanoEnDirBasico(FullPath)
                    Else
                        planos = clsI.ExisteFicheroPlanoEnDirProfundoInv(FullPath)
                    End If

                    If planos IsNot Nothing AndAlso planos.Count > 0 Then
                        For Each queF As String In planos
                            If result.Contains(queF) = False Then result.Add(queF)
                        Next
                    End If
                    planos = Nothing
                    '' ***************************************************************************
                End If

                ' Only Inventor files recursive
                If Not oFileDescriptor.ReferencedFileType = FileTypeEnum.kForeignFileType Then
                    Call FileReferencesReadAllInventorRecursive(oFileDescriptor.ReferencedFile, result, addDrawing, BasicSearch)
                End If
            End If
        Next
    End Sub
    ''
    '' Change all old references files for new references files in parentFile.
    '' FullPath Parent File (parentFile)
    '' Hastable (colOldNew) is Key=Old FullPath, Value=New FullPath
    '' Optional resursive for recursive Childs.
    Public Sub FileReferenciaChangeAllInventor(ByVal parentFile As String, ByVal colOldNew As Hashtable, Optional ByVal recursive As Boolean = True)
        Dim oDoc1 As Inventor.Document = Nothing
        Dim oEns1 As Inventor.AssemblyDocument = Nothing
        Dim oPie1 As Inventor.PartDocument = Nothing
        Dim oIdw1 As Inventor.DrawingDocument = Nothing
        Dim isopen As Boolean = False
        ''
        oAppCls.SilentOperation = True
        If Me.FicheroAbierto(parentFile) = True Then
            oDoc1 = oAppCls.Documents.ItemByName(parentFile)
            isopen = True
        Else
            '' Open with options
            Dim oNameVP As NameValueMap = Me.oAppCls.TransientObjects.CreateNameValueMap
            ''
            If parentFile.ToLower.EndsWith(".iam") Or parentFile.ToLower.EndsWith(".ipt") Then
                oNameVP.Add("SkipAllUnresolvedFiles", True)
            ElseIf parentFile.ToLower.EndsWith(".idw") Or parentFile.ToLower.EndsWith(".dwg") Then
                oNameVP.Add("DeferUpdates", True)
                oNameVP.Add("SkipAllUnresolvedFiles", True)
            End If
            oDoc1 = Me.oAppCls.Documents.OpenWithOptions(parentFile, oNameVP, False)
            isopen = False
        End If

        If oDoc1.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
            oEns1 = oAppCls.Documents.ItemByName(parentFile)
            Me.RepresentacionActivaCrea(oEns1, False)    ' Activate "Principal"
            'oEns.Update2() : oEns.Save2()
            Call FileReferenciaChangeAllInventorRecursive(oEns1.File, colOldNew, recursive)
            'oEns1.Rebuild2(True)
            If oEns1.RequiresUpdate Then oEns1.Update2(True)
            oEns1.Save2(True)
        ElseIf oDoc1.DocumentType = DocumentTypeEnum.kDrawingDocumentObject Then
            oIdw1 = oAppCls.Documents.ItemByName(parentFile) 'oDoc
            If oIdw1.FullFileName.EndsWith(".dwg") AndAlso oIdw1.IsInventorDWG = False Then Exit Sub
            Call FileReferenciaChangeAllInventorRecursive(oIdw1.File, colOldNew, recursive)
            oIdw1.DrawingSettings.DeferUpdates = False
            oIdw1.Update2(True) : oIdw1.Save2(True)
        ElseIf oDoc1.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
            oPie1 = oAppCls.Documents.ItemByName(parentFile) 'oDoc
            Call FileReferenciaChangeAllInventorRecursive(oPie1.File, colOldNew, recursive)
            'oPie1.Rebuild2(True)
            If oPie1.RequiresUpdate Then oPie1.Update2(True)
            oPie1.Save2(True)
        End If
        ''
        Try
            If oEns1 IsNot Nothing Then
                oEns1.Close(True)
            ElseIf oIdw1 IsNot Nothing Then
                oIdw1.Close(True)
            ElseIf oPie1 IsNot Nothing Then
                oPie1.Close(True)
            End If
        Catch ex As Exception
            'oDoc.Save2(False)
            'Console.WriteLine(ex.Message)
        End Try

        oAppCls.SilentOperation = False
        oEns1 = Nothing
        oIdw1 = Nothing
        oPie1 = Nothing
        oDoc1 = Nothing
    End Sub
    ''
    Private Sub FileReferenciaChangeAllInventorRecursive(ByVal oFile As Inventor.File, ByVal colOldNew As Hashtable, Optional ByVal recursive As Boolean = True)
        For Each oFD As FileDescriptor In oFile.ReferencedFileDescriptors
            '' Continue if
            If oFD.FullFileName.Contains("newVer.") Then Continue For
            '' In colOldNew.Values
            If colOldNew.ContainsValue(oFD.FullFileName) Then Continue For
            '' Not in colOldNew.Keys
            If colOldNew.ContainsKey(oFD.FullFileName) = False Then Continue For
            '' Not file in final folder
            If IO.File.Exists(colOldNew(oFD.FullFileName)) = False Then Continue For
            ''
            If oFD.LocationType = LocationTypeEnum.kLibraryLocation Then Continue For
            '' Cambiamos la referencia si la plantilla inicial esta y si existe el fichero final.
            If colOldNew.ContainsKey(oFD.FullFileName) AndAlso IO.File.Exists(colOldNew(oFD.FullFileName)) Then
                oFD.ReplaceReference(colOldNew(oFD.FullFileName))
            End If

            Me.DoEventsInventor(True)
            ' Recursivamente salvo que sea un kForeignFileType (Otros documentos NO de Inventor)
            If Not oFD.ReferencedFileType = FileTypeEnum.kForeignFileType Then
                Call FileReferenciaChangeAllInventorRecursive(oFD.ReferencedFile, colOldNew, recursive)
            End If
            'End If
        Next
    End Sub
#End Region
#End Region


    '' Devuelve un NameValueMap con todos los componentes desactivados.
    '' ** Solo para leer propiedades
    ''' <summary>
    ''' Para usar con OpenWithOptions(fullname, NameValueMap, Visible)
    ''' </summary>
    ''' <returns>Devuelve un NameValueMap con todos los componentes desactivados.</returns>
    ''' <remarks></remarks>
    Public Function OpenLight() As NameValueMap
        Dim resultado As NameValueMap = oAppCls.TransientObjects.CreateNameValueMap
        '' Principal (1) 56065
        '' Todos los componentes desactivados (2) 56066
        '' Todas las piezas desactivadas (3) 56067
        '' Todo el Centro de contenido desactivado (4) 56068
        ' Set the representations to use when opening the document.
        Call resultado.Add("LevelOfDetailRepresentation", "Todos los componentes desactivados")
        'Call resultado.Add("PositionalRepresentation", "MyPositionalRep")
        'Call resultado.Add("DesignViewRepresentation", "MyDesignViewRep")

        Return resultado
    End Function

    Public Sub GuardaTodo(cerrarTodo As Boolean, Optional guardar As Boolean = True)
        oAppCls.SilentOperation = True
        Try
            For Each oDoc As Inventor.Document In oAppCls.Documents
                If guardar = True Then
                    If oDoc.Dirty = True Then oDoc.Save2(False)
                End If
                If cerrarTodo = True Then
                    oDoc.Close(True)
                End If
                Next
            If cerrarTodo = True Then oAppCls.Documents.CloseAll()
        Catch ex As Exception
            ''
        End Try
        oAppCls.SilentOperation = False
    End Sub
    Public Sub AbreActualizaGuarda(ByVal queFichero As String, Optional ByVal elIDW As Boolean = False, Optional ByVal elDWG As Boolean = False)
        'Dim estaVisible As Boolean = False
        If queFichero.Contains("OldVersions") Then Exit Sub
        Dim procesado As Integer = 0
TambienDibujo:
        Dim oDoc As Inventor.Document = Nothing

        Try
            oAppCls.SilentOperation = True
            For Each oDoch In oAppCls.Documents
                If oDoch.FullFileName = queFichero Then
                    oDoc = oDoch
                    Exit For
                End If
            Next

            oAppCls.SilentOperation = True
            If Dir(queFichero) = "" Then Exit Sub
            If (procesado = 2) AndAlso oAppCls.FileManager.IsInventorDWG(queFichero) = False Then Exit Sub
            If oDoc Is Nothing Then
                oDoc = oAppCls.Documents.Open(queFichero, False)
            End If
            '' Si es un dibujo DWG y NO es de Inventor, salimos.
            If oDoc.DocumentType = DocumentTypeEnum.kDrawingDocumentObject AndAlso CType(oDoc, DrawingDocument).IsInventorDWG = False Then Exit Sub

            If procesado = 0 Then   ' Si es el ensamblaje. Ponemos Representación nivelDetalleDefecto como activa.
                Me.RepresentacionActivaCrea(CType(oDoc, AssemblyDocument), True, nivelDetalleDefecto)
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
                    planos = clsI.ExisteFicheroPlanoEnDirBasico(queFichero)
                Else
                    planos = clsI.ExisteFicheroPlanoEnDirProfundoInv(queFichero)
                End If
                Dim planoIDW As String = ""
                If planos IsNot Nothing AndAlso planos.Count > 0 Then
                    planoIDW = clsI.ExisteFicheroPlanoEnArray(planos, IO.Path.ChangeExtension(queFichero, ".idw"))
                End If
                '' *************************************************
                procesado = 1
                GoTo TambienDibujo
                '' *****************************************************************************
            ElseIf elDWG = True And procesado = 1 Then
                '' ***** Para sacar el plano DWG que tenga
                Dim planos As ArrayList
                If Busquedabasica = True Then
                    planos = clsI.ExisteFicheroPlanoEnDirBasico(queFichero)
                Else
                    planos = clsI.ExisteFicheroPlanoEnDirProfundoInv(queFichero)
                End If
                Dim planoDWG As String = ""
                If planos IsNot Nothing AndAlso planos.Count > 0 Then
                    planoDWG = clsI.ExisteFicheroPlanoEnArray(planos, IO.Path.ChangeExtension(queFichero, ".dwg"))
                End If
                '' *************************************************
            End If

            Me.oAppCls.SilentOperation = False
        Catch ex As Exception
            Debug.Print("ALBERTO-->Error en clsInventor-->AbreActualizaGuarda" & vbCrLf & ex.Message)
            oAppCls.SilentOperation = False
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
                    oAppCls.SilentOperation = True
                    Try
                        Dim oD As Inventor.Document = oV.ReferencedDocumentDescriptor.ReferencedDocument
                        If oD IsNot Nothing AndAlso procesados.Contains(oD.FullFileName) = False Then
                            oD.Update2() : oD.Save2()
                            procesados.Add(oD.FullFileName)
                        End If
                    Catch ex As Exception
                        Continue For
                    End Try
                    Me.oAppCls.SilentOperation = False
                Next oV
            Next oS
            oDoc.Save2()
        Catch ex As Exception
            'MsgBox("ALBERTO-->Error en clsInventor-->ActualizaGuardaDibujo" & vbCrLf & ex.Message)
            Debug.Print("ALBERTO-->Error en clsInventor-->AbreActualizaGuarda" & vbCrLf & ex.Message)
        End Try
    End Sub

    Public Function DameComponentesTreeNode(ByVal oD As Inventor.AssemblyDocument, _
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
            If solopadres = True AndAlso clsI.PropiedadLeeUsuario(oE, "_TIPO", , True, "").ToUpper <> "PADRE" Then _
            Continue For

            '' Si esH = True. Buscaremos sólo los elementos Horizontales
            '' Si esH = False. Buscaremos sólo los elementos Verticales (PI en Category)
            '' Category = PI o Category = FAMILIA·Tipo (FUTURA·fu90)
            Dim categoria As String = clsI.PropiedadLeeCategoria(oE)
            'Dim categoria As String = clsI.PropiedadLeeCategoriaApprentice(oE.FullFileName)

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

    '' Devuelve un Arraylist de Componentoccurrences
    Public Function DameComponentesTODOS(ByVal oEns As Inventor.AssemblyDocument, ByVal primernivel As Boolean, Optional ByVal nivelPrincipal As Boolean = True) As ArrayList
        Dim resultado As New ArrayList
        '' Para guardar representacion nivel detalle activa.
        Dim repActiva As LevelOfDetailRepresentation = Nothing
        repActiva = oEns.ComponentDefinition.RepresentationsManager.ActiveLevelOfDetailRepresentation

        If nivelPrincipal = True Then clsI.RepresentacionActivaCrea(oEns, False)
        'MsgBox ("Total occurrences = ( " & oEns.ComponentDefinition.Occurrences.AllLeafOccurrences.Count & " )")
        'MsgBox ("Total occurrences = ( " & oEns.ComponentDefinition.Occurrences.AllReferencedOccurrences(oEns.ComponentDefinition).Count & " )")

        Dim oOcu As ComponentOccurrence
        If primernivel = True Then
            For Each oOcu In oEns.ComponentDefinition.Occurrences
                On Error Resume Next
                Dim oD As Inventor.Document = oOcu.ReferencedDocumentDescriptor.ReferencedDocument
                If (oD Is Nothing) Then Continue For
                If oD.FullFileName = "" Then Continue For
                resultado.Add(oOcu)
            Next
        Else
            For Each oOcu In oEns.ComponentDefinition.Occurrences.AllReferencedOccurrences(oEns.ComponentDefinition)
                On Error Resume Next
                Dim oD As Inventor.Document = oOcu.ReferencedDocumentDescriptor.ReferencedDocument
                If (oD Is Nothing) Then Continue For
                If oD.FullFileName = "" Then Continue For
                resultado.Add(oOcu)
            Next
        End If

        clsI.RepresentacionActivaCrea(oEns, True, repActiva.Name)
        DameComponentesTODOS = resultado
        Exit Function
    End Function

    Public Function DameComponentesTODOSrecursivo(ByVal oEns As Inventor.AssemblyDocument, ByVal primernivel As Boolean, _
                                                  Optional soloPrimerComponente As Boolean = False, _
                                                  Optional soloPiezas As Boolean = False) As ArrayList
        Dim resultado As New ArrayList
        Dim oOcu As ComponentOccurrence

        If soloPrimerComponente = True Then
            oOcu = oEns.ComponentDefinition.Occurrences(1)
            If resultado.Contains(oOcu) = False Then resultado.Add(oOcu)
            GoTo FINAL
        Else
            For Each oOcu In oEns.ComponentDefinition.Occurrences
                ' Check if it's child occurrence (leaf node)
                If soloPiezas = True AndAlso oOcu.DefinitionDocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                    ' No lo añadimos si es un ensamblaje.
                    'If resultado.Contains(oOcu) = False Then resultado.Add(oOcu)
                Else
                    If resultado.Contains(oOcu) = False Then resultado.Add(oOcu)
                End If
                Try
                    If primernivel = False And oOcu.SubOccurrences.Count > 0 Then
                        Call DameComponentesTODOSrecursivoSub(oOcu, resultado, soloPiezas)
                    End If
                Catch ex As Exception
                    Continue For
                End Try
            Next
        End If
FINAL:
        oOcu = Nothing
        Return resultado
    End Function

    ' This function is called for processing sub assembly.  It is called recursively
    ' to iterate through the entire assembly tree.
    Private Sub DameComponentesTODOSrecursivoSub(ByVal oCompOcc As ComponentOccurrence, ByRef queArray As ArrayList, Optional soloPiezas As Boolean = False)
        Dim oSubCompOcc As ComponentOccurrence
        'Try
        For Each oSubCompOcc In oCompOcc.SubOccurrences
            If soloPiezas = True AndAlso oSubCompOcc.DefinitionDocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                ' No lo añadimos si es un ensamblaje.
                'If queArray.Contains(oSubCompOcc) = False Then queArray.Add(oSubCompOcc)
            Else
                If queArray.Contains(oSubCompOcc) = False Then queArray.Add(oSubCompOcc)
            End If
            ' Check if it's child occurrence (leaf node)
            Try
                If oSubCompOcc.SubOccurrences.Count > 0 Then
                    Call DameComponentesTODOSrecursivoSub(oSubCompOcc, queArray)
                End If
            Catch ex As Exception
                Continue For
            End Try
        Next
    End Sub


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
        Dim tieneTriada As Boolean = oAppCls.GeneralOptions.Show3DIndicator
        oAppCls.GeneralOptions.Show3DIndicator = False
        Dim fd As New System.Windows.Forms.SaveFileDialog
        fd.AddExtension = True
        fd.DefaultExt = "png"
        fd.Filter = "Imagen PNG|*.png|Imagen BMP|*.bmp|Imagen JPG|*.jpg"
        fd.FilterIndex = 1
        fd.InitialDirectory = My.Application.Info.DirectoryPath
        fd.FileName = "PantallaInventor"
        If fd.ShowDialog = System.Windows.Forms.DialogResult.OK Then
            Me.oAppCls.ActiveView.SaveAsBitmap(fd.FileName, 1024, 768)
            Threading.Thread.Sleep(2000)
            Call Process.Start(fd.FileName)
        End If
        oAppCls.GeneralOptions.Show3DIndicator = tieneTriada
    End Sub


    Public Sub CopiaConApprentice(ByVal ficheroOrigen As String, ByVal ficheroDestino As String, _
                                  Optional ByVal sobreescribir As Boolean = False, Optional ByVal cerrar As Boolean = True)
        If Dir(ficheroOrigen) <> "" Then
            ' Create a new instance of Apprentice.

            Dim oApprentice As New ApprenticeServerComponent
            ' Open a document.
            Dim oDoc As ApprenticeServerDocument

            oDoc = oApprentice.Open(ficheroOrigen)
            Try
                If My.Computer.FileSystem.FileExists(ficheroDestino) = True And _
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

    Public Function DameComponentesArrTreeNodes(ByVal oD As Inventor.AssemblyDocument, _
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
            Dim CaminoDir As String = DameParteCamino(CaminoTodo, ParteCamino.CaminoConFicheroSinExtensionBarra)
            Dim NombreSolo As String = DameParteCamino(CaminoTodo, ParteCamino.SoloFicheroSinExtension)
            If oCo.ReferencedDocumentDescriptor.ReferencedDocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                oDocCom = oCo.ReferencedDocumentDescriptor.ReferencedDocument
            Else
                Continue For
            End If
            '' Si la propiedad _TIPO no es "PADRE" este no es un ensamblaje padre y
            '' pasamos al siguiente componente. Si solopadres=true
            If solopadres = True AndAlso clsI.PropiedadLeeUsuario(oDocCom, "_TIPO", , True, "").ToUpper <> "PADRE" Then _
            Continue For

            '' Si esH = True. Buscaremos sólo los elementos Horizontales
            '' Si esH = False. Buscaremos sólo los elementos Verticales (PI en Category)
            '' Category = PI o Category = (FU90)    Con FU leemos ADAPRETERSA.ini [CAMINOS]-->FU=FUTURA 'FAMILIA·Tipo (FUTURA·fu90)
            Dim categoria As String = clsI.PropiedadLeeCategoria(oDocCom)
            'Dim categoria As String = clsI.PropiedadLeeCategoriaApprentice(oDocCom.FullFileName)

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
                If Dir(CaminoDir) <> "" And _
                    Dir(CaminoDir & NombreSolo & "·armado.iam") <> "" And _
                    Dir(CaminoDir & NombreSolo & "·armado.idw") <> "" And _
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

    Public Sub ZoomTodoAjustar3D(Optional ByVal ponerIso As Boolean = True)
        Dim oVie As Inventor.View
        Dim oCam As Inventor.Camera
        oVie = Me.oAppCls.ActiveView  'oPie.Views.Item(1)   ' oApp.ActiveView
        oCam = oVie.Camera
        If ponerIso Then oCam.ViewOrientationType = ViewOrientationTypeEnum.kIsoTopRightViewOrientation
        oCam.ApplyWithoutTransition()
        oCam.Fit()
        oCam.ApplyWithoutTransition()
    End Sub

    Public Sub ZoomTodoAjustar2D(ByVal oSh As Sheet)
        Dim oVie As Inventor.View
        Dim oCam As Inventor.Camera
        oVie = Me.oAppCls.ActiveView  'oPie.Views.Item(1)   ' oApp.ActiveView
        oCam = oVie.Camera
        oCam.SetExtents(oSh.Width + 1, oSh.Height + 1)
        'oCam.Fit()
        oCam.ApplyWithoutTransition()
    End Sub


    Public Sub ZoomObjeto(ByVal queEns As AssemblyDocument, ByVal oC As ComponentOccurrence, ByVal pt3D As Point, ByVal pt2D As Point2d)
        Dim ancho, alto, ancho2d, alto2d As Double
        Dim pt1, pt2, ptCentro3D As Inventor.Point
        Dim pt1_2d, pt2_2d, ptC2d_Destino As Point2d
        Dim oVie As Inventor.View = oAppCls.ActiveView
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
        Dim RepPrincipal As String = oAppCls.FileManager.GetLevelOfDetailRepresentations(queFichero)(queRep)
        'RepPrincipal.Name
        Return oAppCls.FileManager.GetFullDocumentName(queFichero, RepPrincipal)
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
        straLOD = oAppCls.FileManager.GetLevelOfDetailRepresentations(queFichero)

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
            resultado = oAppCls.FileManager.GetFullDocumentName(queFichero, straLOD(2))  'LevelOfDetailEnum.kAllComponentsSuppressedLevelOfDetail.ToString)    ' straLOD(2))
        End If


        System.GC.WaitForPendingFinalizers()
        System.GC.Collect()
        Return resultado
    End Function

    Public Sub ExportarDibujosFormatosSaveAs(ByVal oDoc As Inventor.DrawingDocument, _
                                             ByVal queDestino As String, _
                                             ByVal queTipos As ArrayList)    ', ByVal abrirlo As Boolean)

        ' Todos los tipos de una Enum como array de cadenas: [Enum].GetNames(GetType(ENUM))
        ' Un tipo concreto de una Enum como cadena: [Enum].GetName(GetType(GuardaTipo), ultimoGT)
        If queDestino.EndsWith("\") = False Then queDestino &= "\"
        '' Nombre final del fichero (SIN EXTENSION)
        Dim ficheroFin As String = queDestino & DameParteCamino(oDoc.FullFileName, ParteCamino.SoloFicheroSinExtension)

        'For Each oSh As Sheet In oDoc.Sheets
        'clsI.ZoomTodoAjustar2D(oSh)
        'Next

        Dim ultimoGT As GuardaTipo
        'Publish document.
        For Each queG As GuardaTipo In queTipos
            ultimoGT = queG
            '' SaveAs no permite guardar formatos DWG ni DXF. Utilizar TranslatorAddIn.
            If queG = GuardaTipo.dwg Or queG = GuardaTipo.dxf Or queG = GuardaTipo.pdf Or queG = GuardaTipo.dwf Or queG = GuardaTipo.dwfx Then Continue For
            Dim fFin As String = ficheroFin & "." & [Enum].GetName(GetType(GuardaTipo), ultimoGT)
            Try
                If IO.File.Exists(fFin) Then IO.File.Delete(fFin)
                oDoc.SaveAs(fFin, True)
                Console.WriteLine("Guardado correctamente : " & fFin)
                'If abrirlo = True Then Call Process.Start(fFin)
            Catch ex As Exception
                MsgBox("Error SaveAs (" & [Enum].GetName(GetType(GuardaTipo), ultimoGT) & ") con " & fFin)
                Console.WriteLine("Error SaveAs (" & [Enum].GetName(GetType(GuardaTipo), ultimoGT) & ") con " & fFin)
            End Try
        Next
    End Sub

    Public Sub ExportarDibujosPDFAddIn(ByVal oDoc As Inventor.DrawingDocument, _
                                    ByVal queDestino As String)
        If IO.Directory.Exists(queDestino) = False Then
            Call IO.Directory.CreateDirectory(queDestino)
        End If
        If queDestino.EndsWith("\") = False Then queDestino &= "\"
        '' Nombre del fichero PDF destino.
        Dim ficheroPDF As String = queDestino & DameParteCamino(oDoc.FullFileName, ParteCamino.SoloFicheroSinExtension) & ".pdf"
        ' Get the PDF translator Add-In.
        Dim PDFAddIn As TranslatorAddIn
        PDFAddIn = Me.oAppCls.ApplicationAddIns.ItemById("{0AC6FD96-2F4D-42CE-8BE0-8AEA580399E4}")

        Dim oContext As TranslationContext
        oContext = Me.oAppCls.TransientObjects.CreateTranslationContext
        oContext.Type = IOMechanismEnum.kFileBrowseIOMechanism

        ' Create a NameValueMap object
        Dim oOptions As NameValueMap
        oOptions = Me.oAppCls.TransientObjects.CreateNameValueMap

        ' Create a DataMedium object
        Dim oDataMedium As DataMedium
        oDataMedium = Me.oAppCls.TransientObjects.CreateDataMedium

        ' Check whether the translator has 'SaveCopyAs' options
        If PDFAddIn.HasSaveCopyAsOptions(oDoc, oContext, oOptions) Then

            ' Options for drawings...

            oOptions.Value("All_Color_AS_Black") = False     'True o 1 / 0 o False

            'oOptions.Value("Remove_Line_Weights") = True	' 0 o False
            oOptions.Value("Vector_Resolution") = 400
            oOptions.Value("Sheet_Range") = PrintRangeEnum.kPrintAllSheets
            'oOptions.Value("Custom_Begin_Sheet") = 2
            'oOptions.Value("Custom_End_Sheet") = 4

        End If

        'Set the destination file name
        oDataMedium.FileName = ficheroPDF   ' "c:\temp\test.pdf"

        Try
            If IO.File.Exists(ficheroPDF) Then IO.File.Delete(ficheroPDF)
            'Publish document.
            Call PDFAddIn.SaveCopyAs(oDoc, oContext, oOptions, oDataMedium)
            'Call Process.Start(ficheroPDF)
        Catch ex As Exception
            Console.WriteLine("Error exportar PDF con " & ficheroPDF)
            Debug.Print("Error exportar DWG con " & ficheroPDF)
        End Try
    End Sub

    Public Sub ExportToSat(ByVal oDoc As Inventor.Document)
        ' Set reference to active document.
        'Dim oDoc As Inventor.Document
        'oDoc = oApp.ActiveDocument

        ' Check the Document type is an assembly or part
        If (oDoc.DocumentType <> DocumentTypeEnum.kAssemblyDocumentObject And _
          oDoc.DocumentType <> DocumentTypeEnum.kPartDocumentObject) Then
            MsgBox("Error:Document type is not assembly/part")
            oDoc = Nothing
            Exit Sub
        End If

        ' Get document's full file name
        Dim sFname As String
        sFname = oDoc.FullFileName

        ' The file format will depend on the extension
        ' Set file name extension to ".SAT"
        sFname = Microsoft.VisualBasic.Left(sFname, Len(sFname) - 3) & "sat"
        Try
            If IO.File.Exists(sFname) Then IO.File.Delete(sFname)
            ' Do a 'Save Copy As' to SAT format
            Call oDoc.SaveAs(sFname, True)
        Catch ex As Exception
            Debug.Print("Error guardando " & sFname)
        End Try

        oDoc = Nothing
    End Sub


    ' PARA EXPORTAR A DWG (Solo la opción del fichero .INI)
    ' Dim strIniFile As String
    ' strIniFile = "C:\tempDWGOut.ini"
    ' Create the name-value that specifies the ini file to use.
    ' oOptions.Value("Export_Acad_IniFile") = strIniFile
    Public Sub ExportarDibujosDWGDXFAddin(ByVal oDoc As Inventor.DrawingDocument, nombreFIN As String, ByVal dirDestino As String, ByVal ficheroINI As String, ByVal esDWG As Boolean)
        If dirDestino.EndsWith("\") = False Then dirDestino &= "\"
        '' Nombre final del fichero exportado (DWG o DXF)
        Dim ficheroFIN As String
        If esDWG = True Then
            If nombreFIN <> "" Then
                ficheroFIN = dirDestino & nombreFIN & ".dwg"
            Else
                ficheroFIN = dirDestino & DameParteCamino(oDoc.FullFileName, ParteCamino.SoloFicheroSinExtension) & ".dwg"
            End If
        Else
            If nombreFIN <> "" Then
                ficheroFIN = dirDestino & nombreFIN & ".dxf"
            Else
                ficheroFIN = dirDestino & DameParteCamino(oDoc.FullFileName, ParteCamino.SoloFicheroSinExtension) & ".dxf"
            End If
        End If
        If IO.File.Exists(ficheroFIN) = True Then Exit Sub

        ' Get the PDF translator Add-In.
        Dim DWGDXFAddIn As TranslatorAddIn
        If esDWG = True Then
            DWGDXFAddIn = oAppCls.ApplicationAddIns.ItemById("{C24E3AC2-122E-11D5-8E91-0010B541CD80}")
        Else
            DWGDXFAddIn = oAppCls.ApplicationAddIns.ItemById("{C24E3AC4-122E-11D5-8E91-0010B541CD80}")
        End If
        If DWGDXFAddIn.Activated = False Then DWGDXFAddIn.Activate()

        'For Each oSh As Sheet In oDoc.Sheets
        'clsI.ZoomTodoAjustar2D(oSh)
        'Next

        Dim oContext As TranslationContext
        oContext = oAppCls.TransientObjects.CreateTranslationContext
        oContext.Type = IOMechanismEnum.kFileBrowseIOMechanism

        ' Create a NameValueMap object
        Dim oOptions As NameValueMap
        oOptions = oAppCls.TransientObjects.CreateNameValueMap

        ' Create a DataMedium object
        Dim oDataMedium As DataMedium
        oDataMedium = oAppCls.TransientObjects.CreateDataMedium


        ' Check whether the translator has 'SaveCopyAs' options
        If DWGDXFAddIn.HasSaveCopyAsOptions(oDoc, oContext, oOptions) AndAlso IO.File.Exists(ficheroINI) = True Then
            ' Create the name-value that specifies the ini file to use.
            oOptions.Value("Export_Acad_IniFile") = ficheroINI
        End If

        'Set the destination file name
        oDataMedium.FileName = ficheroFIN   ' "c:\temp\test.pdf"

        Try
            '' Borramos el documento DWG final, si existía antes.
            If IO.File.Exists(ficheroFIN) Then IO.File.Delete(ficheroFIN)
            'Publish document.
            Call DWGDXFAddIn.SaveCopyAs(oDoc, oContext, oOptions, oDataMedium)
            '' Abrir el fichero, una vez se ha creado
            'Call Process.Start(ficheroDWG)
        Catch ex As Exception
            Console.WriteLine("Error exportar DWG con " & ficheroFIN)
            Debug.Print("Error exportar DWG con " & ficheroFIN)
        End Try
    End Sub

    Public Sub ExportarDibujosDWFAddin(ByVal oDoc As Inventor.DrawingDocument, ByVal dirDestino As String, ByVal esDWF As Boolean)
        If dirDestino.EndsWith("\") = False Then dirDestino &= "\"
        '' Nombre final del fichero exportado (dwf o dwfx)
        Dim ficheroFIN As String
        If esDWF = True Then
            ficheroFIN = dirDestino & DameParteCamino(oDoc.FullFileName, ParteCamino.SoloFicheroSinExtension) & ".dwf"
        Else
            ficheroFIN = dirDestino & DameParteCamino(oDoc.FullFileName, ParteCamino.SoloFicheroSinExtension) & ".dwfx"
        End If
        If IO.File.Exists(ficheroFIN) = True Then Exit Sub

        ' Get the DWF translator Add-In.
        Dim DWFAddIn As TranslatorAddIn
        DWFAddIn = Me.oAppCls.ApplicationAddIns.ItemById("{0AC6FD95-2F4D-42CE-8BE0-8AEA580399E4}")

        Dim oContext As TranslationContext
        oContext = Me.oAppCls.TransientObjects.CreateTranslationContext
        oContext.Type = IOMechanismEnum.kFileBrowseIOMechanism

        ' Create a NameValueMap object
        Dim oOptions As NameValueMap
        oOptions = Me.oAppCls.TransientObjects.CreateNameValueMap

        ' Create a DataMedium object
        Dim oDataMedium As DataMedium
        oDataMedium = Me.oAppCls.TransientObjects.CreateDataMedium

        ' Check whether the translator has 'SaveCopyAs' options
        If DWFAddIn.HasSaveCopyAsOptions(oDoc, oContext, oOptions) Then

            oOptions.Value("Launch_Viewer") = False     'True o 1  /  False o 0

            ' Other options...
            'oOptions.Value("Publish_All_Component_Props") = 1
            'oOptions.Value("Publish_All_Physical_Props") = 1
            'oOptions.Value("Password") = 0

            If TypeOf oDoc Is DrawingDocument Then

                ' Drawing options
                oOptions.Value("Publish_Mode") = DWFPublishModeEnum.kCustomDWFPublish
                oOptions.Value("Publish_All_Sheets") = 1    '0

                ' The specified sheets will be ignored if
                ' the option "Publish_All_Sheets" is True (1)
                'Dim oSheets As NameValueMap
                'oSheets = Me.oAppCls.TransientObjects.CreateNameValueMap

                ' Publish the first sheet AND its 3D model
                'Dim oSheet1Options As NameValueMap
                'oSheet1Options = Me.oAppCls.TransientObjects.CreateNameValueMap

                'oSheet1Options.Add("Name", "Sheet:1")
                'oSheet1Options.Add("3DModel", True)
                'oSheets.Value("Sheet1") = oSheet1Options

                ' Publish the third sheet but NOT its 3D model
                'Dim oSheet3Options As NameValueMap
                'oSheet3Options = Me.oAppCls.TransientObjects.CreateNameValueMap

                'oSheet3Options.Add("Name", "Sheet3:3")
                'oSheet3Options.Add("3DModel", False)

                'oSheets.Value("Sheet2") = oSheet3Options

                'Set the sheet options object in the oOptions NameValueMap
                'oOptions.Value("Sheets") = oSheets
            End If

        End If

        'Set the destination file name
        oDataMedium.FileName = ficheroFIN
        Try
            If IO.File.Exists(ficheroFIN) Then IO.File.Delete(ficheroFIN)
            'Publish document.
            Call DWFAddIn.SaveCopyAs(oDoc, oContext, oOptions, oDataMedium)
        Catch ex As Exception
            Console.WriteLine("Error guardando " & ficheroFIN)
            Debug.Print("Error exportar DWG con " & ficheroFIN)
        End Try
    End Sub

    Public Sub TextoCrearEnHoja(ByVal oDoc As DrawingDocument, ByVal queTexto As String, Optional ByVal queSk As String = "2acad", Optional ByVal borra As Boolean = False)
        If oDoc Is Nothing Then Exit Sub
        If Me.oAppCls Is Nothing Then Exit Sub
        AppActivate(clsI.oAppCls.Caption)

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
        oTG = clsI.oAppCls.TransientGeometry

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


    Public Function DameImagen(ByVal img As IPictureDisp) As System.Drawing.Image
        Dim oImage As System.Drawing.Image
        oImage = clsIpic.GetImageFromIPictureDisp(img)
        'oImage PictureDispConverter.ToImage(oPict)
        Return oImage
    End Function

    Public Function DameImagenInventor(ByVal img As System.Drawing.Image) As IPictureDisp
        Dim oImage As IPictureDisp   ' Image
        'oImage = Microsoft.VisualBasic.Compatibility.VB6.Support.ImageToIPictureDisp(oPict)
        'oImage = PictureDispConverter.ToIPictureDisp(img)
        oImage = clsIpic.GetIPictureDispFromImage(img)

        Return oImage
    End Function

    Sub ChangeThumbnail(ByVal oDoc As Inventor.Document, ByVal queImagen As String)

        ' Set a reference to the active document
        'Dim oDoc As Document
        'oDoc = oAp.ActiveDocument

        ' Get the "Summary Information" property set
        Dim oPropSet As PropertySet
        oPropSet = oDoc.PropertySets("{F29F85E0-4FF9-1068-AB91-08002B27B3D9}")

        ' Get the "Thumbnail" property    
        Dim oProp As Inventor.Property = Nothing
        oProp = oPropSet.ItemByPropId(17)

        Dim oDisp As IPictureDisp
        oDisp = DameImagenInventor(System.Drawing.Image.FromFile(queImagen))
        'oDisp = LoadPicture("C:\temp\thumbnail.bmp")

        ' Set the value of the thumbnail file property
        oProp.Value = oDisp

        ' Save the document
        oDoc.Save2(False)
    End Sub

    Public Function DameThumbnailAprenticeDoc(ByVal queDoc As String) As System.Drawing.Image
        ' Declare the Apprentice object
        Dim oApprentice As New ApprenticeServerComponent

        ' Open a document using Apprentice
        Dim oApprenticeDoc As ApprenticeServerDocument
        oApprenticeDoc = oApprentice.Open(queDoc)  '  "C:\Test\part.ipt")

        '' Tiempo de espera 40 segundos máximo.
        Dim tiempo As Date = Date.Now
        While oApprenticeDoc.Thumbnail Is Nothing
            If oApprenticeDoc.Thumbnail IsNot Nothing Then Exit While
            If Date.Now > tiempo.AddSeconds(40) Then Exit While
        End While
        Dim oImagen As System.Drawing.Image = Nothing
        oImagen = DameImagen(oApprenticeDoc.Thumbnail)
        oApprentice.Close()

        DameThumbnailAprenticeDoc = oImagen
        Exit Function
    End Function

    Public Function DameThumbnailAprenticeProp(ByVal queDoc As String) As System.Drawing.Image
        ' Declare the Apprentice object
        Dim oApprentice As New ApprenticeServerComponent

        ' Open a document using Apprentice
        Dim oApprenticeDoc As ApprenticeServerDocument
        oApprenticeDoc = oApprentice.Open(queDoc)  '  "C:\Test\part.ipt")

        '' TAMBIÉN VALDRÍA
        ' oApprenticeDoc.Thumbnail


        ' Obtain the PropertySets collection
        Dim oPropsets As PropertySets
        oPropsets = oApprenticeDoc.PropertySets

        ' Get the "Summary Information" property set
        Dim oPropSet As Inventor.PropertySet
        oPropSet = oPropsets("{F29F85E0-4FF9-1068-AB91-08002B27B3D9}")

        ' Get the "Thumbnail" property    
        Dim oProp As Inventor.Property = Nothing
        oProp = oPropSet.ItemByPropId(17)
        'Debug.Print(oProp.Type.ToString & " / " & oProp.GetType.ToString)

        Dim oImagen As System.Drawing.Image
        oImagen = DameImagen(oProp.Value)
        oApprentice.Close()

        DameThumbnailAprenticeProp = oImagen
        Exit Function
    End Function

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

        oAppCls.SilentOperation = True
        For Each queF As String In colPlanos
            '' Si el fichero está en un directorio OldVersions
            If queF.ToLower.Contains("oldversions") Then Continue For
            DoEventsInventor()
            If queF.ToLower.EndsWith(".dwg") AndAlso clsI.oAppCls.FileManager.IsInventorDWG(queF) = False Then
                Continue For
            End If

            If queTb IsNot Nothing Then queTb.Text &= queF & vbCrLf
            Dim oDocRef As Inventor.Document = Nothing
            oDocRef = clsI.oAppCls.Documents.Open(queF, False)

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
        clsI.oAppCls.SilentOperation = False

        Return colFinal
    End Function

    Public Sub CreaDibujo(ByVal oDoc As Inventor.Document, Optional ByVal quePlantilla As String = "", Optional ByVal creacotasBase As Boolean = True)
        'AppActivate(Me.titulo)
        Dim queCamino As String = oDoc.FullFileName
        'Dim oDoc As Inventor.Document = oAppCls.Documents.ItemByName(queCamino)
        '' Si no indicamos plantilla. Cogeremos el nombre plantilla por defecto para planos.
        If quePlantilla = "" Then quePlantilla = oAppCls.FileManager.GetTemplateFile(DocumentTypeEnum.kDrawingDocumentObject)
        '' Añadimos el nuevo dibujo IDW
        Dim oIdw As DrawingDocument = oAppCls.Documents.Add(DocumentTypeEnum.kDrawingDocumentObject, quePlantilla)
        '' Creamos el nombre del plano desde el nombre de la pieza o ensamblaje que indicamos "queCamino"
        Dim extension As String = DameParteCamino(quePlantilla, ParteCamino.SoloExtension)
        Dim NombrePlano As String = DameParteCamino(queCamino, ParteCamino.SoloCambiaExtension, extension)
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
        oView1 = oSheet.DrawingViews.AddBaseView(oDoc, _
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
        oView2 = oSheet.DrawingViews.AddProjectedView(oView1, _
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
        oView3 = oSheet.DrawingViews.AddProjectedView(oView1, _
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
        oView4 = oSheet.DrawingViews.AddProjectedView(oView1, _
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
                dd = vista.Parent.DrawingDimensions.GeneralDimensions.AddLinear( _
               oTg.CreatePoint2d(pt1, pt2), vista.Parent.CreateGeometryIntent(dc))
            Else
                pt1 = dc.CenterPoint.X + 2
                pt2 = dc.CenterPoint.Y + 2
                dd = vista.Parent.DrawingDimensions.GeneralDimensions.AddRadius( _
                  oTg.CreatePoint2d(pt1, pt2), vista.Parent.CreateGeometryIntent(dc))
            End If
        Next
    End Sub

    ''' <summary>
    ''' Pone un texto en la pantalla de Inventor, por un tiempo determinado en milisegundos
    ''' </summary>
    ''' <param name="queTexto">Texto a mostras</param>
    ''' <param name="queTiempo">Milisegundos que estará en pantalla</param>
    ''' <remarks></remarks>
    Public Sub TextoPonEnPantallaPon(ByVal queTexto As String, ByVal queTiempo As Integer)
        If oAppCls.Documents.Count = 0 Then Exit Sub
        ' Set a reference to the document.
        Dim oDoc As Document
        oDoc = oAppCls.ActiveEditDocument

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
            oAppCls.ActiveView.Update()
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
            oAnchorPoint = oAppCls.TransientGeometry.CreatePoint(1, 1, 1)

            ' Set the text's anchor in model space.
            oTextGraphics.Anchor = oAnchorPoint

            ' Anchor the text graphics in the view.
            Call oTextGraphics.SetViewSpaceAnchor( _
                oAnchorPoint, oAppCls.TransientGeometry.CreatePoint2d(30, 30), ViewLayoutEnum.kTopLeftViewCorner)

            ' Update the view to see the text.
            oAppCls.ActiveView.Update()
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

    Public Sub TextoPonEnPantallaBorra()
        If oAppCls.Documents.Count = 0 Then Exit Sub
        Dim contador As Integer = 0
        Try
            ' Set a reference to the document.
            Dim oDoc As Document
            oDoc = oAppCls.ActiveEditDocument

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
            If contador > 0 Then oAppCls.ActiveView.Update()
        Catch ex As Exception
            '' No hacemos nada
        End Try
    End Sub



    Public Function EstilosActualiza_EnsPiePre(ByVal queDoc As Inventor.Document, _
                                           ByVal queFichero As String) As String
        Dim resultado As String = ""
        Dim contadorL As Integer = 0
        Dim contadorM As Integer = 0
        Dim contadorR As Integer = 0
        Dim estababierto As Boolean = True
        oAppCls.SilentOperation = True

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
                queDoc = oAppCls.Documents.ItemByName(queFichero)
            Else
                queDoc = oAppCls.Documents.Open(queFichero, False)
            End If
        End If


        Dim EstilosL As Inventor.LightingStyles = Nothing
        Dim Materiales As Inventor.Materials = Nothing
        Dim Renders As Inventor.RenderStyles = Nothing
        Dim ChapaEs As Inventor.SheetMetalStyles = Nothing

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
            '' Es una chapa
            If queDoc.SubType = "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}" Then
                ChapaEs = CType(quePie.ComponentDefinition, SheetMetalComponentDefinition).SheetMetalStyles
            End If
        ElseIf queDoc.DocumentType = DocumentTypeEnum.kPresentationDocumentObject Then
            Dim quePre As PresentationDocument
            quePre = queDoc
            EstilosL = quePre.LightingStyles
            'Materiales = quePre.Materials
            Renders = quePre.RenderStyles
        End If

        '' ***** LightinStyle (Iluminacion)
        If EstilosL IsNot Nothing Then
            '' Creamos barra de progreso.
            Dim oPb1 As Inventor.ProgressBar = Me.ProgressBarInventor(True, EstilosL.Count, "Procesando Estilos...")
            Dim estilo As LightingStyle
            For Each estilo In EstilosL
                If estilo.StyleLocation = StyleLocationEnum.kLocalStyleLocation Then
                    'Debug.Print("LOCAL : " & estilo.Name & " / " & estilo.InUse)
                ElseIf estilo.StyleLocation = StyleLocationEnum.kLibraryStyleLocation Then
                    'Debug.Print("LIBRERIA : " & estilo.Name & " / " & estilo.InUse)
                ElseIf estilo.StyleLocation = StyleLocationEnum.kBothStyleLocation Then
                    'Debug.Print("AMBOS : " & estilo.Name & " / " & estilo.InUse)
                    If estilo.UpToDate = False Then
                        Call estilo.UpdateFromGlobal()
                        contadorL += 1
                    End If
                End If
                oPb1.UpdateProgress()
            Next
            oPb1.Close()
            oPb1 = Nothing
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
                        Call material.UpdateFromGlobal()
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
                        Call render.UpdateFromGlobal()
                        contadorR += 1
                    End If
                End If
            Next
            resultado &= IIf(contadorR = 0, "", "  (" & contadorR & ") Colores")
        End If
        '' ***** Estilos de Chapa
        If ChapaEs IsNot Nothing Then
            Dim ChapaE As Inventor.SheetMetalStyle
            For Each ChapaE In ChapaEs
                If ChapaE.StyleLocation = StyleLocationEnum.kLocalStyleLocation Then
                    'Debug.Print("LOCAL : " & render.Name & " / " & render.InUse)
                ElseIf ChapaE.StyleLocation = StyleLocationEnum.kLibraryStyleLocation Then
                    'Debug.Print("LIBRERIA : " & render.Name & " / " & render.InUse)
                ElseIf ChapaE.StyleLocation = StyleLocationEnum.kBothStyleLocation Then
                    'Debug.Print("AMBOS : " & render.Name & " / " & render.InUse)
                    If ChapaE.UpToDate = False Then
                        Call ChapaE.UpdateFromGlobal()
                        contadorR += 1
                    End If
                End If
            Next
            resultado &= IIf(contadorR = 0, "", "  (" & contadorR & ") Estilo Chapas")
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
        oAppCls.SilentOperation = False

        Return resultado
    End Function

    Public Function EstilosActualiza_Dib(ByRef queDib As DrawingDocument, _
                                           ByVal queFichero As String) As String
        Dim resultado As String = ""
        Dim contador As Integer = 0
        Dim estababierto As Boolean = True

        oAppCls.SilentOperation = True
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
                queDib = oAppCls.Documents.ItemByName(queFichero)
            Else
                queDib = oAppCls.Documents.Open(queFichero, False)
            End If
        End If

        Dim Estilos As Inventor.Styles
        Estilos = queDib.StylesManager.Styles

        If Estilos IsNot Nothing Then
            '' Creamos una barra de progreso en Inventor. En la barra de tareas.
            Dim oPb1 As Inventor.ProgressBar = Me.ProgressBarInventor(True, Estilos.Count, "Procesando estilos...")
            ''
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
                oPb1.UpdateProgress()
            Next
            oPb1.Close()
            oPb1 = Nothing
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
        oAppCls.SilentOperation = False

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
        If oAppCls.SoftwareVersion.Major <= 14 Then
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

    Public Enum nAtributo As Integer
        alineacion = 0
        punto1 = 1
        punto2 = 2
        cara_cabeza = 3
        alt_men = 4
        quePadre = 5
        pilar = 6       '(nombre del pilar)
        tirada = 7      ' Número de tirada de arriba a abajo
        seleccion = 8   ' pilar & cara & tirada
    End Enum

    Public Enum OperacionValor As Integer
        cambiar = 0
        sumar = 1
        restar = 2
    End Enum

    Public Enum GuardaTipo As Integer
        dwg = 0
        dxf = 1
        dwf = 2
        dwfx = 3
        pdf = 4
        bmp = 5
        gif = 6
        jpg = 7
        png = 8
        tiff = 9
    End Enum

    Public Enum EstructuraBOM
        '("Sin nombre")  ("Estructurado")   ("Sólo piezas")
        Modelo = Inventor.BOMViewTypeEnum.kModelDataBOMViewType
        Estructurado = Inventor.BOMViewTypeEnum.kStructuredBOMViewType
        Piezas = Inventor.BOMViewTypeEnum.kPartsOnlyBOMViewType
    End Enum

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

    Public Sub BOMActiva(ByRef oEns As AssemblyDocument, Optional queVista As EstructuraBOM = EstructuraBOM.Piezas)
        Dim oBom As BOM = oEns.ComponentDefinition.BOM
        If oBom.StructuredViewFirstLevelOnly = True Then oBom.StructuredViewFirstLevelOnly = False
        If oBom.StructuredViewEnabled = False Then oBom.StructuredViewEnabled = True
        If oBom.PartsOnlyViewEnabled = False Then oBom.PartsOnlyViewEnabled = True
    End Sub

    Public Function BOMRowBuscaCantidadIAM(ByVal oBOMRows As BOMRowsEnumerator, fullfi As String) As Double
        Dim resultado As Double = 0
        ''
        If oBOMRows.Count = 0 Then
            Return resultado
            Exit Function
        End If
        ''
        Dim i As Long
        For i = 1 To oBOMRows.Count
            ' Get the current row.
            Dim oRow As BOMRow
            oRow = oBOMRows.Item(i)
            'Set a reference to the primary ComponentDefinition of the row
            Dim oCompDef As Inventor.ComponentDefinition
            oCompDef = oRow.ComponentDefinitions.Item(1)
            Dim od As Inventor.Document
            od = oCompDef.Document  ' co.Definition.Document

            If TypeOf oCompDef Is VirtualComponentDefinition Or IO.Path.GetExtension(od.FullFileName) <> IO.Path.GetExtension(fullfi) Then
                ' No hacemos nada.
                Continue For
            Else
                '' Si no es un ensamblaje, pasamos al siguiente.
                ''
                If od.FullFileName = fullfi Then
                    Dim CantUni As Double = CDbl(oRow.ItemQuantity)
                    Dim CantTot As Double = CDbl(oRow.TotalQuantity)
                    If CantTot > CantUni Then
                        resultado += CantTot
                    Else
                        resultado += CantUni
                    End If
                Else
                    'Si tiene hijos
                    If Not oRow.ChildRows Is Nothing Then
                        resultado += BOMRowBuscaIAMHaciaAbajo(oRow.ChildRows, fullfi)
                    End If
                End If
            End If
        Next
        ''
        Return resultado
    End Function

    Public Function BOMRowBuscaIAMHaciaAbajo(ByVal oBOMRows As BOMRowsEnumerator, fullfi As String) As Double
        Dim resultado As Double = 0
        Dim i As Long
        For i = 1 To oBOMRows.Count
            ' Get the current row.
            Dim oRow As BOMRow
            oRow = oBOMRows.Item(i)
            'Set a reference to the primary ComponentDefinition of the row
            Dim oCompDef As Inventor.ComponentDefinition
            oCompDef = oRow.ComponentDefinitions.Item(1)
            Dim od As Inventor.Document
            od = oCompDef.Document  ' co.Definition.Document
            If TypeOf oCompDef Is VirtualComponentDefinition Or IO.Path.GetExtension(od.FullFileName) <> IO.Path.GetExtension(fullfi) Then
                ' No hacemos nada.
                Continue For
            End If
            '' Si no es un ensamblaje, pasamos al siguiente.
            ''
            If od.FullFileName = fullfi Then
                Dim CantUni As Double = CDbl(oRow.ItemQuantity)
                Dim CantTot As Double = CDbl(oRow.TotalQuantity)
                Dim cantidadFinal As Double = 0
                If CantTot > CantUni Then
                    cantidadFinal = CantTot
                Else
                    cantidadFinal = CantUni
                End If
                resultado = BOMRowBuscaCantidadIAMHaciaArriba(oRow, cantidadFinal)
            Else
                'Si tiene hijos
                If Not oRow.ChildRows Is Nothing Then
                    resultado = BOMRowBuscaIAMHaciaAbajo(oRow.ChildRows, fullfi)
                End If
            End If
        Next
        Return resultado
    End Function

    Public Function BOMRowBuscaCantidadIAMHaciaArriba(oRow As BOMRow, cantidad As Integer) As Double
        Dim resultado As Double = cantidad
        ''
        ''
        If TypeOf oRow.Parent Is BOMView Then
            'resultado *= CDbl(CType(oRow.Parent, BOMRow).ItemQuantity)
            'resultado
            'Exit Function
        ElseIf TypeOf oRow.Parent Is BOMRow Then
            resultado *= CDbl(CType(oRow.Parent, BOMRow).ItemQuantity)
            BOMRowBuscaCantidadIAMHaciaArriba(oRow.Parent, resultado)
        End If
        ''
        Return resultado
    End Function

    ''' <summary>
    ''' Le damos objeto Document y le indicamos que propiedad queremos: "Peso", "Volumen" o "Area"
    ''' </summary>
    ''' <param name="oDoc">Objecto Inventor.Document</param>
    ''' <param name="queDoy">string "Peso", "Volumen" o "Area"</param>
    ''' <returns>una cadena con el valor de la propiedad solicitada</returns>
    ''' <remarks></remarks>
    Public Function PesoDame(oDoc As Inventor.Document, Optional queDoy As String = "Peso") As String
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
        oMp.CacheResultsOnCompute = False
        '' *********************************
        Select Case queDoy
            Case "Peso"
                resultado = oMp.Mass    ' clsI.PropiedadLeeDesignTracking(oDoc, "Mass")  '(oCoGeo.ReferencedDocumentDescriptor.ReferencedDocument, "Mass")
            Case "Volumen"
                resultado = oMp.Volume
            Case "Area"
                resultado = oMp.Area
        End Select

        Return resultado
    End Function

    ''' <summary>
    ''' Le damos objeto MassProperties del objeto y le indicamos que propiedad queremos: "Peso", "Volumen" o "Area"
    ''' </summary>
    ''' <param name="oMp">objetoc MassProperties</param>
    ''' <param name="queDoy">"Peso", "Volumen" o "Area"</param>
    ''' <returns>una cadena con el valor de la propiedad solicitada</returns>
    ''' <remarks></remarks>
    Public Function PesoDameCom(oMp As MassProperties, Optional queDoy As String = "Peso") As String
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
                resultado = oMp.Mass    ' clsI.PropiedadLeeDesignTracking(oDoc, "Mass")  '(oCoGeo.ReferencedDocumentDescriptor.ReferencedDocument, "Mass")
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
        oProgressBar = oAppCls.CreateProgressBar(EnBarraTareas, TotalPasos, Titulo)

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

    Public Function EstaEnDirectoriosProyecto(queF As String) As Boolean
        Dim resultado As Boolean = False
        arrDirsPro = New ArrayList
        Dim oDpM As DesignProjectManager = oAppCls.DesignProjectManager
        '' Directorio del proyecto
        arrDirsPro.Add(IO.Path.GetDirectoryName(oDpM.ActiveDesignProject.FullFileName))
        '' Directorio del espacio de trabajo (debería ser igual que el anterior)
        If arrDirsPro.Contains(oDpM.ActiveDesignProject.WorkspacePath) = False Then _
            arrDirsPro.Add(oDpM.ActiveDesignProject.WorkspacePath)
        '' Directorios del grupo de trabajo
        For Each ProPath As ProjectPath In oDpM.ActiveDesignProject.WorkgroupPaths
            If arrDirsPro.Contains(ProPath.Path) = False Then arrDirsPro.Add(ProPath.Path)
        Next
        ''
        '' Ya tenemos todos los directorios de trabajo y acceso del proyecto (WorkSpace y Grupos de trabajo)
        For Each camino In arrDirsPro
            If queF.StartsWith(camino) Then
                resultado = True
                Exit For
            End If
        Next
        Return resultado
    End Function
End Class

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

#Region "PROPIEDADES"
'Public Sub Propiedades()
'    '***** Declare the Application object
'    Dim oApplication As Inventor.Application

'    ' Obtain the Inventor Application object.
'    ' This assumes Inventor is already running.
'    oApplication = GetObject(, "Inventor.Application")

'    ' Set a reference to the active document.
'    ' This assumes a document is open.
'    Dim oDoc As Document
'    oDoc = oApplication.ActiveDocument

'    ' Obtain the PropertySets collection object
'    Dim oPropsets As PropertySets
'    oPropsets = oDoc.PropertySets

'    '***** Iterate through all the PropertySets one by one using for loop
'    Dim oPropSet As PropertySet
'    For Each oPropSet In oPropsets
'        Dim Nombre As String
'        ' Obtain the DisplayName of the PropertySet
'        'Debug.Print "Display name: " & oPropSet.DisplayName
'        Nombre = oPropSet.DisplayName & " / "

'        ' Obtain the InternalName of the PropertySet
'        'Debug.Print "Internal name: " & oPropSet.InternalName
'        Nombre = Nombre & oPropSet.DisplayName '& vbCrLf

'        Debug.Print("" & Nombre & "")

'        ' Write a blank line to separate each pair.
'        Debug.Print()

'        '***** Todas las Propiedades
'        'Dim oPropSet As PropertySet
'        'For Each oPropSet In oPropsets
'        ' Iterate through all the Properties in the current set.
'        Dim oProp As Property
'        For Each oProp In oPropSet
'            ' Obtain the Name of the Property
'            Dim Name As String
'            Name = oProp.Name

'            ' Obtain the Value of the Property
'            Dim Value As Object
'            On Error Resume Next
'            Value = oProp.Value

'            ' Obtain the PropertyId of the Property
'            Dim PropertyId As Long
'            PropertyId = oProp.PropId
'            Debug.Print(vbTab & "Nombre: " & Name & " (" & oProp.DisplayName & ") / Valor: " & CStr(Value) & " / Id: " & CStr(PropertyId)) '& vbCrLf
'        Next
'        'Next
'        Nombre = "" : Name = "" : Value = Nohting : PropertyId = 0
'    Next
'    ' Write a blank line to separate each pair.
'    Debug.Print()
'End Sub


''***** RESULTADO DEL PROCEDIMIENTO QUE SE IMPRIME. Es una chapa *****

'Información de resumen de Inventor / Información de resumen de Inventor
'Internal name: {F29F85E0-4FF9-1068-AB91-08002B27B3D9}

'    Nombre: Title (Título) / Valor:  / Id: 2
'    Nombre: Subject (Asunto) / Valor:  / Id: 3
'    Nombre: Author (Autor) / Valor: Raul / Id: 4
'    Nombre: Keywords (Palabras clave) / Valor:  / Id: 5
'    Nombre: Comments (Comentarios) / Valor:  / Id: 6
'    Nombre: Last Saved By (Guardado por última vez por) / Valor:  / Id: 8
'    Nombre: Revision Number (Nº de revisión) / Valor:  / Id: 9
'    Nombre: Thumbnail (Miniatura) / Valor:  / Id: 17
'Información resumen documento Inventor / Información resumen documento Inventor
'Internal name: {D5CDD502-2E9C-101B-9397-08002B2CF9AE}

'    Nombre: Category (Categoría) / Valor:  / Id: 2
'    Nombre: Manager (Responsable) / Valor:  / Id: 14
'    Nombre: Company (Empresa) / Valor:  / Id: 15
'Propiedades de Design Tracking / Propiedades de Design Tracking
'Internal name: {32853F0F-3444-11D1-9E93-0060B03C1CA6}


'    Nombre: Creation Time (Fecha de creación) / Valor: 22/04/2008 8:05:14 / Id: 4
'    Nombre: Part Number (Nº de pieza) / Valor: FRONTAL_Grosor / Id: 5
'    Nombre: Project (Proyecto) / Valor:  / Id: 7
'    Nombre: Cost Center (Centro de costes) / Valor:  / Id: 9
'    Nombre: Checked By (Revisado por) / Valor:  / Id: 10
'    Nombre: Date Checked (Fecha de comprobación) / Valor: 01/01/1601 / Id: 11
'    Nombre: Engr Approved By (ING. aprobada por) / Valor:  / Id: 12
'    Nombre: Engr Date Approved (Fecha de aprobación de diseño ing.) / Valor: 01/01/1601 / Id: 13
'    Nombre: User Status (Estado del usuario) / Valor:  / Id: 17
'    Nombre: Material (Material) / Valor: Scotch / Id: 20
'    Nombre: Part Property Revision Id (Revisión de la pieza) / Valor: {827906D5-CB5E-4C98-B02F-7F109188604C} / Id: 21
'    Nombre: Catalog Web Link (Enlace Web de catálogo) / Valor:  / Id: 23
'    Nombre: Part Icon (Icono de la pieza) / Valor:  / Id: 28
'    Nombre: Description (Descripción) / Valor:  / Id: 29
'    Nombre: Vendor (Proveedor) / Valor:  / Id: 30
'    Nombre: Document SubType (Tipo de pieza) / Valor: {9C464203-9BAE-11D3-8BAD-0060B0CE6BB4} / Id: 31
'    Nombre: Document SubType Name (Nombre del tipo de pieza) / Valor: Chapa / Id: 32
'    Nombre: Proxy Refresh Date (Fecha de actualización de proxy) / Valor: 01/01/1601 / Id: 33
'    Nombre: Mfg Approved By (FAB. aprobada por) / Valor:  / Id: 34
'    Nombre: Mfg Date Approved (Fecha de aprobación de fabricación) / Valor: 01/01/1601 / Id: 35
'    Nombre: Cost (Coste) / Valor: 0 / Id: 36
'    Nombre: Standard (Norma) / Valor:  / Id: 37
'    Nombre: Design Status (Estado del diseño) / Valor: 1 / Id: 40
'    Nombre: Designer (Diseñador) / Valor: Raul / Id: 41
'    Nombre: Engineer (Ingeniero) / Valor:  / Id: 42
'    Nombre: Authority (Responsable) / Valor:  / Id: 43
'    Nombre: Parameterized Template (Plantilla parametrizada) / Valor: False / Id: 44
'    Nombre: Template Row (Fila de la plantilla) / Valor:  / Id: 45
'    Nombre: External Property Revision Id (Revisión externa de la pieza) / Valor: {4D29B490-49B2-11D0-93C3-7E0706000000} / Id: 46
'    Nombre: Standard Revision (Revisión de la norma) / Valor:  / Id: 47
'    Nombre: Manufacturer (Fabricante) / Valor:  / Id: 48
'    Nombre: Standards Organization (Organismo de normalización) / Valor:  / Id: 49
'    Nombre: Language (Idioma) / Valor:  / Id: 50
'    Nombre: Defer Updates (Aplazar actualizaciones) / Valor: False / Id: 51
'    Nombre: Standard Revision (Revisión de la norma) / Valor:  / Id: 47
'    Nombre: Manufacturer (Fabricante) / Valor:  / Id: 48
'    Nombre: Standards Organization (Organismo de normalización) / Valor:  / Id: 49
'    Nombre: Language (Idioma) / Valor:  / Id: 50
'    Nombre: Defer Updates (Aplazar actualizaciones) / Valor: False / Id: 51
'    Nombre: Size Designation (Designación del tamaño) / Valor:  / Id: 52
'    Nombre: Categories (Categorias) / Valor:  / Id: 56
'    Nombre: Stock Number (Nº de almacenamiento) / Valor:  / Id: 55
'    Nombre: Weld Material (Material de soldadura) / Valor:  / Id: 57
'    Nombre: Mass (Masa) / Valor: 867,514997290746 / Id: 58
'    Nombre: SurfaceArea (Área de superficie) / Valor: 2775,73350191644 / Id: 59
'    Nombre: Volume (Volumen) / Valor: 110,511464623025 / Id: 60
'    Nombre: Density (Densidad) / Valor: 7,85 / Id: 61
'    Nombre: Valid MassProps (Propiedades másicas válidas) / Valor: 31 / Id: 62
'    Nombre: Flat Pattern Width (FlatPatternExtentsWidth) / Valor: 25,9828672105435 / Id: 63
'    Nombre: Flat Pattern Length (FlatPatternExtentsLength) / Valor: 54,1219114736935 / Id: 64
'    Nombre: Flat Pattern Area (FlatPatternExtentsArea) / Valor: 1406,24243900177 / Id: 65
'Propiedades de Inventor definidas por el usuario / Propiedades de Inventor definidas por el usuario
'Internal name: {D5CDD505-2E9C-101B-9397-08002B2CF9AE}

'    Nombre: ExtensionX (ExtensionX) / Valor: 542 mm / Id: 3
'    Nombre: ExtensionY (ExtensionY) / Valor: 261 mm / Id: 4
'    Nombre: DENOMINACION (DENOMINACION) / Valor: CHAPA / Id: 6
'    Nombre: LETRA (LETRA) / Valor:  / Id: 7
'    Nombre: NºORDEN (NºORDEN) / Valor: 0 / Id: 8
'    Nombre: ELEMENTO (ELEMENTO) / Valor: 0 / Id: 10
'    Nombre: Espesor (Espesor) / Valor: 0,8000 mm / Id: 13
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