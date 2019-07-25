Option Compare Text

Imports Inventor
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports Microsoft.Win32
Imports System.Linq
''Imports System.IO
Imports Microsoft.VisualBasic
Imports Microsoft.WindowsAPICodePack.Shell
Imports System.Collections.Generic
'
Partial Public Class Inventor2acad
    Public Sub ComponentOccurrence_ClearAppearanceOverrides(assDef As AssemblyComponentDefinition)
        For Each oCo As ComponentOccurrence In assDef.Occurrences
            If oCo.Enabled = False Then Continue For
            If oCo.Excluded = True Then Continue For
            If oCo.Visible = False Then Continue For
            '
            On Error Resume Next
            If oCo.ReferencedDocumentDescriptor.ReferencedDocumentType = DocumentTypeEnum.kPartDocumentObject Then
                CType(oCo.Definition, PartComponentDefinition).ClearAppearanceOverrides()
            ElseIf oCo.ReferencedDocumentDescriptor.ReferencedDocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                ComponentOccurrence_ClearAppearanceOverrides(CType(oCo.Definition, AssemblyComponentDefinition))
            End If
            oAppI.UserInterfaceManager.DoEvents()
        Next
    End Sub
    '' Devuelve un Arraylist de Componentoccurrences
    Public Function ComponentOccurrences_DameTODOS(ByVal oEns As Inventor.AssemblyDocument, ByVal primernivel As Boolean, Optional ByVal nivelPrincipal As Boolean = True) As ArrayList
        Dim resultado As New ArrayList
        '' Para guardar representacion nivel detalle activa.
        Dim repActiva As LevelOfDetailRepresentation = Nothing
        repActiva = oEns.ComponentDefinition.RepresentationsManager.ActiveLevelOfDetailRepresentation

        If nivelPrincipal = True Then RepresentacionActivaCrea(oEns, False)
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
                oAppI.UserInterfaceManager.DoEvents()
            Next
        Else
            For Each oOcu In oEns.ComponentDefinition.Occurrences.AllReferencedOccurrences(oEns.ComponentDefinition)
                On Error Resume Next
                Dim oD As Inventor.Document = oOcu.ReferencedDocumentDescriptor.ReferencedDocument
                If (oD Is Nothing) Then Continue For
                If oD.FullFileName = "" Then Continue For
                resultado.Add(oOcu)
                oAppI.UserInterfaceManager.DoEvents()
            Next
        End If

        RepresentacionActivaCrea(oEns, True, repActiva.Name)
        ComponentOccurrences_DameTODOS = resultado
        Exit Function
    End Function
    ''
    '' Declarar e iniciar antes queCol = new System.Collections.Generic.iList(Of ComponentOccurrence)
    Public Sub ComponentOccurrencesListRecursivo(oACd As AssemblyComponentDefinition, queCol As System.Collections.Generic.List(Of ComponentOccurrence))
        For Each oCo As ComponentOccurrence In oACd.Occurrences
            queCol.Add(oCo)
            If Not (oCo.SubOccurrences Is Nothing) And oCo.ReferencedDocumentDescriptor.ReferencedDocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                ComponentOccurrencesListRecursivo(oCo.Definition, queCol)
            End If
            oAppI.UserInterfaceManager.DoEvents()
        Next
    End Sub
    Public Function ComponentOccurrences_DameTODOSrecursivo(ByVal oEns As Inventor.AssemblyDocument, ByVal primernivel As Boolean,
                                              Optional soloPrimerComponente As Boolean = False,
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
                        Call ComponentOccurrences_DameTODOSrecursivoSub(oOcu, resultado, soloPiezas)
                    End If
                Catch ex As Exception
                    Continue For
                End Try
                oAppI.UserInterfaceManager.DoEvents()
            Next
        End If
FINAL:
        oOcu = Nothing
        Return resultado
    End Function

    ' This function is called for processing sub assembly.  It is called recursively
    ' to iterate through the entire assembly tree.
    Private Sub ComponentOccurrences_DameTODOSrecursivoSub(ByVal oCompOcc As ComponentOccurrence, ByRef queArray As ArrayList, Optional soloPiezas As Boolean = False)
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
                    Call ComponentOccurrences_DameTODOSrecursivoSub(oSubCompOcc, queArray)
                End If
            Catch ex As Exception
                Continue For
            End Try
            oAppI.UserInterfaceManager.DoEvents()
        Next
    End Sub
    ''
    Public Sub ComponentOccurrence_Borra(queAsm As AssemblyDocument,
                                    todos As Boolean,
                                    queCarpeta As String,
                                    Optional listaExcluidos As Collections.Generic.List(Of String) = Nothing,
                                    Optional quetipo As DocumentTypeEnum = DocumentTypeEnum.kPartDocumentObject Or DocumentTypeEnum.kAssemblyDocumentObject,
                                    Optional borrarFichero As Boolean = True)
        ''
        Dim listaBorrar As New Collections.Generic.List(Of String)
        For Each oCo As ComponentOccurrence In queAsm.ComponentDefinition.Occurrences
            Try
                Dim fullFi As String = oCo.ReferencedFileDescriptor.FullFileName
                Dim folder As String = IO.Path.GetDirectoryName(fullFi)
                If fullFi = "" OrElse IO.File.Exists(fullFi) = False Then Continue For
                If listaExcluidos IsNot Nothing AndAlso
                listaExcluidos.Count > 0 AndAlso
                listaExcluidos.Contains(fullFi) Then Continue For
                ''
                If todos = True Then
                    If listaBorrar.Contains(fullFi) = False And folder = queCarpeta Then
                        listaBorrar.Add(fullFi)
                    End If
                    oCo.Delete()
                ElseIf todos = False And oCo.DefinitionDocumentType = quetipo Then
                    If listaBorrar.Contains(fullFi) = False And folder = queCarpeta Then
                        listaBorrar.Add(fullFi)
                    End If
                    oCo.Delete()
                Else
                    Continue For
                End If
            Catch ex As Exception
                Continue For
            End Try
            oAppI.UserInterfaceManager.DoEvents()
        Next
        ''
        queAsm.ReleaseReference()
        queAsm.Rebuild2()
        queAsm.Update2()
        queAsm.Save2()
        ''
        If listaBorrar.Count > 0 Then
            '' Recorrer todos los documentos abiertos, para cerrar los que haya que borrar
            For Each oDAbierto As Document In oAppI.Documents
                Try
                    If listaBorrar.Contains(oDAbierto.FullFileName) Then
                        oDAbierto.ReleaseReference()
                        oDAbierto.Close(True)
                    End If
                Catch ex As Exception
                    Continue For
                End Try
            Next
            ''
            '' Recorrer listaBorrar y borrar del disco duro los ficheros.
            For Each queFi As String In listaBorrar
                Try
                    IO.File.Delete(queFi)
                Catch ex As Exception
                    Continue For
                End Try
            Next
        End If
        oAppI.Documents.CloseAll(True)
    End Sub

    ''
    Public Sub ComponentOccurrence_BorraUno(ByRef queAsmDef As AssemblyComponentDefinition,
                                       ByRef queComBorro As ComponentOccurrence,
                                       Optional borrarFichero As Boolean = True)
        ''
        Dim fiBorrar As String = queComBorro.ReferencedFileDescriptor.FullFileName
        ''
        Try
            CType(queComBorro.ReferencedDocumentDescriptor.ReferencedDocument, Inventor.Document).ReleaseReference()
            queAsmDef.Occurrences.ItemByName(queComBorro.Name).Delete()
        Catch ex As Exception
            Exit Sub
        End Try
        ''
        'queAsmDef.Rebuild2()
        'queAsmDef.Update2()
        'queAsmDef.Save2()
        ''
        If borrarFichero = True Then
            '' Recorrer todos los documentos abiertos, para cerrar los que haya que borrar
            For Each oDAbierto As Document In oAppI.Documents
                Try
                    If oDAbierto.FullFileName.ToUpper = fiBorrar.ToUpper Then
                        oDAbierto.ReleaseReference()
                        oDAbierto.Close(True)
                    End If
                Catch ex As Exception
                    Continue For
                End Try
            Next
            ''
            '' Recorrer listaBorrar y borrar del disco duro los ficheros.
            Try
                IO.File.Delete(fiBorrar)
            Catch ex As Exception
                If Log Then PonLog("Error to delete --> " & fiBorrar)
            End Try
        End If
        oAppI.Documents.CloseAll(True)
    End Sub
    Public Sub ComponentOccurence_LlenaDatos(ByRef oCo As ComponentOccurrence, zMin As Double,
               Optional conMensaje As Boolean = False)
        Me.oCo = oCo
        Me.Ptmin = oTg.CreatePoint(10000, 10000, 10000)
        Me.Ptmax = oTg.CreatePoint(-10000, -10000, -10000)
        ''
        If oCo.Grounded = True Then oCo.Grounded = False
        ''
        '' Datos del ComponentOccurrence del ensamblaje
        Me.colP3D = New List(Of SketchPoint3D)
        'Me.colSbD = New Dictionary(Of SurfaceBody, ComponentOccurrence)
        'Me.colFcD = New Dictionary(Of Face, ComponentOccurrence)
        'Me.colFpD = New Dictionary(Of Face, ComponentOccurrence)
        Me.colSb = New List(Of SurfaceBodyProxy)
        Me.colFc = New List(Of Face)
        Me.colFp = New List(Of Face)
        Me.colFcProx = New List(Of FaceProxy)
        Me.colFpProx = New List(Of FaceProxy)
        Me.colFProxAll = New List(Of FaceProxy)
        '' Llenar colSB con todos los SurfaceBodyProxy
        Me.Componentoccurrence_SurfaceBodyLlenaRecursivo(oCo)    '' Llena colSb
        '' Llenar colFe con todos los Face exteriores (Cilindricos y Conicos)
        For Each oSbp As SurfaceBodyProxy In colSb
            For Each oFsh As FaceShell In oSbp.FaceShells
                '' Punto mínimo y máximo de cada SurfaceBodyProxy.
                Dim ptmi As Point = oFsh.RangeBox.MinPoint
                Dim ptma As Point = oFsh.RangeBox.MaxPoint
                ''
                If ptmi.X < Me.Ptmin.X Then Me.Ptmin.X = ptmi.X
                If ptmi.Y < Me.Ptmin.Y Then Me.Ptmin.Y = ptmi.Y
                If ptmi.Z < Me.Ptmin.Z Then Me.Ptmin.Z = ptmi.Z
                ''
                If ptma.X > Me.Ptmax.X Then Me.Ptmax.X = ptma.X
                If ptma.Y > Me.Ptmax.Y Then Me.Ptmax.Y = ptma.Y
                If ptma.Z > Me.Ptmax.Z Then Me.Ptmax.Z = ptma.Z
            Next
            ''
            For Each oFa As Face In oSbp.Faces
                Dim oFaceProxyTemp As Object = Nothing ' FaceProxy = Nothing
                oSbp.ContainingOccurrence.CreateGeometryProxy(oFa, oFaceProxyTemp)
                Dim oFaceProxy As FaceProxy = CType(oFaceProxyTemp, FaceProxy)
                ''
                Dim minArea As Double = 2   '' Area mínima para incluir en List(of
                If oFa.SurfaceType = SurfaceTypeEnum.kPlaneSurface AndAlso FacePlaneDireccionHaciaAbajo(oFa, Me.oCo.Parent) Then
                    Try
                        Dim oPlane As Plane = CType(oFa.Geometry, Plane)
                        '' Solo los que tengan un area minima
                        'If oPlane.Evaluator.Area < minArea Then
                        '    Continue For
                        'End If
                        ' Es Plane y es exterior
                        colFp.Add(oFa)
                        If oFaceProxy IsNot Nothing Then
                            colFpProx.Add(CType(oFaceProxy, FaceProxy))
                            colFProxAll.Add(oFaceProxy)
                        End If
                        'ElseIf (oFa.SurfaceType = SurfaceTypeEnum.kCylinderSurface Or
                        '    oFa.SurfaceType = SurfaceTypeEnum.kConeSurface) AndAlso
                        '    FaceEsExteriorCylinderCone(oFa) Then
                    Catch ex As Exception
                        Continue For
                    End Try
                ElseIf (oFa.SurfaceType = SurfaceTypeEnum.kCylinderSurface) AndAlso
                FaceEsExteriorCylinderCone(oFa) Then
                    Dim oCylinder As Cylinder = CType(oFa.Geometry, Cylinder)
                    Try
                        '' Solo los que tengan un area minima
                        'If oCylinder.Evaluator.Area < minArea Then
                        '    Continue For
                        'End If
                        ' Es Cylinder/Cone y es exterior
                        colFc.Add(oFa)
                        If oFaceProxy IsNot Nothing Then
                            colFcProx.Add(oFaceProxy)
                            colFProxAll.Add(oFaceProxy)
                        End If
                    Catch ex As Exception
                        Continue For
                    End Try
                End If
            Next
        Next
        ''
        largo = Ptmax.X - Ptmin.X
        ancho = Ptmax.Y - Ptmin.Y
        alto = Ptmax.Z - Ptmin.Z
        ''
        minX = Ptmin.X
        minY = Ptmin.Y
        minZ = Ptmin.Z
        maxX = Ptmax.X
        maxY = Ptmax.Y
        maxZ = Ptmax.Z
        ''
        'MsgBox("Punto minimo : " & Me.Ptmin.X & ", " & Me.Ptmin.Y & ", " & Me.Ptmin.Z & vbCrLf &
        '       "Punto máximo : " & Me.Ptmax.X & ", " & Me.Ptmax.Y & ", " & Me.Ptmax.Z & vbCrLf & vbCrLf &
        '       "largo : " & largo & vbCrLf &
        '        "ancho : " & ancho & vbCrLf &
        '         "alto : " & alto)
        '' Evaluar y mover
        'Dim oV As Vector = oTg.CreateVector(Ptmin.X, Ptmin.Y, Ptmin.Z)
        Dim corregido As Boolean = False
        ''
        Dim oMatTemp As Inventor.Matrix = oCo.Transformation
        Dim VectorX As Double = IIf(minX < 0, oMatTemp.Translation.X + (minX * -1), oMatTemp.Translation.X - minX)
        Dim VectorY As Double = IIf(minY < 0, oMatTemp.Translation.Y + (minY * -1), oMatTemp.Translation.Y - minY)
        Dim VectorZ As Double = IIf(minZ < 0, oMatTemp.Translation.Z + (minZ * -1), oMatTemp.Translation.Z - minZ)
        If (minZ < zMin / 10) Then
            corregido = True
            VectorZ = IIf(minZ < 0, oMatTemp.Translation.Z + (minZ * -1), oMatTemp.Translation.Z - minZ) + (zMin / 10)
        Else
            corregido = False
            VectorZ = oMatTemp.Translation.Z
        End If
        ''
        Dim vectorFin As Vector = oTg.CreateVector(VectorX, VectorY, VectorZ)
        oMatTemp.SetTranslation(vectorFin)
        oCo.Transformation = oMatTemp
        ''
        oCo.Grounded = True
        ''
        oAppI.ActiveEditDocument.Rebuild2()
        oAppI.ActiveView.Update()
        'Me.colP3D = Nothing
        'Me.colSb = Nothing
        'Me.colFc = Nothing
        'Me.colFp = Nothing
        If corregido = True And conMensaje = True Then
            MsgBox("Se ha corregido Z a " & zMin & " mm")
        End If
        'GoTo repite
    End Sub

    Public Sub Componentoccurrence_SurfaceBodyLlenaRecursivo(ByRef oOcu As ComponentOccurrence)  ', dicProxys As System.Collections.Generic.Dictionary(Of String, Inventor.SurfaceBody))

        If oOcu.DefinitionDocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
            '' Es un ensamblaje
            For Each oOcuHijo As ComponentOccurrence In oOcu.SubOccurrences
                Componentoccurrence_SurfaceBodyLlenaRecursivo(oOcuHijo)
            Next
        ElseIf oOcu.DefinitionDocumentType = DocumentTypeEnum.kPartDocumentObject Then
            '' Es una pieza o ensamblaje vacio
            For Each oBody As SurfaceBody In oOcu.SurfaceBodies
                Dim oBodyProxy As Object = Nothing    ' SurfaceBodyProxy = Nothing
                ' Create a proxy.
                Call oOcu.CreateGeometryProxy(oBody, oBodyProxy)
                If oBodyProxy IsNot Nothing Then
                    colSb.Add(CType(oBodyProxy, SurfaceBodyProxy))
                End If
            Next
        End If
    End Sub
    Public Function ObjectCollection_SketchPoint3DOrdena(ByRef colPt As List(Of SketchPoint3D),
                                               Optional queOrdeno As IEnum.Datos = IEnum.Datos.X,
                                               Optional orden As IEnum.Datos = IEnum.Datos.Ascending,
                                                Optional conmensaje As Boolean = False) As List(Of SketchPoint3D)
        ''
        Dim menAntes As String = ""
        For Each oObj As SketchPoint3D In colPt
            menAntes &=
            oObj.Geometry.X & ", " &
            oObj.Geometry.Y & ", " &
            oObj.Geometry.Z & vbCrLf
        Next
        'listTemp.Sort(New clsiComparerXYSketchPoint3D)
        ''
        'Dim oBox As Box = queOcu.Definition.RangeBox
        Dim qry As New System.Collections.Generic.List(Of SketchPoint3D)
        'If (oBox.MaxPoint.X - oBox.MinPoint.X) > (oBox.MaxPoint.Y - oBox.MinPoint.Y) Then
        Select Case queOrdeno
            Case IEnum.Datos.X
                If orden = IEnum.Datos.Ascending Then
                    qry = (From p As SketchPoint3D In colPt
                           Order By p.Geometry.X, p.Geometry.Y, p.Geometry.Z Ascending).ToList
                ElseIf orden = IEnum.Datos.Descending Then
                    qry = (From p As SketchPoint3D In colPt
                           Order By p.Geometry.X, p.Geometry.Y, p.Geometry.Z Descending).ToList
                End If
            Case IEnum.Datos.Y
                If orden = IEnum.Datos.Ascending Then
                    qry = (From p As SketchPoint3D In colPt
                           Order By p.Geometry.Y, p.Geometry.X, p.Geometry.Z Ascending).ToList
                ElseIf orden = IEnum.Datos.Descending Then
                    qry = (From p As SketchPoint3D In colPt
                           Order By p.Geometry.Y, p.Geometry.X, p.Geometry.Z Descending).ToList
                End If
            Case IEnum.Datos.Z
                If orden = IEnum.Datos.Ascending Then
                    qry = (From p As SketchPoint3D In colPt
                           Order By p.Geometry.Z, p.Geometry.X, p.Geometry.Y Ascending).ToList
                ElseIf orden = IEnum.Datos.Descending Then
                    qry = (From p As SketchPoint3D In colPt
                           Order By p.Geometry.Z, p.Geometry.X, p.Geometry.Y Descending).ToList
                End If
        End Select
        ''
        'colPt.Clear()
        Dim menDespues As String = ""
        For Each oObj As SketchPoint3D In qry
            'colPt.Add(oObj)
            menDespues &=
            oObj.Geometry.X & ", " &
            oObj.Geometry.Y & ", " &
            oObj.Geometry.Z & vbCrLf
        Next
        ''
        If conmensaje Then
            MsgBox("Antes : " & vbCrLf & menAntes)
            MsgBox("Despues : " & vbCrLf & menDespues)
        End If
        '
        Return qry
    End Function
End Class