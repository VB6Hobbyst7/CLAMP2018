Imports Inventor
'
Partial Public Class Inventor2acad
    Public Function ToProjectExtrude_ListFaceProxies(crearExtrusion As Boolean,
                                         oCoOut As ComponentOccurrence) As PlanarSketch
        ''
        Dim resultado As PlanarSketch = Nothing
        ''
        If oAppI.ActiveEditDocument.DocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then
            MsgBox("Only for ActiveEditDocument = Assembly...", MsgBoxStyle.Critical, "ERROR")
            Return resultado
            Exit Function
        End If
        '' queOut = "Ent" (Entidadas), "Rec" (Rectangulo 2D) o "Loo" (EdgeLoop Exterior)
        'Dim colqueOut As String() = New String() {"Ent", "Rec", "Loo"}
        ''
        Try
            ' Set a reference to the assembly component definintion.
            ' This assumes an assembly document is open.
            Dim oAsmCompDef As AssemblyComponentDefinition
            oAsmCompDef = oAppI.ActiveEditDocument.ComponentDefinition
            '
            Dim oCoOutDef As PartComponentDefinition = oCoOut.Definition
            ''
            '' 1.- ***** Primero creamos la extrusión de todas las oFace.Geometry = Plane
            Dim oFaceProxy As FaceProxy = Nothing
            'For Each oFaceProxy In clsCo.colFpProx
            'Next
            '' 2.- ***** Creamos la extrusión de todas las oFace.Geometry = Cylinder
            'For Each oFaceProxy In clsCo.colFcProx
            'Next
            '' 1.- ***** Creamos la extrusión de todas las Faces
            For Each oFaceProxy In colFProxAll
                Try
                    ' Create a proxy for the sketch in the newly created part.
                    'Dim oSketchProxyTemp As Object = Nothing    ' PlanarSketchProxy = Nothing
                    'Dim oSketchProxy As PlanarSketchProxy = Nothing
                    'Call oCoOut.CreateGeometryProxy(oSketch, oSketchProxyTemp)
                    'oSketchProxy = CType(oSketchProxyTemp, PlanarSketchProxy)
                    ''
                    '' Crear el objeto NonParametricBaseFeature con el FaceProxy
                    Dim oParFea As NonParametricBaseFeature = Nothing
                    If TypeOf oFaceProxy.Geometry Is Plane Or TypeOf oFaceProxy.Geometry Is Cylinder Then
                        oParFea = BaseFeatureSurfaceBodyDame_InFace(oFaceProxy.NativeObject, oFaceProxy.ContainingOccurrence, oCoOut)
                        If oParFea Is Nothing Then  ' OrElse TypeOf oParFea.Faces(1).Geometry Is Cone Then
                            Continue For
                        End If
                    End If
                    ''
                    Dim tipoExt As String = "SOLIDO"     ' "SURFACE"
                    Dim oSketch As PlanarSketch = Nothing
                    oSketch = oCoOutDef.Sketches.Add(oCoOutDef.WorkPlanes(3), False) '' Crear PlanarSketch en plano 3 (XY)
                    Dim oFace As Face = oParFea.Faces(1)
                    '' Si oParFea.Face(1).Geometry = Cylinder. Crear eje entre los 2 circulos
                    If TypeOf oFace.Geometry Is Cylinder Then
                        Dim oSk3D As Sketch3D = FaceCylinderConeCreaWorkPointEje_DameSketch3D(oFace, oCoOutDef)
                        If oSk3D Is Nothing Then
                            Continue For
                        Else
                            Dim escorrecto As Boolean = SurfaceCilindricalCreaContorno2DSketchLine3D(oSk3D, oCoOutDef, oSketch.Name)
                            If escorrecto = False Then
                                Continue For
                            End If
                        End If
                    ElseIf TypeOf oFace.Geometry Is Plane Then
                        Dim oSketchEnt As SketchEntity = Nothing
                        Dim oSketchEntEnum As SketchEntitiesEnumerator = Nothing
                        Dim queTipo As String = "Loo"
                        ''
                        Select Case queTipo
                            Case "Ent"          '' Proyectas todas las entidas
                                For Each oEdge As Edge In oFace.Edges
                                    oSketchEnt = oSketch.AddByProjectingEntity(oEdge)
                                Next
                            Case "Rec"          '' Crear Rectangulo 2D con el Evaluator--RangeBox (Coger solo coordenadas 2D)
                                Dim oBox As Box = oFaceProxy.Evaluator.RangeBox
                                Dim ptMin2D As Point2d = oTg.CreatePoint2d(oBox.MinPoint.X, oBox.MinPoint.Y)
                                Dim ptMax2D As Point2d = oTg.CreatePoint2d(oBox.MaxPoint.X, oBox.MaxPoint.Y)
                                oSketchEntEnum = oSketch.SketchLines.AddAsTwoPointRectangle(ptMin2D, ptMax2D)
                            Case "Loo"          '' Proyectar sólo las entidades externas
                                For Each oLoop As EdgeLoop In oFaceProxy.EdgeLoops
                                    If oLoop.IsOuterEdgeLoop = False Then Continue For
                                    ''
                                    For Each oEdge As Edge In oFace.Edges
                                        oSketchEnt = oSketch.AddByProjectingEntity(oEdge)
                                    Next
                                    Exit For
                                Next
                        End Select
                    End If
                    '
                    If crearExtrusion = True Then
                        '' Crear la extrusión contra la cara de oParFea
                        Dim oExf As ExtrudeFeature = ExtrudeFaceFeature(oParFea, oSketch)
                    End If
                    ''
                    resultado = oSketch
                Catch ex1 As Exception
                    Debug.Print(ex1.ToString)
                    Continue For
                End Try
            Next
        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.ToString)
            resultado = Nothing
        End Try
        ''
        '' Borramos todos los PlanarSketch que estén sin utilizar, Sketch3D y WorkPoints
        Dim oPCd As PartComponentDefinition = oCoOut.Definition
        For Each oPsk As PlanarSketch In oPCd.Sketches
            Try
                If oPsk.Consumed = False Then oPsk.Delete()
            Catch ex As Exception
                oPsk.Visible = False
                Continue For
            End Try
        Next
        For Each oSk3D As Sketch3D In oPCd.Sketches3D
            Try
                If oSk3D.Consumed = False Then oSk3D.Delete()
            Catch ex As Exception
                oSk3D.Visible = False
                Continue For
            End Try
        Next
        For Each oWp As WorkPoint In oPCd.WorkPoints
            Try
                If oWp.Consumed = False Then oWp.Delete()
            Catch ex As Exception
                oWp.Visible = False
                Continue For
            End Try
        Next
        For Each oWpl As WorkPlane In oPCd.WorkPlanes
            oWpl.Visible = False
        Next
        ''
        For Each oSb As SurfaceBody In oPCd.SurfaceBodies
            If oSb.IsSolid Then Continue For
            oSb.Visible = False
        Next
        ''
        For Each oNon As NonParametricBaseFeature In oPCd.Features.NonParametricBaseFeatures
            oNon.SurfaceBody.Visible = False
        Next
        ''
        oPCd.CompactModelHistory = True
        Return resultado
    End Function
    Public Function EmbossFeatureCreaCode(ByRef oCd As PartComponentDefinition, ByRef oFace As Face, CODEPIEZA As String) As EmbossFeature
        Dim oEmbossfeature As EmbossFeature = Nothing
        '
        Dim nSketch As String = "Sketch_Code"
        Dim nEmboss As String = "Emboss_Code"
        Boceto2DBorra(oCd, nSketch)
        ' Get one of the edges of the face to use as the sketch x-axis.
        Dim oEdge As Edge = oFace.Edges.Item(3)     ' El 3 corresponde con el inferior.
        ' Get the start vertex of the edge to use as the origin of the sketch.
        Dim oVertex As Vertex = oEdge.StartVertex
        ' Create a new sketch.  This last argument is set to true to cause the
        ' creation of sketch geometry from the edges of the face.
        ' Crear el boceto orientado y ponerle el texto.
        Dim oSketch As PlanarSketch = oCd.Sketches.AddWithOrientation(oFace, oEdge, True,
                                                 True, oVertex, False)

        'Dim oSketch As PlanarSketch = oCd.Sketches.Add(oFace)
        oSketch.Name = nSketch
        Dim oText As Inventor.TextBox = oSketch.TextBoxes.AddByRectangle(oTg.CreatePoint2d(0.2, 0.2), oTg.CreatePoint2d(5.2, 0.7), CODEPIEZA)
        '
        Dim oProfile As Profile = oSketch.Profiles.AddForSolid
        oSketch.UpdateProfiles()
        oAppI.ActiveView.Update()
        oEmbossfeature = oCd.Features.EmbossFeatures.AddEngraveFromFace(oProfile, 0.1, PartFeatureExtentDirectionEnum.kNegativeExtentDirection)
        oEmbossfeature.Name = nEmboss
        oAppI.ActiveView.Update()
        '
        Try
            Dim oAsset As Asset = AssetApparenceCreaDame("Rojo", System.Drawing.Color.Red)
            oEmbossfeature.Appearance = oCd.Document.AppearanceAssets.Item("Rojo")
            oAppI.ActiveView.Update()
        Catch ex As Exception
            Debug.Print(ex.ToString)
        End Try
        '
        Return oEmbossfeature
    End Function
End Class