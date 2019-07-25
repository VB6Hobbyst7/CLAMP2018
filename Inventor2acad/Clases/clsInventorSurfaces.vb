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
    Public Function BaseFeatureSurfaceBodyDame_InBodyProxy(oBodyProxy As SurfaceBodyProxy, ByRef oCo As ComponentOccurrence) As NonParametricBaseFeature
        'FaceShell, Wire, Face, Edge and Vertex
        If oBodyProxy.IsEntityValid(oBodyProxy.Faces) = True Or
        oCo.ReferencedDocumentDescriptor.ReferencedDocumentType <> DocumentTypeEnum.kPartDocumentObject Then
            Return Nothing : Exit Function
        End If
        ''
        Dim resultado As NonParametricBaseFeature = Nothing
        '' 
        Dim oCd As PartComponentDefinition = oCo.Definition
        oNopfs = oCd.Features.NonParametricBaseFeatures
        oNopfd = oNopfs.CreateDefinition
        oNopf = Nothing
        ''
        ' The selected body is a body proxy in the context of
        ' the assembly. However, there's a problem with the
        ' TransientBrep.Copy method and it creates a copy of the
        ' body that ignores the transorm.  The code below creates
        ' the copy and then performs an extra step to apply the
        ' transform.
        Dim newBody As SurfaceBody = oTBr.Copy(oBodyProxy)
        Call oTBr.Transform(newBody, oBodyProxy.ContainingOccurrence.Transformation)
        ' Transform the body into the parts space of the target occurrence.
        oMatrix = oBodyProxy.ContainingOccurrence.Transformation
        oMatrix.Invert()
        oTBr.Transform(newBody, oMatrix)

        '' Crear coleccion de objectos
        Dim oCollection As ObjectCollection = oTo.CreateObjectCollection
        oCollection.Add(newBody)

        ' Create a non-associative solid base feature in the second part.
        'Dim oFeatureDef2 As NonParametricBaseFeatureDefinition
        'Set oFeatureDef2 = oCoNd.Features.NonParametricBaseFeatures.CreateDefinition

        oNopfd.BRepEntities = oCollection
        oNopfd.OutputType = BaseFeatureOutputTypeEnum.kSurfaceOutputType
        oNopfd.TargetOccurrence = oCo
        oNopfd.IsAssociative = False
        ''
        resultado = oNopfs.AddByDefinition(oNopfd)
        ''
        Return resultado
    End Function
    Public Function BaseFeatureSurfaceBodyDame_InBody(oBody As SurfaceBody, ByRef oCoIn As ComponentOccurrence, ByRef oCoOut As ComponentOccurrence) As NonParametricBaseFeature
        'FaceShell, Wire, Face, Edge and Vertex
        'If oBody.IsEntityValid(oBody.Faces) = True Or
        '    oCoOut.ReferencedDocumentDescriptor.ReferencedDocumentType <> DocumentTypeEnum.kPartDocumentObject Then
        '    Return Nothing : Exit Function
        'End If
        ''
        Dim resultado As NonParametricBaseFeature = Nothing
        '' 
        Dim oCd As PartComponentDefinition = oCoOut.Definition
        oNopfs = oCd.Features.NonParametricBaseFeatures
        oNopfd = oNopfs.CreateDefinition
        oNopf = Nothing
        ''
        Dim oBodyProxy As SurfaceBodyProxy = Nothing
        oCoIn.CreateGeometryProxy(oBody, CType(oBodyProxy, SurfaceBodyProxy))
        ' The selected body is a body proxy in the context of
        ' the assembly. However, there's a problem with the
        ' TransientBrep.Copy method and it creates a copy of the
        ' body that ignores the transorm.  The code below creates
        ' the copy and then performs an extra step to apply the
        ' transform.
        Dim newBody As SurfaceBody = oTBr.Copy(oBodyProxy)
        Call oTBr.Transform(newBody, oBodyProxy.ContainingOccurrence.Transformation)
        ' Transform the body into the parts space of the target occurrence.
        oMatrix = oBodyProxy.ContainingOccurrence.Transformation
        oMatrix.Invert()
        oTBr.Transform(newBody, oMatrix)

        '' Crear coleccion de objectos
        Dim oCollection As ObjectCollection = oTo.CreateObjectCollection
        oCollection.Add(newBody)

        ' Create a non-associative solid base feature in the second part.
        'Dim oFeatureDef2 As NonParametricBaseFeatureDefinition
        'Set oFeatureDef2 = oCoNd.Features.NonParametricBaseFeatures.CreateDefinition

        oNopfd.BRepEntities = oCollection
        oNopfd.OutputType = BaseFeatureOutputTypeEnum.kSurfaceOutputType
        oNopfd.TargetOccurrence = oCoOut
        oNopfd.IsAssociative = False
        ''
        resultado = oNopfs.AddByDefinition(oNopfd)
        ''
        Return resultado
    End Function
    Public Function BaseFeatureSurfaceBodyDame_InFaceProxy(ByRef oFaceProxy As FaceProxy, ByRef oCo As ComponentOccurrence) As NonParametricBaseFeature
        'FaceShell, Wire, Face, Edge and Vertex
        'If oFaceProxy.IsParamReversed Or
        '    oCo.ReferencedDocumentDescriptor.ReferencedDocumentType <> DocumentTypeEnum.kPartDocumentObject Then
        '    Return Nothing : Exit Function
        'End If
        ''
        Dim resultado As NonParametricBaseFeature = Nothing
        '' 
        Dim oCd As PartComponentDefinition = oCo.Definition
        oNopfs = oCd.Features.NonParametricBaseFeatures
        oNopfd = oNopfs.CreateDefinition
        oNopf = Nothing
        ''
        ' The selected body is a body proxy in the context of
        ' the assembly. However, there's a problem with the
        ' TransientBrep.Copy method and it creates a copy of the
        ' body that ignores the transorm.  The code below creates
        ' the copy and then performs an extra step to apply the
        ' transform.
        Dim newFace As Face = oTBr.Copy(oFaceProxy)
        Call oTBr.Transform(newFace, oFaceProxy.ContainingOccurrence.Transformation)
        ' Transform the body into the parts space of the target occurrence.
        oMatrix = oFaceProxy.ContainingOccurrence.Transformation
        oMatrix.Invert()
        oTBr.Transform(newFace, oMatrix)

        '' Crear coleccion de objectos
        Dim oCollection As ObjectCollection = oTo.CreateObjectCollection
        oCollection.Add(newFace)

        ' Create a non-associative solid base feature in the second part.
        'Dim oFeatureDef2 As NonParametricBaseFeatureDefinition
        'Set oFeatureDef2 = oCoNd.Features.NonParametricBaseFeatures.CreateDefinition

        oNopfd.BRepEntities = oCollection
        oNopfd.OutputType = BaseFeatureOutputTypeEnum.kSurfaceOutputType
        oNopfd.TargetOccurrence = oCo
        oNopfd.IsAssociative = False
        ''
        resultado = oNopfs.AddByDefinition(oNopfd)
        ''
        '' Si oNopfs.Face(1).Geometry = Cylinder o Cono. Crear eje entre los 2 circulos
        Call FaceCylinderConeCreaWorkPointEje_DameSketch3D(resultado.Faces(1), oCd)
        ''
        Return resultado
    End Function
    Public Function BaseFeatureSurfaceBodyDame_InFace(oFace As Face,
                                                  ByRef oCoIn As ComponentOccurrence,
                                                  ByRef oCoOut As ComponentOccurrence) As NonParametricBaseFeature
        'FaceShell, Wire, Face, Edge and Vertex
        'If oFace.IsEntityValid(Faces) = True Or
        '    oCoOut.ReferencedDocumentDescriptor.ReferencedDocumentType <> DocumentTypeEnum.kPartDocumentObject Then
        '    Return Nothing : Exit Function
        'End If
        ''
        Dim resultado As NonParametricBaseFeature = Nothing
        '' 
        Dim oCd As PartComponentDefinition = oCoOut.Definition
        oNopfs = oCd.Features.NonParametricBaseFeatures
        oNopfd = oNopfs.CreateDefinition
        oNopf = Nothing
        ''
        'Dim oFaceProxyTemp As Object = Nothing  ' FaceProxy = Nothing
        'oCoIn.CreateGeometryProxy(oFace, oFaceProxyTemp)
        'Dim oFaceProxy As FaceProxy = CType(oFaceProxyTemp, FaceProxy)
        ' The selected body is a body proxy in the context of
        ' the assembly. However, there's a problem with the
        ' TransientBrep.Copy method and it creates a copy of the
        ' body that ignores the transorm.  The code below creates
        ' the copy and then performs an extra step to apply the
        ' transform.
        'Dim newFace As Face = oTBr.Copy(oFace)
        Dim newFace As Object = oTBr.Copy(oFace)
        Call oTBr.Transform(newFace, oCoIn.Transformation)
        ' Transform the body into the parts space of the target occurrence.
        'oMatrix = oCoIn.Transformation
        'oMatrix.Invert()
        'oTBr.Transform(newFace, oMatrix)

        '' Crear coleccion de objectos
        Dim oCollection As ObjectCollection = oTo.CreateObjectCollection
        oCollection.Add(newFace)

        ' Create a non-associative solid base feature in the second part.
        'Dim oFeatureDef2 As NonParametricBaseFeatureDefinition
        'Set oFeatureDef2 = oCoNd.Features.NonParametricBaseFeatures.CreateDefinition

        oNopfd.BRepEntities = oCollection
        oNopfd.OutputType = BaseFeatureOutputTypeEnum.kSurfaceOutputType
        oNopfd.TargetOccurrence = oCoOut
        oNopfd.IsAssociative = False
        ''
        resultado = oNopfs.AddByDefinition(oNopfd)
        ''
        Return resultado
    End Function
    Public Function FaceCylinderConeCreaWorkPointEje_DameSketch3D(oFace As Face, oCd As PartComponentDefinition) As Sketch3D
        Dim resultado As Sketch3D = Nothing
        '' Si no es Cylinder ni Cono salimos con Nothing
        If TypeOf oFace.Geometry Is Cylinder = False And TypeOf oFace.Geometry Is Cone = False Then
            Return resultado
            Exit Function
        End If
        ''
        Dim oUv As UnitVector = Nothing
        If TypeOf oFace.Geometry Is Cylinder Then
            oUv = CType(oFace.Geometry, Cylinder).AxisVector
        ElseIf TypeOf oFace.Geometry Is Cone Then
            oUv = CType(oFace.Geometry, Cone).AxisVector
        End If
        ''
        Dim colPoints As New Collections.Generic.List(Of WorkPoint)
        ''
        '' Coger los 2 circulos que estarán al inicio y final
        Dim Width As Double = 100
        For Each oEd As Edge In oFace.Edges
            If TypeOf oEd.Geometry Is Circle Then
                Dim oCir As Circle = oEd.Geometry
                If oCir.Normal.IsParallelTo(oUv, 0.1) = False Then Continue For
                ''
                Dim oWp As WorkPoint = oCd.WorkPoints.AddFixed(oCir.Center)
                colPoints.Add(oWp)
                If oCir.Radius * 1.8 < Width Then
                    Width = oCir.Radius * 1.8
                End If
            ElseIf TypeOf oEd.Geometry Is Arc3d Then
                Dim oArc As Arc3d = oEd.Geometry
                If oArc.Normal.IsParallelTo(oUv, 0.1) = False Then Continue For
                ''
                Dim oWp As WorkPoint = oCd.WorkPoints.AddFixed(oArc.Center)
                colPoints.Add(oWp)
                If oArc.Radius * 1.8 < Width Then
                    Width = oArc.Radius * 1.8
                End If
            End If
        Next
        '' Crear el parámetro "Width" y ponerle valor
        Dim oUp As UserParameter = Nothing
        Try
            oUp = oCd.Parameters.UserParameters.Item("Width")
            oUp.Value = Width
        Catch ex As Exception
            oUp = oCd.Parameters.UserParameters.AddByValue("Width", Width, oAppI.ActiveDocument.UnitsOfMeasure.LengthUnits.ToString)
        End Try
        ''
        '' Crear la linea 3D en el boceto3d.
        resultado = oCd.Sketches3D.Add()
        ''
        Dim oSkl3D As SketchLine3D = Nothing
        If colPoints.Count >= 2 Then
            For x = 0 To colPoints.Count - 1
                'colPoints(x).Name = "WP" & x + 1
                '' Crear la linea después de crear el punto 2
                If x = 1 Then
                    oSkl3D = resultado.SketchLines3D.AddByTwoPoints(colPoints(0).Point, colPoints(1).Point)
                    Exit For
                End If
            Next
        Else
            Return Nothing
            Exit Function
        End If
        '' Si la longitud es menor de 2 cm. No crear Boceto 3D
        If oSkl3D.Length < 1.0 Then
            colPoints = Nothing
            Return Nothing
        Else
            colPoints = Nothing
            Return resultado
        End If
    End Function
    Public Function ExtrudeFaceFeature(noPf As NonParametricBaseFeature,
                                   Optional ByRef oSk As PlanarSketch = Nothing,
                                   Optional nWorkPlane As Integer = 3) As ExtrudeFeature

        Dim oCd As PartComponentDefinition = noPf.Parent
        If CType(oCd.Document, Inventor.Document).DocumentType <> DocumentTypeEnum.kPartDocumentObject Then
            Return Nothing : Exit Function
        End If
        ''
        Dim resultado As ExtrudeFeature = Nothing
        Dim oFa As Face = noPf.Faces(1)
        Dim esnuevo As Boolean = False
        ''
        Dim oProf As Profile = Nothing
        If oSk Is Nothing Then
            Dim tipoEnt As String = "LOO"
            esnuevo = True
            '' 1=YZ,2=XZ,3=XY
            oSk = oCd.Sketches.Add(oCd.WorkPlanes.Item(nWorkPlane))
            Dim oSEnt As SketchEntity = Nothing
            Dim oSketchEntEnum As SketchEntitiesEnumerator = Nothing
            ''
            Try
                Select Case tipoEnt
                    Case "ENT"
                        '' Proyectas todas las entidas
                        For Each oEd As Edge In oFa.Edges
                            oSEnt = oSk.AddByProjectingEntity(oEd)
                            If oSEnt Is Nothing Then Continue For
                        Next
                    Case "LOO"
                        '' Proyectar los Edge (Face--EdgeLoop (Solo exteriores)--Edge)
                        '' Solo el primer EdgeLoop
                        For Each oLoop As EdgeLoop In oFa.EdgeLoops
                            If oLoop.IsOuterEdgeLoop = False Then Continue For
                            ''
                            For Each oEdge As Edge In oLoop.Edges
                                oSEnt = oSk.AddByProjectingEntity(oEdge)
                            Next
                            Exit For
                        Next
                    Case "REC"
                        '' Crear Rectangulo 2D con el Evaluator--RangeBox (Coger solo coordenadas 2D)
                        Dim oBox As Box = oFa.Evaluator.RangeBox
                        Dim ptMin2D As Point2d = oTg.CreatePoint2d(oBox.MinPoint.X, oBox.MinPoint.Y)
                        Dim ptMax2D As Point2d = oTg.CreatePoint2d(oBox.MaxPoint.X, oBox.MaxPoint.Y)
                        oSketchEntEnum = oSk.SketchLines.AddAsTwoPointRectangle(ptMin2D, ptMax2D)
                End Select
                ''
                oProf = oSk.Profiles.AddForSolid(False)
            Catch ex As Exception
                If esnuevo = True Then oSk.Delete()
                Return Nothing
                Exit Function
            End Try
        ElseIf oSk IsNot Nothing Then
            esnuevo = False
            '' Crear la extrusión
            Try
                oProf = oSk.Profiles.AddForSolid()
            Catch ex As Exception
                If esnuevo = True Then oSk.Delete()
                Return Nothing
                Exit Function
            End Try
        End If
        ''
        'Dim areaminima As Double = 3
        'If oProf IsNot Nothing AndAlso oProf.RegionProperties.Area > areaminima Then
        Try
            Dim oExfd As ExtrudeDefinition = oCd.Features.ExtrudeFeatures.CreateExtrudeDefinition(oProf, PartFeatureOperationEnum.kNewBodyOperation)
            Call oExfd.SetToExtent(oFa, True)
            ''
            resultado = oCd.Features.ExtrudeFeatures.Add(oExfd)
        Catch ex As Exception
            resultado = Nothing
        End Try
        'Else
        'resultado = Nothing
        'End If
        ''
        Return resultado
    End Function
    ''
    Public Sub SurfaceCilindricalCreaContorno2DProxy(oFaProxy As FaceProxy, ByRef oSk As PlanarSketchProxy)
        ' Salimos si no ex Cylinder o Cone
        If TypeOf oFaProxy.Geometry IsNot Cylinder And TypeOf oFaProxy IsNot Cone Then Exit Sub
        ''
        Dim oVecUni As UnitVector = Nothing
        Dim oVec As Vector = Nothing
        Dim largo As Double = 0
        Dim ancho1 As Double = 0
        Dim ancho2 As Double = 0
        Dim ancho As Double = 0
        Dim basePt As Point = Nothing
        ''
        If TypeOf oFaProxy.Geometry Is Cylinder Then
            Dim oCyl As Cylinder = oFaProxy.Geometry
            oVecUni = oCyl.AxisVector
            oVec = oVecUni.AsVector
            basePt = oCyl.BasePoint
            largo = oVec.Length
            ancho1 = oCyl.Radius * 2
            ancho2 = ancho1
        ElseIf TypeOf oFaProxy.Geometry Is Cone Then
            Dim oCon As Cone = oFaProxy.Geometry
            oVecUni = oCon.AxisVector
            oVec = oVecUni.AsVector
            basePt = oCon.BasePoint
            largo = oVec.Length
            ancho1 = oCon.Radius * 2
            ancho2 = ancho1
        End If
        ''
        If oVec IsNot Nothing Then
            Dim oSl1 As SketchLine = oSk.SketchLines.AddByTwoPoints(
            oTg.CreatePoint2d(basePt.X, basePt.Y),
            oTg.CreatePoint2d(oVec.X, oVec.Y))
            oSl1.Construction = True
            Call oSk.GeometricConstraints.AddGround(oSl1)
            'Dim oAxisLine As Line
            'Set oAxisLine = ThisApplication.TransientGeometry.CreateLine _
            '(oCylinder.BasePoint, oCylinder.AxisVector.AsVector)
            ''
            Dim oSl2 As SketchLine = oSk.SketchLines.AddByTwoPoints(
                            oTg.CreatePoint2d(0, 0),
            oTg.CreatePoint2d(0, ancho1 / 2))
            oSl2.Construction = True
            ''
            Dim oSp As SketchPoint = oSk.SketchPoints.Add(oTg.CreatePoint2d(0, 0))
            ''
            Call oSk.GeometricConstraints.AddMidpoint(oSp, oSl1)
            Call oSk.GeometricConstraints.AddCoincident(oSl2.StartSketchPoint, oSl1.EndSketchPoint)
            Call oSk.GeometricConstraints.AddPerpendicular(oSl1, oSl2)
            ''
            Dim oSR As SketchEntitiesEnumerator = oSk.SketchLines.AddAsThreePointRectangle(
            oSp, oSl1.EndSketchPoint, oSl2.EndSketchPoint.Geometry)
        End If
    End Sub
    Public Sub SurfaceCilindricalCreaContorno2D(oFace As Face, ByRef oSk As PlanarSketch)
        ' Salimos si no ex Cylinder o Cone
        If TypeOf oFace.Geometry IsNot Cylinder And TypeOf oFace IsNot Cone Then Exit Sub
        ''
        Dim oVecUni As UnitVector = Nothing
        Dim oVec As Vector = Nothing
        Dim largo As Double = 0
        Dim ancho1 As Double = 0
        Dim ancho2 As Double = 0
        Dim ancho As Double = 0
        Dim basePt As Point = Nothing
        ''
        If TypeOf oFace.Geometry Is Cylinder Then
            Dim oCyl As Cylinder = oFace.Geometry
            oVecUni = oCyl.AxisVector
            oVec = oVecUni.AsVector
            basePt = oCyl.BasePoint
            largo = oVec.Length
            ancho1 = oCyl.Radius * 2
            ancho2 = ancho1
        ElseIf TypeOf oFace.Geometry Is Cone Then
            Dim oCon As Cone = oFace.Geometry
            oVecUni = oCon.AxisVector
            oVec = oVecUni.AsVector
            basePt = oCon.BasePoint
            largo = oVec.Length
            ancho1 = oCon.Radius * 2
            ancho2 = ancho1
        End If
        ''
        If oVec IsNot Nothing Then
            Dim oSl1 As SketchLine = oSk.SketchLines.AddByTwoPoints(
            oTg.CreatePoint2d(basePt.X, basePt.Y),
            oTg.CreatePoint2d(oVec.X, oVec.Y))
            oSl1.Construction = True
            Call oSk.GeometricConstraints.AddGround(oSl1)
            'Dim oAxisLine As Line
            'Set oAxisLine = ThisApplication.TransientGeometry.CreateLine _
            '(oCylinder.BasePoint, oCylinder.AxisVector.AsVector)
            ''
            Dim oSl2 As SketchLine = oSk.SketchLines.AddByTwoPoints(
                            oSl1.EndSketchPoint.Geometry,
            oTg.CreatePoint2d(oSl1.EndSketchPoint.Geometry.X, oSl1.EndSketchPoint.Geometry.Y + (ancho1 / 2)))
            oSl2.Construction = True
            ''
            Dim oSp As SketchPoint = oSk.SketchPoints.Add(oTg.CreatePoint2d(0, 0))
            ''
            Call oSk.GeometricConstraints.AddMidpoint(oSp, oSl1)
            'Call oSk.GeometricConstraints.AddCoincident(oSl2.StartSketchPoint, oSl1.EndSketchPoint)
            Call oSk.GeometricConstraints.AddPerpendicular(oSl1, oSl2)
            ''
            Dim oSR As SketchEntitiesEnumerator = oSk.SketchLines.AddAsThreePointRectangle(
            oSp, oSl1.EndSketchPoint, oSl2.EndSketchPoint.Geometry)
        End If
    End Sub
    ''
    Public Sub SurfaceCilindricalCreaContorno2DWorkPoints(oWp1 As WorkPoint, oWp2 As WorkPoint, ByRef oCd As PartComponentDefinition, ByRef oSk As PlanarSketch)
        ' Salimos si alguno de los valores de entrada es Nothing
        If oWp1 Is Nothing Or oWp2 Is Nothing Or oSk Is Nothing Then Exit Sub
        '' Coger parámetro ancho
        Dim ancho As Double = 10
        Try
            ancho = oCd.Parameters.UserParameters.Item("Width").Value
        Catch ex As Exception
            '' No debería dar error si ya exite.
        End Try
        '
        Dim oSl1 As SketchLine = oSk.SketchLines.AddByTwoPoints(
        oTg.CreatePoint2d(oWp1.Point.X, oWp1.Point.Y),
        oTg.CreatePoint2d(oWp2.Point.X, oWp2.Point.Y))
        oSl1.Construction = True
        Call oSk.GeometricConstraints.AddGround(oSl1)
        ''
        Dim oSl2 As SketchLine = oSk.SketchLines.AddByTwoPoints(
                            oSl1.EndSketchPoint.Geometry,
            oTg.CreatePoint2d(oSl1.EndSketchPoint.Geometry.X, oSl1.EndSketchPoint.Geometry.Y + (ancho / 2)))
        oSl2.Construction = True
        ''
        Dim oSp As SketchPoint = oSk.SketchPoints.Add(oTg.CreatePoint2d(0, 0))
        ''
        Call oSk.GeometricConstraints.AddMidpoint(oSp, oSl1)
        'Call oSk.GeometricConstraints.AddCoincident(oSl2.StartSketchPoint, oSl1.EndSketchPoint)
        Call oSk.GeometricConstraints.AddPerpendicular(oSl1, oSl2)
        ''
        Dim oSR As SketchEntitiesEnumerator = oSk.SketchLines.AddAsThreePointRectangle(
            oSp,
            oSl1.EndSketchPoint.Geometry,
            oSl2.EndSketchPoint.Geometry)
    End Sub
    ''
    Public Function SurfaceCilindricalCreaContorno2DSketchLine3D(oS3D As Sketch3D, ByRef oCd As PartComponentDefinition,
                                                        Optional nombreSk As String = "") As Boolean
        Dim resultado As Boolean = False
        ' Salimos si alguno de los valores de entrada es Nothing
        If oS3D Is Nothing Or oCd Is Nothing Then
            Return False
            Exit Function
        End If
        Dim oSk As PlanarSketch = Nothing
        If nombreSk <> "" Then
            Try
                ' Coger el PlanarSketch con el nombre "nombreSk"
                oSk = oCd.Sketches.Item(nombreSk)
            Catch ex As Exception
                '' No existe "nombreSk"
                Try
                    ' Coger el primer PlanarSketch
                    oSk = oCd.Sketches.Item(1)
                Catch ex1 As Exception
                    ' No hay ningún PlanarSketch. Crear uno en el plano XY (3)
                    oSk = oCd.Sketches.Add(oCd.WorkPlanes.Item(3))
                End Try
            End Try
        End If
        ' Coger parámetro ancho
        Dim ancho As Double = 10
        Try
            ancho = oCd.Parameters.UserParameters.Item("Width").Value
        Catch ex As Exception
            '' No debería dar error si ya exite.
        End Try
        ''
        Try
            ' Coger la linea 3D que habremos creado como eje
            Dim oSl3D As SketchLine3D = oS3D.SketchLines3D.Item(1)
            Dim oSl1 As SketchLine = oSk.SketchLines.AddByTwoPoints(
            oTg.CreatePoint2d(oSl3D.StartSketchPoint.Geometry.X, oSl3D.StartSketchPoint.Geometry.Y),
            oTg.CreatePoint2d(oSl3D.EndSketchPoint.Geometry.X, oSl3D.EndSketchPoint.Geometry.Y))
            oSl1.Construction = True
            Call oSk.GeometricConstraints.AddGround(oSl1)
            ''
            Dim oSl2 As SketchLine = oSk.SketchLines.AddByTwoPoints(
                                oSl1.EndSketchPoint.Geometry,
                oTg.CreatePoint2d(oSl1.EndSketchPoint.Geometry.X, oSl1.EndSketchPoint.Geometry.Y + (ancho / 2)))
            oSl2.Construction = True
            ''
            Dim oSp As SketchPoint = oSk.SketchPoints.Add(oTg.CreatePoint2d(0, 0))
            ''
            Call oSk.GeometricConstraints.AddMidpoint(oSp, oSl1)
            'Call oSk.GeometricConstraints.AddCoincident(oSl2.StartSketchPoint, oSl1.EndSketchPoint)
            Call oSk.GeometricConstraints.AddPerpendicular(oSl1, oSl2)
            ''
            Dim oSR As SketchEntitiesEnumerator = oSk.SketchLines.AddAsThreePointCenteredRectangle(
                oSp.Geometry,
                oSl1.EndSketchPoint.Geometry,
                oSl2.EndSketchPoint.Geometry)
            ''
            oSk.UpdateProfiles()
            ''
            If oSk.HealthStatus = HealthStatusEnum.kInErrorHealth Or oSk.HealthStatus = HealthStatusEnum.kInconsistentHealth Then
                'oSk.Delete()
                resultado = False
            Else
                resultado = True
            End If
        Catch ex As Exception
            'oSk.Delete()
            resultado = False
        End Try
        ''
        '' Borrar Sketch3D. Ya no nos hace falta.
        Return resultado
    End Function
    ''
    Public Function FacePlaneDireccionHaciaAbajo(oFace As FaceProxy, oAsmCompDef As AssemblyComponentDefinition) As Boolean
        '' Si queFace.Geometry no es Plane, salimos.
        If TypeOf oFace.Geometry IsNot Plane Then
            Return False
            Exit Function
        End If
        ''
        Dim oPlane As Plane = oFace.Geometry
        Dim Params(1) As Double
        Dim Points(2) As Double
        Dim Normals(2) As Double
        'If Faces is planar, then the Normal wil be the same all over the face
        Params(0) = 0 : Params(1) = 0
        oFace.Evaluator.GetPointAtParam(Params, Points)
        oFace.Evaluator.GetNormalAtPoint(Points, Normals)
        ''
        Dim oUnitNormal As UnitVector
        oUnitNormal = oTg.CreateUnitVector(Normals(0), Normals(1), Normals(2))
        ''
        Dim angulo As Double = oAsmCompDef.WorkAxes(3).Line.Direction.AngleTo(oUnitNormal)
        '' Comparar la Z para ver si está hacia arriba o hacia abajo.
        '' Angulo tiene que ser menor de 90º (Las perpendiculares están a 90º)
        If Points(2) > Normals(2) And angulo < 90 Then
            'MsgBox("Hacia abajo" & vbCrLf & "Angulo = " & RadGra(angulo))
            Return True
        Else
            Return False
        End If
        'If Points(2) < Normals(2) And angulo < 90 Then
        ' MsgBox("Hacia arriba" & vbCrLf & "Angulo = " & RadGra(angulo))
        'ElseIf Points(2) > Normals(2) Then
        '    MsgBox("Hacia abajo" & vbCrLf & "Angulo = " & RadGra(angulo))
        'ElseIf Points(2) = Normals(2) Then
        '    MsgBox("Perpendicular a XY" & vbCrLf & "Angulo = " & RadGra(angulo))
        'End If
    End Function
    ''
    Public Function FacePlaneDameAnguloSobreXY(oFace As FaceProxy, oAsmCompDef As AssemblyComponentDefinition) As Double
        '' Si queFace.Geometry no es Plane, salimos.
        If TypeOf oFace.Geometry IsNot Plane Then
            Return 0
            Exit Function
        End If
        ''
        Dim oPlane As Plane = oFace.Geometry
        Dim Params(1) As Double
        Dim Points(2) As Double
        Dim Normals(2) As Double
        'If Faces is planar, then the Normal wil be the same all over the face
        Params(0) = 0 : Params(1) = 0
        oFace.Evaluator.GetPointAtParam(Params, Points)
        oFace.Evaluator.GetNormalAtPoint(Points, Normals)
        ''
        Dim oUnitNormal As UnitVector
        oUnitNormal = oTg.CreateUnitVector(Normals(0), Normals(1), Normals(2))
        ''
        Dim angulo As Double = oAsmCompDef.WorkAxes(3).Line.Direction.AngleTo(oUnitNormal)
        Return angulo
    End Function
    ''' <summary>
    ''' Copia SurfaceBody de una pieza en otra (Ambas en el mismo ensamblaje)
    ''' y devuelve el objeto NonParametricBaseFeature creado
    ''' </summary>
    ''' <param name="oAssemblyDoc">Objeto AssemblyDocument</param>
    ''' <param name="PartOccu1">ComponentOccurrence 1 (Pieza Origen)</param>
    ''' <param name="PartOccu2">ComponentOccurrence 2 (Pieza Destino)</param>
    ''' <param name="asociativo">Crear en PartOccu2 como asociativo o no</param>
    ''' <returns></returns>
    Public Function SurfaceBodyCopia(ByRef oAssemblyDoc As AssemblyDocument,
                                 PartOccu1 As ComponentOccurrence,
                                ByRef PartOccu2 As ComponentOccurrence,
                                Optional asociativo As Boolean = False) As NonParametricBaseFeature
        ''
        Dim oPartDef1 As PartComponentDefinition = PartOccu1.Definition
        Dim oPartDef2 As PartComponentDefinition = PartOccu2.Definition
        Dim oBaseFeature1 As NonParametricBaseFeature = Nothing

        ' Get the source solid body from the first part.
        Dim oSourceBody As SurfaceBody = oPartDef1.SurfaceBodies.Item(1)
        Dim oSurfaceBodyProxy As Object = Nothing    ' SurfaceBodyProxy = Nothing
        Call PartOccu1.CreateGeometryProxy(oSourceBody, oSurfaceBodyProxy)

        ' Create an associative surface base feature in the second part.
        Dim oFeatureDef1 As NonParametricBaseFeatureDefinition
        oFeatureDef1 = oPartDef2.Features.NonParametricBaseFeatures.CreateDefinition

        Dim oCollection As ObjectCollection = oTo.CreateObjectCollection
        oCollection.Add(CType(oSurfaceBodyProxy, SurfaceBodyProxy))
        ''
        If asociativo = True Then
            ' Create a associative solid base feature in the second part.
            oFeatureDef1.BRepEntities = oCollection
            oFeatureDef1.OutputType = BaseFeatureOutputTypeEnum.kSurfaceOutputType '' Superficie
            'oFeatureDef1.OutputType = BaseFeatureOutputTypeEnum.kSolidOutputType   '' Solido
            oFeatureDef1.TargetOccurrence = PartOccu2
            oFeatureDef1.IsAssociative = True

            oBaseFeature1 = oPartDef2.Features.NonParametricBaseFeatures.AddByDefinition(oFeatureDef1)
        Else
            ' Create a non-associative solid base feature in the second part.
            oFeatureDef1.BRepEntities = oCollection
            oFeatureDef1.OutputType = BaseFeatureOutputTypeEnum.kSurfaceOutputType  '' Superficie
            'oFeatureDef1.OutputType = BaseFeatureOutputTypeEnum.kSolidOutputType   '' Solido
            oFeatureDef1.TargetOccurrence = PartOccu2
            oFeatureDef1.IsAssociative = False

            oBaseFeature1 = oPartDef2.Features.NonParametricBaseFeatures.AddByDefinition(oFeatureDef1)
        End If
        ''
        oAssemblyDoc.Update()
        ''
        Return oBaseFeature1
    End Function
End Class