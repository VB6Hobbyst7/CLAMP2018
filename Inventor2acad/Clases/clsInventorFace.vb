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
    Public Function Get3dPointInFace(inFace As Boolean) As Inventor.Point
        Dim modelPoint As Inventor.Point = Nothing
        Dim oSelectedPointInFace As New clsGetPointInFace(oAppI, inFace)
        modelPoint = oSelectedPointInFace.GetPoint
        ''
        Return modelPoint
    End Function
    Public Function FaceEsExteriorCylinderCone(oFace As Face) As Boolean
        Dim resultado As Boolean = True
        '' Evaluaremos si Face es CylinderSurface o ConeSurface
        If oFace.SurfaceType <> SurfaceTypeEnum.kCylinderSurface And oFace.SurfaceType <> SurfaceTypeEnum.kConeSurface Then
            Return False
            Exit Function
        End If
        ''
        Dim oCylinder As Object = Nothing
        oCylinder = oFace.Geometry
        'Dim oCylinder As Inventor.Cylinder = Nothing
        'Dim oCone As Inventor.Cone = Nothing
        'If oFace.SurfaceType = SurfaceTypeEnum.kCylinderSurface Then
        '    oCylinder = CType(oFace.Geometry, Inventor.Cylinder)
        'ElseIf oFace.SurfaceType = SurfaceTypeEnum.kConeSurface Then
        '    oCone = CType(oFace.Geometry, Inventor.Cone)
        'End If

        Dim params(1) As Double
        params(0) = 0.5
        params(1) = 0.5

        ' Get point on surface at param .5,.5
        Dim points(2) As Double
        Call oFace.Evaluator.GetPointAtParam(params, points)

        ' Create point object
        Dim oPoint As Point
        oPoint = oTg.CreatePoint(points(0), points(1), points(2))

        ' Get normal at this point
        Dim normals(2) As Double
        Call oFace.Evaluator.GetNormal(params, normals)

        ' Create normal vector object
        Dim oNormal As Vector
        oNormal = oTg.CreateVector(normals(0), normals(1), normals(2))

        ' Scale vector by radius of the cylinder
        'oNormal.ScaleBy(IIf(oCylinder IsNot Nothing, oCylinder.Radius, oCone.Radius))
        ' Cylinder y Cone tiene propiedad Radius
        oNormal.ScaleBy(oCylinder.Radius)
        'If oFace.SurfaceType = SurfaceTypeEnum.kPlaneSurface Then
        '    oNormal.ScaleBy(2)
        'ElseIf oFace.SurfaceType = SurfaceTypeEnum.kCylinderSurface Then
        '    oNormal.ScaleBy(CType(oCylinder, Cylinder).Radius)
        'ElseIf oFace.SurfaceType = SurfaceTypeEnum.kConeSurface Then
        '    oNormal.ScaleBy(CType(oCylinder, Cone).Radius)
        'End If

        ' Find the sampler point on the normal by adding the
        ' scaled normal vector to the point at .5,.5 param.
        Dim oSamplePoint As Point
        oSamplePoint = oPoint

        oSamplePoint.TranslateBy(oNormal)

        ' Check if the sample point lies on the cylinder axis.
        ' If it does, we have a hollow face.

        ' Create a line describing the cylinder axis
        Dim oAxisLine As Line = Nothing
        'oAxisLine = oTg.CreateLine(
        '    IIf(oCylinder IsNot Nothing, oCylinder.BasePoint, oCone.BasePoint),
        '    IIf(oCylinder IsNot Nothing, oCylinder.AxisVector.AsVector, oCone.AxisVector.AsVector))
        oAxisLine = oTg.CreateLine(oCylinder.BasePoint, oCylinder.AxisVector.AsVector)

        'Create a line parallel to the axis passing thru the sample point.
        Dim oSampleLine As Line
        'oSampleLine = oTg.CreateLine(oSamplePoint,
        '                             IIf(oCylinder IsNot Nothing, oCylinder.AxisVector.AsVector, oCone.AxisVector.AsVector))
        oSampleLine = oTg.CreateLine(oSamplePoint, oCylinder.AxisVector.AsVector)

        If oSampleLine.IsColinearTo(oAxisLine) Then
            resultado = False
        Else
            resultado = True
        End If
        ''
        Return resultado
    End Function
    Public Function FaceDameUnitVector(oFace As Face, queDato As IEnum.FaceData) As UnitVector
        Dim resultado As UnitVector = Nothing
        '
        Dim facePoint As Point = oFace.PointOnFace
        Dim surfEval As SurfaceEvaluator = oFace.Evaluator
        Dim points(2) As Double
        points(0) = facePoint.X
        points(1) = facePoint.Y
        points(2) = facePoint.Z
        Dim guessparams As Double() = New Double(1) {0, 0}
        Dim maxDeviations As Double() = New Double() {}
        Dim params As Double() = New Double() {}
        Dim solutionNatures As SolutionNatureEnum() = New SolutionNatureEnum() {}
        '
        Call surfEval.GetParamAtPoint(points, guessparams, maxDeviations, params, solutionNatures)
        '
        ' Calcular la Normal
        Dim normal(2) As Double
        Call surfEval.GetNormal(params, normal)
        Dim oNormal As UnitVector = oTg.CreateUnitVector(normal(0), normal(1), normal(2))
        '
        ' Calcular xDir
        Dim uTangents As Double() = New Double() {}
        Dim vTangents As Double() = New Double() {}
        Call surfEval.GetTangents(params, uTangents, vTangents)
        Dim xDir As UnitVector = oTg.CreateUnitVector(uTangents(0), uTangents(1), uTangents(2))
        '
        ' Calcular yDir
        Dim yDir As UnitVector = oNormal.CrossProduct(xDir)
        '
        ' Create a transform to position the text on the mid face.
        Dim transform As Matrix = oTg.CreateMatrix
        Call transform.SetCoordinateSystem(facePoint, xDir.AsVector, yDir.AsVector, oNormal.AsVector)
        Dim ori As Point = Nothing
        Dim xVec As Vector = Nothing
        Dim yVec As Vector = Nothing
        Dim zVec As Vector = Nothing
        transform.GetCoordinateSystem(ori, xVec, yVec, zVec)

        Select Case queDato
            Case IEnum.FaceData.Normal, IEnum.FaceData.DireccionZ
                resultado = oNormal
            Case IEnum.FaceData.DireccionX
                resultado = xDir
            Case IEnum.FaceData.DireccionY
                resultado = yDir
            Case IEnum.FaceData.DireccionZMedio
                resultado = zVec
        End Select
        ''
        Return resultado
    End Function
    ' Utility function that given a face returns a normal.  This is only useful
    ' for planar faces, since they have a consistent normal anywhere on the face.
    Private Function FaceDameNormal(InputFace As Face) As UnitVector
        If TypeOf InputFace.Geometry IsNot Plane Then
            Return Nothing
            Exit Function
        End If
        '
        Dim eval As SurfaceEvaluator = InputFace.Evaluator
        ' Get the center of the parametric range.
        Dim center(1) As Double
        center(0) = (eval.ParamRangeRect.MinPoint.X + eval.ParamRangeRect.MaxPoint.X) / 2
        center(1) = (eval.ParamRangeRect.MinPoint.Y + eval.ParamRangeRect.MaxPoint.Y) / 2

        ' Calculate the normal.
        Dim normal(2) As Double
        Call eval.GetNormal(center, normal)

        ' Create a unit vector to pass back the result.
        Return oAppI.TransientGeometry.CreateUnitVector(normal(0), normal(1), normal(2))
    End Function
    Public Function FaceDamePuntoMedio(oFace As Face) As Point
        Dim resultado As Point = Nothing
        '
        Dim facePoint As Point = oFace.PointOnFace
        Dim surfEval As SurfaceEvaluator = oFace.Evaluator
        Dim points(2) As Double
        points(0) = facePoint.X
        points(1) = facePoint.Y
        points(2) = facePoint.Z
        Dim guessparams As Double() = New Double(1) {0, 0}
        Dim maxDeviations As Double() = New Double() {}
        Dim params As Double() = New Double() {}
        Dim solutionNatures As SolutionNatureEnum() = New SolutionNatureEnum() {}
        '
        Call surfEval.GetParamAtPoint(points, guessparams, maxDeviations, params, solutionNatures)
        '
        ' Calcular la Normal
        Dim normal(2) As Double
        Call surfEval.GetNormal(params, normal)
        Dim oNormal As UnitVector = oTg.CreateUnitVector(normal(0), normal(1), normal(2))
        '
        ' Calcular xDir
        Dim uTangents As Double() = New Double() {}
        Dim vTangents As Double() = New Double() {}
        Call surfEval.GetTangents(params, uTangents, vTangents)
        Dim xDir As UnitVector = oTg.CreateUnitVector(uTangents(0), uTangents(1), uTangents(2))
        '
        ' Calcular yDir
        Dim yDir As UnitVector = oNormal.CrossProduct(xDir)
        '
        ' Create a transform to position the text on the mid face.
        Dim transform As Matrix = oTg.CreateMatrix
        Call transform.SetCoordinateSystem(facePoint, xDir.AsVector, yDir.AsVector, oNormal.AsVector)
        Dim ori As Point = Nothing
        Dim xVec As Vector = Nothing
        Dim yVec As Vector = Nothing
        Dim zVec As Vector = Nothing
        transform.GetCoordinateSystem(ori, xVec, yVec, zVec)
        '
        Return ori
    End Function
    Public Function FaceCylindrical_DameCirIniFin(oFace As FaceProxy) As Object()
        ' List(0) y List(1)
        Dim resultado(2) As Object ' As New List(Of Inventor.Circle)
        resultado(0) = Nothing : resultado(1) = Nothing
        '
        If TypeOf oFace.Geometry Is Cylinder = False Then
            Return resultado
            Exit Function
        End If
        '
        Dim oCylinder As Cylinder = oFace.Geometry
        '
        For Each oEdge As Edge In oFace.Edges
            Dim oCiTemp As Inventor.Circle = Nothing
            If TypeOf oEdge.Geometry Is Arc3d Then
                Dim oArc3D As Arc3d = oEdge.Geometry
                Dim centro(2) As Double : Dim axisvector(2) As Double : Dim refvector(2) As Double
                Dim radius, starangel, swepangle As Double
                oArc3D.GetArcData(centro, axisvector, refvector, radius, starangel, swepangle)
                Dim oArc3DNormal As UnitVector = oTg.CreateUnitVector(axisvector(0), axisvector(1), axisvector(2))
                oCiTemp = oTg.CreateCircle(oArc3D.Center, oArc3DNormal, oArc3D.Radius)
            ElseIf TypeOf oEdge.Geometry Is Circle Then
                oCiTemp = oEdge.Geometry
            Else
                Continue For
            End If
            ' Verificar si el circulo tiene el mismo radio con el cilindro
            If oCiTemp IsNot Nothing AndAlso oCiTemp.Radius <> oCylinder.Radius Then ' oCiTemp.Normal.IsParallelTo(oCylinder.AxisVector) = False Then
                Continue For
            End If
            ' Si el centro del circulo coincide con el inicio de cilindro va a resultado(0)
            If oCiTemp.Center.Equals(oCylinder.BasePoint) Then
                resultado(0) = oCiTemp
            Else
                resultado(1) = oCiTemp
            End If
        Next
        '
        Return resultado
    End Function
    Public Sub FaceCylindrical_RellenaCirIniFin(oFace As FaceProxy, ByRef oCi1 As Circle, ByRef oCi2 As Circle)
        If TypeOf oFace.Geometry Is Cylinder = False Then
            oCi1 = Nothing : oCi2 = Nothing
            Exit Sub
        End If
        '
        Dim oCylinder As Cylinder = oFace.Geometry
        '
        For Each oEdge As Edge In oFace.Edges
            Dim oCiTemp As Inventor.Circle = Nothing
            If TypeOf oEdge.Geometry Is Arc3d Then
                Dim oArc3D As Arc3d = oEdge.Geometry
                Dim centro(2) As Double : Dim axisvector(2) As Double : Dim refvector(2) As Double
                Dim radius, starangel, swepangle As Double
                oArc3D.GetArcData(centro, axisvector, refvector, radius, starangel, swepangle)
                Dim oArc3DNormal As UnitVector = oTg.CreateUnitVector(axisvector(0), axisvector(1), axisvector(2))
                oCiTemp = oTg.CreateCircle(oArc3D.Center, oArc3DNormal, oArc3D.Radius)
            ElseIf TypeOf oEdge.Geometry Is Circle Then
                oCiTemp = oEdge.Geometry
            Else
                Continue For
            End If
            ' Verificar si el circulo tiene el mismo radio con el cilindro
            If oCiTemp IsNot Nothing AndAlso oCylinder.AxisVector.IsParallelTo(oCiTemp.Normal, 0.5) = False Then ' oCiTemp.Normal.IsParallelTo(oCylinder.AxisVector) = False Then
                Continue For
            End If
            ' Si el centro del circulo coincide con el inicio de cilindro va a resultado(0)
            If oCi1 Is Nothing Then
                oCi1 = oCiTemp
            ElseIf oCi1 IsNot Nothing AndAlso oCi1.Center.X > oCiTemp.Center.X Then
                oCi2 = oCi1
                oCi1 = oCiTemp
            Else
                oCi2 = oCiTemp
            End If
        Next
    End Sub
    ' Damos por hecho que tendrá algún circulo dentro de la Cara
    Public Function FaceDameGrosorAgujeros(oFace As FaceProxy) As Double
        Dim resultado As Double = 0
        ' Recorrer sólo los EdgeLoop interiores de la cara. Cogeremos el que tenga el radio más grande
        'Dim oHig As HighlightSet = oApp.ActiveEditDocument.HighlightSets.Add
        'oHig.Color = oApp.TransientObjects.CreateColor(0, 0, 255)
        Dim oCd As AssemblyComponentDefinition = CType(oAppI.ActiveEditDocument, AssemblyDocument).ComponentDefinition
        Dim oEl As EdgeLoop = Nothing
        Dim oWp1 As WorkPoint = Nothing
        Dim oWp2 As WorkPoint = Nothing
        Dim oPt1 As Point = Nothing
        Dim oPt2 As Point = Nothing
        Dim oNormalF As UnitVector = FaceDameNormal(oFace)
        Dim oNormal As UnitVector = Nothing
        Dim esExterior As Boolean = False
        For Each oEl In oFace.EdgeLoops
            ' Si es exterior, saltar
            If oEl.IsOuterEdgeLoop Then Continue For
            '
            For Each oEdge As Edge In oEl.Edges
                If TypeOf oEdge.Geometry Is Circle Or TypeOf oEdge.Geometry Is Arc3d Then
                    For Each oFa As Face In oEdge.Faces
                        ' No tenemos en cuenta la cara enviada. Continuar
                        If oFa.Equals(oFace) Then Continue For
                        '
                        'Call oHig.AddItem(oFa)
                        oAppI.ActiveView.Update()
                        oNormal = oEdge.Geometry.Normal
                        Dim s As SurfaceEvaluator = oFa.Evaluator
                        Dim paramRange As Box2d = s.ParamRangeRect
                        ' Aquí sacamos la medida de esta Face
                        Dim altoFace As Double = paramRange.MaxPoint.X - paramRange.MinPoint.X
                        Dim anchoFace As Double = paramRange.MaxPoint.Y - paramRange.MinPoint.Y
                        'If altoFace > resultado Then resultado = altoFace
                        ' Crear WorkPoints al inicio y final
                        Dim pt(2) As Double
                        Dim params(1) As Double
                        params(0) = paramRange.MinPoint.X
                        params(1) = paramRange.MinPoint.Y
                        Call s.GetPointAtParam(params, pt)
                        oPt1 = oAppI.TransientGeometry.CreatePoint(pt(0), pt(1), pt(2))
                        'oWp1 = oCd.WorkPoints.AddFixed(oApp.TransientGeometry.CreatePoint(pt(0), pt(1), pt(2)))
                        '
                        Dim pt1(2) As Double
                        Dim params1(1) As Double
                        params1(0) = paramRange.MinPoint.X + altoFace
                        params1(1) = paramRange.MinPoint.Y
                        Call s.GetPointAtParam(params1, pt1)
                        oPt2 = oAppI.TransientGeometry.CreatePoint(pt1(0), pt1(1), pt1(2))
                        'oWp2 = oCd.WorkPoints.AddFixed(oApp.TransientGeometry.CreatePoint(pt1(0), pt1(1), pt1(2)))
                        esExterior = FaceEsExteriorCylinderCone(oFa)
                    Next
                    'resultado = oWp1.Point.DistanceTo(oWp2.Point)
                    resultado = oPt1.DistanceTo(oPt2)
                End If
                Exit For
            Next
        Next
        '
        'Dim oVector As Vector = oApp.TransientGeometry.CreateVector(oWp2.Point.X, oWp2.Point.Y, oWp2.Point.Z)
        'Dim oLine As Line = oApp.TransientGeometry.CreateLine(oWp1.Point, oVector)
        'Dim direccion As UnitVector = oLine.Direction
        'Dim angulo As Double = oLine.Direction.AngleTo(oNormalF)
        If esExterior Then
            Return resultado
        Else
            Return -resultado
        End If
        'If Math.Round(oNormalF.X, 2).Equals(Math.Round(oNormal.X, 2)) And
        '        Math.Round(oNormalF.Y, 2).Equals(Math.Round(oNormal.Y, 2)) And
        '        Math.Round(oNormalF.Z, 2).Equals(Math.Round(oNormal.Z, 2)) Then
        '    Return resultado
        'Else
        '    Return -resultado
        'End If
    End Function
End Class