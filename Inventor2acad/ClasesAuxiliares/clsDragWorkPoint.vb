Option Explicit On
Imports Inventor
'# Descripcion 
'This sample demonstrates the use Of the OnDrag Event To drag fixed work points When no command Is active.
'This sample only allows drags parallel To the X-Y plane.
'This sample Is dependent On events And VB only supports events within a Class Module.
'
'# Para utilizar esta clase en VB.NET (EN VBA cambar oApp por ThisApplication y Poner Set= para variables)
'Public oDragWorkPoint As clsDragWorkPoint
'Sub WorkPointDrag()
'    Set oDragWorkPoint = New clsDragWorkPoint
'    oDragWorkPoint.Initialize()
'End Sub

Public Class clsDragWorkPoint
    Private WithEvents oUserInputEvents As UserInputEvents
    Private oIE As InteractionEvents
    Private WithEvents oMouseEvents As MouseEvents
    Private oIntGraphics As InteractionGraphics
    Private oWP As WorkPoint
    Private ThisApplication As Inventor.Application

    Public Sub Initialize(oInventorApp As Inventor.Application)
        oUserInputEvents = oInventorApp.CommandManager.UserInputEvents
        Me.ThisApplication = oInventorApp
    End Sub

    Private Sub oUserInputEvents_OnDrag(ByVal DragState As Inventor.DragStateEnum, ByVal ShiftKeys As Inventor.ShiftStateEnum, ByVal ModelPosition As Inventor.Point, ByVal ViewPosition As Inventor.Point2d, ByVal View As Inventor.View, ByVal AdditionalInfo As Inventor.NameValueMap, HandlingCode As Inventor.HandlingCodeEnum)
        Dim oSS As SelectSet = ThisApplication.ActiveDocument.SelectSet
        If DragState = DragStateEnum.kDragStateDragHandlerSelection Then
            If oSS.Count = 1 And TypeOf oSS.Item(1) Is WorkPoint Then   'oSS.Item(1).Type = ObjectTypeEnum.kWorkPointObject Then
                oWP = oSS.Item(1)
                If oWP.DefinitionType = WorkPointDefinitionEnum.kFixedWorkPoint Then
                    HandlingCode = HandlingCodeEnum.kEventCanceled
                    oIE = ThisApplication.CommandManager.CreateInteractionEvents
                    oMouseEvents = oIE.MouseEvents
                    oMouseEvents.MouseMoveEnabled = True
                    oIntGraphics = oIE.InteractionGraphics
                    Call oIE.SetCursor(CursorTypeEnum.kCursorBuiltInCommonSketchDrag)
                    oIE.Start()
                End If
            End If
        End If
    End Sub

    Private Sub oMouseEvents_OnMouseMove(ByVal Button As MouseButtonEnum, ByVal ShiftKeys As ShiftStateEnum, ByVal ModelPosition As Point, ByVal ViewPosition As Point2d, ByVal View As View)
        Dim oSS As SelectSet = ThisApplication.ActiveDocument.SelectSet
        If oSS.Count = 1 And TypeOf oSS.Item(1) Is WorkPoint Then   'oSS.Item(1).Type = ObjectTypeEnum.kWorkPointObject Then
            Dim oWPDef As FixedWorkPointDef = oWP.Definition
            Dim oProjectedPoint As Inventor.Point = Nothing
            Call ProjectPoint(ModelPosition, oWPDef.Point, oProjectedPoint)
            ' Set a reference to the transient geometry object for user later.
            Dim oTransGeom As TransientGeometry = ThisApplication.TransientGeometry
            ' Create a graphics data set object.  This object contains all of the
            ' information used to define the graphics.
            Dim oDataSets As GraphicsDataSets = oIntGraphics.GraphicsDataSets
            If oDataSets.Count <> 0 Then
                oDataSets.Item(1).Delete()
            End If
            ' Create a coordinate set.
            Dim oCoordSet As GraphicsCoordinateSet = oDataSets.CreateCoordinateSet(1)
            ' Create an array that contains coordinates that define a set
            ' of outwardly spiraling points.
            Dim oPointCoords(2) As Double
            ' Define the X, Y, and Z components of the point.
            oPointCoords(0) = oProjectedPoint.X
            oPointCoords(1) = oProjectedPoint.Y
            oPointCoords(2) = oProjectedPoint.Z
            ' Assign the points into the coordinate set.
            Call oCoordSet.PutCoordinates(oPointCoords)
            ' Create the ClientGraphics object.
            Dim oClientGraphics As ClientGraphics = oIntGraphics.PreviewClientGraphics
            If oClientGraphics.Count <> 0 Then
                oClientGraphics.Item(1).Delete()
            End If
            ' Create a new graphics node within the client graphics objects.
            Dim oPtNode As GraphicsNode = oClientGraphics.AddNode(1)
            ' Create a PointGraphics object within the node.
            Dim oPtGraphics As PointGraphics = oPtNode.AddPointGraphics
            ' Assign the coordinate set to the line graphics.
            oPtGraphics.CoordinateSet = oCoordSet
            oPtGraphics.PointRenderStyle = PointRenderStyleEnum.kCrossPointStyle
            ThisApplication.ActiveView.Update()
        End If
    End Sub

    Private Sub oMouseEvents_OnMouseUp(ByVal Button As MouseButtonEnum, ByVal ShiftKeys As ShiftStateEnum, ByVal ModelPosition As Point, ByVal ViewPosition As Point2d, ByVal View As View)
        Dim oSS As SelectSet = ThisApplication.ActiveDocument.SelectSet
        If oSS.Count = 1 And oSS.Item(1).Type = ObjectTypeEnum.kWorkPointObject Then
            Dim oWPDef As FixedWorkPointDef = oWP.Definition
            Dim oProjectedPoint As Inventor.Point = Nothing
            Call ProjectPoint(ModelPosition, oWPDef.Point, oProjectedPoint)
            ' Reposition the fixed work point
            oWPDef.Point = oProjectedPoint
            ThisApplication.ActiveDocument.Update()
            oIE.Stop()
            '
            oWP = Nothing
        End If
    End Sub
    ' Project the ModelPosition to a plane parallel to the
    ' X-Y plane on which the work point currently is.
    Private Sub ProjectPoint(ByVal ModelPosition As Inventor.Point, ByVal WorkPointPosition As Inventor.Point, ProjectedPoint As Inventor.Point)
        ' Set a reference to the camera object
        Dim oCamera As Inventor.Camera = ThisApplication.ActiveView.Camera
        Dim oVec As Vector = oCamera.Eye.VectorTo(oCamera.Target)
        Dim oLine As Line = ThisApplication.TransientGeometry.CreateLine(ModelPosition, oVec)
        ' Create the z-axis vector
        Dim oZAxis As Vector = ThisApplication.TransientGeometry.CreateVector(0, 0, 1)
        ' Create a plane parallel to the X-Y plane
        Dim oWPPlane As Plane = ThisApplication.TransientGeometry.CreatePlane(WorkPointPosition, oZAxis)
        '
        ProjectedPoint = oWPPlane.IntersectWithLine(oLine)
    End Sub
End Class
