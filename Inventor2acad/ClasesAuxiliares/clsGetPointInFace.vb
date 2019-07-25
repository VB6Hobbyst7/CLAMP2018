Option Explicit On
''
Imports Inventor
Public Class clsGetPointInFace

    ' Declare the event objects
    Private WithEvents oInteractEvents As InteractionEvents = Nothing
    Private WithEvents oUserInputEvents As UserInputEvents
    Private WithEvents oMouseEvents As MouseEvents = Nothing
    Private WithEvents oSelect As SelectEvents
    ''
    Private queApp As Inventor.Application = Nothing

    ' Declare a flag that's used to determine when selection stops.
    Private bStillSelecting As Boolean
    Private modelPoint As Point
    Private inFace As Boolean = True

    Public Sub New(ByRef oInv As Inventor.Application, Optional bolinFace As Boolean = True)
        Me.queApp = oInv
        Me.inFace = bolinFace
    End Sub

    Public Function GetPoint() As Point
        ' Initialize flag.
        bStillSelecting = True

        ' Create an InteractionEvents object.
        oInteractEvents = queApp.CommandManager.CreateInteractionEvents
        ' Set a reference to the User Input Events.
        oUserInputEvents = queApp.CommandManager.UserInputEvents
        ' Set a reference to the mouse events.
        oMouseEvents = oInteractEvents.MouseEvents
        ''
        If Me.inFace = True Then
            oInteractEvents.StatusBarText = "Select Point in Face"
        Else
            oInteractEvents.StatusBarText = "Select Point in Screen"
        End If
        '' Select events = Yes / Mouse Events = false (True)
        oInteractEvents.SelectionActive = False
        ' Ensure interaction is enabled.
        oInteractEvents.InteractionDisabled = False

        oMouseEvents.PointInferenceEnabled = True
        oMouseEvents.MouseMoveEnabled = True

        ' Start the InteractionEvents object.
        oInteractEvents.Start()

        ' Loop until a (3D) point in the model is selected.
        Do While bStillSelecting = True
            System.Windows.Forms.Application.DoEvents()
        Loop

        ' Stop the InteractionEvents object.
        oInteractEvents.Stop()

        ' Clean up.
        oMouseEvents = Nothing
        oInteractEvents = Nothing

        GetPoint = modelPoint
    End Function

    Private Sub oInteractEvents_OnTerminate()
        ' Set the flag to indicate we're done.
        bStillSelecting = False
    End Sub

    Private Function MovePtToFace(pt As Point, v As View) As Point
        ' Get the view direction, i.e. the vector pointing
        ' from the Eye to the Target
        Dim e2t As Vector
        e2t = v.Camera.Eye.VectorTo(v.Camera.Target)

        ' The vector that will take the Model Point from the
        ' Target plane to the Screen plane is the opposite of e2t
        Dim m2s As Vector
        m2s = e2t.Copy
        m2s.ScaleBy(-1)
        Call pt.TranslateBy(m2s)

        Dim doc As PartDocument
        doc = v.Document

        ' Now we can shoot a ray from the Screen plane
        ' towards the model along the view direction to
        ' find the first object it hits and the intersection point
        Dim objects As ObjectsEnumerator = Nothing
        Dim pts As ObjectsEnumerator = Nothing
        Call doc.ComponentDefinition.FindUsingRay(
            pt, e2t.AsUnitVector(),
            0.001, objects, pts)

        If pts.Count > 0 Then
            MovePtToFace = pts(1)
        Else
            MovePtToFace = Nothing
        End If
    End Function

    Private Sub oMouseEvents_OnMouseClick(Button As MouseButtonEnum, ShiftKeys As ShiftStateEnum, ModelPosition As Point, ViewPosition As Point2d, View As View) Handles oMouseEvents.OnMouseClick
        bStillSelecting = False

        ' ModelPosition will be on the Target Plane
        ' which is a plane parallel to the screen's plane
        ' but instead of including the Camera.Eye position
        ' this includes the Camera.Target position
        ''
        If inFace = True Then
            modelPoint = MovePtToFace(ModelPosition, View)
        Else
            modelPoint = ModelPosition
        End If
    End Sub

    Private Sub oMouseEvents_OnMouseMove(Button As MouseButtonEnum, ShiftKeys As ShiftStateEnum, ModelPosition As Point, ViewPosition As Point2d, View As View) Handles oMouseEvents.OnMouseMove
        Dim newPos As Point
        Dim txtPre As String = ""
        ''
        If inFace = True Then
            newPos = MovePtToFace(ModelPosition, View)
            txtPre = "Select Point in Face : "
        Else
            newPos = ModelPosition
            txtPre = "Select Point in Screen : "
        End If

        If Not newPos Is Nothing Then ModelPosition = newPos

        queApp.StatusBarText = txtPre &
            ModelPosition.X & " : " &
            ModelPosition.Y & " : " &
            ModelPosition.Z
    End Sub
End Class
