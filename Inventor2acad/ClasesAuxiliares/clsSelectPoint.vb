Imports Inventor
Imports System.Windows.Forms
Imports System.Runtime.InteropServices
Imports System.Configuration
Imports Microsoft.Win32
'
' * Nos da el punto pulsado en la pantalla.

'Then use this in whichever sub you want to call ClsSelectPoint  from
'Dim oGetPoint As New clsSelectPoint
'Dim oCP As Point2d = oGetPoint.GetPoint
Public Class clsSelectPoint

    Public oModelX As Double
    Public oModelY As Double
    Public oModelZ As Double

    ' Declare the event objects
    Private WithEvents oInteraction As Inventor.InteractionEvents
    Public WithEvents oMouseEvents As Inventor.MouseEvents
    Public Event MouseClick As MouseEventHandler

    ' Declare a Flag that's used to determine when selection stops.
    Private bStillSelecting As Boolean
    Private oSelectedPoint As Inventor.Point
    Private InvApp As Inventor.Application
    ''
    Public Sub New(AppObj As Inventor.Application)
        InvApp = AppObj
    End Sub
    Public Function GetPoint(Optional Prompt As String = "Select Point") As Inventor.Point
        ' Initialize flag.
        bStillSelecting = True

        ' Create an InteractionEvents object.
        oInteraction = InvApp.CommandManager.CreateInteractionEvents

        ' Set a reference to the mouse events.
        oMouseEvents = oInteraction.MouseEvents

        ' Disable mouse move since we only need the click.
        oMouseEvents.MouseMoveEnabled = False
        oInteraction.SetCursor(CursorTypeEnum.kCursorBuiltInCrosshair)
        oInteraction.StatusBarText = Prompt
        ' The InteractionEvents object.
        oInteraction.Start()
        ' Loop until a selection is made.
        Do While bStillSelecting
            InvApp.UserInterfaceManager.DoEvents()
        Loop
        ' Set the return variable with the point.
        GetPoint = oSelectedPoint
        ' Stop the InteractionEvents object.
        oInteraction.Stop()
        ' Clean up.
        oInteraction.SetCursor(CursorTypeEnum.kCursorTypeDefault)
        oMouseEvents = Nothing
        oInteraction = Nothing
    End Function

    Private Sub oInteraction_OnTerminate()
        '    Private Sub oInteraction_OnTerminate()
        ' Set the flag to indicate we're done.
        bStillSelecting = False
    End Sub
    '
    Public oPointX, oPointY, oPointZ As Double
    Private Sub oMouseEvents_OnMouseClick(ByVal Button As MouseButtonEnum, ByVal ShiftKeys As ShiftStateEnum, ByVal ModelPosition As Inventor.Point, ByVal ViewPosition As Point2d, ByVal View As Inventor.View) Handles oMouseEvents.OnMouseClick

        ' These are in cm
        oPointX = ModelPosition.X
        oPointY = ModelPosition.Y
        oPointZ = ModelPosition.Z

        bStillSelecting = False
        'your code here
        '################
        'MsgBox("X: " & oPointX & vbCr & "Y: " & oPointY)
        '#################
    End Sub
End Class


