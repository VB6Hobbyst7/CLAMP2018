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
Public Class clsGetPoint2D
    Private WithEvents m_interaction As InteractionEvents
    Private WithEvents m_mouse As MouseEvents
    Private m_position As Point2d
    Private m_button As MouseButtonEnum
    Private m_continue As Boolean
    Private ThisApplication As Inventor.Application
    'Public Sub TestGetDrawingPoint()
    '    Dim getPoint As New clsGetPoint
    '    Dim pnt As Point2d
    '    Do
    '    pnt = getPoint.GetDrawingPoint("Click the desired location", kLeftMouseButton)
    '   If Not pnt Is Nothing Then
    '       MsgBox "Click is at " & Format(pnt.x, "0.0000") & ", " & Format(pnt.y, "0.0000")
    '   End If
    '    Loop While Not pnt Is Nothing
    '    End Sub
    Public Sub New(AppObj As Inventor.Application)
        ThisApplication = AppObj
    End Sub
    Public Function GetDrawingPoint(Prompt As String, button As MouseButtonEnum) As Point2d
        m_position = Nothing
        m_button = button
        m_interaction = ThisApplication.CommandManager.CreateInteractionEvents
        m_mouse = m_interaction.MouseEvents
        m_interaction.StatusBarText = Prompt
        m_interaction.Start
        m_continue = True
        Do
            ThisApplication.UserInterfaceManager.DoEvents()
        Loop While m_continue
        m_interaction.Stop
        GetDrawingPoint = m_position
    End Function
    Private Sub m_mouse_OnMouseClick(ByVal button As MouseButtonEnum, ByVal ShiftKeys As ShiftStateEnum, ByVal ModelPosition As Point, ByVal ViewPosition As Point2d, ByVal View As Inventor.View)
        If button = m_button Then
            m_position = ThisApplication.TransientGeometry.CreatePoint2d(ModelPosition.X, ModelPosition.Y)
        End If
        m_continue = False
    End Sub

    Private Sub m_Key_KeyPress(KeyAscii As Integer)
        If KeyAscii = 27 Then

            m_continue = False
        End If
    End Sub
End Class
