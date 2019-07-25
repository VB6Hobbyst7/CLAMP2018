Imports Inventor
Public Class clsSelectMulti
    ''
    '' # Como utilizar esta clase para selecionar una Cara (face)
    ''
    '' # Declare a variable and create a new instance of the select class.
    'Dim oSelect As New clsSelect(m_inApp)          ' m_inApp es el Objecto Application de Inventor.
    '' # Call the Pick method of the clsSelect object and set
    '' # the filter to pick any face.
    'Dim oFace As Face
    'oFace = oSelect.Pick(SelectionFilterEnum.kPartFaceFilter,mensaje,True/False para singleselection)
    'oSelect = Nothing
    '' #
    '' #
    ''
    '*************************************************************
    ' The declarations and functions below need to be copied into
    ' a class module whose name is "clsSelect". The name can be
    ' changed but you'll need to change the declaration in the
    ' calling function "TestSelection" to use the new na
    ' Declare the event objects
    Private oInteractEvts As InteractionEvents
    Private oSelectEvents As SelectEvents
    ' Declare a flag that's used to determine when selection stops.
    Private bStillSelecting As Boolean
    Private ThisApplication As Application
    Private filter As SelectionFilterEnum
    Private mensaje As String = ""
    Private bolSingleSelection As Boolean = True
    Private oCol As Inventor.ObjectCollection = Nothing
    ''
    Public Sub New(AppObj As Inventor.Application)
        ThisApplication = AppObj
        oCol = ThisApplication.TransientObjects.CreateObjectCollection
    End Sub
    ''
    Public Function Pick(fil As SelectionFilterEnum, Optional men As String = "", Optional bolSingleSel As Boolean = True) As Object
        Dim oReturn As Object = Nothing
        ' Initialize flag.
        bStillSelecting = True
        filter = fil
        mensaje = men
        bolSingleSelection = bolSingleSel
        ' Create an InteractionEvents object.
        oInteractEvts = ThisApplication.CommandManager.CreateInteractionEvents
        AddHandler oInteractEvts.OnTerminate, AddressOf oInteractEvts_OnTerminate
        ' Ensure interaction is enabled.
        oInteractEvts.InteractionDisabled = False
        ' Set a reference to the select events.
        oSelectEvents = oInteractEvts.SelectEvents
        oSelectEvents.SingleSelectEnabled = bolSingleSelection
        oSelectEvents.ResetSelections()
        oSelectEvents.ClearSelectionFilter()
        oSelectEvents.ClearWindowSelectionFilter()
        AddHandler oSelectEvents.OnPreSelect, AddressOf oSelectEvents_OnPreSelect
        AddHandler oSelectEvents.OnSelect, AddressOf oSelectEvents_OnSelect
        ' Set the filter using the value passed in.
        oSelectEvents.AddSelectionFilter(filter)
        ' Start the InteractionEvents object.
        oInteractEvts.Start()
        ' Loop until a selection is made.
        Do While bStillSelecting
            If mensaje <> "" Then ThisApplication.StatusBarText = mensaje
            ThisApplication.UserInterfaceManager.DoEvents()
        Loop
        ' Devolver sólo una entidad (bolSingleSelection = True)
        ' O devolver todas las entidades (bolSingleSelection = False)
        Dim oSelectedEnts As ObjectsEnumerator
        oSelectedEnts = oSelectEvents.SelectedEntities
        If bolSingleSelection = True And oSelectedEnts.Count > 0 Then
            oReturn = oSelectedEnts.Item(1)
        ElseIf bolSingleSelection = False And oCol.Count > 0 Then
            oReturn = oCol
        Else
            oReturn = Nothing
        End If
        ' Stop the InteractionEvents object.
        oInteractEvts.Stop()
        ' Clean up.
        oSelectEvents = Nothing
        oInteractEvts = Nothing
        ThisApplication = Nothing
        Return oReturn
    End Function
    ''
    Private Sub oInteractEvts_OnTerminate()
        ' Set the flag to indicate we're done.
        bStillSelecting = False
    End Sub
    ''
    Private Sub oSelectEvents_OnPreSelect(ByRef PreSelectEntity As Object, ByRef DoHighlight As Boolean, ByRef MorePreSelectEntities As Inventor.ObjectCollection, SelectionDevice As Inventor.SelectionDeviceEnum, ModelPosition As Inventor.Point, ViewPosition As Inventor.Point2d, View As Inventor.View)
        DoHighlight = True
    End Sub
    ''
    Private Sub oSelectEvents_OnSelect(ByVal JustSelectedEntities As ObjectsEnumerator, ByVal SelectionDevice As SelectionDeviceEnum, ByVal ModelPosition As Point, ByVal ViewPosition As Point2d, ByVal View As View)
        ' Set the flag to indicate we're done.
        If bolSingleSelection Then
            bStillSelecting = False
        ElseIf bolSingleSelection = False Then
            'oCol.Clear()
            For Each oEnt As Object In JustSelectedEntities
                Dim existe As Boolean = False
                For Each oObj As Object In oCol
                    If oObj.Equals(oEnt) = True Then
                        existe = True
                        Exit For
                    End If
                Next
                '
                If existe = False Then
                    oCol.Add(oEnt)
                End If
            Next
        End If
    End Sub
End Class

