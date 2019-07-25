Imports Inventor
Public Class clsSelectMulti1
    Private objInteraction As InteractionEvents
    Private sle As SelectEvents
    Private uiecc As UserInputEvents
    Private curSelection As Boolean = False
    Dim objSelectSet As ObjectCollection = Nothing
    Private appInventor As Application
    Private filtro As SelectionFilterEnum
    Private mensaje As String
    '
    Public Sub New(AppObj As Inventor.Application, fil As SelectionFilterEnum, men As String)
        appInventor = AppObj
        objSelectSet = appInventor.TransientObject.CreateObjectCollection()
        filtro = fil
        mensaje = men
        StartGetObjects()
    End Sub
    '
    Public Sub StartGetObjects()
        sle = Nothing
        objInteraction = Nothing
        objInteraction = appInventor.CommandManager.CreateInteractionEvents()
        sle = objInteraction.SelectEvents
        sle.ResetSelections()
        sle.ClearSelectionFilter()
        sle.SingleSelectEnabled = False
        sle.AddSelectionFilter(filtro)  ' SelectionFilterEnum.kAssemblyOccurrenceFilter)
        '
        objInteraction.StatusBarText = mensaje  ' "Select Components"
        curSelection = True
        AddHandler sle.OnSelect, AddressOf sle_OnSelect
        AddHandler sle.OnUnSelect, AddressOf sle_OnUnSelect
        'sle.OnSelect += New SelectEventsSink_OnSelectEventHandler(sle_OnSelect)
        'sle.OnUnSelect += New SelectEventsSink_OnUnSelectEventHandler(sle_OnUnSelect)
        objInteraction.Start()
    End Sub


    Private Sub sle_OnUnSelect(UnSelectedEntities As ObjectsEnumerator, SelectionDevice As SelectionDeviceEnum, ModelPosition As Inventor.Point, ViewPosition As Point2d, View As Inventor.View)
        If curSelection Then
            'Remove Object from Collection
            If UnSelectedEntities.Count > 0 Then
                For Each obj As Object In UnSelectedEntities
                    objSelectSet.Remove(obj)
                Next
            End If
        End If
    End Sub

    Private Sub sle_OnSelect(JustSelectedEntities As ObjectsEnumerator, SelectionDevice As SelectionDeviceEnum, ModelPosition As Inventor.Point, ViewPosition As Point2d, View As Inventor.View)
        If curSelection Then
            'Add Objects to Collection
            If JustSelectedEntities.Count > 0 Then
                For Each obj As Object In JustSelectedEntities
                    objSelectSet.Add(obj)
                Next
            End If
        End If
    End Sub
End Class
