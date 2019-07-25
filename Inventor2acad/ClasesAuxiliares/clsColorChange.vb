Imports Inventor
Public Class clsColorChange
    '*************************************************************
    ' The declarations and functions below need to be copied into
    ' a class module whose name is "clsColorChange".  The name
    ' can be changed but you'll need to change the declaration in
    ' the calling function "ChangeAppearanceMiniToolbarSample" to use the new name.
    Private WithEvents m_MiniToolbar As MiniToolbar
    Private WithEvents m_Colors As MiniToolbarComboBox
    Private WithEvents m_Filter As MiniToolbarDropdown
    Private WithEvents m_Preview As MiniToolbarCheckBox
    Private m_PreviewColor As MiniToolbarCheckBox

    Private WithEvents oInteractionEvents As InteractionEvents
    Private WithEvents m_SelectEvents As SelectEvents
    Private m_ChangeColorTransaction As Transaction

    Private m_Doc As PartDocument
    Private m_DefaultColor As Asset

    Private bIsinteractionStarted As Boolean

    Private bNeedTransaction As Boolean
    Private bStop As Boolean
    Private ThisApplication As Inventor.Application

    Public Sub Init(oDoc As PartDocument)
        m_Doc = oDoc
        m_DefaultColor = m_Doc.ActiveAppearance
        ThisApplication = oDoc.PrintManager.Application
        ' Create interaction events
        oInteractionEvents = ThisApplication.CommandManager.CreateInteractionEvents
        'oInteractionEvents.InteractionDisabled = False

        m_SelectEvents = oInteractionEvents.SelectEvents
        'm_SelectEvents.ClearSelectionFilter
        'm_SelectEvents.SingleSelectEnabled = False
        'm_SelectEvents.Enabled = True

        ' Create mini-tool bar for changing appearance
        m_MiniToolbar = oInteractionEvents.CreateMiniToolbar
        Call InitiateMiniToolbar()

        bStop = False
        m_ChangeColorTransaction = ThisApplication.TransactionManager.StartTransaction(m_Doc, "Change Appearance")

        Do
            ThisApplication.UserInterfaceManager.DoEvents
        Loop Until bStop

    End Sub

    Private Sub InitiateMiniToolbar()
        m_MiniToolbar.ShowOK = True
        m_MiniToolbar.ShowApply = True
        m_MiniToolbar.ShowCancel = True

        Dim oControls As MiniToolbarControls
        oControls = m_MiniToolbar.Controls
        oControls.Item("MTB_Options").Visible = False

        m_Filter = m_MiniToolbar.Controls.AddDropdown("Filter", False, True, True, False)
        Call m_Filter.AddItem("Part", "Part", "Filter_Part", False, False)
        Call m_Filter.AddItem("Feature", "Feature", "Filter_Feature", False, False)
        Call m_Filter.AddItem("Face", "Face", "Filter_Face", False, False)


        m_Colors = oControls.AddComboBox("Colors", True, True, 50)
        Call m_Colors.AddItem("Default", "Use default color", "Default", False)
        Call m_Colors.AddItem("Red", "Red", "Red", False)
        Call m_Colors.AddItem("Orange", "Orange", "Orange", False)
        Call m_Colors.AddItem("Yellow", "Yellow", "Yellow", False)
        Call m_Colors.AddItem("Green", "Green", "Green", False)
        Call m_Colors.AddItem("Blue", "Blue", "Blue", False)
        Call m_Colors.AddItem("Indigo", "Indigo", "Indigo", False)
        Call m_Colors.AddItem("Purple", "Purple", "Purple", False)

        oControls.AddNewLine

        ' Specify if preview the color when hover a color item
        m_PreviewColor = m_MiniToolbar.Controls.AddCheckBox("PreviewColor", "Hover color preview", "Whether preview color when hover on it", True)

        ' Position the mini-tool bar to the top-left.
        Dim oPosition As Point2d
        oPosition = ThisApplication.TransientGeometry.CreatePoint2d(0, 0)

        m_MiniToolbar.Visible = True
        m_MiniToolbar.Position = oPosition
    End Sub

    Private Sub m_Colors_OnItemHoverStart(ByVal ListItem As MiniToolbarListItem)
        ' Preview the color when hover on it.
        If m_PreviewColor.Checked Then
            Call ChangeColor(ListItem.Text)
        End If
    End Sub

    Private Sub m_Colors_OnSelect(ByVal ListItem As MiniToolbarListItem)
        ' Check if the selected color is already used for the part/objects
        If m_Filter.SelectedItem.Text = "Part" Then
            If m_Doc.ActiveAppearance.Name = ListItem.Text Then
                bNeedTransaction = False
            Else
                bNeedTransaction = True
            End If
        Else
            bNeedTransaction = True
        End If

        Call ChangeColor(ListItem.Text)

    End Sub
    ' Change filter for assigning color
    Private Sub m_Filter_OnSelect(ByVal ListItem As MiniToolbarListItem)
        If ThisApplication.TransactionManager.CurrentTransaction.DisplayName = "Change Appearance" Then
            ThisApplication.TransactionManager.CurrentTransaction.Abort
        End If

        m_ChangeColorTransaction = ThisApplication.TransactionManager.StartTransaction(m_Doc, "Change Appearance")

        Select Case ListItem.Text
            Case "Part"
                m_Doc.SelectSet.Clear
                m_SelectEvents.ResetSelections
                m_SelectEvents.ClearSelectionFilter
                m_SelectEvents.AddSelectionFilter(SelectionFilterEnum.kPartDefaultFilter)
                oInteractionEvents.SetCursor(CursorTypeEnum.kCursorTypeDefault)
            Case "Feature"
                m_Doc.SelectSet.Clear

                m_SelectEvents.ResetSelections
                m_SelectEvents.ClearSelectionFilter
                m_SelectEvents.AddSelectionFilter(SelectionFilterEnum.kPartFeatureFilter)

                If Not bIsinteractionStarted Then
                    oInteractionEvents.Start
                    bIsinteractionStarted = True
                End If
            Case "Face"
                m_Doc.SelectSet.Clear
                m_SelectEvents.ResetSelections
                m_SelectEvents.ClearSelectionFilter
                m_SelectEvents.AddSelectionFilter(SelectionFilterEnum.kPartFaceFilter)
                If Not bIsinteractionStarted Then
                    oInteractionEvents.Start
                    bIsinteractionStarted = True
                End If

        End Select
        m_Doc.Views(1).Update
        Call ChangeColor(ListItem.Text)

    End Sub

    Private Sub m_MiniToolbar_OnApply()

        If (m_Filter.SelectedItem.Text = "Feature" Or m_Filter.SelectedItem.Text = "Face") And (m_SelectEvents.SelectedEntities.Count = 0) Then
            m_ChangeColorTransaction.Abort
            m_Doc.Views(1).Update
            m_ChangeColorTransaction = ThisApplication.TransactionManager.StartTransaction(m_Doc, "Change Appearance")
            Exit Sub
        Else
            If bNeedTransaction Then ' Change color style
                Call ChangeColor(m_Colors.SelectedItem.Text)
                m_ChangeColorTransaction.End
            Else ' If no change to the color style
                m_ChangeColorTransaction.Abort
            End If

            ' Clear current selection for Feature and Face filter.
            If (m_Filter.SelectedItem.Text = "Feature" Or m_Filter.SelectedItem.Text = "Face") Then
                m_Doc.SelectSet.Clear
                m_SelectEvents.ResetSelections
            End If
        End If

        m_ChangeColorTransaction = ThisApplication.TransactionManager.StartTransaction(m_Doc, "Change Appearance")
    End Sub

    Private Sub m_MiniToolbar_OnCancel()
        bStop = True
        If ThisApplication.TransactionManager.CurrentTransaction Is m_ChangeColorTransaction Then
            m_ChangeColorTransaction.Abort
        End If
        m_SelectEvents.AddSelectionFilter(SelectionFilterEnum.kPartDefaultFilter)
        If bIsinteractionStarted Then oInteractionEvents.Stop
        m_Doc.Views(1).Update

    End Sub

    Private Sub m_MiniToolbar_OnOK()
        bStop = True
        If bNeedTransaction Then ' Change color
            Call ChangeColor(m_Colors.SelectedItem.Text)
            m_ChangeColorTransaction.End
        Else ' If no change to the color style
            m_ChangeColorTransaction.Abort
        End If
    End Sub

    Private Sub oInteractionEvents_OnTerminate()

        If ThisApplication.TransactionManager.CurrentTransaction Is m_ChangeColorTransaction Then
            m_ChangeColorTransaction.Abort
        End If
        If bIsinteractionStarted Then
            oInteractionEvents.Stop
        End If
        m_Doc.Views(1).Update
    End Sub

    Private Sub ChangeColor(sColor As String)
        Debug.Print("Passed in:" & sColor)
        If m_Filter.SelectedItem.Text = "Part" Then
            Select Case sColor
                Case "Default"
                    m_Doc.ActiveAppearance = m_DefaultColor
                Case "Red", "Orange", "Yellow", "Green", "Blue", "Indigo", "Purple"
                    m_Doc.ActiveAppearance = m_Doc.AppearanceAssets.Item(sColor)
            End Select
        ElseIf m_Filter.SelectedItem.Text = "Feature" Then
            If m_SelectEvents.SelectedEntities.Count Then
                Dim oFeature As PartFeature, oSelectedObj As Object

                For Each oSelectedObj In m_SelectEvents.SelectedEntities
                    If InStr(1, TypeName(oSelectedObj), "Feature") Then
                        oFeature = oSelectedObj

                        Select Case sColor 'm_Colors.SelectedItem.Text
                            Case "Default"
                                oFeature.Appearance = m_DefaultColor
                            Case "Red", "Orange", "Yellow", "Green", "Blue", "Indigo", "Purple"
                                oFeature.Appearance = m_Doc.AppearanceAssets.Item(sColor)
                        End Select
                    End If
                Next
            End If
        ElseIf m_Filter.SelectedItem.Text = "Face" Then
            If m_SelectEvents.SelectedEntities.Count Then
                Dim oFace As Face

                For Each oSelectedObj In m_SelectEvents.SelectedEntities
                    If InStr(1, TypeName(oSelectedObj), "Face") Then
                        oFace = oSelectedObj

                        Select Case sColor
                            Case "Default"
                                oFace.Appearance = m_DefaultColor
                            Case "Red", "Orange", "Yellow", "Green", "Blue", "Indigo", "Purple"
                                oFace.Appearance = m_Doc.AppearanceAssets.Item(sColor)
                        End Select
                    End If
                Next
            End If
        End If
    End Sub
End Class
