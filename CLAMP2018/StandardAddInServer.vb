Imports Inventor
Imports System.Runtime.InteropServices
Imports Microsoft.Win32

Namespace CLAMP2018
    <ProgIdAttribute("CLAMP2018.StandardAddInServer"),
    GuidAttribute("aae1c1d7-01fc-4822-8835-fd043e2daff5")>
    Public Class StandardAddInServer
        Implements Inventor.ApplicationAddInServer

        Private WithEvents oAppUIEv As UserInterfaceEvents
        'Private WithEvents m_sampleButton As ButtonDefinition
        ' Inventor application object.
        Public WithEvents btn2aCADWeb As ButtonDefinition
        Public WithEvents btn2aCADSoporte As ButtonDefinition
        ''
        Public WithEvents btnNewClamp As ButtonDefinition
        Public WithEvents btnClamp As ButtonDefinition
        Public WithEvents btnHousing As ButtonDefinition
        Public WithEvents btnOptions As ButtonDefinition
#Region "ApplicationAddInServer Members"
        ' This method is called by Inventor when it loads the AddIn. The AddInSiteObject provides access  
        ' to the Inventor Application object. The FirstTime flag indicates if the AddIn is loaded for
        ' the first time. However, with the introduction of the ribbon this argument is always true.
        Public Sub Activate(ByVal addInSiteObject As Inventor.ApplicationAddInSite, ByVal firstTime As Boolean) Implements Inventor.ApplicationAddInServer.Activate
            ' Initialize AddIn members.
            oApp = addInSiteObject.Application
            oAppEv = oApp.ApplicationEvents
            ' Connect to the user-interface events to handle a ribbon reset.
            oAppUIEv = oApp.UserInterfaceManager.UserInterfaceEvents

            ' TODO: Add button definitions.

            ' Sample to illustrate creating a button definition.
            'Dim largeIcon As stdole.IPictureDisp = PictureDispConverter.ToIPictureDisp(My.Resources.YourBigImage)
            'Dim smallIcon As stdole.IPictureDisp = PictureDispConverter.ToIPictureDisp(My.Resources.YourSmallImage)
            'Dim controlDefs As Inventor.ControlDefinitions = g_inventorApplication.CommandManager.ControlDefinitions
            'm_sampleButton = controlDefs.AddButtonDefinition("Command Name", "Internal Name", CommandTypesEnum.kShapeEditCmdType, AddInClientID)

            ' Add to the user interface, if it's the first time.
            If firstTime Then
                AddToUserInterface()
            End If
        End Sub

        ' This method is called by Inventor when the AddIn is unloaded. The AddIn will be
        ' unloaded either manually by the user or when the Inventor session is terminated.
        Public Sub Deactivate() Implements Inventor.ApplicationAddInServer.Deactivate

            ' TODO:  Add ApplicationAddInServer.Deactivate implementation

            ' Release objects.
            oAppUIEv = Nothing
            oApp = Nothing
            oAppEv = Nothing

            System.GC.Collect()
            System.GC.WaitForPendingFinalizers()
        End Sub

        ' This property is provided to allow the AddIn to expose an API of its own to other 
        ' programs. Typically, this  would be done by implementing the AddIn's API
        ' interface in a class and returning that class object through this property.
        Public ReadOnly Property Automation() As Object Implements Inventor.ApplicationAddInServer.Automation
            Get
                Return Nothing
            End Get
        End Property

        ' Note:this method is now obsolete, you should use the 
        ' ControlDefinition functionality for implementing commands.
        Public Sub ExecuteCommand(ByVal commandID As Integer) Implements Inventor.ApplicationAddInServer.ExecuteCommand
        End Sub

#End Region

#Region "User interface definition"
        ' Sub where the user-interface creation is done.  This is called when
        ' the add-in loaded and also if the user interface is reset.
        Private Sub AddToUserInterface()
            Dim strG As String = AddInClientID()
            ''
            '' CARGAR RIBBON, RIBBONPANNELS Y RIBBONBUTTONs
            Try
                'MsgBox("Activate--> Create Ribbons")
                ''***** Crearemos los botones de la aplicación. Cargar iconos y definir botones
                ''***** 16x16 para icono pequeño y 32x32 para grande (24x24 en entorno anterior)
                Dim oCd As ControlDefinitions = oApp.CommandManager.ControlDefinitions
                Dim smallPic As stdole.IPictureDisp
                Dim largePic As stdole.IPictureDisp
                '** AQUI INICIO DE BOTONES: Botón inicial de Configuracion'' Creamos los botones antes de crear el Ribbon, 
                ' El internal name debe ser igual en el fichero ribbon.xml... Si lo usamos.
                ' btn2aCADWeb
                smallPic = PictureDispConverter.ToIPictureDisp(New System.Drawing.Icon(My.Resources._2aCAD_2019_ICO, 16, 16)) '  (My.Resources.TipsICO)
                largePic = PictureDispConverter.ToIPictureDisp(New System.Drawing.Icon(My.Resources._2aCAD_2019_ICO, 32, 32))
                btn2aCADWeb = oCd.AddButtonDefinition("2aCAD", PreApp & "2aCADWeb", CommandTypesEnum.kQueryOnlyCmdType,
                                                             strG, "Web 2aCAD", "Web 2aCAD",
                                                             smallPic, largePic, ButtonDisplayEnum.kDisplayTextInLearningMode)
                '
                ' btn2aCADSoporte
                smallPic = PictureDispConverter.ToIPictureDisp(New System.Drawing.Icon(My.Resources.Soporte_Icono, 16, 16)) '  (My.Resources.TipsICO)
                largePic = PictureDispConverter.ToIPictureDisp(New System.Drawing.Icon(My.Resources.Soporte_Icono, 32, 32))
                btn2aCADSoporte = oCd.AddButtonDefinition("2aCAD", PreApp & "2aCADSupport", CommandTypesEnum.kQueryOnlyCmdType,
                                                             strG, "Support on-line 2aCAD", "Support on-line 2aCAD",
                                                             smallPic, largePic, ButtonDisplayEnum.kDisplayTextInLearningMode)
                '
                ' btnNewClamp
                smallPic = PictureDispConverter.ToIPictureDisp(New System.Drawing.Icon(My.Resources.CLAMPS_ICO, 16, 16)) '  (My.Resources.TipsICO)
                largePic = PictureDispConverter.ToIPictureDisp(New System.Drawing.Icon(My.Resources.CLAMPS_ICO, 32, 32))
                btnNewClamp = oCd.AddButtonDefinition(rbButtonNewClamp, rbButtonNewClampId, CommandTypesEnum.kQueryOnlyCmdType,
                                                             strG, rbButtonNewClamp, rbButtonNewClamp,
                                                             smallPic, largePic, ButtonDisplayEnum.kDisplayTextInLearningMode)
                '
                ' btnClamp
                smallPic = PictureDispConverter.ToIPictureDisp(New System.Drawing.Icon(My.Resources.CLAM_ICO, 16, 16)) '  (My.Resources.TipsICO)
                largePic = PictureDispConverter.ToIPictureDisp(New System.Drawing.Icon(My.Resources.CLAM_ICO, 32, 32))
                btnClamp = oCd.AddButtonDefinition(rbButtonClamp, rbButtonClampId, CommandTypesEnum.kQueryOnlyCmdType,
                                                             strG, rbButtonClamp, rbButtonClamp,
                                                             smallPic, largePic, ButtonDisplayEnum.kDisplayTextInLearningMode)
                '
                ' btnHousing
                smallPic = PictureDispConverter.ToIPictureDisp(New System.Drawing.Icon(My.Resources.Housing_ICO, 16, 16)) '  (My.Resources.TipsICO)
                largePic = PictureDispConverter.ToIPictureDisp(New System.Drawing.Icon(My.Resources.Housing_ICO, 32, 32))
                btnHousing = oCd.AddButtonDefinition(rbButtonHousing, rbButtonHousingId, CommandTypesEnum.kQueryOnlyCmdType,
                                                             strG, rbButtonHousing, rbButtonHousing,
                                                             smallPic, largePic, ButtonDisplayEnum.kDisplayTextInLearningMode)
                '
                ' btnOptions
                smallPic = PictureDispConverter.ToIPictureDisp(New System.Drawing.Icon(My.Resources.Configurar_ICO, 16, 16)) '  (My.Resources.TipsICO)
                largePic = PictureDispConverter.ToIPictureDisp(New System.Drawing.Icon(My.Resources.Configurar_ICO, 32, 32))
                btnOptions = oCd.AddButtonDefinition(rbButtonOptions, rbButtonOptionsId, CommandTypesEnum.kQueryOnlyCmdType,
                                                             strG, rbButtonOptions, rbButtonOptions,
                                                             smallPic, largePic, ButtonDisplayEnum.kDisplayTextInLearningMode)
                ''
                '' btnIFeatures
                'smallPic = PictureDispConverter.ToIPictureDisp(New System.Drawing.Icon(My.Resources.iFeatures_ICO, 16, 16)) '  (My.Resources.TipsICO)
                'largePic = PictureDispConverter.ToIPictureDisp(New System.Drawing.Icon(My.Resources.iFeatures_ICO, 32, 32))
                'btnIFeatures = oCd.AddButtonDefinition(rbButtonIFeatures, rbButtonIFeaturesId, CommandTypesEnum.kQueryOnlyCmdType,
                '                                             strG, rbButtonIFeatures, rbButtonIFeatures,
                '                                             smallPic, largePic, ButtonDisplayEnum.kDisplayTextInLearningMode)
                ''
                '' btnRotate
                'smallPic = PictureDispConverter.ToIPictureDisp(New System.Drawing.Icon(My.Resources.Gira3DEje_ICO, 16, 16)) '  (My.Resources.TipsICO)
                'largePic = PictureDispConverter.ToIPictureDisp(New System.Drawing.Icon(My.Resources.Gira3DEje_ICO, 32, 32))
                'btnRotate = oCd.AddButtonDefinition(rbButtonRotate, rbButtonRotateId, CommandTypesEnum.kQueryOnlyCmdType,
                '                                             strG, rbButtonRotate, rbButtonRotate,
                '                                             smallPic, largePic, ButtonDisplayEnum.kDisplayTextInLearningMode)
                '' btnTower
                'smallPic = PictureDispConverter.ToIPictureDisp(New System.Drawing.Icon(My.Resources.GAUGE_T_ICO, 16, 16)) '  (My.Resources.TipsICO)
                'largePic = PictureDispConverter.ToIPictureDisp(New System.Drawing.Icon(My.Resources.GAUGE_T_ICO, 32, 32))
                'btnTower = oCd.AddButtonDefinition(rbButtonTower, rbButtonTowerId, CommandTypesEnum.kQueryOnlyCmdType,
                '                                             strG, rbButtonTower, rbButtonTower,
                '                                             smallPic, largePic, ButtonDisplayEnum.kDisplayTextInLearningMode)
                '' btnTowerR
                'smallPic = PictureDispConverter.ToIPictureDisp(New System.Drawing.Icon(My.Resources.GAUGE_TR_ICO, 16, 16)) '  (My.Resources.TipsICO)
                'largePic = PictureDispConverter.ToIPictureDisp(New System.Drawing.Icon(My.Resources.GAUGE_TR_ICO, 32, 32))
                'btnTowerR = oCd.AddButtonDefinition(rbButtonTowerR, rbButtonTowerRId, CommandTypesEnum.kQueryOnlyCmdType,
                '                                             strG, rbButtonTowerR, rbButtonTowerR,
                '                                             smallPic, largePic, ButtonDisplayEnum.kDisplayTextInLearningMode)
                '' btnMark
                'smallPic = PictureDispConverter.ToIPictureDisp(New System.Drawing.Icon(My.Resources.MARK_ICO, 16, 16)) '  (My.Resources.TipsICO)
                'largePic = PictureDispConverter.ToIPictureDisp(New System.Drawing.Icon(My.Resources.MARK_ICO, 32, 32))
                'btnMark = oCd.AddButtonDefinition(rbButtonMark, rbButtonMarkId, CommandTypesEnum.kQueryOnlyCmdType,
                '                                             strG, rbButtonMark, rbButtonMark,
                '                                             smallPic, largePic, ButtonDisplayEnum.kDisplayTextInLearningMode)
                ''
                '' btnTowerSE
                'smallPic = PictureDispConverter.ToIPictureDisp(New System.Drawing.Icon(My.Resources.GAUGE_T_START_END_ICO, 16, 16)) '  (My.Resources.TipsICO)
                'largePic = PictureDispConverter.ToIPictureDisp(New System.Drawing.Icon(My.Resources.GAUGE_T_START_END_ICO, 32, 32))
                'btnTowerSE = oCd.AddButtonDefinition(rbButtonTowerSE, rbButtonTowerSEId, CommandTypesEnum.kQueryOnlyCmdType,
                '                                             strG, rbButtonTowerSE, rbButtonTowerSE,
                '                                             smallPic, largePic, ButtonDisplayEnum.kDisplayTextInLearningMode)

                '' btnLibrary
                'smallPic = PictureDispConverter.ToIPictureDisp(New System.Drawing.Icon(My.Resources.biblioteca_ico, 16, 16)) '  (My.Resources.TipsICO)
                'largePic = PictureDispConverter.ToIPictureDisp(New System.Drawing.Icon(My.Resources.biblioteca_ico, 32, 32))
                'btnLibrary = oCd.AddButtonDefinition(rbButtonLibrary, rbButtonLibraryId, CommandTypesEnum.kQueryOnlyCmdType,
                '                                             strG, rbButtonLibrary, rbButtonLibrary,
                '                                             smallPic, largePic, ButtonDisplayEnum.kDisplayTextInLearningMode)
                ''''
                '' *** Creamos el nuevo RibbonTab en Entornos  {"ZeroDoc" y "Assembly"}
                'Dim arrRibbon As String() = New String() {"ZeroDoc", "Assembly", "Part", "Drawing", "Presentation", "UnknownDocument"}
                Dim arrRibbon As String() = New String() {"ZeroDoc", "Assembly"}
                For Each queR As String In arrRibbon
                    '** Cargamos el Ribbon (CLAMP)
                    Dim oRib As Inventor.Ribbon = oApp.UserInterfaceManager.Ribbons.Item(queR)
                    '** Cargamos o creamos el nuevo RibbonTab de la aplicación (menuNom)
                    Dim oRibTab As Inventor.RibbonTab
                    Try
                        oRibTab = oRib.RibbonTabs.Item(rbTab)
                    Catch ex As Exception
                        oRibTab = oRib.RibbonTabs.Add(rbTab, rbTabId, strG, , , False)
                    End Try
                    ''
                    '' ***** Creamos el RibbonPanel rbTab.2aCADUtilities
                    Dim oRibPanel As RibbonPanel
                    Try
                        oRibPanel = oRibTab.RibbonPanels.Item(rbPanel2ACADId)
                    Catch ex As Exception
                        oRibPanel = oRibTab.RibbonPanels.Add(rbPanel2ACAD, rbPanel2ACADId, strG)
                    End Try
                    '' ***** Añadimos los 2 RibbonButton a mkTools2aCADUtilities en una Galería
                    Dim oCol As ObjectCollection = oApp.TransientObjects.CreateObjectCollection
                    oCol.Add(btn2aCADWeb)
                    oCol.Add(btn2aCADSoporte)
                    Call oRibPanel.CommandControls.AddButtonPopup(oCol, True)
                    '
                    ' ***** Creamos el RibbonPanel PreApp & Clamps
                    Try
                        oRibPanel = oRibTab.RibbonPanels.Item(rbPanelClampsId)
                    Catch ex As Exception
                        oRibPanel = oRibTab.RibbonPanels.Add(rbPanelClamps, rbPanelClampsId, strG)
                    End Try
                    '
                    '' Añadimos los RibbonButton a PreApp & Clamps
                    'Call oRibPanel.CommandControls.AddSeparator()
                    Call oRibPanel.CommandControls.AddButton(btnNewClamp, True)  ' Icono Grande
                    Call oRibPanel.CommandControls.AddSeparator()
                    Call oRibPanel.CommandControls.AddButton(btnClamp, True)  ' Icono Grande
                    Call oRibPanel.CommandControls.AddSeparator()
                    Call oRibPanel.CommandControls.AddButton(btnHousing, True)  ' Icono Grande
                    'Call oRibPanel.CommandControls.AddSeparator()
                    'Call oRibPanel.CommandControls.AddButton(btnCerramiento, True)  ' Icono Grande
                    'Call oRibPanel.CommandControls.AddSeparator()
                    'Call oRibPanel.CommandControls.AddButton(btnEscalera, True)  ' Icono Grande
                    'Call oRibPanel.CommandControls.AddSeparator()
                    'Call oRibPanel.CommandControls.AddButton(btnConfActual, True)  ' Icono Grande
                    ''
                    '' ***** Creamos el RibbonPanel PreApp & Options
                    Try
                        oRibPanel = oRibTab.RibbonPanels.Item(rbPanelOptionsId)
                    Catch ex As Exception
                        oRibPanel = oRibTab.RibbonPanels.Add(rbPanelOptions, rbPanelOptionsId, strG)
                    End Try
                    '
                    '' Añadimos los RibbonButton a PreApp & Options
                    'Call oRibPanel.CommandControls.AddSeparator()
                    Call oRibPanel.CommandControls.AddButton(btnOptions, True)  ' Icono Grande
                    'Call oRibPanel.CommandControls.AddSeparator()
                Next
                ''
            Catch ex As Exception
                MsgBox("Activate--> Create Ribbons--> " & ex.Message)
                Exit Sub
            End Try
            'MsgBox(oApp.Caption)
            ' This is where you'll add code to add buttons to the ribbon.

            '** Sample to illustrate creating a button on a new panel of the Tools tab of the Part ribbon.

            '' Get the part ribbon.
            'Dim partRibbon As Ribbon = g_inventorApplication.UserInterfaceManager.Ribbons.Item("Part")

            '' Get the "Tools" tab.
            'Dim toolsTab As RibbonTab = partRibbon.RibbonTabs.Item("id_TabTools")

            '' Create a new panel.
            'Dim customPanel As RibbonPanel = toolsTab.RibbonPanels.Add("Sample", "MysSample", AddInClientID)

            '' Add a button.
            'customPanel.CommandControls.AddButton(m_sampleButton)
        End Sub

        Private Sub oAppUIEv_OnResetRibbonInterface(Context As NameValueMap) Handles oAppUIEv.OnResetRibbonInterface
            ' The ribbon was reset, so add back the add-ins user-interface.
            AddToUserInterface()
        End Sub

        ' Sample handler for the button.
        'Private Sub m_sampleButton_OnExecute(Context As NameValueMap) Handles m_sampleButton.OnExecute
        '    MsgBox("Button was clicked.")
        'End Sub
#End Region
#Region "Botones"
        Private Sub btn2aCADWeb_OnExecute(Context As Inventor.NameValueMap) Handles btn2aCADWeb.OnExecute
            Dim Input_URL As String = "http://www.2acad.es"
            'ShellExecute(0&, vbNullString, Input_URL, vbNullString, vbNullString, SW_SHOWNORMAL)
            System.Diagnostics.Process.Start(Input_URL)
        End Sub


        Private Sub btn2aCADSoporte_OnExecute(Context As Inventor.NameValueMap) Handles btn2aCADSoporte.OnExecute
            Dim Input_URL As String = "http://www.2acad.es/soporte-tecnico/"
            'ShellExecute(0&, vbNullString, Input_URL, vbNullString, vbNullString, SW_SHOWNORMAL)
            System.Diagnostics.Process.Start(Input_URL)
        End Sub
        '
        '
        Private Sub btnNewClamp_OnExecute(Context As Inventor.NameValueMap) Handles btnNewClamp.OnExecute
            MsgBox("In construction...")
            '' Solo una instancia del formulario
            'If frmO IsNot Nothing Then Exit Sub
            ''
            '' Stop active command
            'Try
            '    oApp.CommandManager.StopActiveCommand()
            'Catch ex As Exception
            '    'Console.Write(ex.Message)
            'End Try
            '' Close other forms
            'closeForms()
            '' Instance Form
            'frmO = New frmOpciones
            'frmO.Show(New WindowWrapper(oApp.MainFrameHWND))
            'Call frmO.Focus()
        End Sub

        Private Sub btnClamp_OnExecute(Context As Inventor.NameValueMap) Handles btnClamp.OnExecute
            MsgBox("In construction...")
            '' Solo una instancia del formulario
            'If frmO IsNot Nothing Then Exit Sub
            ''
            '' Stop active command
            'Try
            '    oApp.CommandManager.StopActiveCommand()
            'Catch ex As Exception
            '    'Console.Write(ex.Message)
            'End Try
            '' Close other forms
            'closeForms()
            '' Instance Form
            'frmO = New frmOpciones
            'frmO.Show(New WindowWrapper(oApp.MainFrameHWND))
            'Call frmO.Focus()
        End Sub


        Private Sub btnHousing_OnExecute(Context As Inventor.NameValueMap) Handles btnHousing.OnExecute
            MsgBox("In construction...")
            '' Solo una instancia del formulario
            'If frmO IsNot Nothing Then Exit Sub
            ''
            '' Stop active command
            'Try
            '    oApp.CommandManager.StopActiveCommand()
            'Catch ex As Exception
            '    'Console.Write(ex.Message)
            'End Try
            '' Close other forms
            'closeForms()
            '' Instance Form
            'frmO = New frmOpciones
            'frmO.Show(New WindowWrapper(oApp.MainFrameHWND))
            'Call frmO.Focus()
        End Sub

        Private Sub btnOptions_OnExecute(Context As Inventor.NameValueMap) Handles btnOptions.OnExecute
            MsgBox("In construction...")
            '' Solo una instancia del formulario
            'If frmO IsNot Nothing Then Exit Sub
            ''
            '' Stop active command
            'Try
            '    oApp.CommandManager.StopActiveCommand()
            'Catch ex As Exception
            '    'Console.Write(ex.Message)
            'End Try
            '' Close other forms
            'closeForms()
            '' Instance Form
            'frmO = New frmOpciones
            'frmO.Show(New WindowWrapper(oApp.MainFrameHWND))
            'Call frmO.Focus()
        End Sub
        '
        '
        '
        '
        ' VIEJOS, DE EJEMPLO
        'Private Sub btnPlataforma_OnExecute(Context As Inventor.NameValueMap) Handles btnPlataforma.OnExecute
        '    'Dim resultado As String = UserAutorizacion()
        '    'If resultado <> "" Then
        '    '    MsgBox(resultado & vbCrLf & app_nameandversion)
        '    '    Exit Sub
        '    'End If
        '    ''
        '    '' Solo una instancia del formulario
        '    If frmP IsNot Nothing Then Exit Sub
        '    ''
        '    '' Stop active command
        '    Try
        '        oApp.CommandManager.StopActiveCommand()
        '    Catch ex As Exception
        '        'Console.Write(ex.Message)
        '    End Try
        '    '' Close other forms
        '    closeForms()
        '    IAMFinal = ""                       ' Ensamblaje final a abrir. Ya copiado y revinculado
        '    IDWFinal = ""                            ' Plano final a abrir. Ya copiado y revinculado
        '    ultimoIamOrigen = ""
        '    ultimoIamDestino = ""
        '    ultimoPROJECT = ""
        '    '' Instance Form
        '    frmP = New frmPlataforma
        '    frmP.Show(New WindowWrapper(oApp.MainFrameHWND))
        '    Call frmP.Focus()
        'End Sub
        'Private Sub btnCerramiento_OnExecute(Context As Inventor.NameValueMap) Handles btnCerramiento.OnExecute
        '    'Dim resultado As String = UserAutorizacion()
        '    'If resultado <> "" Then
        '    '    MsgBox(resultado & vbCrLf & app_nameandversion)
        '    '    Exit Sub
        '    'End If
        '    ''
        '    '' Solo una instancia del formulario
        '    If frmC IsNot Nothing Then Exit Sub
        '    ''
        '    '' Stop active command
        '    Try
        '        oApp.CommandManager.StopActiveCommand()
        '    Catch ex As Exception
        '        'Console.Write(ex.Message)
        '    End Try
        '    '' Close other forms
        '    closeForms()
        '    IAMFinal = ""                       ' Ensamblaje final a abrir. Ya copiado y revinculado
        '    IDWFinal = ""                            ' Plano final a abrir. Ya copiado y revinculado
        '    '' Instance Form
        '    frmC = New frmCerramiento
        '    'frmC.Show(New WindowWrapper(oApp.MainFrameHWND))
        '    'Call frmC.Focus()
        '    If frmC.ShowDialog(New WindowWrapper(oApp.MainFrameHWND)) = Windows.Forms.DialogResult.OK Then

        '    End If
        'End Sub
        'Private Sub btnEscalera_OnExecute(Context As Inventor.NameValueMap) Handles btnEscalera.OnExecute
        '    'Dim resultado As String = UserAutorizacion()
        '    'If resultado <> "" Then
        '    '    MsgBox(resultado & vbCrLf & app_nameandversion)
        '    '    Exit Sub
        '    'End If
        '    ''
        '    '' Solo una instancia del formulario
        '    If frmE IsNot Nothing Then Exit Sub
        '    ''
        '    '' Stop active command
        '    Try
        '        oApp.CommandManager.StopActiveCommand()
        '    Catch ex As Exception
        '        'Console.Write(ex.Message)
        '    End Try
        '    '' Close other forms
        '    closeForms()
        '    IAMFinal = ""                       ' Ensamblaje final a abrir. Ya copiado y revinculado
        '    IDWFinal = ""                            ' Plano final a abrir. Ya copiado y revinculado
        '    '' Instance Form
        '    frmE = New frmEscalera
        '    'frmE.Show(New WindowWrapper(oApp.MainFrameHWND))
        '    'Call frmE.Focus()
        '    If frmE.ShowDialog(New WindowWrapper(oApp.MainFrameHWND)) = Windows.Forms.DialogResult.OK Then

        '    End If
        'End Sub

        'Private Sub btnConfActual_OnExecute(Context As Inventor.NameValueMap) Handles btnConfActual.OnExecute
        '    'Dim resultado As String = UserAutorizacion()
        '    'If resultado <> "" Then
        '    '    MsgBox(resultado & vbCrLf & app_nameandversion)
        '    '    Exit Sub
        '    'End If
        '    '
        '    ' Si no hay documentos abierto, salir con mensaje.
        '    If oApp.Documents.Count = 0 Then
        '        Dim mensaje As String = "Open an assembly prepared for the configurator first."
        '        MsgBox(mensaje, MsgBoxStyle.Critical, "NOTICE TO USER")
        '        Exit Sub
        '    End If
        '    ''
        '    'Dim iamConfigurar As String = ""
        '    '' Poner el proyecto necesario para el configurador
        '    'If oApp.FileLocations.FileLocationsFile <> IPJ Then
        '    '    ' El proyecto activo NOo es el mismo que InventorProject
        '    '    Dim mensaje As String = "All open files will be closed to change the active project to :" & vbCrLf &
        '    '    IPJ & vbCrLf &
        '    '    ", necessary to work with the configurators."
        '    '    If MsgBox(mensaje, MsgBoxStyle.OkCancel, "NOTICE TO USER") = MsgBoxResult.Ok Then
        '    '        iamConfigurar = oApp.ActiveEditDocument.FullFileName
        '    '        oApp.Documents.CloseAll()
        '    '        oApp.FileLocations.FileLocationsFile = IPJ
        '    '        '
        '    '    End If
        '    'End If
        '    '' Solo una instancia del formulario
        '    If frmPnew IsNot Nothing Then Exit Sub
        '    ''
        '    '' Stop active command
        '    Try
        '        oApp.CommandManager.StopActiveCommand()
        '    Catch ex As Exception
        '        'Console.Write(ex.Message)
        '    End Try
        '    '' Close other forms
        '    closeForms()
        '    '' Instance Form
        '    'frmPnew = New frmNewPlatform
        '    'frmP.Show(New WindowWrapper(oApp.MainFrameHWND))
        '    'Call frmP.Focus()
        '    '
        '    'If iamConfigurar <> "" And IO.File.Exists(iamConfigurar) Then
        '    '    Call oApp.Documents.Open(iamConfigurar)
        '    'End If
        '    Dim queTipo As String = clsIAv.PropiedadLeeUsuario(oApp.ActiveEditDocument, "TYPE")
        '    ' ***** Finalmento abrimos el nuevo ensamblaje.
        '    If oApp.ActiveDocumentType = DocumentTypeEnum.kAssemblyDocumentObject AndAlso queTipo <> "" AndAlso queTipo = tipoPlatform Then    ' Plataformas
        '        If oApp.ActiveEditDocument.RequiresUpdate Then
        '            oApp.ActiveEditDocument.Update2()
        '            oApp.ActiveEditDocument.Rebuild2()
        '        End If
        '        If oApp.ActiveEditDocument.Dirty Then oApp.ActiveEditDocument.Save2()
        '        ' Poner visible y activar el documento
        '        'Call oApp.ActiveEditDocument.Views.Add()
        '        '' Instance Form
        '        frmPnew = New frmNewPlatform
        '        frmPnew.Show(New WindowWrapper(oApp.MainFrameHWND))
        '        Call frmPnew.Focus()
        '    ElseIf oApp.ActiveDocumentType = DocumentTypeEnum.kAssemblyDocumentObject AndAlso queTipo <> "" AndAlso queTipo = tipoLadders Then ' Escaleras
        '    ElseIf oApp.ActiveDocumentType = DocumentTypeEnum.kAssemblyDocumentObject AndAlso queTipo <> "" AndAlso queTipo = tipoFences Then  ' Cercados
        '    Else
        '        Dim mensaje As String = "Open an assembly prepared for the configurator first."
        '        MsgBox(mensaje, MsgBoxStyle.Critical, "NOTICE TO USER")
        '        Exit Sub
        '    End If
        'End Sub

        'Private Sub btnBomExcel_OnExecute(Context As Inventor.NameValueMap) Handles btnBomExcel.OnExecute
        '    'Dim resultado As String = UserAutorizacion()
        '    'If resultado <> "" Then
        '    '    MsgBox(resultado & vbCrLf & app_nameandversion)
        '    '    Exit Sub
        '    'End If
        '    '
        '    ' Si no hay documentos abierto, salir con mensaje.
        '    If oApp.Documents.Count = 0 OrElse oApp.ActiveEditDocument.DocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then
        '        Dim mensaje As String = "Open an assembly prepared for the configurator first."
        '        MsgBox(mensaje, MsgBoxStyle.Critical, "NOTICE TO USER")
        '        Exit Sub
        '    End If
        '    '
        '    ' Solo una instancia del formulario
        '    'If frmPnew IsNot Nothing Then Exit Sub
        '    ''
        '    '' Stop active command
        '    Try
        '        oApp.CommandManager.StopActiveCommand()
        '    Catch ex As Exception
        '        'Console.Write(ex.Message)
        '    End Try
        '    '' Close other forms
        '    closeForms()
        '    '' Instance Form
        '    'frmPnew = New frmNewPlatform
        '    'frmP.Show(New WindowWrapper(oApp.MainFrameHWND))
        '    'Call frmP.Focus()
        '    '
        '    'If iamConfigurar <> "" And IO.File.Exists(iamConfigurar) Then
        '    '    Call oApp.Documents.Open(iamConfigurar)
        '    'End If
        '    Dim queTipo As String = clsIAv.PropiedadLeeUsuario(oApp.ActiveEditDocument, "TYPE")
        '    ' ***** Finalmento abrimos el nuevo ensamblaje.
        '    If oApp.ActiveDocumentType = DocumentTypeEnum.kAssemblyDocumentObject AndAlso queTipo <> "" Then    ' Plataformas
        '        If oApp.ActiveEditDocument.RequiresUpdate Then
        '            oApp.ActiveEditDocument.Update2()
        '            oApp.ActiveEditDocument.Rebuild2()
        '        End If
        '        If oApp.ActiveEditDocument.Dirty Then oApp.ActiveEditDocument.Save2()
        '        ' Poner visible y activar el documento
        '        'Call oApp.ActiveEditDocument.Views.Add()
        '        '' Instance Form
        '        'frmPnew = New frmNewPlatform
        '        'frmPnew.Show(New WindowWrapper(oApp.MainFrameHWND))
        '        'Call frmPnew.Focus()
        '        Dim oAsm As AssemblyDocument = CType(oApp.ActiveDocument, AssemblyDocument)
        '        Dim cCos As New clsCOcus(oAsm)
        '        If cCos IsNot Nothing Then
        '            cCos.ListaTotales
        '        End If

        '        If IO.File.Exists(cCos.fullFiCSV) Then
        '            If MsgBox("Open " & cCos.fullFiCSV, MsgBoxStyle.OkCancel, "Open Excel file") = MsgBoxResult.Ok Then
        '                Process.Start(cCos.fullFiCSV)
        '            End If
        '        End If

        '        'Dim fiExcel As String = IO.Path.ChangeExtension(oAsm.FullFileName, "BOM1.xls")
        '        ''
        '        'Dim oACd As AssemblyComponentDefinition = oAsm.ComponentDefinition
        '        '' Poner representación Master antes o dará error
        '        'oACd.RepresentationsManager.LevelOfDetailRepresentations.Item(1).Activate()
        '        'Dim oBom As BOM = oACd.BOM
        '        ''
        '        'oBom.StructuredViewFirstLevelOnly = False
        '        ' Activar las vistas "Estructurado" y "Solo Piezas"
        '        'oBom.PartsOnlyViewEnabled = True
        '        'oBom.StructuredViewEnabled = True
        '        ''
        '        ''Set oBOMView = oBOM.BOMViews.Item(2)   '(1)Sin nombre '(2)Estructurado   '(3)Sólo piezas
        '        'Dim oPVi As BOMView = oBom.BOMViews.Item(3)
        '        'oPVi.Export(fiExcel, FileFormatEnum.kMicrosoftExcelFormat)
        '        'If IO.File.Exists(fiExcel) Then
        '        '    If MsgBox("Open " & fiExcel, MsgBoxStyle.OkCancel, "Open Excel file") = MsgBoxResult.Ok Then
        '        '        Process.Start(fiExcel)
        '        '    End If
        '        'End If
        '    Else
        '        Dim mensaje As String = "Open an assembly prepared for the configurator first."
        '        MsgBox(mensaje, MsgBoxStyle.Critical, "NOTICE TO USER")
        '        Exit Sub
        '    End If
        'End Sub

        'Private Sub btnPackAndGo_OnExecute(Context As Inventor.NameValueMap) Handles btnPackAndGo.OnExecute
        '    'Dim resultado As String = UserAutorizacion()
        '    'If resultado <> "" Then
        '    '    MsgBox(resultado & vbCrLf & app_nameandversion)
        '    '    Exit Sub
        '    'End If
        '    ' Si no hay documentos abierto, salir con mensaje.
        '    If oApp.Documents.Count = 0 Then
        '        Dim mensaje As String = "Open an assembly prepared for the configurator first."
        '        MsgBox(mensaje, MsgBoxStyle.Critical, "NOTICE TO USER")
        '        Exit Sub
        '    End If
        '    '
        '    If oApp.ActiveDocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then
        '        Dim mensaje As String = "PackAndGo. Only for assembly documents (IAM)"
        '        MsgBox(mensaje, MsgBoxStyle.Critical, "NOTICE TO USER")
        '        Exit Sub
        '    End If
        '    '
        '    ' Stop active command
        '    Try
        '        oApp.CommandManager.StopActiveCommand()
        '    Catch ex As Exception
        '        'Console.Write(ex.Message)
        '    End Try
        '    '' Close other forms
        '    closeForms()
        '    '
        '    ' ***** Actualizar y guardar el ensamblaje.
        '    If oApp.ActiveEditDocument.RequiresUpdate Then
        '        oApp.ActiveEditDocument.Update2()
        '        oApp.ActiveEditDocument.Rebuild2()
        '    End If
        '    If oApp.ActiveEditDocument.Dirty Then oApp.ActiveEditDocument.Save2()
        '    '
        '    ' llamar a nuestro pag2aCAD.exe.
        '    Dim oPi As New ProcessStartInfo
        '    oPi.FileName = _apppag2acad
        '    oPi.Arguments = "" & Chr(34) & "" & oApp.ActiveDocument.FullFileName & "" & Chr(34) & ""
        '    Process.Start(oPi)
        '    oPi = Nothing
        'End Sub
#End Region
    End Class
End Namespace


Public Module Globals
    ' Inventor application object.
    'Public g_inventorApplication As Inventor.Application

#Region "Function to get the add-in client ID."
    ' This function uses reflection to get the GuidAttribute associated with the add-in.
    Public Function AddInClientID() As String
        Dim guid As String = ""
        Try
            Dim t As Type = GetType(CLAMP2018.StandardAddInServer)
            Dim customAttributes() As Object = t.GetCustomAttributes(GetType(GuidAttribute), False)
            Dim guidAttribute As GuidAttribute = CType(customAttributes(0), GuidAttribute)
            guid = "{" + guidAttribute.Value.ToString() + "}"
        Catch
        End Try

        Return guid
    End Function
#End Region

#Region "hWnd Wrapper Class"
    ' This class is used to wrap a Win32 hWnd as a .Net IWind32Window class.
    ' This is primarily used for parenting a dialog to the Inventor window.
    '
    ' For example:
    ' myForm.Show(New WindowWrapper(g_inventorApplication.MainFrameHWND))
    '
    Public Class WindowWrapper
        Implements System.Windows.Forms.IWin32Window
        Public Sub New(ByVal handle As IntPtr)
            _hwnd = handle
        End Sub

        Public ReadOnly Property Handle() As IntPtr _
          Implements System.Windows.Forms.IWin32Window.Handle
            Get
                Return _hwnd
            End Get
        End Property

        Private _hwnd As IntPtr
    End Class
#End Region

#Region "Image Converter"
    ' Class used to convert bitmaps and icons from their .Net native types into
    ' an IPictureDisp object which is what the Inventor API requires. A typical
    ' usage is shown below where MyIcon is a bitmap or icon that's available
    ' as a resource of the project.
    '
    ' Dim smallIcon As stdole.IPictureDisp = PictureDispConverter.ToIPictureDisp(My.Resources.MyIcon)

    Public NotInheritable Class PictureDispConverter
        <DllImport("OleAut32.dll", EntryPoint:="OleCreatePictureIndirect", ExactSpelling:=True, PreserveSig:=False)> _
        Private Shared Function OleCreatePictureIndirect( _
            <MarshalAs(UnmanagedType.AsAny)> ByVal picdesc As Object, _
            ByRef iid As Guid, _
            <MarshalAs(UnmanagedType.Bool)> ByVal fOwn As Boolean) As stdole.IPictureDisp
        End Function

        Shared iPictureDispGuid As Guid = GetType(stdole.IPictureDisp).GUID

        Private NotInheritable Class PICTDESC
            Private Sub New()
            End Sub

            'Picture Types
            Public Const PICTYPE_BITMAP As Short = 1
            Public Const PICTYPE_ICON As Short = 3

            <StructLayout(LayoutKind.Sequential)> _
            Public Class Icon
                Friend cbSizeOfStruct As Integer = Marshal.SizeOf(GetType(PICTDESC.Icon))
                Friend picType As Integer = PICTDESC.PICTYPE_ICON
                Friend hicon As IntPtr = IntPtr.Zero
                Friend unused1 As Integer
                Friend unused2 As Integer

                Friend Sub New(ByVal icon As System.Drawing.Icon)
                    Me.hicon = icon.ToBitmap().GetHicon()
                End Sub
            End Class

            <StructLayout(LayoutKind.Sequential)> _
            Public Class Bitmap
                Friend cbSizeOfStruct As Integer = Marshal.SizeOf(GetType(PICTDESC.Bitmap))
                Friend picType As Integer = PICTDESC.PICTYPE_BITMAP
                Friend hbitmap As IntPtr = IntPtr.Zero
                Friend hpal As IntPtr = IntPtr.Zero
                Friend unused As Integer

                Friend Sub New(ByVal bitmap As System.Drawing.Bitmap)
                    Me.hbitmap = bitmap.GetHbitmap()
                End Sub
            End Class
        End Class

        Public Shared Function ToIPictureDisp(ByVal icon As System.Drawing.Icon) As stdole.IPictureDisp
            Dim pictIcon As New PICTDESC.Icon(icon)
            Return OleCreatePictureIndirect(pictIcon, iPictureDispGuid, True)
        End Function

        Public Shared Function ToIPictureDisp(ByVal bmp As System.Drawing.Bitmap) As stdole.IPictureDisp
            Dim pictBmp As New PICTDESC.Bitmap(bmp)
            Return OleCreatePictureIndirect(pictBmp, iPictureDispGuid, True)
        End Function
    End Class
#End Region

End Module
