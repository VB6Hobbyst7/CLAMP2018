Imports System
Imports System.Windows.Forms
'
Imports Inventor2acad
Imports ua = UtilesAlberto
Imports uau = UtilesAlberto.Utiles

Module modVar
    ' ***** INVENTOR OBJECTS
    Public oApp As Inventor.Application
    Public oAppEv As Inventor.ApplicationEvents
    'Public oAppUIEv As Inventor.UserInterfaceEvents
    '
    ' ***** CLASS
    Public clsI As Inventor2acad.Inventor2acad = Nothing
    Public cfg As ua.Conf
    '
    '***** FORMULARIOS
    'Public frmCre As frmCreate
    'Public frmOpt As frmOpciones
    'Public frmR3D As frmRotate3D
    'Public frmT As frmTower
    'Public frmTR As frmTowerR
    'Public frmM As frmMark
    'Public frmTse As frmTowerStartEnd
    'Public frmL As frmLibrary
    '
    '' VARIABLES UI INVENTOR
    ' RIBBON
    Public rbTab As String = CLAMP
    Public rbTabId As String = PreApp & rbTab
    ' RIBBONPANNEL 2aCAD
    Public rbPanel2ACAD As String = "2aCAD Utilities"
    Public rbPanel2ACADId As String = rbTabId & "." & rbPanel2ACAD
    Public rbButton2aCADWeb As String = "2aCAD Web"
    Public rbButton2aCADWebId As String = PreApp & rbButton2aCADWeb
    Public rbButton2aCADSupport As String = "2aCAD Support"
    Public rbButton2aCADSupportId As String = PreApp & rbButton2aCADSupport
    ' RIBBONPANNELS CLAMP
    Public rbPanelClamps As String = "Clamps"
    Public rbPanelClampsId As String = PreApp & rbPanelClamps
    Public rbPanelOptions As String = "Options"
    Public rbPanelOptionsId As String = PreApp & rbPanelOptions
    '
    ' RIBBONBUTTONS
    Public rbButtonNewClamp As String = "New Clamp"
    Public rbButtonNewClampId As String = PreApp & rbButtonNewClamp
    Public rbButtonClamp As String = "Clamp"    ' Alojamiento
    Public rbButtonClampId As String = PreApp & rbButtonClamp
    Public rbButtonHousing As String = "Housing"    ' Alojamiento
    Public rbButtonHousingId As String = PreApp & rbButtonHousing

    Public rbButtonOptions As String = "Options"
    Public rbButtonOptionsId As String = PreApp & rbButtonOptions
    '
    ' ***** CONFIGURATION
    Public Const CLAMP As String = "CLAMP"
    Public Const PreCLAMP As String = "CLAMP·"
    Public Const PreApp As String = "CLAMP2018·"
    '
    Public Log As Boolean = False
    Public FolderProjects As String = "PROJECTS"
    Public FolderTemplates As String = "TEMPLATES"
    Public TemplateIPJ As String = "CLAMP.ipj"
    '
    ' ***** GENERALES
    Public _appCLAMP As String = ""

    '
    Public Function Config_LeeTodo() As String
        '[OPTIONS]
        'Log = 1
        ';
        '[PATHS]
        'FolderProjects=PROJECTS
        'FolderTemplates=TEMPLATES
        'TemplateIPJ = CLAMP.ipj
        Dim resultado As String = ""
        ' ***** Run Start
        If clsI Is Nothing Then clsI = New Inventor2acad.Inventor2acad(oApp)
        If cfg Is Nothing Then cfg = New ua.Conf(System.Reflection.Assembly.GetExecutingAssembly)

        Try
            _appCLAMP = IO.Path.Combine(cfg._appfolder, CLAMP)
            ' ***** OPTIONS
            Dim LogTemp As String = uau.IniGet(cfg._appini, "OPTIONS", "Log")
            If LogTemp = "1" Then Log = True
            '
            Try
                cfg._Log = Log
            Catch ex As ArgumentException
            End Try
            '

            FolderProjects = uau.IniGet(cfg._appini, "PATHS", "FolderProjects").ToString
            FolderProjects = IO.Path.Combine(_appCLAMP, FolderProjects)
            FolderTemplates = uau.IniGet(cfg._appini, "PATHS", "FolderTemplates").ToString
            FolderTemplates = IO.Path.Combine(_appCLAMP, FolderTemplates)
            TemplateIPJ = uau.IniGet(cfg._appini, "PATHS", "TemplateIPJ").ToString
            TemplateIPJ = IO.Path.Combine(_appCLAMP, TemplateIPJ)
            '
            ' ***** PATHS
            ' Path principal, si este falla, no continuamos con el resto.
            If IO.Directory.Exists(_appCLAMP) = False Then
                Try
                    IO.Directory.CreateDirectory(_appCLAMP)
                Catch ex As Exception
                    resultado &= "Error creating the directory:" & vbCrLf & vbCrLf & _appCLAMP & vbCrLf & vbCrLf & "Check permissions..."
                    GoTo FINAL
                    Exit Function
                End Try
            End If
            '
            ' El resto de los Path, que cuelgan de CLAMP
            If IO.Directory.Exists(FolderProjects) = False Then
                Try
                    IO.Directory.CreateDirectory(FolderProjects)
                Catch ex As Exception
                    resultado &= "Error creating the directory:" & vbCrLf & vbCrLf & FolderProjects & vbCrLf & vbCrLf & "Check permissions..."
                End Try
            End If
            '
            If IO.Directory.Exists(FolderTemplates) = False Then
                Try
                    IO.Directory.CreateDirectory(FolderTemplates)
                Catch ex As Exception
                    resultado &= "Error creating the directory:" & vbCrLf & vbCrLf & FolderTemplates & vbCrLf & vbCrLf & "Check permissions..."
                End Try
            End If
            '
            ' El fichero .ipj
            If IO.File.Exists(TemplateIPJ) = False Then
                Try
                    IO.Directory.CreateDirectory(TemplateIPJ)
                Catch ex As Exception
                    resultado &= "Not exist:" & vbCrLf & vbCrLf & TemplateIPJ & vbCrLf & vbCrLf & "Impossible to continue ......"
                End Try
            End If
            '
        Catch ex As Exception
            resultado &= ex.ToString
        End Try
        '
FINAL:
        Return resultado
    End Function
    '
    Public Sub closeForms()
        On Error Resume Next
        'If Not (frmL Is Nothing) Then frmL.Close()
        'frmL = Nothing
        'If Not (frmO Is Nothing) Then frmO.Close()
        'frmO = Nothing
        'If Not (frmP Is Nothing) Then frmP.Close()
        'frmP = Nothing
        'If Not (frmC Is Nothing) Then frmC.Close()
        'frmC = Nothing
        'If Not (frmE Is Nothing) Then frmE.Close()
        'frmE = Nothing
        ''
        'If frmPnew IsNot Nothing Then frmPnew.Close()
        'frmPnew = Nothing
    End Sub
End Module
