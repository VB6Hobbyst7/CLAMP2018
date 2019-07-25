Imports Inventor
Partial Public Class Inventor2acad
    Public Sub Export_DWF3DFull(DocIn As Inventor.Document, Optional FileOut As String = "", Optional ExtReplace As Boolean = False, Optional DWFx As Boolean = True)
        'oApp.ActiveEditDocument.SaveAs(oApp.ActiveEditDocument.FullFileName & ".dwf", True)
        Dim FileIn As String = DocIn.FullFileName
        If FileOut = "" Then
            If ExtReplace Then
                FileOut = IO.Path.ChangeExtension(FileIn, ".dwf" & IIf(DWFx, "x", ""))
            Else
                FileOut = FileIn & ".dwf"
            End If
        End If
        DocIn.SaveAs(FileOut, True)
    End Sub
    Public Sub Export_DWF3DLight(DocIn As Inventor.Document, Optional FileOut As String = "", Optional ExtReplace As Boolean = False, Optional DWFx As Boolean = True)
        Dim FileIn As String = DocIn.FullFileName
        If FileOut = "" Then
            If ExtReplace Then
                FileOut = IO.Path.ChangeExtension(FileIn, ".dwf" & IIf(DWFx, "x", ""))
            Else
                FileOut = FileIn & ".dwf"
            End If
        End If

        ' Get the DWF translator Add-In.
        Dim DWFAddIn As TranslatorAddIn
        DWFAddIn = oAppI.ApplicationAddIns.ItemById("{0AC6FD95-2F4D-42CE-8BE0-8AEA580399E4}")

        Dim oContext As TranslationContext
        oContext = oAppI.TransientObjects.CreateTranslationContext
        oContext.Type = IOMechanismEnum.kFileBrowseIOMechanism

        ' Create a NameValueMap object
        Dim oOptions As NameValueMap
        oOptions = oAppI.TransientObjects.CreateNameValueMap

        ' Create a DataMedium object
        Dim oDataMedium As DataMedium
        oDataMedium = oAppI.TransientObjects.CreateDataMedium

        ' Check whether the translator has 'SaveCopyAs' options
        If DWFAddIn.HasSaveCopyAsOptions(DocIn, oContext, oOptions) Then
            oOptions.Value("Launch_Viewer") = 0
            oOptions.Value("Publish_All_Component_Props") = False
            oOptions.Value("Publish_All_Physical_Props") = False
            If DocIn.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                oOptions.Value("Enable_Large_Assembly_Mode") = False
                oOptions.Value("BOM_Structured") = False
                oOptions.Value("BOM_Parts_Only") = False
            ElseIf DocIn.DocumentType = DocumentTypeEnum.kPresentationDocumentObject Then
                oOptions.Value("BOM_Structured") = False
                oOptions.Value("BOM_Parts_Only") = False
                oOptions.Value("Animations") = False
                oOptions.Value("Instructions") = False
            End If
            oOptions.Value("Include_Empty_Properties") = False
            oOptions.Value("Output_Path") = FileOut
            oOptions.Value("Facet_Quality") = AccuracyEnum.kLow

            '    If TypeOf oDocument Is DrawingDocument Then
            '        Drawing options
            '        oOptions.Value("Publish_Mode") = kCustomDWFPublish
            '        oOptions.Value("Publish_All_Sheets") = 0
            '        The specified sheets will be ignored if
            '         the option "Publish_All_Sheets" Is True (1)
            '        Dim oSheets As NameValueMap
            '    Set oSheets = ThisApplication.TransientObjects.CreateNameValueMap
            '     Publish the first sheet And its 3D model
            '    Dim oSheet1Options As NameValueMap
            '    Set oSheet1Options = ThisApplication.TransientObjects.CreateNameValueMap

            '    oSheet1Options.Add "Name", "Sheet:1"
            '    oSheet1Options.Add "3DModel", True
            '    oSheets.Value("Sheet1") = oSheet1Options

            '        Publish the third sheet but Not its 3D model
            '        Dim oSheet3Options As NameValueMap
            '    Set oSheet3Options = ThisApplication.TransientObjects.CreateNameValueMap

            '    oSheet3Options.Add "Name", "Sheet:3"
            '    oSheet3Options.Add "3DModel", False

            '    oSheets.Value("Sheet2") = oSheet3Options

            '        Set the sheet options object in the oOptions NameValueMap
            '        oOptions.Value("Sheets") = oSheets
            'End If

        End If

        'Set the destination file name
        oDataMedium.FileName = FileOut

        'Publish document.
        Call DWFAddIn.SaveCopyAs(DocIn, oContext, oOptions, oDataMedium)
    End Sub

    Public Sub Export_DrawingSaveAs(ByVal oDoc As Inventor.DrawingDocument,
                                         ByVal queDestino As String,
                                         ByVal queTipos As ArrayList)    ', ByVal abrirlo As Boolean)

        ' Todos los tipos de una Enum como array de cadenas: [Enum].GetNames(GetType(ENUM))
        ' Un tipo concreto de una Enum como cadena: [Enum].GetName(GetType(GuardaTipo), ultimoGT)
        If queDestino.EndsWith("\") = False Then queDestino &= "\"
        '' Nombre final del fichero (SIN EXTENSION)
        Dim ficheroFin As String = queDestino & DameParteCamino(oDoc.FullFileName, IEnum.ParteCamino.SoloFicheroSinExtension)

        'For Each oSh As Sheet In oDoc.Sheets
        'ZoomTodoAjustar2D(oSh)
        'Next

        Dim ultimoGT As IEnum.GuardaTipo
        'Publish document.
        For Each queG As IEnum.GuardaTipo In queTipos
            ultimoGT = queG
            '' SaveAs no permite guardar formatos DWG ni DXF. Utilizar TranslatorAddIn.
            If queG = IEnum.GuardaTipo.dwg Or queG = IEnum.GuardaTipo.dxf Or queG = IEnum.GuardaTipo.pdf Or queG = IEnum.GuardaTipo.dwf Or queG = IEnum.GuardaTipo.dwfx Then Continue For
            Dim fFin As String = ficheroFin & "." & [Enum].GetName(GetType(IEnum.GuardaTipo), ultimoGT)
            Try
                If IO.File.Exists(fFin) Then IO.File.Delete(fFin)
                oDoc.SaveAs(fFin, True)
                Console.WriteLine("Guardado correctamente : " & fFin)
                'If abrirlo = True Then Call Process.Start(fFin)
            Catch ex As Exception
                MsgBox("Error SaveAs (" & [Enum].GetName(GetType(IEnum.GuardaTipo), ultimoGT) & ") con " & fFin)
                Console.WriteLine("Error SaveAs (" & [Enum].GetName(GetType(IEnum.GuardaTipo), ultimoGT) & ") con " & fFin)
            End Try
        Next
    End Sub

    Public Sub Export_DrawingPDFAddIn(ByVal oDoc As Inventor.DrawingDocument,
                                ByVal queDestino As String)
        If IO.Directory.Exists(queDestino) = False Then
            Call IO.Directory.CreateDirectory(queDestino)
        End If
        If queDestino.EndsWith("\") = False Then queDestino &= "\"
        '' Nombre del fichero PDF destino.
        Dim ficheroPDF As String = queDestino & DameParteCamino(oDoc.FullFileName, IEnum.ParteCamino.SoloFicheroSinExtension) & ".pdf"
        ' Get the PDF translator Add-In.
        Dim PDFAddIn As TranslatorAddIn
        PDFAddIn = Me.oAppI.ApplicationAddIns.ItemById("{0AC6FD96-2F4D-42CE-8BE0-8AEA580399E4}")

        Dim oContext As TranslationContext
        oContext = Me.oAppI.TransientObjects.CreateTranslationContext
        oContext.Type = IOMechanismEnum.kFileBrowseIOMechanism

        ' Create a NameValueMap object
        Dim oOptions As NameValueMap
        oOptions = Me.oAppI.TransientObjects.CreateNameValueMap

        ' Create a DataMedium object
        Dim oDataMedium As DataMedium
        oDataMedium = Me.oAppI.TransientObjects.CreateDataMedium

        ' Check whether the translator has 'SaveCopyAs' options
        If PDFAddIn.HasSaveCopyAsOptions(oDoc, oContext, oOptions) Then

            ' Options for drawings...

            oOptions.Value("All_Color_AS_Black") = False     'True o 1 / 0 o False

            'oOptions.Value("Remove_Line_Weights") = True	' 0 o False
            oOptions.Value("Vector_Resolution") = 400
            oOptions.Value("Sheet_Range") = PrintRangeEnum.kPrintAllSheets
            'oOptions.Value("Custom_Begin_Sheet") = 2
            'oOptions.Value("Custom_End_Sheet") = 4

        End If

        'Set the destination file name
        oDataMedium.FileName = ficheroPDF   ' "c:\temp\test.pdf"

        Try
            If IO.File.Exists(ficheroPDF) Then IO.File.Delete(ficheroPDF)
            'Publish document.
            Call PDFAddIn.SaveCopyAs(oDoc, oContext, oOptions, oDataMedium)
            'Call Process.Start(ficheroPDF)
        Catch ex As Exception
            Console.WriteLine("Error exportar PDF con " & ficheroPDF)
            Debug.Print("Error exportar DWG con " & ficheroPDF)
        End Try
    End Sub

    Public Sub Export_3DToSat(ByVal oDoc As Inventor.Document)
        ' Set reference to active document.
        'Dim oDoc As Inventor.Document
        'oDoc = oApp.ActiveDocument

        ' Check the Document type is an assembly or part
        If (oDoc.DocumentType <> DocumentTypeEnum.kAssemblyDocumentObject And
      oDoc.DocumentType <> DocumentTypeEnum.kPartDocumentObject) Then
            MsgBox("Error:Document type is not assembly/part")
            oDoc = Nothing
            Exit Sub
        End If

        ' Get document's full file name
        Dim sFname As String
        sFname = oDoc.FullFileName

        ' The file format will depend on the extension
        ' Set file name extension to ".SAT"
        sFname = Microsoft.VisualBasic.Left(sFname, Len(sFname) - 3) & "sat"
        Try
            If IO.File.Exists(sFname) Then IO.File.Delete(sFname)
            ' Do a 'Save Copy As' to SAT format
            Call oDoc.SaveAs(sFname, True)
        Catch ex As Exception
            Debug.Print("Error guardando " & sFname)
        End Try

        oDoc = Nothing
    End Sub


    ' PARA EXPORTAR A DWG (Solo la opción del fichero .INI)
    ' Dim strIniFile As String
    ' strIniFile = "C:\tempDWGOut.ini"
    ' Create the name-value that specifies the ini file to use.
    ' oOptions.Value("Export_Acad_IniFile") = strIniFile
    Public Sub Export_DrawingDXFAddin(ByVal oDoc As Inventor.DrawingDocument, ByVal dirDestino As String, ByVal ficheroINI As String, ByVal esDWG As Boolean)
        If dirDestino.EndsWith("\") = False Then dirDestino &= "\"
        '' Nombre final del fichero exportado (DWG o DXF)
        Dim ficheroFIN As String
        If esDWG = True Then
            ficheroFIN = dirDestino & DameParteCamino(oDoc.FullFileName, IEnum.ParteCamino.SoloFicheroSinExtension) & ".dwg"
        Else
            ficheroFIN = dirDestino & DameParteCamino(oDoc.FullFileName, IEnum.ParteCamino.SoloFicheroSinExtension) & ".dxf"
        End If
        If IO.File.Exists(ficheroFIN) = True Then Exit Sub

        ' Get the PDF translator Add-In.
        Dim DWGDXFAddIn As TranslatorAddIn
        If esDWG = True Then
            DWGDXFAddIn = oAppI.ApplicationAddIns.ItemById("{C24E3AC2-122E-11D5-8E91-0010B541CD80}")
        Else
            DWGDXFAddIn = oAppI.ApplicationAddIns.ItemById("{C24E3AC4-122E-11D5-8E91-0010B541CD80}")
        End If
        If DWGDXFAddIn.Activated = False Then DWGDXFAddIn.Activate()

        'For Each oSh As Sheet In oDoc.Sheets
        'ZoomTodoAjustar2D(oSh)
        'Next

        Dim oContext As TranslationContext
        oContext = oAppI.TransientObjects.CreateTranslationContext
        oContext.Type = IOMechanismEnum.kFileBrowseIOMechanism

        ' Create a NameValueMap object
        Dim oOptions As NameValueMap
        oOptions = oAppI.TransientObjects.CreateNameValueMap

        ' Create a DataMedium object
        Dim oDataMedium As DataMedium
        oDataMedium = oAppI.TransientObjects.CreateDataMedium


        ' Check whether the translator has 'SaveCopyAs' options
        If DWGDXFAddIn.HasSaveCopyAsOptions(oDoc, oContext, oOptions) AndAlso IO.File.Exists(ficheroINI) = True Then
            ' Create the name-value that specifies the ini file to use.
            oOptions.Value("Export_Acad_IniFile") = ficheroINI
        End If

        'Set the destination file name
        oDataMedium.FileName = ficheroFIN   ' "c:\temp\test.pdf"

        Try
            If IO.File.Exists(ficheroFIN) Then IO.File.Delete(ficheroFIN)
            'Publish document.
            Call DWGDXFAddIn.SaveCopyAs(oDoc, oContext, oOptions, oDataMedium)
            '' Abrir el fichero, una vez se ha creado
            'Call Process.Start(ficheroDWG)
        Catch ex As Exception
            Console.WriteLine("Error exportar DWG con " & ficheroFIN)
            Debug.Print("Error exportar DWG con " & ficheroFIN)
        End Try
    End Sub

    Public Sub Export_DrawingDWFAddin(ByVal oDoc As Inventor.DrawingDocument, ByVal dirDestino As String, ByVal esDWF As Boolean)
        If dirDestino.EndsWith("\") = False Then dirDestino &= "\"
        '' Nombre final del fichero exportado (dwf o dwfx)
        Dim ficheroFIN As String
        If esDWF = True Then
            ficheroFIN = dirDestino & DameParteCamino(oDoc.FullFileName, IEnum.ParteCamino.SoloFicheroSinExtension) & ".dwf"
        Else
            ficheroFIN = dirDestino & DameParteCamino(oDoc.FullFileName, IEnum.ParteCamino.SoloFicheroSinExtension) & ".dwfx"
        End If
        If IO.File.Exists(ficheroFIN) = True Then Exit Sub

        ' Get the DWF translator Add-In.
        Dim DWFAddIn As TranslatorAddIn
        DWFAddIn = Me.oAppI.ApplicationAddIns.ItemById("{0AC6FD95-2F4D-42CE-8BE0-8AEA580399E4}")

        Dim oContext As TranslationContext
        oContext = Me.oAppI.TransientObjects.CreateTranslationContext
        oContext.Type = IOMechanismEnum.kFileBrowseIOMechanism

        ' Create a NameValueMap object
        Dim oOptions As NameValueMap
        oOptions = Me.oAppI.TransientObjects.CreateNameValueMap

        ' Create a DataMedium object
        Dim oDataMedium As DataMedium
        oDataMedium = Me.oAppI.TransientObjects.CreateDataMedium

        ' Check whether the translator has 'SaveCopyAs' options
        If DWFAddIn.HasSaveCopyAsOptions(oDoc, oContext, oOptions) Then

            oOptions.Value("Launch_Viewer") = False     'True o 1  /  False o 0

            ' Other options...
            'oOptions.Value("Publish_All_Component_Props") = 1
            'oOptions.Value("Publish_All_Physical_Props") = 1
            'oOptions.Value("Password") = 0

            If TypeOf oDoc Is DrawingDocument Then

                ' Drawing options
                oOptions.Value("Publish_Mode") = DWFPublishModeEnum.kCustomDWFPublish
                oOptions.Value("Publish_All_Sheets") = 1    '0

                ' The specified sheets will be ignored if
                ' the option "Publish_All_Sheets" is True (1)
                'Dim oSheets As NameValueMap
                'oSheets = Me.oAppCls.TransientObjects.CreateNameValueMap

                ' Publish the first sheet AND its 3D model
                'Dim oSheet1Options As NameValueMap
                'oSheet1Options = Me.oAppCls.TransientObjects.CreateNameValueMap

                'oSheet1Options.Add("Name", "Sheet:1")
                'oSheet1Options.Add("3DModel", True)
                'oSheets.Value("Sheet1") = oSheet1Options

                ' Publish the third sheet but NOT its 3D model
                'Dim oSheet3Options As NameValueMap
                'oSheet3Options = Me.oAppCls.TransientObjects.CreateNameValueMap

                'oSheet3Options.Add("Name", "Sheet3:3")
                'oSheet3Options.Add("3DModel", False)

                'oSheets.Value("Sheet2") = oSheet3Options

                'Set the sheet options object in the oOptions NameValueMap
                'oOptions.Value("Sheets") = oSheets
            End If

        End If

        'Set the destination file name
        oDataMedium.FileName = ficheroFIN
        Try
            If IO.File.Exists(ficheroFIN) Then IO.File.Delete(ficheroFIN)
            'Publish document.
            Call DWFAddIn.SaveCopyAs(oDoc, oContext, oOptions, oDataMedium)
        Catch ex As Exception
            Console.WriteLine("Error guardando " & ficheroFIN)
            Debug.Print("Error exportar DWG con " & ficheroFIN)
        End Try
    End Sub
End Class
