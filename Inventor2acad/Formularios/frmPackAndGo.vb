Option Compare Text

Imports Inventor
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports Microsoft.Win32
Imports System.Linq
Imports System.IO
Imports Microsoft.VisualBasic
Imports System.IO.Compression
'Imports Microsoft.WindowsAPICodePack.Shell
Public Class frmPackAndGo
    Protected Friend oAsm As AssemblyDocument = Nothing
    Private oCd As AssemblyComponentDefinition = Nothing
    Protected Friend clsI As Inventor2acad = Nothing
    '
    Public PathASM As String = ""
    Public dirDestino As String = ""
    Private colFi As Dictionary(Of String, String)  ' Dictionary con los fullPath viejos y nuevos (Solo de los que se van a copiar)
    Private activado As Boolean = False
    Private Revincular As Boolean = False
    Private comprimir As Boolean = True

    Private Sub frmInicio_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        colFi = New Dictionary(Of String, String)
        PathASM = oAsm.FullFileName
        oCd = oAsm.ComponentDefinition
    End Sub
    '
    Private Sub btnStart_Click(sender As Object, e As EventArgs) Handles btnStart.Click
        ' 1.- Guardar como, el ensamblaje actual en el destino.
        oAsm.SaveAs(IO.Path.Combine(dirDestino, oAsm.FullFileName), False)
        ' 
        ' 2.- Recorrer todos los subcomponentes y borrar los que este Excluidos en la representación actual.
        ' rellenar "colFi" solo con los que se copian (Incluidos y Visibles)
        ComponentOccurences_BorraExcluidos(oCd)
        '
        ' 3.- Activar Representación Principal y borrar el resto.

        CopiaFicheros()
        'If Revincular Then RehacerLinks_TODOS()
        'If comprimir Then
        '    Dim arrFi As String() = IO.Directory.GetFiles(dirDestino, "*.*", SearchOption.AllDirectories)
        '    Dim fiZip As String = dirDestino & ".zip"
        '    Dim ElZip As New Ionic.Zip.ZipFile
        '    Using ElZip
        '        'ElZip.AddFiles(colFi.Values.ToArray)
        '        ElZip.AddFiles(arrFi)
        '        ElZip.Save(fiZip)
        '    End Using
        'End If
    End Sub
    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub
    '
    Private Sub ComponentOccurences_BorraExcluidos(queCd As AssemblyComponentDefinition)
        For Each oCo As ComponentOccurrence In oCd.Occurrences
            If oCo.Excluded = True OrElse oCo.Visible = False Then
                oCo.Delete2(True)
            Else
                Dim PathActual As String = CType(oCo.Definition.Document, Inventor.Document).FullFileName
                Dim PathNuevo As String = IO.Path.Combine(dirDestino, IO.Path.GetFileName(PathActual))
                '
                colFi.Add(PathActual, PathNuevo)
                CType(oCo.Definition.Document, Inventor.Document).SaveAs(PathNuevo, False)
                If oCo.DefinitionDocumentType = DocumentTypeEnum.kAssemblyDocumentObject AndAlso oCo.SubOccurrences IsNot Nothing AndAlso oCo.SubOccurrences.Count > 0 Then
                    ComponentOccurences_BorraExcluidos(CType(oCo.Definition.Document, Inventor.AssemblyDocument).ComponentDefinition)
                End If
            End If
        Next
    End Sub

    Private Sub Representaciones_BorraTodas()
        oAsm.Save2()
        ' Activar representación Principal
        oAsm.ComponentDefinition.RepresentationsManager.LevelOfDetailRepresentations.Item(1).Activate()
        '
        For Each oRep As LevelOfDetailRepresentation In oAsm.ComponentDefinition.RepresentationsManager.LevelOfDetailRepresentations
            If oRep.LevelOfDetail = LevelOfDetailEnum.kCustomLevelOfDetail Then
                oRep.Delete()
            End If
        Next
        oAsm.Save2()
    End Sub

    Private Sub RehacerLinks_TODOS()
        If colFi Is Nothing OrElse colFi.Count = 0 Then Exit Sub
        '
        Dim contador As Integer = 0
        lblDatos.Text = "Remaking links ---> " & contador.ToString & " of " & colFi.Count
        pb1.Value = 0 : pb1.Maximum = colFi.Count
        Dim extensiones() As String = New String() {".iam", ".ipt", ".idw", ".dwg", ".ipn"}
        Dim apprenticeServer As Inventor.ApprenticeServerComponent = New Inventor.ApprenticeServerComponent
        For Each queFi As String In colFi.Keys
            Dim fiDestino As String = colFi(queFi)
            contador += 1
            lblDatos.Text = "Remaking links ---> " & contador.ToString & " of " & colFi.Count
            If pb1.Value <= pb1.Maximum Then pb1.Value += 1
            ' No revincular los ficheros que no son de Inventor.
            If extensiones.Contains(IO.Path.GetExtension(fiDestino)) = False Then Continue For
            '
            Dim oDoc As Inventor.ApprenticeServerDocument = apprenticeServer.Open(fiDestino)
            ' Remplazar Documentos
            For Each oD As DocumentDescriptor In oDoc.ReferencedDocumentDescriptors
                If colFi.ContainsKey(oD.ReferencedFileDescriptor.FullFileName) Then
                    oD.ReferencedFileDescriptor.ReplaceReference(colFi(oD.ReferencedFileDescriptor.FullFileName))
                End If
            Next
            ' Remplazar OLEFileDescriptor
            For Each refOle As ReferencedOLEFileDescriptor In oDoc.ReferencedOLEFileDescriptors
                If colFi.ContainsKey(refOle.FullFileName) Then
                    refOle.FileDescriptor.ReplaceReference(colFi(refOle.FullFileName))
                End If
            Next
            oDoc.Close()
        Next
        apprenticeServer.Close()
        apprenticeServer = Nothing
        '
        lblDatos.Text &= " (TERMINADO)"
        btnClose.Enabled = True
        btnStart.Enabled = False
        Me.Cursor = Cursors.Default
        If cbOpen.Checked Then
            Process.Start(dirDestino)
            SendKeys.Send("{F5}")
            btnClose_Click(Nothing, Nothing)
        End If
    End Sub

    Private Sub Rellena_colFi_Recursivo(oDoc As ApprenticeServerDocument)
        ' Añadir el documento, el plano (Si tiene), los hijos y recursivamente recorrer los hijos.
        '
        ' El documento padre
        Dim fullAhora As String = oDoc.FullFileName
        Dim fullLuego As String = IO.Path.Combine(dirDestino, IO.Path.GetFileName(fullAhora))
        If colFi.ContainsKey(fullAhora) = False Then colFi.Add(fullAhora, fullLuego)
        '
        ' Si tiene IDW
        Dim planoIDW As String = IO.Path.ChangeExtension(fullAhora, ".idw")
        If IO.File.Exists(planoIDW) AndAlso colFi.ContainsKey(planoIDW) = False Then
            fullLuego = IO.Path.Combine(dirDestino, IO.Path.GetFileName(planoIDW))
            colFi.Add(planoIDW, fullLuego)
        End If
        ' Si tiene DWG
        Dim planoDWG As String = IO.Path.ChangeExtension(fullAhora, ".dwg")
        If IO.File.Exists(planoDWG) AndAlso colFi.ContainsKey(planoDWG) = False Then
            fullLuego = IO.Path.Combine(dirDestino, IO.Path.GetFileName(planoDWG))
            colFi.Add(planoDWG, fullLuego)
        End If
        ' Si tiene IPN
        Dim ficheroIPN As String = IO.Path.ChangeExtension(fullAhora, ".ipn")
        If IO.File.Exists(ficheroIPN) AndAlso colFi.ContainsKey(ficheroIPN) = False Then
            fullLuego = IO.Path.Combine(dirDestino, IO.Path.GetFileName(ficheroIPN))
            colFi.Add(ficheroIPN, fullLuego)
        End If
        ' Archivos referenciados en el documento hijo
        For Each oRef As ReferencedOLEFileDescriptor In oDoc.ReferencedOLEFileDescriptors
            If colFi.ContainsKey(oRef.FullFileName) = False Then
                fullLuego = IO.Path.Combine(dirDestino, IO.Path.GetFileName(oRef.FullFileName))
                colFi.Add(oRef.FullFileName, fullLuego)
            End If
        Next
        '
        ' Componentes Hijos
        Dim oDocs As ApprenticeServerDocuments = oDoc.AllReferencedDocuments
        If oDocs IsNot Nothing AndAlso oDocs.Count > 0 Then
            For Each oDocH As ApprenticeServerDocument In oDocs
                Rellena_colFi_Recursivo(oDocH)
            Next
        End If
    End Sub
    '
    Private Sub CopiaFicheros()
        Me.Cursor = Cursors.AppStarting
        If pb1.Value <= pb1.Maximum Then pb1.Value += 1
        Dim contador As Integer = 1
        Dim total As Integer = pb1.Maximum
        ' Copiar la plantilla
        lblDatos.Text = "PackAndGo ---> " & contador.ToString & " of " & total & " Files"
        Dim ipjFin As String = IO.Path.Combine(dirDestino, "_" & IO.Path.GetFileNameWithoutExtension(PathASM) & ".ipj")
        Dim b As Byte() = My.Resources.Template
        My.Computer.FileSystem.WriteAllBytes(ipjFin, b, False)
        '
        ' El resto de ficheros
        For Each fiOri As String In colFi.Keys
            If pb1.Value <= pb1.Maximum Then pb1.Value += 1
            contador += 1
            lblDatos.Text = "PackAndGo ---> " & contador.ToString & " of " & total & " Files"
            Dim fiDes As String = colFi(fiOri)
            IO.File.Copy(fiOri, fiDes, True)
        Next
        ' Si no revinculamos. Final aquí y abrir carpeta
        If Revincular = False Then
            lblDatos.Text &= " (TERMINADO)"
            btnClose.Enabled = True
            btnStart.Enabled = False
            Me.Cursor = Cursors.Default
            If cbOpen.Checked Then
                Process.Start(dirDestino)
                SendKeys.Send("{F5}")
                btnClose_Click(Nothing, Nothing)
            End If
        End If
    End Sub
End Class


