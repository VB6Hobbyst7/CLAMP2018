Option Compare Text

Imports Inventor
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports Microsoft.Win32
Imports System.Linq
Imports System.IO
Imports Microsoft.VisualBasic
Imports Microsoft.WindowsAPICodePack.Shell
'
Partial Public Class Inventor2acad
    Private DirDestino As String = ""           ' directorio destino donde vamos a copiar todos los ficheros
    '
    ' 0.- Primero pedir el directorio donde copiar todos los ficheros (Directorio_Elegir para rellenar "DirDestino")
    '
    ' 1.- Si todo es correcto. Llamar al formulario. Formulario_Crear
    '
    Public Sub PackAndGo_ElegirDirectorio()
        If oAppI.ActiveDocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then
            MsgBox("Utility only for Inventor Assembly...", MsgBoxStyle.Critical, "ATTENTION")
            Exit Sub
        End If
        '
        ' Elegir el directorio para copiar todos los ficheros.
        Dim oFb As New FolderBrowserDialog
Repite:
        oFb.Description = "Select destination directory (Different from the current IAM)"
        oFb.SelectedPath = IO.Path.GetDirectoryName(oAppI.ActiveDocument.FullFileName)
        oFb.ShowNewFolderButton = True
        Dim resultado As DialogResult = oFb.ShowDialog()
        If resultado = DialogResult.OK Then
            If oFb.SelectedPath = IO.Path.GetDirectoryName(oAppI.ActiveDocument.FullFileName) Then
                MsgBox("Destination directory must be different from the current IAM")
                GoTo Repite
            End If
            'dirDestino = IO.Path.Combine(oFb.SelectedPath, IO.Path.GetFileNameWithoutExtension(PathASM))
            DirDestino = oFb.SelectedPath
        Else
            Exit Sub
        End If

        If IO.Directory.Exists(DirDestino) = False Then
            MsgBox("Folder error or not exist...", MsgBoxStyle.Critical, "ATTENTION")
            Exit Sub
        Else
            Dim frmPg As New frmPackAndGo
            frmPg.oAsm = CType(oAppI.ActiveDocument, AssemblyDocument)
            frmPg.dirDestino = DirDestino
            frmPg.Show((New WindowWrapper(oAppI.MainFrameHWND)))
        End If
    End Sub
End Class
