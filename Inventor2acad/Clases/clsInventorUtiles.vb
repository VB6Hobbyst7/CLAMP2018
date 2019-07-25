Option Compare Text

Imports Inventor
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports Microsoft.Win32
Imports System.Linq
Imports System.IO
Imports Microsoft.VisualBasic
'Imports Microsoft.WindowsAPICodePack.Shell
Partial Public Class Inventor2acad
    Public Function DistanciaDame(oPt1 As Point, oPt2 As Point, Optional distEn As IEnum.DistancieEn = IEnum.DistancieEn.real) As Double
        Dim resultado As Double
        Select Case distEn
            Case IEnum.DistancieEn.real
                resultado = oAppI.MeasureTools.GetMinimumDistance(oPt1, oPt2)   ' oPt1.DistanceTo(oPt2)
            Case IEnum.DistancieEn.X
                Dim oPt2T As Point = oTg.CreatePoint(oPt2.X, oPt1.Y, oPt1.Z)
                resultado = oAppI.MeasureTools.GetMinimumDistance(oPt1, oPt2T)
            Case IEnum.DistancieEn.Y
                Dim oPt2T As Point = oTg.CreatePoint(oPt1.X, oPt2.Y, oPt1.Z)
                resultado = oAppI.MeasureTools.GetMinimumDistance(oPt1, oPt2T)
            Case IEnum.DistancieEn.Z
                Dim oPt2T As Point = oTg.CreatePoint(oPt1.X, oPt1.Y, oPt2.Z)
                resultado = oAppI.MeasureTools.GetMinimumDistance(oPt1, oPt2T)
        End Select
        '
        Return resultado
    End Function
    Public Function DistanciaDame(ByVal x1 As Double, ByVal y1 As Double, ByVal z1 As Double, ByVal x2 As Double, ByVal y2 As Double, ByVal z2 As Double) As Double
        Return (((x2 - x1) ^ 2) + ((y2 - y1) ^ 2) + ((z2 - z1) ^ 2)) ^ 0.5
        Exit Function
    End Function

    Public Function GraRad(ByVal queGra As Double) As Double
        Return (queGra * Math.PI) / 180
    End Function

    Public Function RadGra(ByVal queRad As Double) As Double
        Return (queRad * 180) / Math.PI
    End Function

    ' PackAndGo tiene que ejecutarse fuera del proceso Inventor (Usa Apprentice)
    'Public Sub EmpaquetadoInventor(dirDestino As String, Optional oIam As Inventor.AssemblyDocument = Nothing)
    '    If oIam Is Nothing And oAppI.ActiveDocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
    '        oIam = TryCast(oAppI.ActiveDocument, Inventor.AssemblyDocument)
    '    End If
    '    '
    '    ' Crear objeto oPacknGoComp
    '    Dim oPacknGoComp As New PackAndGoLib.PackAndGoComponent
    '    '
    '    ' Crear objeto oPacknGo
    '    Dim oPacknGo As PackAndGoLib.PackAndGo
    '    oPacknGo = oPacknGoComp.CreatePackAndGo(oIam.FullFileName, dirDestino)
    '    '
    '    ' Fichero de proyecto de Inventor. Por defecto es el activo.
    '    oPacknGo.ProjectFile = oAppI.DesignProjectManager.ActiveDesignProject.FullFileName
    '    Dim sRefFiles = New String() {}
    '    Dim sMissFiles = New Object
    '    '
    '    ' Opciones de empaquetado
    '    oPacknGo.SkipLibraries = True
    '    oPacknGo.SkipStyles = True
    '    oPacknGo.SkipTemplates = True
    '    oPacknGo.CollectWorkgroups = False
    '    oPacknGo.KeepFolderHierarchy = True
    '    oPacknGo.IncludeLinkedFiles = True
    '    '
    '    ' Cargar todas las referencias en sRefFiles y en sMissFiles los no encontrados
    '    oPacknGo.SearchForReferencedFiles(sRefFiles, sMissFiles)
    '    '
    '    ' Añadir ficheros de referencia al paquete
    '    oPacknGo.AddFilesToPackage(sRefFiles)
    '    '
    '    ' Iniciar el empaquetado (False=Sin sobrescribir, True=sobrescribir)
    '    oPacknGo.CreatePackage(False)
    'End Sub
End Class
