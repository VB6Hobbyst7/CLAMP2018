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
    ' ***** OBJECTOS INVENTOR
    Public oAppI As Inventor.Application = Nothing
    Public WithEvents oAppIEv As Inventor.ApplicationEvents = Nothing
    Public oGN As Inventor.GraphicsNode = Nothing
    Public oRS As Inventor.RenderStyle = Nothing
    Public oTg As Inventor.TransientGeometry = Nothing
    Public oTo As Inventor.TransientObjects = Nothing
    Public oTBr As Inventor.TransientBRep = Nothing
    Public oSelSet As Inventor.SelectSet = Nothing
    Public oCm As Inventor.CommandManager = Nothing
    Private oMatrix As Matrix = Nothing
    Private oNopfs As NonParametricBaseFeatures = Nothing
    Private oNopf As NonParametricBaseFeature = Nothing
    Private oNopfd As NonParametricBaseFeatureDefinition = Nothing
    Public oIam As AssemblyDocument = Nothing
    Public oIpt As PartDocument = Nothing
    Public oIdw As DrawingDocument = Nothing
    Public oCo As Inventor.ComponentOccurrence = Nothing    '' ComponentOccurrence de la que sacamos las Faces.
    '
    ' ***** CLASES
    Public clsIp As clsiPictureToImage = Nothing    '' Inicializar al inicio
    Public clsCc As clsColorChange = Nothing        '' Inicializar sólo cuando se necesite (Con PartDocument)
    Public clsSe As clsSelectMulti = Nothing             '' Inicializar sólo cuando se necesite (Con oAppCls)
    '
    ' ***** CONSTANTES
    Public Const nivelDetalleDefecto As String = "Desactivados"
    Public Const nivelDetalleDefectoCompleto As String = "<Desactivados>"
    '
    ' ***** VARIABLES GENERALES
    Public WithEvents Timer1 As System.Windows.Forms.Timer = Nothing
    Public cadenaMensajes As String = ""
    Public dirProyectoInv As String = ""    '' Proyecto de Inventor que activaremos.
    Public accion As String = ""
    Public ptIAM As String = "Centro"   ' Nombre del origen del ensamblaje (punto 0,0,0)
    Public ptIAM1 As String = "Centro1"   ' Nombre de otro punto del ensamblaje (punto opuesto a ptIAM)
    Public Busquedabasica As Boolean = True
    Public arrDirsPro As ArrayList = Nothing
    '
    ' De clsComponentOccurrence
    '' Puntos minimos y maximos
    Public oBox As Box = Nothing        ''
    Public Ptmin As Point = Nothing     '' Punto mínimo del Box del ComponentOccurrence (Enamblaje Tubo)
    Public Ptmax As Point = Nothing     '' Punto máximo del Box...
    Public largo As Double = 0      '' Largo componente en X
    Public ancho As Double = 0      '' Ancho componente en Y
    Public alto As Double = 0       '' Alto componente en Z
    '' Coordenadas sueltas mínimas y máximas.
    Public minX As Double = 0
    Public minY As Double = 0
    Public minZ As Double = 0
    Public maxX As Double = 0
    Public maxY As Double = 0
    Public maxZ As Double = 0
    '
    ' ***** COLECCIONES De clsSurface
    Public colMatrix As System.Collections.Generic.List(Of Matrix)
    Public colP3D As System.Collections.Generic.List(Of SketchPoint3D)
    Public colSb As System.Collections.Generic.List(Of SurfaceBodyProxy)
    Public colFc As System.Collections.Generic.List(Of Face)    '' Faces exteriores (Cilindricas y Conicas)
    Public colFp As System.Collections.Generic.List(Of Face)    '' Faces exteriores (Planas)
    Public colFcProx As System.Collections.Generic.List(Of FaceProxy)    '' Faces exteriores (Cilindricas y Conicas)
    Public colFpProx As System.Collections.Generic.List(Of FaceProxy)    '' Faces exteriores (Planas)
    Public colFProxAll As System.Collections.Generic.List(Of FaceProxy)    '' Faces exteriores (Planas)
    '
    '***** DE FICHEROS DE CONFIGURACION
    Public app_folder As String = My.Application.Info.DirectoryPath     '' Solo Directorio
    Public app_name As String = My.Application.Info.AssemblyName        '' 
    Public app_folderandname As String = IO.Path.Combine(app_folder, app_name)        '' 
    Public app_folderandnameExt As String = IO.Path.Combine(app_folder, app_name & ".dll")
    Public app_version As String = My.Application.Info.Version.ToString
    Public app_nameandversion As String = app_name & " - v" & app_version
    Public app_log As String = app_folderandname & ".log"
    'Public app_ini As String = app_folderandname & ".ini"
    'Public app_config As String = app_folderandnameExt & ".config"
    Public Log As Boolean = True
    '
    Public Sub VaciaTodo()
        'If Not (oAppI Is Nothing) Then Marshal.ReleaseComObject(oAppI)
        oAppI = Nothing
        'If Not (oAppIEv Is Nothing) Then Marshal.ReleaseComObject(oAppIEv)
        oAppIEv = Nothing
        'If Not (oGN Is Nothing) Then Marshal.ReleaseComObject(oGN)
        oGN = Nothing
        'If Not (oRS Is Nothing) Then Marshal.ReleaseComObject(oRS)
        oRS = Nothing
        'If Not (oTg Is Nothing) Then Marshal.ReleaseComObject(oTg)
        oTg = Nothing
        'If Not (oTo Is Nothing) Then Marshal.ReleaseComObject(oTo)
        oTo = Nothing
        'If Not (oTBr Is Nothing) Then Marshal.ReleaseComObject(oTBr)
        oTBr = Nothing
        'If Not (oSelSet Is Nothing) Then Marshal.ReleaseComObject(oSelSet)
        oSelSet = Nothing
        'If Not (oCm Is Nothing) Then Marshal.ReleaseComObject(oCm)
        oCm = Nothing
        'If Not (oMatrix Is Nothing) Then Marshal.ReleaseComObject(oMatrix)
        oMatrix = Nothing
        'If Not (oNopfs Is Nothing) Then Marshal.ReleaseComObject(oNopfs)
        oNopfs = Nothing
        'If Not (oNopf Is Nothing) Then Marshal.ReleaseComObject(oNopf)
        oNopf = Nothing
        'If Not (oNopfd Is Nothing) Then Marshal.ReleaseComObject(oNopfd)
        oNopfd = Nothing
        'If Not (oIam Is Nothing) Then Marshal.ReleaseComObject(oIam)
        oIam = Nothing
        'If Not (oIpt Is Nothing) Then Marshal.ReleaseComObject(oIpt)
        oIpt = Nothing
        'If Not (oIdw Is Nothing) Then Marshal.ReleaseComObject(oIdw)
        oIdw = Nothing
        'If Not (oCo Is Nothing) Then Marshal.ReleaseComObject(oCo)
        oCo = Nothing
        '
        clsIp = Nothing    '' Inicializar al inicio
        clsCc = Nothing        '' Inicializar sólo cuando se necesite (Con PartDocument)
        clsSe = Nothing             '' Inicializar sólo cuando se necesite (Con oAppCls)
        '
        colMatrix = Nothing
        colP3D = Nothing
        colSb = Nothing
        colFc = Nothing
        colFp = Nothing
        colFcProx = Nothing
        colFpProx = Nothing
        colFProxAll = Nothing
        '
        System.GC.Collect()
        System.GC.WaitForPendingFinalizers()
        System.GC.Collect()
    End Sub
    Public Sub LimpiaMemoria()
        GC.WaitForPendingFinalizers()
        GC.Collect()
        GC.WaitForPendingFinalizers()
        GC.Collect()
    End Sub
    Public Sub PonLog(text As String, Optional borrar As Boolean = False)
        Try
            If text.EndsWith(vbCrLf) = False Then text &= vbCrLf
            text = Date.Now.ToString & vbTab & text
            If borrar Then
                IO.File.WriteAllText(Me.app_log, text)
            Else
                IO.File.AppendAllText(Me.app_log, text)
            End If
        Catch ex As Exception
            ''
        End Try
    End Sub

    ''' <summary>
    ''' Si le damos una cadena completa (unidad:\directorio\fichero.extension) nos devuelve la parte que le indiquemos.
    ''' </summary>
    ''' <param name="queCamino">Cadena completa con el camino a procesar DIR+FICHERO+EXT</param>
    ''' <param name="queParte">Que queremos que nos devuelva</param>
    ''' <param name="queExtension">"" o extensión (Ej: ".bak"), si queremos cambiarla</param>
    ''' <returns>Retorna la cadena de texto con la opción indicada</returns>
    ''' <remarks></remarks>
    Public Function DameParteCamino(ByVal queCamino As String, Optional ByVal queParte As IEnum.ParteCamino = 0, Optional ByVal queExtension As String = "") As String
        Dim resultado As String = ""

        Select Case queParte
            Case 0  'ParteCamino.SoloCambiaExtension (dwg) Sin punto
                If queExtension <> "" And IO.Path.HasExtension(queCamino) Then
                    queCamino = IO.Path.ChangeExtension(queCamino, queExtension)
                End If
                resultado = queCamino
            Case 1  'ParteCamino.CaminoSinFichero
                resultado = IO.Path.GetDirectoryName(queCamino)
            Case 2  'ParteCamino.CaminoSinFicheroBarra
                resultado = IO.Path.GetDirectoryName(queCamino) & "\"
            Case 3  'ParteCamino.CaminoConFicheroSinExtension
                resultado = IO.Path.ChangeExtension(queCamino, Nothing)
            Case 4  'ParteCamino.CaminoConFicheroSinExtension
                resultado = IO.Path.ChangeExtension(queCamino, Nothing) & "\"
            Case 5  'ParteCamino.SoloFicheroConExtension
                resultado = IO.Path.GetFileName(queCamino)
            Case 6  'ParteCamino.SoloFicheroSinExtension
                resultado = IO.Path.GetFileNameWithoutExtension(queCamino)
            Case 7  'ParteCamino.SoloExtension
                resultado = IO.Path.GetExtension(queCamino)
            Case 8  'ParteCamino.SoloRaiz
                resultado = IO.Path.GetPathRoot(queCamino)
            Case 9  'ParteCamino.SoloNombreDirectorio
                Dim trozos() As String = queCamino.Split("\")
                resultado = trozos(trozos.GetUpperBound(0) - 1)
            Case 10  'ParteCamino.PenultimoDirectorioSinBarra
                Dim trozos() As String = queCamino.Split("\")
                If trozos.GetUpperBound(0) > 2 Then
                    Dim final(trozos.GetUpperBound(0) - 1) As String
                    Array.Copy(trozos, final, trozos.GetUpperBound(0) - 1)
                    resultado = String.Join("\", final)
                ElseIf trozos.GetUpperBound(0) > 1 Then
                    resultado = trozos(0)
                Else
                    resultado = "C:"
                End If
                If resultado.EndsWith("\") Then resultado = Mid(resultado, 1, resultado.Length - 1)
            Case 11  'ParteCamino.PenultimoDirectorioConBarra
                Dim trozos() As String = queCamino.Split("\")
                If trozos.GetUpperBound(0) > 1 Then
                    Dim final(trozos.GetUpperBound(0) - 2) As String
                    Array.Copy(trozos, final, trozos.GetUpperBound(0) - 1)
                    resultado = String.Join("\", final)
                ElseIf trozos.GetUpperBound(0) > 1 Then
                    resultado = trozos(0)
                Else
                    resultado = "C:"
                End If
                If resultado.EndsWith("\") = False Then resultado &= "\"
            Case 12  'ParteCamino.AntePenultimoDirectorioSinBarra
                Dim trozos() As String = queCamino.Split("\")
                If trozos.GetUpperBound(0) > 2 Then
                    Dim final(trozos.GetUpperBound(0) - 2) As String
                    Array.Copy(trozos, final, trozos.GetUpperBound(0) - 2)
                    resultado = String.Join("\", final)
                ElseIf trozos.GetUpperBound(0) > 1 Then
                    resultado = trozos(0)
                Else
                    resultado = "C:"
                End If
                If resultado.EndsWith("\") Then resultado = Mid(resultado, 1, resultado.Length - 1)
            Case 13  'ParteCamino.AntePenultimoDirectorioConBarra
                Dim trozos() As String = queCamino.Split("\")
                If trozos.GetUpperBound(0) > 2 Then
                    Dim final(trozos.GetUpperBound(0) - 2) As String
                    Array.Copy(trozos, final, trozos.GetUpperBound(0) - 2)
                    resultado = String.Join("\", final)
                ElseIf trozos.GetUpperBound(0) > 1 Then
                    resultado = trozos(0)
                Else
                    resultado = "C:"
                End If
                If resultado.EndsWith("\") = False Then resultado &= "\"
        End Select
        DameParteCamino = resultado
        Exit Function
    End Function

    Public Sub Retardo(ByVal segundos As Integer)
        Const NSPerSecond As Long = 10000000
        Dim ahora As Long = Date.Now.Ticks
        Debug.Print(Date.Now.Ticks)
        Do
            ' No hacemos nada
            'System.Windows.Forms.Application.DoEvents
        Loop While Date.Now.Ticks < ahora + (segundos * NSPerSecond)
        Debug.Print(Date.Now.Ticks)
    End Sub
End Class
'

Namespace IEnum
#Region "ENUMERACIONES"
    ''' <summary>
    ''' Diferentes vistas soportadas para visualizar una imagen
    ''' </summary>
    Public Enum TipoVista
        Small
        Medium
        Large
        ExtraLarge
    End Enum
    Public Enum FaceData
        Normal
        DireccionX
        DireccionY
        DireccionZ
        DireccionZMedio
    End Enum
    Public Enum ParteCamino
        SoloCambiaExtension = 0
        CaminoSinFichero = 1
        CaminoSinFicheroBarra = 2
        CaminoConFicheroSinExtension = 3
        CaminoConFicheroSinExtensionBarra = 4
        SoloFicheroConExtension = 5
        SoloFicheroSinExtension = 6
        SoloExtension = 7
        SoloRaiz = 8
        SoloNombreDirectorio = 9
        PenultimoDirectorioSinBarra = 10
        PenultimoDirectorioConBarra = 11
        AntePenultimoDirectorioSinBarra = 10
        AntePenultimoDirectorioConBarra = 11
    End Enum
    Public Enum nAtributo As Integer
        alineacion = 0
        punto1 = 1
        punto2 = 2
        cara_cabeza = 3
        alt_men = 4
        quePadre = 5
        pilar = 6       '(nombre del pilar)
        tirada = 7      ' Número de tirada de arriba a abajo
        seleccion = 8   ' pilar & cara & tirada
    End Enum
    Public Enum OperacionValor As Integer
        cambiar = 0
        sumar = 1
        restar = 2
    End Enum
    Public Enum GuardaTipo As Integer
        dwg = 0
        dxf = 1
        dwf = 2
        dwfx = 3
        pdf = 4
        bmp = 5
        gif = 6
        jpg = 7
        png = 8
        tiff = 9
    End Enum
    Public Enum EstructuraBOM
        '("Estructurado")   ("Sin nombre")  ("Sólo piezas")
        Estructurado = Inventor.BOMViewTypeEnum.kStructuredBOMViewType
        Piezas = Inventor.BOMViewTypeEnum.kPartsOnlyBOMViewType
        Modelo = Inventor.BOMViewTypeEnum.kModelDataBOMViewType
    End Enum
    Public Enum Datos
        EjeX
        EjeY
        EjeZ
        Origen
        PlanoXY
        PlanoXZ
        PlanoYZ
        X
        Y
        Z
        Ascending
        Descending
    End Enum
    Public Enum nombreColor
        Rojo
        Verde
        Azul
        Negro
        Blanco
    End Enum
    Public Enum DistancieEn
        X
        Y
        Z
        real
    End Enum
#End Region
End Namespace