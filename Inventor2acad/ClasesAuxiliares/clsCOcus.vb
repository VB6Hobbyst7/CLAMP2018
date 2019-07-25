Imports System
Imports Inventor
Imports System.Runtime.InteropServices
Imports Microsoft.Win32
'Imports System.Xml
'Imports System.Windows.Forms
Imports System.Threading
Imports System.Security.Permissions
Imports System.Collections
Imports System.Collections.Generic
Imports System.Linq
Imports Inventor2acad.Inventor2acad
''
Imports ua = UtilesAlberto
Imports uau = UtilesAlberto.Utiles
'
'
' ***** Para instanciar esta clase y crear el fichero CSV con los Componentes de la Representación actual.
'
'Dim oAsm As AssemblyDocument = CType(oApp.ActiveDocument, AssemblyDocument)
'Dim cCos As New clsCOcus(oAsm)
'If cCos IsNot Nothing Then
' cCos.ListaTotales
'End If

'If IO.File.Exists(cCos.fullFiCSV) Then
' If MsgBox("Open " & cCos.fullFiCSV, MsgBoxStyle.OkCancel, "Open Excel file") = MsgBoxResult.Ok Then
'  Process.Start(cCos.fullFiCSV)
' End If
'End If
'
Public Class clsCOcus
    Public oApp As Inventor.Application
    Public dicCOcu As SortedDictionary(Of String, clsCOcu) = Nothing
    Public lisCOcu As List(Of ComponentOccurrence) = Nothing
    Public oIam As AssemblyDocument = Nothing
    Public fullFiCSV As String = ""   'IO.Path.ChangeExtension(Me.oIam.FullFileName, "BOM.csv")
    Public ActiveRep As String = ""
    '
    Public Sub New(oAp As Inventor.Application, oI As AssemblyDocument)
        oApp = oAp
        lisCOcu = New List(Of ComponentOccurrence)
        Me.oIam = oI
        If oI.ComponentDefinition.Occurrences.Count > 0 Then
            PonCOcu_Recursivo(oI.ComponentDefinition.Occurrences)
        End If
        If lisCOcu.Count > 0 Then
            dicCOcu_Rellena()
        End If
        ActiveRep = Me.oIam.ComponentDefinition.RepresentationsManager.ActiveLevelOfDetailRepresentation.Name
        fullFiCSV = IO.Path.ChangeExtension(Me.oIam.FullFileName, ActiveRep.Trim & ".csv")
    End Sub
    Public Sub dicCOcu_Rellena()
        dicCOcu = New SortedDictionary(Of String, clsCOcu)
        For Each oCu As ComponentOccurrence In lisCOcu
            Dim cCu As New clsCOcu(oApp, oCu)
            If dicCOcu.ContainsKey(cCu.Name) Then
                dicCOcu(cCu.Name).Count += 1
            Else
                dicCOcu.Add(cCu.Name, cCu)
            End If
            cCu = Nothing
        Next
    End Sub
    '
    Public Sub PonCOcu_Recursivo(oCus As ComponentOccurrences)
        For Each oCu As ComponentOccurrence In oCus
            Try
                ' No incluir la pieza que termina en "_wireframe.ipt"
                If CType(oCu.ContextDefinition.Document, Inventor.Document).FullFileName.Contains("_wireframe") = True Or oCu.Name.Contains("_wireframe") = True Then
                    Continue For
                End If
                ' No incluir las que esten desactivadas o invisibles
                If oCu.Enabled = False Or oCu.Visible = False Or oCu.Suppressed = True Then
                    Continue For
                End If
            Catch ex As Exception
                Continue For
            End Try
            '
            lisCOcu.Add(oCu)
            If oCu.SubOccurrences IsNot Nothing AndAlso oCu.SubOccurrences.Count > 0 Then
                PonCOcu_Recursivo(oCu.SubOccurrences)
            End If
        Next
    End Sub
    '
    Public Sub ListaTotales()
        Dim mensaje As String = ""
        Dim sep As String = ";"   ' "," '";"
        Dim comi As String = Chr(34)
        Dim cabeceras As String = "ITEM;NAME;COUNT;MASS;TOTAL MASS;STRUCTURE;FULLPATH" & vbCrLf
        ' Cerrar Excel, si estaba abierta.
        For Each oPro As Process In Process.GetProcessesByName("excel")
            oPro.Kill()
        Next
        IO.File.WriteAllText(fullFiCSV, cabeceras, Text.Encoding.UTF8)
        '
        Dim contador As Integer = 1
        For Each key As String In dicCOcu.Keys
            Dim cCu As clsCOcu = dicCOcu(key)
            'Dim linea As String =
            '    contador.ToString & sep &
            '    IIf(IsNumeric(cCu.Name), comi & cCu.Name & comi, cCu.Name) & sep &
            '    cCu.Count & sep &
            '    cCu.Mass & sep &
            '    Math.Round(cCu.Mass * cCu.Count, 3) & sep &
            '    cCu.BOME.ToString.Replace("BOMStructure", "").Substring(1) & sep &
            '    cCu.File & vbCrLf
            Dim linea As String =
                contador.ToString & sep &
                IIf(IsNumeric(cCu.Name), "=" & comi & cCu.Name & comi, comi & cCu.Name & comi) & sep &
                cCu.Count & sep &
                cCu.Mass & sep &
                Math.Round(cCu.Mass * cCu.Count, 3) & sep &
                comi & cCu.BOME.ToString.Replace("BOMStructure", "").Substring(1) & comi & sep &
                comi & cCu.File & comi & vbCrLf
            '
            'If contador < dicCOcu.Keys.Count Then
            '    linea &= vbCrLf
            'End If
            '
            IO.File.AppendAllText(fullFiCSV, linea, Text.Encoding.UTF8)
            cCu = Nothing
            linea = ""
            contador += 1
        Next
        ' Poner totales
        Dim tCount As Integer = (From t As clsCOcu In dicCOcu.Values
                                 Where t.Count > 0
                                 Select t.Count).Sum()

        Dim tMass As Double = (From t As clsCOcu In dicCOcu.Values
                               Where (t.Mass * t.Count) > 0
                               Select (t.Mass * t.Count)).Sum()

        Dim totales As String = sep & comi & "TOTALS" & comi & sep & tCount.ToString & sep & sep & Math.Round(tMass, 3).ToString & sep & sep '& vbCrLf
        IO.File.AppendAllText(fullFiCSV, totales, Text.Encoding.UTF8)
    End Sub
End Class

Public Class clsCOcu
    Public Name As String = ""
    Public NameFull As String = ""
    Public File As String = ""
    Public FileName As String = ""
    Public Tipo As TipoCom = Nothing
    Public BOME As BOMStructureEnum = Nothing
    Public Mass As Double = -1
    Public Count As Integer = -1

    Public Sub New(oAp As Inventor.Application, oCo As ComponentOccurrence)
        NameFull = oCo.Name
        Name = NameFull.Split(":"c)(0)
        File = CType(oCo.Definition.Document, Inventor.Document).FullFileName
        FileName = IO.Path.GetFileNameWithoutExtension(File)
        Count = 1
        Mass = Math.Round(oCo.MassProperties.Mass, 3)
        '
        Me.BOME = oCo.Definition.BOMStructure
        If oCo.DefinitionDocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
            Me.Tipo = TipoCom.Assembly
        ElseIf oCo.DefinitionDocumentType = DocumentTypeEnum.kPartDocumentObject Then
            If CType(oCo.Definition.Document, PartDocument).ComponentDefinition.IsContentMember Then
                Me.Tipo = TipoCom.CC
                If clsI Is Nothing Then clsI = New Inventor2acad(oAp)
                If FileName.Contains("·") Then
                    clsI.PropiedadEscribeDesignTracking(oCo.Definition.Document, "Part Number", FileName.Split("·")(0))
                    CType(oCo.Definition.Document, Inventor.Document).Save2()
                End If
            ElseIf CType(oCo.Definition.Document, PartDocument).SubType = "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}" Then
                Me.Tipo = TipoCom.SheetMetal
            End If
        End If
    End Sub
End Class

Public Enum TipoCom
    Assembly
    Part
    SheetMetal
    CC
End Enum
