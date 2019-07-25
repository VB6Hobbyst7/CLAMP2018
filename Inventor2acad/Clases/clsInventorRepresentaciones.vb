Option Compare Text

Imports Inventor
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports Microsoft.Win32
Imports System.Linq
''Imports System.IO
Imports Microsoft.VisualBasic
Imports Microsoft.WindowsAPICodePack.Shell
Imports System.Collections.Generic
Partial Public Class Inventor2acad

    Public Sub Representation_CreateSubstitute_ALL(ByVal folderEnd As String, Optional insert As Boolean = True)
        ' Set a reference to the active assembly document
        Dim oDocASM As AssemblyDocument = oAppI.ActiveDocument
        Dim oDefASM As AssemblyComponentDefinition = oDocASM.ComponentDefinition
        If folderEnd = "" OrElse IO.Directory.Exists(folderEnd) = False Then
            folderEnd = IO.Path.GetDirectoryName(oDocASM.FullFileName)
        End If
        '


        ' Create a new part document that will be the shrinkwrap substitute
        Dim oPartDoc As PartDocument = oAppI.Documents.Add(DocumentTypeEnum.kPartDocumentObject, , False)
        Dim oPartDef As PartComponentDefinition = oPartDoc.ComponentDefinition

        Dim oDerivedAssemblyDef As DerivedAssemblyDefinition = oPartDef.ReferenceComponents.DerivedAssemblyComponents.CreateDefinition(oDocASM.FullDocumentName)

        ' Set various shrinkwrap related options
        oDerivedAssemblyDef.DeriveStyle = DerivedComponentStyleEnum.kDeriveAsSingleBodyNoSeams
        oDerivedAssemblyDef.IncludeAllTopLevelWorkFeatures = DerivedComponentOptionEnum.kDerivedExcludeAll
        oDerivedAssemblyDef.IncludeAllTopLevelSketches = DerivedComponentOptionEnum.kDerivedExcludeAll ' = kDerivedIncludeAll
        oDerivedAssemblyDef.IncludeAllTopLeveliMateDefinitions = DerivedComponentOptionEnum.kDerivedExcludeAll ' = kDerivedExcludeAll
        oDerivedAssemblyDef.IncludeAllTopLevelParameters = DerivedComponentOptionEnum.kDerivedExcludeAll   ' = kDerivedExcludeAll
        oDerivedAssemblyDef.ReducedMemoryMode = True
        oDerivedAssemblyDef.RemoveInternalVoids = True
        oDerivedAssemblyDef.UseColorOverridesFromSource = True

        Call oDerivedAssemblyDef.SetHolePatchingOptions(DerivedHolePatchEnum.kDerivedPatchAll)
        'Call oDerivedAssemblyDef.SetRemoveByVisibilityOptions(kDerivedRemovePartsAndFaces, 25)

        ' Create the shrinkwrap component
        Dim oDerivedAssembly As DerivedAssemblyComponent = oPartDef.ReferenceComponents.DerivedAssemblyComponents.Add(oDerivedAssemblyDef)
        oDerivedAssembly.BreakLinkToFile()
        '
        ' Ocultar WorkFeatures
        Dim oWa As WorkAxis
        For Each oWa In oPartDef.WorkAxes
            oWa.Visible = False
        Next
        '
        Dim oWpl As WorkPlane
        For Each oWpl In oPartDef.WorkPlanes
            oWpl.Visible = False
        Next
        '
        Dim oWpt As WorkPoint
        For Each oWpt In oPartDef.WorkPoints
            oWpt.Visible = False
        Next
        '
        oWa = Nothing
        oWpl = Nothing
        oWpt = Nothing
        '
        ' Save the part
        Dim strSubstituteFileName As String
        If folderEnd <> "" AndAlso IO.Directory.Exists(folderEnd) Then
            strSubstituteFileName = IO.Path.Combine(folderEnd, IO.Path.GetFileNameWithoutExtension(oDocASM.FullFileName) & "_SS.ipt")
        Else
            strSubstituteFileName = IO.Path.Combine(
                IO.Path.GetDirectoryName(oDocASM.FullFileName), IO.Path.GetFileNameWithoutExtension(oDocASM.FullFileName) & "_SS.ipt")
            'strSubstituteFileName = Left$(oDoc.FullFileName, Len(oDoc.FullFileName) - 4)
            'strSubstituteFileName = strSubstituteFileName & "_SS.ipt"
        End If

        oAppI.SilentOperation = True
        Call oPartDoc.SaveAs(strSubstituteFileName, False)

        If insert = False Then
            ' Create a substitute level of detail using the shrinkwrap part.
            Dim oSubstituteLOD As LevelOfDetailRepresentation = Nothing
            oSubstituteLOD = oDefASM.RepresentationsManager.LevelOfDetailRepresentations.AddSubstitute(strSubstituteFileName, IO.Path.GetFileNameWithoutExtension(strSubstituteFileName), True)
            oSubstituteLOD.Activate(True)
            ' Release reference of the invisibly opened part document.
            oPartDoc.ReleaseReference()
        ElseIf insert = True Then
            ' Borrar todos los componentes que tenga
            Dim oCo As ComponentOccurrence
            For Each oCo In oDefASM.Occurrences
                oCo.Delete()
            Next
            '
            ' Insertar el componente generado
            Dim oMat As Matrix = oAppI.TransientGeometry.CreateMatrix
            Dim newOcc As ComponentOccurrence = oDefASM.Occurrences.Add(strSubstituteFileName, oMat)
            newOcc.Grounded = False
        End If
        If oDocASM.RequiresUpdate Then oDocASM.Update2()
        If oDocASM.Dirty Then oDocASM.Save2()
        oAppI.SilentOperation = False
    End Sub

    Public Sub Representation_CreateSubstitute_LEVELONE(ByVal FolderEnd As String, ByVal propBusco As String)
        ' Set a reference to the active assembly document
        Dim oDocASM As AssemblyDocument = oAppI.ActiveDocument
        Dim oDefASM As AssemblyComponentDefinition = oDocASM.ComponentDefinition
        Dim newPartName As String = ""
        If FolderEnd = "" OrElse IO.Directory.Exists(FolderEnd) = False Then
            FolderEnd = IO.Path.GetDirectoryName(oDocASM.FullFileName)
        End If
        '
        Call oAppI.CommandManager.ControlDefinitions.Item("PartHideAllOccurrenceWorkFeaturesCtxCmd").Execute2(True)

        oAppI.SilentOperation = True
        '
        ' Recorrer subensamblajes del nivel 1
        Dim oCo As ComponentOccurrence
        For Each oCo In oDefASM.Occurrences
            ' Solo procesamos los ensamblajes
            If oCo.DefinitionDocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                Dim oDef1 As AssemblyComponentDefinition = oCo.Definition
                Dim oDoc1 As AssemblyDocument = oDef1.Document
                Dim fullPathOccu As String = oDoc1.FullFileName
                newPartName = oCo.Name.Replace(":", "·") & "·" & FolderEnd & "·" & Guid.NewGuid.ToString("N").Substring(0, 5) & ".ipt"
                ' Repetir mientras exista el fichero.
                While IO.File.Exists(newPartName) = True
                    newPartName = oCo.Name.Replace(":", "·") & "·" & FolderEnd & "·" & Guid.NewGuid.ToString("N").Substring(0, 5) & ".ipt"
                End While
                '
                If propBusco <> "" Then
                    Dim pFlag As String = ""
                    Try
                        ' Solo lo convertimos en Piezas si propBusco = True
                        pFlag = PropiedadLeeUsuario(oDoc1, propBusco)
                        If pFlag.ToUpper = "FALSE" Or pFlag = "" Then
                            Continue For
                        End If
                    Catch ex As Exception
                        Continue For
                    End Try
                End If
                '
                ' Create a new part document that will be the shrinkwrap substitute
                Dim oPartDoc As PartDocument = oAppI.Documents.Add(DocumentTypeEnum.kPartDocumentObject, , False)
                Dim oPartDef As PartComponentDefinition = oPartDoc.ComponentDefinition

                Dim oDerivedAssemblyDef As DerivedAssemblyDefinition =
                        oPartDef.ReferenceComponents.DerivedAssemblyComponents.CreateDefinition(oDocASM.FullFileName)
                ' Set various shrinkwrap related options
                oDerivedAssemblyDef.DeriveStyle = DerivedComponentStyleEnum.kDeriveAsSingleBodyNoSeams
                oDerivedAssemblyDef.IncludeAllTopLevelWorkFeatures = DerivedComponentOptionEnum.kDerivedExcludeAll ' = kDerivedIncludeAll
                oDerivedAssemblyDef.IncludeAllTopLevelSketches = DerivedComponentOptionEnum.kDerivedExcludeAll ' = kDerivedIncludeAll
                oDerivedAssemblyDef.IncludeAllTopLeveliMateDefinitions = DerivedComponentOptionEnum.kDerivedExcludeAll ' = kDerivedExcludeAll
                oDerivedAssemblyDef.IncludeAllTopLevelParameters = DerivedComponentOptionEnum.kDerivedExcludeAll   ' = kDerivedExcludeAll
                oDerivedAssemblyDef.ReducedMemoryMode = True
                oDerivedAssemblyDef.RemoveInternalVoids = True
                oDerivedAssemblyDef.UseColorOverridesFromSource = True
                oDerivedAssemblyDef.SetHolePatchingOptions(DerivedHolePatchEnum.kDerivedPatchAll)
                'Call oDerivedAssemblyDef.SetRemoveByVisibilityOptions(kDerivedRemovePartsAndFaces, 25)

                ' Excluir todos menos el ensamblaje a simplificar
                Dim oDOco As DerivedAssemblyOccurrence
                For Each oDOco In oDerivedAssemblyDef.Occurrences
                    Dim fullPathOccuD As String = CType(oDOco.ReferencedOccurrence.ReferencedDocumentDescriptor.ReferencedDocument, Inventor.Document).FullFileName
                    If fullPathOccuD <> fullPathOccu Then
                        oDOco.InclusionOption = DerivedComponentOptionEnum.kDerivedExcludeAll
                    End If
                    'oDOco.InclusionOption = DerivedComponentOptionEnum.kDerivedIndividualDefined
                    'Else
                    '    oDOco.InclusionOption = DerivedComponentOptionEnum.kDerivedExcludeAll
                    'End If
                Next
                '

                ' Create the shrinkwrap component
                Dim oDerivedAssembly As DerivedAssemblyComponent = oPartDef.ReferenceComponents.DerivedAssemblyComponents.Add(oDerivedAssemblyDef)
                oDerivedAssembly.BreakLinkToFile()
                '
                ' Ocultar WorkFeatures
                Dim oWa As WorkAxis
                For Each oWa In oPartDef.WorkAxes
                    oWa.Visible = False
                Next
                '
                Dim oWpl As WorkPlane
                For Each oWpl In oPartDef.WorkPlanes
                    oWpl.Visible = False
                Next
                '
                Dim oWpt As WorkPoint
                For Each oWpt In oPartDef.WorkPoints
                    oWpt.Visible = False
                Next
                '
                oWa = Nothing
                oWpl = Nothing
                oWpt = Nothing
                '
                ' Save the part
                Dim strSubstituteFileName As String
                If FolderEnd <> "" AndAlso IO.Directory.Exists(FolderEnd) Then
                    'strSubstituteFileName = IO.Path.Combine(FolderEnd, IO.Path.GetFileNameWithoutExtension(oDoc1.FullFileName) & "_SS.ipt")
                    strSubstituteFileName = IO.Path.Combine(FolderEnd, newPartName)
                Else
                    '    strSubstituteFileName = IO.Path.Combine(
                    'IO.Path.GetDirectoryName(oDoc1.FullFileName), IO.Path.GetFileNameWithoutExtension(oDoc1.FullFileName) & "_SS" & contador.ToString & ".ipt")
                    strSubstituteFileName = IO.Path.Combine(IO.Path.GetDirectoryName(oDoc1.FullFileName), newPartName)
                    ''strSubstituteFileName = Left$(oDoc.FullFileName, Len(oDoc.FullFileName) - 4)
                    'strSubstituteFileName = strSubstituteFileName & "_SS.ipt"
                End If
                '' Si existe, nuevo nombre incrementando contador. Si no da error si está ya insertado y en uso
                If IO.File.Exists(strSubstituteFileName) Then
                    strSubstituteFileName = IO.Path.Combine(IO.Path.GetDirectoryName(strSubstituteFileName), newPartName)
                End If
                '
                Call oPartDoc.SaveAs(strSubstituteFileName, True)
                '
                ' Esto es sólo si queremos crear la nueva Representación--Nivel de Detalle
                ' Create a substitute level of detail using the shrinkwrap part.
                'Dim oSubstituteLOD As LevelOfDetailRepresentation =
                'oDef1.RepresentationsManager.LevelOfDetailRepresentations.AddSubstitute(strSubstituteFileName)
                'oSubstituteLOD.Activate()
                ' Release reference of the invisibly opened part document.
                'oPartDoc.ReleaseReference
                'oDoc1.Save2
                '
                Dim oMat As Matrix = oAppI.TransientGeometry.CreateMatrix
                Dim newOcc As ComponentOccurrence = oDefASM.Occurrences.Add(strSubstituteFileName, oMat)
                newOcc.Grounded = True
                oCo.Delete()
            ElseIf oCo.DefinitionDocumentType = DocumentTypeEnum.kPartDocumentObject Then
                Dim oDefP As PartComponentDefinition = oCo.Definition
                ' Ocultar WorkFeatures
                Dim oWa As WorkAxis
                For Each oWa In oDefP.WorkAxes
                    oWa.Visible = False
                Next
                '
                Dim oWpl As WorkPlane
                For Each oWpl In oDefP.WorkPlanes
                    oWpl.Visible = False
                Next
                '
                Dim oWpt As WorkPoint
                For Each oWpt In oDefP.WorkPoints
                    oWpt.Visible = False
                Next
            End If
        Next
        If oDocASM.RequiresUpdate Then
            oDocASM.Update2()
        End If
        If oDocASM.Dirty Then oDocASM.Save2()
        oAppI.SilentOperation = False
    End Sub
End Class
