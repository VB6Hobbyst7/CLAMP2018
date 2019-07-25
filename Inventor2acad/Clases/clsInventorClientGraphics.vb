Option Compare Text

Imports Inventor
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports Microsoft.Win32
Imports System.Linq
Imports System.IO
Imports Microsoft.VisualBasic
Imports Microsoft.WindowsAPICodePack.Shell
Partial Public Class Inventor2acad
    Public Const nombreCG As String = "CG2acad"
    ''draw the point graphics
    'Dim oCoordSet As GraphicsCoordinateSet = Nothing
    'Dim oGraphicsNode As GraphicsNode = Nothing
    'Dim oDataSets As GraphicsDataSets = Nothing

    ''get datasets, dataset, graphics node for client graphics
    'getCG(oGraphicsNode, oCoordSet, oDataSets)

    'oGraphicsNode.Selectable = True

    ''add surface Graphics
    'Dim oBody1Graphics As SurfaceGraphics
    'oBody1Graphics = oGraphicsNode.AddSurfaceGraphics(newBody1)
    'oBody1Graphics.ChildrenAreSelectable = True
    'For i = 1 To oBody1Graphics.DisplayedFaces.Count
    '   oBody1Graphics.DisplayedFaces.Item(i).Selectable = True
    'Next
    'oAppI.ActiveView.Update()
    Public Sub CG_Dame(ByRef oGraphicsNode As Object,
                     Optional ByRef oCoordSet As Object = Nothing,
                     Optional ByRef oOutDataSets As Object = Nothing)

        Dim oDoc As Document = oAppI.ActiveDocument

        Dim oDataOwner As Object = Nothing
        Dim oGraphicsOwner As Object = Nothing

        'check the document type and get the owner of the datasets and graphics
        If oDoc.DocumentType = DocumentTypeEnum.kPartDocumentObject Or oDoc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
            oDataOwner = oDoc
            oGraphicsOwner = oDoc.ComponentDefinition
        ElseIf oDoc.DocumentType = DocumentTypeEnum.kDrawingDocumentObject Then
            If oDoc.ActiveSheet Is Nothing Then
                MsgBox("The current document is a drawing. The command is supposed to draw client graphics on active sheet! But active sheet is null!")
                Exit Sub
            Else
                oDataOwner = oDoc.ActiveSheet
                oGraphicsOwner = oDoc.ActiveSheet
            End If
        End If

        'delete the data sets and graphics if they exist
        Try
            oDataOwner.GraphicsDataSetsCollection(nombreCG).Delete()
        Catch ex As Exception
        End Try

        Try
            oGraphicsOwner.ClientGraphicsCollection(nombreCG).Delete()
        Catch ex As Exception
        End Try

        'create DataSets 
        Dim oDataSets As GraphicsDataSets = oDataOwner.GraphicsDataSetsCollection.Add(nombreCG)
        oOutDataSets = oDataSets

        'create one coordinate data set
        oCoordSet = oDataSets.CreateCoordinateSet(oDataSets.Count + 1)

        'create graphics node
        Dim oClientGraphics As Inventor.ClientGraphics = oGraphicsOwner.ClientGraphicsCollection.Add(nombreCG)
        oGraphicsNode = oClientGraphics.AddNode(oClientGraphics.Count + 1)
    End Sub
    Public Sub CG_CrearFlecha3D()
        If oAppI.ActiveDocumentType <> DocumentTypeEnum.kAssemblyDocumentObject And oAppI.ActiveDocumentType <> DocumentTypeEnum.kPartDocumentObject Then
            Exit Sub
        End If
        '
        Dim oDoc As Document = oAppI.ActiveDocument
        ' Set a reference to component definition of the active document.
        ' This assumes that a part or assembly document is active
        Dim oCompDef As ComponentDefinition = oAppI.ActiveDocument.ComponentDefinition

        ' Check to see if the test graphics data object already exists.
        ' If it does clean up by removing all associated of the client 
        ' graphics from the document. If it doesn't create it
        On Error Resume Next
        Dim oClientGraphics As ClientGraphics = oCompDef.ClientGraphicsCollection.Item(nombreCG)
        If Err.Number = 0 Then
            On Error GoTo 0
            ' An existing client graphics object was successfully 
            ' obtained so clean up
            oClientGraphics.Delete()

            ' Update the display to see the results
            oAppI.ActiveView.Update()
        Else
            Err.Clear()
            On Error GoTo 0

            ' Set a reference to the transient geometry object 
            ' for user later
            Dim oTransGeom As TransientGeometry = oAppI.TransientGeometry

            ' Create the ClientGraphics object.
            oClientGraphics = oCompDef.ClientGraphicsCollection.Add(nombreCG)

            ' Create a new graphics node within the client graphics objects
            Dim oSurfacesNode As GraphicsNode = oClientGraphics.AddNode(1)

            Dim oTransientBRep As TransientBRep = oAppI.TransientBRep

            ' Create a point representing the center of the bottom of 
            ' the cone
            Dim oBottom As Point = oAppI.TransientGeometry.CreatePoint(0, 0, 0)

            ' Create a point representing the tip of the cone
            Dim oTop As Point = oAppI.TransientGeometry.CreatePoint(0, 10, 0)

            ' Create a transient cone body
            Dim oBody As SurfaceBody = oTransientBRep.CreateSolidCylinderCone(oBottom, oTop, 5, 5, 0)

            ' Reset the top point indicating the center of the top of 
            ' the cylinder
            oTop = oAppI.TransientGeometry.CreatePoint(0, -40, 0)

            ' Create a transient cylinder body
            Dim oCylBody As SurfaceBody = oTransientBRep.CreateSolidCylinderCone(oBottom, oTop, 2.5, 2.5, 2.5)

            ' Union the cone and cylinder bodies
            Call oTransientBRep.DoBoolean(oBody, oCylBody, BooleanTypeEnum.kBooleanTypeUnion)

            ' Create client graphics based on the transient body
            Dim oSurfaceGraphics As SurfaceGraphics = oSurfacesNode.AddSurfaceGraphics(oBody)

            ' Update the view to see the resulting curves
            oAppI.ActiveView.Update()
        End If
    End Sub
End Class
