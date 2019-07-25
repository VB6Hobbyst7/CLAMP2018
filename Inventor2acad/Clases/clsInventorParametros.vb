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
    Public Function ParametroLeeDouble(ByVal oDoc As Inventor.Document, ByVal quePar As String, Optional queF As String = "") As Double
        Dim resultado As Double = 0
        Dim oPar As Inventor.Parameter = Nothing
        '' Por si le damos fullFilename (queF) en vez de Document
        Dim estabaabierto As Boolean = True
        If oDoc Is Nothing AndAlso queF <> "" AndAlso IO.File.Exists(queF) = True Then
            estabaabierto = FicheroAbierto(queF)
            oAppI.SilentOperation = True
            If estabaabierto = True Then
                oDoc = oAppI.Documents.ItemByName(queF)
            Else
                oDoc = oAppI.Documents.Open(queF, False)
            End If
            oAppI.SilentOperation = False
        End If
        '' ***********************************************
        Try
            If oDoc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                oPar = CType(oDoc, AssemblyDocument).ComponentDefinition.Parameters.Item(quePar)
            ElseIf oDoc.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
                oPar = CType(oDoc, PartDocument).ComponentDefinition.Parameters.Item(quePar)
            End If
            resultado = oPar.Value
        Catch ex As Exception
            'MsgBox("Error ParametroASMLee. El parametro (" & quePar & ") no existe.")
        End Try
        If estabaabierto = False Then oDoc.Close(True)
        ParametroLeeDouble = resultado
    End Function
    Public Function ParametroLeeString(ByVal oDoc As Inventor.Document, ByVal quePar As String, Optional queF As String = "") As String
        Dim resultado As String = ""
        Dim oPar As Inventor.Parameter = Nothing
        '' Por si le damos fullFilename (queF) en vez de Document
        Dim estabaabierto As Boolean = True
        If oDoc Is Nothing AndAlso queF <> "" AndAlso IO.File.Exists(queF) = True Then
            estabaabierto = FicheroAbierto(queF)
            oAppI.SilentOperation = True
            If estabaabierto = True Then
                oDoc = oAppI.Documents.ItemByName(queF)
            Else
                oDoc = oAppI.Documents.Open(queF, False)
            End If
            oAppI.SilentOperation = False
        End If
        '' ***********************************************
        Try
            If oDoc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                oPar = CType(oDoc, AssemblyDocument).ComponentDefinition.Parameters.Item(quePar)
            ElseIf oDoc.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
                oPar = CType(oDoc, PartDocument).ComponentDefinition.Parameters.Item(quePar)
            End If
            resultado = oPar.Value
        Catch ex As Exception
            'MsgBox("Error ParametroASMLee. El parametro (" & quePar & ") no existe.")
        End Try
        If estabaabierto = False Then oDoc.Close(True)
        ParametroLeeString = resultado
    End Function

    Public Function ParametroLeeBoolean(ByVal oDoc As Inventor.Document, ByVal quePar As String, Optional queF As String = "") As Boolean
        Dim resultado As Boolean = False
        Dim oPar As Inventor.Parameter = Nothing
        '' Por si le damos fullFilename (queF) en vez de Document
        Dim estabaabierto As Boolean = True
        If oDoc Is Nothing AndAlso queF <> "" AndAlso IO.File.Exists(queF) = True Then
            estabaabierto = FicheroAbierto(queF)
            oAppI.SilentOperation = True
            If estabaabierto = True Then
                oDoc = oAppI.Documents.ItemByName(queF)
            Else
                oDoc = oAppI.Documents.Open(queF, False)
            End If
            oAppI.SilentOperation = False
        End If
        '' ***********************************************
        Try
            If oDoc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                oPar = CType(oDoc, AssemblyDocument).ComponentDefinition.Parameters.Item(quePar)
            ElseIf oDoc.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
                oPar = CType(oDoc, PartDocument).ComponentDefinition.Parameters.Item(quePar)
            End If
            resultado = oPar.Value
        Catch ex As Exception
            'MsgBox("Error ParametroASMLee. El parametro (" & quePar & ") no existe.")
        End Try
        If estabaabierto = False Then oDoc.Close(True)
        ParametroLeeBoolean = resultado
    End Function
    Public Sub Parameter_ListWrite(ByVal queDoc As Inventor.Document,
                                 ByVal queFi As String,
                                 ByVal lPar As List(Of UtilesAlberto.Parameter),
                                 Optional ByVal cerrar As Boolean = False)
        If queDoc Is Nothing And (queFi = "" OrElse IO.File.Exists(queFi) = False) Then
            Debug.Print("Object queDoc Nothing. And queFi Not exist")
            Exit Sub
        End If
        '
        If queFi <> "" AndAlso IO.File.Exists(queFi) Then
            oAppI.SilentOperation = True
            queDoc = oAppI.Documents.Open(queFi, False)
            oAppI.SilentOperation = False
        End If
        oAppI.ScreenUpdating = False
        '
        ' List of UtilesAlberto.Parameter
        For Each oP As UtilesAlberto.Parameter In lPar
            Dim quePar As String = oP._Parameter
            Dim queVal As String = oP._Value.ToString.Trim
            Dim queValDbl As String = IIf(oP._Value.ToString = "", "", oP._Value.ToString.Split(" "c)(0).Trim)
            '
            Dim oPar As Inventor.Parameter = Nothing
            Try
                If queDoc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                    oPar = CType(queDoc, AssemblyDocument).ComponentDefinition.Parameters.Item(quePar)
                ElseIf queDoc.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
                    oPar = CType(queDoc, PartDocument).ComponentDefinition.Parameters.Item(quePar)
                End If
            Catch ex As Exception
                Continue For
            End Try
            ' Si no está en uso, continuar.
            If oPar.InUse = False Then
                'Continue For
            End If
            '' Solo lo actualizaremos si es un parámetro modificable.
            If oPar.ParameterType <> ParameterTypeEnum.kDerivedParameter And
        oPar.ParameterType <> ParameterTypeEnum.kReferenceParameter And
        oPar.ParameterType <> ParameterTypeEnum.kTableParameter Then
                Dim queUnits As String = oPar.Units
                Dim queValUnits As String = ""
                If queUnits = "Text" Or queUnits = "Boolean" Then
                    If oPar.Value <> queVal Then
                        Try
                            oPar.Value = queVal
                        Catch ex As Exception
                            Debug.Print("Value : " & ex.ToString)
                        End Try
                    End If
                Else
                    queValUnits = queValDbl.Trim & " " & queUnits.Trim
                    If oPar.Expression <> queValUnits Then
                        Try
                            oPar.Expression = queValUnits
                        Catch ex As Exception
                            Debug.Print("Expression : " & ex.ToString)
                            If IsNumeric(queValDbl) Then
                                Try
                                    oPar.Value = CDbl(queValDbl)
                                Catch ex1 As Exception
                                    Debug.Print("Value : " & ex1.ToString)
                                End Try
                            End If
                        End Try
                    End If
                End If
            End If
            oPar = Nothing
        Next
        '
        oAppI.ScreenUpdating = True
        If queDoc.RequiresUpdate Then queDoc.Update2()
        If cerrar = True Then
            queDoc.Save2()
            queDoc.Close(True)
        End If
    End Sub

    'Public Sub Parameter_ListWriteMK(ByVal queDoc As Inventor.Document,
    '                             ByVal lPar As List(Of UtilesAlberto.Parameter),
    '                             Optional crear As Boolean = True)
    '    If queDoc Is Nothing Then
    '        Debug.Print("Object queDoc Nothing. And queFi Not exist")
    '        Exit Sub
    '    End If
    '    '
    '    oAppI.SilentOperation = True
    '    oAppI.ScreenUpdating = False
    '    '
    '    ' List of UtilesAlberto.Parameter
    '    For Each oP As UtilesAlberto.Parameter In lPar
    '        Dim quePar As String = oP._Parameter
    '        Dim queVal As String = oP._Value.ToString.Trim
    '        Dim queValDbl As String = IIf(oP._Value.ToString = "", "", oP._Value.ToString.Split(" "c)(0).Trim)
    '        '
    '        Dim oPar As Inventor.Parameter = Nothing
    '        Dim oCdIAM As AssemblyComponentDefinition = Nothing
    '        Dim oCdIPT As PartComponentDefinition = Nothing
    '        Try
    '            If crear = True AndAlso quePar.StartsWith("sv_") Then
    '                Dim uPar As UserParameter = Nothing
    '                Dim unitSpec As Inventor.UnitsTypeEnum
    '                If IsNumeric(queValDbl) = False Then
    '                    ' Not numeric = Text or Boolen
    '                    If queValDbl.ToUpper = "TRUE" OrElse queValDbl.ToUpper = "FALSE" Then
    '                        unitSpec = UnitsTypeEnum.kBooleanUnits
    '                    Else
    '                        unitSpec = UnitsTypeEnum.kTextUnits
    '                    End If
    '                Else
    '                    ' Is numeric (mm, su, gr)
    '                    If oP._Units.ToLower = "mm" Then
    '                        unitSpec = UnitsTypeEnum.kMillimeterLengthUnits
    '                    ElseIf oP._Units.ToLower = "su" Then
    '                        unitSpec = UnitsTypeEnum.kUnitlessUnits
    '                    ElseIf oP._Units.ToLower = "gr" Then
    '                        unitSpec = UnitsTypeEnum.kGradAngleUnits
    '                    Else
    '                        unitSpec = UnitsTypeEnum.kTextUnits
    '                    End If
    '                End If

    '                If oCdIAM IsNot Nothing Then
    '                    uPar = oCdIAM.Parameters.UserParameters.AddByExpression(quePar, queVal, unitSpec)
    '                ElseIf oCdIPT IsNot Nothing Then
    '                    uPar = oCdIPT.Parameters.UserParameters.AddByExpression(quePar, queVal, unitSpec)
    '                End If
    '            End If
    '            '
    '            If queDoc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
    '                oCdIAM = CType(queDoc, AssemblyDocument).ComponentDefinition
    '                oPar = oCdIAM.Parameters.Item(quePar)
    '            ElseIf queDoc.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
    '                oCdIPT = CType(queDoc, PartDocument).ComponentDefinition
    '                oPar = oCdIPT.Parameters.Item(quePar)
    '            End If
    '        Catch ex As Exception
    '            ' No existe el parámetro
    '            Continue For
    '        End Try
    '        ' Si no está en uso, continuar.
    '        'If oPar.InUse = False Then
    '        'Continue For
    '        'End If
    '        '' Solo lo actualizaremos si es un parámetro modificable.
    '        If oPar.ParameterType <> ParameterTypeEnum.kDerivedParameter And
    '    oPar.ParameterType <> ParameterTypeEnum.kReferenceParameter And
    '    oPar.ParameterType <> ParameterTypeEnum.kTableParameter Then
    '            Dim queUnits As String = oPar.Units
    '            Dim queValUnits As String = ""
    '            If queUnits = "txt" OrElse queUnits = "Text" OrElse queUnits = "Boolean" OrElse queUnits = "" Then
    '                If oPar.Value <> queVal Then
    '                    Try
    '                        oPar.Value = queVal
    '                    Catch ex As Exception
    '                        Debug.Print("Value : " & ex.ToString)
    '                    End Try
    '                End If
    '            Else
    '                queValUnits = queValDbl.Trim & " " & queUnits.Trim
    '                If oPar.Expression <> queValUnits Then
    '                    Try
    '                        oPar.Expression = queValUnits
    '                    Catch ex As Exception
    '                        Debug.Print("Expression : " & ex.ToString)
    '                        If IsNumeric(queValDbl) Then
    '                            Try
    '                                oPar.Value = CDbl(queValDbl)
    '                            Catch ex1 As Exception
    '                                Debug.Print("Value : " & ex1.ToString)
    '                            End Try
    '                        End If
    '                    End Try
    '                End If
    '            End If
    '        End If
    '        oPar = Nothing
    '    Next
    '    '
    '    If queDoc.RequiresUpdate Then queDoc.Update2()
    '    oAppI.ScreenUpdating = True
    '    oAppI.SilentOperation = False
    'End Sub
    Public Sub ParametroEscribeDouble(ByVal queDoc As Inventor.Document, ByVal queFi As String, ByVal quePar As String, ByVal queVal As Object, Optional ByVal queOperacion As IEnum.OperacionValor = IEnum.OperacionValor.cambiar, Optional ByVal cerrar As Boolean = False)
        If queFi <> "" AndAlso IO.File.Exists(queFi) Then
            oAppI.SilentOperation = True
            queDoc = oAppI.Documents.Open(queFi, False)
            oAppI.SilentOperation = False
        End If
        oAppI.ScreenUpdating = False
        ' queVal vendrá siempre en cm. Ya cambiamos a mm si procede.
        Dim oPar As Inventor.Parameter = Nothing
        Try
            If queDoc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                oPar = CType(queDoc, AssemblyDocument).ComponentDefinition.Parameters.Item(quePar)
            ElseIf queDoc.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
                oPar = CType(queDoc, PartDocument).ComponentDefinition.Parameters.Item(quePar)
            End If
        Catch ex As Exception
            'MsgBox("Error ParametroASMEscribe. El parametro (" & quePar & ") no existe. O valor (" & queVal.ToString & ") incorrecto.")
            'Debug.Print("Error ParametroASMEscribe. El parametro (" & quePar & ") no existe. O valor (" & queVal.ToString & ") incorrecto.")
            Exit Sub
        End Try
        '' Solo lo actualizaremos si es un parámetro modificable.
        If oPar.ParameterType <> ParameterTypeEnum.kDerivedParameter And
    oPar.ParameterType <> ParameterTypeEnum.kReferenceParameter And
    oPar.ParameterType <> ParameterTypeEnum.kTableParameter And
    IsNumeric(Left(oPar.Expression, 1)) = True Then
            ' Dim valor As Object = queVal
            Select Case queOperacion
                Case IEnum.OperacionValor.cambiar
                    If oPar.Value.ToString <> queVal.ToString Then
                        If IsNumeric(queVal) Then
                            If oPar.Value <> CDbl(queVal) Then oPar.Value = CDbl(queVal)
                        Else
                            Try
                                If oPar.Expression <> queVal.ToString Then oPar.Expression = queVal.ToString
                            Catch ex As Exception
                                If Log Then PonLog("Error en ParametroEscribe con parametro " & oPar.Name)
                            End Try
                        End If
                    End If
                Case IEnum.OperacionValor.sumar
                    If IsNumeric(queVal) Then
                        If oPar.Value <> oPar.Value + CDbl(queVal) Then _
                    oPar.Value = oPar.Value + CDbl(queVal)
                    End If
                Case IEnum.OperacionValor.restar
                    If IsNumeric(queVal) Then
                        If oPar.Value <> oPar.Value - CDbl(queVal) Then _
                    oPar.Value = oPar.Value - CDbl(queVal)
                    End If
            End Select
        End If
        oPar = Nothing
        If queDoc.RequiresUpdate Then queDoc.Update2()
        If cerrar = True Then
            queDoc.Save2()
            queDoc.Close(True)
        End If
        oAppI.ScreenUpdating = True
    End Sub
    '
    Public Sub ParametroEscribeString(ByVal queDoc As Inventor.Document, ByVal queFi As String, ByVal quePar As String, ByVal queVal As String, Optional ByVal cerrar As Boolean = False)
        If queFi <> "" AndAlso IO.File.Exists(queFi) Then
            oAppI.SilentOperation = True
            queDoc = oAppI.Documents.Open(queFi, False)
            oAppI.SilentOperation = False
        End If
        oAppI.ScreenUpdating = False
        ' queVal vendrá siempre en cm. Ya cambiamos a mm si procede.
        Dim oPar As Inventor.Parameter = Nothing
        Try
            If queDoc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                oPar = CType(queDoc, AssemblyDocument).ComponentDefinition.Parameters.Item(quePar)
            ElseIf queDoc.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
                oPar = CType(queDoc, PartDocument).ComponentDefinition.Parameters.Item(quePar)
            End If
        Catch ex As Exception
            'MsgBox("Error ParametroASMEscribe. El parametro (" & quePar & ") no existe. O valor (" & queVal.ToString & ") incorrecto.")
            'Debug.Print("Error ParametroASMEscribe. El parametro (" & quePar & ") no existe. O valor (" & queVal.ToString & ") incorrecto.")
            Exit Sub
        End Try
        '' Solo lo actualizaremos si es un parámetro modificable.
        If oPar.ParameterType <> ParameterTypeEnum.kDerivedParameter And
    oPar.ParameterType <> ParameterTypeEnum.kReferenceParameter And
    oPar.ParameterType <> ParameterTypeEnum.kTableParameter And
    oPar.ModelValueType = ModelValueTypeEnum.kNominalValue Then
            ' Dim valor As Object = queVal
            If oPar.Value.ToString <> queVal.ToString Then
                oPar.Value = queVal
            End If
        End If
        oPar = Nothing
        If queDoc.RequiresUpdate Then queDoc.Update2()
        If cerrar = True Then
            queDoc.Save2()
            queDoc.Close(True)
        End If
        oAppI.ScreenUpdating = True
    End Sub
    '
    '' Le indicamos nombre completo del fichero Inventor y colDatos de la clase clsDatosFila (arrFilas(COM_CLAVE))
    Public Sub ParametrosEscribeTODOSCaminoHash(ByVal queFichero As String, ByVal colDatosCls As Hashtable)
        '' Hemos configurado las propiedades de algunas operaciones para que se activen
        '' si el parámetro que las controla es superior a 0,1 (los valores 0 los convertimos a 0,1)
        '' En este arraylist tendremos los nombres de los Parametros que se externalizan como Propiedades
        '' para no volver a escribirlos cuando pongamos las propiedades.
        Dim oDoc As Inventor.Document = Nothing
        If Dir(queFichero) = "" Then
            MsgBox("El fichero " & queFichero & vbCrLf & "NO EXISTE...")
            Exit Sub
        End If
        Dim EstabaAbierto As Boolean = Me.FicheroAbierto(queFichero)
        Me.oAppI.SilentOperation = True
        If EstabaAbierto = True Then
            oDoc = Me.oAppI.Documents.ItemByName(queFichero)
        Else
            oDoc = Me.oAppI.Documents.Open(queFichero, False)
        End If
        ParametrosEscribeTODOSCaminoHashDoc(oDoc, colDatosCls)
        If oDoc IsNot Nothing AndAlso EstabaAbierto = False Then oDoc.Close(True)
        If Not (oDoc Is Nothing) Then Marshal.ReleaseComObject(oDoc)
        oDoc = Nothing
        '
        System.GC.Collect()
        System.GC.WaitForPendingFinalizers()
        System.GC.Collect()
    End Sub
    ' Le indicamos nombre completo del fichero Inventor y colDatos de la clase clsDatosFila (arrFilas(COM_CLAVE))
    Public Sub ParametrosEscribeTODOSCaminoHashDoc(ByRef queDoc As Inventor.Document, ByVal colDatosCls As Hashtable)
        '' Hemos configurado las propiedades de algunas operaciones para que se activen
        '' si el parámetro que las controla es superior a 0,1 (los valores 0 los convertimos a 0,1)
        '' En este arraylist tendremos los nombres de los Parametros que se externalizan como Propiedades
        '' para no volver a escribirlos cuando pongamos las propiedades.
        Dim arrExternas As New ArrayList
        Dim oUps As Inventor.UserParameters = Nothing
        Dim nP As String = ""
        Me.oAppI.SilentOperation = True
        oAppI.ScreenUpdating = False
        '' ***** RECORREMOS TODOS LOS PARÁMETROS
        If queDoc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
            'oDoc = CType(oDoc, Inventor.AssemblyDocument)
            oUps = CType(queDoc, Inventor.AssemblyDocument).ComponentDefinition.Parameters.UserParameters
        ElseIf queDoc.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
            'oDoc = CType(oDoc, Inventor.PartDocument)
            oUps = CType(queDoc, Inventor.PartDocument).ComponentDefinition.Parameters.UserParameters
        End If
        For Each quePar As UserParameter In oUps
            nP = quePar.Name
            '' ***** Si no existe el nombre en colDatosCls, pasamos al siguiente parámetro.
            If Not colDatosCls.ContainsKey(quePar.Name) Then
                'If quePar.Name <> "COM_DIMX" And quePar.Name <> "COM_DIMY" And quePar.Name <> "COM_DIMZ" Then
                Continue For
                'End If
            End If
            Dim valor As Object
            Try
                valor = colDatosCls(quePar.Name)
            Catch ex As Exception
                Continue For
            End Try

            Dim parEx As Object = quePar.Expression
            Dim parExSin As String = Trim(quePar.Expression.Replace(quePar.Units, ""))

            '' ***** TODO ESTO PARA SABER SI LO RELLENAMOS O PASAMOS AL SIGUIENTE *****

            '' Si tiene una expresion, pasamos al siguiente.
            'If parEx IsNot Nothing Then Continue For
            Try
                If IsNumeric(parExSin) = False Then Continue For
                If quePar.Expression.Contains("_") Then Continue For
                If IsNumeric(quePar.Expression.Substring(0, 1)) = False Then Continue For

                If IsDBNull(valor) AndAlso quePar.Value.GetTypeCode = TypeCode.String Then valor = ""
                If IsDBNull(valor) AndAlso
        (quePar.Value.GetTypeCode = TypeCode.Decimal Or quePar.Value.GetTypeCode = TypeCode.Double) _
        Then valor = 0

                If valor.ToString = "" Then Continue For
                If valor = quePar.Value Then Continue For
                '' No cambiamos los diámetros si vienen con valor 0 o ""
                If nP.StartsWith("di_") And IsNumeric(valor) AndAlso valor = 0 Then Continue For
                If nP.StartsWith("di_") And IsNumeric(valor) = False AndAlso valor = "" Then Continue For
                '' ***************************************************************************************
                '' Si el parámetro se ha externalizado lo guardamos en la colección para
                '' no sobrescribir su valor cuando escribamos las iProperties (bucle siguiente)
                If quePar.ExposedAsProperty = False Then quePar.ExposedAsProperty = True
                If arrExternas.Contains(quePar.Name) = False Then arrExternas.Add(quePar.Name)

                If quePar.Units <> "su" Then
                    If quePar.Value.GetTypeCode = TypeCode.String AndAlso quePar.Value = "0" Then _
                If quePar.Value <> "0,01" Then _
                    quePar.Value = "0,01"
                    If quePar.Value.GetTypeCode = TypeCode.Decimal AndAlso quePar.Value = 0 Then _
                If quePar.Value <> 0.01 Then _
                    quePar.Value = 0.01
                    If quePar.Value.GetTypeCode = TypeCode.Double AndAlso quePar.Value = 0 Then _
                If quePar.Value <> 0.01 Then _
                    quePar.Value = 0.01
                    '' Valor que viene de colDatosCls, para evaluarlo
                End If
            Catch ex As Exception
                ''***** LOG PARA CONTROL DE ERRORES *****
                If Log Then PonLog(vbCrLf & "Error en ParametrosEscribeTODOSCaminoHashDoc con " & nP & " y valor " & valor.ToString & vbCrLf)
                ''*****************************************
                MsgBox("Error en ParametrosEscribeTODOSCaminoHashDoc. Al evaluar valores")
                Continue For
            End Try

            Try
                '' Parametro de la BD a parametro Inventor.
                'Dim nombreReal As String = quePar.Name
                'If quePar.Name = "ds_alt" Then
                'If quePar.Value <> CDbl(colDatosCls("ds_alt")) Then quePar.Value = CDbl(colDatosCls("AltTot")) ' AltTol ya tiene el valor calculado previamente. CDbl(valor) + CDbl(colDatosCls("lo_emp"))
                'Continue For
                'End If

                '' Si los valores son iguales. No hacemos nada con el valor.
                If valor = quePar.Value Then Continue For
                If valor.ToString = quePar.Expression Then Continue For
                If quePar.Units = "gr" Then Continue For

                If valor = 0 And quePar.Units = "su" Then
                    If quePar.Expression <> ("0 " & quePar.Units) Then _
                quePar.Expression = ("0 " & quePar.Units)
                    'quePar.Value = 1
                ElseIf valor = 0 And quePar.Units = "cm" Then
                    'If quePar.Expression <> ("0,01 " & quePar.Units) Then quePar.Expression = ("0,01 " & quePar.Units)
                    If quePar.Value <> 0.01 Then _
                quePar.Value = 0.01
                ElseIf valor = 0 And quePar.Units = "mm" Then
                    'If quePar.Expression <> ("0,01 " & quePar.Units) Then quePar.Expression = ("0,01 " & quePar.Units)
                    If quePar.Value <> 0.1 Then _
                quePar.Value = 0.1
                    'ElseIf quePar.Name = "ds_lar" Then
                    'If quePar.Value <> CDbl(colDatosCls("ds_lar")) Then quePar.Value = CDbl(colDatosCls("ds_lar"))
                Else
                    'If quePar.Expression <> (FormatNumber(valor, 2, , , Microsoft.VisualBasic.TriState.False) & " " & quePar.Units) Then _
                    'quePar.Expression = FormatNumber(valor, 2, , , Microsoft.VisualBasic.TriState.False) & " " & quePar.Units
                    If quePar.Expression <> valor & " " & quePar.Units Then _
                quePar.Expression = valor & " " & quePar.Units
                    'quePar.Expression = Format(valor, "f") & " " & quePar.Units
                    'quePar.Value = FormatNumber(valor, 2)
                End If
                'oAp.UserInterfaceManager.DoEvents()
            Catch ex As Exception
                ''***** LOG PARA CONTROL DE ERRORES *****
                If Log Then PonLog(vbCrLf & "Error en ParametrosEscribeTODOSCaminoHashDoc con parametro (" & nP & ") y valor " & valor.ToString & vbCrLf)
                ''*****************************************
                'MsgBox("Error en ParametrosEscribeTODOSCaminoHash. Con parametro (" & nP & ") con el valor --> " & valor.ToString)
                Continue For
            End Try
            Me.DoEventsInventor(True)
            'oAppCls.UserInterfaceManager.DoEvents()
        Next
        'oAp.UserInterfaceManager.DoEvents()
        '' RECORREMOS TODAS LAS PROPIEDADES de Usuario
        Dim oPS As PropertySet = queDoc.PropertySets.Item("{D5CDD505-2E9C-101B-9397-08002B2CF9AE}")
        For Each oPro As Inventor.Property In oPS
            nP = oPro.Name
            'Dim nP As String = oPro.Name
            Dim valor As Object = Nothing
            Try
                'If oPro.Expression IsNot Nothing Then Continue For
                If arrExternas.Contains(oPro.Name) = True Then Continue For
                If colDatosCls.ContainsKey(oPro.Name) = False Then Continue For
                If oPro.Expression <> oPro.Value Then Continue For

                valor = colDatosCls(oPro.Name)
                If IsDBNull(valor) Then valor = ""

                If valor.ToString = "" Then Continue For
                If oPro.Value.ToString = valor.ToString Then Continue For

                oPro.Value = Trim(valor.ToString)
            Catch ex As Exception
                ''***** LOG PARA CONTROL DE ERRORES *****
                If Log Then PonLog(vbCrLf & "Error en ParametrosEscribeTODOSCaminoHashDoc con propiedad (" & nP & ") y valor " & valor.ToString & vbCrLf)
                ''*****************************************
                'MsgBox("Error en ParametrosEscribeTODOSCaminoHashDoc. Con propiedad (" & nP & ") con el valor --> " & valor.ToString)
                Continue For
            End Try
            Me.DoEventsInventor(True)
        Next
        queDoc.Update2()
        queDoc.Save2(False)

        If Not (oPS Is Nothing) Then Marshal.ReleaseComObject(oPS)
        oPS = Nothing
        If Not (oUps Is Nothing) Then Marshal.ReleaseComObject(oUps)
        oUps = Nothing

        System.GC.Collect()
        System.GC.WaitForPendingFinalizers()
        System.GC.Collect()

        oAppI.ScreenUpdating = True
        Me.oAppI.SilentOperation = False
    End Sub
    Public Sub ParametrosEscribeHijos(ByVal docPadreCamino As String, ByVal docHijo As PartDocument)
        '' "docPadre" es el ensamblaje PADRE origen (el que tiene en item(1) la pieza con todos los parámetros)
        '' Buscaremos en él todos los documentos referenciados (no incluir item(1))
        '' "docHijo" contiene todos los Userparameters a crear/cambiar en el resto.
        '' "pOrigen" el la colección de parametros originales a crear/cambiar referencias de "docPadre"
        Dim pOrigen As UserParameters = docHijo.ComponentDefinition.Parameters.UserParameters
        Dim dirModificar As String = DameParteCamino(docPadreCamino, IEnum.ParteCamino.CaminoConFicheroSinExtensionBarra) ' docPadre.FullFileName.Replace(".iam", "\")
        If IO.Directory.Exists(dirModificar) = False Then Exit Sub

        For Each fichero As String In IO.Directory.GetFiles(dirModificar, "*.i*", IO.SearchOption.TopDirectoryOnly)
            If fichero.ToLower.EndsWith(".iam") Or fichero.ToLower.EndsWith(".ipt") Then
                Me.ParametrosEscribeHijosUserParameters(fichero, pOrigen)
            End If
            oAppI.UserInterfaceManager.DoEvents()
        Next
    End Sub

    Public Sub ParametrosEscribeHijosUserParameters(ByVal queF As String, ByVal pOrigen As UserParameters)
        oAppI.ScreenUpdating = False
        Dim oD As Inventor.Document = Nothing
        Try
            oD = oAppI.Documents.Open(queF, False)
        Catch ex As Exception
            '' Si da error al abrir salimos fuera, porque no será de inventor. O está bloqueado.
            Exit Sub
        End Try
        Me.PropiedadEscribe(oD, "Nº de pieza", DameParteCamino(oD.FullFileName, IEnum.ParteCamino.SoloFicheroSinExtension))
        Dim oPs As UserParameters = Nothing
        If oD.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
            oPs = CType(oD, AssemblyDocument).ComponentDefinition.Parameters.UserParameters
        ElseIf oD.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
            oPs = CType(oD, PartDocument).ComponentDefinition.Parameters.UserParameters
        End If

        For Each oP As UserParameter In oPs
            'oAp.UserInterfaceManager.DoEvents()
            Try
                If oP.Units <> pOrigen.Item(oP.Name).Units Then _
            oP.Units = pOrigen.Item(oP.Name).Units
                If oP.Expression <> pOrigen.Item(oP.Name).Expression Then _
            oP.Expression = pOrigen.Item(oP.Name).Expression
            Catch ex As Exception
                Continue For
                '' No hacemos nada. El parametro no existe
            End Try
        Next
        If oD.RequiresUpdate Then oD.Update2()

        oAppI.ScreenUpdating = True
        'oD.Close()  ' Cerramos guardando todo
    End Sub
    '' Para poner la expresion de un parámetro dentro de una fórmula.
    Public Sub ParametroPonFormula(ByRef quePie As Inventor.PartDocument, ByVal nombreP As String, ByVal formula As String)
        If quePie Is Nothing Then Exit Sub
        If quePie.ComponentDefinition.Parameters.UserParameters Is Nothing Then Exit Sub
        If quePie.ComponentDefinition.Parameters.UserParameters.Count = 0 Then Exit Sub
        oAppI.ScreenUpdating = False
        Try
            Dim queP As Inventor.UserParameter = quePie.ComponentDefinition.Parameters.UserParameters.Item(nombreP)
            ''***** Comprobamos si existe el parámetro nombreP Ejemplo: "ds_lar_tot"
            '' y si tiene la fórmula correcta formula(valores) Ejemplo: "floor(valores sumanos)"
            If Not queP.Expression.Contains(formula) Then
                'isolate(floor(largo + alto + ancho);su;mm)
                Dim expresion As String = "isolate(" & formula & "(" & queP.Expression & ");su;" & queP.Units & ")"
                queP.Expression = expresion
            End If
            quePie.Save2(False)
            ''********************************************************
        Catch ex As Exception
            '' No existe el parámetro en UserParameters.
            Exit Sub
        Finally
            oAppI.ScreenUpdating = True
        End Try
    End Sub
End Class