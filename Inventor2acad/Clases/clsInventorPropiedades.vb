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
    Public Sub PropiedadesCopiadasModelo_Actualiza()
        '' UpdateCopiedModeliPropertiesCmd / Act&ualizar iProperties de modelo copiadas / Vuelve a copiar las iProperties elegidas del modelo de origen en el dibujo
        If oAppI.ActiveDocumentType = DocumentTypeEnum.kDrawingDocumentObject Then
            Try
                oAppI.CommandManager.ControlDefinitions.Item("UpdateCopiedModeliPropertiesCmd").Execute2(True)
            Catch ex As Exception
                '' Si da error sería que no tiene propiedades copiadas.
                '' Si da error sería que no tiene propiedades copiadas.
            End Try
        End If
    End Sub
    Public Function PropiedadLeeTodasInventorArray(ByVal queDoc As String) As Object()
        Dim resultado(2) As Object
        Dim colEn As New Hashtable
        Dim colEs As New Hashtable
        Dim imagen As System.Drawing.Image = Nothing
        Dim estababierto As Boolean = True

        oAppI.SilentOperation = True

        estababierto = FicheroAbierto(queDoc)

        Dim queDocSin As String = ""
        If queDoc.ToUpper.EndsWith(".iam") = True Then
            queDocSin = Me.DameCaminoSinComponentesInventor(queDoc)
        Else    'If queDoc.EndsWith(".ipt") Then
            queDocSin = queDoc
        End If

        ' Abrir un documento.
        Dim oDoc As Inventor.Document = Nothing
        Dim oProSs As PropertySets = Nothing

        Try
            If estababierto = True Then
                oDoc = oAppI.Documents.ItemByName(queDoc)
            Else
                oDoc = oAppI.Documents.Open(queDocSin, False)
            End If
            oDoc = oAppI.Documents.Open(queDocSin, False)
            oProSs = oDoc.PropertySets
        Catch ex As Exception
            MsgBox("Error : " & vbCrLf & vbCrLf & ex.Message)
        End Try

        Try
            For Each oProS As PropertySet In oProSs
                For Each oPro As Inventor.Property In oProS
                    If oPro.Name = "Thumbnail" Or oPro.DisplayName = "Miniatura" Then
                        imagen = clsIp.GetImageFromIPictureDisp(oPro.Value)
                    Else
                        If colEs.ContainsKey(oPro.DisplayName) Then _
                    colEs.Add(oPro.DisplayName, oPro.Value.ToString)
                        If colEn.ContainsKey(oPro.Name) Then _
                    colEn.Add(oPro.Name, oPro.Value.ToString)
                    End If
                Next
            Next
        Catch ex As Exception
            MsgBox("Error PropiedadLeeTodasApprenticeArray...")
        End Try

        If estababierto = False Then oDoc.Close(True)

        ''Liberar Objetos.
        If Not (oDoc Is Nothing) Then Marshal.ReleaseComObject(oDoc)
        oDoc = Nothing
        If Not (oProSs Is Nothing) Then Marshal.ReleaseComObject(oProSs)
        oProSs = Nothing

        System.GC.WaitForPendingFinalizers()
        System.GC.Collect()
        oAppI.SilentOperation = False
        '' Guardamos en el Array los 3 valores Es, En, imagen
        resultado(0) = colEs
        resultado(1) = colEn
        resultado(2) = imagen

        Return resultado
        Exit Function
    End Function
    Public Function PropiedadLeeCategoria(ByVal queDoc As Inventor.Document, Optional queF As String = "", Optional ByRef quePss As PropertySets = Nothing) As String
        '' Por si le damos fullFilename (queF) en vez de Document
        Dim estabaabierto As Boolean = True
        If queDoc Is Nothing AndAlso queF <> "" AndAlso IO.File.Exists(queF) = True Then
            estabaabierto = FicheroAbierto(queF)
            oAppI.SilentOperation = True
            If estabaabierto = True Then
                queDoc = oAppI.Documents.ItemByName(queF)
            Else
                queDoc = oAppI.Documents.Open(queF, False)
            End If
            oAppI.SilentOperation = False
        End If
        '' ***********************************************
        '' Lee un valor de texto en una iProperty de usuario. Si no existe la crea con valor "".
        Dim resultado As String = ""
        '' Información resumen documento Inventor / Información resumen documento Inventor
        '' Internal name: {D5CDD502-2E9C-101B-9397-08002B2CF9AE}
        '' Nombre: Category (Categoría) / Valor:  / Id: 2

        Dim oProS As PropertySet = queDoc.PropertySets.Item("{D5CDD502-2E9C-101B-9397-08002B2CF9AE}")
        Dim oPro As Inventor.Property = oProS.ItemByPropId(2)
        resultado = oPro.Value.ToString
        If estabaabierto = False Then queDoc.Close(True)
        PropiedadLeeCategoria = resultado
        Exit Function
    End Function
    Public Sub PropiedadEscribe(ByRef queCom As Inventor.Document, ByVal quePro As String, ByVal queVal As Object, Optional queF As String = "")
        '' Por si le damos fullFilename (queF) en vez de Document
        Dim estabaabierto As Boolean = True
        If queCom Is Nothing AndAlso queF <> "" AndAlso IO.File.Exists(queF) = True Then
            estabaabierto = FicheroAbierto(queF)
            oAppI.SilentOperation = True
            If estabaabierto = True Then
                queCom = oAppI.Documents.ItemByName(queF)
            Else
                queCom = oAppI.Documents.Open(queF, False)
            End If
            oAppI.SilentOperation = False
        End If
        '' ***********************************************
        Dim oProSs As PropertySets = queCom.PropertySets

        Try
            For Each oProS As PropertySet In oProSs
                For Each oPro As Inventor.Property In oProS
                    If oPro.Name = quePro Or oPro.DisplayName = quePro Or oPro.PropId.ToString = quePro Then
                        '' Usaremos esto como expresión. Por si es Nothing oPro.Expression.
                        Dim oProExp As String = IIf(oPro.Expression Is Nothing, "", oPro.Expression)

                        '' Solo escribiremos el valor si: la propiedad no es una expresion o le mandamos una nueva expresion y si
                        '' no es una expresión pero es diferente.
                        If queVal = oPro.Value Or queVal = oProExp Then
                            GoTo FINAL
                        ElseIf queVal.ToString.StartsWith("=") = True And queVal <> oProExp Then
                            oPro.Expression = queVal.ToString
                            GoTo FINAL
                        ElseIf queVal.ToString.StartsWith("=") = True And queVal = oProExp Then
                            GoTo FINAL
                        ElseIf oProExp.StartsWith("=") Then     '' Si hay una expresión ya. No hacemos nada.
                            GoTo FINAL
                        ElseIf queVal <> oPro.Value Then
                            oPro.Value = queVal
                            GoTo FINAL
                        End If
                    End If
                Next
            Next
        Catch ex As Exception
            ''***** LOG PARA CONTROL DE ERRORES *****
            If Log Then PonLog(vbCrLf & "Error PropiedadEscribe. El parametro (" & quePro & ") no existe. O valor (" & queVal.ToString & ") incorrecto." & vbCrLf)
            ''*****************************************
            'MsgBox("Error PropiedadEscribe. El parametro (" & quePro & ") no existe. O valor (" & queVal.ToString & ") incorrecto.")
        End Try
FINAL:
        Try
            If queCom.Dirty = True Then queCom.Save2(False)
            If estabaabierto = False Then queCom.Close(True)
        Catch ex As Exception
            ' No hacemos nada.
            Debug.Print(ex.Message)
        End Try
    End Sub
    Public Sub PropiedadEscribeUsuario(ByRef queDoc As Inventor.Document,
                               ByVal quePro As String,
                               ByVal queVal As Object,
                               Optional ByVal queFi As String = "",
                               Optional ByVal cerrar As Boolean = False,
                               Optional CREAR As Boolean = True,
                               Optional sobrescribir As Boolean = True)
        If queVal = "" Or queVal Is Nothing Then Exit Sub

        Dim estabaabierto As Boolean = False
        If queDoc Is Nothing Then
            If FicheroAbierto(queFi) = False Then
                queDoc = oAppI.Documents.Open(queFi, False)
                estabaabierto = False
            Else
                queDoc = oAppI.Documents.ItemByName(queFi)
                estabaabierto = True
            End If
        End If
        '' Escribe un valor de texto en una iProperty. Si no existe la crea con valor "".
        Dim oProS As PropertySet = queDoc.PropertySets.Item("{D5CDD505-2E9C-101B-9397-08002B2CF9AE}")
        Dim oPro As Inventor.Property = Nothing

        Try
            oPro = oProS.Item(quePro)
        Catch ex As Exception
            If CREAR = True Then
                oPro = oProS.Add(queVal.ToString, quePro)
                GoTo FINAL
            Else
                GoTo FINAL
            End If
        End Try
        '' Usaremos esto como expresión. Por si es Nothing oPro.Expression.
        Dim oProExp As String = IIf(oPro.Expression Is Nothing, "", oPro.Expression)
        '' Si sobrescribir=true. Siempre sobrescribimos el valor
        If sobrescribir Then
            Try
                oPro.Expression = queVal.ToString
                GoTo FINAL
            Catch ex As Exception
                If Log Then PonLog("Error en PropiedadEscribeUsuario con oPro.Expression")
                GoTo FINAL
            End Try
        End If
        '' Solo escribiremos el valor si: la propiedad no es una expresion o le mandamos una nueva expresion y si
        '' no es una expresión pero es diferente.
        If queVal.ToString.StartsWith("=") = True AndAlso queVal <> oProExp Then
            Try
                oPro.Expression = queVal.ToString
                GoTo FINAL
            Catch ex As Exception
                If Log Then PonLog("Error en PropiedadEscribeUsuario con oPro.Expression")
                GoTo FINAL
            End Try
        ElseIf queVal.ToString.StartsWith("=") = True AndAlso queVal = oProExp Then
            GoTo FINAL
        ElseIf oProExp.StartsWith("=") Then     '' Si hay una expresión ya. No hacemos nada.
            GoTo FINAL
        ElseIf queVal <> oPro.Value Then
            oPro.Value = queVal
            GoTo FINAL
        ElseIf queVal = oPro.Value Then
            GoTo FINAL
        End If

FINAL:

        If estabaabierto = False And cerrar = True Then
            '' Si queremos que se actualice y se guarde. Lo quitamos para ganar tiempo. Lo haremos al final.
            Try
                If queDoc.RequiresUpdate Then queDoc.Update2()
            Catch ex As Exception
                '' Continuamos.
            End Try
            Try

                If queDoc.Dirty = True Then queDoc.Save2(False)
            Catch ex As Exception
                '' No lo guardamos si da error.
            End Try
            Try
                queDoc.Close(True)
            Catch ex As Exception
                '' No hacemos nada. Lo dejamos abierto.
            End Try
        End If
        oPro = Nothing : oProS = Nothing
    End Sub
    Public Sub PropiedadEscribeDesignTracking(ByRef queDoc As Inventor.Document, ByVal quePro As String, ByVal queVal As Object, Optional ByVal queFi As String = "", Optional ByVal cerrar As Boolean = False)
        ''Propiedades de Design Tracking / Propiedades de Design Tracking
        'Internal name: {32853F0F-3444-11D1-9E93-0060B03C1CA6}
        If queVal = "" Or queVal Is Nothing Then Exit Sub

        Dim estabaabierto As Boolean = False
        If queDoc Is Nothing Then
            If FicheroAbierto(queFi) = False Then
                queDoc = oAppI.Documents.Open(queFi, False)
                estabaabierto = False
            Else
                queDoc = oAppI.Documents.ItemByName(queFi)
                estabaabierto = True
            End If
        End If
        '' Escribe un valor de texto en una iProperty. Si no existe la crea con valor "".
        Dim oProS As PropertySet = queDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}")
        Dim oPro As Inventor.Property = Nothing

        Try
            oPro = IIf(IsNumeric(quePro), oProS.ItemByPropId(CInt(quePro)), oProS.Item(quePro))
        Catch ex As Exception
            '' No hacemos nada. Aquí no se pueden crear propiedades nuevas.
        End Try
        '' Usaremos esto como expresión. Por si es Nothing oPro.Expression.
        Dim oProExp As String = IIf(oPro.Expression Is Nothing, "", oPro.Expression)

        '' Solo escribiremos el valor si: la propiedad no es una expresion o le mandamos una nueva expresion y si
        '' no es una expresión pero es diferente.
        If queVal.ToString.StartsWith("=") = True And queVal <> oProExp Then
            'oPro.Expression = queVal.ToString
            oPro.Expression = queVal.ToString
            GoTo FINAL
        ElseIf queVal.ToString.StartsWith("=") = True And queVal = oProExp Then
            GoTo FINAL
        ElseIf oProExp.StartsWith("=") Then     '' Si hay una expresión ya. No hacemos nada.
            GoTo FINAL
        ElseIf queVal <> oPro.Value Then
            oPro.Value = queVal
            GoTo FINAL
        ElseIf queVal = oPro.Value Then
            GoTo FINAL
        End If

FINAL:
        If estabaabierto = False And cerrar = True Then
            '' Si queremos que se actualice y se guarde. Lo quitamos para ganar tiempo. Lo haremos al final.
            Try
                If queDoc.RequiresUpdate Then queDoc.Update2()
            Catch ex As Exception
                '' Continuamos.
                Debug.Print("error")
            End Try
            Try

                If queDoc.Dirty = True Then queDoc.Save2(False)
            Catch ex As Exception
                '' No lo guardamos si da error.
                Debug.Print("error")
            End Try
            Try
                queDoc.Close(True)
            Catch ex As Exception
                '' No hacemos nada. Lo dejamos abierto.
                Debug.Print("error")
            End Try
        End If
        oPro = Nothing : oProS = Nothing
    End Sub
    Public Function PropiedadLeeUsuario(ByVal queDoc As Inventor.Document, ByVal quePro As String,
                                Optional queF As String = "",
                                Optional crear As Boolean = False,
                                Optional valor As String = "Centro",
                                Optional ByRef quePss As PropertySets = Nothing) As String
        '' Por si le damos fullFilename (queF) en vez de Document
        Dim estabaabierto As Boolean = True
        If queDoc Is Nothing AndAlso queF <> "" AndAlso IO.File.Exists(queF) = True Then
            estabaabierto = FicheroAbierto(queF)
            oAppI.SilentOperation = True
            If estabaabierto = True Then
                queDoc = oAppI.Documents.ItemByName(queF)
            Else
                queDoc = oAppI.Documents.Open(queF, False)
            End If
            oAppI.SilentOperation = False
        End If
        '' ***********************************************
        '' Lee un valor de texto en una iProperty de usuario. Si no existe la crea con valor "".
        Dim resultado As String = ""
        If queDoc Is Nothing Then
            Return resultado
            Exit Function
        End If
        Dim oProS As PropertySet = Nothing
        If quePss IsNot Nothing Then
            oProS = quePss.Item("{D5CDD505-2E9C-101B-9397-08002B2CF9AE}")
        Else
            oProS = queDoc.PropertySets.Item("{D5CDD505-2E9C-101B-9397-08002B2CF9AE}")
        End If
        Dim oPro As Inventor.Property = Nothing
        Try
            oPro = oProS.Item(quePro)
            resultado = oPro.Value.ToString
        Catch ex As Exception
            If crear = True Then
                'Me.PropiedadEscribeUsuario(queDoc, quePro, "Centro", , False)
                Me.PropiedadEscribeUsuario(queDoc, quePro, valor, , False)
                oPro = oProS.Item(quePro)
                resultado = oPro.Value.ToString
            Else
                resultado = ""
            End If
        End Try
        'resultado = oPro.Value.ToString
        oProS = Nothing
        oPro = Nothing
        If estabaabierto = False Then queDoc.Close(True)
        Return resultado
    End Function
    Public Function PropiedadLeeResumenInventor(ByRef queDoc As Inventor.Document, ByVal quePro As String, Optional ByRef quePss As PropertySets = Nothing) As String
        'Información de resumen de Inventor / Información de resumen de Inventor
        'Internal name: {F29F85E0-4FF9-1068-AB91-08002B27B3D9}
        '
        ' Nombre: Title (Título) / Valor:  / Id: 2
        ' Nombre: Subject (Asunto) / Valor:  / Id: 3
        ' Nombre: Author (Autor) / Valor: Raul / Id: 4
        ' Nombre: Keywords (Palabras clave) / Valor:  / Id: 5
        ' Nombre: Comments (Comentarios) / Valor:  / Id: 6
        ' Nombre: Last Saved By (Guardado por última vez por) / Valor:  / Id: 8
        ' Nombre: Revision Number (Nº de revisión) / Valor:  / Id: 9
        ' Nombre: Thumbnail (Miniatura) / Valor:  / Id: 17

        '' Lee un valor de texto en una iProperty de Resumen Inventor. Si no existe devuelve "".
        Dim resultado As String = ""
        Dim oProS As PropertySet = Nothing
        Dim oPro As Inventor.Property = Nothing
        Try
            If quePss IsNot Nothing Then
                oProS = quePss.Item("{D5CDD505-2E9C-101B-9397-08002B2CF9AE}")
            Else
                oProS = queDoc.PropertySets.Item("{F29F85E0-4FF9-1068-AB91-08002B27B3D9}")
            End If
            oPro = IIf(IsNumeric(quePro), oProS.ItemByPropId(CInt(quePro)), oProS.Item(quePro))
            resultado = oPro.Value.ToString
        Catch ex As Exception
            ' No existe la Propiedad indicada en quePro
        End Try

        oProS = Nothing
        oPro = Nothing
        Return resultado
    End Function
    Public Function PropiedadLeeResumenDocumento(ByRef queDoc As Inventor.Document, ByVal quePro As String, Optional ByRef quePss As PropertySets = Nothing) As String
        'Información resumen documento Inventor / Información resumen documento Inventor
        'Internal name: {D5CDD502-2E9C-101B-9397-08002B2CF9AE}

        'Nombre: Category (Categoría) / Valor:  / Id: 2
        'Nombre: Manager (Responsable) / Valor:  / Id: 14
        'Nombre: Company (Empresa) / Valor:  / Id: 15
        '' Lee un valor de texto en una iProperty de Resumen Documento. Si no existe devuelve "".
        Dim resultado As String = ""
        Dim oProS As PropertySet = Nothing
        Dim oPro As Inventor.Property = Nothing
        Try
            If quePss IsNot Nothing Then
                oProS = quePss.Item("{D5CDD505-2E9C-101B-9397-08002B2CF9AE}")
            Else
                oProS = queDoc.PropertySets.Item("{D5CDD502-2E9C-101B-9397-08002B2CF9AE}")
            End If
            oPro = IIf(IsNumeric(quePro), oProS.ItemByPropId(CInt(quePro)), oProS.Item(quePro))
            resultado = oPro.Value.ToString
        Catch ex As Exception
            ' No existe la Propiedad indicada en quePro
        End Try

        oProS = Nothing
        oPro = Nothing
        Return resultado
    End Function
    Public Function PropiedadLeeDesignTracking(queDoc As Inventor.Document, ByVal quePro As String, Optional ByRef quePss As PropertySets = Nothing) As String
        '' Lee un valor de texto en una iProperty de DesignTracking. Si no existe devuelve "".
        Dim resultado As String = ""
        Dim oProS As PropertySet = Nothing
        Dim oPro As Inventor.Property = Nothing
        Try
            If quePss IsNot Nothing Then
                oProS = quePss.Item("{D5CDD505-2E9C-101B-9397-08002B2CF9AE}")
            Else
                oProS = queDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}")
            End If
            oPro = IIf(IsNumeric(quePro), oProS.ItemByPropId(CInt(quePro)), oProS.Item(quePro))
            resultado = oPro.Value.ToString
        Catch ex As Exception
            ' No existe la Propiedad indicada en quePro
        End Try

        oProS = Nothing
        oPro = Nothing
        Return resultado
    End Function
    Public Function PropiedadLeeUsuarioDoc(ByRef queDoc As Inventor.Document, ByVal quePro As String, Optional ByRef quePss As PropertySets = Nothing) As String
        '' Propiedades de Inventor definidas por el usuario / Propiedades de Inventor definidas por el usuario
        '' Internal name: {D5CDD505-2E9C-101B-9397-08002B2CF9AE}
        '' Lee un valor de texto en una iProperty de DesignTracking. Si no existe devuelve "".
        Dim resultado As String = ""
        Dim oProS As PropertySet = Nothing
        Dim oPro As Inventor.Property = Nothing
        Try
            If quePss IsNot Nothing Then
                oProS = quePss.Item("{D5CDD505-2E9C-101B-9397-08002B2CF9AE}")
            Else
                oProS = queDoc.PropertySets.Item("{D5CDD505-2E9C-101B-9397-08002B2CF9AE}")
            End If
            oPro = IIf(IsNumeric(quePro), oProS.ItemByPropId(CInt(quePro)), oProS.Item(quePro))
            resultado = oPro.Value.ToString
        Catch ex As Exception
            ' No existe la Propiedad indicada en quePro
        End Try

        oProS = Nothing
        oPro = Nothing
        Return resultado
    End Function
    Public Function PropiedadLeeUsuarioHashtable(ByVal queDoc As Inventor.Document, ByVal queFichero As String, Optional ByRef quePss As PropertySets = Nothing) As Hashtable
        '' Lee un valor de texto en una iProperty de usuario. Si no existe la crea con valor "".
        Dim resultado As Hashtable = Nothing
        Dim oProS As PropertySet
        Dim estaabierto As Boolean = False


        If queDoc IsNot Nothing Then
            estaabierto = True
        ElseIf (queDoc Is Nothing) AndAlso queFichero <> "" AndAlso IO.File.Exists(queFichero) Then
            estaabierto = FicheroAbierto(queFichero)
            If estaabierto = True Then
                queDoc = oAppI.Documents.ItemByName(queFichero)
            Else
                '' El FullDocumentName sin componentes. Más rápido para abrir.
                If queFichero.ToLower.EndsWith(".iam") = True Then
                    queFichero = Me.DameCaminoSinComponentesInventor(queFichero)
                End If

                oAppI.SilentOperation = True
                queDoc = oAppI.Documents.Open(queFichero, False)
                oAppI.SilentOperation = False
            End If
        End If

        oProS = queDoc.PropertySets.Item("{D5CDD505-2E9C-101B-9397-08002B2CF9AE}")

        Try
            For Each oPro As Inventor.Property In oProS
                If resultado.ContainsKey(oPro.Name) = False Then resultado.Add(oPro.Name, oPro.Value)
            Next
        Catch ex As Exception
            '' Error leyendo propiedades
        End Try
        '' Si no estaba abierto antes. Lo cerramos.
        If estaabierto = False Then queDoc.Close(True)

        oProS = Nothing
        Return resultado
    End Function
End Class
#Region "PROPIEDADES"
'Public Sub Propiedades()
'    '***** Declare the Application object
'    Dim oApplication As Inventor.Application

'    ' Obtain the Inventor Application object.
'    ' This assumes Inventor is already running.
'    oApplication = GetObject(, "Inventor.Application")

'    ' Set a reference to the active document.
'    ' This assumes a document is open.
'    Dim oDoc As Document
'    oDoc = oApplication.ActiveDocument

'    ' Obtain the PropertySets collection object
'    Dim oPropsets As PropertySets
'    oPropsets = oDoc.PropertySets

'    '***** Iterate through all the PropertySets one by one using for loop
'    Dim oPropSet As PropertySet
'    For Each oPropSet In oPropsets
'        Dim Nombre As String
'        ' Obtain the DisplayName of the PropertySet
'        'Debug.Print "Display name: " & oPropSet.DisplayName
'        Nombre = oPropSet.DisplayName & " / "

'        ' Obtain the InternalName of the PropertySet
'        'Debug.Print "Internal name: " & oPropSet.InternalName
'        Nombre = Nombre & oPropSet.DisplayName '& vbCrLf

'        Debug.Print("" & Nombre & "")

'        ' Write a blank line to separate each pair.
'        Debug.Print()

'        '***** Todas las Propiedades
'        'Dim oPropSet As PropertySet
'        'For Each oPropSet In oPropsets
'        ' Iterate through all the Properties in the current set.
'        Dim oProp As Property
'        For Each oProp In oPropSet
'            ' Obtain the Name of the Property
'            Dim Name As String
'            Name = oProp.Name

'            ' Obtain the Value of the Property
'            Dim Value As Object
'            On Error Resume Next
'            Value = oProp.Value

'            ' Obtain the PropertyId of the Property
'            Dim PropertyId As Long
'            PropertyId = oProp.PropId
'            Debug.Print(vbTab & "Nombre: " & Name & " (" & oProp.DisplayName & ") / Valor: " & CStr(Value) & " / Id: " & CStr(PropertyId)) '& vbCrLf
'        Next
'        'Next
'        Nombre = "" : Name = "" : Value = Nohting : PropertyId = 0
'    Next
'    ' Write a blank line to separate each pair.
'    Debug.Print()
'End Sub


''***** RESULTADO DEL PROCEDIMIENTO QUE SE IMPRIME. Es una chapa *****

'Información de resumen de Inventor / Información de resumen de Inventor
'Internal name: {F29F85E0-4FF9-1068-AB91-08002B27B3D9}

'    Nombre: Title (Título) / Valor:  / Id: 2
'    Nombre: Subject (Asunto) / Valor:  / Id: 3
'    Nombre: Author (Autor) / Valor: Raul / Id: 4
'    Nombre: Keywords (Palabras clave) / Valor:  / Id: 5
'    Nombre: Comments (Comentarios) / Valor:  / Id: 6
'    Nombre: Last Saved By (Guardado por última vez por) / Valor:  / Id: 8
'    Nombre: Revision Number (Nº de revisión) / Valor:  / Id: 9
'    Nombre: Thumbnail (Miniatura) / Valor:  / Id: 17
'Información resumen documento Inventor / Información resumen documento Inventor
'Internal name: {D5CDD502-2E9C-101B-9397-08002B2CF9AE}

'    Nombre: Category (Categoría) / Valor:  / Id: 2
'    Nombre: Manager (Responsable) / Valor:  / Id: 14
'    Nombre: Company (Empresa) / Valor:  / Id: 15
'Propiedades de Design Tracking / Propiedades de Design Tracking
'Internal name: {32853F0F-3444-11D1-9E93-0060B03C1CA6}

'    Nombre: Creation Time (Fecha de creación) / Valor: 22/04/2008 8:05:14 / Id: 4
'    Nombre: Part Number (Nº de pieza) / Valor: FRONTAL_Grosor / Id: 5
'    Nombre: Project (Proyecto) / Valor:  / Id: 7
'    Nombre: Cost Center (Centro de costes) / Valor:  / Id: 9
'    Nombre: Checked By (Revisado por) / Valor:  / Id: 10
'    Nombre: Date Checked (Fecha de comprobación) / Valor: 01/01/1601 / Id: 11
'    Nombre: Engr Approved By (ING. aprobada por) / Valor:  / Id: 12
'    Nombre: Engr Date Approved (Fecha de aprobación de diseño ing.) / Valor: 01/01/1601 / Id: 13
'    Nombre: User Status (Estado del usuario) / Valor:  / Id: 17
'    Nombre: Material (Material) / Valor: Scotch / Id: 20
'    Nombre: Part Property Revision Id (Revisión de la pieza) / Valor: {827906D5-CB5E-4C98-B02F-7F109188604C} / Id: 21
'    Nombre: Catalog Web Link (Enlace Web de catálogo) / Valor:  / Id: 23
'    Nombre: Part Icon (Icono de la pieza) / Valor:  / Id: 28
'    Nombre: Description (Descripción) / Valor:  / Id: 29
'    Nombre: Vendor (Proveedor) / Valor:  / Id: 30
'    Nombre: Document SubType (Tipo de pieza) / Valor: {9C464203-9BAE-11D3-8BAD-0060B0CE6BB4} / Id: 31
'    Nombre: Document SubType Name (Nombre del tipo de pieza) / Valor: Chapa / Id: 32
'    Nombre: Proxy Refresh Date (Fecha de actualización de proxy) / Valor: 01/01/1601 / Id: 33
'    Nombre: Mfg Approved By (FAB. aprobada por) / Valor:  / Id: 34
'    Nombre: Mfg Date Approved (Fecha de aprobación de fabricación) / Valor: 01/01/1601 / Id: 35
'    Nombre: Cost (Coste) / Valor: 0 / Id: 36
'    Nombre: Standard (Norma) / Valor:  / Id: 37
'    Nombre: Design Status (Estado del diseño) / Valor: 1 / Id: 40
'    Nombre: Designer (Diseñador) / Valor: Raul / Id: 41
'    Nombre: Engineer (Ingeniero) / Valor:  / Id: 42
'    Nombre: Authority (Responsable) / Valor:  / Id: 43
'    Nombre: Parameterized Template (Plantilla parametrizada) / Valor: False / Id: 44
'    Nombre: Template Row (Fila de la plantilla) / Valor:  / Id: 45
'    Nombre: External Property Revision Id (Revisión externa de la pieza) / Valor: {4D29B490-49B2-11D0-93C3-7E0706000000} / Id: 46
'    Nombre: Standard Revision (Revisión de la norma) / Valor:  / Id: 47
'    Nombre: Manufacturer (Fabricante) / Valor:  / Id: 48
'    Nombre: Standards Organization (Organismo de normalización) / Valor:  / Id: 49
'    Nombre: Language (Idioma) / Valor:  / Id: 50
'    Nombre: Defer Updates (Aplazar actualizaciones) / Valor: False / Id: 51
'    Nombre: Standard Revision (Revisión de la norma) / Valor:  / Id: 47
'    Nombre: Manufacturer (Fabricante) / Valor:  / Id: 48
'    Nombre: Standards Organization (Organismo de normalización) / Valor:  / Id: 49
'    Nombre: Language (Idioma) / Valor:  / Id: 50
'    Nombre: Defer Updates (Aplazar actualizaciones) / Valor: False / Id: 51
'    Nombre: Size Designation (Designación del tamaño) / Valor:  / Id: 52
'    Nombre: Categories (Categorias) / Valor:  / Id: 56
'    Nombre: Stock Number (Nº de almacenamiento) / Valor:  / Id: 55
'    Nombre: Weld Material (Material de soldadura) / Valor:  / Id: 57
'    Nombre: Mass (Masa) / Valor: 867,514997290746 / Id: 58
'    Nombre: SurfaceArea (Área de superficie) / Valor: 2775,73350191644 / Id: 59
'    Nombre: Volume (Volumen) / Valor: 110,511464623025 / Id: 60
'    Nombre: Density (Densidad) / Valor: 7,85 / Id: 61
'    Nombre: Valid MassProps (Propiedades másicas válidas) / Valor: 31 / Id: 62
'    Nombre: Flat Pattern Width (FlatPatternExtentsWidth) / Valor: 25,9828672105435 / Id: 63
'    Nombre: Flat Pattern Length (FlatPatternExtentsLength) / Valor: 54,1219114736935 / Id: 64
'    Nombre: Flat Pattern Area (FlatPatternExtentsArea) / Valor: 1406,24243900177 / Id: 65
'Propiedades de Inventor definidas por el usuario / Propiedades de Inventor definidas por el usuario
'Internal name: {D5CDD505-2E9C-101B-9397-08002B2CF9AE}

'    Nombre: ExtensionX (ExtensionX) / Valor: 542 mm / Id: 3
'    Nombre: ExtensionY (ExtensionY) / Valor: 261 mm / Id: 4
'    Nombre: DENOMINACION (DENOMINACION) / Valor: CHAPA / Id: 6
'    Nombre: LETRA (LETRA) / Valor:  / Id: 7
'    Nombre: NºORDEN (NºORDEN) / Valor: 0 / Id: 8
'    Nombre: ELEMENTO (ELEMENTO) / Valor: 0 / Id: 10
'    Nombre: Espesor (Espesor) / Valor: 0,8000 mm / Id: 13
#End Region
