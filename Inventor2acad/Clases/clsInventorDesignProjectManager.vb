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
    Public Function Project_GetPath()
        Dim oDpm As DesignProjectManager = oAppI.DesignProjectManager
        Return oDpm.ActiveDesignProject.FullFileName
    End Function
    '
    Public Sub Project_Activate()

    End Sub

    Public Function Project_LoadAndActivate(queFullPro As String) As Boolean
        Dim resultado As Boolean = False
        '
        Try
            If oAppI.DesignProjectManager.ActiveDesignProject.FullFileName <> queFullPro Then
                ' El proyecto activo No es el mismo que InventorProject
                oAppI.Documents.CloseAll()
                Dim oDpro As Inventor.DesignProject = Nothing
                For Each oDp In oAppI.DesignProjectManager.DesignProjects
                    If oDp.FullFileName = queFullPro Then
                        oDp.Activate()
                        oDpro = oDp
                        resultado = True
                        Exit For
                    End If
                Next
                '
                If oDpro Is Nothing And resultado = False Then
                    oDpro = oAppI.DesignProjectManager.DesignProjects.AddExisting(queFullPro)
                    oDpro.Activate()
                    resultado = True
                End If
                oDpro = Nothing
            Else
                ' The Project is Active
                resultado = True
            End If
        Catch ex As Exception
            resultado = False
        End Try
        '
        Return resultado
    End Function
    Public Function Project_FileInFoldersProject(queF As String) As Boolean
        Dim resultado As Boolean = False
        arrDirsPro = New ArrayList
        Dim oDpM As DesignProjectManager = oAppI.DesignProjectManager
        '' Directorio del proyecto
        arrDirsPro.Add(IO.Path.GetDirectoryName(oDpM.ActiveDesignProject.FullFileName))
        '' Directorio del espacio de trabajo (debería ser igual que el anterior)
        If arrDirsPro.Contains(oDpM.ActiveDesignProject.WorkspacePath) = False Then _
        arrDirsPro.Add(oDpM.ActiveDesignProject.WorkspacePath)
        '' Directorios del grupo de trabajo
        For Each ProPath As ProjectPath In oDpM.ActiveDesignProject.WorkgroupPaths
            If arrDirsPro.Contains(ProPath.Path) = False Then arrDirsPro.Add(ProPath.Path)
        Next
        ''
        '' Ya tenemos todos los directorios de trabajo y acceso del proyecto (WorkSpace y Grupos de trabajo)
        For Each camino In arrDirsPro
            If queF.StartsWith(camino) Then
                resultado = True
                Exit For
            End If
        Next
        Return resultado
    End Function
End Class
