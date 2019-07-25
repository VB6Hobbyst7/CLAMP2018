Imports Inventor
Partial Public Class Inventor2acad
    Public Sub iLogicRunRule(queApp As Inventor.Application, ByRef queDoc As Inventor.Document, rulename As String)
        Dim iLogicAddIn As Object = iLogicGetAutomation(queApp)
        If (iLogicAddIn Is Nothing) Then Exit Sub
        ''
        Try
            '' Run internal Rule
            iLogicAddIn.RunRule(queDoc, rulename)
        Catch ex As Exception
            Try
                '' Run external Rule
                iLogicAddIn.RunExternalRule(queDoc, rulename)
            Catch ex1 As Exception
                MsgBox("Rule --> " & rulename & " not found")
            End Try
        End Try
    End Sub
    ''
    Private Function iLogicGetAutomation(queApp As Inventor.Application) As Object
        Dim addIn As Inventor.ApplicationAddIn
        Try
            addIn = queApp.ApplicationAddIns.ItemById("{3bdd8d79-2179-4b11-8a5a-257b1c0263ac}")
            If addIn.Activated = False Then addIn.Activate()
            iLogicGetAutomation = addIn.Automation
        Catch
            iLogicGetAutomation = Nothing
        End Try
    End Function
End Class
