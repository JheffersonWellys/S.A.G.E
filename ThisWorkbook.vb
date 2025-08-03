
Public Class ThisWorkbook

    Private Sub ThisWorkbook_Startup() Handles Me.Startup

    End Sub

    Private Sub ThisWorkbook_Shutdown() Handles Me.Shutdown

    End Sub

    Protected Overrides Function CreateRibbonExtensibilityObject() As Microsoft.Office.Core.IRibbonExtensibility
        Return New Ribbon_Menu()
    End Function

End Class
