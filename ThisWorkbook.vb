
Public Class ThisWorkbook

    Private Sub ThisWorkbook_Startup() Handles Me.Startup

        AlternarModoFoco(True)

    End Sub

    Private Sub ThisWorkbook_Shutdown() Handles Me.Shutdown

        AlternarModoFoco(False)

    End Sub

    Private Sub ThisWorkbook_SheetActivate(Sh As Object) Handles Me.SheetActivate

        AlternarModoFoco(True)

    End Sub

    Private Sub ThisWorkbook_SheetDeactivate(Sh As Object) Handles Me.SheetDeactivate

        AlternarModoFoco(False)

    End Sub

    Protected Overrides Function CreateRibbonExtensibilityObject() As Microsoft.Office.Core.IRibbonExtensibility

        Return New Ribbon_SAGE()

    End Function

End Class
