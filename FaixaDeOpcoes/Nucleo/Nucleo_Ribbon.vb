Public Module Nucleo_ControlesRibbon

#Region "Funções"

    Public Sub AtualizarRibbon()

        UIRibbon_SAGE.Invalidate()

    End Sub

    Public Sub AtualizarComponenteRibbon(IdComponente As String)

        UIRibbon_SAGE.InvalidateControl(IdComponente)

    End Sub

#End Region

End Module
