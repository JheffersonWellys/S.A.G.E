Public Module Ribbon_SAGE_Functions

    Public Sub OnActionRibbon_AbrirAba(TabSelecionada As String)
        RibbonUI_TabAtiva = TabSelecionada
        RibbonUI_SAGE.Invalidate()
    End Sub

End Module
