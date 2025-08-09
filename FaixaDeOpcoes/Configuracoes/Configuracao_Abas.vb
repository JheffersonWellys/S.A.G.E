Public Module Configuracao_Abas

#Region "Variáveis"

    Public TabAtual As String = "Tb_Logon"

#End Region

#Region "Funções"

    Public Sub NavegarParaAba(TabSelecionada As String)

        TabAtual = TabSelecionada
        Call AtualizarRibbon()

    End Sub

#End Region

End Module
