Public Module Configuracoes_Botoes

#Region "Funções"

    Public Sub Configurar_VisualizacaoCronograma_CalendarioAcademico(Status As Boolean)

        Visualizacao_Cronograma_CalendarioAcademico = Status

        Call AtualizarComponenteRibbon("Bttn_EdicaoDeCronograma_CalendarioAcademico_Cronograma_VisualizarCronograma")
        Call AtualizarComponenteRibbon("Bttn_EdicaoDeCronograma_CalendarioAcademico_Cronograma_VisualizarCalendarioAcademico")

    End Sub

    Public Sub Configurar_VisualizacaoCronograma_MapaDeSala(Status As Boolean)

        Visualizacao_Cronograma_MapaDeSala = Status

        Call AtualizarComponenteRibbon("Bttn_EdicaoDeCronograma_MapaDeSala_Cronograma_Visualizacao_VisualizarCronograma")
        Call AtualizarComponenteRibbon("Bttn_EdicaoDeCronograma_MapaDeSala_Cronograma_Visualizacao_VisualizarMapaDeSala")

    End Sub

#End Region

End Module
