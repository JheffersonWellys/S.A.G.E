Public Module Acoes_Botoes

#Region "OnAction"

#Region "Ribbon"

    Public Sub OnAction_IniciarSessao()
        Call NavegarParaAba("Tb_MenuInicial")
    End Sub

    Public Sub OnAction_SolicitarAcessoViaTeams()
        Informacao_BancoDeDados_EmUso = False
        Call AtualizarRibbon()
    End Sub

    Public Sub OnAction_FinalizarSessao()
        Call NavegarParaAba("Tb_Logon")
    End Sub

    Public Sub OnAction_GerenciarUsuarios()
        MessageBox.Show("Em breve...")
    End Sub

    Public Sub OnAction_GerenciarAcessosDeUsuarios()
        MessageBox.Show("Em breve...")
    End Sub

    Public Sub OnAction_GerenciarUnidadesEducacionais()
        MessageBox.Show("Em breve...")
    End Sub

    Public Sub OnAction_GerenciarBlocos()
        MessageBox.Show("Em breve...")
    End Sub

    Public Sub OnAction_GerenciarAndares()
        MessageBox.Show("Em breve...")
    End Sub

    Public Sub OnAction_GerenciarSalas()
        MessageBox.Show("Em breve...")
    End Sub

    Public Sub OnAction_GerenciarDocentes()
        MessageBox.Show("Em breve...")
    End Sub

    Public Sub OnAction_GerenciarAutorizacoesParaLecionar()
        MessageBox.Show("Em breve...")
    End Sub

    Public Sub OnAction_GerenciarAtestados()
        MessageBox.Show("Em breve...")
    End Sub

    Public Sub OnAction_GerenciarAreasProfissionais()
        MessageBox.Show("Em breve...")
    End Sub

    Public Sub OnAction_GerenciarCursos()
        MessageBox.Show("Em breve...")
    End Sub

    Public Sub OnAction_GerenciarUnidadesCurriculares()
        MessageBox.Show("Em breve...")
    End Sub

    Public Sub OnAction_GerenciarFeriados()
        MessageBox.Show("Em breve...")
    End Sub

    Public Sub OnAction_GerenciarRecessos()
        MessageBox.Show("Em breve...")
    End Sub

    Public Sub OnAction_GerenciarDatasEventuaisPorUnidadeEducacional()
        MessageBox.Show("Em breve...")
    End Sub

    Public Sub OnAction_GestaoDeCalendariosAcademicos()
        Call NavegarParaAba("Tb_EdicaoDeCronograma_CalendarioAcademico")
    End Sub

    Public Sub OnAction_GestaoDeMapasDeSala()
        Call NavegarParaAba("Tb_EdicaoDeCronograma_MapaDeSala")
    End Sub

    Public Sub OnAction_CalendarioAcademico_VoltarParaMenuInicial()
        Call NavegarParaAba("Tb_MenuInicial")
    End Sub

    Public Sub OnAction_CalendarioAcademico_CriarComIA()
        MessageBox.Show("Em breve...")
    End Sub

    Public Sub OnAction_CalendarioAcademico_RecriarComIA()
        MessageBox.Show("Em breve...")
    End Sub

    Public Sub OnAction_CalendarioAcademico_Edicao_EditarCronogramaManualmente()
        MessageBox.Show("Em breve...")
    End Sub

    Public Sub OnAction_CalendarioAcademico_GerenciarDatasEventuaisPorTurma()
        MessageBox.Show("Em breve...")
    End Sub

    Public Sub OnAction_CalendarioAcademico_GerenciarDatasDeProjetosIntegradores()
        MessageBox.Show("Em breve...")
    End Sub

    Public Sub OnAction_CalendarioAcademico_GerenciarDatasDeEstagiosProfissionaisSupervisionados()
        MessageBox.Show("Em breve...")
    End Sub

    Public Sub OnAction_CalendarioAcademico_ExportacaoVisualizarErros()
        MessageBox.Show("Em breve...")
    End Sub

    Public Sub OnAction_CalendarioAcademico_Cronograma_VisualizarCronograma()
        Call Configurar_VisualizacaoCronograma_CalendarioAcademico(True)
    End Sub

    Public Sub OnAction_CalendarioAcademico_Cronograma_VisualizarCalendarioAcademico()
        Call Configurar_VisualizacaoCronograma_CalendarioAcademico(False)
    End Sub

    Public Sub OnAction_CalendarioAcademico_ExportacaoEmPDF()
        MessageBox.Show("Em breve...")
    End Sub

    Public Sub OnAction_CalendarioAcademico_ExportacaoEmXLSX()
        MessageBox.Show("Em breve...")
    End Sub

    Public Sub OnAction_MapaDeSala_VoltarParaMenuInicial()
        Call NavegarParaAba("Tb_MenuInicial")
    End Sub

    Public Sub OnAction_MapaDeSala_CriarComIA()
        MessageBox.Show("Em breve...")
    End Sub

    Public Sub OnAction_MapaDeSala_RecriarComIA()
        MessageBox.Show("Em breve...")
    End Sub

    Public Sub OnAction_MapaDeSala_ExportacaoVisualizarErros()
        MessageBox.Show("Em breve...")
    End Sub

    Public Sub OnAction_MapaDeSala_VisualizarCronograma()
        Call Configurar_VisualizacaoCronograma_MapaDeSala(True)
    End Sub

    Public Sub OnAction_MapaDeSala_VisualizarMapaDeSala()
        Call Configurar_VisualizacaoCronograma_MapaDeSala(False)
    End Sub

    Public Sub OnAction_MapaDeSala_ExportacaoEmPDF()
        MessageBox.Show("Em breve...")
    End Sub

    Public Sub OnAction_MapaDeSala_ExportacaoEmXLSX()
        MessageBox.Show("Em breve...")
    End Sub

#End Region

#Region "Backstage"

    Public Sub OnAction_Configuracoes_ConfigurarBancoDeDados()
        Informacao_BancoDeDados_Configurado = True
        Call AtualizarRibbon()
    End Sub

    Public Sub OnAction_Configuracoes_ReconfigurarBancoDeDados()
        Informacao_BancoDeDados_Configurado = False
        Call AtualizarRibbon()
    End Sub

#End Region

#End Region

End Module
