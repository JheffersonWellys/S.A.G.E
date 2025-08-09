Imports System.Drawing

Module Controles_Botoes

#Region "GetImage"

#Region "Ribbon"

    Public Function GetImage_IniciarSessao() As Bitmap
        Return My.Resources.Icn_Bttn_Logon_Login_IniciarSessao
    End Function

    Public Function GetImage_SolicitarAcessoViaTeams() As Bitmap
        Return My.Resources.Icn_Bttn_Logon_Configuracao_Informacoes_UsuarioLogado_SolicitarAcessoViaTeams
    End Function

    Public Function GetImage_FinalizarSessao() As Bitmap
        Return My.Resources.Icn_Bttn_MenuInicial_Logout_FinalizarSessao
    End Function

    Public Function GetImage_GerenciarUsuarios() As Bitmap
        Return My.Resources.Icn_Bttn_MenuInicial_Configuracoes_GestaoDeAcessos_GerenciarUsuarios
    End Function

    Public Function GetImage_GerenciarAcessosDeUsuarios() As Bitmap
        Return My.Resources.Icn_Bttn_MenuInicial_Configuracoes_GestaoDeAcessos_GerenciarAcessosDeUsuarios
    End Function

    Public Function GetImage_GerenciarUnidadesEducacionais() As Bitmap
        Return My.Resources.Icn_Bttn_MenuInicial_Configuracoes_GestaoDeInfraestrutura_GerenciarUnidadesEducacionais
    End Function

    Public Function GetImage_GerenciarBlocos() As Bitmap
        Return My.Resources.Icn_Bttn_MenuInicial_Configuracoes_GestaoDeInfraestrutura_GerenciarBlocos
    End Function

    Public Function GetImage_GerenciarAndares() As Bitmap
        Return My.Resources.Icn_Bttn_MenuInicial_Configuracoes_GestaoDeInfraestrutura_GerenciarAndares
    End Function

    Public Function GetImage_GerenciarSalas() As Bitmap
        Return My.Resources.Icn_Bttn_MenuInicial_Configuracoes_GestaoDeInfraestrutura_GerenciarSalas
    End Function

    Public Function GetImage_GerenciarDocentes() As Bitmap
        Return My.Resources.Icn_Bttn_MenuInicial_Configuracoes_GestaoEducacional_GerenciarDocentes
    End Function

    Public Function GetImage_GerenciarAutorizacoesParaLecionar() As Bitmap
        Return My.Resources.Icn_Bttn_MenuInicial_Configuracoes_GestaoEducacional_GerenciarAutorizacoesParaLecionar
    End Function

    Public Function GetImage_GerenciarAtestados() As Bitmap
        Return My.Resources.Icn_Bttn_MenuInicial_Configuracoes_GestaoEducacional_GerenciarAtestados
    End Function

    Public Function GetImage_GerenciarAreasProfissionais() As Bitmap
        Return My.Resources.Icn_Bttn_MenuInicial_Configuracoes_GestaoAcademica_GerenciarAreasProfissionais
    End Function

    Public Function GetImage_GerenciarCursos() As Bitmap
        Return My.Resources.Icn_Bttn_MenuInicial_Configuracoes_GestaoAcademica_GerenciarCursos
    End Function

    Public Function GetImage_GerenciarUnidadesCurriculares() As Bitmap
        Return My.Resources.Icn_Bttn_MenuInicial_Configuracoes_GestaoAcademica_GerenciarUnidadesCurriculares
    End Function

    Public Function GetImage_GerenciarFeriados() As Bitmap
        Return My.Resources.Icn_Bttn_MenuInicial_Configuracoes_GestaoDeEventos_GerenciarFeriados
    End Function

    Public Function GetImage_GerenciarRecessos() As Bitmap
        Return My.Resources.Icn_Bttn_MenuInicial_Configuracoes_GestaoDeEventos_GerenciarRecessos
    End Function

    Public Function GetImage_GerenciarDatasEventuaisPorUnidadeEducacional() As Bitmap
        Return My.Resources.Icn_Bttn_MenuInicial_Configuracoes_GestaoDeEventos_GerenciarDatasEventuaisPorUnidadeEducacional
    End Function

    Public Function GetImage_GestaoDeCalendariosAcademicos() As Bitmap
        Return My.Resources.Icn_Bttn_MenuInicial_Cronogramas_GestaoDeCalendariosAcademicos
    End Function

    Public Function GetImage_GestaoDeMapasDeSala() As Bitmap
        Return My.Resources.Icn_Bttn_MenuInicial_Cronogramas_GestaoDeMapasDeSala
    End Function

    Public Function GetImage_CalendarioAcademico_VoltarParaMenuInicial() As Bitmap
        Return My.Resources.Icn_Bttn_EdicaoDeCronograma_CalendarioAcademico_VoltarPara_MenuInicial
    End Function

    Public Function GetImage_CalendarioAcademico_CriarComIA() As Bitmap
        Return My.Resources.Icn_Bttn_EdicaoDeCronograma_CalendarioAcademico_Cronograma_Criacao_CriarComIA
    End Function

    Public Function GetImage_CalendarioAcademico_RecriarComIA() As Bitmap
        Return My.Resources.Icn_Bttn_EdicaoDeCronograma_CalendarioAcademico_Cronograma_Criacao_RecriarComIA
    End Function

    Public Function GetImage_CalendarioAcademico_Edicao_EditarCronogramaManualmente() As Bitmap
        Return My.Resources.Icn_Bttn_EdicaoDeCronograma_CalendarioAcademico_Cronograma_Edicao_EditarCronogramaManualmente
    End Function

    Public Function GetImage_CalendarioAcademico_GerenciarDatasEventuaisPorTurma() As Bitmap
        Return My.Resources.Icn_Bttn_EdicaoDeCronograma_CalendarioAcademico_GestaoDeEventos_GerenciarDatasEventuaisPorTurma
    End Function

    Public Function GetImage_CalendarioAcademico_GerenciarDatasDeProjetosIntegradores() As Bitmap
        Return My.Resources.Icn_Bttn_EdicaoDeCronograma_CalendarioAcademico_GestaoDeEventos_GerenciarDatasDeProjetosIntegradores
    End Function

    Public Function GetImage_CalendarioAcademico_GerenciarDatasDeEstagiosProfissionaisSupervisionados() As Bitmap
        Return My.Resources.Icn_Bttn_EdicaoDeCronograma_CalendarioAcademico_GestaoDeEventos_GerenciarDatasDeEstagiosProfissionaisSupervisionados
    End Function

    Public Function GetImage_CalendarioAcademico_VisualizarErros() As Bitmap
        Return My.Resources.Icn_Bttn_EdicaoDeCronograma_CalendarioAcademico_Cronograma_Validacao_VisualizarErros
    End Function

    Public Function GetImage_CalendarioAcademico_VisualizarCronograma() As Bitmap
        Return My.Resources.Icn_Bttn_EdicaoDeCronograma_CalendarioAcademico_Cronograma_VisualizarCronograma
    End Function

    Public Function GetImage_CalendarioAcademico_VisualizarCalendarioAcademico() As Bitmap
        Return My.Resources.Icn_Bttn_EdicaoDeCronograma_CalendarioAcademico_Cronograma_VisualizarCalendarioAcademico
    End Function

    Public Function GetImage_CalendarioAcademico_ExportacaoEmPDF() As Bitmap
        Return My.Resources.Icn_Bttn_EdicaoDeCronograma_CalendarioAcademico_Cronograma_Exportacao_EmPDF
    End Function

    Public Function GetImage_CalendarioAcademico_ExportacaoEmXLSX() As Bitmap
        Return My.Resources.Icn_Bttn_EdicaoDeCronograma_CalendarioAcademico_Cronograma_Exportacao_EmXLSX
    End Function

    Public Function GetImage_MapaDeSala_VoltarParaMenuInicial() As Bitmap
        Return My.Resources.Icn_Bttn_EdicaoDeCronograma_MapaDeSala_VoltarPara_MenuInicial
    End Function

    Public Function GetImage_MapaDeSala_CriarComIA() As Bitmap
        Return My.Resources.Icn_Bttn_EdicaoDeCronograma_MapaDeSala_Cronograma_Criacao_CriarComIA
    End Function

    Public Function GetImage_MapaDeSala_RecriarComIA() As Bitmap
        Return My.Resources.Icn_Bttn_EdicaoDeCronograma_MapaDeSala_Cronograma_Criacao_RecriarComIA
    End Function

    Public Function GetImage_MapaDeSala_VisualizarErros() As Bitmap
        Return My.Resources.Icn_Bttn_EdicaoDeCronograma_MapaDeSala_Cronograma_Validacao_VisualizarErros
    End Function

    Public Function GetImage_MapaDeSala_VisualizarCronograma() As Bitmap
        Return My.Resources.Icn_Bttn_EdicaoDeCronograma_MapaDeSala_Cronograma_Visualizacao_VisualizarCronograma
    End Function

    Public Function GetImage_MapaDeSala_VisualizarMapaDeSala() As Bitmap
        Return My.Resources.Icn_Bttn_EdicaoDeCronograma_MapaDeSala_Cronograma_Visualizacao_VisualizarMapaDeSala
    End Function

    Public Function GetImage_MapaDeSala_ExportacaoEmPDF() As Bitmap
        Return My.Resources.Icn_Bttn_EdicaoDeCronograma_MapaDeSala_Cronograma_Exportacao_EmPDF
    End Function

    Public Function GetImage_MapaDeSala_ExportacaoEmXLSX() As Bitmap
        Return My.Resources.Icn_Bttn_EdicaoDeCronograma_MapaDeSala_Cronograma_Exportacao_EmXLSX
    End Function

#End Region

#Region "Backstage"

    Public Function GetImage_Configuracoes_ConfigurarBancoDeDados() As Bitmap
        Return My.Resources.Icn_Bttn_InformacoesSobreOSistema_Configuracoes_ConfigurarBancoDeDados
    End Function

    Public Function GetImage_Configuracoes_ReconfigurarBancoDeDados() As Bitmap
        Return My.Resources.Icn_Bttn_InformacoesSobreOSistema_Configuracoes_BancoDeDados_ReconfigurarBancoDeDados
    End Function

#End Region

#End Region

#Region "GetEnabled"

#Region "Ribbon"

    Public Function GetEnabled_IniciarSessao() As Boolean
        Return True
    End Function

    Public Function GetEnabled_FinalizarSessao() As Boolean
        Return True
    End Function

    Public Function GetEnabled_GerenciarUsuarios() As Boolean
        Return True
    End Function

    Public Function GetEnabled_GerenciarAcessosDeUsuarios() As Boolean
        Return True
    End Function

    Public Function GetEnabled_GerenciarUnidadesEducacionais() As Boolean
        Return True
    End Function

    Public Function GetEnabled_GerenciarBlocos() As Boolean
        Return True
    End Function

    Public Function GetEnabled_GerenciarAndares() As Boolean
        Return True
    End Function

    Public Function GetEnabled_GerenciarSalas() As Boolean
        Return True
    End Function

    Public Function GetEnabled_GerenciarDocentes() As Boolean
        Return True
    End Function

    Public Function GetEnabled_GerenciarAutorizacoesParaLecionar() As Boolean
        Return True
    End Function

    Public Function GetEnabled_GerenciarAtestados() As Boolean
        Return True
    End Function

    Public Function GetEnabled_GerenciarAreasProfissionais() As Boolean
        Return True
    End Function

    Public Function GetEnabled_GerenciarCursos() As Boolean
        Return True
    End Function

    Public Function GetEnabled_GerenciarUnidadesCurriculares() As Boolean
        Return True
    End Function

    Public Function GetEnabled_GerenciarFeriados() As Boolean
        Return True
    End Function

    Public Function GetEnabled_GerenciarRecessos() As Boolean
        Return True
    End Function

    Public Function GetEnabled_GerenciarDatasEventuaisPorUnidadeEducacional() As Boolean
        Return True
    End Function

    Public Function GetEnabled_GestaoDeCalendariosAcademicos() As Boolean
        Return True
    End Function

    Public Function GetEnabled_GestaoDeMapasDeSala() As Boolean
        Return True
    End Function

    Public Function GetEnabled_CalendarioAcademico_CriarComIA() As Boolean
        Return True
    End Function

    Public Function GetEnabled_CalendarioAcademico_RecriarComIA() As Boolean
        Return True
    End Function

    Public Function GetEnabled_CalendarioAcademico_EditarCronogramaManualmente() As Boolean
        Return True
    End Function

    Public Function GetEnabled_CalendarioAcademico_GerenciarDatasEventuaisPorTurma() As Boolean
        Return True
    End Function

    Public Function GetEnabled_CalendarioAcademico_GerenciarDatasDeProjetosIntegradores() As Boolean
        Return True
    End Function

    Public Function GetEnabled_CalendarioAcademico_GerenciarDatasDeEstagiosProfissionaisSupervisionados() As Boolean
        Return True
    End Function

    Public Function GetEnabled_CalendarioAcademico_VisualizarErros() As Boolean
        Return True
    End Function

    Public Function GetEnabled_CalendarioAcademico_VisualizarCronograma() As Boolean
        Return True
    End Function

    Public Function GetEnabled_CalendarioAcademico_VisualizarCalendarioAcademico() As Boolean
        Return True
    End Function

    Public Function GetEnabled_CalendarioAcademico_ExportacaoEmPDF() As Boolean
        Return True
    End Function

    Public Function GetEnabled_CalendarioAcademico_ExportacaoEmXLSX() As Boolean
        Return True
    End Function

    Public Function GetEnabled_MapaDeSala_CriarComIA() As Boolean
        Return True
    End Function

    Public Function GetEnabled_MapaDeSala_RecriarComIA() As Boolean
        Return True
    End Function

    Public Function GetEnabled_MapaDeSala_VisualizarErros() As Boolean
        Return True
    End Function

    Public Function GetEnabled_MapaDeSala_VisualizarCronograma() As Boolean
        Return True
    End Function

    Public Function GetEnabled_MapaDeSala_VisualizarMapaDeSala() As Boolean
        Return True
    End Function

    Public Function GetEnabled_MapaDeSala_ExportacaoEmPDF() As Boolean
        Return True
    End Function

    Public Function GetEnabled_MapaDeSala_ExportacaoEmXLSX() As Boolean
        Return True
    End Function

#End Region

#End Region

#Region "GetVisible"

#Region "Ribbon"

    Public Function GetVisible_CalendarioAcademico_Cronograma_VisualizarCronograma() As Boolean
        Return Visualizacao_Cronograma_CalendarioAcademico = False
    End Function

    Public Function GetVisible_CalendarioAcademico_Cronograma_VisualizarCalendarioAcademico() As Boolean
        Return Visualizacao_Cronograma_CalendarioAcademico = True
    End Function

    Public Function GetVisible_MapaDeSala_Cronograma_Visualizacao_VisualizarCronograma() As Boolean
        Return Visualizacao_Cronograma_MapaDeSala = False
    End Function

    Public Function GetVisible_MapaDeSala_Cronograma_Visualizacao_VisualizarMapaDeSala() As Boolean
        Return Visualizacao_Cronograma_MapaDeSala = True
    End Function

#End Region

#End Region

End Module
